import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import serial
import time
import json
import threading
import sys
import os

# --- OPTIONAL: For a modern UI theme ---
try:
    import sv_ttk
except ImportError:
    sv_ttk = None

# --- OPTIONAL: For plotting and styling ---
try:
    import numpy as np
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    import matplotlib.lines as mlines

    # Use Seaborn for a more professional look if available
    try:
        import seaborn as sns

        sns.set_style("whitegrid")
        print("Seaborn style applied.")
    except ImportError:
        sns = None
except ImportError:
    np = None
    plt = None
    FigureCanvasTkAgg = None
    Figure = None
    mlines = None
    sns = None


# ==============================================================================
# ## Asset Path Helper for PyInstaller ##
# ==============================================================================
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ==============================================================================
# ## Backend Controller ##
# ==============================================================================
class MinipulsController:
    """A Python class to control the Gilson MINIPULS 3 via the GSIOC protocol."""

    def __init__(self, port, unit_id=30, baudrate=19200, logger_func=print, debug_mode=False, command_interval=0.2):
        self.port, self.unit_id, self.baudrate = port, unit_id, baudrate
        self.ser, self.is_connected, self.logger = None, False, logger_func
        self.debug_mode = debug_mode
        self.command_interval = command_interval

    def connect(self):
        if self.debug_mode:
            self.logger("DEBUG MODE: Virtual connection established.")
            self.is_connected = True
            return True
        try:
            self.ser = serial.Serial(self.port, self.baudrate, bytesize=serial.EIGHTBITS, parity=serial.PARITY_EVEN,
                                     stopbits=serial.STOPBITS_ONE, timeout=1)
            self.logger(f"Serial port {self.port} opened.")
            self.logger(f"Connecting to Unit ID: {self.unit_id}...")
            self.ser.write(bytes([255]));
            time.sleep(0.05)
            connect_command = bytes([self.unit_id + 128])
            self.ser.write(connect_command)
            response = self.ser.read(1)
            if response == connect_command:
                self.logger(f"Successfully connected to pump (ID: {self.unit_id}).");
                self.is_connected = True;
                return True
            else:
                self.logger(f"Connection failed. Expected {connect_command.hex()} but received {response.hex()}");
                self.ser.close();
                return False
        except serial.SerialException as e:
            self.logger(f"Connection Error: {e}");
            return False

    def disconnect(self):
        if self.debug_mode:
            self.logger("DEBUG MODE: Virtual connection closed.")
            self.is_connected = False
            return
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(bytes([255]));
                time.sleep(0.05);
                self.ser.close()
                self.logger(f"Serial port {self.port} closed.")
            except Exception as e:
                self.logger(f"Error during disconnect: {e}")
        self.is_connected = False

    def send_buffered_command(self, command, wait=True):
        if not self.is_connected:
            self.logger("Error: Pump not connected.")
            return

        if self.debug_mode:
            self.logger(f"DEBUG CMD > {command}")
            if wait:
                time.sleep(self.command_interval)
            return

        full_command = b"\n" + command.encode("ascii") + b"\r"
        self.logger(f"Sending command: {command}")
        self.ser.write(full_command)
        if wait:
            time.sleep(self.command_interval)

    def set_command_interval(self, interval):
        self.command_interval = interval
        self.logger(f"Command interval set to: {self.command_interval}s")

    def set_remote_mode(self):
        self.send_buffered_command("SR")

    def set_keypad_mode(self):
        self.send_buffered_command("SK")

    def start_forward(self):
        # Do not wait internally so caller controls timing
        self.send_buffered_command("K>", wait=False)

    def start_backward(self):
        # Do not wait internally so caller controls timing
        self.send_buffered_command("K<", wait=False)

    def stop(self):
        self.send_buffered_command("KH")

    def set_speed(self, rpm, wait=True):
        if not (0 <= rpm <= 48):
            rpm = max(0, min(48, rpm))
            self.logger(f"Warning: RPM value clamped to {rpm}.")
        self.send_buffered_command(f"R{int(rpm * 100)}", wait=wait)


# ==============================================================================
# ## Dialog Windows for Adding Steps ##
# ==============================================================================
class AddPhaseDialog(tk.Toplevel):
    def __init__(self, parent, existing_data=None):
        super().__init__(parent);
        self.transient(parent)
        self.title("Edit Phase" if existing_data else "Add New Phase");
        self.result = None
        body = ttk.Frame(self, padding="15");
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}");
        body.pack(padx=10, pady=10)

        ttk.Label(body, text="Direction:").grid(row=0, column=0, sticky='w', pady=5)
        self.direction_cb = ttk.Combobox(body, values=["Forward", "Backward"], state="readonly", width=15);
        self.direction_cb.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(body, text="Speed Mode:").grid(row=1, column=0, sticky='w', pady=5)
        self.speed_mode_cb = ttk.Combobox(body, values=["Fixed", "Ramp"], state="readonly", width=15);
        self.speed_mode_cb.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(body, text="Target RPM:").grid(row=2, column=0, sticky='w', pady=5)
        self.rpm_entry = ttk.Entry(body, width=17);
        self.rpm_entry.grid(row=2, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(body, text="Duration:").grid(row=3, column=0, sticky='w', pady=5)
        self.duration_entry = ttk.Entry(body, width=17);
        self.duration_entry.grid(row=3, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(body, text="Unit:").grid(row=4, column=0, sticky='w', pady=5)
        self.unit_combobox = ttk.Combobox(body, values=["s", "min", "hr"], state="readonly", width=15);
        self.unit_combobox.grid(row=4, column=1, sticky='ew', pady=5, padx=5)

        if existing_data:
            self.direction_cb.set(existing_data.get("direction", "Forward"));
            self.speed_mode_cb.set(existing_data.get("mode", "Fixed"))
            self.rpm_entry.insert(0, existing_data.get("rpm", 0.0));
            self.duration_entry.insert(0, existing_data.get("duration", 10))
            self.unit_combobox.set(existing_data.get("unit", "s"))
        else:
            self.direction_cb.set("Forward");
            self.speed_mode_cb.set("Fixed");
            self.unit_combobox.set("s")

        button_frame = ttk.Frame(self, padding=(0, 0, 0, 10))
        ok_button = ttk.Button(button_frame, text="OK", command=self.on_ok, style='Accent.TButton');
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.destroy)
        ok_button.pack(side='left', expand=True, fill='x', padx=10);
        cancel_button.pack(side='right', expand=True, fill='x', padx=10)
        button_frame.pack(fill='x', padx=10);
        self.grab_set();
        self.wait_window(self)

    def on_ok(self):
        try:
            rpm = float(self.rpm_entry.get());
            duration = float(self.duration_entry.get())
            if not (0 <= rpm <= 48 and duration >= 0): raise ValueError
            self.result = {"type": "Phase", "direction": self.direction_cb.get(), "mode": self.speed_mode_cb.get(),
                           "rpm": rpm, "duration": duration, "unit": self.unit_combobox.get()};
            self.destroy()
        except (ValueError, TypeError):
            messagebox.showerror("Input Error", "RPM must be 0-48 and Duration must be a positive number.", parent=self)


class AddCycleDialog(tk.Toplevel):
    def __init__(self, parent, existing_data=None):
        super().__init__(parent);
        self.transient(parent);
        self.title("Edit Cycle" if existing_data else "Add New Cycle");
        self.result = None
        body = ttk.Frame(self, padding="15");
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}");
        body.pack(padx=10, pady=10)

        ttk.Label(body, text="Start Phase #:").grid(row=0, column=0, sticky='w', pady=5)
        self.start_entry = ttk.Entry(body, width=15);
        self.start_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(body, text="End Phase #:").grid(row=1, column=0, sticky='w', pady=5)
        self.end_entry = ttk.Entry(body, width=15);
        self.end_entry.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(body, text="Number of Repeats:").grid(row=2, column=0, sticky='w', pady=5)
        self.repeats_entry = ttk.Entry(body, width=15);
        self.repeats_entry.grid(row=2, column=1, sticky='ew', pady=5, padx=5)

        if existing_data:
            self.start_entry.insert(0, existing_data.get("start_phase", 1));
            self.end_entry.insert(0, existing_data.get("end_phase", 1));
            self.repeats_entry.insert(0, existing_data.get("repeats", 10))

        button_frame = ttk.Frame(self, padding=(0, 0, 0, 10))
        ok_button = ttk.Button(button_frame, text="OK", command=self.on_ok, style='Accent.TButton');
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.destroy)
        ok_button.pack(side='left', expand=True, fill='x', padx=10);
        cancel_button.pack(side='right', expand=True, fill='x', padx=10)
        button_frame.pack(fill='x', padx=10);
        self.grab_set();
        self.wait_window(self)

    def on_ok(self):
        try:
            start = int(self.start_entry.get());
            end = int(self.end_entry.get());
            repeats = int(self.repeats_entry.get())
            if not (start > 0 and end >= start and repeats > 0): raise ValueError
            self.result = {"type": "Cycle", "start_phase": start, "end_phase": end, "repeats": repeats};
            self.destroy()
        except (ValueError, TypeError):
            messagebox.showerror("Input Error",
                                 "All fields must be positive integers, and End Phase must be >= Start Phase.",
                                 parent=self)


class ConfirmationDialog(tk.Toplevel):
    def __init__(self, parent, sequence_data):
        super().__init__(parent);
        self.parent = parent
        self.transient(parent);
        self.title("Confirm Sequence Execution");
        self.confirmed = False
        self.geometry(f"+{parent.winfo_rootx() + 100}+{parent.winfo_rooty() + 100}")

        main_frame = ttk.Frame(self, padding="10");
        main_frame.pack(fill=tk.BOTH, expand=True)

        if np and plt and FigureCanvasTkAgg:
            self._create_plot(main_frame, sequence_data)
        else:
            ttk.Label(main_frame,
                      text="Plotting libraries not found.\nFor visualization, run:\npip install matplotlib numpy seaborn",
                      foreground="red").pack(pady=20)

        button_frame = ttk.Frame(self, padding=10)
        ok_button = ttk.Button(button_frame, text="Confirm & Run", command=self.on_confirm, style='Accent.TButton');
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.destroy)
        ok_button.pack(side='left', padx=10);
        cancel_button.pack(side='right', padx=10)
        button_frame.pack();
        self.grab_set();
        self.wait_window(self)

    def _create_plot(self, parent, sequence_data):
        fig = Figure(figsize=(8, 5), dpi=100);
        ax = fig.add_subplot(111)

        plot_data = self.parent._get_expanded_sequence_data(sequence_data)

        total_time_s = plot_data['time_points'][-1] if plot_data['time_points'] else 0
        if total_time_s < 120:
            time_unit, time_factor = 's', 1.0
        elif total_time_s < 7200:
            time_unit, time_factor = 'min', 1.0 / 60.0
        else:
            time_unit, time_factor = 'hr', 1.0 / 3600.0

        time_scaled = [t * time_factor for t in plot_data['time_points']]
        rpm_data = plot_data['rpm_points']
        phase_markers = plot_data['phase_markers']
        cycle_spans = plot_data['cycle_spans']

        ax.plot(time_scaled, rpm_data, color='lightgray', linestyle='--', label='Plan', zorder=1)

        marker_times = [p['time'] * time_factor for p in phase_markers]
        marker_rpms = [p['rpm'] for p in phase_markers]
        ax.plot(marker_times, marker_rpms, 'o', color='#333333', markersize=5, zorder=2)

        for marker in phase_markers:
            ax.text(marker['time'] * time_factor, marker['rpm'] + 1.2, str(marker['step']), ha='center', va='bottom',
                    fontsize=8, fontweight='bold')

        for span in cycle_spans:
            ax.axvspan(span[0] * time_factor, span[1] * time_factor, color='lightskyblue', alpha=0.3, label='Cycle',
                       zorder=0)

        handles, labels = ax.get_legend_handles_labels()
        by_label = dict(zip(labels, handles))
        ax.legend(by_label.values(), by_label.keys())

        ax.set_title("Sequence Execution Plan", fontsize=14, fontweight='bold');
        ax.set_xlabel(f"Time ({time_unit})", fontsize=10);
        ax.set_ylabel("Speed (RPM)", fontsize=10)
        ax.set_ylim(bottom=-2, top=52)
        if time_scaled:
            ax.set_xlim(left=min(time_scaled) - 0.05 * max(time_scaled), right=max(time_scaled) * 1.05)

        if not sns:
            ax.grid(True, linestyle='--', alpha=0.6)

        fig.tight_layout(pad=1.5)
        canvas = FigureCanvasTkAgg(fig, master=parent);
        canvas.draw();
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def on_confirm(self):
        self.confirmed = True;
        self.destroy()


# ==============================================================================
# ## Main GUI Application ##
# ==============================================================================
class PumpControlUI(tk.Tk):
    RAMP_STEP_INTERVAL_S = 0.1
    UPDATE_INTERVAL_MS = 100

    COLOR_FORWARD = '#29b6f6'  # Blue
    COLOR_BACKWARD = '#f44336'  # Red

    def _log(self, message):
        def _append():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
            self.log_text.config(state=tk.DISABLED);
            self.log_text.see(tk.END)

        self.after(0, _append)

    def _update_speed_label(self, value):
        val = float(value)
        self.speed_label.config(text=f"{val:.1f} RPM")
        if self.focus_get() != self.manual_rpm_entry:
            self.manual_rpm_entry.delete(0, tk.END)
            self.manual_rpm_entry.insert(0, f"{val:.1f}")

    def _update_speed_from_entry(self, event=None):
        try:
            val = float(self.manual_rpm_entry.get())
            if not (0 <= val <= 48):
                val = max(0, min(48, val))
                self.manual_rpm_entry.delete(0, tk.END)
                self.manual_rpm_entry.insert(0, f"{val:.1f}")
            self.speed_scale.set(val)
            self.speed_label.config(text=f"{val:.1f} RPM")
        except (ValueError, TypeError):
            pass

    def _set_manual_controls_state(self, state):
        for child in self.manual_frame.winfo_children():
            try:
                child.config(state=state)
                if isinstance(child, ttk.Frame):
                    for grandchild in child.winfo_children():
                        grandchild.config(state=state)
            except tk.TclError:
                pass

    def _connect_pump(self):
        port, unit_id = self.com_port_entry.get(), int(self.unit_id_entry.get())
        is_debug = self.debug_mode_var.get()
        try:
            command_interval = float(self.command_interval_entry.get())
            if command_interval < 0:
                command_interval = 0.05
                self.command_interval_entry.delete(0, tk.END)
                self.command_interval_entry.insert(0, str(command_interval))
        except (ValueError, TypeError):
            messagebox.showerror("Input Error", "Command Interval must be a positive number.")
            return

        self.pump_controller = MinipulsController(port, unit_id, logger_func=self._log, debug_mode=is_debug,
                                                  command_interval=command_interval)

        if self.pump_controller.connect():
            status_text = "Status: Connected (DEBUG)" if is_debug else "Status: Connected"
            self.status_label.config(text=status_text, foreground="#27ae60")
            self.connect_btn.config(state=tk.DISABLED);
            self.disconnect_btn.config(state=tk.NORMAL)
            self.set_interval_btn.config(state=tk.NORMAL)
            self._set_manual_controls_state(tk.NORMAL);
            self.run_seq_btn.config(state=tk.NORMAL)
            self.pump_controller.set_remote_mode()
        else:
            self.status_label.config(text="Status: Connection Failed", foreground="#c0392b")
            self.pump_controller = None

    def _disconnect_pump(self):
        if self.pump_controller: self.pump_controller.set_keypad_mode(); self.pump_controller.disconnect(); self.pump_controller = None
        self.status_label.config(text="Status: Disconnected", foreground="#c0392b")
        self.connect_btn.config(state=tk.NORMAL);
        self.disconnect_btn.config(state=tk.DISABLED)
        self.set_interval_btn.config(state=tk.DISABLED)
        self._set_manual_controls_state(tk.DISABLED);
        self.run_seq_btn.config(state=tk.DISABLED)

    def _set_command_interval(self):
        if not (self.pump_controller and self.pump_controller.is_connected):
            self._log("Cannot set interval. Pump not connected.")
            return

        try:
            command_interval = float(self.command_interval_entry.get())
            if command_interval < 0:
                messagebox.showwarning("Input Warning", "Command interval must be a non-negative number.", parent=self)
                return

            self.pump_controller.set_command_interval(command_interval)

        except (ValueError, TypeError):
            messagebox.showerror("Input Error", "Command Interval must be a valid number.", parent=self)

    def _manual_start_fwd(self):
        if self.pump_controller: self.pump_controller.set_speed(
            self.speed_scale.get()); self.pump_controller.start_forward()

    def _manual_start_rev(self):
        if self.pump_controller: self.pump_controller.set_speed(
            self.speed_scale.get()); self.pump_controller.start_backward()

    def _manual_stop(self):
        if self.pump_controller: self.pump_controller.stop()

    def _add_phase(self):
        dialog = AddPhaseDialog(self)
        if dialog.result:
            self.sequence_data.append(dialog.result)
            self._update_treeview()

    def _add_cycle(self):
        dialog = AddCycleDialog(self)
        if dialog.result:
            self.sequence_data.append(dialog.result)
            self._update_treeview()

    def _edit_item(self, event):
        selected_item_id = self.sequence_tree.focus()
        if not selected_item_id: return

        index = self.sequence_tree.index(selected_item_id)
        item_data = self.sequence_data[index]

        dialog_class = AddPhaseDialog if item_data['type'] == "Phase" else AddCycleDialog
        dialog = dialog_class(self, existing_data=item_data)

        if dialog.result:
            self.sequence_data[index] = dialog.result
            self._update_treeview()

    def _move_item(self, direction):
        selected_item_id = self.sequence_tree.focus()
        if not selected_item_id: return

        index = self.sequence_tree.index(selected_item_id)

        if direction == 'up' and index > 0:
            self.sequence_data[index], self.sequence_data[index - 1] = self.sequence_data[index - 1], \
                self.sequence_data[index]
            new_selection_index = index - 1
        elif direction == 'down' and index < len(self.sequence_data) - 1:
            self.sequence_data[index], self.sequence_data[index + 1] = self.sequence_data[index + 1], \
                self.sequence_data[index]
            new_selection_index = index + 1
        else:
            return

        self._update_treeview()
        new_item_id = self.sequence_tree.get_children()[new_selection_index]
        self.sequence_tree.selection_set(new_item_id)
        self.sequence_tree.focus(new_item_id)

    def _remove_item(self):
        selected_items = self.sequence_tree.selection()
        if not selected_items: return
        if messagebox.askyesno("Confirm", "Remove selected step(s)?"):
            indices_to_remove = sorted([self.sequence_tree.index(item) for item in selected_items], reverse=True)
            for index in indices_to_remove:
                del self.sequence_data[index]
            self._update_treeview()

    def _clear_sequence(self):
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all sequence steps?"):
            self.sequence_data.clear()
            self._update_treeview()

    def _update_treeview(self):
        selected = self.sequence_tree.focus()
        scroll_pos = self.sequence_tree.yview()
        for item in self.sequence_tree.get_children(): self.sequence_tree.delete(item)

        for i, step_data in enumerate(self.sequence_data):
            step_num = i + 1
            if step_data['type'] == 'Phase':
                p = step_data
                details = f"{p['direction']}, {p['mode']} to {p['rpm']} RPM"
                duration_str = f"{p['duration']} {p['unit']}"
                values = (step_num, "Phase", details, duration_str)
            elif step_data['type'] == 'Cycle':
                c = step_data;
                total_duration_s = 0
                if c['start_phase'] <= len(self.sequence_data) and c['end_phase'] <= len(self.sequence_data):
                    for phase_idx in range(c['start_phase'] - 1, c['end_phase']):
                        phase = self.sequence_data[phase_idx]
                        if phase['type'] == 'Phase':
                            d = phase['duration']
                            if phase['unit'] == 'min':
                                d *= 60
                            elif phase['unit'] == 'hr':
                                d *= 3600
                            total_duration_s += d
                details = f"Loop Phases {c['start_phase']}-{c['end_phase']} ({c['repeats']} times)"
                duration_str = f"~{total_duration_s:.1f} s/cycle"
                values = (step_num, "Cycle", details, duration_str)
            self.sequence_tree.insert("", tk.END, values=values)

        if selected:
            try:
                self.sequence_tree.selection_set(selected);
                self.sequence_tree.focus(selected)
            except tk.TclError:
                pass
        self.sequence_tree.yview_moveto(scroll_pos[0])

    def _get_expanded_sequence_data(self, sequence_data):
        command_interval = getattr(self, "pump_controller", None)
        if command_interval:
            command_interval = command_interval.command_interval
        else:
            command_interval = self.RAMP_STEP_INTERVAL_S

        time_points, rpm_points = [0], [0]
        cycle_spans, phase_markers, phase_directions = [], [], []
        current_rpm, elapsed_time_s = 0.0, 0.0
        pc, cycle_counters, iterations, max_iterations = 0, {}, 0, 10000

        sequence_copy = [dict(s) for s in sequence_data]
        for i, step in enumerate(sequence_copy): step['original_index'] = i

        while pc < len(sequence_copy) and iterations < max_iterations:
            instruction = sequence_copy[pc]
            if instruction['type'] == 'Phase':
                phase_markers.append(
                    {
                        "time": elapsed_time_s,
                        "rpm": current_rpm,
                        "step": instruction["original_index"] + 1,
                    }
                )

                duration_s = instruction["duration"]
                if instruction["unit"] == "min":
                    duration_s *= 60
                elif instruction["unit"] == "hr":
                    duration_s *= 3600

                # Store the end time of this phase
                phase_directions.append(
                    {
                        "end_time": elapsed_time_s + duration_s,
                        "direction": instruction["direction"],
                    }
                )

                target_rpm = instruction["rpm"]

                if instruction["mode"] == "Fixed":
                    if current_rpm != target_rpm:
                        time_points.append(elapsed_time_s)
                        rpm_points.append(target_rpm)

                elif instruction["mode"] == "Ramp":
                    num_steps = int(duration_s / command_interval) if duration_s > 0 else 1
                    if num_steps == 0:
                        num_steps = 1
                    time_step = duration_s / num_steps
                    rpm_increment = (target_rpm - current_rpm) / num_steps
                    for i in range(num_steps):
                        time_points.append(elapsed_time_s + (i + 1) * time_step)
                        rpm_points.append(current_rpm + (i + 1) * rpm_increment)

                elapsed_time_s += duration_s
                if instruction["mode"] != "Ramp":
                    time_points.append(elapsed_time_s)
                    rpm_points.append(target_rpm)
                current_rpm = target_rpm
                pc += 1

            elif instruction['type'] == 'Cycle':
                start_idx, end_idx, repeats = instruction['start_phase'] - 1, instruction['end_phase'] - 1, instruction[
                    'repeats']
                if not (0 <= start_idx <= end_idx < len(sequence_copy)): break

                if pc not in cycle_counters: cycle_counters[pc] = {'count': repeats, 'start_time': -1}

                if cycle_counters[pc]['count'] > 0:
                    if cycle_counters[pc]['start_time'] == -1:
                        cycle_counters[pc]['start_time'] = elapsed_time_s

                    cycle_counters[pc]['count'] -= 1
                    pc = start_idx
                else:
                    cycle_start_time = cycle_counters[pc]['start_time']
                    if cycle_start_time != -1:
                        cycle_spans.append((cycle_start_time, elapsed_time_s))
                    del cycle_counters[pc]
                    pc += 1
            iterations += 1

        return {'time_points': time_points, 'rpm_points': rpm_points, 'cycle_spans': cycle_spans,
                'phase_markers': phase_markers, 'phase_directions': phase_directions}

    def _get_total_sequence_time(self):
        # Use the pre-calculated plan if it exists
        if hasattr(self, 'plan_time_points') and self.plan_time_points:
            return self.plan_time_points[-1]

        plot_data = self._get_expanded_sequence_data(self.sequence_data)
        return plot_data['time_points'][-1] if plot_data['time_points'] else 0.0

    def _run_sequence(self):
        if not self.sequence_data: messagebox.showwarning("Warning", "Sequence is empty. Cannot execute."); return

        dialog = ConfirmationDialog(self, self.sequence_data)
        if not dialog.confirmed:
            self._log("Sequence run cancelled by user.");
            return

        # Pre-calculate and store the entire plan
        self.plot_data = self._get_expanded_sequence_data(self.sequence_data)
        self.plan_time_points = self.plot_data['time_points']
        self.plan_rpm_points = self.plot_data['rpm_points']
        self.plan_directions = self.plot_data['phase_directions']
        self.total_sequence_time = self.plan_time_points[-1] if self.plan_time_points else 0.0

        if self.total_sequence_time < 120:
            self.time_unit = 's'
            self.time_factor = 1.0
        elif self.total_sequence_time < 7200:
            self.time_unit = 'min'
            self.time_factor = 1.0 / 60.0
        else:
            self.time_unit = 'hr'
            self.time_factor = 1.0 / 3600.0

        self._prepare_live_plot()
        self.notebook.select(self.live_plot_tab)

        self.stop_event.clear()
        self.sequence_is_running = True
        self.sequence_start_time = time.perf_counter()

        with self.state_lock:
            self.current_step_num = 0
            self.last_plot_point = (0, 0)

        self.sequence_thread = threading.Thread(target=self._sequence_worker, args=(self.sequence_data[:],),
                                                daemon=True)
        self.sequence_thread.start()

        self.run_seq_btn.config(state=tk.DISABLED);
        self.stop_seq_btn.config(state=tk.NORMAL)

        self.after(self.UPDATE_INTERVAL_MS, self._periodic_updater)

    def _sequence_worker(self, sequence):
        self._log("Sequence started...")
        self.progress_bar['maximum'] = self.total_sequence_time if self.total_sequence_time > 0 else 1

        pc = 0
        cycle_counters = {}

        # This is now the source of truth for the worker's current RPM state
        current_rpm = 0.0

        while pc < len(sequence) and not self.stop_event.is_set():
            with self.state_lock:
                self.current_step_num = pc + 1

            instruction = sequence[pc]
            if instruction['type'] == 'Phase':
                # The worker now tracks its own RPM state internally
                current_rpm = self._execute_phase(instruction, current_rpm)
                pc += 1
            elif instruction['type'] == 'Cycle':
                start_idx, end_idx = instruction['start_phase'] - 1, instruction['end_phase'] - 1
                if not (0 <= start_idx <= end_idx < len(sequence)):
                    self._log(f"Error: Invalid phase range in Cycle at step {pc + 1}. Aborting.");
                    break
                if pc not in cycle_counters: cycle_counters[pc] = instruction['repeats']
                if cycle_counters[pc] > 0:
                    self._log(
                        f"Cycle at step {pc + 1}: {cycle_counters[pc]} repeats left. Jumping to step {start_idx + 1}.")
                    cycle_counters[pc] -= 1;
                    pc = start_idx
                else:
                    self._log(f"Cycle at step {pc + 1} finished. Continuing.")
                    del cycle_counters[pc];
                    pc += 1

        if not self.stop_event.is_set():
            self._log("Sequence finished. Stopping pump.")
            self.pump_controller.stop()

        self.after(0, self._on_sequence_finish)

    def _execute_phase(self, phase, start_rpm):
        if self.stop_event.is_set(): return start_rpm

        if phase['direction'] == 'Forward':
            self.pump_controller.start_forward()
        else:
            self.pump_controller.start_backward()

        duration_s = phase['duration']
        if phase['unit'] == 'min':
            duration_s *= 60
        elif phase['unit'] == 'hr':
            duration_s *= 3600

        target_rpm = phase['rpm']
        phase_start_time = time.perf_counter()

        if phase['mode'] == 'Fixed':
            self.pump_controller.set_speed(target_rpm, wait=False)
            end_time = phase_start_time + duration_s
            while time.perf_counter() < end_time:
                if self.stop_event.is_set():
                    break
                self.stop_event.wait(0.05)

        elif phase['mode'] == 'Ramp':
            cmd_interval = self.pump_controller.command_interval
            num_steps = int(duration_s / cmd_interval) if duration_s > 0 else 1
            if num_steps == 0:
                num_steps = 1
            step_time = duration_s / num_steps
            for i in range(1, num_steps + 1):
                target_time = phase_start_time + i * step_time
                ramp_fraction = i / num_steps
                current_rpm_in_phase = start_rpm + (target_rpm - start_rpm) * ramp_fraction
                wait_time = target_time - time.perf_counter()
                if wait_time > 0:
                    if self.stop_event.wait(wait_time):
                        break
                if self.stop_event.is_set():
                    break
                self.pump_controller.set_speed(current_rpm_in_phase, wait=False)

        if not self.stop_event.is_set():
            # Ensure total phase duration is respected
            remaining = (phase_start_time + duration_s) - time.perf_counter()
            if remaining > 0:
                self.stop_event.wait(remaining)
            return target_rpm
        else:
            final_elapsed = time.perf_counter() - phase_start_time
            if phase['mode'] == 'Ramp' and duration_s > 0:
                ramp_fraction = min(1.0, final_elapsed / duration_s)
                return start_rpm + (target_rpm - start_rpm) * ramp_fraction
            else:
                return target_rpm

    def _periodic_updater(self):
        if not self.sequence_is_running:
            return

        elapsed_time = time.perf_counter() - self.sequence_start_time

        # Ensure plot doesn't run past the end time
        if elapsed_time >= self.total_sequence_time:
            elapsed_time = self.total_sequence_time

        # Find the correct direction from the pre-calculated plan
        planned_direction = "Forward"  # Default
        if hasattr(self, 'plan_directions') and self.plan_directions:
            for phase_info in self.plan_directions:
                if elapsed_time <= phase_info['end_time'] + 0.001:  # Add tolerance for float comparison
                    planned_direction = phase_info['direction']
                    break

        with self.state_lock:
            step_num = self.current_step_num

        if self.plan_time_points:
            planned_rpm = np.interp(elapsed_time, self.plan_time_points, self.plan_rpm_points)
        else:
            planned_rpm = 0

        self._update_progress(step_num, elapsed_time, self.total_sequence_time)
        self._update_live_plot(elapsed_time * self.time_factor, planned_rpm, planned_direction)

        # Continue updating as long as the sequence is marked as running
        if self.sequence_is_running:
            self.after(self.UPDATE_INTERVAL_MS, self._periodic_updater)

    def _update_progress(self, current_step_num, elapsed_time, total_time):
        self.progress_bar['value'] = elapsed_time
        et_m, et_s = divmod(int(elapsed_time), 60)
        tt_m, tt_s = divmod(int(total_time), 60)

        if current_step_num > len(self.sequence_data): current_step_num = len(self.sequence_data)

        self.progress_step_label.config(text=f"Current Step: {current_step_num}/{len(self.sequence_data)}")
        self.progress_time_label.config(text=f"Time: {et_m:02d}:{et_s:02d} / {tt_m:02d}:{tt_s:02d}")

    def _on_sequence_finish(self):
        if not self.sequence_is_running: return  # Prevent double-calls
        self.sequence_is_running = False

        self.run_seq_btn.config(state=tk.NORMAL);
        self.stop_seq_btn.config(state=tk.DISABLED)

        # Final plot update to ensure it reaches the very end
        if hasattr(self, 'plan_directions') and self.plan_directions:
            final_rpm = self.plan_rpm_points[-1] if self.plan_rpm_points else 0
            final_direction = self.plan_directions[-1]['direction'] if self.plan_directions else "Forward"
            self._update_live_plot(self.total_sequence_time * self.time_factor, final_rpm, final_direction)

        self._update_progress(len(self.sequence_data), self.total_sequence_time, self.total_sequence_time)
        if self.live_ax: self.live_canvas.draw_idle()

    def _stop_sequence(self):
        if self.sequence_thread and self.sequence_thread.is_alive():
            self.stop_event.set()
            # The worker will call _on_sequence_finish when it exits
            self._log("Stop requested by user.")

    def _save_sequence(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Sequence Files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.sequence_data, f, indent=4)
            self._log(f"Sequence saved to: {filepath}")
        except Exception as e:
            messagebox.showerror("Save Failed", f"Could not save file: {e}")

    def _load_sequence(self):
        filepath = filedialog.askopenfilename(filetypes=[("JSON Sequence Files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                self.sequence_data = json.load(f)
            self._update_treeview()
            self._log(f"Sequence loaded from: {filepath}")
        except Exception as e:
            messagebox.showerror("Load Failed", f"Could not load file: {e}")

    def _on_closing(self):
        self.stop_event.set()
        self.sequence_is_running = False
        if self.pump_controller and self.pump_controller.is_connected:
            if messagebox.askyesno("Exit", "The pump is still connected. Do you want to disconnect before exiting?"):
                self._disconnect_pump()
        self.destroy()

    def __init__(self):
        super().__init__()
        self._setup_theme_and_style()
        self._initialize_variables()
        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_theme_and_style(self):
        if sv_ttk:
            sv_ttk.set_theme("light")
        self.title("MINIPULS 3 Controller")
        self.geometry("1050x810")
        try:
            self.iconbitmap(resource_path("minipuls3_icon.ico"))
        except Exception as e:
            print(f"Could not set runtime icon: {e}")

        self.bold_font = font.Font(family="Segoe UI", size=10, weight="bold")
        self.status_font = font.Font(family="Segoe UI", size=9, weight="bold")
        self.style = ttk.Style()
        self.style.configure("TLabelframe.Label", font=self.bold_font)
        self.style.configure("Stop.TButton", font=self.bold_font)
        self.style.configure("Run.TButton", font=self.bold_font)

    def _initialize_variables(self):
        self.pump_controller = None
        self.sequence_thread = None
        self.stop_event = threading.Event()
        self.sequence_data = []
        self.debug_mode_var = tk.BooleanVar(value=False)

        self.state_lock = threading.Lock()
        self.sequence_is_running = False
        self.sequence_start_time = 0.0
        self.total_sequence_time = 0.0
        self.current_step_num = 0

        self.plan_time_points = []
        self.plan_rpm_points = []
        self.plan_directions = []

        self.time_unit = 'min'
        self.time_factor = 1.0 / 60.0

        self.live_fig = None
        self.live_ax = None
        self.live_canvas = None
        self.last_plot_point = (0, 0)

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        left_frame = ttk.Frame(main_frame, width=300)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10));
        left_frame.pack_propagate(False)
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self._create_connection_panel(left_frame)
        self._create_manual_control_panel(left_frame)
        self._create_file_panel(left_frame)
        self._create_execution_panel(left_frame)
        self._create_contact_panel(left_frame)

        paned_window = ttk.PanedWindow(right_frame, orient=tk.VERTICAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        notebook_frame = ttk.Frame(paned_window)
        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        editor_tab = ttk.Frame(self.notebook, padding=0)
        plot_tab = ttk.Frame(self.notebook, padding=0)
        self.live_plot_tab = plot_tab

        self.notebook.add(editor_tab, text="Sequence Editor")
        self.notebook.add(plot_tab, text="Live Visualization")

        self._create_sequence_editor(editor_tab)
        self._create_live_plot_panel(plot_tab)

        log_frame = ttk.Frame(paned_window)
        self._create_log_panel(log_frame)

        paned_window.add(notebook_frame, weight=3)
        paned_window.add(log_frame, weight=1)

    def _create_connection_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="1. Connection & Status", padding="10")
        frame.pack(fill=tk.X, pady=5, anchor='n')
        ttk.Label(frame, text="COM Port:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.com_port_entry = ttk.Entry(frame, width=15)
        self.com_port_entry.insert(0, "COM4")
        self.com_port_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(frame, text="Unit ID:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.unit_id_entry = ttk.Entry(frame, width=15)
        self.unit_id_entry.insert(0, "30")
        self.unit_id_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(frame, text="Cmd Interval (s):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        interval_frame = ttk.Frame(frame)
        interval_frame.grid(row=2, column=1, sticky="ew")
        self.command_interval_entry = ttk.Entry(interval_frame, width=8)
        self.command_interval_entry.insert(0, "0.2")
        self.command_interval_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.set_interval_btn = ttk.Button(interval_frame, text="Set", command=self._set_command_interval,
                                           state=tk.DISABLED, width=5)
        self.set_interval_btn.pack(side=tk.LEFT, padx=(5, 0))

        self.debug_mode_check = ttk.Checkbutton(frame, text="Debug Mode (No Pump)", variable=self.debug_mode_var)
        self.debug_mode_check.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=(5, 5))

        self.connect_btn = ttk.Button(frame, text="✔ Connect", command=self._connect_pump, style='Accent.TButton')
        self.connect_btn.grid(row=4, column=0, pady=5, padx=5, sticky="ew")
        self.disconnect_btn = ttk.Button(frame, text="✖ Disconnect", command=self._disconnect_pump, state=tk.DISABLED)
        self.disconnect_btn.grid(row=4, column=1, pady=5, padx=5, sticky="ew")

        self.status_label = ttk.Label(frame, text="Status: Disconnected", foreground="#c0392b", font=self.status_font)
        self.status_label.grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

    def _create_manual_control_panel(self, parent):
        self.manual_frame = ttk.LabelFrame(parent, text="2. Manual Control", padding="10")
        self.manual_frame.pack(fill=tk.X, pady=5, anchor='n')
        speed_frame = ttk.Frame(self.manual_frame)
        speed_frame.pack(fill=tk.X)
        ttk.Label(speed_frame, text="Speed (RPM):").pack(side=tk.LEFT, anchor='w')
        self.manual_rpm_entry = ttk.Entry(speed_frame, width=6)
        self.manual_rpm_entry.pack(side=tk.RIGHT, anchor='e')
        self.manual_rpm_entry.bind("<Return>", self._update_speed_from_entry)
        self.manual_rpm_entry.bind("<FocusOut>", self._update_speed_from_entry)
        self.speed_scale = ttk.Scale(self.manual_frame, from_=0, to=48, orient=tk.HORIZONTAL,
                                     command=self._update_speed_label)
        self.speed_scale.pack(fill=tk.X, pady=2)
        self.speed_label = ttk.Label(self.manual_frame, text="0.0 RPM")
        self.speed_label.pack(anchor=tk.W)
        btn_frame = ttk.Frame(self.manual_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 5))
        self.fwd_btn = ttk.Button(btn_frame, text="▶ Forward", command=self._manual_start_fwd)
        self.fwd_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.rev_btn = ttk.Button(btn_frame, text="◀ Backward", command=self._manual_start_rev)
        self.rev_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.stop_btn = ttk.Button(self.manual_frame, text="⏹️ STOP", command=self._manual_stop, style="Stop.TButton")
        self.stop_btn.pack(fill=tk.X, pady=(5, 0), ipady=5)
        self._set_manual_controls_state(tk.DISABLED)

    def _create_file_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Sequence Operations", padding="10")
        frame.pack(fill=tk.X, pady=5, anchor='n')
        self.save_btn = ttk.Button(frame, text="💾 Save", command=self._save_sequence)
        self.save_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.load_btn = ttk.Button(frame, text="📂 Load", command=self._load_sequence)
        self.load_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

    def _create_execution_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Sequence Execution", padding="10")
        frame.pack(fill=tk.X, pady=5, anchor='n')
        self.run_seq_btn = ttk.Button(frame, text="▶️ Run Sequence", command=self._run_sequence, state=tk.DISABLED,
                                      style="Run.TButton")
        self.run_seq_btn.pack(fill=tk.X, ipady=5, pady=2)
        self.stop_seq_btn = ttk.Button(frame, text="⏹️ Stop Sequence", command=self._stop_sequence, state=tk.DISABLED,
                                       style="Stop.TButton")
        self.stop_seq_btn.pack(fill=tk.X, ipady=5, pady=2)
        progress_frame = ttk.Frame(frame, padding=(0, 5))
        progress_frame.pack(fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(5, 2))
        self.progress_step_label = ttk.Label(progress_frame, text="Current Step: -")
        self.progress_step_label.pack(side=tk.LEFT)
        self.progress_time_label = ttk.Label(progress_frame, text="Time: 00:00 / 00:00")
        self.progress_time_label.pack(side=tk.RIGHT)

    def _create_contact_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Contact", padding="10")
        frame.pack(fill=tk.X, pady=5, anchor='n')
        contact_widget = tk.Text(frame, height=2, wrap=tk.WORD, relief=tk.FLAT, font=("Segoe UI", 9))
        try:
            contact_widget.config(bg=self.style.lookup('TFrame', 'background'))
        except tk.TclError:
            pass
        contact_widget.insert(tk.END, "Any problem or suggestion, contact:\n")
        contact_widget.insert(tk.END, "qiyaolin3776@gmail.com")
        contact_widget.tag_add("email", "2.0", "2.end")
        contact_widget.tag_config("email", foreground="blue", underline=True)
        contact_widget.config(state=tk.DISABLED)
        contact_widget.pack(fill=tk.X)

    def _create_sequence_editor(self, parent):
        frame = ttk.LabelFrame(parent, text="Sequence Editor", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, pady=0, padx=0)
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        cols = ("#", "Type", "Details", "Duration")
        self.sequence_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col, width in zip(cols, [40, 80, 400, 120]):
            self.sequence_tree.heading(col, text=col)
            self.sequence_tree.column(col, width=width, anchor=tk.W if col == "Details" else tk.CENTER)
        self.sequence_tree.bind("<Double-1>", self._edit_item)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.sequence_tree.yview)
        self.sequence_tree.configure(yscrollcommand=vsb.set);
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.sequence_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        editor_controls = ttk.Frame(frame, padding=(0, 10, 0, 0))
        editor_controls.pack(fill=tk.X)
        ttk.Button(editor_controls, text="➕ Add Phase", command=self._add_phase, style='Accent.TButton').pack(
            side=tk.LEFT, padx=2)
        ttk.Button(editor_controls, text="🔄 Add Cycle", command=self._add_cycle, style='Accent.TButton').pack(
            side=tk.LEFT, padx=2)
        ttk.Button(editor_controls, text="🗑️ Remove", command=self._remove_item).pack(side=tk.LEFT, padx=(10, 2))
        ttk.Button(editor_controls, text="❌ Clear All", command=self._clear_sequence).pack(side=tk.LEFT, padx=2)
        ttk.Button(editor_controls, text="⬆️ Move Up", command=lambda: self._move_item('up')).pack(side=tk.LEFT,
                                                                                                   padx=(10, 2))
        ttk.Button(editor_controls, text="⬇️ Move Down", command=lambda: self._move_item('down')).pack(side=tk.LEFT,
                                                                                                       padx=2)

    def _create_live_plot_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Live Process Visualization", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, pady=0, padx=0)

        if not (Figure and mlines): return

        self.live_fig = Figure(figsize=(5, 4), dpi=100)
        self.live_ax = self.live_fig.add_subplot(111)

        self.live_canvas = FigureCanvasTkAgg(self.live_fig, master=frame)
        self.live_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self._prepare_live_plot()

    def _prepare_live_plot(self):
        if not self.live_ax: return

        self.live_ax.clear()

        # Use the pre-calculated plan if available, otherwise calculate it
        if hasattr(self, 'plan_time_points') and self.plan_time_points:
            plot_data = {'time_points': self.plan_time_points, 'rpm_points': self.plan_rpm_points,
                         'cycle_spans': self.plot_data['cycle_spans']}
        else:
            plot_data = self._get_expanded_sequence_data(self.sequence_data)

        time_scaled = [t * self.time_factor for t in plot_data['time_points']]
        rpm_data = plot_data['rpm_points']

        self.live_ax.plot(time_scaled, rpm_data, color='lightgray', linestyle='--', label='Plan', zorder=1)

        for span in plot_data.get('cycle_spans', []):
            self.live_ax.axvspan(span[0] * self.time_factor, span[1] * self.time_factor, color='lightskyblue',
                                 alpha=0.3, zorder=0)

        self.live_ax.set_title("Real-time Sequence Monitoring", fontsize=14, fontweight='bold')
        self.live_ax.set_xlabel(f"Time ({self.time_unit})", fontsize=10)
        self.live_ax.set_ylabel("Speed (RPM)", fontsize=10)
        self.live_ax.set_ylim(bottom=-2, top=52)
        if time_scaled:
            self.live_ax.set_xlim(left=min(time_scaled) - 0.05 * max(time_scaled), right=max(time_scaled) * 1.05)

        plan_line = mlines.Line2D([], [], color='lightgray', linestyle='--', label='Plan')
        fwd_line = mlines.Line2D([], [], color=self.COLOR_FORWARD, label='Forward')
        bwd_line = mlines.Line2D([], [], color=self.COLOR_BACKWARD, label='Backward')
        self.live_ax.legend(handles=[plan_line, fwd_line, bwd_line])

        if not sns:
            self.live_ax.grid(True, linestyle='--', alpha=0.6)

        self.live_fig.tight_layout(pad=1.5)
        self.live_canvas.draw()

    def _update_live_plot(self, time_scaled, rpm, direction):
        if not self.live_ax: return

        color = self.COLOR_FORWARD if direction == 'Forward' else self.COLOR_BACKWARD
        new_point = (time_scaled, rpm)

        self.live_ax.plot(
            [self.last_plot_point[0], new_point[0]],
            [self.last_plot_point[1], new_point[1]],
            color=color,
            linewidth=2.5,
            solid_capstyle='round'
        )
        self.last_plot_point = new_point
        self.live_canvas.draw_idle()

    def _create_log_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Status & Log", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.log_text = tk.Text(frame, height=10, state=tk.DISABLED, wrap=tk.WORD)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=vsb.set);
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    if np and plt and Figure and FigureCanvasTkAgg:
        plt.ion()
    app = PumpControlUI()
    app.mainloop()