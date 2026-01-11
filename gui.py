"""
Panasonic DC/DC Test Dashboard
------------------------------
A Tkinter-based control panel for the DS1000Z oscilloscope.
Provides safety interlocks, input validation, and test session logging
for critical facilities testing.
"""

import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import json
import datetime
from ds1000z import DS1000Z

# Global instrument reference
scope = None

def log_message(message):
    """Appends a timestamped entry to the scrolling log console."""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_window.configure(state='normal')
    log_window.insert(tk.END, f"[{timestamp}] {message}\n")
    log_window.see(tk.END)
    log_window.configure(state='disabled')

def load_config():
    """Initializes the UI with network settings from config.json."""
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
        ip_entry.insert(0, config.get('oscilloscope_ip', ''))
        log_message("System Initialized. Configuration loaded.")
    except FileNotFoundError:
        log_message("System Alert: config.json not found.")

def connect_to_scope():
    """Establishes connection to hardware and unlocks control interface."""
    global scope
    target_ip = ip_entry.get()
    
    log_message(f"Initiating handshake with {target_ip}...")
    
    try:
        scope = DS1000Z(target_ip)
        scope.reset()
        
        # Update Interface State
        status_indicator.config(bg="#4CAF50")
        status_label.config(text="CONNECTED", fg="#4CAF50")
        run_btn.config(state="normal")
        stop_btn.config(state="normal")
        log_message("Connection established. Hardware reset complete.")
        
    except Exception as e:
        status_indicator.config(bg="#F44336")
        status_label.config(text="DISCONNECTED", fg="#F44336")
        log_message(f"Handshake Failed: {e}")
        messagebox.showerror("Connection Error", str(e))

def run_test_sequence():
    """Validates inputs and executes the defined waveform parameters."""
    if not scope:
        messagebox.showerror("Hardware Error", "No instrument connected.")
        return

    # Compliance Check
    tech_name = tech_entry.get().strip()
    if not tech_name:
        messagebox.showwarning("Compliance Error", "Technician Name is required to proceed.")
        return

    # Parameter Validation
    try:
        volts = float(volts_entry.get())
        freq = float(freq_entry.get())
    except ValueError:
        messagebox.showerror("Input Error", "Voltage and Frequency must be numeric values.")
        return

    # Safety Interlock
    if volts > 20:
        log_message("SAFETY INTERLOCK: Voltage > 20V rejected.")
        return

    log_message(f"--- SESSION STARTED: {tech_name.upper()} ---")
    log_message(f"Setting Source: {volts}V @ {freq}Hz")
    
    try:
        scope.set_source_amplitude(volts)
        scope.set_source_frequency(freq)
        scope.enable_source()
        log_message("Output Active. Test sequence running.")
        log_message("--- SEQUENCE COMPLETE ---")
    except Exception as e:
        log_message(f"Command Error: {e}")

def disable_output():
    """Emergency Stop: Resets instrument state immediately."""
    if not scope:
        return
    try:
        scope.reset()
        log_message("SAFETY STOP: Output disabled / Instrument Reset.")
    except Exception as e:
        log_message(f"Stop Failed: {e}")

def save_report():
    """Exports the current session log to a text file."""
    tech_name = tech_entry.get().strip() or "Unknown"
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    default_filename = f"TestReport_{tech_name}_{timestamp}.txt"
    
    filepath = filedialog.asksaveasfilename(
        defaultextension=".txt",
        initialfile=default_filename,
        title="Export Test Report"
    )
    
    if filepath:
        try:
            with open(filepath, "w") as f:
                f.write(log_window.get("1.0", tk.END))
            log_message(f"Report exported: {filepath}")
            messagebox.showinfo("Export Complete", "Log file saved successfully.")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

# --- Main Interface Construction ---
root = tk.Tk()
root.title("Panasonic DC/DC Diagnostic | Test Bench")
root.geometry("600x650")

# Header
header_frame = tk.Frame(root, pady=15)
header_frame.pack()
tk.Label(header_frame, text="DC/DC Converter Diagnostic Tool", font=("Segoe UI", 14, "bold")).pack()

# Session Metadata
session_frame = tk.Frame(root, pady=5)
session_frame.pack()
tk.Label(session_frame, text="Technician:", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=5)
tech_entry = tk.Entry(session_frame, width=25)
tech_entry.pack(side=tk.LEFT, padx=5)

# Connectivity Controls
conn_frame = tk.LabelFrame(root, text="Network Configuration", padx=15, pady=15)
conn_frame.pack(fill="x", padx=15, pady=10)

tk.Label(conn_frame, text="Instrument IP:").grid(row=0, column=0, padx=5)
ip_entry = tk.Entry(conn_frame, width=18)
ip_entry.grid(row=0, column=1, padx=5)

connect_btn = tk.Button(conn_frame, text="Initialize", command=connect_to_scope, bg="#E0E0E0", width=12)
connect_btn.grid(row=0, column=2, padx=10)

status_indicator = tk.Frame(conn_frame, width=16, height=16, bg="gray")
status_indicator.grid(row=0, column=3, padx=5)
status_label = tk.Label(conn_frame, text="Offline", fg="gray", font=("Segoe UI", 9, "bold"))
status_label.grid(row=0, column=4)

# Test Parameters
param_frame = tk.LabelFrame(root, text="Signal Parameters", padx=15, pady=15)
param_frame.pack(fill="x", padx=15, pady=5)

tk.Label(param_frame, text="Amplitude (V):").grid(row=0, column=0, padx=5)
volts_entry = tk.Entry(param_frame, width=10)
volts_entry.insert(0, "5.0")
volts_entry.grid(row=0, column=1, padx=5)

tk.Label(param_frame, text="Frequency (Hz):").grid(row=0, column=2, padx=5)
freq_entry = tk.Entry(param_frame, width=10)
freq_entry.insert(0, "50000")
freq_entry.grid(row=0, column=3, padx=5)

# Execution Controls
action_frame = tk.Frame(root, pady=15)
action_frame.pack()

run_btn = tk.Button(action_frame, text="EXECUTE TEST", command=run_test_sequence, 
                    font=("Segoe UI", 10, "bold"), bg="#4CAF50", fg="white", state="disabled", width=18)
run_btn.pack(side=tk.LEFT, padx=10)

stop_btn = tk.Button(action_frame, text="EMERGENCY STOP", command=disable_output, 
                    font=("Segoe UI", 10, "bold"), bg="#F44336", fg="white", state="disabled", width=18)
stop_btn.pack(side=tk.LEFT, padx=10)

# System Log
log_frame = tk.LabelFrame(root, text="Operations Log", padx=5, pady=5)
log_frame.pack(fill="both", expand=True, padx=15, pady=10)
log_window = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', font=("Consolas", 9))
log_window.pack(fill="both", expand=True)

# Footer
footer_frame = tk.Frame(root, pady=10)
footer_frame.pack()
save_btn = tk.Button(footer_frame, text="Export Session Log", command=save_report, width=20)
save_btn.pack()

if __name__ == "__main__":
    load_config()
    root.mainloop()