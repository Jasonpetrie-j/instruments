import tkinter as tk
from tkinter import messagebox, scrolledtext
import json
from ds1000z import DS1000Z


scope = None

def log_message(message):
    """Adds a timestamped message to the scrolling log window."""
    log_window.configure(state='normal')  
    log_window.insert(tk.END, f"> {message}\n")
    log_window.see(tk.END) 
    log_window.configure(state='disabled') # Freeze again so user can't type in it

def load_config():
    """Loads IP from JSON and puts it into the UI Entry box."""
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
        ip_entry.insert(0, config.get('oscilloscope_ip', ''))
        log_message("Config loaded successfully.")
    except FileNotFoundError:
        log_message("Error: config.json not found.")

def connect_to_scope():
    """Tries to connect using the IP currently typed in the box."""
    global scope
    target_ip = ip_entry.get()
    
    log_message(f"Attempting connection to {target_ip}...")
    
    try:
        # The Real Driver Call
        scope = DS1000Z(target_ip)
        scope.reset()
        
        # Update UI to show success
        status_indicator.config(bg="#4CAF50") # Green
        status_label.config(text="CONNECTED", fg="#4CAF50")
        run_btn.config(state="normal") # Enable the Run button
        log_message("Success: Hardware connected and reset.")
        
    except Exception as e:
        status_indicator.config(bg="red")
        status_label.config(text="DISCONNECTED", fg="red")
        log_message(f"Failed: {e}")
        messagebox.showerror("Connection Error", str(e))

def run_test_sequence():
    """Runs the actual test commands using values from the UI."""
    if not scope:
        messagebox.showerror("Error", "No Hardware Connected!")
        return

    # 1. Get Values from UI
    try:
        volts = float(volts_entry.get())
        freq = float(freq_entry.get())
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numbers for Voltage/Freq")
        return

    # 2. Validation (The "Guardrails")
    if volts > 20:
        log_message("Error: Voltage too high! Safety limit is 20V.")
        return

    # 3. Execute Commands
    log_message(f"Configuring Source: {volts}V at {freq}Hz...")
    try:
        scope.set_source_amplitude(volts)
        scope.set_source_frequency(freq)
        scope.enable_source()
        log_message("Test Running. Waveform output active.")
    except Exception as e:
        log_message(f"Command Failed: {e}")

# --- GUI SETUP ---
root = tk.Tk()
root.title("Panasonic DC/DC Test Dashboard")
root.geometry("600x500")

# 1. HEADER
header_frame = tk.Frame(root, pady=10)
header_frame.pack()
tk.Label(header_frame, text="Test Bench", font=("Arial", 16, "bold")).pack()

# 2. CONNECTION SECTION
conn_frame = tk.LabelFrame(root, text="Hardware Connection", padx=10, pady=10)
conn_frame.pack(fill="x", padx=10, pady=5)

tk.Label(conn_frame, text="Oscilloscope IP:").grid(row=0, column=0, padx=5)
ip_entry = tk.Entry(conn_frame, width=20)
ip_entry.grid(row=0, column=1, padx=5)

connect_btn = tk.Button(conn_frame, text="Connect", command=connect_to_scope, bg="#E0E0E0")
connect_btn.grid(row=0, column=2, padx=10)

# Status Light (A small square frame)
status_indicator = tk.Frame(conn_frame, width=20, height=20, bg="gray")
status_indicator.grid(row=0, column=3, padx=5)
status_label = tk.Label(conn_frame, text="Not Connected", fg="gray", font=("Arial", 9, "bold"))
status_label.grid(row=0, column=4)

# 3. PARAMETERS SECTION
param_frame = tk.LabelFrame(root, text="Test Parameters", padx=10, pady=10)
param_frame.pack(fill="x", padx=10, pady=5)

tk.Label(param_frame, text="Amplitude (Volts):").grid(row=0, column=0, padx=5)
volts_entry = tk.Entry(param_frame, width=10)
volts_entry.insert(0, "5.0") # Default value
volts_entry.grid(row=0, column=1, padx=5)

tk.Label(param_frame, text="Frequency (Hz):").grid(row=0, column=2, padx=5)
freq_entry = tk.Entry(param_frame, width=10)
freq_entry.insert(0, "50000") # Default value
freq_entry.grid(row=0, column=3, padx=5)

# 4. ACTION SECTION
action_frame = tk.Frame(root, pady=10)
action_frame.pack()
# State='disabled' prevents clicking until connected
run_btn = tk.Button(action_frame, text="RUN DIAGNOSTIC", command=run_test_sequence, 
                    font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", state="disabled")
run_btn.pack(ipadx=20, ipady=5)

# 5. LOG CONSOLE
log_frame = tk.LabelFrame(root, text="System Log", padx=5, pady=5)
log_frame.pack(fill="both", expand=True, padx=10, pady=10)
log_window = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', font=("Consolas", 9))
log_window.pack(fill="both", expand=True)

# Start Up
load_config() # Auto-fill the IP box when app starts
root.mainloop()