import tkinter as tk
from tkinter import messagebox
import json
from ds1000z import DS1000Z  

def run_connection_test():
 
    target_ip = "Unknown"
    
    try:
       
        with open('config.json', 'r') as f:
            config = json.load(f)
        
        target_ip = config.get('oscilloscope_ip', 'Unknown IP')
        print(f"Attempting connection to {target_ip}...")
        
        
        instrument = DS1000Z(target_ip)
        instrument.reset()
        
        messagebox.showinfo("Connection Success", f"Successfully verified link to {target_ip}!")
        
    except FileNotFoundError:
        messagebox.showerror("Config Error", "config.json not found!")
        
    except Exception as e:
        
        print(f"Error details: {e}")
        messagebox.showerror("Connection Failed", 
            f"Could not reach device at {target_ip}.\n\n"
            f"Check Ethernet cable and ensure device is powered on.\n"
            f"System Error: {e}"
        )


root = tk.Tk()
root.title("Panasonic Test Controller")
root.geometry("450x250")


header_label = tk.Label(root, text="DCDC Test Controller", font=("Segoe UI", 14, "bold"))
header_label.pack(pady=20)


status_label = tk.Label(root, text="System Ready", fg="gray")
status_label.pack()


connect_btn = tk.Button(root, text="Initialize Hardware", command=run_connection_test, height=2, width=25, bg="#e1e1e1")
connect_btn.pack(pady=15)

if __name__ == "__main__":
    root.mainloop()