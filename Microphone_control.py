import tkinter as tk
from tkinter import ttk
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER

# --- Microphone control ---
def get_microphone_volume_control():
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return volume

mic_volume = get_microphone_volume_control()
mic_active = False
toggle_mode = False
selected_ptt_key = "space"
pressed_keys = set()

# --- Microphone status functions ---
def activate_mic():
    global mic_active
    if not mic_active:
        mic_volume.SetMute(0, None)
        status_label.config(text="Mic: ACTIVE", foreground="green")
        mic_active = True

def deactivate_mic():
    global mic_active
    if mic_active:
        mic_volume.SetMute(1, None)
        status_label.config(text="Mic: MUTED", foreground="red")
        mic_active = False

def toggle_mic():
    if mic_active:
        deactivate_mic()
    else:
        activate_mic()

# --- Key Handling ---
def on_key_press(event):
    key = event.keysym
    if mode_var.get() == "hold":
        if key == selected_ptt_key and key not in pressed_keys:
            pressed_keys.add(key)
            activate_mic()
    elif mode_var.get() == "toggle":
        if key == selected_ptt_key:
            toggle_mic()

def on_key_release(event):
    key = event.keysym
    if mode_var.get() == "hold":
        if key == selected_ptt_key and key in pressed_keys:
            pressed_keys.remove(key)
            deactivate_mic()

# --- Update Selected Key ---
def update_selected_key(var, label):
    global selected_ptt_key
    selected_ptt_key = var.get()
    label.config(text=f"Current PTT Key: {selected_ptt_key.upper()}")
    deactivate_mic()

# --- GUI Setup ---
root = tk.Tk()
root.title("Push-to-Talk Microphone Controller")
root.geometry("420x520")

main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill=tk.BOTH, expand=True)

# Title
ttk.Label(main_frame, text="Push-to-Talk Microphone Controller", font=("Arial", 14, "bold"))\
    .grid(row=0, column=0, columnspan=2, pady=(0, 20))

# --- Key Settings ---
settings_frame = ttk.LabelFrame(main_frame, text="Key Settings", padding="15")
settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

ttk.Label(settings_frame, text="Select PTT Key:", font=("Arial", 10, "bold"))\
    .grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

key_var = tk.StringVar(value="space")
key_options = [
    ("Space Bar", "space"),
    ("Left Ctrl", "Control_L"),
    ("Right Ctrl", "Control_R"),
    ("Left Alt", "Alt_L"),
    ("Right Alt", "Alt_R"),
    ("Tab", "Tab"),
    ("F1", "F1"),
    ("F2", "F2"),
    ("F3", "F3"),
    ("F4", "F4"),
]

for i, (display_name, key_value) in enumerate(key_options):
    row = i // 2 + 1
    col = i % 2
    ttk.Radiobutton(settings_frame, text=display_name, variable=key_var,
                    value=key_value).grid(row=row, column=col, sticky=tk.W, padx=(0, 20), pady=2)

# Apply + Status
current_key_label = ttk.Label(settings_frame, text="Current PTT Key: SPACE", font=("Arial", 9, "bold"), foreground="blue")
current_key_label.grid(row=6, column=0, columnspan=2, pady=(15, 5))

ttk.Button(settings_frame, text="Apply Key Setting",
           command=lambda: update_selected_key(key_var, current_key_label))\
    .grid(row=7, column=0, pady=(5, 0))

# --- Mode Select: Hold or Toggle ---
ttk.Label(settings_frame, text="Mic Activation Mode:", font=("Arial", 10, "bold"))\
    .grid(row=8, column=0, sticky=tk.W, pady=(15, 5))

mode_var = tk.StringVar(value="hold")
ttk.Radiobutton(settings_frame, text="Hold to Talk", variable=mode_var, value="hold")\
    .grid(row=9, column=0, sticky=tk.W)
ttk.Radiobutton(settings_frame, text="Toggle On/Off", variable=mode_var, value="toggle")\
    .grid(row=9, column=1, sticky=tk.W)

# --- Status Display ---
status_frame = ttk.Frame(main_frame)
status_frame.grid(row=2, column=0, columnspan=2, pady=15)

ttk.Label(status_frame, text="Status:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
status_label = ttk.Label(status_frame, text="READY TO START", font=("Arial", 12, "bold"), foreground="blue")
status_label.grid(row=1, column=0, pady=(5, 0))

# --- Key Bindings ---
root.bind("<KeyPress>", on_key_press)
root.bind("<KeyRelease>", on_key_release)
root.focus_set()
deactivate_mic()

root.mainloop()
