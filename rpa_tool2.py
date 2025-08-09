"""
RPA Dynamic Flow Builder & Runner - single-file Tkinter app
Features:
- Create flows composed of steps: Click, Type, Wait, OCR_Check, ExcelLoop
- Record screen regions by clicking and dragging (region selector overlay)
- Save / Load flows to JSON
- Execute flows using pyautogui, pytesseract, pandas
- Logs shown in GUI

Dependencies:
- pyautogui
- pillow
- pytesseract
- pandas (only if using ExcelLoop)

Run:
    py -m pip install pyautogui pillow pytesseract pandas
    python rpa_dynamic_flow_tool.py

Note: Running GUI scripts that control mouse/keyboard may require you to disable failsafes or run with proper privileges.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import json
import threading
import time
import pyautogui
from PIL import ImageGrab, Image
import pytesseract
import pandas as pd
import os
import sys

pyautogui.FAILSAFE = True
# Change this path to your local tesseract if needed
pytesseract.pytesseract.tesseract_cmd = r"C:\\Users\\Poovarasan.P\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe"

# -----------------------------
# Helper utilities
# -----------------------------

def timestamp():
    return time.strftime('%Y-%m-%d %H:%M:%S')

# -----------------------------
# Execution engine
# -----------------------------
class FlowRunner:
    def __init__(self, steps, logger):
        self.steps = steps
        self.logger = logger
        self._stop = False

    def stop(self):
        self._stop = True

    def log(self, msg):
        # logger is a ScrolledText widget
        try:
            self.logger.insert(tk.END, f"[{timestamp()}] {msg}\n")
            self.logger.see(tk.END)
        except Exception:
            print(f"[{timestamp()}] {msg}")

    def run(self, context=None):
        """Execute steps sequentially.
        context: dict used for Excel rows or variables
        """
        self._stop = False
        ctx = context or {}
        try:
            for idx, step in enumerate(self.steps):
                if self._stop:
                    self.log('Execution stopped by user')
                    break
                act = step.get('action')
                self.log(f"Running step {idx+1}: {act}")

                if act == 'click':
                    x = int(step['x']); y = int(step['y'])
                    pyautogui.click(x, y)

                elif act == 'move':
                    x = int(step['x']); y = int(step['y'])
                    pyautogui.moveTo(x, y)

                elif act == 'type':
                    text = step.get('text','')
                    # support variable replacement like {col}
                    try:
                        text = text.format(**ctx) if ctx else text
                    except Exception as e:
                        self.log(f"Warning: failed to format text with context: {e}")
                    pyautogui.write(text, interval=0.03)

                elif act == 'hotkey':
                    keys = step.get('keys','')
                    keys_list = [k.strip() for k in keys.split('+') if k.strip()]
                    if keys_list:
                        pyautogui.hotkey(*keys_list)

                elif act == 'wait':
                    secs = float(step.get('seconds', '1'))
                    time.sleep(secs)

                elif act == 'ocr_check':
                    region = step.get('region')
                    expected = step.get('expected','')
                    if region and len(region)==4:
                        # Expecting region as bbox: (left, top, right, bottom)
                        bbox = tuple(region)
                        img = ImageGrab.grab(bbox=bbox)
                        text = pytesseract.image_to_string(img).strip()
                        self.log(f"OCR read: '{text}' (expecting '{expected}')")
                        if expected and expected not in text:
                            raise Exception(f"OCR mismatch. Expected '{expected}', got '{text}'")

                elif act == 'excel_loop':
                    # Reads Excel and loops nested steps for each row
                    path = step.get('path')
                    sheet = step.get('sheet') or 0
                    range_steps = step.get('steps') or []
                    if not os.path.exists(path):
                        raise Exception(f"Excel not found: {path}")
                    df = pd.read_excel(path, sheet_name=sheet)
                    for ridx, row in df.iterrows():
                        if self._stop:
                            break
                        # build simple context with column names as keys
                        row_ctx = {str(col): row[col] for col in df.columns}
                        self.log(f"Excel row {ridx+1} -> {row_ctx}")
                        nested = FlowRunner(range_steps, self.logger)
                        nested.run(context=row_ctx)

                else:
                    self.log(f"Unknown action: {act}")

            self.log('Flow finished')
        except Exception as e:
            self.log(f"Execution error: {e}")
            raise

# -----------------------------
# GUI - region selector overlay (fixed normalization)
# -----------------------------
class RegionSelector(tk.Toplevel):
    """Full-screen transparent overlay to let user drag a rectangle.
    Returns absolute screen coordinates in bbox format: (left, top, right, bottom)
    """
    def __init__(self, master, callback):
        super().__init__(master)
        self.callback = callback
        self.overrideredirect(True)
        self.attributes('-alpha', 0.25)
        self.attributes('-topmost', True)
        self.start_x = self.start_y = None
        self.rect_id = None

        self.canvas = tk.Canvas(self, cursor='cross')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Make the toplevel cover the entire screen
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")

        # Bind pointer events
        self.bind('<ButtonPress-1>', self.on_press)
        self.bind('<B1-Motion>', self.on_drag)
        self.bind('<ButtonRelease-1>', self.on_release)

    def on_press(self, event):
        # Use absolute pointer coords so orientation doesn't matter
        self.start_x = self.winfo_pointerx()
        self.start_y = self.winfo_pointery()
        if self.rect_id:
            try: self.canvas.delete(self.rect_id)
            except: pass
            self.rect_id = None

    def on_drag(self, event):
        cur_x = self.winfo_pointerx()
        cur_y = self.winfo_pointery()
        # Convert absolute screen coords to canvas-local coords for drawing
        x1 = self.start_x - self.winfo_rootx(); y1 = self.start_y - self.winfo_rooty()
        x2 = cur_x - self.winfo_rootx(); y2 = cur_y - self.winfo_rooty()
        if self.rect_id:
            self.canvas.coords(self.rect_id, x1, y1, x2, y2)
        else:
            self.rect_id = self.canvas.create_rectangle(x1, y1, x2, y2, outline='red', width=2)

    def on_release(self, event):
        end_x = self.winfo_pointerx()
        end_y = self.winfo_pointery()
        left = min(self.start_x, end_x)
        top = min(self.start_y, end_y)
        right = max(self.start_x, end_x)
        bottom = max(self.start_y, end_y)

        if right > left and bottom > top:
            region = [left, top, right, bottom]
            # Callback with normalized bbox (absolute screen coords)
            try:
                self.callback(region)
            finally:
                self.destroy()
        else:
            messagebox.showerror('Invalid selection', 'Please select a proper rectangular region')
            self.destroy()

# -----------------------------
# Main App
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('RPA Dynamic Flow Builder')
        self.geometry('1000x700')

        self.steps = []
        self._runner = None

        self._build_ui()
        # handle window close (Exit)
        self.protocol("WM_DELETE_WINDOW", self.on_exit)

    def _build_ui(self):
        left = ttk.Frame(self)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=8)

        ttk.Label(left, text='Flow Steps').pack()
        self.listbox = tk.Listbox(left, width=40, height=30)
        self.listbox.pack()
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

        btn_frame = ttk.Frame(left)
        btn_frame.pack(pady=6)
        ttk.Button(btn_frame, text='Add Step', command=self.add_step_dialog).grid(row=0,column=0)
        ttk.Button(btn_frame, text='Remove', command=self.remove_step).grid(row=0,column=1)
        ttk.Button(btn_frame, text='Move Up', command=lambda:self.move_step(-1)).grid(row=1,column=0)
        ttk.Button(btn_frame, text='Move Down', command=lambda:self.move_step(1)).grid(row=1,column=1)

        middle = ttk.Frame(self)
        middle.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)

        ttk.Label(middle, text='Step Details').pack()
        self.detail_frame = ttk.Frame(middle)
        self.detail_frame.pack(fill=tk.BOTH, expand=True)

        ctrl = ttk.Frame(middle)
        ctrl.pack(fill=tk.X, pady=6)
        ttk.Button(ctrl, text='Save Flow', command=self.save_flow).pack(side=tk.LEFT)
        ttk.Button(ctrl, text='Load Flow', command=self.load_flow).pack(side=tk.LEFT)
        ttk.Button(ctrl, text='Run Flow', command=self.run_flow).pack(side=tk.LEFT)
        ttk.Button(ctrl, text='Stop', command=self.stop_flow).pack(side=tk.LEFT)
        ttk.Button(ctrl, text='Exit', command=self.on_exit).pack(side=tk.LEFT)

        right = ttk.Frame(self)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        ttk.Label(right, text='Logs').pack()
        self.log_box = ScrolledText(right, width=80, height=40)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    # -----------------------------
    # Step list management
    # -----------------------------
    def add_step_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title('Add Step')
        ttk.Label(dlg, text='Action:').grid(row=0, column=0, sticky='w', padx=6, pady=6)
        action_cb = ttk.Combobox(dlg, values=['click','move','type','hotkey','wait','ocr_check','excel_loop'])
        action_cb.grid(row=0, column=1, padx=6, pady=6)
        action_cb.current(0)

        params_frame = ttk.Frame(dlg)
        params_frame.grid(row=1, column=0, columnspan=2, pady=6, padx=6)

        def on_action_change(event=None):
            for w in params_frame.winfo_children(): w.destroy()
            act = action_cb.get()
            if act in ('click','move'):
                ttk.Button(params_frame, text='Select Region', command=lambda:self.select_region_for_dlg(dlg, action_cb.get())).pack(anchor='w')
                ttk.Label(params_frame, text='Or enter X,Y:').pack(side=tk.LEFT, padx=4)
                x_ent = ttk.Entry(params_frame, width=8); x_ent.pack(side=tk.LEFT, padx=2)
                y_ent = ttk.Entry(params_frame, width=8); y_ent.pack(side=tk.LEFT, padx=2)

            elif act == 'type':
                ttk.Label(params_frame, text='Text:').pack(side=tk.LEFT, padx=4); text_ent = ttk.Entry(params_frame, width=40); text_ent.pack(side=tk.LEFT, padx=2)

            elif act == 'hotkey':
                ttk.Label(params_frame, text='Keys (ctrl+shift+s):').pack(side=tk.LEFT, padx=4); keys_ent = ttk.Entry(params_frame, width=30); keys_ent.pack(side=tk.LEFT, padx=2)

            elif act == 'wait':
                ttk.Label(params_frame, text='Seconds:').pack(side=tk.LEFT, padx=4); sec_ent = ttk.Entry(params_frame, width=6); sec_ent.pack(side=tk.LEFT, padx=2)

            elif act == 'ocr_check':
                ttk.Button(params_frame, text='Select Region', command=lambda:self.select_region_for_dlg(dlg, action_cb.get())).pack(anchor='w')
                ttk.Label(params_frame, text='Expected Text:').pack(anchor='w'); exp_ent = ttk.Entry(params_frame, width=40); exp_ent.pack(anchor='w')

            elif act == 'excel_loop':
                ttk.Label(params_frame, text='Excel Path:').pack(anchor='w')
                exc_ent = ttk.Entry(params_frame, width=40); exc_ent.pack(side=tk.LEFT, padx=4)
                ttk.Button(params_frame, text='Browse', command=lambda: self.set_entry_file(exc_ent)).pack(side=tk.LEFT)
                ttk.Label(params_frame, text='Sheet (name or index):').pack(anchor='w'); sheet_ent = ttk.Entry(params_frame, width=20); sheet_ent.pack(anchor='w')
                ttk.Label(params_frame, text='Note: after adding a loop, edit the saved flow to add nested steps')

        action_cb.bind('<<ComboboxSelected>>', on_action_change)
        on_action_change()

        def add_and_close():
            act = action_cb.get()
            d = {'action': act}

            # Collect simple params using a blocking input helper for reliability
            if act in ('click','move'):
                val = simple_input('Enter X,Y (leave blank if you used Select Region):')
                if val:
                    try:
                        x,y = val.split(',')
                        d['x'] = int(x.strip()); d['y'] = int(y.strip())
                    except Exception:
                        messagebox.showwarning('Invalid', 'Failed to parse X,Y')

            elif act == 'type':
                val = simple_input('Enter text to type (you can use {col} for Excel columns):')
                d['text'] = val

            elif act == 'hotkey':
                val = simple_input('Enter keys like ctrl+shift+s:')
                d['keys'] = val

            elif act == 'wait':
                val = simple_input('Seconds to wait:')
                try: d['seconds'] = float(val or 1)
                except: d['seconds'] = 1

            elif act == 'ocr_check':
                val = simple_input('Enter expected text (optional):')
                d['expected'] = val

            elif act == 'excel_loop':
                path = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx;*.xls')])
                if not path:
                    messagebox.showwarning('Missing', 'Excel not selected')
                    return
                d['path'] = path
                d['sheet'] = 0
                d['steps'] = []

            self.steps.append(d)
            self.refresh_steps()
            dlg.destroy()

        ttk.Button(dlg, text='Add', command=add_and_close).grid(row=2,column=0, pady=6)
        ttk.Button(dlg, text='Cancel', command=dlg.destroy).grid(row=2,column=1, pady=6)

    def set_entry_file(self, entry_widget):
        p = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx;*.xls')])
        if p:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, p)

    def select_region_for_dlg(self, parent_dlg, action):
        """Open the region selector and append a step when done.
        Region callback provides bbox: [left,top,right,bottom]
        """
        def cb(region):
            left, top, right, bottom = region
            cx = (left + right) // 2
            cy = (top + bottom) // 2
            step = {'action': action, 'x': cx, 'y': cy, 'region': [left, top, right, bottom]}
            self.steps.append(step)
            self.refresh_steps()

        parent_dlg.withdraw()
        # allow overlay to show after a short delay
        self.after(200, lambda: RegionSelector(self, cb))
        parent_dlg.deiconify()

    def remove_step(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        del self.steps[idx]
        self.refresh_steps()

    def move_step(self, direction):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        new = idx + direction
        if new < 0 or new >= len(self.steps): return
        self.steps[idx], self.steps[new] = self.steps[new], self.steps[idx]
        self.refresh_steps()
        self.listbox.select_set(new)

    def on_select(self, event=None):
        sel = self.listbox.curselection()
        for w in self.detail_frame.winfo_children(): w.destroy()
        if not sel: return
        idx = sel[0]
        step = self.steps[idx]
        ttk.Label(self.detail_frame, text=json.dumps(step, indent=2)).pack(anchor='nw')

    def refresh_steps(self):
        self.listbox.delete(0, tk.END)
        for i,s in enumerate(self.steps):
            label = f"{i+1}. {s.get('action')}"
            if s.get('action') in ('click','move') and 'x' in s:
                label += f" @({s.get('x')},{s.get('y')})"
            self.listbox.insert(tk.END, label)

    # -----------------------------
    # Flow file IO
    # -----------------------------
    def save_flow(self):
        path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON','*.json')])
        if not path: return
        with open(path,'w', encoding='utf-8') as f:
            json.dump(self.steps, f, indent=2)
        messagebox.showinfo('Saved', f'Flow saved to {path}')

    def load_flow(self):
        path = filedialog.askopenfilename(filetypes=[('JSON','*.json')])
        if not path: return
        with open(path, 'r', encoding='utf-8') as f:
            self.steps = json.load(f)
        self.refresh_steps()
        messagebox.showinfo('Loaded', f'Flow loaded from {path}')

    # -----------------------------
    # Run / Stop
    # -----------------------------
    def run_flow(self):
        if not self.steps:
            messagebox.showwarning('Empty', 'No steps to run')
            return
        runner = FlowRunner(self.steps, self.log_box)
        self._runner = runner
        def task():
            try:
                runner.run()
            except Exception as e:
                messagebox.showerror('Run Error', str(e))
        t = threading.Thread(target=task, daemon=True)
        t.start()

    def stop_flow(self):
        if self._runner:
            self._runner.stop()
            self.log_box.insert(tk.END, f"[{timestamp()}] Stop requested\n")

    def on_exit(self):
        """Called when the user requests to exit the app (Exit button or window close).
        Stops any running FlowRunner and then closes the GUI.
        """
        if messagebox.askokcancel('Exit', 'Do you really want to exit?'):
            if self._runner:
                try:
                    self._runner.stop()
                    self.log_box.insert(tk.END, f"[{timestamp()}] Exit requested - stopping runner")
                except Exception:
                    pass
            self.destroy()


# -----------------------------
# Simple input helper (blocking)
# -----------------------------
def simple_input(prompt):
    def on_ok():
        nonlocal val
        val = entry.get()
        win.destroy()
    val = ''
    win = tk.Toplevel()
    win.title('Input')
    tk.Label(win, text=prompt).pack(padx=8, pady=6)
    entry = tk.Entry(win, width=60)
    entry.pack(padx=8, pady=6)
    entry.focus()
    tk.Button(win, text='OK', command=on_ok).pack(pady=6)
    win.transient()
    win.grab_set()
    win.wait_window()
    return val

# -----------------------------
# Run app
# -----------------------------
if __name__ == '__main__':
    app = App()
    app.mainloop()
