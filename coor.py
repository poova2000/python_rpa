import tkinter as tk
import pyautogui

class OverlaySelector:
    def __init__(self):
        self.start_x = self.start_y = 0
        self.rect = None

        self.root = tk.Tk()
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.3)
        self.root.configure(bg='gray')

        self.canvas = tk.Canvas(self.root, bg="black", cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.root.mainloop()

    def on_mouse_down(self, event):
        pos = pyautogui.position()
        self.start_x, self.start_y = pos.x, pos.y
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="red", width=2
        )

    def on_mouse_drag(self, event):
        pos = pyautogui.position()
        current_x, current_y = pos.x, pos.y
        self.canvas.coords(self.rect, self.start_x, self.start_y, current_x, current_y)

    def on_mouse_up(self, event):
        end_pos = pyautogui.position()
        end_x, end_y = end_pos.x, end_pos.y

        x = min(self.start_x, end_x)
        y = min(self.start_y, end_y)
        width = abs(end_x - self.start_x)
        height = abs(end_y - self.start_y)

        with open("coords.txt", "w") as f:
            f.write(f"x: {x}\ny: {y}\nwidth: {width}\nheight: {height}\n")

        print(f"\n✅ Area selected: x={x}, y={y}, width={width}, height={height}")
        self.root.destroy()

if __name__ == "__main__":
    OverlaySelector()
