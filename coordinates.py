import pyautogui
import time
print("Move mouse to TOP-LEFT corner... (you have 3 seconds)")
time.sleep(3)
top_left = pyautogui.position()
print(f"Top-left: {top_left}")

print("Move mouse to BOTTOM-RIGHT corner... (you have 3 seconds)")
time.sleep(3)
bottom_right = pyautogui.position()
print(f"Bottom-right: {bottom_right}")

width = bottom_right.x - top_left.x
height = bottom_right.y - top_left.y

print("\n📦 Rectangle Area:")
print(f"x: {top_left.x}, y: {top_left.y}, width: {width}, height: {height}")

# x: 661, y: 428,

# x: 733, y: 446,
# w=112
# h=18

# x: 1392, y: 357,
# x: 1591, y: 414,
# w=199
# h=57