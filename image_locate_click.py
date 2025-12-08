import pyautogui
import os
import cv2
import numpy as np
import time

# --- Configuration ---
# Ensure the reference image is in the same directory as the script or provide the full path
gray_image = 'G:\\python\\gray.png'
reference_image_name = 'G:\\python\\amazon.png'
cancel = 'G:\\python\\cancel.png'
screenshot_output_name = 'G:\\python\\found_area.png'
# The confidence argument requires opencv-python to be installed
confidence_level = 0.9

# Add a pause to switch to the application after running the script
print("Switch to the target application within 5 seconds...")
pyautogui.sleep(5)
pyautogui.press('f6')
time.sleep(2)



try:
    # 1. Locate the reference image on the screen
    # This function returns a 4-integer tuple: (left, top, width, height)
    location = pyautogui.locateOnScreen(
        reference_image_name,
        confidence=confidence_level
    )

    if location is not None:
        print(f"Image found at coordinates: {location}")

        # Unpack the location tuple for clarity
        left, top, width, height = location
        
        print(left)
        print(top)
        print(width)
        print(height)
        pyautogui.click(location)
        pyautogui.press('enter')
        time.sleep(5)
        pyautogui.press('enter')
        time.sleep(2)
        # if cancel_btn == pyautogui.locateCenterOnScreen("cancel.png", confidence=0.8):
        #    pyautogui.click(cancel_btn)
        

        # 2. Take a screenshot of the specific region where the image was found
        # The region argument takes a 4-integer tuple: (left, top, width, height)
        screenshot_region = (left, top, width, height)
        screenshot = pyautogui.screenshot(region=screenshot_region)

        

        # 3. Save the captured screenshot
        screenshot.save(screenshot_output_name)
        
        
        print(f"Screenshot saved as {screenshot_output_name} in the current directory: {os.getcwd()}")
    else:
        print(f"Could not find '{reference_image_name}' on the screen. Ensure it is visible and the file path is correct.")

except pyautogui.ImageNotFoundException:
    print(f"Could not find '{reference_image_name}' on the screen. Ensure it is visible.")
except Exception as e:
    print(f"An error occurred: {e}")

