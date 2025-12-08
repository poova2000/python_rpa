import pandas as pd
import time
import pyautogui
from PIL import Image, ImageOps
import pytesseract
import os
import sys
import traceback
import pyperclip
import cv2
import numpy
import logging
from logging import StreamHandler

LOG_FILE = "rpa_log.txt"
EXCEL_PATH = r"G:\\python\\saleorder.xlsx"
customer = r"G:\\python\\customer.png"
pyautogui.FAILSAFE = False
TESSERACT_PATH = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

gray_image = 'G:\\python\\gray.png'
reference_image_name = 'G:\\python\\customer_advance.png'
cancel = 'G:\\python\\cancel.png'
screenshot_output_name = 'G:\\python\\found_area.png'

# customer_advance = 'G:\\python\\customer_advance.png'
# amazon = 'G:\\python\\amazon.png'
# flipkart= 'G:\\python\\flipkart.png'
def select_tender(reference_image_name):
    confidence_level=0.9
    print("Switch to the target application within 5 seconds...")
    pyautogui.sleep(5)

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
            pyautogui.click(location)
            time.sleep(2)
            # pyautogui.press('enter')
            # time.sleep(2)
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

def wait_and_click(image, timeout=12, confidence=0.7):
    start = time.time()

    while time.time() - start < timeout:
        try:
            location = pyautogui.locateOnScreen(image, confidence=confidence)
        except:
            location = None

        if location:
            x, y = pyautogui.center(location)
            pyautogui.click(x, y)
            logging.info(f"Clicked: {image}")
            return True

        time.sleep(0.5)

    logging.error(f"Not found: {image}")
    return False


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        StreamHandler(sys.stdout)
    ]
)
def process_excel_and_transfer(excel_file):
   

    # Test if image is detected
    # test = pyautogui.locateOnScreen(customer, confidence=0.7)
    # wait_and_click(test)
    # print("DETECTION TEST:", test)
    df = pd.read_excel(EXCEL_PATH)
    for index, row in df.iterrows():
        pyautogui.keyDown('shift')
        pyautogui.keyDown('f3')
        pyautogui.keyUp('f3')
        pyautogui.keyUp('shift')
        time.sleep(1)
        print("Opened Sale Order Filter Screen")
        pyautogui.press('enter')
        time.sleep(1)
        
        name = str(row['name'])
        amount = str(row['amount'])
        time.sleep(1)
        pyautogui.press('ctrl')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.write(name)
        time.sleep(2)
        print(name)
        pyautogui.press('enter')
        time.sleep(2)
        invoice_region = (1433,331,165,50)
        # x=722 
        # y=500
        # pyautogui.click(x, y, duration=1)
        screenshot = pyautogui.screenshot(region=invoice_region)
        time.sleep(2)
        screenshot.save('invoice.png')
        time.sleep(2)

        custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789.'
        # img = Image.open('invoice.png').convert('L')
        # img = ImageOps.invert(img).strip()
        # img = img.point(lambda x: 0 if x < 180 else 255)
        text = pytesseract.image_to_string(screenshot, config=custom_config).strip()
        excel_int = int(float(amount))
        detected_int = int(float(text))
        print('Excel amount',excel_int)
        print('Detected amount',detected_int)
        if excel_int != detected_int:
            print("Amount not matched")
            pyautogui.press('f7')
            time.sleep(2)

            # select_tender(amazon)
            # time.sleep(5)
            pyautogui.press('enter')
            time.sleep(2)
            # sys.exit()
        # else:
        #     print("Amount matched")
        pyautogui.press('f6')
        time.sleep(2)
        select_tender(reference_image_name)
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)

        # # select_tender(amazon)
        # # time.sleep(5)
        # pyautogui.press('enter')
        # time.sleep(2)
        # sys.exit()



if __name__ == "__main__":
    time.sleep(5)
    print("RPA script started...")
    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('s')
    pyautogui.keyUp('s')
    pyautogui.keyUp('ctrl')
    time.sleep(1)
    print("Opened Sale BIll Screen")
    process_excel_and_transfer(EXCEL_PATH)
    time.sleep(2)
    pyautogui.press('esc')
    logging.info("Process completed successfully.")
    pyautogui.alert('Process completed successfully')

   
