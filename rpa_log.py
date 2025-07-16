import rpa as r
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

# === CONFIGURATION ===
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
EXCEL_PATH = r"D:\\python\\out.xlsx"
TEMP_IMAGE_PATH = r"D:\\python\\temp_area.png"
OCR_REGION = (1392, 357, 199, 57)  # x, y, width, height - Adjust for your screen
WAIT_SHORT = 1
WAIT_LONG = 3
LOG_FILE = "rpa_log.txt"

# === SETUP ===
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# === LOGGING CONFIGURATION ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stdout)
    ]
)

# === UTILITIES ===
def read_text_from_screen(region):
    screenshot = pyautogui.screenshot(region=region)
    screenshot.save(TEMP_IMAGE_PATH)
    r.wait(2)
    img = Image.open(TEMP_IMAGE_PATH)
    text = pytesseract.image_to_string(img)
    return text

def handle_failure(error_message):
    traceback_str = traceback.format_exc()
    full_message = f"[ERROR] {error_message}\nTraceback:\n{traceback_str}"
    logging.error(full_message)
    r.close()
    sys.exit(1)

# === MAIN TASK ===
def process_excel_and_transfer(excel_file):
    df = pd.read_excel(excel_file)
    for index, row in df.iterrows():
        try:
            outlet = str(row['store'])
            invoice = str(row['invoice'])

            logging.info(f"Processing row {index+1}: Outlet='{outlet}', Invoice='{invoice}'")

            r.keyboard('[enter]')
            time.sleep(WAIT_SHORT)
            r.keyboard(outlet)
            time.sleep(WAIT_SHORT)
            r.keyboard('[enter]')
            time.sleep(WAIT_SHORT)
            r.keyboard('[shift][F7]')
            time.sleep(WAIT_SHORT)
            r.keyboard('1')
            time.sleep(WAIT_SHORT)
            r.keyboard('[enter]')
            time.sleep(WAIT_SHORT)
            r.keyboard(invoice)
            r.wait(2)
            # r.keyboard(invoice + ' [left]')

            invoice_region = (661,428,98,16)
            # x=722 
            # y=500
            # pyautogui.click(x, y, duration=1)
            screenshot = pyautogui.screenshot(region=invoice_region)
            r.wait(2)
            screenshot.save('invoice.png')
            r.wait(1)

            custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=MR0123456789'
            # img = Image.open('invoice.png').convert('L')
            # img = ImageOps.invert(img)
            # img = img.point(lambda x: 0 if x < 180 else 255)
            text = pytesseract.image_to_string(screenshot, config=custom_config).strip()
            r.wait(1)

            if text != invoice:
                logging.warning(f"Invoice mismatch at row {index+1}. Expected: {invoice}, Detected: {text} Retrying...")
                r.keyboard('[esc]')
                time.sleep(WAIT_SHORT)
                r.keyboard('[F7]')
                continue

            time.sleep(WAIT_SHORT)
            r.keyboard('[enter]')
            time.sleep(WAIT_SHORT)

            max_attempts = 8
            for attempt in range(max_attempts):
                logging.info(f"OCR validation attempt {attempt+1}")
                time.sleep(10)
                text = read_text_from_screen(OCR_REGION)


                logging.info(f"OCR Read: {text.strip()}")
                logging.info(f"OCR Invoice: {invoice.strip()}")

                if text and not text.startswith("0"):

                    logging.info("OCR validation successful. Proceeding to save.")
                    # r.hover('temp_area.png')
                    # time.sleep(WAIT_SHORT)
                    # r.hover('save.png')
                   
                    r.keyboard('[F6]')
                    time.sleep(WAIT_LONG)
                    r.keyboard('[enter]')
                    time.sleep(WAIT_LONG)
                    r.keyboard('[right]')
                    r.wait(1)
                    r.keyboard('[enter]')
                    r.wait(1)
                    break
                elif attempt == max_attempts - 1:
                    raise Exception(f"OCR failed after {max_attempts} attempts for invoice '{invoice}'")
                    

        except Exception as row_err:
            handle_failure(f"Failed at row {index+1}: {row_err}")
            continue

# === MAIN FLOW ===
if __name__ == "__main__":
    try:
        logging.info("Starting RPA...")
        r.init(visual_automation=True, chrome_browser=False)
        time.sleep(1)
        r.keyboard('[ctrl][k]')
        time.sleep(1)

        process_excel_and_transfer(EXCEL_PATH)

        logging.info("Process completed successfully.")
        r.keyboard('[esc]')
        r.close()

    except Exception as main_err:
        handle_failure(f"Script crashed: {main_err}")