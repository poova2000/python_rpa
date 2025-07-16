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
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\\tesseract.exe"
# EXCEL_PATH = r"D:\\InwardFiles\\InwardFiles_TSC.xlsx"
TEMP_IMAGE_PATH = r"D:\\python\\temp_area.png"
VEDNOR_IMAGE_PATH = r"D:\\python\\vednor.png"
OCR_REGION = (553,238,110,29)
VENDOR_REGION = (553,238,110,29)  # x, y, width, height - Adjust for your screen
WAIT_SHORT = 1
WAIT_LONG = 3
LOG_FILE = "rpa_log.txt"
VENDOR=['AARAV KNIT GARMENTS','AATARSH CLOTHING COMPANY','ATITHYA CLOTHING COMPANY','RAMRAJ HANDLOOMS','RAM AND RAM FABRIC','VIVEAGHAM GARMENTS','R AND R TEXTILE','TARA SAREE COLLECTION','VARUN TEXTILE','THARUN TEXTILE','B AND B TEXTILE(A UNIT OF ENES TEXTILE MILLS)']
# === UTILITIES ===
def read_text_from_screen(region):
    screenshot = pyautogui.screenshot(region=region)
    screenshot.save(TEMP_IMAGE_PATH)
    r.wait(2)
    img = Image.open(TEMP_IMAGE_PATH)
    text = pytesseract.image_to_string(img)
    return text
def read_vendor(region):
    screenshot = pyautogui.screenshot(region=VENDOR_REGION)
    screenshot.save(VEDNOR_IMAGE_PATH)
    r.wait(2)
    img = Image.open(VEDNOR_IMAGE_PATH)
    vednor = pytesseract.image_to_string(img)
    return vednor

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
def handle_failure(error_message):
    traceback_str = traceback.format_exc()
    full_message = f"[ERROR] {error_message}\nTraceback:\n{traceback_str}"
    logging.error(full_message)
    r.close()
    sys.exit(1)



def process_excel_and_transfer(): 
    time.sleep(5)
    logging.info('wait 3')
    r.keyboard('[F7]')
    logging.info('clicked F7')
    r.wait(WAIT_SHORT)
    logging.info('wait 2')
    r.keyboard('[F7]')
    logging.info('clicked F7 2')
    r.wait(WAIT_SHORT)
    logging.info('wait 2')
    # r.wait(WAIT_SHORT)
    # logging.info('wait 2')
    # r.keyboard('[F7]')
    # r.wait(WAIT_SHORT)
    # logging.info('wait 2')
    # r.keyboard('[F7]')
    pyautogui.click(191,114)
    r.wait(WAIT_LONG)
    pyautogui.click(193,167)
    r.wait(WAIT_LONG)
    r.keyboard('[enter]')
    r.wait(WAIT_SHORT)
    r.keyboard('[enter]')
    r.wait(WAIT_SHORT)
    # r.keyboard('[down]')
    # r.wait(WAIT_SHORT)
    # r.keyboard('[down]') 
    # r.wait(WAIT_SHORT)

 
    # r.keyboard(VENDOR[name-1])
    r.wait(WAIT_SHORT)

    r.keyboard('[enter]')
    r.wait(WAIT_SHORT) 
    r.keyboard('[down]')
    r.wait(WAIT_SHORT)
    r.keyboard('[down]') 
    r.wait(WAIT_SHORT) 
    r.keyboard('[down]') 
    r.wait(WAIT_SHORT) 
    # r.keyboard('[down]') 
    # r.wait(WAIT_SHORT) 
    r.keyboard('[enter]')
    r.wait(WAIT_SHORT)
    max_attempts = 12
    for attempt in range(max_attempts):
        logging.info(f"OCR validation attempt {attempt+1}")
        time.sleep(5)
        text = read_text_from_screen(OCR_REGION)


        logging.info(f"OCR Read: {text.strip()}")
        # logging.info(f"OCR Invoice: {invoice.strip()}")

        if text and not text.startswith("0"):

            logging.info("OCR validation successful. Proceeding to save.")
            # r.hover('temp_area.png')
            time.sleep(WAIT_SHORT)
            # r.hover('save.png')
            r.keyboard('[F6]')
            time.sleep(WAIT_LONG)
            r.keyboard('[enter]')

            # location1=pyautogui.locateOnScreen('yes.png',confidence=0.8)
            # if location1:
            #    pyautogui.click(location1)
            time.sleep(WAIT_LONG)
            save_attempts = 12
            for save_attempt in range(save_attempts):
                logging.info(f"OCR Save attempt {save_attempt + 1}")
                time.sleep(5)
                if save_attempts > save_attempt + 1:
                  
                    r.wait(WAIT_LONG)
                    r.keyboard('[enter]')
                    # location=pyautogui.locateOnScreen('ok.png',confidence=0.8)
                    # if location:
                    #     pyautogui.click(location)
                    time.sleep(5)
                    r.keyboard('[F7]')
                    time.sleep(2)
                    r.keyboard('[F7]')
                    process_excel_and_transfer()
                    
                elif save_attempts == save_attempt + 1:
                     time.sleep(2)
                     r.keyboard('[F7]')
                     time.sleep(2)
                     r.keyboard('[F7]')
                     time.sleep(2)
                     process_excel_and_transfer()
                    # raise Exception(f"OCR failed after {save_attempts} attempts to save invoice")
                    # return handle_failure('failed')
                    
        elif attempt == max_attempts - 1:
                     time.sleep(2)
                     r.keyboard('[F7]')
                     time.sleep(2)
                     r.keyboard('[F7]')
                     time.sleep(2)
                     process_excel_and_transfer()
            
    r.close()




# === MAIN FLOW ===
if __name__ == "__main__":
    try:
        logging.info("Starting Inward RPA...")
        # name=int(input("ENTER VENDOR 1.'AARAV KNIT GARMENTS 2.AATARSH CLOTHING COMPANY' 3.ATITHYA CLOTHING COMPANY 4.RAMRAJ HANDLOOMS 5.RAM AND RAM FABRIC 6.VIVEAGHAM GARMENTS 7.R AND R TEXTILE 8.TARA SAREE COLLECTION 9.VARUN TEXTILE 10.THARUN TEXTILE 11.B AND B TEXTILE(A UNIT OF ENES TEXTILE MILLS)' "))
       
        r.init(visual_automation=True, chrome_browser=False)
        time.sleep(2)
        r.keyboard('[ctrl][p]')
        time.sleep(10)

        process_excel_and_transfer()

        logging.info("Process completed successfully.")
        r.keyboard('[esc]')
        r.close()

    except Exception as main_err:
        handle_failure(f"Script crashed: {main_err}")