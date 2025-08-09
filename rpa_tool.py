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
import numpy
import logging
r.init(visual_automation=True, chrome_browser=False)
excel='D:\\test.xlsx'
df = pd.read_excel(excel)
for index, row in df.iterrows():
        outlet = str(row['filename'])
        invoice = str(row['number'])
        print(outlet)
        print(invoice)
        text="MR23"
        if text != invoice:
            
            # logging.warning(f"Invoice mismatch at row {index+1}. Expected: {invoice}, Detected: {text} Retrying...")
            print("invoice doest not matched")

        max_attempts = 8
        for attempt in range(max_attempts):
            print("outside irn",attempt)
            if text and not text.startswith("0"):
                        print("invoice match")
                        print("CR validation successful. Proceeding to save.")

                        # logging.info("OCR validation successful. Proceeding to save.")
                        # r.hover('temp_area.png')
                        # time.sleep(WAIT_SHORT)
                    
                        time.sleep(1)
                        max_attempts1 = 8
                        for attempt1 in range(max_attempts1):
                            print("inside irn ",attempt1)
                            print("OCR IRN validation attempt",attempt1+1)

                            # logging.info(f"OCR IRN validation attempt {attempt+1}")
                            time.sleep(1)
                            if r.present('no.png') and r.present('ok.png'):
                                if r.present('run.png'):
                                  r.hover('run.png')
                                if r.present('terminal.png'):
                                  r.hover('terminal.png')    
                                print("invoice saved")
                                break   
                        
                            elif attempt1==max_attempts1-1:
                                   raise Exception(f"OCR failed to save invoice after {max_attempts} attempts")
                        break      
            elif attempt == max_attempts - 1:
                raise Exception(f"OCR failed after {max_attempts} attempts for invoice '{invoice}'")
                break

