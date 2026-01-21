import rpa as r
import time
import pyautogui
from PIL import Image
import pytesseract
import sys
import traceback
import logging
import re
import oracledb
import cv2
import numpy as np
import pandas as pd
from datetime import datetime


# === CONFIGURATION ===
TESSERACT_PATH =  r"C:\Program Files\Tesseract-OCR\\tesseract.exe"
TEMP_IMAGE_PATH = r"D:\\python\\temp_area.png"
CONFIRM_IMAGE_PATH = r"D:\\python\\confirm_image.png"
CONFIRM_YES_IMAGE_PATH = r"D:\\python\\confirm_yes_image.png"
EXCEL_PATH = r"D:\\python\\out.xlsx"
# OCR_REGION = (1708,353,207,65)
OCR_REGION = (1711,354,205,60)
  # Adjust for your screen
WAIT_SHORT = 1
WAIT_LONG = 3
LOG_FILE = "rpa_log.txt"

# === ORACLE DB CONFIG ===
DB_USER = "RAYMEDI_RAMRAJ"
DB_PASS = "raymedi_hq"
DB_DSN = "ramraj-qa.cugyvyz68ru0.ap-south-1.rds.amazonaws.com:1521/ramrajqa"

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
    return pytesseract.image_to_string(img)

def handle_failure(error_message):
    traceback_str = traceback.format_exc()
    logging.error(f"[ERROR] {error_message}\nTraceback:\n{traceback_str}")
    r.close()
    sys.exit(1)
def stock_not():
   print('present')
   time.sleep(2)
   r.keyboard('[right]')
   time.sleep(2)
   r.keyboard('[enter]')
   time.sleep(2)
   r.keyboard('[down]')
   time.sleep(2)
   r.keyboard('[down]')
   time.sleep(2)
   r.keyboard('[enter]')    

# === FETCH DATA FROM ORACLE DB ===
def fetch_outlet_invoice():
    connection = oracledb.connect(user=DB_USER, password=DB_PASS, dsn=DB_DSN)
    cursor = connection.cursor()

    query = """
   SELECT 
TRANS_.OUTLET_NAME AS OUTLET,
MR_NUMBER
FROM
(SELECT
    hroi.MSM_SHOP_NAME AS OUTLET_NAME,
    MRC_NUMBER AS MR_NUMBER
FROM
    (
    SELECT
        mmh.MMH_DIST_BILL_NO AS InvoiceNo,
        mmh.MMH_DIST_BILL_DT AS Invoice_Date,
        mmh.MMH_MRC_PREFIX || MMH_MRC_NO AS MRC_NUMBER,
        mmh.MMH_CREATED_DATE AS MRC_GENERATED_DATE,
        msh.MSH_DN_PREFIX || msh.MSH_DN_NO AS TO_ID,
        'WH-DEPOT' AS Transfer_From_SR,
             msh.MSH_DN_CUST_NAME AS Transferred_To_SR,
        mmh.MMH_NARRATION AS H_CARDCODE
    FROM
        RAYMEDI_CS.med_sdn_hdr msh
    RIGHT JOIN RAYMEDI_CS.med_mrc_hdr mmh ON
        msh.RETAIL_OUTLET_ID = mmh.RETAIL_OUTLET_ID
        AND msh.MSH_OTHERS4 = to_char(mmh.MMH_NO)
        WHERE
        mmh.retail_outlet_id = 318
        -- AND mmh.MMH_NARRATION != 'U05'
        -- AND mmh.MMH_NARRATION != 'U04'
        AND mmh.MMH_NARRATION != 'U06'
        AND mmh.MMH_NARRATION != 'R262'
        AND mmh.MMH_NARRATION != 'R261'
        AND mmh.MMH_NARRATION != 'R258'
      
        AND MMH_DIST_BILL_DT >= to_date('2025-12-01', 'YYYY-MM-DD')
        -- AND MMH_DIST_BILL_DT <= to_date('2025-12-25', 'YYYY-MM-DD')
    ORDER BY
        MMH_DIST_BILL_DT ASC)k
LEFT JOIN RAYMEDI_hq.HQ_RETAIL_OUTLET_INFO hroi ON
    k.H_CARDCODE = hroi.ERP_REF_CODE
WHERE
    TO_ID IS NULL AND hroi.MSM_RAD_LIC_NO IS NOT NULL  )TRANS_ 
WHERE  TRANS_.MR_NUMBER  NOT IN 
( SELECT torl.MR_NO FROM RAYMEDI_RAMRAJ.TRANSFER_OUT_RPA_LOG torl WHERE TRANS_.MR_NUMBER=torl.MR_NO)
FETCH FIRST 20 ROWS ONLY
    """

    cursor.execute(query)
    results = cursor.fetchall()
    cursor.close()
    connection.close()

    return results  # List of tuples (Outlet, Invoice)

def load_excel():
    df = pd.read_excel(EXCEL_PATH, sheet_name="Sheet1")
    if 'STATUS' not in df.columns:
        df['STATUS'] = ''
    return df


def save_excel(df):
    df.to_excel(EXCEL_PATH, index=False)    


    # === MAIN TASK ===

def process_data_from_db():
    # connection = oracledb.connect(user=DB_USER, password=DB_PASS, dsn=DB_DSN)
    # cursor = connection.cursor()
    # data = fetch_outlet_invoice()
    # datetime_str = datetime.now()
    # current_datetime = datetime_str.strftime("%Y-%m-%d %H:%M:%S")

    
    # if not data: 
    #     print("No data available")
    # else:    
    #     logging.info(f"Fetched {len(data)} records from DB")
    df = load_excel()
    records = df[df['STATUS'].isna() | (df['STATUS'] == '')]

    if records.empty:
        logging.info("No pending records in Excel")
        return

    logging.info(f"Processing {len(records)} records from Excel")

    for idx, row in records.iterrows():
        outlet = str(row['OUTLET']).strip()
        invoice = str(row['MR_NUMBER']).strip()
        # print(outlet)
        # print(invoice)
        # r.close()
        # sys.exit(1)
     
        try:
                logging.info(f"Processing {idx}: Outlet='{outlet}', Invoice='{invoice}'")

                r.keyboard('[enter]')
                time.sleep(WAIT_SHORT)
                r.keyboard(outlet)
                # outlet_region = (1373,390,398,21)
                # # pyautogui.click(x, y, duration=1)
                # screenshot = pyautogui.screenshot(region=outlet_region)
                # screenshot.save('outlet.png')
                # outlet_img = Image.open('D:\\python\\outlet.png')
                # custom_config = r'--oem 3 --psm 8 -c tessedit_char_whitelist= ABCDEFGHIJKLMNOPQRSTUVWXYZ-'
                # outlet_name = pytesseract.image_to_string(outlet_img,config=custom_config).strip()
                # logging.info(f"OCR REGION OUTLET NAME='{outlet_name}'")
                # if outlet_name!=outlet:
                #     time.sleep(2)
                #     r.keyboard('[esc]')
                #     time.sleep(2)
                #     continue
                time.sleep(WAIT_LONG)
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

                # invoice_region =(979,595,99,24)
                invoice_region =(980,598,98,19)
                screenshot = pyautogui.screenshot(region=invoice_region)
                screenshot.save('invoice.png')
                r.wait(1)

                custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=MR0123456789'
                img = cv2.imread('D:\\python\\invoice.png')
                # gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                # thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
                text = pytesseract.image_to_string(img,config=custom_config).strip()
        
                if text != invoice:
                    logging.warning(f"Invoice mismatch. Expected: {invoice}, Detected: {text}. Skipping row.")
                    r.keyboard('[esc]')
                    time.sleep(WAIT_SHORT)
                    r.keyboard('[F7]')
                    continue

                time.sleep(5)
                
                r.keyboard('[enter]')
                time.sleep(4)

                for attempt in range(22):
                    logging.info(f"OCR validation attempt {attempt+1}")
                    time.sleep(5)
                    text = read_text_from_screen(OCR_REGION)
                    if r.present('stock_not.png'):
                            stock_not()
                            pass

                    if text and not text.startswith("0"):
                        logging.info("OCR validation successful.")
                        time.sleep(5)
                        r.keyboard('[F6]')
                        time.sleep(5)
                        r.keyboard('[enter]')
                                
                        time.sleep(5)
                        if outlet.startswith("TN-"):
                            time.sleep(8)
                            r.click('no.png')
                            time.sleep(5)
                            #ins_query = "INSERT INTO RAYMEDI_RAMRAJ.TRANSFER_OUT_RPA_LOG (OUTLET_NAME,MR_NO,CREATED_DATE,CREATED_TIME,STATUS) values (TRIM(outlet), TRIM(invoice), SYSDATE, SYSDATE, 'Saved')"
                            # ins_query = (
                            #    "INSERT INTO RAYMEDI_RAMRAJ.TRANSFER_OUT_RPA_LOG "
                            #     "(OUTLET_NAME, MR_NO, CREATED_DATE, CREATED_TIME, STATUS) "
                            #     "VALUES ('" + outlet.strip() + "', '" + invoice.strip() + "', TO_DATE('"+current_datetime+"','YYYY-MM-DD HH24:MI:SS') , SYSDATE, 'Saved')"
                            # )
                            # # print ("insquery:", ins_query)
                            # cursor.execute(ins_query)
                            # connection.commit()
                            logging.info("Invoice saved successfully")
                            

                        else:
                            for attempt2 in range(12):
                                r.wait(5)
                                if r.present('ok.png'):
                                    time.sleep(4)
                                    r.click('ok.png')
                                    r.wait(3)
                                    r.click('ok.png')
                                    r.wait(3)
                                    r.click('no.png')
                                    time.sleep(3)
                                   #  ins_query = (
                                   # "INSERT INTO RAYMEDI_RAMRAJ.TRANSFER_OUT_RPA_LOG "
                                   #  "(OUTLET_NAME, MR_NO, CREATED_DATE, CREATED_TIME, STATUS) "
                                   #  "VALUES ('" + outlet.strip() + "', '" + invoice.strip() + "', TO_DATE('"+current_datetime+"','YYYY-MM-DD HH24:MI:SS'), SYSDATE, 'Saved')"
                                   #  )
                                   #  # print ("insquery:", ins_query)
                                   #  cursor.execute(ins_query)
                                   #  connection.commit()
                                    logging.info("Invoice saved successfully")
                                   
                                    break
                                elif r.present('no.png'):
                                    time.sleep(4)
                                    r.click('no.png')
                                    time.sleep(4)
                                   #  ins_query = (
                                   # "INSERT INTO RAYMEDI_RAMRAJ.TRANSFER_OUT_RPA_LOG "
                                   #  "(OUTLET_NAME, MR_NO, CREATED_DATE, CREATED_TIME, STATUS) "
                                   #  "VALUES ('" + outlet.strip() + "', '" + invoice.strip() + "', TO_DATE('"+current_datetime+"','YYYY-MM-DD HH24:MI:SS'), SYSDATE, 'Saved')"
                                   #  )
                                   #  # print ("insquery:", ins_query)
                                   #  cursor.execute(ins_query)
                                   #  connection.commit()
                                    logging.info("Invoice saved successfully")
                                   
                                    break
                        break
                    elif attempt == 22:
                        if r.present('stock_not.png'):
                            stock_not()
                            pass

                        raise Exception(f"OCR failed after {attempt+1} attempts for invoice '{invoice}'")

        except Exception as row_err:
                handle_failure(f"Failed at record {index}: {row_err}")
                continue

# === MAIN FLOW ===
if __name__ == "__main__":
    try:        

        while True:
            logging.info("Starting RPA with DB source...")
            r.init(visual_automation=True, chrome_browser=False)
            time.sleep(1)
            r.keyboard('[ctrl][k]')
            time.sleep(2)            
            process_data_from_db()
            logging.info("Process completed successfully.")
            r.keyboard('[esc]')
            r.close()
            sys.exit(1)

       

    except Exception as main_err:
        handle_failure(f"Script crashed: {main_err}")
