import pyautogui
import time
import pytesseract
from PIL import Image
time.sleep(2)
region = (0,215,177,31)

screenshot = pyautogui.screenshot(region=region)
screenshot.save('coor.png')
# r.wait(2)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\\tesseract.exe"
img = Image.open('D:\\python\\coor.png')
# custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=MR0123456789'
text = pytesseract.image_to_string(img).strip()
print("Extracted text:", text)
# r.hover('temp_area.png')
# r.close()