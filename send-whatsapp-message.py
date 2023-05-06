# Dependencies: import libraries
import time
import pywhatkit
import pyautogui
from pynput.keyboard import Key, Controller
from win32 import win32clipboard
import pandas as pd

keyboard = Controller()


def send_whatsapp_message(msg: str, cellphone: str):
    try:
        # pywhatkit.sendwhatmsg_instantly(
        #     phone_no="(+ PLUS SYMBOL) + COUNTRYCODE + NUMBER)",
        #     message=msg,
        #     tab_close=True
        # )
        # IMPORTANT: Only images in JPEG format.
        pywhatkit.sendwhats_image(
            "+COUNTRY CODE" + cellphone, "\\FILE PATH", msg, 10, True, 3)
        time.sleep(6)

        pyautogui.click()
        time.sleep(2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        print("Message to: " + cellphone + " was sent correctly.")
    except Exception as e:
        print(str(e))


if __name__ == "__main__":
    # IMPORTANT: pd.read_excel(FILE NAME.xlsx, Sheet Name)
    data = pd.read_excel("FILE NAME.xlsx", sheet_name='SHEET NAME')
    # IMPORTANT: Print only the 3 first registers
    print(data.head(3))
    for i in range(len(data)):
        # IMPORTANT: convert to string with str(data.loc[i, 'COLUMN NAME'])
        number = str(data.loc[i, 'COLUMN NAME'])
        # msg: message
        send_whatsapp_message(
            msg="*This is a test message.*", cellphone=number)

# pywhatkit.sendwhatmsg(
#     phone_no="<phone-number>",
#     message="This is a scheduled message.",
#     time_hour=9,
#     time_min=47
# )
