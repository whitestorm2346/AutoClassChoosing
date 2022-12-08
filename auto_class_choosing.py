'''
淡江大學自動選課程式

使用方式：
    。 將預加選之課程的"開課序號"逐一列在 classID.txt 裡
    。 執行 AutoClassChoosing.exe
    。 根據程式的指示輸入正確的學號、密碼以及選課的起始時間
    。 等待程式於指定時間自動登入進行選課
    。 程式執行完畢後，將會把選課資訊整理(成功或失敗)在 result.txt 裡

設計者： 資工二A 游昃勛
專案版本： 1.0
專案完成日期： 2022/09/12
'''

from datetime import datetime
from time import sleep
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
import ddddocr  # for detecting the confirm code
import base64   # for reading the image present in base 64

LOGIN_URL = 'https://www.ais.tku.edu.tw/EleCos/login.aspx'
TARGET_URL = 'https://www.ais.tku.edu.tw/EleCos/action.aspx'

RESULT_FILE = 'result.txt'
CLASS_ID_FILE = 'classID.txt'

LOGIN_FAIL = "E999 淡江大學個人化入口網帳密驗證失敗或驗證伺服器忙碌中, 請重新輸入或請參考密碼說明..."
CONFIRM_FAIL = "請輸入學號、密碼及驗證碼...(「淡江大學單一登入(SSO)」單一帳密驗證密碼)\n" \
               "※105學年度入學新生(含轉學生)起，預設為西元生日(西元年/月/日)後6碼，" \
               "例如西元生日為1997/01/05，則後6碼為970105※\nE903 驗證碼輸入錯誤,請重新輸入 !!!"
WRONG_TIME = "E999 登入失敗(非帳號密碼錯誤) ???\nE051 目前不是您的選課開放時間"

ADD_SUCCESS = "加選成功"
ADD_FAIL = "加選失敗"


class AutoClassChoosing:
    def __init__(self, student_num='', password='', starting_time='', expiry_time='', driver=None) -> None:
        self.student_num = student_num
        self.password = password
        self.starting_time = starting_time
        self.expiry_time = expiry_time
        self.driver = driver

    def run(self, entries=-1) -> int:
        if entries != -1:
            # get user's student number and password
            self.student_num = input('請輸入學號：')
            self.password = input('請輸入密碼：')
            self.starting_time = input('請輸入選課起始時間(例如： 2022/08/07 12:30): ')

        self.starting_time = datetime.strptime(
            self.starting_time, '%Y/%m/%d %H:%M')

        print('基本設定完成，將於指定時間自動登入選課')

        while not self.clock_on_time():
            sleep(1)

        edge_options = webdriver.EdgeOptions()
        edge_options.add_argument('--log-level=3')

        self.driver = webdriver.Edge(
            options=edge_options,
            executable_path='msedgedriver.exe',
            service_log_path='NUL'
        )

        while True:
            login_status = self.login()

            if login_status == 0:  # login success
                print('登入成功!!')
                break
            elif login_status == 1:  # wrong password or wrong student number
                print('學號或密碼錯誤!')
                self.student_num = input('請輸入學號： ')
                self.password = input('請輸入密碼： ')
            elif login_status == 2:  # wrong confirm code
                print('驗證碼錯誤，程式將重新判讀一次')
            elif login_status == 3:  # wrong login time
                self.starting_time = input(
                    '非登入時間，請輸入正確的選課起始時間(例如： 2022/08/07 12:30): ')
            else:  # other situations
                print('登入錯誤，程式將中斷執行')
                exit(1)

        if entries == -1:
            class_choosing_status = self.choose_classes()
        else:
            class_choosing_status = self.choose_classes(entries=entries)

        if class_choosing_status == 0:  # class choosing successfully
            print('完成自動選課，詳細結果紀錄於 result.txt')
        else:
            print('選課錯誤，程式將中斷執行')
            exit(1)

        return 0

    def clock_on_time(self) -> bool:
        curr_time = datetime.now()
        is_expired = curr_time >= self.starting_time

        return is_expired

    def login(self) -> int:
        self.driver.get(LOGIN_URL)

        try:
            # student number input
            student_num_input = self.driver.find_element(
                By.XPATH, '//*[@id="txtStuNo"]')
            student_num_input.clear()
            student_num_input.send_keys(self.student_num)

            # password input
            password_input = self.driver.find_element(
                By.XPATH, '//*[@id="txtPSWD"]')
            password_input.clear()
            password_input.send_keys(self.password)

            # confirm code input
            confirm_code_input = self.driver.find_element(
                By.XPATH, '//*[@id="txtCONFM"]')
        except Exception as e:
            print('網頁HTML錯誤')

            return 4

        confirm_code_input.clear()
        confirm_code_input.send_keys(self.auto_detect_confirm_code())

        login_btn = self.driver.find_element(By.XPATH, '//*[@id="btnLogin"]')
        login_btn.click()

        if self.driver.current_url == TARGET_URL:
            return 0
        else:
            msg = self.driver.find_element(
                By.XPATH, '//*[@id="form1"]/div[3]/table/tbody/tr[6]/td[2]')

            if msg.text == LOGIN_FAIL:
                return 1
            elif msg.text == CONFIRM_FAIL:
                return 2
            elif WRONG_TIME in msg.text:
                return 3
            else:
                return 4

    def auto_detect_confirm_code(self) -> str:
        # get the image(base64) using javascript
        captchaBase64 = self.driver.execute_async_script(
            """
            let canvas = document.createElement('canvas');
            let context = canvas.getContext('2d');
            let img = document.querySelector('#imgCONFM');

            canvas.height = img.naturalHeight;
            canvas.width = img.naturalWidth;

            context.drawImage(img, 0, 0);

            callback = arguments[arguments.length - 1];
            callback(canvas.toDataURL());
            """
        )

        # decode image(base64) to confirm code
        img = base64.b64decode(captchaBase64.split(',')[1])
        ocr = ddddocr.DdddOcr()
        confirm_code = ocr.classification(img)

        return confirm_code

    def choose_classes(self) -> int:
        with open('result.txt', 'w') as result_file:
            with open('classID.txt', 'r') as class_id_file:
                for id in class_id_file:
                    line = '開課序號：' + id + ' '

                    # class id input
                    class_id_input = self.driver.find_element(
                        By.XPATH, '//*[@id="txtCosEleSeq"]')
                    class_id_input.clear()
                    class_id_input.send_keys(id)

                    # add button click
                    add_btn = self.driver.find_element(
                        By.XPATH, '//*[@id="btnAdd"]')
                    add_btn.click()

                    msg = self.driver.find_element(
                        By.XPATH, '//*[@id="form1"]/div[3]/table/tbody/tr[2]/td[3]')

                    msg_in_line = msg.text.split('\n')

                    if ADD_SUCCESS in msg.text:
                        line += ADD_SUCCESS
                    elif ADD_FAIL in msg.text:
                        line += (ADD_FAIL + " ")
                        line += msg_in_line[1]
                    else:
                        line += "ERROR"

                    result_file.write(line + '\n')

        return 0

    def choose_classes(self, entries) -> int:
        for entry in entries:
            id = entry.value.get()
            line = '開課序號：' + id + ' '

            # class id input
            class_id_input = self.driver.find_element(
                By.XPATH, '//*[@id="txtCosEleSeq"]')
            class_id_input.clear()
            class_id_input.send_keys(id)

            # add button click
            add_btn = self.driver.find_element(
                By.XPATH, '//*[@id="btnAdd"]')
            add_btn.click()

            msg = self.driver.find_element(
                By.XPATH, '//*[@id="form1"]/div[3]/table/tbody/tr[2]/td[3]')

            msg_in_line = msg.text.split('\n')

            if ADD_SUCCESS in msg.text:
                line += ADD_SUCCESS
            elif ADD_FAIL in msg.text:
                line += (ADD_FAIL + " ")
                line += msg_in_line[1]
            else:
                line += "ERROR"

        return 0
