import threading
from tkinter import *
from tkinter import ttk
from datetime import datetime, time as Time
from time import sleep
from selenium import webdriver  # for operating the website
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import ddddocr  # for detecting the confirm code
import base64   # for reading the image present in base 64

LOGIN_URL = 'https://www.ais.tku.edu.tw/EleCos/login.aspx'
TARGET_URL = 'https://www.ais.tku.edu.tw/EleCos/action.aspx'

LOGIN_URL_ENG = 'https://www.ais.tku.edu.tw/EleCos_English/loginE.aspx'
TARGET_URL_ENG = 'https://www.ais.tku.edu.tw/EleCos_English/actionE.aspx'

LOGIN_FAIL = "E999 淡江大學個人化入口網帳密驗證失敗或驗證伺服器忙碌中, 請重新輸入或請參考密碼說明..."
CONFIRM_FAIL = "請輸入學號、密碼及驗證碼...(「淡江大學單一登入(SSO)」單一帳密驗證密碼)\n" \
               "※105學年度入學新生(含轉學生)起，預設為西元生日(西元年/月/日)後6碼，" \
               "例如西元生日為1997/01/05，則後6碼為970105※\nE903 驗證碼輸入錯誤,請重新輸入 !!!"
WRONG_TIME = "E999 登入失敗(非帳號密碼錯誤) ???\nE051 目前不是您的選課開放時間"
MAINTAIN_TIME = "E999 登入失敗(非帳號密碼錯誤) ???\nE071 現在為選課每日系統維護時間(11:30-12:30)"

LOGIN_FAIL_ENG = "E901 Student ID number error???"
CONFIRM_FAIL_ENG = "Please enter your Student ID number, Password and Verify code !!!\n" \
    "(Since Fall 2016, the default password (for freshmen and transfer students) " \
    "of \"TamKang University Single Sign On (SSO)\" will be set as the last six digits " \
    "of your date of birth (yyyy/mm/dd), for example, if your birthday is 1997/01/05, " \
    "your password will be 970105.)\nE903 Confirm code input Error !!!"
WRONG_TIME_ENG = "E999 Login Unsuccessful???\nE051 Currently not open for you"
MAINTAIN_TIME_ENG = "E999 Login Unsuccessful???\nE071 The daily maintain hour(11:30-12:30)"

ADD_SUCCESS = "加選成功"
ADD_FAIL = "加選失敗"

DROP_SUCCESS = "退選成功"
DROP_FAIL = "退選失敗"

ADD_SUCCESS_ENG = "Add successfully"
ADD_FAIL_ENG = "Add failed"

DROP_SUCCESS_ENG = "Drop successfully"
DROP_FAIL_ENG = "Drop Failed"

RESULT_FILE = 'result.txt'
DATA_FILE = '_data.txt'


class AutoClassChoosing:
    def __init__(self, student_num, password) -> None:
        self.student_num = student_num
        self.password = password
        self.__init_driver__()

    def __init_driver__(self) -> None:
        chrome_option = Options()
        chrome_option.add_argument('--log-level=3')
        self.driver = webdriver.Chrome(
            executable_path=ChromeDriverManager(
                version='114.0.5735.90').install(),
            options=chrome_option
        )

    def run(self, entries) -> int:
        print('Finish the set up. The program will start automatically after countdown.')

        while True:
            login_status = self.login()

            if login_status == 0:  # login success
                print('Login successfully!!')
                break
            elif login_status == 1:  # wrong password or wrong student number
                print('ID or password incorrect!')
                self.student_num = input('Please enter your ID: ')
                self.password = input('Please enter your password: ')
            elif login_status == 2:  # wrong confirm code
                print('Wrong confirm code, program will retry now.')
            elif login_status == 3:  # wrong login time
                if datetime.now() > self.starting_time:
                    print('Time Expired!')

                    return -1

                print(
                    f'The program will start at {datetime.strftime(self.starting_time, "%Y/%m/%d %H:%M:%S")}')

                while not self.clock_on_time():
                    sleep(1)
            else:  # other situations
                print('Login error!')

                return -1

        class_choosing_status = self.choose_classes(entries=entries)

        if class_choosing_status == 0:  # class choosing successfully
            print('The program finished! Log messages are written in result.txt')

            logout_btn = self.driver.find_element(
                By.XPATH, '//*[@id="btnLogout"]')
            logout_btn.click()

            self.driver.close()
        else:
            print('Class choosing error! End up the program.')

            self.driver.close()

            return -1

        return 0

    def clock_on_time(self) -> bool:
        curr_time = datetime.now()
        is_expired = curr_time >= self.starting_time

        return is_expired

    def login(self) -> int:
        self.driver.get(LOGIN_URL_ENG)

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
            print('HTML error!')

            return 4

        confirm_code_input.clear()
        confirm_code_input.send_keys(self.auto_detect_confirm_code())

        login_btn = self.driver.find_element(By.XPATH, '//*[@id="btnLogin"]')
        login_btn.click()

        if self.driver.current_url == TARGET_URL_ENG:
            return 0
        else:
            msg = self.driver.find_element(
                By.XPATH, '//*[@id="TABLE1"]/tbody/tr[6]/td[2]')

            if msg.text == LOGIN_FAIL_ENG:
                return 1
            elif CONFIRM_FAIL_ENG in msg.text:
                return 2
            elif WRONG_TIME_ENG in msg.text:
                msg_text = msg.text.split('\n')
                time_info = msg_text[2].split(' ')
                time = time_info[7] + ' ' + time_info[8]

                self.starting_time = datetime.strptime(
                    time, '%Y-%m-%d %H:%M:%S')

                return 3
            elif MAINTAIN_TIME_ENG in msg.text:
                current_date = datetime.now().date()
                noon_time = Time(hour=12, minute=30)

                self.starting_time = datetime.combine(current_date, noon_time)

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
        ocr = ddddocr.DdddOcr(show_ad=False)
        confirm_code = ocr.classification(img)

        return confirm_code

    def choose_classes(self, entries) -> int:
        with open('result.txt', 'w') as result_file:
            for entry in entries:
                id = entry.value.get()
                line = 'Class ID: ' + id + ' '

                # class id input
                class_id_input = self.driver.find_element(
                    By.XPATH, '//*[@id="txtCosEleSeq"]')
                class_id_input.clear()
                class_id_input.send_keys(id)

                if entry.add_btn_value:
                    # add button click
                    add_btn = self.driver.find_element(
                        By.XPATH, '//*[@id="btnAdd"]')
                    add_btn.click()
                elif entry.drop_btn_value:
                    # del button click
                    drop_btn = self.driver.find_element(
                        By.XPATH, '//*[@id="btnDel"]')
                    drop_btn.click()
                else:
                    print('UI error!')

                    return -1

                msg = self.driver.find_element(
                    By.XPATH, '//*[@id="form1"]/div[3]/table/tbody/tr[2]/td[3]')

                msg_in_line = msg.text.split('\n')

                if ADD_SUCCESS_ENG in msg.text:
                    line += ADD_SUCCESS_ENG
                elif ADD_FAIL_ENG in msg.text:
                    line += (ADD_FAIL_ENG + " -> ")
                    line += msg_in_line[1]
                elif DROP_SUCCESS_ENG in msg.text:
                    line += DROP_SUCCESS_ENG
                elif DROP_FAIL_ENG in msg.text:
                    line += (DROP_FAIL_ENG + " -> ")
                    line += msg_in_line[1]
                else:
                    line += "ERROR"

                result_file.write(line + '\n\n')

        return 0

    def close(self) -> None:
        self.driver.close()


ENGLISH = 'Times New Roman'
CHINESE = '微軟正黑體'


class InputObject:
    def __init__(self, container) -> None:
        self.value = StringVar(container)
        self.label = Label(container)
        self.entry = Entry(container, textvariable=self.value)

        self.add_btn = Button(container)
        self.add_btn_value = True
        self.add_btn.config(text='Add', font=(ENGLISH, 12, 'bold'), foreground='black',
                            height=1, width=5, command=self.add_btn_onclick)

        self.drop_btn = Button(container)
        self.drop_btn_value = False
        self.drop_btn.config(text='Drop', font=(ENGLISH, 12), foreground='gray',
                             height=1, width=5, command=self.drop_btn_onclick)

    def set_label(self, text) -> None:
        self.label.config(text=text, font=(ENGLISH, 14))

    def set_entry(self) -> None:
        self.entry.config(font=(ENGLISH, 14))

    def add_btn_onclick(self) -> None:
        self.add_btn_value = True
        self.add_btn.config(font=(ENGLISH, 12, 'bold'), foreground='black')
        self.drop_btn_value = False
        self.drop_btn.config(font=(ENGLISH, 12), foreground='gray')

    def drop_btn_onclick(self) -> None:
        self.drop_btn_value = True
        self.drop_btn.config(font=(ENGLISH, 12, 'bold'), foreground='black')
        self.add_btn_value = False
        self.add_btn.config(font=(ENGLISH, 12), foreground='gray')

    def place(self, index) -> None:
        self.label.grid(row=index, column=0, padx=15, pady=10)
        self.entry.grid(row=index, column=1, padx=10, pady=10)
        self.add_btn.grid(row=index, column=2, padx=5, pady=10)
        self.drop_btn.grid(row=index, column=3, padx=5, pady=10)

    def destroy(self) -> None:
        self.entry.destroy()
        self.label.destroy()
        self.add_btn.destroy()
        self.drop_btn.destroy()


class MainUI:
    def __init__(self) -> None:
        self.init_main_frame()
        self.init_login_frame()
        self.init_class_id_frame()
        self.init_buttons()

        self.threads = []  # init thread

    def __show_password_onclick__(self):
        if self.show_password_value.get():
            self.password_entry.config(show='')
        else:
            self.password_entry.config(show='\u25CF')

    def init_main_frame(self) -> None:
        self.root = Tk()
        self.root.resizable(False, False)
        self.root.geometry("450x625")
        self.root.title('AutoClassChoosing Set-up')

    def init_login_frame(self) -> None:
        self.login_frame = LabelFrame(self.root)
        self.login_frame.config(text=' Login ', font=(ENGLISH, 12))
        self.login_frame.pack(side=TOP, fill='x', padx=10, pady=10)

        self.student_id_label = Label(self.login_frame)
        self.student_id_label.config(
            text='Student ID', font=(ENGLISH, 14, 'bold'))
        self.student_id_label.pack(side=TOP, pady=5)

        self.student_id = StringVar(self.login_frame)
        self.student_id_entry = Entry(self.login_frame)
        self.student_id_entry.config(
            font=(ENGLISH, 12), textvariable=self.student_id)
        self.student_id_entry.pack(side=TOP, fill='x', padx=5, pady=10)

        self.password_label = Label(self.login_frame)
        self.password_label.config(
            text='Password', font=(ENGLISH, 14, 'bold'))
        self.password_label.pack(side=TOP, pady=5)

        self.password = StringVar(self.login_frame)
        self.password_entry = Entry(self.login_frame)
        self.password_entry.config(
            font=(ENGLISH, 12), textvariable=self.password, show='\u25CF')
        self.password_entry.pack(side=TOP, fill='x', padx=5, pady=10)

        # Set Login Inner Frame

        self.login_inner_frame = Frame(self.login_frame)
        self.login_inner_frame.pack(side=TOP, fill='x')

        self.show_password_value = BooleanVar()
        self.show_password_checkbox = Checkbutton(
            self.login_inner_frame, text='Show Password',
            variable=self.show_password_value, command=self.__show_password_onclick__)
        self.show_password_checkbox.pack(side=LEFT, padx=10)

        self.checkbox_value = BooleanVar()
        self.checkbox = Checkbutton(
            self.login_inner_frame, text='Remember Password', variable=self.checkbox_value)
        self.checkbox.pack(side=RIGHT, padx=10)

        with open(DATA_FILE, 'r') as data_file:
            data = data_file.readline().replace('\n', '').split(' ')

            if data[0] != 'null':
                self.student_id_entry.insert(0, data[0])

            if data[1] != 'null':
                self.password_entry.insert(0, data[1])

            if data[2] == 'True':
                self.checkbox_value.initialize(True)

    def init_class_id_frame(self) -> None:
        self.class_id_frame = LabelFrame(self.root)
        self.class_id_frame.config(text=' Class ID Input ', font=(ENGLISH, 12))
        self.class_id_frame.pack(side=TOP, fill='x', padx=10, pady=10)

        self.inner_frame = Frame(self.class_id_frame)
        self.inner_frame.pack(fill=BOTH, expand=True)

        self.inner_canvas = Canvas(self.inner_frame)
        self.inner_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.inner_frame, orient=VERTICAL,
                                       command=self.inner_canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.inner_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.inner_canvas.bind('<Configure>', lambda e: self.inner_canvas.configure(
            scrollregion=self.inner_canvas.bbox('all')))

        self.window_frame = Frame(self.inner_canvas)

        self.inner_canvas.create_window(
            (0, 0), window=self.window_frame, anchor='nw')
        self.window_frame.bind('<Configure>', self.scrollbar_resize)

        self.entries = [InputObject(self.window_frame)]

    def init_buttons(self) -> None:
        self.add_btn = Button(self.root)
        self.add_btn.config(text='add', font=(ENGLISH, 14, 'bold'),
                            height=2, width=6, command=self.add_btn_onclick)
        self.add_btn.pack(side=LEFT, padx=15)

        self.del_btn = Button(self.root)
        self.del_btn.config(text='del', font=(ENGLISH, 14, 'bold'),
                            height=2, width=6, command=self.del_btn_onclick)
        self.del_btn.pack(side=LEFT, padx=10)

        self.start_btn = Button(self.root)
        self.start_btn.config(text='start', font=(ENGLISH, 14, 'bold'),
                              height=2, width=8, command=self.start_btn_onclick)
        self.start_btn.pack(side=LEFT, padx=10)

        self.quit_btn = Button(self.root)
        self.quit_btn.config(text='quit', font=(ENGLISH, 14, 'bold'),
                             height=2, width=8, command=self.quit_btn_onclick)
        self.quit_btn.pack(side=LEFT, padx=15)

    def place_entries(self):
        for i in range(0, len(self.entries)):
            self.entries[i].set_label('Class ID ' + str(i + 1))
            self.entries[i].place(i)

    def scrollbar_resize(self, event):
        size = (self.window_frame.winfo_reqwidth(),
                self.window_frame.winfo_reqheight())
        self.inner_canvas.config(scrollregion="0 0 %s %s" % size)

        if self.window_frame.winfo_reqwidth() != self.inner_canvas.winfo_width():
            self.inner_canvas.config(width=self.window_frame.winfo_reqwidth())

    def auto_class_choosing(self):
        self.save_data()

        self.bot = AutoClassChoosing(
            student_num=self.student_id.get(),
            password=self.password.get()
        )

        self.bot.run(entries=self.entries)

    def save_data(self):
        bool_value = self.checkbox_value.get()

        if bool_value:
            password = self.password.get()
        else:
            password = 'null'

        with open(DATA_FILE, 'w') as data_file:
            data_file.write(self.student_id.get() + ' ' +
                            password + ' ' + str(bool_value))

    def add_btn_onclick(self):
        idx = len(self.entries)

        self.entries.append(InputObject(self.window_frame))
        self.entries[idx].set_label('Class ID ' + str(idx + 1))
        self.entries[idx].place(idx)

    def del_btn_onclick(self):
        if len(self.entries) <= 1:
            return

        self.entries[-1].destroy()
        self.entries.pop()

    def start_btn_onclick(self):
        self.threads.append(threading.Thread(target=self.auto_class_choosing))
        self.threads[-1].start()

    def quit_btn_onclick(self):
        size = len(self.threads)

        for i in range(0, size):
            self.threads.pop()

        try:
            self.bot.close()
        except Exception as e:
            print('Browser closed.')

        self.root.quit()
        exit(1)

    def run(self) -> None:
        self.place_entries()
        self.root.mainloop()


if __name__ == "__main__":
    main_ui = MainUI()
    main_ui.run()
