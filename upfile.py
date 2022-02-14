from tkinter import *
from tkinter import filedialog
from selenium import webdriver
# from webdriver_manager.chrome import ChromeDriverManager
import os
import time
import re
import requests
import zipfile

class YunTG():
    def __init__(self,driver):
        self.driver = driver
    # 登录云通关
    def login(self,personname,password):
        #打开浏览器
        self.driver.get('http://yun.etongguan.com/LogIn.aspx')
        # self.driver.implicitly_wait(5)
        self.driver.find_element_by_id('PersonName').send_keys(personname)
        self.driver.find_element_by_id('Password').send_keys(password)
        self.driver.find_element_by_xpath('//button').click()  #登录
        # self.driver.implicitly_wait(5)

    #查询流水号
    def queryNumber(self,customNumber):
        time.sleep(1)
        self.driver.get(f'http://yun.etongguan.com/wdbg/DecQuery2019.aspx?v=VjUwMTM2Mg==')
        # self.driver.find_element_by_xpath('''//a[@onclick="OpenNew('wdbg')"]''').click() #我的通关
        self.driver.implicitly_wait(5)
        self.driver.find_element_by_id('txtEntry_id').send_keys(customNumber) #输入报关单号
        self.driver.find_element_by_id('btnSearchOrder').click() #查询
        self.driver.implicitly_wait(5)
        try:
            return self.driver.find_element_by_xpath("//table[@id='fixedTableBodyColumns']/tbody/tr/td/a").text #流水号
        except:
            return 0

    #上传留底
    def upload(self,run_number,filename,temp):
        time.sleep(1)
        self.driver.get(f'http://yun.etongguan.com/wztg/Input_JKDeclInfo.aspx?serialNumber={run_number}001&depNo=001002008&userId=6929&type=1')
        # self.driver.find_element_by_xpath("//table[@id='fixedTableBodyColumns']/tbody/tr/td/a").click()
        self.driver.implicitly_wait(5)
        # self.driver.switch_to_window(self.driver.window_handles[1])
        # self.driver.implicitly_wait(5)
        self.driver.find_element_by_xpath('//li[@class="current"]').click() #单证仓库
        self.driver.implicitly_wait(5)
        self.driver.find_element_by_xpath('//button').click() #下拉
        self.driver.find_element_by_xpath('//div[@class="dropdown-menu open"]/ul/li[@data-original-index="58"]').click() #选择留底
        self.driver.implicitly_wait(5)
        self.driver.find_element_by_xpath('//input[@multiple="multiple"]').send_keys(rf'{filename}/{temp}') #输入文件地址
        self.driver.implicitly_wait(5)

    #关闭浏览器
    def closed(self):
        time.sleep(5)
        self.driver.close()

class MY_GUI():
    def __init__(self,window_name):
        self.window_name = window_name
#         self.text_2 = '''1、账号密码为云通关账号\n
# 2、输入留底资料所在文件夹\n
# 3、留底文件名为报关单号\n
# 4、请确认谷歌浏览器与chromedriver是否匹配\n
# 5、chromedriver如已加入环境变量请忽略第六条\n
# 6、输入chromedriver地址\n
# 7、有头模式为浏览器展示所有操作步骤\n
# 8、建议有头模式，无头会卡\n'''
        # self.option = None

    #视图界面
    def set_window(self):
        self.window_name.title('自动化上传 V2.1')
        self.window_name.minsize(500,450)
        self.window_name.maxsize(500,450)
        # self.window_name.iconbitmap(default=r'C:\Users\ASUS\Documents\Exercise\venv\upfile\flash.ico')
        # self.text_instruction = Text(self.window_name,bg='#efefef')
        # self.text_instruction.insert(END,self.text_2)
        # self.text_instruction.place(x=469,y=16,width=264,height=264)
        # self.text_instruction.configure()
        self.text_output = Text(self.window_name)
        self.text_output.place(x=35,y=185,width=430,height=245)
        self.label_personname = Label(self.window_name,text='账号:')
        self.label_personname.place(x=35,y=25,width=75,height=30)
        self.label_password = Label(self.window_name,text='密码:')
        self.label_password.place(x=35,y=67,width=75,height=30)
        self.label_path_1 = Label(self.window_name,text='留底路径:')
        self.label_path_1.place(x=35,y=135,width=75,height=30)
        # self.label_path_2 = Label(self.window_name,text='组件路径:')
        # self.label_path_2.place(x=35,y=200,width=75,height=30)
        self.entry_personname = Entry(self.window_name)
        self.entry_personname.place(x=108,y=25,width=94,height=28)
        self.entry_password = Entry(self.window_name)
        self.entry_password.place(x=108,y=67,width=94,height=28)
        self.entry_path_1 = Entry(self.window_name)  #留底路径
        self.entry_path_1.place(x=110,y=135,width=247,height=30)
        # self.entry_path_2 = Entry(self.window_name)  #chromedriver路径
        # self.entry_path_2.place(x=110,y=200,width=247,height=30)
        self.button_path_1 = Button(self.window_name,text='浏览',command=self.click_1)
        self.button_path_1.place(x=387,y=135,width=64,height=30)
        # self.button_path_2 = Button(self.window_name,text='浏览',command=self.click_2)
        # self.button_path_2.place(x=387,y=200,width=64,height=30)
        self.button_path_3 = Button(self.window_name,text='开始',command=self.start)
        self.button_path_3.place(x=355,y=40,width=100,height=45)
        # self.radiobutton_1 = Radiobutton(self.window_name,text='有头模式',value=1,command=self.radio_1)
        # self.radiobutton_1.place(x=495,y=310,width=100,height=35)
        # self.radiobutton_1 = Radiobutton(self.window_name,text='无头模式',value=2,command=self.radio_2)
        # self.radiobutton_1.place(x=635,y=310,width=100,height=35)

    #提取留底文件路径
    def click_1(self):
        filename_1 = StringVar()
        filename_1 = filedialog.askdirectory()
        self.entry_path_1.delete(0,END)
        self.entry_path_1.insert(0,filename_1)

    #提取chromedriver文件路径
    def click_2(self):
        filename_2 = StringVar()
        filename_2 = filedialog.askopenfilename()
        self.entry_path_2.delete(0,END)
        self.entry_path_2.insert(0,filename_2)

    #对text_1进行输出展示
    def get_text(self,var):
        self.text_output.insert(END,f'{var}\n')

    #单选框1
    # def radio_1(self):
    #     self.option = webdriver.ChromeOptions()
    #     self.option.add_argument('User-Agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36 Edg/96.0.1054.62"')

    #单选框2
    # def radio_2(self):
    #     self.option = Options()
    #     self.option.add_argument('headless')

    #下载驱动
    def install(self):
        #查询代码地址
        url_2 = r'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'        
        #下载地址
        url = rf'https://oss.npmmirror.com/dist/chromedriver/{requests.get(url_2).text}/chromedriver_win32.zip'
        r = requests.get(url)
        with open(r'chromedriver.zip','wb') as f:
            f.write(r.content)
        zip_file = zipfile.ZipFile('chromedriver.zip')
        zip_list = zip_file.namelist()
        for f in zip_list:
            zip_file.extract(f,r'.\chromedriver')
        zip_file.close()
    
    #开始按钮
    def start(self):
        #提取文件名中的报关单号
        self.install()
        files_name = os.listdir(self.entry_path_1.get())
        self.get_text(f'共计有{len(files_name)}个文件')
        dict_temp = list()
        for i in files_name:
            for j in re.findall('\d{18,18}',i):
                dict_temp.append([j,i])
        self.get_text(f'其中报关单号有{len(dict_temp)}条')
        #输入账号密码
        personname = self.entry_personname.get()
        password = self.entry_password.get()
        # temp_ = self.entry_path_2.get()
        # if len(temp_) > 0:
            # options = webdriver.Chrome(chrome_options=self.option,executable_path=self.entry_path_2.get())
        # else:
        options = webdriver.Chrome(executable_path='.\chromedriver\chromedriver.exe')
        options.maximize_window()
        driver = YunTG(options)
        driver.login(personname,password)
        count = 0
        for customNumber,custom_file in dict_temp:
            run_number = driver.queryNumber(customNumber)
            if run_number == 0:
                self.get_text(f'{customNumber}未查询到流水号。。。')
                continue
            driver.upload(run_number,self.entry_path_1.get(),custom_file)
            count += 1
            self.get_text(f'已上传{count}条')
        self.get_text(f'剩余{len(dict_temp)-count}条未上传')
        driver.closed()

if __name__ == '__main__':
    root = Tk()
    window = MY_GUI(root)
    window.set_window()
    root.mainloop()