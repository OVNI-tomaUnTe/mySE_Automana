# -*- coding: utf-8 -*-

# Programmed  by Mingyu 'Ozone' CUI | DCOE, Sonepar APAC

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import re
import os, sys

import warnings
warnings.filterwarnings("ignore")

import tkinter as tk
from PIL import Image, ImageTk

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import os, sys
def base_path(path):
    if getattr(sys, 'frozen', None):
        basedir = sys._MEIPASS
    else:
        basedir = os.getcwd()
    return os.path.join(basedir, path)

os.chdir(base_path(''))

pic = Image.open('Hagemeyer.png')

print("Warnings surpressed.")

# -----[GUI]-----

class CustomText(tk.Text):
    def __init__(self, *args, **kwargs):
        """自定义多行文本框类，可实时监控变化事件"""
        tk.Text.__init__(self, *args, **kwargs)
        self._orig = self._w + '_orig'
        self.tk.call('rename', self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def _proxy(self, command, *args):
        if command == 'get' and (args[0] == 'sel.first' and args[1] == 'sel.last') and not self.tag_ranges('sel'):
            return
        if command == 'delete' and (args[0] == 'sel.first' and args[1] == 'sel.last') and not self.tag_ranges('sel'):
            return
        cmd = (self._orig, command) + args
        result = self.tk.call(cmd)
        if command in ('insert', 'delete', 'replace'):
            self.event_generate('<<TextModified>>')
        return result

class interface():
    def __init__(self, master, bg, ttl, btn_cap, email, password):
        
        self.master = master
        self.master.config(bg=bg)
        self.master.title(ttl)
        self.master.geometry('600x500')
        
        self.entry_var0 = tk.StringVar()
        self.entry_var1 = tk.StringVar()
        self.entry_var2 = tk.StringVar()
        self.entry_var3 = tk.IntVar()
        
        self.div0 = tk.Frame(self.master)
        self.div0.pack(side = 'top')
        self.div1 = tk.Frame(self.master)
        self.div1.pack(side = 'top')
        self.sep = tk.Frame(self.master, height = 10)
        self.sep.pack(side = 'top')
        self.div2 = tk.Frame(self.master)
        self.div2.pack(side = 'top')
        self.sep = tk.Frame(self.master, height = 20)
        self.sep.pack(side = 'top')
        self.div3 = tk.Frame(self.master)
        self.div3.pack(side = 'top')
        self.div4 = tk.Frame(self.master)
        self.div4.pack(side = 'top')
        self.div5 = tk.Frame(self.master)
        self.div5.pack(side = 'bottom')
        
        def on_modify(event):
            chars = event.widget.get('1.0', 'end-1c')
            char_counter.configure(text='%s chars' % len(chars))
            self.entry_var2.set(chars)
        
        btn_cmd = self.process
        
        self.label = tk.Label(self.div0, text = sticknotes, fg = '#10387d', width=400, height=2)
        self.label.pack(side = 'top')
        
        tk.Label(self.div1, text = label_email).pack(side = 'left', expand = 'no')
        input_email = tk.Entry(self.div1, width = 40, textvariable = self.entry_var0)
        input_email.pack(side = 'left', expand = 'no')
        self.entry_var0.set(email)
        
        tk.Label(self.div2, text = label_password).pack(side = 'left', expand = 'no')
        input_password = tk.Entry(self.div2, width = 40, textvariable = self.entry_var1, show = '*')
        input_password.pack(side = 'left', expand = 'no')
        self.entry_var1.set(password)
        
        search_annotation = tk.Label(self.div3, text = annotation, fg = '#10387d', anchor = 'w')
        search_annotation.pack(side = 'top', fill = 'x')
        search_capture = CustomText(self.div3, height = 10)
        search_capture.pack()
        search_capture.insert(1.0, current_search)
        search_capture.focus()
        search_capture.bind('<<TextModified>>', on_modify)
        char_counter = tk.Label(self.div3, anchor = 'w')
        char_counter.pack(side = 'bottom', fill = 'x')
        
        exception_annotation = tk.Label(self.div4, text = exceptions, fg = '#b80101')
        exception_annotation.pack(side = 'top', fill = 'x')
        tk.Label(self.div4, text = label_modulus).pack(side = 'left', expand = 'no')
        input_modulus = tk.Entry(self.div4, width = 10, textvariable = self.entry_var3)
        input_modulus.pack(side = 'left', expand = 'no')
        self.entry_var3.set(seg_modulus)
        
        tk.Button(self.div5,text=btn_cap, font=('Arial', 12, 'bold'), command=btn_cmd, width=20, height=2).pack(side = 'top')
        
        tk.Label(self.div5, image = trademark, text = footnote, font=('Arial', 6), compound='left', width=800, height=38).pack(side = 'bottom')
        
    def process(self):
        
        current_email = self.entry_var0.get()
        current_password = self.entry_var1.get()
        current_search = self.entry_var2.get()
        current_search = current_search.strip()
        current_search = current_search.replace(',', '')
        current_search = current_search.replace('.00', '')
        current_search = re.sub('\t', ',', current_search)  #\s: space, \t: tab, \n: enter.
        current_search = re.sub('\s\n', '\n', current_search)
        current_search = re.sub('[\n]+', '\n', current_search)  # Delete extra null lines.
        
        current_modulus = self.entry_var3.get()
        login_credential.append(current_email)
        login_credential.append(current_password)
        
        # Initiate a Chrome browser driver & login with default or user-entered credentials.
        
        driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
        driver.get('https://myse.schneider-electric.com.cn/mySchneider/#!/login')
        driver.maximize_window()
        time.sleep(3)
        login_tag = driver.find_element_by_xpath('//*[@id="login-idms-button"]')
        login_tag.click()
        time.sleep(3)
        
        email_path = driver.find_element_by_xpath('//*[@id="usernameMMM"]')
        next_button = driver.find_element_by_xpath('//*[@id="next_btn"]')
        
        email_path.send_keys(current_email)
        time.sleep(1)
        next_button.click()
        time.sleep(3)
        
        password_path = driver.find_element_by_xpath('//*[@id="userPswd"]')
        login_button = driver.find_element_by_xpath('//*[@id="login_btn"]')
        
        password_path.send_keys(current_password)
        time.sleep(1)
        login_button.click()
        time.sleep(3)
        
        batch_search = driver.find_element_by_xpath('//*[@id="wrap-content-v1"]/div/div/div[1]/div/myse-panda-home/form/div/div[1]/span/a')
        
        batch_search.click()
        time.sleep(3)
        
        local_source = driver.find_element_by_xpath('//*[@id="importIntoPALink"]')
        
        local_source.click()
        time.sleep(3)
        
        # Perform multiple search split before next steps (on batch search page).
        
        search_list = []
        n_counter = 0
        seg_counter = 0
        for i in current_search:
            if i == '\n':
                n_counter += 1
        seg_counter = (n_counter // current_modulus) + 1 # JupyterLab Python 3.5: '%': Remainder; '//': Modulus.
        seg_remainder = (n_counter % current_modulus)
        print('To split list into seg modulo: ', current_modulus)
        print('Total product line received: ', n_counter + 1)
        print('Total querries to be executed: ', seg_counter)
        
        if seg_counter == 1:
            search_list.append(current_search)
            
            # CODE BLOCK ZERO.
            
            input_search = driver.find_element_by_xpath('//*[@id="cutAndPasteText"]')
            input_confirm = driver.find_elements_by_xpath('//*[@id="importpa_importCutAndPaste"]')
            
            input_search.click()
            input_search.send_keys(search_list[0])
            print('Content pasted to Entry Box.')
            time.sleep(1)
            
            for i in input_confirm:
                i.click()
                print(i)
            print('ADD button clicked.')
            time.sleep(1)
            
            for i in range(3):
                driver.find_element_by_tag_name('body').send_keys(Keys.END)
                time.sleep(1)
                i += 1
            print('Page bottom reached.')
            driver.find_element_by_xpath('//*[@id="paSearchButton"]/span').click()
            time.sleep(10)
            print('SEARCH button clicked.')
            print('ETA 20 to next action.')
            time.sleep(20)
            download_tag = driver.find_element_by_xpath('//*[@id="main-div"]/div[3]/div[7]/div[1]/a')
            time.sleep(1)
            download_tag.click()
            print('DOWNLOAD button clicked')
            time.sleep(1)
            print('Cool Down ETA 60.')
            time.sleep(60)
            
            pass
        else:
            for i in range(seg_counter):
                print(i)
                j = i * current_modulus
                str = ''
                if i == seg_counter - 1:
                    while j < (i * current_modulus + seg_remainder):
                        str += current_search.split('\n')[j] + '\n'
                        j += 1
                    search_list.append(str)
                else:
                    while j < (i * current_modulus + current_modulus):
                        str += current_search.split('\n')[j] + '\n'
                        j += 1
                    search_list.append(str)
            list_range = range(len(search_list))
            print('Variable "search_list" range: ', list_range)
            print(search_list)
            
            for s in list_range:
                
                print('-----[ CURRENT LIST RANGE #: ', s, ' ]-----')
                
                if s == 0:
                    pass
                else:
                    # Loop to search remaining segments.
                    
                    print('[CURRENT SEARCH LOOP ', s, ']')
                    
                    driver.back()
                    print('BACK IN TIME.')
                    
                    time.sleep(5)
                    for i in range(5):
                        driver.find_element_by_tag_name('body').send_keys(Keys.END)
                        time.sleep(1)
                        i += 1
                    print('Page bottom reached.')
                    driver.find_element_by_xpath('//*[@id="pa_clear_form"]/span').click()
                    print('CLEAR button clicked.')
                    time.sleep(2)
                    driver.find_element_by_xpath('//*[@id="importIntoPALink"]').click()
                    time.sleep(2)
                
                # CODE BLOCK ZERO.
                
                input_search = driver.find_element_by_xpath('//*[@id="cutAndPasteText"]')
                input_confirm = driver.find_elements_by_xpath('//*[@id="importpa_importCutAndPaste"]')
                
                input_search.click()
                input_search.send_keys(search_list[s])
                print('Content pasted to Entry Box.')
                time.sleep(1)
                
                for i in input_confirm:
                    i.click()
                    print(i)
                print('ADD button clicked.')
                time.sleep(1)
                
                for i in range(3):
                    driver.find_element_by_tag_name('body').send_keys(Keys.END)
                    time.sleep(1)
                    i += 1
                print('Page bottom reached.')
                driver.find_element_by_xpath('//*[@id="paSearchButton"]/span').click()
                time.sleep(20)
                print('Search button clicked.')
                print('ETA 20 to next action.')
                time.sleep(20)
                download_tag = driver.find_element_by_xpath('//*[@id="main-div"]/div[3]/div[7]/div[1]/a')
                time.sleep(1)
                download_tag.click()
                print('DOWNLOAD button clicked.')
                time.sleep(1)
                print('Cool Down ETA 60.')
                time.sleep(60)
            
        print('\n[END OF LINE]\n')
        print(poem)
        self.master.destroy()

# INSTANTIATE GUI

global current_email
global current_password
global current_search
global current_modulus

current_email = ''
current_password = ''
current_search = ''

seg_modulus = 50

current_modulus = seg_modulus

sheet_names = ['mySE货期查询自动化程序']
default_credential = ['amy.song@hagemeyercn.com','hag.8888']
login_credential = []
colors = ['#EEEEEE']
button_caps = ['登录mySE查询 >>']
sticknotes = '输入登录邮箱、密码；按提示要求将查询信息粘贴至输入框，最后点击按钮，等待系统执行自动查询。'
annotation = '请从Excel复制粘贴，或按以下格式在此输入待查询的物料订货号和数量（英文半角逗号分隔，每行一组）：\n(订货号), (数量)\nABFH14H020,5\nQO120M100C,10'
exceptions = '若mySE网页超过10秒显示加载图标或无动作，请关闭浏览器，再次点击查询按钮。'
footnote = 'Developed by:\n   Hagemeyer | DCOE Sonepar\n Yiping LU | Ozone CUI'
label_email = '登录邮箱：'
label_password = '登录密码：'
label_modulus = '自动分组查询产品数（50以下网页下载文件到本地；超过50发送至登录邮箱）：'
poem = 'I would rather be ashes than dust!\nI would rather that my spark should burn out in a brilliant blaze than it should be stifled by dry-rot.\nI would rather be a superb meteor, every atom of me in magnificent glow, than a sleepy and permanent planet.\nThe proper function of man is to live, not to exist.\nI shall not waste my days in trying to prolong them.\nI shall use my time.'

print('GUI loaded: Ready to launch mySE querry.')

if __name__ == '__main__':
    root = tk.Tk()
    trademark = ImageTk.PhotoImage(pic.resize((180, 30)))
    interface(root, colors[0], sheet_names[0], button_caps[0], default_credential[0], default_credential[1])
    root.mainloop()