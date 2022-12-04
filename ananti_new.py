from selenium import webdriver
from selenium.common.exceptions import (ElementNotVisibleException, ElementNotSelectableException, StaleElementReferenceException, SessionNotCreatedException, WebDriverException)

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

import time
import datetime

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import font

import numpy as np
import pandas as pd
import cv2
import threading

ignore_list = [ElementNotVisibleException, ElementNotSelectableException, StaleElementReferenceException]   # wait에서 제외할 예외들

def th():
    """ 메인 프로그램을 실행할 쓰레드 """
    th = threading.Thread(target=run_program)
    th.daemon = True
    th.start()


def select_reserve_mode():
    """ 예약 모드 선택시 실행하는 함수 """
    global test
    
    window.title('Ananti Reservation(예약 모드)')
    test = 0
    
    log.insert(tk.END, '예약 모드를 선택했습니다.\n')
    

def select_test_mode():
    """ 테스트 모드 선택시 실행하는 함수 """
    global test
    
    window.title('Ananti Reservation(테스트 모드)')
    test = 1
    
    log.insert(tk.END, '테스트 모드를 선택했습니다.\n')


def select_id_pw(event):
    """ 예약자를 선택하는 함수 """
    global user
    global id
    global pw 
    
    user = combo_user.get()
    if user == '박태종':
        id = "2511167621"
        pw = "600315"
    elif user == '김양미':
        id = "2511167600"
        pw = "680615"
    elif user == '박기택':
        id = "2511167622"
        pw = "930216"
    elif user == '박준택':
        id = "2511167623"
        pw = "950331"
        
    if test: log.insert(tk.END, 'user : ' + user + ',  id : ' + id + ',  pw : ' + pw + '\n')
    

def select_place():
    """ 가평/부산/남해 선택 """
    global selected_place
    global combo_reserve_type
    global place_info
    global platform_value
    global h
    
    # 장소를 선택하면 예약 타입 선택을 초기화
    for widgets in frame_reserve_type.winfo_children():
          widgets.destroy()
          
    # 장소를 선택하면 룹타입을 초기화
    for widgets in frame_roomtype.winfo_children():
          widgets.destroy()
    
    # 예약 타입 선택 combo box
    combo_reserve_type = ttk.Combobox(frame_reserve_type, width=12, font=style)
    combo_reserve_type['values'] = ['일반예약(4주)', '특별예약(8주)', '특별예약(12주)']
    combo_reserve_type.pack(padx=5, pady=5)
    combo_reserve_type.bind('<<ComboboxSelected>>', select_reserve_type)

    # 장소 선택에 따라 정보 선택
    selected_place = var_place.get()
    if selected_place == 1: 
        if test: log.insert(tk.END, '가평 아난티를 선택했습니다.\n')
        platform_value = '24'
        place_info = '가평'
        h = 10
    if selected_place == 2: 
        if test: log.insert(tk.END, '부산 아난티를 선택했습니다.\n')
        platform_value = '25'
        place_info = '부산'
        h = 10
    if selected_place == 3: 
        if test: log.insert(tk.END, '남해 아난티를 선택했습니다.\n')
        platform_value = '12'
        place_info = '남해'
        h = 9
        

def select_reserve_type(event):
    """ 예약 타입 선택 """
    global reserve_type
    global combo_roomtype
    global combo_reserve_type
        
    # 예약 타입을 선택하면 날짜, 숙박일수, 인원 선택 활성화
    combo_year.configure(state='enabled')
    combo_month.configure(state='enabled')
    combo_day.configure(state='enabled')
    
    combo_nights.configure(state='enabled')
    combo_guests.configure(state='enabled')
    
    select_date(0)
    select_nights(0)
    select_guests(0)
    
    
    # 예약 타입을 선택하면 룸타입을 초기화
    for widgets in frame_roomtype.winfo_children():
            widgets.destroy()
        
    # 예약 타입과 장소에 따라 룸타입 출력
    reserve_type = combo_reserve_type.get()
    
    if reserve_type == '일반예약(4주)':
        if test: log.insert(tk.END, 'reserve type : ' + reserve_type + '\n')
        
        combo_roomtype = ttk.Combobox(frame_roomtype, width=33, font=style)
        # 가평 
        if selected_place == 1: 
            combo_roomtype['values'] = ["테라스 하우스 (트윈+트윈)", "테라스 하우스 (킹+트윈)", "무라타 하우스 (킹)", "풀 하우스 (킹)", "더 하우스 일반 (킹2)", "더 하우스 확장 (킹2+트윈)", "스위트 (킹2+트윈)"]
        # 부산 
        elif selected_place == 2: 
            combo_roomtype['values'] = ["테라스풀하우스 A (트윈2)", "테라스풀하우스 D (킹2)", "테라스풀하우스 H (온수풀/트윈2)", "테라스풀하우스 B (복층/트윈2)", 
                                        "레지던스 오션 A (트윈-거실일체형)", "레지던스 오션 B (트윈-거실분리형)", "레지던스 오션 C (트윈-커넥팅 불가)", "레지던스 오션 C (킹-커넥팅 불가)", "레지던스 오션 (트윈, 휠체어 이용)", 
                                        "레지던스 마운틴 A (트윈-거실일체형)", "레지던스 마운틴 B (트윈-거실분리형)", "레지던스 마운틴 C (킹-커넥팅 불가)", "커넥팅 마운틴 (트윈 + 트윈)", 
                                        "패밀리 마운틴 (트윈 + 킹 + 온돌)", 
                                        "Seaside 단층 (트윈 + 트윈)", "Seaside 복층 (트윈 + 트윈 + 트윈)", "Seaside Family (트윈 + 트윈)", "Seaside Special (트윈 + 트윈 + 트윈 + 트윈)"]
        # 남해 
        elif selected_place == 3:
            combo_roomtype['values'] = ["펜트하우스A (킹+트윈)", "펜트하우스A (트윈+트윈)", "펜트하우스A (트윈+트윈 -1시 체크아웃 / 저층)", "휠체어 이용 펜트하우스A (킹+트윈)", "펫룸 펜트하우스A (트윈+트윈)",
                                        "펜트하우스B (52평형 트윈+트윈)", "펜트하우스B (52평형 킹+트윈)",
                                        "펜트하우스C (35평형 킹)", "펜트하우스C (35평형 트윈)", "펜트하우스C (트윈 -1시 체크아웃 / 저층)",
                                        "더하우스 풀 (78평형)", "더하우스 가든 (78평형)"]
            
        combo_roomtype.pack(padx=5, pady=5)
        combo_roomtype.bind('<<ComboboxSelected>>', select_roomtype)
        
    elif reserve_type == '특별예약(8주)':
        if test: log.insert(tk.END, 'reserve type : ' + reserve_type + '\n')
            
        combo_roomtype = ttk.Combobox(frame_roomtype, width=33, font=style)
        # 가평
        if selected_place == 1: 
            combo_roomtype['values'] = ["테라스 하우스 (트윈+트윈)", "테라스 하우스 (킹+트윈)"]
        # 부산
        elif selected_place == 2: 
            combo_roomtype['values'] = ["테라스풀하우스 A (트윈2)", "테라스풀하우스 D (킹2)",
                                        "레지던스 오션 A (트윈-거실일체형)", "레지던스 오션 B (트윈-거실분리형)", "레지던스 오션 C (트윈-커넥팅 불가)", "레지던스 오션 C (킹-커넥팅 불가)", "커넥팅 오션 (트윈 + 트윈)",
                                        "레지던스 마운틴 A (트윈-거실일체형)", "레지던스 마운틴 B (트윈-거실분리형)", "레지던스 마운틴 C (킹-커넥팅 불가)", "커넥팅 마운틴 (트윈 + 트윈)"]
        # 남해
        elif selected_place == 3:
            combo_roomtype['values'] = ["펜트하우스A (킹+트윈)", "펜트하우스A (트윈+트윈)", "펜트하우스A (트윈+트윈 -1시 체크아웃 / 저층)"]
            
        combo_roomtype.pack(padx=5, pady=5)
        combo_roomtype.bind('<<ComboboxSelected>>', select_roomtype)
        
    elif reserve_type == '특별예약(12주)':
        if test: log.insert(tk.END, 'reserve type : ' + reserve_type + '\n')
            
        combo_roomtype = ttk.Combobox(frame_roomtype, width=33, font=style)
        # 가평
        if selected_place == 1: 
            combo_roomtype['values'] = ["테라스 하우스 (트윈+트윈)", "테라스 하우스 (킹+트윈)"]
        # 부산
        elif selected_place == 2: 
            combo_roomtype['values'] = ["테라스풀하우스 A (트윈2)", "테라스풀하우스 D (킹2)",
                                        "레지던스 오션 A (트윈-거실일체형)", "레지던스 오션 B (트윈-거실분리형)", "레지던스 오션 C (트윈-커넥팅 불가)", "레지던스 오션 C (킹-커넥팅 불가)", "커넥팅 오션 (트윈 + 트윈)",
                                        "레지던스 마운틴 A (트윈-거실일체형)", "레지던스 마운틴 B (트윈-거실분리형)", "레지던스 마운틴 C (킹-커넥팅 불가)", "커넥팅 마운틴 (트윈 + 트윈)"]
        # 남해
        elif selected_place == 3:
            combo_roomtype['values'] = ["펜트하우스A (킹+트윈)", "펜트하우스A (트윈+트윈)"]
            
        combo_roomtype.pack(padx=5, pady=5)
        combo_roomtype.bind('<<ComboboxSelected>>', select_roomtype)
        

def select_roomtype(event):
    """ 룸타입 선택 """
    global roomtype
    
    roomtype = combo_roomtype.get()
    
    if test: log.insert(tk.END, 'room type : ' + roomtype + '\n')
    
    
def select_date(event):
    """ 날짜 선택 """
    global checkin
    
    checkin = '%d%02d%02d' % (int(combo_year.get()), int(combo_month.get()), int(combo_day.get()))
    select_nights(0) # check-in 날짜를 선택하면 숙박일수를 이용해서 check-out 날짜를 자동 선택
    

def select_nights(event):
    """ 숙박일수 선택 """
    global checkout
    global checkin
    global combo_nights
    global date_info
    
    # 숙박일수를 선택한 것을 입력받음
    days_name = ['(월)', '(화)', '(수)', '(목)', '(금)', '(토)', '(일)']
    nights = int(combo_nights.get())

    # 체크인 날짜로 date object 생성
    checkin_str = "{}-{}-{}".format(checkin[0:4], checkin[4:6], checkin[6:8])
    checkin_obj = datetime.datetime.strptime(checkin_str, "%Y-%m-%d")
    checkin_info = checkin_str + days_name[checkin_obj.weekday()]
    
    # 체크인 date object에 숙박일수를 더해서 체크아웃 날짜를 결정
    checkout_obj = checkin_obj + datetime.timedelta(nights-1)
    checkout_str = '{}-{}-{}'.format(checkout_obj.year, checkout_obj.month, checkout_obj.day)
    checkout_info = checkout_str + days_name[checkout_obj.weekday()]
    
    checkout = '%d%02d%02d' % (checkout_obj.year, checkout_obj.month, checkout_obj.day)
    nights_info = "{}박 {}일".format(nights-1, nights)
    
    date_info = checkin_info + '~' + checkout_info + '({})'.format(nights_info)
    
    if test: log.insert(tk.END, 'checkin/out : ' + date_info + '\n')
    

def select_guests(event):
    """ 숙박인원 선택 """
    global combo_guests
    global guests
    global guests_info
    
    guests = int(combo_guests.get())
    guests_info = str(guests) + '명'
    
    if test: log.insert(tk.END, 'guests : ' + str(guests) + '명\n')


def confirm():
    """ 예약 확정 """
    global date_info
    global guests
    global user
    global place_info
    global reserve_type
    global roomtype
    global guests_info

    try:
        reservation_info = '===== 예약 정보를 확인하세요 =====' \
                            + '\n예 약 자 : ' + user \
                            + '\n장    소 : ' + place_info \
                            + '\n예약타입 : ' + reserve_type \
                            + '\n룸 타 입 : ' + roomtype \
                            + '\n날    짜 : ' + date_info\
                            + '\n인    원 : ' + guests_info\
                            + '\n========================='       
    except NameError:
        messagebox.showwarning(title='ERROR', message='모든 항목을 선택하세요.')
    else:
        answer = messagebox.askquestion(title='CONFIRM', message=reservation_info)
        if answer == 'yes':
            th()
        elif answer:
            log.insert(tk.END, '취소하셨습니다.\n')


def load_font():
    """ 미리 저장된 Captcha 폰트 로드 """
    font1 = []
    font2 = []
    font3 = []
    for i in range(10):
        tmp = pd.read_excel("font1.xlsx", sheet_name=i).to_numpy()
        font1.append(tmp)
        tmp = pd.read_excel("font2.xlsx", sheet_name=i).to_numpy()
        font2.append(tmp)
        tmp = pd.read_excel("font3.xlsx", sheet_name=i).to_numpy()
        font3.append(tmp)
    # log.insert(tk.END, "status message : font loaded\n")
    
    return font1, font2, font3

  
def SolveCaptcha(font1, font2, font3, img):
    """ Captcha 문자 해독 코드 """
    # 이미지에서 글자를 1, 배경을 0로 변환
    img[img>0] = 255
    img[img==0] = 1
    img[img==255] = 0
    [h0, w0] = img.shape
    
    captcha_txt = ""
    cand = []   # Captcha 문자의 후보들을 저장하는 배열
    info = []
    
    # i는 Capthca에서 찾을 숫자 (0 ~ 9)
    for i in range(10):
        [h1, w1] = font1[i].shape
        [h2, w2] = font2[i].shape
        [h3, w3] = font3[i].shape
    
        # font에 저장된 각 숫자의 pixel을 이미지에서 한 칸씩 shift하며 xor 연산
        # 이때 xor 결과가 가장 작은(가장 겹치는 부분이 많은) 숫자를 Captcha 숫자 후보로 선정
        j = 0
        while j < w0-30:
            xor1 = (img[0:h0, j:j+w1] + font1[i]) % 2
            xor2 = (img[0:h0, j:j+w2] + font2[i]) % 2
            xor3 = (img[0:h0, j:j+w3] + font3[i]) % 2
            sum1 = np.sum(xor1)
            sum2 = np.sum(xor2)
            sum3 = np.sum(xor3)
        
            m = min(sum1, sum2, sum3)
            if m < 100:
                cand.append([m, i, j])
            j = j + 1
            
        cand.sort(key=lambda x:x[2])    # 각 후보의 x좌표 순서대로 정렬
    cand = np.array(cand)  
    
    # pixel을 한 칸씩 shift 하다보면 숫자가 정확히 일치한 것이 아니라 이미지 상의 두 숫자 사이에 걸쳤는데도 xor 결과가 낮아서 cand에 포함된 경우가 있음
    # 그런 경우를 제거하기 위해 이미지의 숫자들 중 첫 번째부터 시작해서 본인과 바로 다음 숫자 사이의 좌표를 가지는 후보를 제거
    while cand[0, 2] + 15 < w0:
        sub_cand = cand[cand[:, 2]<cand[0, 2]+15]               # cand의 첫 번째 숫자의 좌표 + 15(숫자 하나의 폭이 15 이상임)보다 작은 좌표를 가지는 cand를 sub_cand에 저장
        cand = np.delete(cand, list(range(len(sub_cand))), 0)   # cand에서는 삭제(즉, 해당 cand들을 cand -> sub_cand로 이동한 것)
        sub_cand = sub_cand.tolist()
        sub_cand.sort()                                         # sub_cand에서 pixel 수가 더 낮은(실제 숫자와 겹쳤을 확률이 더 높은) 순서로 정렬
        info.append(sub_cand[0])                                # 둘 중 더 낮은 것을 실제 Captcha로 판단하고 info에 추가
        if len(cand) == 0:
            break
    log.insert(tk.END, "status message : captcha info = {}".format(info) + '\n')
    
    # 혹시라도 잘못된 정보를 덜 걸렀으면 다시 걸러주는 과정 수행
    if len(info) > 6:
        log.insert(tk.END, "status message : too many numbers, %d\n" % len(info))
        i = 0
        while i < len(info) - 1:
            # print(i)
            if info[i+1][2] - info[i][2] < 10:                  # info에 들어간 숫자들의 좌표 차이가 10이하로 나는 경우(숫자 폭이 15 이상이므로 10 차이가 날 수 없음)
                if info[i][0] < info[i+1][0]:                   # 둘의 pixel 수가 더 낮은 것을 실제 Captcha 정보로 판단하고 다른 것을 삭제
                    del info[i+1]
                else:
                    del info[i]
            i = i + 1
        log.insert(tk.END, "status message : reduced info = " + info + '\n')
                      
    # 남은 6자리 숫자를 Captcha text로 판단
    for i in range(6):                                        
        captcha_txt = captcha_txt + str(info[i][1])
    
    return captcha_txt


def run_program():
    """ 메인 예약 프로그램 """
    global id
    global pw
    
    
    # 크롬 드라이버 확인
    log.insert(tk.END, "status message : checking 'chromedriver.exe'\n")
    try:
        driver = webdriver.Chrome()
        url = "https://www.google.com"
        driver.get(url)
    except SessionNotCreatedException:
        messagebox.showwarning(title='ERROR', message='chromedriver.exe 버전을 확인하세요.')
        log.insert(tk.END, "exception message : 'chromedriver.exe' version error\n")
        return -1
    except WebDriverException:
        messagebox.showwarning(title='ERROR', message='chormedriver.exe 파일 존재/이름을 확인하세요.')
        log.insert(tk.END, "exception message : 'chromedriver.exe' not exist\n")
        return -1
    finally:
        driver.close()
        log.insert(tk.END, "status message : 'chromedriver.exe' is OK\n")

    
    # Captcha font 로드
    font1, font2, font3 = load_font()
    log.insert(tk.END, "status message : captcha font loaded\n")

     
    # 예약시간 2분전까지 대기
    if not test:
        now = datetime.datetime.now()
        th = now.replace(hour=h-1, minute=58, second=0, microsecond=0)
        log.insert(tk.END, "\r{} / {} : waiting   ".format(now, th))
        while(now < th):
            now = datetime.datetime.now()
            log.delete('end-1l', tk.END)
            log.insert(tk.END, "\n{} / {} : waiting    ".format(now, th))
            time.sleep(0.5)
            log.delete('end-1l', tk.END)
            log.insert(tk.END, "\n{} / {} : waiting.   ".format(now, th))
            time.sleep(0.5)
            log.delete('end-1l', tk.END)
            log.insert(tk.END, "\n{} / {} : waiting..  ".format(now, th))
            time.sleep(0.5)
            log.delete('end-1l', tk.END)
            log.insert(tk.END, "\n{} / {} : waiting... ".format(now, th))
            time.sleep(0.5)
            log.delete('end-1l', tk.END)
            log.insert(tk.END, "\n{} / {} : waiting....".format(now, th))
            time.sleep(0.5)
        log.insert(tk.END, "\nstatus message : program start\n")
    
    
    # 웹페이지 열기
    driver = webdriver.Chrome()
    actions = ActionChains(driver)
    url = "https://ananti.kr/ko/login"
    driver.get(url)
    driver.maximize_window()


    # 로그인
    time.sleep(3)
    driver.find_element(By.XPATH, "//input[contains(@type, 'text')]").send_keys(id)
    driver.find_element(By.XPATH, "//input[contains(@type, 'password')]").send_keys(pw)
    driver.find_element(By.ID, "btnLogin").click()
    
    log.insert(tk.END, "status message : login done\n")


    # 예약페이지로 이동
    time.sleep(3)
    driver.get("https://ananti.kr/ko/reservation")
    
    log.insert(tk.END, "status message : reservation page open\n")


    # # 회원권 타입 선택
    # time.sleep(1)
    # driver.execute_script("document.getElementById('bookingMemberType').style.display='block';")
    # select = Select(driver.find_element(By.ID, "bookingMemberType"))
    # select.select_by_value(id)
    # driver.execute_script("document.getElementById('bookingMemberType').style.display='none';")
    # time.sleep(1)
    # # actions.move_to_element(driver.find_element(By.ID, "bookingMemberType")).click().perform()
    # driver.find_element(By.XPATH, "/html/body/header/div/div[1]/div[3]/div").click()
    
    # log.insert(tk.END, "status message : booking type selected\n")


    # 플랫폼 타입 선택
    time.sleep(1)
    driver.execute_script("document.getElementById('bookingPlatform').style.display='block';")
    select = Select(driver.find_element(By.ID, "bookingPlatform"))
    select.select_by_value(platform_value)
    
    log.insert(tk.END, "status message : platform type selected\n")


    # 객실 타입 선택
    time.sleep(1)
    driver.execute_script("document.getElementById('bookingPackage').style.display='block';")
    select = Select(driver.find_element(By.ID, "bookingPackage"))
    select.select_by_value("RR")
    driver.execute_script("document.getElementById('bookingPackage').style.display='none';")
    time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="resultInfoDiv"]/div').click()
    
    log.insert(tk.END, "status message : room group selected\n")


    # 예약 시작하기 클릭
    time.sleep(1)
    driver.find_element(By.ID, "startRsvn").click()
    
    log.insert(tk.END, "status message : start reservatino button clicked\n")


    # 예약일에 맞는 달력으로 이동
    try:
        WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.CLASS_NAME, "passed")))
        log.insert(tk.END, "status message : calendar loaded\n")
    except:
        log.insert(tk.END, "exception message : calendar not loaded\n")
        
    n_click = 0                 # 예약일에 맞는 월까지 넘겨야하는 페이지 수
    date_button = driver.find_element(By.XPATH, "//td[contains(@data-day, '%s')]" % (checkout))
    while date_button.is_displayed() is not True:
        time.sleep(1)
        try:
            next_button = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.element_to_be_clickable((By.CLASS_NAME, "icon.icon-black.icon-chevron-right")))
            next_button.click()
            n_click = n_click + 1
        except:
            log.insert(tk.END, "exception message : no next button\n")
        time.sleep(1)  
    log.insert(tk.END, "status message : n click = " + str(n_click) + '\n')
    
    
    # 예약 시작 시간까지 대기
    if not test:  
        now = datetime.datetime.now()
        th = now.replace(hour=h-1, minute=59, second=59, microsecond=900000)
        log.insert(tk.END, "{} / {} : still waiting".format(now, th))
        while(now < th):
            now = datetime.datetime.now()
            log.delete('end-1l', tk.END)
            log.insert(tk.END, "\n{} / {} : still waiting".format(now, th))
        log.insert(tk.END, "\nstatus message : reservation start\n")
    
    
    # 서버가 열릴때(날짜 버튼이 활성화)까지 새로고침 반복
    driver.refresh()
    try:
        WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.CLASS_NAME, "passed")))
        log.insert(tk.END, "status message : page loaded\n")
    except:
        log.insert(tk.END, "exception message : page not loaded\n")
    
    button = len(driver.find_elements(By.CLASS_NAME, "date-cherry"))
    while button == 0:
        driver.refresh()
        log.insert(tk.END, "status message : page refreshed\n") #, datetime.datetime.now())
        try:
            WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.CLASS_NAME, "passed")))
            button = len(driver.find_elements(By.CLASS_NAME, "date-cherry"))
        except:
            log.insert(tk.END, "exception message : page not loaded\n")
    log.insert(tk.END, "status message : server open\n")    
    
    
    # 예약일에 맞는 월 페이지로 이동
    if n_click != 0:
        for n in range(n_click):
            try:
                next_button = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.element_to_be_clickable((By.CLASS_NAME, "icon.icon-black.icon-chevron-right")))
                next_button.click()
            except:
                log.insert(tk.END, "exception message : no next button\n")   
        log.insert(tk.END, "status message : moved to reservation month\n")    
    
    
    # 체크인/체크아웃 날짜 선택
    active_first = 0
    while active_first == 0:
        try:
            checkin_button = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.element_to_be_clickable((By.XPATH, "//td[contains(@data-day, '%s')]" % (checkin))))
            checkin_button.click()
            log.insert(tk.END, "status message : checkin date clicked\n")
        except:
            log.insert(tk.END, "exception message : no checkin date clicked\n")
        active_first = len(driver.find_elements(By.XPATH, "//td[contains(@class, 'on')]"))
    
    active_last = 0
    while active_last == 0:
        try:
            checkout_button = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.element_to_be_clickable((By.XPATH, "//td[contains(@data-day, '%s')]" % (checkout))))
            checkout_button.click()
            log.insert(tk.END, "status message : checkout date clicked\n")
        except:
            log.insert(tk.END, "exception message : no checkout date button\n")             
        active_last = len(driver.find_elements(By.XPATH, "//td[contains(@class, 'last')]"))


    # 룸타입 선택
    try:
        room_button = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '%s')]" % roomtype)))
        actions.move_to_element(room_button).click().perform()
        log.insert(tk.END, "status message : room type selected\n")
    except:
        log.insert(tk.END, "exception message : no room type button\n")
    
    
    # 예약하기 버튼 클릭
    reserve_button = driver.find_element(By.XPATH, "//*[contains(text(), '예약하기')]")
    actions.move_to_element(reserve_button).click().perform()
    log.insert(tk.END, "status message : reservation button clicked\n")


    # 숙박 인원 선택
    driver.execute_script("document.getElementById('gst').style.display='block';")
    select = Select(driver.find_element(By.ID, "gst"))
    select.select_by_index(guests)
    actions.move_to_element(driver.find_element(By.ID, "gst")).click().perform()
    log.insert(tk.END, "status message : number of guests selected\n")
    
    
    # 약관 동의
    # 이 코드는 체크박스 상태를 확인하며 클릭될 때까지 반복
    while not driver.find_element(By.XPATH, "//*[contains(@type, 'checkbox') and contains(@id, 'reservation-check')]").is_selected():
        log.insert(tk.END, 'status message : checkbox click loop\n')
        agree_button = driver.find_element(By.XPATH, "//*[contains(text(), '약관')]")
        actions.move_to_element(agree_button).click().perform()
    log.insert(tk.END, "status message : agreement clicked\n") 
    #-----------------------------------------------------------------------
    ## 이 코드는 체크박스 상태를 확인 안한 버전
    # agree_button = driver.find_element(By.XPATH, "//*[contains(text(), '약관')]")
    # actions.move_to_element(agree_button).click().perf;'orm()
    # print("status message : agreement clicked")        
    
    
    ##=========================================== captcha 입력으로 이동
    # Captcha 텍스트 입력    
    try:
        rsv_button = driver.find_element(By.ID, "reservation")
        captchaTxt = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.ID, "captchaText")))
        actions.move_to_element(rsv_button).perform()
        captchaTxt.click()
        #actions.move_to_element(captchaTxt).click().perform()
        # captchaTxt.send_keys(captcha)
        log.insert(tk.END, "status message : captcha click\n")
    except:
        log.insert(tk.END, "exception message : no captcha input\n")
    #======================================================================================
        
    # # Captcha 풀기
    # img_src = driver.find_element(By.ID, "captchaImg")
    # img_src.screenshot("captcha.png")
    # img = cv2.imread("captcha.png", cv2.IMREAD_GRAYSCALE)
    # captcha_txt = SolveCaptcha(font1, font2, font3, img)
    # log.insert(tk.END, "status message : captcha text = " + captcha_txt + '\n')
        
        
    # # Captcha 텍스트 입력    
    # try:
    #     captchaTxt = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.ID, "captchaText")))
    #     actions.move_to_element(captchaTxt).click().send_keys(captcha_txt).perform()
    #     # captchaTxt.send_keys(captcha)
    #     log.insert(tk.END, "status message : captcha input\n")
    # except:
    #     log.insert(tk.END, "exception message : no captcha input\n")
            
            
    # # Captcha 확인 버튼 클릭        
    # try:
    #     captchaTxt = WebDriverWait(driver, 5, poll_frequency=0.0001, ignored_exceptions=ignore_list).until(EC.presence_of_element_located((By.ID, "captchaBtn")))
    #     actions.move_to_element(captchaTxt).click().perform()
    #     log.insert(tk.END, "status message : captcha confirm clicked\n")
    # except:
    #     log.insert(tk.END, "exception message : no captcha confirm\n")
    
    
    # # 예약 접수 버튼 클릭 
    # confirm_box = len(driver.find_elements(By.XPATH, "//*[contains(text(), '까?')]"))
    # while confirm_box == 0:
    #     # print(confirm_box)
    #     rsv_button = driver.find_element(By.ID, "reservation")
    #     actions.move_to_element(rsv_button).click().perform()
    #     confirm_box = len(driver.find_elements(By.XPATH, "//*[contains(text(), '까?')]"))
    # log.insert(tk.END, "status message : captcha OK clicked\n")
    # log.insert(tk.END, "status message : reservation button clicked\n")
        

    # # 최종 확인 버튼 클릭
    # if not test:
    #     while confirm_box != 0:
    #         # print(confirm_box)
    #         confirm_button = driver.find_element(By.XPATH, "//button[contains(@type, 'submit') and contains(@name, 'fn_confirm')]")
    #         actions.move_to_element(confirm_button).click().perform()
    #         confirm_box = len(driver.find_elements(By.XPATH, "//*[contains(text(), '까?')]"))
    #     log.insert(tk.END, "status message : reservation done\n")

    time.sleep(3600)



if __name__ == '__main__':
    #-------------------------- 전역변수 선언 --------------------------
    global test                 # 모드 선택 : 0 = reserve mode / 1 = test mode
    test = 0
    global combo_nights
    global combo_reserve_type
    global frame_roomtype
    global selected_place
    global combo_year
    global combo_month
    global combo_day
    global combo_guests
    global h
    
    
    #-------------------------- 윈도우 생성 --------------------------
    window = tk.Tk()
    window.title('Ananti Reservation(예약 모드)')  # 윈도우 이름 설정
    win_w = 490
    win_h = 560
    scr_w = window.winfo_screenwidth()
    scr_h = window.winfo_screenheight()
    pos_x = (scr_w - win_w) / 2
    pos_y = (scr_h - win_h) / 2
    window.geometry('%dx%d+%d+%d' % (win_w, win_h, pos_x, pos_y))   # 윈도우 크기 & 초기 좌표 설정
    window.resizable(False, False)      # 윈도우 크기 변경 가능 여부 (width, height 순)
    
    style = font.Font(size=10)
    
    
    #-------------------------- 메뉴 생성 --------------------------
    main_menu = tk.Menu(window)
    window.config(menu=main_menu)
    
    option_menu = tk.Menu(main_menu)
    main_menu.add_cascade(label='Option', menu=option_menu, font=style)
    option_menu.add_command(label='예약 모드', command=select_reserve_mode, font=style)
    # option_menu.add_seperator()
    option_menu.add_command(label='테스트 모드', command=select_test_mode, font=style)
    
    
    #-------------------------- 기본정보(예약자, 장소) 선택 --------------------------
    frame_basic = tk.LabelFrame(window, text='기본 정보', font=style)
    frame_basic.pack(padx=1, pady=1)
    frame_basic.place(x=20, y=10, width=450, height=80)
    
    # 예약자 선택
    frame_user = tk.LabelFrame(frame_basic, text='예약자', font=style)
    frame_user.place(x=15, y=2, width=120, height=50)
    
    combo_user = ttk.Combobox(frame_user, width=8, font=style)
    combo_user['values'] = ['박태종', '김양미', '박기택', '박준택']
    combo_user.pack(padx=5, pady=5)
    # combo_user.current(0)
    combo_user.bind('<<ComboboxSelected>>', select_id_pw)
    
    # 예약 장소 선택
    frame_place = tk.LabelFrame(frame_basic, text='장소', font=style)
    frame_place.place(x=150, y=2, width=280, height=50)
    
    var_place = tk.IntVar()
    radio_place1 = tk.Radiobutton(frame_place, text='가평', variable=var_place, value=1, command=select_place, font=style, padx=15, pady=5).grid(row=0, column=1)
    radio_place2 = tk.Radiobutton(frame_place, text='부산', variable=var_place, value=2, command=select_place, font=style, padx=15, pady=5).grid(row=0, column=2)
    radio_place3 = tk.Radiobutton(frame_place, text='남해', variable=var_place, value=3, command=select_place, font=style, padx=15, pady=5).grid(row=0, column=3)
    
    
    #-------------------------- 예약정보 선택 --------------------------
    frame_reserve = tk.LabelFrame(window, text='예약 정보', font=style)
    frame_reserve.pack(padx=1, pady=1)
    frame_reserve.place(x=20, y=100, width=450, height=150)
    
    
    #------------- 예약 모드(4주전/8주전/12주전) 선택 -------------
    frame_reserve_type = tk.LabelFrame(frame_reserve, text='예약모드', font=style)
    frame_reserve_type.place(x=15, y=2, width=120, height=50)


    #------------- 룸타입 선택 -------------
    frame_roomtype = tk.LabelFrame(frame_reserve, text='룸타입', font=style)
    frame_roomtype.place(x=150, y=2, width=280, height=50)

    
    #------------- 체크인 날짜 선택 -------------
    frame_date = tk.LabelFrame(frame_reserve, text='체크인 날짜', font=style)
    frame_date.place(x=15, y=55, width=175, height=65)
    
    # 직접 입력 : 년
    frame_year = tk.LabelFrame(frame_date, text='년', font=style)
    frame_year.place(x=5, y=0)
    
    combo_year = ttk.Combobox(frame_year, width=4, font=style)
    combo_year['values'] = [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]
    combo_year.pack(padx=5, pady=5)
    combo_year.current(0)
    combo_year.bind('<<ComboboxSelected>>', select_date)
    
    # 직접 입력 : 월
    frame_month = tk.LabelFrame(frame_date, text='월', font=style)
    frame_month.place(x=68, y=0)
    
    combo_month = ttk.Combobox(frame_month, width=2, font=style)
    combo_month['values'] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    combo_month.pack(padx=5, pady=5)
    combo_month.current(0)
    combo_month.bind('<<ComboboxSelected>>', select_date)
    
    # 직접 입력 : 일
    frame_day = tk.LabelFrame(frame_date, text='일', font=style)
    frame_day.place(x=116, y=0)
    
    combo_day = ttk.Combobox(frame_day, width=2, font=style)
    combo_day['values'] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
                            17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
    combo_day.pack(padx=5, pady=5)
    combo_day.current(0)
    combo_day.bind('<<ComboboxSelected>>', select_date)
    
    combo_year.configure(state='disabled')
    combo_month.configure(state='disabled')
    combo_day.configure(state='disabled')

    
    #------------- 숙박일수 선택 -------------
    frame_nights = tk.LabelFrame(frame_reserve, text='숙박일수', font=style)
    frame_nights.place(x=210, y=55, width=100, height=65)

    combo_nights = ttk.Combobox(frame_nights, width=4, font=style)
    combo_nights.place(x=23, y=13)
    combo_nights['values'] = [2, 3, 4, 5, 6]
    combo_nights.current(0)
    combo_nights.bind('<<ComboboxSelected>>', select_nights)
    
    combo_nights.configure(state='disabled')

    
    #------------- 인원 선택 -------------
    frame_guests = tk.LabelFrame(frame_reserve, text='인원', font=style)
    frame_guests.place(x=330, y=55, width=100, height=65)

    combo_guests = ttk.Combobox(frame_guests, width=4, font=style)
    combo_guests.place(x=23, y=13)
    combo_guests['values'] = [1, 2, 3, 4]
    combo_guests.current(0)
    combo_guests.bind('<<ComboboxSelected>>', select_guests)
    
    combo_guests.configure(state='disabled')

    
    #-------------------------- 시작 버튼 --------------------------
    btn_confirm = tk.Button(window, text='예약 시작', command=confirm, width=20, font=style)
    btn_confirm.pack()
    btn_confirm.place(x=170, y=260)
    
    
    #-------------------------- 로그 생성 --------------------------
    frame_log = tk.LabelFrame(window, text='Log', font=style)
    frame_log.pack()
    frame_log.place(x=20, y=285)
    
    scrollbar = tk.Scrollbar(frame_log)
    scrollbar.pack(side='right', fill='y')
    
    log = tk.Text(frame_log, width=61, height=18, yscrollcommand=scrollbar.set, font=style)
    log.pack(side='left', fill='both')
    
    log.insert(tk.END, '아난티 예약 프로그램입니다.\n')
    
    
    #-------------------------- 프로그램 실행 --------------------------
    window.mainloop()