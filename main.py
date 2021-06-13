import os
import time

from urllib.request import urlretrieve
from urllib.error import URLError, HTTPError

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# to be edited
account = ""
pwd = ""

Login_url = "https://sso.nutc.edu.tw/ePortal/Default.aspx"
MyArea_url = "https://sso.nutc.edu.tw/ePortal/myarea/MyArea.aspx"
app_form_registration_url = "https://ais.nutc.edu.tw/student/apps/app_form_registration.aspx"

t_chk = ""
t_sess = ""
foldername = "./download/"
subFileName = [".docx", ".pdf"]

# to be edited
nameList = [
    "s1111111111",
]

options = Options()
options.add_argument("--disable-notifications")

# to be edited
driver = webdriver.Chrome('chromedriver.exe', chrome_options=options)
driver.set_window_size(800,600)

def BasicBuilding():
    # 建立目錄
    if not os.path.isdir(foldername):
        os.mkdir(foldername)

def getUrlAttr(url, attrName):
    urlList = url.split("?")[1].split("&")
    for s in urlList:
        if s[0:len(attrName)+1] == "%s=" % attrName:
            return s.split("=")[1]
    return ""

def download_file(filename, file_type=0, t=0):
    if len(t_sess) == 0:
        print("[System - Error]", "sess error", "sess="+t_sess)
        return
    if len(t_chk) == 0:
        print("[System - Error]", "chk error", "chk="+t_chk)
        return 
    file_url = "https://report.nutc.edu.tw/export_doc.aspx?sess=%s&doc=ChineseTranscript&chk=%s&report_no=4&rank_type=0&total_grade=0&stu_list=%s&sign_in=1&sch_dep=1&pdf=%s" % (t_sess, t_chk, filename, file_type)

    try:
        urlretrieve(file_url, foldername + filename + subFileName[file_type])
    except HTTPError as exception:
        print("\t[System - Error]Status=%s: %s" % (exception.code, exception.reason))
        Update()
        if t<3:
            print("\tTry again: t=%s" % t)
            download_file(filename, file_type, t+1)
    else:
        print("\t%s%s is downloaded" % (filename, subFileName[file_type]))

def download_list():
    for name in nameList:
        print("Download", name.split('s')[1])
        download_file(name.split('s')[1])

def Login():
    driver.get(Login_url)

    driver.find_element_by_xpath('/html/body/div[1]/div[2]/form/div[3]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/input').send_keys(account)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/form/div[3]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/input').send_keys(pwd)
    os.system("pause") ## 輸入驗證碼
    time.sleep(1)
    print("[System - Info]", "is Login")
    

def Update():
    driver.get("https://sso.nutc.edu.tw/ePortal/myarea/MyArea.aspx")
    print("[System - Info]", "進入 我的專區")
    driver.find_element_by_xpath('/html/body/div[1]/form/div[3]/div[2]/div/div[3]/div/table[1]/tbody/tr/td/ul/li[9]/a').click()
    print("[System - Info]", "進入 學生管理系統")
    time.sleep(2)
    driver.switch_to_window(driver.window_handles[0]) 
    # driver.find_element_by_xpath('/html/body/div[2]/div[3]/nav/div[2]/ul/li[9]/a').click()
    driver.get("https://ais.nutc.edu.tw/student/apps/app_form_registration.aspx")
    print("[System - Info]", "進入 線上文件申請書選項")
    driver.find_element_by_xpath('/html/body/div[2]/section/form/div/ul/li[2]/a').click()
    print("[System - Info]", "進入 申請紀錄")
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/section/form/div/div/div[2]/table/tbody/tr[6]/td[4]/input').click()
    print("[System - Info]", "點擊 下載檔案")
    time.sleep(1)

    ## 取得Download url
    data = driver.find_element_by_xpath('/html/body/iframe')
    src = data.get_attribute("src")
    updateInfo(src) ## 更新sess, chk

def updateInfo(src):
    print("[Update Info]")
    global t_sess, t_chk
    t_sess = getUrlAttr(src, "sess")
    t_chk = getUrlAttr(src, "chk")
    print("\tsess=%s" % t_sess)
    print("\tchk=%s" % t_chk)

if __name__ == "__main__":
    BasicBuilding()
    Login()
    Update()
    download_list()
    