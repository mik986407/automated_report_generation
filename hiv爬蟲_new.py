from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
import os , time
from datetime import datetime , date , timedelta
import calendar
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from dateutil.relativedelta import relativedelta

#=======================================================設定路徑及時間==================================================================
# 路徑設定
# 目前selenium版本為 4.8.1 , chrome版本為 110.0.5481.178
# chromedriver 路徑
chromedriver_file = r"C:\Users\mik986407\python_code\chromedriver.exe"
# 下載檔存放路徑,需與chrome的預設下載路徑配合
download_dir = r"C:\Users\mik986407\py_data\hiv資料\壓縮檔專區"
# 下載時的頁面截圖，用以檢查
check_folder = r'C:\Users\mik986407\py_data\hiv資料\hiv截圖'

#時間設定
# 設定製作報表時間
# today = date.today()
# day_now = today.strftime('%Y%m%d')
day_now = '20230306'
today = datetime.strptime(day_now,"%Y%m%d")

# 获取上个月的最後一天日期-----------------------------------------
last_month = today.replace(day=1) - timedelta(days=1)
# 获取上个月的第一天和最后一天
first_day = last_month.replace(day=1).strftime('%Y%m%d')
last_day = last_month.replace(day=calendar.monthrange(last_month.year, last_month.month)[1]).strftime('%Y%m%d')

# 计算前三个月的第一天-------------------------------------
first_day_of_third_month_ago = (today.replace(day=1) - timedelta(days=3*30)).replace(day=1)
# 获取前三个月的月份和年份
month = first_day_of_third_month_ago.month
year = first_day_of_third_month_ago.year
# 获取前三个月的天数
num_days = calendar.monthrange(year, month)[1]
# 結果
day_of_third_month_ago_first = first_day_of_third_month_ago.strftime("%Y%m%d")
day_of_third_month_ago_last = (first_day_of_third_month_ago.replace(day=num_days)).strftime("%Y%m%d")
# 计算前四个月的第一天-------------------------------------
first_day_of_fourth_month_ago = (today.replace(day=1) - timedelta(days=30*4)).replace(day=1)
# 获取前四个月的月份和年份
month = first_day_of_fourth_month_ago.month
year = first_day_of_fourth_month_ago.year
# 获取前四个月的天数
num_days = calendar.monthrange(year, month)[1]
# 結果
day_of_fourth_month_ago_first = first_day_of_fourth_month_ago.strftime("%Y%m%d")
day_of_fourth_month_ago_last = (first_day_of_fourth_month_ago.replace(day=num_days)).strftime("%Y%m%d")
# 计算前六个月的日期-------------------------------------
# 计算前6个月的日期
six_months_ago = (today - relativedelta(months=6)).strftime('%Y%m%d')

# 取得去年的年份-----------------------------------------
last_year = datetime.now().year - 1
# 計算去年的第一天
first_day_last_year = datetime(last_year, 1, 1).strftime("%Y%m%d")
# 計算去年的最後一天
last_day_last_year = datetime(last_year, 12, 31).strftime("%Y%m%d")

# 计算前1年的第一天-------------------------------------
first_day_of_1_years_ago = (today.replace(day=1) - timedelta(days=30*12)).replace(day=1)
# 获取前1年的月份和年份
month = first_day_of_1_years_ago.month
year = first_day_of_1_years_ago.year
# 获取前1年第一個月的天数
num_days = calendar.monthrange(year, month)[1]
# 結果
day_of_1_year_ago_first = first_day_of_1_years_ago.strftime("%Y%m%d")
day_of_1_year_ago_last = (first_day_of_1_years_ago.replace(day=num_days)).strftime("%Y%m%d")
#=============================================================刪除舊檔================================================================
# 確保資料夾存在，如果不存在，則建立資料夾
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

if not os.path.exists(check_folder):
    os.makedirs(check_folder)    
    
# 清空資料夾
for filename in os.listdir(download_dir):
    file_path = os.path.join(download_dir, filename)
    try:
        if os.path.isfile(file_path) or os.path.islink(file_path):
            os.unlink(file_path)
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path)
    except Exception as e:
        print('Failed to delete %s. Reason: %s' % (file_path, e))
        
check_fname_num = len(os.listdir(download_dir))
print('開始時間:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))


options = webdriver.ChromeOptions()
options.add_argument('ignore-certificate-errors')
prefs = {"download.default_directory": download_dir}
options.add_experimental_option("prefs", prefs)
s = Service(chromedriver_file)
driver = webdriver.Chrome(service=s,options=options)
driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/Security/Public/LoginForm.aspx?ReturnUrl=%2fHasweb")
driver.maximize_window()
# 點選自然人憑證按鈕
btnOK = driver.find_element(By.ID, "ctl00_main_HALoginCtrl_LoginCardRBL_1")
btnOK.click()
# 點擊確認
btnOK1 = driver.find_element(By.ID, "ctl00_main_HALoginCtrl_LoginButton")
btnOK1.click()
#----
driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/ExportSystem/HASExportOperationForm.aspx?EditPositionType=HIV")
# 在日期欄位中輸入日期
date_field = driver.find_element(By.ID, 'ctl00_main_HASExportCaseDataCtrl_t_EndDiagnosisDate_Editor')
date_field.send_keys(Keys.CONTROL, "a")
date_field.send_keys(Keys.BACKSPACE)
date_field.send_keys(last_day)

select1 = Select(driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_ExportType"))
select1.select_by_value('2')

time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_TabPanelCBL_0").click()
time.sleep(1)
driver.save_screenshot(check_folder + r'\個案基本資料.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_CsvBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_CsvBtn").click()
time.sleep(1)

try:
    driver.switch_to.alert.accept()
except:
    print('')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('個案基本資料下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#--------------------------------------------------急性初期感染個案管理月報表--------------------------------------------------------------
driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/Reports/HASWrningReport1_23Page.aspx")
# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_23Ctrl_Conditions_t_BeginDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(first_day)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_23Ctrl_Conditions_t_EndDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(last_day)

# ctl00_main_HASWrningReport1_23Ctrl_Conditions_t_Nationality
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASWrningReport1_23Ctrl_Conditions_t_Nationality"))
select1.select_by_visible_text('全部')

# ctl00_main_HASWrningReport1_23Ctrl_DownLoadCSVBtn
driver.save_screenshot(check_folder + r'\急性初期感染個案管理月報表.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_23Ctrl_DownLoadCSVBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASWrningReport1_23Ctrl_DownLoadCSVBtn").click()
time.sleep(1)

try:
    driver.switch_to.alert.accept()
except:
    print('可直接下載')
    
# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('急性初期感染個案管理月報表下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#-----------------------------------------------愛滋感染者新通報疾病個案管理季報表--------------------------------------------------------

# HIV診斷日期起迄日
date1_start = '19800101'

# 指定疾病通報日期起迄日
date2_start = '19800101'

# 計算就診率的就診日期區間
date3_start = '20220201'

# 「就醫資料的病毒量檢驗日期」截止日
date4 = day_now


driver.get(r"https://hcdweb.cdc.gov.tw/Hasweb/Reports/HASWrningReport1_25Page.aspx")
# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_t_DiagnosisDateBegin_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_t_DiagnosisDateEnd_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(last_day)

# 指定疾病通報日期開始
date_field3 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_Conditions_t_BeginDate_Editor")
date_field3.send_keys(Keys.CONTROL, "a")
date_field3.send_keys(Keys.BACKSPACE)
date_field3.send_keys(date2_start)

# 指定疾病通報日期結束
date_field4 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_Conditions_t_EndDate_Editor")
date_field4.send_keys(Keys.CONTROL, "a")
date_field4.send_keys(Keys.BACKSPACE)
date_field4.send_keys(last_day)

# 國籍別
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_Conditions_t_Nationality"))
select1.select_by_visible_text('全部')

# 計算就診率的就診日期開始
date_field5 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_t_MedicalDateBegin_Editor")
date_field5.send_keys(Keys.CONTROL, "a")
date_field5.send_keys(Keys.BACKSPACE)
date_field5.send_keys(date3_start)

# 計算就診率的就診日期結束
date_field6 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_t_MedicalDateEnd_Editor")
date_field6.send_keys(Keys.CONTROL, "a")
date_field6.send_keys(Keys.BACKSPACE)
date_field6.send_keys(last_day)

# 「就醫資料的病毒量檢驗日期」截止日
# ctl00_main_HASWrningReport1_25Ctrl_t_VirusCheckDateEnd_Editor
date_field7 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_t_VirusCheckDateEnd_Editor")
date_field7.send_keys(Keys.CONTROL, "a")
date_field7.send_keys(Keys.BACKSPACE)
date_field7.send_keys(date4)

time.sleep(1)
driver.save_screenshot(check_folder + r'\愛滋感染者新通報疾病個案管理季報表.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_DownLoadTotalCSVBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASWrningReport1_25Ctrl_DownLoadTotalCSVBtn").click()
time.sleep(1)
try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('愛滋感染者新通報疾病個案管理季報表下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#----------------------------------------------------合併使用成癮性藥物個案管理季報表-------------------------------------------------------

# HIV診斷日期起迄日
date1_start = '19840101'
date1_end = last_day

# 個案使用成癮性藥物判斷日期區間

date2_start = day_of_1_year_ago_first
date2_end = last_day

# 計算就診率的就診日期區間
date3_start = first_day_last_year
date3_end = last_day_last_year

# 「就醫資料的病毒量檢驗日期」截止日
date4 = day_now


driver.get(r"https://hcdweb.cdc.gov.tw/Hasweb/Reports/HASWrningReport1_24Page.aspx")
# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_Conditions_t_BeginDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_Conditions_t_EndDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(date1_end)

# 指定疾病通報日期開始
date_field3 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_t_DrugDateBegin_Editor")
date_field3.send_keys(Keys.CONTROL, "a")
date_field3.send_keys(Keys.BACKSPACE)
date_field3.send_keys(date2_start)

# 指定疾病通報日期結束
date_field4 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_t_DrugDateEnd_Editor")
date_field4.send_keys(Keys.CONTROL, "a")
date_field4.send_keys(Keys.BACKSPACE)
date_field4.send_keys(date2_end)

# 國籍別
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_Conditions_t_Nationality"))
select1.select_by_visible_text('全部')

# 計算就診率的就診日期開始
date_field5 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_t_MedicalDateBegin_Editor")
date_field5.send_keys(Keys.CONTROL, "a")
date_field5.send_keys(Keys.BACKSPACE)
date_field5.send_keys(date3_start)

# 計算就診率的就診日期結束
date_field6 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_t_MedicalDateEnd_Editor")
date_field6.send_keys(Keys.CONTROL, "a")
date_field6.send_keys(Keys.BACKSPACE)
date_field6.send_keys(date3_end)

# 「就醫資料的病毒量檢驗日期」截止日
# ctl00_main_HASWrningReport1_25Ctrl_t_VirusCheckDateEnd_Editor
date_field7 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_t_VirusCheckDateEnd_Editor")
date_field7.send_keys(Keys.CONTROL, "a")
date_field7.send_keys(Keys.BACKSPACE)
date_field7.send_keys(date4)

driver.save_screenshot(check_folder + r'\合併使用成癮性藥物個案管理季報表.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_DownLoadCSVBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASWrningReport1_24Ctrl_DownLoadCSVBtn").click()
time.sleep(1)


try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('合併使用成癮性藥物個案管理季報表下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#-------------------------------------------------------個案管理狀態查詢結果-------------------------------------------------------------

driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/TrackSystem/HCaseManegedStateListPage.aspx?AutoQuery=1&County=00000000-0000-0000-0000-000000000000&Action=35&CanNotSearch=1")


# ctl00_main_HASHCaseManegedStateListCtrl_CsvBtn
driver.save_screenshot(check_folder + r'\個案管理狀態查詢結果.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASHCaseManegedStateListCtrl_CsvBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASHCaseManegedStateListCtrl_CsvBtn").click()
time.sleep(1)


try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')
    
# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('個案管理狀態查詢結果下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#-------------------------------------------------------------特殊個案--------------------------------------------------------------------

# HIV診斷日期起迄日
date1_start = ''
date1_end = last_day


driver.get(r"https://hcdweb.cdc.gov.tw/Hasweb/Reports/HASWrningReport3_1Page.aspx")

# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASWrningReport3_1Ctrl_Conditions_t_BeginDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASWrningReport3_1Ctrl_Conditions_t_EndDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(date1_end)

# 國籍別
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASWrningReport3_1Ctrl_Conditions_t_Nationality"))
select1.select_by_visible_text('全部')

# ctl00_main_HASWrningReport3_1Ctrl_Conditions_t_LawDiseaseKind
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASWrningReport3_1Ctrl_Conditions_t_LawDiseaseKind"))
select1.select_by_visible_text('結核病')

driver.save_screenshot(check_folder + r'\特殊個案.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASWrningReport3_1Ctrl_DownLoadCSVBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASWrningReport3_1Ctrl_DownLoadCSVBtn").click()
time.sleep(1)


try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('特殊個案下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#------------------------------------------存活個案就醫率、服藥率及病毒量測不到率(健保資料+系統就醫資料)-----------------------------------    
# HIV診斷日期起迄日
date1_start = ''
date1_end = last_day

# 個案使用成癮性藥物判斷日期區間

date2_start = day_of_fourth_month_ago_first
date2_end = last_day

# 計算就診率的就診日期區間
date3_start = six_months_ago
date3_end = day_now

driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/Reports/HASWrningReport1_19_2Page.aspx")

# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_Conditions_t_BeginDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_Conditions_t_EndDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(date1_end)

# ctl00_main_HASWrningReport1_19Ctrl_Conditions_t_Nationality
# 國籍別
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_Conditions_t_Nationality"))
select1.select_by_visible_text('全部')

# 計算就診率的就診日期區間開始
date_field3 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_t_MedicalDateBegin_Editor")
date_field3.send_keys(Keys.CONTROL, "a")
date_field3.send_keys(Keys.BACKSPACE)
date_field3.send_keys(date2_start)

# 計算就診率的就診日期區間結束
date_field4 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_t_MedicalDateEnd_Editor")
date_field4.send_keys(Keys.CONTROL, "a")
date_field4.send_keys(Keys.BACKSPACE)
date_field4.send_keys(date2_end)

# 病毒量勾稽區間日期開始
date_field5 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_t_VirusCheckDateStart_Editor")
date_field5.send_keys(Keys.CONTROL, "a")
date_field5.send_keys(Keys.BACKSPACE)
date_field5.send_keys(date3_start)

# 病毒量勾稽區間日期結束
date_field6 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_t_VirusCheckDateEnd_Editor")
date_field6.send_keys(Keys.CONTROL, "a")
date_field6.send_keys(Keys.BACKSPACE)
date_field6.send_keys(date3_end)

# ctl00_main_HASWrningReport1_19Ctrl_DownLoadCSVBtn
driver.save_screenshot(check_folder + r'\存活個案就醫率、服藥率及病毒量測不到率.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_DownLoadCSVBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASWrningReport1_19Ctrl_DownLoadCSVBtn").click()
time.sleep(1)
try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('存活個案就醫率、服藥率及病毒量測不到率下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#----------------------------------------------------------個案資料 (1) 就醫紀錄------------------------------------------------------  
# 診斷日期起迄日
date1_start = day_of_1_year_ago_first
date1_end = last_day

# 通報日期區間

date2_start = ''
date2_end = last_day

driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/ExportSystem/HASExportOperationForm.aspx?EditPositionType=HIV")


# 下載類別
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_ExportType"))
select1.select_by_value("4")

time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_TabPanelRBL_5").click()
time.sleep(2)

# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_BeginMedicalDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_EndMedicalDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(date1_end)

# 通報日期開始
date_field3 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_BeginDiagnosisDate_Editor")
date_field3.send_keys(Keys.CONTROL, "a")
date_field3.send_keys(Keys.BACKSPACE)
date_field3.send_keys(date2_start)

# 通報日期結束
date_field4 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_EndDiagnosisDate_Editor")
date_field4.send_keys(Keys.CONTROL, "a")
date_field4.send_keys(Keys.BACKSPACE)
date_field4.send_keys(date2_end)

driver.save_screenshot(check_folder + r'\個案資料 (1) 就醫紀錄.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_CsvBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_CsvBtn").click()
time.sleep(1)



try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')
    
# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('個案資料 (1) 就醫紀錄下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#-----------------------------------------------------------非本國籍感染者統計表----------------------------------------------------------

# 診斷日期起迄日
date1_start = '19840101'
date1_end = last_day

# 通報日期區間

date2_start = '20220201'
date2_end = last_day

driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/Reports/HASWrningReport1_27Page.aspx")

# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_27Ctrl_Conditions_t_BeginDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_27Ctrl_Conditions_t_EndDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(date1_end)

# 通報日期開始
date_field3 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_27Ctrl_t_MedicalDateBegin_Editor")
date_field3.send_keys(Keys.CONTROL, "a")
date_field3.send_keys(Keys.BACKSPACE)
date_field3.send_keys(date2_start)

# 通報日期結束
date_field4 = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_27Ctrl_t_MedicalDateEnd_Editor")
date_field4.send_keys(Keys.CONTROL, "a")
date_field4.send_keys(Keys.BACKSPACE)
date_field4.send_keys(date2_end)

driver.save_screenshot(check_folder + r'\非本國籍感染者統計表.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASWrningReport1_27Ctrl_DownLoadCSVBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASWrningReport1_27Ctrl_DownLoadCSVBtn").click()
time.sleep(1)


try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('非本國籍感染者統計表下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
#---------------------------------------------------個案資料 (2) 替代治療情形-給藥紀錄---------------------------------------------------

# 診斷日期起迄日
date1_start = ''
date1_end = last_day

# 通報日期區間

date2_start = ''
date2_end = ''

driver.get(r"https://hcdweb.cdc.gov.tw/HASWeb/ExportSystem/HASExportOperationForm.aspx?EditPositionType=HIV")

# hiv 診斷日期開始
date_field1 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_BeginDiagnosisDate_Editor")
date_field1.send_keys(Keys.CONTROL, "a")
date_field1.send_keys(Keys.BACKSPACE)
date_field1.send_keys(date1_start)

# hiv 診斷日期結束
date_field2 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_EndDiagnosisDate_Editor")
date_field2.send_keys(Keys.CONTROL, "a")
date_field2.send_keys(Keys.BACKSPACE)
date_field2.send_keys(date1_end)

# 通報日期開始
date_field3 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_BeginAnnounceDate_Editor")
date_field3.send_keys(Keys.CONTROL, "a")
date_field3.send_keys(Keys.BACKSPACE)
date_field3.send_keys(date2_start)

# 通報日期結束
date_field4 = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_EndAnnounceDate_Editor")
date_field4.send_keys(Keys.CONTROL, "a")
date_field4.send_keys(Keys.BACKSPACE)
date_field4.send_keys(date2_end)

# 下載類別
select1 = Select(driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_ExportType"))
select1.select_by_value("2")

element = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_TabPanelCBL_6")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_t_TabPanelCBL_6").click()
time.sleep(1)
driver.save_screenshot(check_folder + r'\替代治療情形-給藥紀錄.png')
time.sleep(1)
element = driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_CsvBtn")
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(1)
driver.find_element(By.ID, "ctl00_main_HASExportCaseDataCtrl_CsvBtn").click()
time.sleep(1)


try:
    driver.switch_to.alert.accept()
    print('可直接下載')
except:
    print('可直接下載')

# 下載過渡階段,等檔案載好再爬新資料
fname_num = len(os.listdir(download_dir))
for i in range(100):
    time.sleep(3)
    fname_num = len(os.listdir(download_dir))
    if fname_num > check_fname_num :
        break
print('個案資料 (2) 替代治療情形-給藥紀錄下載成功:',datetime.strftime(datetime.today(),'%Y-%m-%d %H:%M:%S'))
check_fname_num = len(os.listdir(download_dir))
time.sleep(5)
print('下載完畢')
driver.close()










