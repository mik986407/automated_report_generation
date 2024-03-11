# -*- coding: UTF-8 –*-
from pathlib import Path
import numpy as np
import pandas as pd
import os, string, re, zipfile
from datetime import datetime , timedelta ,date
from dateutil.relativedelta import relativedelta
import warnings
warnings.filterwarnings('ignore')

#===================================================設置路徑及時間=================================================================
#時間設定
# 設定製作報表時間
# today = date.today()
# day_now = today.strftime('%Y%m%d')
day_now = '20230306'
today = datetime.strptime(day_now,"%Y%m%d")

# 日期計算基準
time_now = datetime.strptime('2023/01/01','%Y/%m/%d') # 須詢問
#time_now = datetime.now()

# 設定下載壓縮檔的目錄與解壓縮檔案的放置資料夾
download_folder = r'C:\Users\mik986407\py_data\hiv資料\壓縮檔專區' # 壓縮路徑

today = date.today()
day_stamp = today.strftime('%Y%m%d')
folder_path =  r"C:\Users\mik986407\py_data\hiv資料\解壓縮檔案區_{}".format(day_stamp) # 解壓縮路徑

# 路徑設定
#灰色部分
sub_path1 =folder_path + r'\個案基本資料.csv'
#綠色部分
sub_path2 = folder_path + r'\急性初期感染個案管理月報表.csv'
sub_path3 = folder_path + r'\愛滋感染者新通報疾病個案管理季報表.csv'
sub_path4 = folder_path + r'\合併使用成癮性藥物個案管理季報表.csv'
sub_path5 = folder_path + r'\個案管理狀態查詢結果.csv'
sub_path6 = folder_path + r'\特殊個案.csv'
#橘色部分
sub_path7 = folder_path + r'\存活個案就醫率、服藥率及病毒量測不到率(健保資料+系統就醫資料).csv'
sub_path8 = folder_path + r'\就醫紀錄.csv'
#粉色部分
sub_path9 = folder_path + r'\在台開始服藥日期.csv'
#紫色部分
sub_path10 = folder_path + r'\非本國籍感染者統計表.csv'
#黃色部分
sub_path11 = folder_path + r'\藥酒癮系統個案戒癮資料.csv'
sub_path12 = folder_path + r'\替代治療情形-給藥紀錄.csv'

#===========================================================建立資料夾===============================================================

# 確保資料夾存在，如果不存在，則建立資料夾
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# # 清空資料夾
# for filename in os.listdir(folder_path):
#     file_path = os.path.join(folder_path, filename)
#     try:
#         if os.path.isfile(file_path) or os.path.islink(file_path):
#             os.unlink(file_path)
#         elif os.path.isdir(file_path):
#             shutil.rmtree(file_path)
#     except Exception as e:
#         print('無法刪除 %s. 原因: %s' % (file_path, e))
#=========================================================壓縮檔解壓================================================================

def output_zipfile(zip_f,check_f):
    zip_fname = os.path.join(download_folder, zip_f)
    check_fname = os.path.join(folder_path, check_f)
    with zipfile.ZipFile(zip_fname) as fn:  # 檔案的路徑與檔案名
        if os.path.exists(check_fname) is not True:
            zip_list = fn.namelist()  # 得到壓縮包里所有檔案
            for f in zip_list:
                true_path = Path(fn.extract(f,pwd="L125370530".encode("utf-8")))  # 回圈解壓檔案到指定目錄
                true_path.rename(f.encode("cp437").decode("cp950"))
                file_name = str(true_path).split('\\')[-1].split('.')[0].encode("cp437").decode("cp950")+'.csv'
                print('檔案: {} 解壓縮成功!'.format(file_name))
        else:
            print('壓縮檔: {} 已解壓縮!'.format(zip_f))
            
os.chdir(folder_path)
zip_list = ['個案資料.zip','急性初期感染個案管理月報表.zip','愛滋感染者新通報疾病個案管理季報表.zip','合併使用成癮性藥物個案管理季報表.zip'
            ,'個案管理狀態查詢結果.zip','特殊個案.zip','存活個案就醫率、服藥率及病毒量測不到率(健保資料+系統就醫資料).zip','個案資料 (1).zip'
            ,'非本國籍感染者統計表.zip','個案資料 (2).zip']

check_list = ['個案基本資料.csv','急性初期感染個案管理月報表.csv','愛滋感染者新通報疾病個案管理季報表.csv'
              ,'合併使用成癮性藥物個案管理季報表.csv','個案管理狀態查詢結果.csv','特殊個案.csv'
              ,'存活個案就醫率、服藥率及病毒量測不到率(健保資料+系統就醫資料).csv','就醫紀錄.csv'
              ,'非本國籍感染者統計表.csv','替代治療情形-給藥紀錄.csv']

file_dict = zip(zip_list,check_list)

for zip_f,check_f in file_dict:
    output_zipfile(zip_f,check_f)
    print('-----'*20)
    
#=========================================================灰色資料部分=====================================================================
print('=====開始處理資料=====')
# 個案基本資料
#time_now = datetime.now()
def notify_process(x):
    time1 = datetime.strptime(x,'%Y/%m/%d')
    time_diff = time_now - time1
    if time_diff >= timedelta(days = 730):
        return('滿2年')
    elif (time_diff >= timedelta(days = 183))&(time_diff < timedelta(days = 730)):
        return('半年至2年')
    else:
        return('半年內')

month_date_start = datetime.now().date() - relativedelta(months=1) # 需考量
month_date_start = datetime.strptime(month_date_start.strftime('%Y/%m')+'/01','%Y/%m/%d')
month_date_end = datetime.strptime(datetime.now().strftime('%Y/%m')+'/01','%Y/%m/%d')    
def new_notify(x):
    time1 = datetime.strptime(x,'%Y/%m/%d')
    if (time1 >= month_date_start)&(time1 < month_date_end):
        return('*')
    else:
        return(None)

df1 = pd.read_csv(sub_path1)
use_cols = ['HIV編號','居留證號','生日','分局別','管理縣市','管理鄉鎮','電腦編號','通報時國籍別','通報時國籍名稱','性Sex','狀態','HIV確認通報定義',
           'HIV通報醫療院所','HIV診斷日期','HIV診斷年齡','HIV足歲診斷年齡','實際年齡','通報縣市','HIV通報日期','AIDS通報日期','婚姻狀況','HIV感染危險因子',
            'HIV感染危險因子_性行為_性行為對象','HIV感染危險因子_性行為_性行為對象是否為毒癮者','HIV感染危險因子_靜脈注射藥癮者_性行為對象','HIV感染危險因子_其他(說明)',
           '目前狀況','服務卡_卡別']
df1_new = df1[use_cols].sort_values('HIV編號')
df1_new['通報時程'] = df1_new['HIV通報日期'].apply(notify_process)
df1_new['新通報'] = df1_new['HIV通報日期'].apply(new_notify)
df1_new['目前未成年'] = df1_new['實際年齡'].apply(lambda x: '*' if int(x) < 18 else None)
print('--個案基本資料處理完成--')
print('====灰色資料處理完成====')
#=========================================================綠色資料部分===================================================================

# 急性初期感染個案管理月報表
df2 = pd.read_csv(sub_path2)
df2_new = df2[['HIV編號','新通報個案']]
df2_new.rename(columns={'新通報個案': '本年度急性初期感染'}, inplace=True)
df2_new = df2_new[df2_new['本年度急性初期感染'] == '*']
print('--本年度急性初期感染處理完成--')

#-------------------------------------------------------------------------------------------------------------------------

# 新通報STD -> 愛滋感染者新通報疾病個案管理季報表
df3 = pd.read_csv(sub_path3)
df3_tmp_1 = df3[df3['共病數'] == 1]
df3_tmp_2 = df3[df3['共病數'] == 2]
df3_tmp_3 = df3[df3['共病數'] >= 3]
#去除區間內通報梅毒且為非活性梅毒,去除通報結核病
df3_tmp_1 = df3_tmp_1[~(df3_tmp_1['梅毒分型']=='非活性梅毒')]
df3_tmp_1 = df3_tmp_1[~(df3_tmp_1['區間內通報結核病']=='*')]
#去除區間內通報含非活性梅毒與結核病
df3_tmp_2 = df3_tmp_2[~((df3_tmp_2['梅毒分型']=='非活性梅毒')&(df3_tmp_2['區間內通報結核病']=='*'))]
#合併資料
df3_new = pd.concat([df3_tmp_1,df3_tmp_2,df3_tmp_3])
df3_new['新通報STD'] = '*'
df3_new = df3_new[['HIV編號','新通報STD']]
print('--新通報STD處理完成--')
#------------------------------------------------------------------------------------------------------------------------

# 近1年合併使用成癮物質 ->合併使用成癮性藥物個案管理季報表
df4 = pd.read_csv(sub_path4)
df4_new = df4[['HIV編號','新通報個案']]
df4_new.dropna(subset=['新通報個案'],inplace=True)
df4_new.rename(columns={'新通報個案': '近1年合併使用成癮物質'}, inplace=True)
df4_new.drop_duplicates(subset=['HIV編號'], keep='last',inplace=True)
print('--近1年合併使用成癮物質處理完成--')

#------------------------------------------------------------------------------------------------------------------------

def define_prison(x):
    x = str(x)
    if '入監或移監' in x :
        return('*')
    else:
        return(np.nan)

# 估計在監各案 -> 個案管理狀態查詢結果(缺)
df5 = pd.read_csv(sub_path5,encoding='ansi',usecols=['HIV編號','個案狀況'])
df5['估計在監各案'] = df5['個案狀況'].apply(define_prison)
df5.dropna(subset=['估計在監各案'],inplace=True)
df5_new = df5[['HIV編號','估計在監各案']]
df5_new.sort_values(by = ['HIV編號'],inplace=True)
print('--估計在監各案處理完成--')

#------------------------------------------------------------------------------------------------------------------------

# TB -> 特殊個案 
df6 = pd.read_csv(sub_path6)
df6_tmp = df6[['電腦編號','結核病確診狀態']]
df6_new = df6_tmp[df6_tmp['結核病確診狀態'] == '確診管理中']
df6_new['TB'] ='*'
df6_new = df6_new[['電腦編號','TB']]
print('--TB處理完成--')
print('====綠色資料處理完成====')

#=========================================================橘色資料部分===================================================================

# 報表19.2
df7 = pd.read_csv(sub_path7)
df7_new = df7[['HIV編號','就醫個案','服藥個案']]
df7_new.rename(columns={'就醫個案': '有就醫','服藥個案':'有服藥'}, inplace=True)
# df7_new.dropna(subset=['有就醫', '有服藥'],inplace=True)
print('--報表19.2處理(含:有就醫、有服藥)完成--')

#------------------------------------------------------------------------------------------------------------------------

#就醫資料:病毒量控制取最新日期，且值大等於200
def define_virus_value(x):
    if x == 'undetectable':x = '1'
    if float(re.findall(r'-?\d+\.?\d*', x)[0]) < 200:return('*')
    else:return('測得到')

usecols = ['HIV編號', '病毒量序號', '病毒量']
df8 = pd.read_csv(sub_path8,usecols = usecols )
df8.dropna(subset=['病毒量'],inplace = True)
df8_new = df8.sort_values(by=['HIV編號', '病毒量序號'])
df8_new = df8_new.drop_duplicates(subset=['HIV編號'], keep='last')  
df8_new['病毒量控制(就醫資料)'] = df8_new['病毒量'].apply(define_virus_value)
df8_new = df8_new[['HIV編號', '病毒量控制(就醫資料)']]
print('--病毒量控制(就醫資料)處理完成--')
print('====橘色資料處理完成====')

#=========================================================粉色資料部分===================================================================

#time_now = datetime.now()
def md_over_twe_year(x):
    try:
        time1 = datetime.strptime(x,'%Y/%m/%d')
    except:
        time1 = time_now
    time_diff = time_now - time1
    if time_diff >= timedelta(days = 730):
        return('*')
    else:
        return(np.nan)
    
# 在台開始服藥日
df9_new = pd.read_csv(sub_path9,usecols=['HIV編號','在台開始服藥日期'])
df9_new.sort_values(by=['HIV編號'],inplace=True)
# df10['在台開始服藥日期'] = df10['在台開始服藥日期'].apply(lambda x: '未服藥')
df9_new['服藥年分'] = df9_new['在台開始服藥日期'].apply(md_over_twe_year)
print('--在台開始服藥日處理完成--')
print('====粉色資料處理完成====')

#=========================================================紫色資料部分===================================================================

#非本國籍感染者統計表
df10 = pd.read_csv(sub_path10)
df10_tmp = df10[['HIV編號','身分別','是否已歸化為本國籍','是否為目前境內外籍愛滋就醫個案']]
df10_new = df10_tmp.fillna(0)
df10_new['職業別'] = np.nan
df10_new = df10_new[['HIV編號','身分別','職業別','是否已歸化為本國籍','是否為目前境內外籍愛滋就醫個案']]
print('--非本國籍感染者處理完成--')
print('====紫色資料處理完成====')
#=========================================================黃色資料部分===================================================================
# 藥酒癮系統個案戒癮資料 -> DETOX 
df11 = pd.read_csv(sub_path11)
df11_tmp = df11.drop_duplicates(subset=['HIV編號'],keep = 'last')
df11_new = df11_tmp[['HIV編號']]
df11_new['DETOX'] = '*'
print('--DETOX處理完成--')
# 替代治療情形-給藥紀錄 -> MMT 
df12 = pd.read_csv(sub_path12)
df12_tmp = df12[df12['出席狀況']=='出席']
df12_tmp['給藥日期'] = df12_tmp['給藥日期'].apply(lambda x: datetime.strptime(x,'%Y/%m/%d'))
df12_tmp.sort_values(by = ['HIV編號','給藥日期'],inplace=True)
df12_tmp = df12_tmp.drop_duplicates(subset=['HIV編號'], keep='last') 
df12_new = df12_tmp[['HIV編號']]
df12_new['MMT'] = '*'
print('--MMT處理完成--')
print('====黃色資料處理完成====')
#==========================================================資料勾稽======================================================================

# 綠色部分
# 本年度急性初期感染
df_all = df1_new.merge(df2_new,on = 'HIV編號',how = 'left')
# 新通報STD
df_all = df_all.merge(df3_new,on = 'HIV編號',how = 'left')
# 近1年合併使用成癮物質
df_all = df_all.merge(df4_new,on = 'HIV編號',how = 'left')
# 估計在監各案  
df_all = df_all.merge(df5_new,on = 'HIV編號',how = 'left')
# TB
df_all = df_all.merge(df6_new,on = '電腦編號',how = 'left')
#橘色部分
#有就醫、有服藥
df_all = df_all.merge(df7_new,on = 'HIV編號',how = 'left')
#就醫資料:病毒量控制
df_all = df_all.merge(df8_new,on = 'HIV編號',how = 'left')
# 粉色部分
# 在台開始服藥日期
df_all = df_all.merge(df9_new,on = 'HIV編號',how = 'left')
#紫色部分
#非本國籍
df_all = df_all.merge(df10_new,on = 'HIV編號',how = 'left')
# 黃色部分
#DETOX
df_all = df_all.merge(df11_new,on = 'HIV編號',how = 'left')
#MMT
df_all = df_all.merge(df12_new,on = 'HIV編號',how = 'left')
print('====資料勾稽處理完成====')

#=====================================================對特定欄位的缺值補0及新增欄位==========================================================

fillna_cols = ['本年度急性初期感染','新通報STD','近1年合併使用成癮物質','估計在監各案','TB','有就醫','有服藥','病毒量控制(就醫資料)',
               '服藥年分','身分別','是否已歸化為本國籍','是否為目前境內外籍愛滋就醫個案']

for col in fillna_cols:
    df_all[col] = df_all[col].apply(lambda x: 0 if x is np.nan else x)

df_all['在台開始服藥日期'] = df_all['在台開始服藥日期'].apply(lambda x: '未服藥' if x is np.nan else x)

print('====對特定欄位的缺值補0完成====')

#新增欄位
#有戒癮紀錄
def f1(x):
    if (x['DETOX'] == '*') or (x['MMT'] == '*'):
        return ('*')
    else:
        return(np.nan)

df_all['有戒癮紀錄'] = df_all.apply(f1,axis=1)

#有戒癮+訪視紀錄
def f2(x):
    if (x['近1年合併使用成癮物質'] == '*') and (x['有戒癮紀錄'] == '*'):
        return ('*')
    else:
        return(np.nan)

df_all['有戒癮+訪視紀錄'] = df_all.apply(f2,axis=1)
print('====已新增有戒癮紀錄以及有戒癮+訪視紀錄====')
#========================================================資料匯出======================================================================

try:
    # filename = r'D:\Users\mik986407\py_data\個案基本資料總表.xlsx'
    filename = r'C:\Users\mik986407\py_data\hiv資料\個案基本資料總表_csv_{}.csv'.format(day_stamp)
    # df_all.to_excel(filename,sheet_name='個案基本資料',index=False)
    df_all.to_csv(filename,index=False,encoding='cp950')
    print('資料匯出成功!')
except:
    print('資料匯出異常!')


























