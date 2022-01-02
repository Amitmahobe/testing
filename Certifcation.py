#import
import time
import json
import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from datetime import date

from datetime import date
current_date = date.today() 
t_day=current_date.day
mth=current_date.month
if(t_day<=10):
    month=mth-1
else:
    month=mth
print(month)

if(month<=3):
    tr=1
    td=month
elif((month>=4) and (month<=6)):
    tr=2
    td=month-3
elif((month>=7) and (month<=9)):
    tr=3
    td=month-6
elif((month>=10) and (month<=12)):
    tr=4
    td=month-9

#url_link
url=input("Enter URL:")
# chromedriver path
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
path="C:/web_driver/chromedriver.exe"
driver=webdriver.Chrome(path,chrome_options=options)
driver.get(url)

#Certification
service=driver.find_element_by_xpath('//*[@id="landing_menu"]/li[2]/a').click()
mertics=ActionChains(driver)
mer=driver.find_element_by_xpath('//*[@id="32_menubar"]/ul/li[3]/a')
mark=driver.find_element_by_xpath('//*[@id="32_menubar"]/ul/li[3]/ul/li[6]/a')
mertics.move_to_element(mer).click().move_to_element(mark).click().perform()

#Monthly Trends
market=ActionChains(driver)
market_share=driver.find_element_by_xpath('//*[@id="61_menubar"]/ul/li[4]/a')
market.move_to_element(market_share).perform()
monthly_trends=driver.find_element_by_xpath('//*[@id="61_menubar"]/ul/li[4]/ul/li[1]/a').click()

#Per-Years

filte=driver.find_element_by_xpath('//*[@id="toggle"]').click()
per_year=driver.find_element_by_xpath('//*[@id="P821_PREVIOUS_YEAR_CONTAINER"]/table/tbody/tr/td/div/div/label').click()
btn=driver.find_element_by_xpath('//*[@id="B33794100989270075"]').click()

#Heading-MT
heading_MT=[]
first='//*[@id="COL'
last='"]'
for n in range(1,14):
    final=first+str('{:02d}'.format(n))+last
    heading=driver.find_element_by_xpath(final)
    heading_MT.append(heading.text)
heading_MT.insert(0,"Metrics")

#Data_MT
first_n='//*[@id="report_PRODUCT_REPORT"]/div/div[2]/table/tbody/tr[1]/td['
last_n=']'
first_np='//*[@id="report_PRODUCT_REPORT"]/div/div[2]/table/tbody/tr[2]/td['
last_np=']'
first_nc='//*[@id="report_PRODUCT_REPORT"]/div/div[2]/table/tbody/tr[3]/td['
last_nc=']'
ftb_data_n=[]
ftb_data_np=[]
ftb_data_nc=[]
for x in range(0,4):
    filte=driver.find_element_by_xpath('//*[@id="toggle"]').click()
    driver.execute_script('document.getElementById("P821_METRIC").style.cssText = "display :block !important;"')
    select_fil=Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="P821_METRIC"]'))))
    select_fil.select_by_index(str(x))
    text_value=select_fil.first_selected_option
    flt_name=text_value.text
    btn_filter=driver.find_element_by_xpath('//*[@id="B33794100989270075"]').click()
    tb_data_n=[]
    tb_data_np=[]
    tb_data_nc=[]
    for n in range(1,14):
        final_n=first_n+str(n)+last_n
        final_np=first_np+str(n)+last_np
        final_nc=first_nc+str(n)+last_nc
        data_n=driver.find_element_by_xpath(final_n)
        data_np=driver.find_element_by_xpath(final_np)
        data_nc=driver.find_element_by_xpath(final_nc)
        tb_data_n.append(data_n.text)
        tb_data_np.append(data_np.text)
        tb_data_nc.append(data_nc.text)
    tb_data_n.insert(0,flt_name)
    tb_data_np.insert(0,flt_name)
    tb_data_nc.insert(0,flt_name)
    ftb_data_n.append(tb_data_n)
    ftb_data_np.append(tb_data_np)
    ftb_data_nc.append(tb_data_nc)
    tb_data_n=[]
    tb_data_np=[]
    tb_data_nc=[]
 

#Dataframe
MT_DICT=dict(zip(heading_MT,zip(ftb_data_n[0],ftb_data_np[0],ftb_data_nc[0],ftb_data_n[1],ftb_data_np[1],ftb_data_nc[1],ftb_data_n[2],ftb_data_np[2],ftb_data_nc[2],
                               ftb_data_n[3],ftb_data_np[3],ftb_data_nc[3])))
MT_PD=pd.DataFrame(MT_DICT)

#YOY-Summary
market=ActionChains(driver)
market_share=driver.find_element_by_xpath('//*[@id="821_menubar"]/ul/li[4]/a')
market.move_to_element(market_share).perform()
monthly_trends=driver.find_element_by_xpath('//*[@id="821_menubar"]/ul/li[4]/ul/li[2]/a').click()

#Month Picker

filte=driver.find_element_by_xpath('//*[@id="toggle"]').click()
per_year=driver.find_element_by_xpath('//*[@id="P822_MONTH_YR"]').click()
month=driver.find_element_by_xpath('//*[@id="P822_MONTH_YR"]').click()
first='//div[contains (@id , "monthpicker_")]/table/tbody/tr['
sec=']/td['
last=']'
final=first+str(tr)+sec+str(td)+last
month=driver.find_element_by_xpath(final).click()
btn=driver.find_element_by_xpath('//*[@id="B35104775343187859"]').click()

#Heading-YOY
heading_yoy=[]
heading=driver.find_element_by_xpath('//*[@id="RGN_DESC"]')
heading_yoy.append(heading.text)
heading=driver.find_element_by_xpath('//*[@id="SUB_AREA_COUNT"]')
heading_yoy.append(heading.text)
heading=driver.find_element_by_xpath('//*[@id="DLR_COUNT"]')
heading_yoy.append(heading.text)
heading=driver.find_element_by_xpath('//*[@id="PRV_VAL"]')
heading_yoy.append(heading.text)
heading=driver.find_element_by_xpath('//*[@id="CURR_VAL"]')
heading_yoy.append(heading.text)
heading=driver.find_element_by_xpath('//*[@id="CHANGE"]')
heading_yoy.append(heading.text)
heading=driver.find_element_by_xpath('//*[@id="NAT_AVG"]')
heading_yoy.append(heading.text)
heading_yoy.insert(0,"Metrics")


#Dataframe
first_p='//*[@id="alt_report_total"]/tbody/tr[1]/td['
last_p=']'
first_s='//*[@id="alt_report_total"]/tbody/tr[2]/td['
last_s=']'
first_c='//*[@id="alt_report_total"]/tbody/tr[3]/td['
last_c=']'
first_n='//*[@id="alt_report_total"]/tbody/tr[4]/td['
last_n=']'
first_t='//*[@id="alt_report_total"]/tbody/tr[5]/td['
last_t=']'
ftb_data_n=[]
ftb_data_p=[]
ftb_data_c=[]
ftb_data_s=[]
ftb_data_t=[]
for x in range(0,4):
    filte=driver.find_element_by_xpath('//*[@id="toggle"]').click()
    driver.execute_script('document.getElementById("P822_METRIC").style.cssText = "display :block !important;"')
    select_fil=Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="P822_METRIC"]'))))
    select_fil.select_by_index(str(x))
    text_value=select_fil.first_selected_option
    flt_name=text_value.text
    btn_filter=driver.find_element_by_xpath('//*[@id="B35104775343187859"]').click()
    try:
        if(driver.find_element_by_xpath('//*[@id="alt_report_total"]/tbody/tr[1]/td[1]').is_displayed()):
            pass
    except:
        continue
    tb_data_n=[]
    tb_data_p=[]
    tb_data_c=[]
    tb_data_s=[]
    tb_data_t=[]
    for n in range(1,8):
        final_n=first_n+str(n)+last_n
        final_p=first_p+str(n)+last_p
        final_c=first_c+str(n)+last_c
        final_s=first_s+str(n)+last_s
        final_t=first_t+str(n)+last_t
        data_n=driver.find_element_by_xpath(final_n)
        data_p=driver.find_element_by_xpath(final_p)
        data_c=driver.find_element_by_xpath(final_c)
        data_s=driver.find_element_by_xpath(final_s)
        data_t=driver.find_element_by_xpath(final_t)
        tb_data_n.append(data_n.text)
        tb_data_p.append(data_p.text)
        tb_data_c.append(data_c.text)
        tb_data_s.append(data_s.text)
        tb_data_t.append(data_t.text)
    tb_data_n.insert(0,flt_name)
    tb_data_p.insert(0,flt_name)
    tb_data_c.insert(0,flt_name)
    tb_data_s.insert(0,flt_name)
    tb_data_t.insert(0,flt_name)
    ftb_data_n.append(tb_data_n)
    ftb_data_p.append(tb_data_p)
    ftb_data_c.append(tb_data_c)
    ftb_data_s.append(tb_data_s)
    ftb_data_t.append(tb_data_t)
    tb_data_n=[]
    tb_data_p=[]
    tb_data_c=[]
    tb_data_s=[]
    tb_data_t=[]

#DataFrame_YOY
YOY_DICT=dict(zip(heading_MT,zip(ftb_data_p[0],ftb_data_n[0],ftb_data_s[0],ftb_data_c[0],ftb_data_t[0],ftb_data_p[1],ftb_data_n[1],ftb_data_s[1],ftb_data_c[1],ftb_data_t[1]
              ,ftb_data_p[2],ftb_data_n[2],ftb_data_s[2],ftb_data_c[2],ftb_data_t[2])))
YOY_PD=pd.DataFrame(YOY_DICT)

#csv_file

final_sheet={'Monthly Trends':MT_PD,'YOY Summary':YOY_PD}
write=pd.ExcelWriter("Certification.xlsx", engine='xlsxwriter')
for sheet_name in final_sheet.keys():
    final_sheet[sheet_name].to_excel(write, sheet_name=sheet_name, index=False)
write.save()