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

#url_link
url=''
with open("C:/Users/Admin/Desktop/MAYA-TASK/vw Scrap/json/link.json") as link:
    url=json.load(link)[0]

# chromedriver path
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
path="C:/web_driver/chromedriver.exe"
driver=webdriver.Chrome(path,chrome_options=options)
driver.get(url)


pur_header=[]
cap_header=[]
cap_header1=[]
pur_header1=[]

#part_sales
link=driver.find_element_by_xpath('//*[@id="landing_menu"]/li[1]/a').click()
first_header='//*[@id="holder_'
last_header='_section"]/figure[1]/div[1]'
for n in range(1,10):
    final_header=first_header+str(n)+last_header
    header_pur=driver.find_element_by_xpath(final_header)
    pur_header.append(header_pur.text)

for sub in pur_header:
    pur_header1.append(sub.replace('\n','_'))
print(pur_header1)
first_caption='//*[@id="holder_'
last_caption='_section"]/figure[1]/div[2]'
for n in range(1,10):
    final_caption=first_caption+str(n)+last_caption
    header_caption=driver.find_element_by_xpath(final_caption)
    cap_header.append(header_caption.text)

for sub in cap_header:
    cap_header1.append(sub.replace('\n','_'))
print(cap_header1)
dict_purchase=dict(zip(pur_header1,cap_header1))
print(dict_purchase)
part_dashboard=pd.DataFrame(dict_purchase, index=[0])

#sale by tactical_sagment
metrics=driver.find_element_by_xpath('//*[@id="53_menubar"]/ul/li[2]/a')
part_sales=driver.find_element_by_xpath('//*[@id="53_menubar"]/ul/li[2]/ul/li[1]/a')

action_part_sales=ActionChains(driver)
action_part_sales.move_to_element(metrics).move_to_element(part_sales).click().perform()

action_tactical=ActionChains(driver)

part_sales1=driver.find_element_by_xpath('//*[@id="54_menubar"]/ul/li[3]/a')
tactical_sagment=driver.find_element_by_xpath('//*[@id="54_menubar"]/ul/li[3]/ul/li[2]/a')
action_tactical.move_to_element(part_sales1).move_to_element(tactical_sagment).click().perform()

#dealer_cost_table

pageSource = driver.execute_script("return document.body.innerHTML;")
page_dealer=BeautifulSoup(pageSource,'html.parser')

#table_header
tab_head=[]
column=page_dealer.findAll(attrs={'class':'t-Report-colHead'})
for n in range(0,16):
    header=column[n].text
    tab_head.append(header)
print(tab_head)


#mechanical_column
data=page_dealer.findAll(attrs={'class':'parent1'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Mechanical';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
mechanical=sale_dollar
print(mechanical)

#Collision_Column
data=page_dealer.findAll(attrs={'class':'parent2'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Collision';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
collision=sale_dollar
print(collision)

#Chemicals_Column
data=page_dealer.findAll(attrs={'class':'parent3'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Chemicals';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
chemical=sale_dollar
print(chemical)

#Accessories_Column
data=page_dealer.findAll(attrs={'class':'parent4'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Accessories';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
accessories=sale_dollar
print(accessories)

#Tires_Column
data=page_dealer.findAll(attrs={'class':'parent5'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Tires';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
tires=sale_dollar
print(tires)

#Miscellaneous_Column
data=page_dealer.findAll(attrs={'class':'parent6'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Miscellaneous';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
miscellaneous=sale_dollar
print(miscellaneous)

#TOTAL_DEALER_COST
data=page_dealer.findAll('table')
raw_column=data[4].text
raw_total=list(raw_column.split('\n'))
total_data=raw_total[624]
step_split=list(total_data.split(' '))
remove_data={'',':'}
column_split=[ele for ele in step_split if ele not in remove_data]
total_sales=column_split[4]
sale_dollar=list(total_sales.split("$"))
sale=sale_dollar[1:-1]
last_sale=sale_dollar[-1:]
last_sa=last_sale[0]
ccrf_sale=last_sa[0:-1]
total_columnh=column_split[0:4]
final_total=total_columnh+sale
final_total.append(ccrf_sale)
dealer_cost_t=final_total
print(dealer_cost_t)

#creating dataframe Dealer_cost

dict_dealer_cost=dict(zip(tab_head,zip(mechanical,collision,chemical,accessories,tires,miscellaneous,dealer_cost_t)))
tactical_sagment_dealercost=pd.DataFrame(dict_dealer_cost)


#tactical_sagment Transaction

action_filter_tactical=ActionChains(driver)
filter_tactical=driver.find_element_by_xpath('//*[@id="toggle"]')
action_filter_tactical.move_to_element(filter_tactical).click().perform()
#time.sleep(15)
driver.execute_script('document.getElementById("P1001_SALES_TYPE_METRIC").style.cssText = "display : block !important;"')
price_level=Select(WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="P1001_SALES_TYPE_METRIC"]'))))
#time.sleep(5)
price_level.select_by_index(0)
btn_filter=driver.find_element_by_xpath('//*[@id="SEARCH"]').click()

#Transaction_table

pageSource = driver.execute_script("return document.body.innerHTML;")
page_transaction=BeautifulSoup(pageSource,'html.parser')

#Mechanical_Trans
data=page_transaction.findAll(attrs={'class':'parent1'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Mechanical';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
mechanical_transaction=sale_dollar
print(mechanical_transaction)

#Collision_Trans
data=page_transaction.findAll(attrs={'class':'parent2'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Collision';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
collision_transaction=sale_dollar
print(collision_transaction)

#Chemicals_trans
data=page_transaction.findAll(attrs={'class':'parent3'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Chemicals';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
chemical_transaction=sale_dollar
print(chemical_transaction)

#Accessories_trans
data=page_transaction.findAll(attrs={'class':'parent4'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Accessories';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
accessories_transaction=sale_dollar
print(accessories_transaction)

#Tires_Trans
data=page_transaction.findAll(attrs={'class':'parent5'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Tires';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
tires_transaction=sale_dollar
print(tires_transaction)

#Miscellaneous_Trans
data=page_transaction.findAll(attrs={'class':'parent6'})
raw_column=data[0].text
column_split=list(raw_column.split('\n'))
column_split
remove_data={"var value = 'Miscellaneous';",""}
column_split= [ele for ele in column_split if ele not in remove_data]
sale_dollar=list(column_split[1].split('$'))
column_header=column_split[0]
sale_dollar.insert(0,column_header)
RAD=sale_dollar[1]
regions=RAD[0]
areas=RAD[1:3]
dealers=RAD[3:6]
sale_dollar.insert(1,regions)
sale_dollar.insert(2,areas)
sale_dollar.insert(3,dealers)
sale_dollar.remove(sale_dollar[4])
miscellaneous_transaction=sale_dollar
print(miscellaneous_transaction)

#TOTAL_Trans
data=page_transaction.findAll('table')
raw_column=data[4].text
raw_total=list(raw_column.split('\n'))
total_data=raw_total[624]
step_split=list(total_data.split(' '))
remove_data={'',':'}
column_split=[ele for ele in step_split if ele not in remove_data]
total_sales=column_split[4]
sale_dollar=list(total_sales.split("$"))
sale=sale_dollar[1:-1]
last_sale=sale_dollar[-1:]
last_sa=last_sale[0]
ccrf_sale=last_sa[0:-1]
total_columnh=column_split[0:4]
final_total=total_columnh+sale
final_total.append(ccrf_sale)
total_transaction=final_total
print(total_transaction)

#creating dataframe Transaction

dict_transaction=dict(zip(tab_head,zip(mechanical_transaction,collision_transaction,chemical_transaction,accessories_transaction,tires_transaction,miscellaneous_transaction,total_transaction)))
tactical_sagment_transaction=pd.DataFrame(dict_transaction)


#csv_file

final_sheet={'Part_sale_Dashboard':part_dashboard,'TACTICAL_SAGMENT_Dealer_Cost':tactical_sagment_dealercost,'TACTICAL_SAGMENT_Transaction':tactical_sagment_transaction}
write=pd.ExcelWriter("C:/Users/Admin/Desktop/MAYA-TASK/vw Scrap/output/part_sales.xlsx", engine='xlsxwriter')
for sheet_name in final_sheet.keys():
    final_sheet[sheet_name].to_excel(write, sheet_name=sheet_name, index=False)
write.save()