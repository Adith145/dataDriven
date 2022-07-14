import openpyxl
import xlOperation
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

driver = webdriver.Chrome(executable_path=r"C:\seleniumwebdriver\chromedriver.exe")
action=ActionChains(webdriver)
driver.implicitly_wait(5)
driver.get(r"https://www.amazon.in/")
driver.maximize_window()
driver.implicitly_wait(10)
path =r"C:\Users\Admin\OneDrive\Desktop\report.xlsx"

rows = xlOperation.getRowCount(path, "Sheet1")

    # perform or read the value from excel file and pass to application
wb = openpyxl.load_workbook(r"C:\Users\Admin\OneDrive\Desktop\report.xlsx")
sheet = wb.active
x1 = sheet['B1'].value
x2 = sheet['B2'].value
x3 = sheet['B3'].value

t1=driver.title
a =driver.find_element(by=By.XPATH, value=" //input[@id='twotabsearchtextbox']")
action.scroll_to_element(a)
a.send_keys("oneplus Mobile under 30000")
driver.implicitly_wait(5)
driver.find_element(by=By.XPATH, value=" //input[@id='nav-search-submit-button']").click()
driver.implicitly_wait(3)
t2 = driver.title
driver.find_element(by=By.XPATH, value=" //span[contains(text(),'Featured')]").click()
t3=driver.title
driver.implicitly_wait(3)
driver.find_element(by=By.XPATH, value=" //a[@id='s-result-sort-select_4']").click()
driver.implicitly_wait(5)

sheet = wb.active

if x1==t1 :
    sheet['C1']="pass"
else:
    sheet['C1']="fail"

if x2==t2 :
    sheet['C2']="pass"
else:
    sheet['C2']="fail"

if x3==t3 :
    sheet['C3']="pass"
else:
    sheet['C3']="fail"

driver.implicitly_wait(5)

wb.save(r"C:\Users\Admin\OneDrive\Desktop\report.xlsx")

driver.close()






