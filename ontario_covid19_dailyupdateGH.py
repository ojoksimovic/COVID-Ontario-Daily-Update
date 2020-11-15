#! python3
import selenium
import time
import datetime
import glob
import os
import openpyxl
import smtplib
import sys
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options  

To_Emails = [**INSERT EMAIL ADDRESSES**]

options = Options()
options.headless = True
driver = webdriver.Chrome(options=options, executable_path=r'C:\Users\olive\AppData\Local\Programs\Python\Python38\chromedriver.exe')

wait = WebDriverWait(driver, 30)
driver.get("https://www.ontario.ca/page/2019-novel-coronavirus")
time.sleep(5)

COVUpdateElement = driver.find_element_by_xpath("//div[@id='pagebody']/p[contains(text(),'Last updated: ')]")    
COVUpdate = COVUpdateElement.text
PMCOVPositiveElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[3]/td[2]")
PMCOVPositive = ("".join(re.findall('\d+', PMCOVPositiveElement.text)))
PMCOVResolvedElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[4]/td[2]")
PMCOVResolved = ("".join(re.findall('\d+', PMCOVResolvedElement.text)))
PMCOVDeceasedElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[5]/td[2]")
PMCOVDeceased = ("".join(re.findall('\d+', PMCOVDeceasedElement.text)))


PMTotal_Cases = (int(PMCOVPositive) + int(PMCOVResolved) + int(PMCOVDeceased))



while "5:30" in COVUpdate:
    time.sleep(60)
    driver.refresh()
    time.sleep(5)
    COVUpdateElement = driver.find_element_by_xpath("//div[@id='pagebody']/p[contains(text(),'Last updated: ')]")

if "10:30" in COVUpdate:
    COVUpdateElement = driver.find_element_by_xpath("//div[@id='pagebody']/p[contains(text(),'Last updated: ')]")
    COVUpdate = COVUpdateElement.text
    COVNegativeElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[1]/td[2]")
    COVNegative = ("".join(re.findall('\d+', COVNegativeElement.text)))
    COVInvestigationElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[2]/td[2]")
    COVInvestigation = ("".join(re.findall('\d+', COVInvestigationElement.text)))
    COVPositiveElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[3]/td[2]")
    COVPositive = ("".join(re.findall('\d+', COVPositiveElement.text)))
    COVResolvedElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[4]/td[2]")
    COVResolved = ("".join(re.findall('\d+', COVResolvedElement.text)))
    COVDeceasedElement = driver.find_element_by_xpath("//div[@id='pagebody']/table/tbody/tr[5]/td[2]")
    COVDeceased = ("".join(re.findall('\d+', COVDeceasedElement.text)))

    Total_Cases = (int(COVPositive) + int(COVResolved) + int(COVDeceased))
    
    New_Cases = Total_Cases - PMTotal_Cases
    
    print('Good morning!\n' + COVUpdate + '\nNumber of negative cases: ' + COVNegative + '\nNumber cases under investigation: ' + COVInvestigation + '\nNumber of positive cases: ' + COVPositive + '\nNumber of resolved cases: ' + COVResolved + '\nNumber of deceased cases: ' + COVDeceased + '\nTotal cases in Ontario are currently: ' + str(Total_Cases))
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    type(smtpObj)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(INSERT EMAIL, INSERT PASSWORD)
    smtpObj.sendmail(INSERT EMAIL, To_Emails,
    'Subject: \nMorning Update\n' + COVUpdate + '\nNew cases in Ontario: ' + str(New_Cases) + '\nNegative cases: ' + COVNegative +'\nUnder investigation: ' + COVInvestigation + '\nPositive cases: ' + COVPositive + '\nResolved cases: ' + COVResolved + '\nDeceased cases: ' + COVDeceased + '\nTotal cases in Ontario: ' + str(Total_Cases))
    smtpObj.quit()
driver.quit()
sys.exit()
