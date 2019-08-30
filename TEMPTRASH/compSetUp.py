from flask import Flask
import requests
from bs4 import BeautifulSoup
import csv
import shutil
import os
import glob
from shutil import copyfile
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions.interaction import KEY
from selenium.webdriver.common import keys
EMAIL = "nathan@certasun.com"
PASS = "Fac21408"
webdriverpath = os.path.join('\\Users', 'NathanWayne', 'Desktop', 'TestCSVPY', 'chromedriver.exe')
browser = webdriver.Chrome(executable_path = 'chromedriver.exe')
wait = WebDriverWait(browser, 30)
browser.get('https://www.office.com/')
element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="hero-banner-sign-in-to-office-365-link"]')))
sign_in_button = browser.find_element_by_xpath('//*[@id="hero-banner-sign-in-to-office-365-link"]')
sign_in_button.click()
email_field = browser.find_element_by_xpath('//*[@id="i0116"]')
email_field.send_keys(EMAIL)
#element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="hero-banner-sign-in-to-office-365-link"]')))
next_button = browser.find_element_by_xpath('//*[@id="idSIButton9"]')
next_button.click()
element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="i0281"]/div/div/div[1]/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div')))
pass_feild = browser.find_element_by_xpath('//*[@id="i0118"]')
pass_feild.send_keys(PASS)
sign_in_button_office = browser.find_element_by_xpath('//*[@id="i0281"]/div/div/div[1]/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div')
sign_in_button_office.click()
element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]')))
stay_signed_in_button = browser.find_element_by_xpath('//*[@id="idSIButton9"]')
stay_signed_in_button.click()
install_office_button = browser.find_element_by_xpath('//*[@id="install-dropdown-link"]/div/span')
install_office_button.click()
element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="install-item--installbutton-OrgId-DirectInstall"]/div/div/div[1]')))
office_365_apps_button = browser.find_element_by_xpath('//*[@id="install-item--installbutton-OrgId-DirectInstall"]/div/div/div[1]')
office_365_apps_button.click()
browser.get('https://products.office.com/en-us/microsoft-teams/download-app')
teams_download = browser.find_element_by_xpath('//*[@id="office-Hero5050-hc9qs8e"]/section/div/div/div/div/div/div[2]/div/a')
teams_download.click()
browser.get('https://www.office.com/?auth=2')
#sharepoint_button = browser.find_element_by_xpath('//*[@id="ShellSites_link"]')
#sharepoint_button.click()

browser.get('https://certasun.sharepoint.com/sites/Admin/SitePages/Home.aspx')
element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="spPageCanvasContent"]/div/div/div/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[1]/button')))
admin_sync_button = browser.find_element_by_xpath('//*[@id="spPageCanvasContent"]/div/div/div/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[1]/button')
#admin_sync_button = browser.find_element_by_xpath('//*[@id="id__150"]')
admin_sync_button.click()
element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div/button/div')))
time.sleep(2)
print("CLICK OPEN ONEDRIVE")
myin = input("PRESS ENTER TO CONTINUE")
print("CONTINUING")
sync_pop_up_x_button = browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div/button/div/i')
sync_pop_up_x_button.click()
time.sleep(2)
browser.get('https://certasun.sharepoint.com/sites/All')
#element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="spPageChromeAppDiv"]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[2]')))
#more_button = browser.find_element_by_xpath('//*[@id="spPageChromeAppDiv"]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[2]')
#more_button.click()
#element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="spPageChromeAppDiv"]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[3]')))
#sync_all_button = browser.find_element_by_xpath('//*[@id="spPageChromeAppDiv"]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[3]')
#sync_all_button.click()
element = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="id__310"]')))
sync_all_button = browser.find_element_by_xpath('//*[@id="id__310"]')
sync_all_button.click()
print("CLICK OPEN ONEDRIVE")
myin = input("PRESS ENTER TO CONTINUE")
print("CONTINUING")
element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[6]/div/div/div/div[2]/div[2]/div/div[1]/div/button/div')))
sync_pop_up_x_button = browser.find_element_by_xpath('/html/body/div[6]/div/div/div/div[2]/div[2]/div/div[1]/div/button/div')
sync_pop_up_x_button.click()
time.sleep(3)
browser.get('https://www.office.com/?auth=2')
onedrive_button = browser.find_element_by_xpath('//*[@id="ShellDocuments_link"]/span/ohp-icon-font/span')
onedrive_button.click()
onedrive_sync_button = browser.find_element_by_xpath('//*[@id="id__63"]')
onedrive_sync_button.click()
browser.get('https://fonts.google.com/specimen/Open+Sans?selection.family=Open+Sans')
family_selected_button = browser.find_element_by_xpath('//*[@id="collection-drawer-layout-container"]/gf-collection-drawer/header/h1')
font_download_button = browser.find_element_by_xpath('//*[@id="collection-drawer-layout-container"]/gf-collection-drawer/main/section/div[2]/div/div/a[2]')
font_download_button.click()
browser.get('https://123.hp.com/us/en/devices/ojpro7740')
get_printer_app_button = browser.find_element_by_xpath('//*[@id="download-button"]')
browser.switchTo().alert().accept()
