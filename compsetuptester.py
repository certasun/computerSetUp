from flask import Flask
import requests
from bs4 import BeautifulSoup
import csv
import shutil
import os
import glob
import pprint
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
from pynput.keyboard import Key, Controller
import subprocess
pp = pprint.PrettyPrinter(indent=2)
EMAIL = "nathan@certasun.com"
PASS = "Fac21408"
webdriverpath = os.path.join('\\Users', 'NathanWayne', 'Desktop', 'TestCSVPY', 'chromedriver.exe')
keyboard = Controller()

def loginOffice(email, password, browser1, wait1, longwait1):
    global browser
    global wait
    global longwait
    browser = browser1
    wait = wait1
    longwait = longwait1
    browser.get('https://www.office.com/')
    element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="hero-banner-sign-in-to-office-365-link"]')))
    sign_in_button = browser.find_element_by_xpath('//*[@id="hero-banner-sign-in-to-office-365-link"]')
    sign_in_button.click()
    email_field = browser.find_element_by_xpath('//*[@id="i0116"]')
    email_field.send_keys(email)
    #element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="hero-banner-sign-in-to-office-365-link"]')))
    next_button = browser.find_element_by_xpath('//*[@id="idSIButton9"]')
    next_button.click()
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0281"]/div/div/div[1]/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div')))
    pass_feild = browser.find_element_by_xpath('//*[@id="i0118"]')
    pass_feild.send_keys(password)
    sign_in_button_office = browser.find_element_by_xpath('//*[@id="i0281"]/div/div/div[1]/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div')
    sign_in_button_office.click()
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="idSIButton9"]')))
    stay_signed_in_button = browser.find_element_by_xpath('//*[@id="idSIButton9"]')
    stay_signed_in_button.click()
    return

############################################# TEST AREA ####################
def printerSetup(Windows_PIN):
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.5)
    keyboard.type("Microsoft Store")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.ctrl)
    keyboard.press('e')
    keyboard.release('e')
    keyboard.release(Key.ctrl)
    time.sleep(.5)
    for i in range(3):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(.5)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.type("awstestemail32@gmail.com")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(1)
    keyboard.type("printersetuppassword32!")
    time.sleep(1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(5)
    keyboard.type(str(Windows_PIN))
    time.sleep(10)
    keyboard.press(Key.ctrl)
    keyboard.press('e')
    keyboard.release('e')
    keyboard.release(Key.ctrl)
    time.sleep(.1)
    keyboard.type("HP Smart")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    for i in range(3):
        keyboard.press(Key.tab)
        keyboard.press(Key.tab)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)

    return

def syncAllSharepoint():
    browser.get('https://certasun.sharepoint.com/sites/All')
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="spPageChromeAppDiv"]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div/button')))
    time.sleep(2)
    more_button = browser.find_element_by_xpath('//*[@id="spPageChromeAppDiv"]/div/div/div[3]/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div/button')
    more_button.click()
    time.sleep(2)
    keyboard.press(Key.down)
    keyboard.release(Key.down)
    keyboard.press(Key.down)
    keyboard.release(Key.down)
    keyboard.press(Key.down)
    keyboard.release(Key.down)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.left)
    keyboard.release(Key.left)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    ################## MAY NOT NEED ON FRESH SETUP#######################
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    ###############################################################
    return

def downloadOffice(username):
    install_office_button = browser.find_element_by_xpath('//*[@id="install-dropdown-link"]/div/span')
    install_office_button.click()
    element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="install-item--installbutton-OrgId-DirectInstall"]/div/div/div[1]')))
    office_365_apps_button = browser.find_element_by_xpath('//*[@id="install-item--installbutton-OrgId-DirectInstall"]/div/div/div[1]')
    office_365_apps_button.click()
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(5)
    office_exe_path = os.path.join("\\Users", username, "Downloads", "Setup.Def.en-us_O365BusinessRetail_0c7a0c18-b54d-4727-976c-12e96e6185de_TX_PR_Platform_def_b_32_.exe")
    if os.path.isfile(office_exe_path) == False:
        for i in range(13):
            keyboard.press(Key.shift)
            keyboard.press(Key.tab)
            keyboard.release(Key.tab)
            keyboard.release(Key.shift)
            time.sleep(.25)
        time.sleep(2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
    return

def installOffice2(username):
    time.sleep(5)
    path = os.path.join("\\Users", username, "Downloads", "Setup.Def.en-us_O365BusinessRetail_0c7a0c18-b54d-4727-976c-12e96e6185de_TX_PR_Platform_def_b_32_.exe")
    subprocess.run(path)
    return
############################# DEPRACATED #############################
def installOffice():
    time.sleep(2)
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(2)
    keyboard.type("File")
    time.sleep(2)
    for i in range(8):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        time.sleep(.1)
    time.sleep(.5)
    keyboard.press(Key.down)
    keyboard.release(Key.down)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.press(Key.enter)
    time.sleep(.5)
    for i in range(6):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.space)
    keyboard.release(Key.space)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(5)
    ###############################################REAL BLOCK #############################################
    #keyboard.press(Key.left)
    #keyboard.release(Key.left)
    #time.sleep(.1)
    #keyboard.press(Key.enter)
    #keyboard.release(Key.enter)
    ###############################
    ######################TEST BLOCK##########################
    ##############################
    return

def downloadTeams():
    browser.get('https://products.office.com/en-us/microsoft-teams/download-app')
    teams_download = browser.find_element_by_xpath('//*[@id="office-Hero5050-hc9qs8e"]/section/div/div/div/div/div/div[2]/div/a')
    teams_download.click()
    teams_download2 = browser.find_element_by_xpath('//*[@id="office-DesktopAppDownload-wqkfqkk"]/div/div/div/div/a')
    teams_download2.click()
    time.sleep(3)
    for i in range(62):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        i = i +1
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    return

def installTeams2():
    path = os.path.join("\\Users", username, "Downloads", "Teams_windows_x64.exe")
    subprocess.run(path)
    return

def installTeams():
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(2)
    keyboard.type("File")
    time.sleep(2)
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.1)
    for i in range(3):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    for i in range(6):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.5)
    keyboard.press(Key.space)
    keyboard.release(Key.space)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    return

def syncAdminSharepoint():
    #loginOffice()
    browser.get('https://certasun.sharepoint.com/sites/Admin/SitePages/Home.aspx')
    time.sleep(2)
########################### PUT BACK #####################
    keyboard.press(Key.ctrl)
    keyboard.press('w')
    keyboard.release('w')
    keyboard.release(Key.ctrl)
    ############################################
    time.sleep(1)
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="spPageCanvasContent"]/div/div/div/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[1]/button')))
    admin_sync_button = browser.find_element_by_xpath('//*[@id="spPageCanvasContent"]/div/div/div/div/div/div/div/div/div/div/div/div[4]/div/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div/div/div[1]/div[1]/button')
    #admin_sync_button = browser.find_element_by_xpath('//*[@id="id__150"]')
    admin_sync_button.click()
    element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div/button/div')))
    time.sleep(2)
    keyboard.press(Key.left)
    keyboard.release(Key.left)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    ################## MAY NOT NEED ON FRESH SETUP#######################
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    ###############################################################
    #keyboard.press(Key.enter)
    #keyboard.release(Key.enter)
    #time.sleep(10)
    #for i in range(2):
#        keyboard.press(Key.tab)
#        keyboard.release(Key.tab)
#    time.sleep(.1)
#    keyboard.press(Key.space)
#    keyboard.release(Key.space)
#    time.sleep(2)
#    keyboard.press(Key.alt)
#    keyboard.press(Key.f4)
#    keyboard.release(Key.f4)
#    keyboard.release(Key.alt)
#    time.sleep(.5)
#    sync_pop_up_x_button = browser.find_element_by_xpath('/html/body/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div/button/div/i')
#    sync_pop_up_x_button.click()
#    time.sleep(2)
    return

def runAll():
    loginOffice()
    downloadOffice()
    downloadTeams()
    syncAdminSharpoint()
    syncAllSharepoint()
    #sync
    return

def syncOneDrive():
    browser.get('https://office.com')
    element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ShellDocuments_link"]/span/ohp-icon-font/span')))
    onedrive_button = browser.find_element_by_xpath('//*[@id="ShellDocuments_link"]/span/ohp-icon-font/span')
    onedrive_button.click()
    time.sleep(10)
    for i in range(14):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
        i = i + 1
    for i in range(2):
        keyboard.press(Key.right)
        keyboard.release(Key.right)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.left)
    keyboard.release(Key.left)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(10)
    for i in range(2):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.1)
    keyboard.press(Key.space)
    keyboard.release(Key.space)
    time.sleep(2)
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.5)
    keyboard.press(Key.space)
    keyboard.release(Key.space)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(2)
    keyboard.press(Key.ctrl)
    keyboard.press('w')
    keyboard.release('w')
    keyboard.release(Key.ctrl)
    time.sleep(1)
    ################## MAY NOT NEED ON FRESH SETUP#######################
    keyboard.press(Key.right)
    keyboard.release(Key.right)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    ###############################################################
    return

############ CHANGE PATH #########
def installOpenSans(username):
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(2)
    #path = os.path.join('Users', username, 'Downloads', 'Open_Sans')
    path = os.path.join('D:\\', 'Open_Sans')
    keyboard.type(path)
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    num_font_files = 0
    os.chdir(path)
    for file in glob.glob("*.ttf"):
        num_font_files = num_font_files+1
    for i in range(num_font_files):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
        keyboard.press(Key.shift)
        keyboard.press(Key.f10)
        keyboard.release(Key.f10)
        keyboard.release(Key.shift)
        for i in range(3):
            keyboard.press(Key.down)
            keyboard.release(Key.down)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        time.sleep(5)
    return


def copyfile(username, file, dest_path):
    source_path = os.path.join('\\Users', username,'Certasun', 'Admin - Documents', 'Templates', 'For new blank documents', file)
    shutil.copy2(source_path, dest_path)
    return

#######MOVE TEMPLATES#########
def moveTemplates(username):
    file1 = "Normal.dotm"
    file2 = "NormalEmail.dotm"
    file3 = "Blank.potx"
    file4 = "Book.xltm"
    file5 = "Certasun Formal Letter.dotx"
    file6 = "Certasun Informal Letter.dotx"
    source_path = os.path.join('\\Users', username,'Certasun', 'Admin - Documents', 'Templates')
    dest_path = os.path.join('\\Users', username, 'AppData', 'Roaming', 'Microsoft', 'Templates')
    excel_dest_path = os.path.join("\\Users", username, "AppData", "Roaming", "Microsoft", "Excel", "XLSTART")
    copyfile(username, file1, dest_path)
    copyfile(username, file2, dest_path)
    copyfile(username, file3, dest_path)
    shutil.copy2(source_path+"\\"+file5, dest_path)
    shutil.copy2(source_path+"\\"+file6, dest_path)
    copyfile(username, file4, excel_dest_path)
    ######SET EXCEL SHEET
    #keyboard.press(Key.cmd_r)
    #keyboard.release(Key.cmd_r)
    #time.sleep(.5)
    #keyboard.type(excel_dest_path)
    #time.sleep(.5)
    #keyboard.press(Key.enter)
    #keyboard.release(Key.enter)
    #time.sleep(2)
    #keyboard.press(Key.down)
    #keyboard.release(Key.down)
    #keyboard.press(Key.enter)
    #keyboard.release(Key.enter)
    #time.sleep(3)
    #keyboard.press(Key.ctrl)
    #keyboard.press('n')
    #keyboard.release('n')
    #keyboard.release(Key.ctrl)
    #time.sleep(2)
    return
    ################## SET OPTIONS IN EXCEL #############
def excelSetup(username):
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.5)
    keyboard.type("excel")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.press(Key.alt)
    keyboard.press('f')
    keyboard.release('f')
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.shift)
    keyboard.press('t')
    keyboard.release('t')
    keyboard.release(Key.shift)
    time.sleep(.5)
    for i in range(21):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.space)
    keyboard.release(Key.space)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.alt)
    keyboard.press('f')
    keyboard.release('f')
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.shift)
    keyboard.press('t')
    keyboard.release('t')
    keyboard.release(Key.shift)
    for i in range(4):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    for i in range(10):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.type("C:\\Users\\" + username + "\\OneDrive - Certasun")
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.type("C:\\Users\\" + username + "\\Certasun\\Admin - Documents\\Templates")
    time.sleep(.5)
    for i in range(7):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(1)
    return
    ################ SETTINGS FOR WORD #############
def word_Powerpoint_Setup(username):
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.5)
    keyboard.type("Word")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press('f')
    keyboard.release('f')
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.shift)
    keyboard.press('t')
    keyboard.release('t')
    keyboard.release(Key.shift)
    time.sleep(.5)
    for i in range(3):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    for i in range(11):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.type("C:\\Users\\"+username+"\\OneDrive - Certasun\\")
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.type("C:\\Users\\"+username+"\\Certasun\\Admin - Documents\\Templates")
    time.sleep(.5)
    for i in range(3):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.5)
    keyboard.type("Powerpoint")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press('f')
    keyboard.release('f')
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.shift)
    keyboard.press('t')
    keyboard.release('t')
    keyboard.release(Key.shift)
    time.sleep(.5)
    keyboard.press('s')
    keyboard.release('s')
    time.sleep(.5)
    for i in range(10):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.type("C:\\Users\\"+username+"\\OneDrive - Certasun")
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    keyboard.type("C:\\Users\\"+username+"\\Certasun\\Admin - Documents\\Templates")
    time.sleep(.5)
    for i in range(3):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    return
    #################OUTLOOK SETTINGS#######
def outlookSetup():
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.5)
    keyboard.type("Outlook")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.press(Key.alt)
    keyboard.press('f')
    keyboard.release('f')
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.shift)
    keyboard.press('t')
    keyboard.release('t')
    keyboard.release(Key.shift)
    time.sleep(1)
    for i in range(10):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(1)
    for i in range(7):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    keyboard.type('Certasun Templates')
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    for i in range(8):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    for i in range(3):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    for i in range(116):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    for i in range(9):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    ########### Set templates for calendar view#######
    time.sleep(.5)
    keyboard.press(Key.ctrl)
    keyboard.press('2')
    keyboard.release('2')
    keyboard.release(Key.ctrl)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press('f')
    keyboard.release('f')
    keyboard.release(Key.alt)
    time.sleep(.5)
    keyboard.press(Key.shift)
    keyboard.press('t')
    keyboard.release('t')
    keyboard.release(Key.shift)
    time.sleep(.5)
    for i in range(10):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    for i in range(7):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.type('Certasun Templates')
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    for i in range(8):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    for i in range(3):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    for i in range(116):
        keyboard.press(Key.down)
        keyboard.release(Key.down)
    time.sleep(.5)
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(.5)
    for i in range(9):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(1)
    return

def testWordSetup(username):
    save_location_test = " "
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.1)
    keyboard.type("word")
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(1)
    keyboard.type("TEST")
    time.sleep(1)
    keyboard.press(Key.ctrl)
    keyboard.press('s')
    keyboard.release('s')
    keyboard.release(Key.ctrl)
    time.sleep(.5)
    keyboard.type("WORDTESTSAVELOCATION")
    time.sleep(.5)
    for i in range(4):
        keyboard.press(Key.tab)
        keyboard.release(Key.tab)
    time.sleep(.1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(5)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(5)
    path = os.path.join("C:\\Users", username, "OneDrive - Certasun")
    if os.path.isfile(path+ "\\WORDTESTSAVELOCATION.docx") == False:
        save_location_test = "FAIL"
    elif os.path.isfile(path+ "\\WORDTESTSAVELOCATION.docx") == True:
        save_location_test = "PASS"
        os.chdir(path)
        os.remove("WORDTESTSAVELOCATION.docx")
    print(save_location_test)
    return save_location_test

def testPPTSetup(username):
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(.1)
    keyboard.type("Powerpoint")
    time.sleep(.5)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(2)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(1)
    keyboard.type('TEST')
    time.sleep(1)
    keyboard.press(Key.ctrl)
    keyboard.press('s')
    keyboard.release('s')
    keyboard.release(Key.ctrl)
    time.sleep(.5)
    keyboard.type('PowerpointTEST')
    time.sleep(.5)
    #for i in range(4):
#        keyboard.press(Key.tab)
    #    keyboard.release(Key.tab)
    time.sleep(1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(5)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(5)
    path = os.path.join("C:\\Users", username, "OneDrive - Certasun")
    if os.path.isfile(path+"\\PowerpointTEST.pptx") == False:
        save_location_test = "FAIL"
    elif os.path.isfile(path+"\\PowerpointTEST.pptx") == True:
        save_location_test = "PASS"
        os.chdir(path)
        os.remove("PowerpointTEST.pptx")
    print("PPT: "+ save_location_test)
    return save_location_test

def testExcelSave(username):
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    time.sleep(1)
    keyboard.type('Excel')
    time.sleep(1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(3)
    keyboard.type("TEST")
    time.sleep(1)
    keyboard.press(Key.ctrl)
    keyboard.press('s')
    keyboard.release('s')
    keyboard.release(Key.ctrl)
    time.sleep(2)
    keyboard.type("ExcelSaveTest")
    time.sleep(1)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(7)
    keyboard.press(Key.alt)
    keyboard.press(Key.f4)
    keyboard.release(Key.f4)
    keyboard.release(Key.alt)
    time.sleep(2)
    i = 0
    path = os.path.join("C:\\Users\\" + username + "\\OneDrive - Certasun")
    ret = "FAIL"
    while(i < 60):
        if os.path.isfile(path+"\\ExcelSaveTest.xlsx") == False:
            time.sleep(.25)
            i = i + 1
            print(i)
        else:
            ret = "PASS"
            os.chdir(path)
            os.remove(path+"\\ExcelSaveTest.xlsx")
            break
    print(ret)
    return ret

def testOneDriveSync(username):
    path = os.path.join("C:\\Users", username, "OneDrive - Certasun")
    if os.path.isdir(path) == True:
        ret = "PASS"
    elif os.path.isdir(path) == False:
        ret = "FAIL"

    return ret

def testSharepointSync(username):
    admin_path = os.path.join("\\Users", username, "Certasun", "Admin - Documents")
    all_path = os.path.join("\\Users", username, "Certasun", "All - Documents")
    sync_statai = {
    "ADMIN" : "",
    "ALL" : ""
    }
    if os.path.isdir(admin_path) == True:
        sync_statai['ADMIN'] = "PASS"
    elif os.path.isdir(admin_path) == False:
        sync_statai["ADMIN"] = "FAIL"
    if os.path.isdir(all_path) == True:
        sync_statai['ALL'] = "PASS"
    elif os.path.isdir(all_path) == False:
        sync_statai['ALL'] = "FAIL"
    return sync_statai

def testTeamsDownload(username):
    path = os.path.join("\\Users", username, "AppData", "Local", "Microsoft", "Teams", "current", "Teams.exe")
    if os.path.isfile(path) == True:
        ret = "PASS"
    elif os.path.isfile(path) ==False:
        ret = "FAIL"
    return ret

def testOfficedownloads():
    word_path = os.path.join("\\Program Files", "Microsoft Office", "root", "Office16", "WINWORD.exe")
    excel_path = os.path.join("\\Program Files", "Microsoft Office", "root", "Office16", "EXCEL.exe")
    outlook_path = os.path.join("\\Program Files", "Microsoft Office", "root", "Office16", "OUTLOOK.exe")
    ppt_path = os.path.join("\\Program Files", "Microsoft Office", "root", "Office16", "POWERPNT.exe")
    download_statai = {
    "WORD" : "",
    "PPT" : "",
    "OUTLOOK" : "",
    "EXCEL" : "",
    }
    if os.path.isfile(word_path) == True:
        download_statai['WORD'] = "PASS"
    elif os.path.isfile(word_path) ==False:
        download_statai['WORD'] = "FAIL"
    if os.path.isfile(excel_path) == True:
        download_statai['EXCEL'] = "PASS"
    elif os.path.isfile(excel_path) ==False:
        download_statai['EXCEL'] = "FAIL"
    if os.path.isfile(outlook_path) == True:
        download_statai['OUTLOOK'] = "PASS"
    elif os.path.isfile(outlook_path) ==False:
        download_statai['OUTLOOK'] = "FAIL"
    if os.path.isfile(ppt_path) == True:
        download_statai['PPT'] = "PASS"
    elif os.path.isfile(ppt_path) ==False:
        download_statai['WORD'] = "FAIL"
    return download_statai

def testFontInstall(username):
    font_files = list()
    missing_font_files = list()
    font_folder_path = os.path.join("\\Users", username, "AppData", "Local", "Microsoft", "Windows", "Fonts")
    file1 = "OpenSans-SemiBoldItalic.ttf"
    file2 = "OpenSans-SemiBold.ttf"
    file3 = "OpenSans-Regular.ttf"
    file4 = "OpenSans-LightItalic.ttf"
    file5 = "OpenSans-Light.ttf"
    file6 = "OpenSans-RegularItalic.ttf"
    file7 = "OpenSans-ExtraBoldItalic.ttf"
    file8 = "OpenSans-ExtraBold.ttf"
    file9 = "OpenSans-BoldItalic.ttf"
    file10 = "OpenSans-Bold.ttf"
    font_files.append(file1)
    font_files.append(file2)
    font_files.append(file3)
    font_files.append(file4)
    font_files.append(file5)
    font_files.append(file6)
    font_files.append(file7)
    font_files.append(file8)
    font_files.append(file9)
    font_files.append(file10)
    os.chdir(font_folder_path)
    for row in font_files:
        if os.path.isfile(row) == False:
            missing_font_files.append(row)
    if len(missing_font_files) == 0:
        missing_font_files.append("All Fonts Installed")
        ret = missing_font_files
    else:
        ret = missing_font_files
    return ret

def testFunctions(username):
    Test_Status = {
    'Word Status': "",
    'PowerPoint Status' : "",
    "Outlook Status" : "",
    "Excel Status" : "",
    "Teams Status" : "",
    "Word Default Save Location Status" : "",
    "PowerPoint Save Location Status" : "",
    "Excel Default Save Status" : "",
    "OneDrive Sync Status" : "",
    "Admin Sharepoint Sync Status" : "",
    "All Sharepoint Sync Status" : "",
    }
    word_setup_status = testWordSetup(username)
    if word_setup_status == "PASS":
        Test_Status['Word Default Save Location Status'] = "Save Location Set Correctly"
    else:
        Test_Status['Word Default Save Location Status'] = "ERROR Save Location Not Set Correctly"
    ppt_setup_status = testPPTSetup(username)
    if ppt_setup_status == "PASS":
        Test_Status['PowerPoint Save Location Status'] = "Save Location Set Correctly"
    else:
         Test_Status['PowerPoint Save Location Status'] = "ERROR Save Location Not Set Correctly"
    teams_download_status = testTeamsDownload(username)
    if teams_download_status == "PASS":
        Test_Status['Teams Status'] = "Microsoft Teams Downloaded Successfully"
    else:
        Test_Status['Teams Status'] = "ERROR Microsoft Teams Downloaded Unsuccessfully"
    office_downloads_statai = testOfficedownloads()
    if office_downloads_statai['WORD'] == "PASS":
        Test_Status['Word Status'] = "Installed Correctly"
    else:
        Test_Status['Word Status'] = "ERROR Installed Incorrectly"
    if office_downloads_statai['EXCEL'] == "PASS":
        Test_Status['Excel Status'] = "Installed Correctly"
    else:
        Test_Status['Excel Status'] = "ERROR Installed Incorrectly"
    if office_downloads_statai['OUTLOOK'] == "PASS":
        Test_Status['Outlook Status'] = "Installed Correctly"
    else:
        Test_Status['Outlook Status'] = "ERROR Installed Incorrectly"
    if office_downloads_statai['PPT'] == "PASS":
        Test_Status['PowerPoint Status'] = "Installed Correctly"
    else:
        Test_Status['PowerPoint Status'] = "ERROR Installed Incorrectly"
    onedrive_sync_status = testOneDriveSync(username)
    if onedrive_sync_status == "PASS":
        Test_Status['OneDrive Sync Status'] = "Sync Successful"
    else:
        Test_Status['OneDrive Sync Status'] = "ERROR Sync Failed"
    sharepoint_sync_status = testSharepointSync(username)
    pp.pprint(sharepoint_sync_status)
    if sharepoint_sync_status['ADMIN'] == "PASS":

        Test_Status['Admin Sharepoint Sync Status'] = "Sync Successful"
    else:
        Test_Status['Admin Sharepoint Sync Status'] = "ERROR Sync Failed"
    if sharepoint_sync_status['ALL'] == "PASS":
        Test_Status['All Sharepoint Sync Status'] = "Sync Successful"
    else:
        Test_Status['All Sharepoint Sync Status'] = "ERROR Sync Failed"
    excel_save_status = testExcelSave(username)
    if excel_save_status == "PASS":
        Test_Status['Excel Default Save Status'] = "Save Location Set Correctly"
    else:
        Test_Status['Excel Default Save Status'] = "Error Save Location Not Set Correctly"
    font_install_status = testFontInstall(username)
    Test_Status['Font Status'] = font_install_status
    path = os.path.join("D:\\")
    os.chdir(path)
    with open("setup_status.txt", "w+") as text_file:
        print("Drive Sync's: ", file = text_file)
        print(f"    All Sharepoint Sync Status: {Test_Status['All Sharepoint Sync Status']}", file = text_file)
        print(f"    Admin Sharepoint Sync Status: {Test_Status['Admin Sharepoint Sync Status']}", file = text_file)
        print("     Onedrive Sync Status: {}".format(Test_Status['OneDrive Sync Status']), file = text_file)
        print(f"WORD", file = text_file)
        print(f"    Word Installation: {Test_Status['Word Status']}", file = text_file)
        print(f"    Word Default Save Location: {Test_Status['Word Default Save Location Status']}", file = text_file)
        print(f"PPT", file=text_file)
        print(f"    PowerPoint Installation: {Test_Status['PowerPoint Status']}", file = text_file)
        print(f"    PowerPoint Default Save Location: {Test_Status['PowerPoint Save Location Status']}", file = text_file)
        print(f"EXCEL", file=text_file)
        print(f"    Excel Installation: {Test_Status['Excel Status']}", file = text_file)
        print(f"    Excel Default Save Location: {Test_Status['Excel Default Save Status']}", file = text_file)
        print(f"OutLook Installation: {Test_Status['Outlook Status']}", file = text_file)
        print(f"Teams Installation: {Test_Status['Teams Status']}", file = text_file)
    time.sleep(2)
    keyboard.press(Key.cmd_r)
    keyboard.release(Key.cmd_r)
    #dir_path = os.path.dirname(os.path.realpath(__file__))
    dir_path = os.path.join("D:\\")
    time.sleep(1)
    #keyboard.type(dir_path)
    time.sleep(.5)
    #keyboard.press(Key.enter)
    #keyboard.release(Key.enter)
    time.sleep(1)
    #dir_path = os.path.dirname(os.path.realpath(__file__))
    subprocess.run("notepad "+ dir_path+"\\setup_status.txt")
    return

#####   #######################################################################
