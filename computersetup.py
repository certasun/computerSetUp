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
#from selenium.webdriver.common import key
#from pyrobot import Robot
import subprocess
from pynput.keyboard import Key, Controller
from compsetuptester import testExcelSave, testPPTSetup, loginOffice, downloadOffice, downloadTeams, syncAdminSharepoint, syncAllSharepoint, keyboard, syncOneDrive, installOffice2, copyfile, moveTemplates, excelSetup, word_Powerpoint_Setup, outlookSetup, testFunctions, printerSetup, syncOneDrive, installOpenSans
username = input("Type your name in this format FirstLast capitilize first letters of each name: ")
email = input("Type your Certasun email: ")
password = input("Type your password: ")
#username = "NathanWayne"
#email = "nathan@certasun.com"
#password = "Fac21408"
browser = webdriver.Chrome()
wait = WebDriverWait(browser, 3)
longwait = WebDriverWait(browser, 30)
loginOffice(email, password, browser, wait, longwait)
downloadOffice(username)
installOffice2(username)
stop = input("PRESS ENTER TO CONTINUE")
print("DONE")
downloadTeams()
syncOneDrive()
syncAdminSharepoint()
syncAllSharepoint()
moveTemplates(username)
excelSetup(username)
word_Powerpoint_Setup(username)
outlookSetup()
testFunctions(username)
#printerSetup(112032)
#installOpenSans(username)
