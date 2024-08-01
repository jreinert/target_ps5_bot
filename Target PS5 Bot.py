# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 19:49:33 2021

@author: reine
"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from PIL import ImageGrab
from win32com.client import Dispatch
import win32com

def screengrab():
    screen = ImageGrab.grab()
    screen.save('success.png')
    
def send_mail():
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '9892843354@vtext.com'
    mail.Subject = 'Wins!!'
    mail.Body = 'PS5 purchased'
    #mail.Attachments.Add('D:\success.png')
    mail.Send()
    
def get_device(driver, device, cvv):
    run = True
    while run == True:
        # go to device url and search for 'Pick it up' button
        driver.get(device)
        button = driver.find_elements(by=By.XPATH, value='//button[text()= "Pick it up"]')
        button = [i.text for i in button]
        
        # if/when 'Pick it up' button appears, complete purchase
        if 'Pick it up' in button:
            wait = WebDriverWait(driver, 5)
            driver.find_element(by=By.XPATH, value='//button[text()= "Pick it up"]').click()
            x = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[text()= "View cart & checkout"]')))
            driver.find_element(by=By.XPATH, value='//button[text()= "View cart & checkout"]').click()
            x = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[text()= "Check out"]')))
            driver.find_element(by=By.XPATH, value='//button[text()= "Check out"]').click()
            x = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[text()= "Place your order"]')))
            driver.find_element(by=By.XPATH, value='//button[text()= "Place your order"]').click()
            x = wait.until(EC.visibility_of_element_located((By.ID, 'creditCardInput-cvv')))
            x.send_keys(f'{cvv}')
            driver.find_element(by=By.XPATH, value='//button[text()= "Confirm"]').click()
            run = False
        time.sleep(5)
        

# set values for url and cvv
ps5_digi = 'https://www.target.com/p/playstation-5-digital-edition-console/-/A-81114596'
ps5_disc = 'https://www.target.com/p/playstation-5-console/-/A-81114595'
test = 'https://www.target.com/p/treasure-x-monster-gold-action-figure/-/A-81959648#lnk=sametab'
cvv = '473'

# create driver
driver = webdriver.Chrome(r"C:\Users\reine\Downloads\chromedriver_win32\chromedriver.exe")

# manually sign into target account
driver.get('https://www.target.com/circle')
time.sleep(30)

# check for device availability and purchase when available
get_device(driver, ps5_disc, cvv)

# close driver
driver.close()

# take screenshot of wins
screengrab()

# send text of wins
send_mail()





