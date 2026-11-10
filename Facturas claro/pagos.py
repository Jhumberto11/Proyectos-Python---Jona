from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from numeros import numeros


card = {
    "nombre" : "Mauricio Orantes",
    "numero" : "5197 1200 0148 2809",
    "expire" : "1227",
    "cvv": "556",
    "email": "imosa253@gmail.com"
 
}


driver_path =r"C:\Users\imosa\Downloads\chromedriver-win64 (9)\chromedriver-win64\chromedriver.exe"

chrome_service = Service(driver_path)

options = Options()


prefs = {
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_setting_values.geolocation": 2,
    "profile.default_content_setting_values.media_stream": 2
}

options.add_experimental_option("prefs", prefs)


driver = webdriver.Chrome(service=chrome_service, options= options)

wait = WebDriverWait(driver=driver , timeout=20)

driver.get("https://sv.mipagoclaro.com/?utm_source=MICLAROWEB&utm_medium=ICONOACCESO&utm_campaign=MCWACCESOS#/")


for i in numeros :
    num_to_pay = wait.until(EC.presence_of_element_located((By.ID,"phone__input")))
    button_to_pass = wait.until(EC.presence_of_element_located((By.ID, "btnAceptarHome")))

    num_to_pay.send_keys(i)
    time.sleep(2)
    button_to_pass.click()
    time.sleep(4)


    pay_type = wait.until( EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="invoices__pay-button"]')))
    time.sleep(2)
    pay_type.click()

    
    card_input = wait.until(EC.presence_of_element_located((By.ID, "pp_card")))
    card_expire_input = wait.until(EC.presence_of_element_located((By.ID, "pp_expire")))
    card_cvv_input = wait.until(EC.presence_of_element_located((By.ID, "pp_cvv")))
    card_email_input = wait.until(EC.presence_of_element_located((By.ID, "pp_email")))
    button_to_pay = wait.until(EC.presence_of_element_located((By.ID , "continueButton")))

    card_input.send_keys(card.get("numero"))
    card_cvv_input.send_keys(card.get("cvv"))
    card_email_input.send_keys(card.get("email"))
    card_expire_input.send_keys(card.get("expire"))





    time.sleep(1)
    button_to_pay.click()
    time.sleep(20)

    driver.get("https://sv.mipagoclaro.com/?utm_source=MICLAROWEB&utm_medium=ICONOACCESO&utm_campaign=MCWACCESOS#/")





time.sleep(1000)


