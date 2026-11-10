from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
import time


casa = [9421995]


options = Options()
driver = webdriver.Chrome(options=options)


wait = WebDriverWait(driver,10)


for i in casa:
    driver.get("https://mi.tigo.com.sv/pago-express/facturas")
    
    time.sleep(3)

    id_count_input = wait.until(EC.presence_of_element_located((By.NAME, "ion-input-0")))
    id_count_input.send_keys(i)
    time.sleep(2)
    btn_siguiente = wait.until(EC.element_to_be_clickable((By.ID, "onSubmitChangeNameForm"))).click()

    time.sleep(15)

    btn_pay1 = wait.until(EC.element_to_be_clickable((By.ID, "payBtn")))
    btn_pay1.click()

    time.sleep(6)

    btn_metodo_pago = wait.until(EC.element_to_be_clickable((By.ID, "choosePayment0"))).click()
    time.sleep(5)


    number_card = wait.until(EC.presence_of_element_located((By.ID, "cardNumber")))
    expire_card = wait.until(EC.presence_of_element_located((By.ID, "expDate")))
    ccv_card = wait.until(EC.presence_of_element_located((By.ID, "cvc")))
    name_card = wait.until(EC.presence_of_element_located((By.ID, "nameOnCard")))
    number_card.send_keys("5197120001482809")
    expire_card.send_keys("1227")
    ccv_card.send_keys("556")
    name_card.send_keys("Mauricio Orantes")

    btn_pay2 = wait.until(EC.element_to_be_clickable((By.ID, "showVerifySection")))
    btn_pay2.click()

    time.sleep(5)
    btn_final_pay = wait.until(EC.element_to_be_clickable((By.ID, "continueButton")))
    btn_final_pay.click()
    time.sleep(1000)
