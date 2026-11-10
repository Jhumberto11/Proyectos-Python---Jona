from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
import time


numeros = [78604878,70529131,78515721,78036808,78417054]
casa = [22241416]


options = Options()
driver = webdriver.Chrome(options=options)


wait = WebDriverWait(driver,10)


for i in numeros:
    driver.get("https://sv.mipagoclaro.com/?utm_source=MICLAROWEB&utm_medium=ICONOACCESO&utm_campaign=MCWACCESOS")
    time.sleep(3)

    tel_space = wait.until(EC.presence_of_element_located((By.ID, "phone__input")))
    tel_space.send_keys(i)
    time.sleep(2)
    btn_siguiente = wait.until(EC.element_to_be_clickable((By.ID, "btnAceptarHome"))).click()

    time.sleep(5)

    btn_metodo_pago = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-testid='invoices__pay-button']"))).click()
    time.sleep(5)


    number_card = wait.until(EC.presence_of_element_located((By.ID, "pp_card")))
    expire_card = wait.until(EC.presence_of_element_located((By.ID, "pp_expire")))
    ccv_card = wait.until(EC.presence_of_element_located((By.ID, "pp_cvv")))
    email = wait.until(EC.presence_of_element_located((By.ID, "pp_email")))
    btn_pay = wait.until(EC.element_to_be_clickable((By.ID, "continueButton")))
    number_card.send_keys("5197120001482809")
    expire_card.send_keys("1227")
    ccv_card.send_keys("556")
    email.send_keys("imosa253@gmail.com")
    btn_pay.click()


    time.sleep(5)
    btn_final_pay = wait.until(EC.element_to_be_clickable((By.ID, "continueButton")))
    btn_final_pay.click()
    time.sleep(10)
