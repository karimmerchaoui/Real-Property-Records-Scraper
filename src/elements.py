from selenium.webdriver.support.ui import WebDriverWait

from selenium.common.exceptions import NoSuchElementException,ElementClickInterceptedException
from selenium.webdriver.common.by import By
import time
import random


def intercept_excep(xpath):
    try:
        xpath.click()
        return True
    except ElementClickInterceptedException:
        return False
def element_exists_id(driver, element_id):
    try:
        driver.find_element('css selector', f"#{element_id}")
        return True
    except NoSuchElementException:
        return False


def element_exists_xpath(driver, xpath):
    try:
        driver.find_element('xpath', xpath)
        return True
    except NoSuchElementException:
        return False

def element_exists_tag(driver, tag):
    try:
        driver.find_element(By.TAG_NAME, tag)
        return True
    except NoSuchElementException:
        return False

def click_close_button(browser):
    while element_exists_id(browser, "mat-mdc-dialog-0"):
        button = browser.find_element('xpath','//*[@id="mat-mdc-dialog-0"]/div/div/rpr-status-dialog/div/div/div[2]/div/div/button')
        button.click()
        time.sleep(random.uniform(0.3, 0.5))

def scroll(browser):
    scroll_px = 0
    down = random.uniform(900, 1200)
    up = random.uniform(-400, -300)
    while scroll_px < down:
        scroll_px = scroll_px + random.uniform(50, 200)
        browser.execute_script(f"window.scrollBy(0, {scroll_px});")
        time.sleep(0.2)

    # scrolling
    scroll_px = 0
    while scroll_px > -500:
        scroll_px = scroll_px - random.uniform(50, 200)
        browser.execute_script(f"window.scrollBy(0, {scroll_px});")
        time.sleep(0.2)
    time.sleep(2)


def login(browser):
    browser.get('https://auth.narrpr.com/auth/sign-in')

    # LOGGING IN
    time.sleep(2)
    WebDriverWait(browser, 25).until(lambda b: b.execute_script("return document.readyState") == "complete")
    if (element_exists_xpath(browser, '//*[@id="mat-mdc-dialog-0"]/div/div')):
        browser.get('https://auth.narrpr.com/auth/sign-in')
        time.sleep(random.uniform(3, 5))
        # LOGGING IN
    WebDriverWait(browser, 25).until(lambda b: b.execute_script("return document.readyState") == "complete")
    email = browser.find_element("css selector", "#SignInEmail")  # By ID
    e = "ben@msvproperties.net"
    email.clear()
    email.send_keys(e)
    time.sleep(0.5)
    password = browser.find_element("css selector", "#SignInPassword")  # By ID

    p = "Monsey108!"
    password.clear()
    password.send_keys(p)

    button = browser.find_element("css selector", "#SignInBtn")  # By ID
    button.click()
    time.sleep(1)