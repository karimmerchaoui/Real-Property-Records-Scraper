
import random
import time
from selenium.webdriver.support.ui import WebDriverWait
from elements import element_exists_xpath
from selenium.common.exceptions import ElementClickInterceptedException

def wait_for_json(driver, timeout=25):

    driver.execute_script("""
        window.jsonLoaded = false;
        let open = XMLHttpRequest.prototype.open;
        XMLHttpRequest.prototype.open = function() {
            if (arguments[1].includes(arguments[1])) {
                window.jsonLoaded = true;
            }
            open.apply(this, arguments);
        };
    """)

    # Wait for the network activity to confirm the JSON load
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return window.jsonLoaded;")
    )
def login(browser):
    browser.get('https://auth.narrpr.com/auth/sign-in')

    time.sleep(2)
    WebDriverWait(browser, 25).until(lambda b: b.execute_script("return document.readyState") == "complete")
    if (element_exists_xpath(browser, '//*[@id="mat-mdc-dialog-0"]/div/div')):
        browser.get('https://auth.narrpr.com/auth/sign-in')
        time.sleep(random.uniform(3, 5))
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


def intercept_excep(element):
    try:
        element.click()
        return True
    except ElementClickInterceptedException:
        return False