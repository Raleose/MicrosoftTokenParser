import json

from selenium.common.exceptions import TimeoutException
from selenium.webdriver import Proxy
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.wait import WebDriverWait

max_wait_time = 10


def get_token_for_account(login, password, proxy=None):
    capabilities = DesiredCapabilities.CHROME
    capabilities["goog:loggingPrefs"] = {"performance": "ALL"}

    options = Options()
    options.add_argument('--disable-extensions')
    options.add_argument('test-type')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('no-sandbox')
    options.add_argument('--headless')
    if proxy:
        options.proxy = Proxy(proxy)

    service = Service(executable_path='C:\\chromedriver\\chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options, desired_capabilities=capabilities)

    try:
        def wait_element_to_be_present(by, value):
            return WebDriverWait(driver, max_wait_time).until(EC.presence_of_element_located((by, value)))

        def wait_element_to_be_clickable(by, value):
            return WebDriverWait(driver, max_wait_time).until(EC.element_to_be_clickable((by, value)))

        driver.get("https:\\redeem.microsoft.com")

        try:
            # ввод логина
            login_field = wait_element_to_be_present(By.NAME, 'loginfmt')
            login_field.send_keys(login)

            ok_button = wait_element_to_be_clickable(By.ID, 'idSIButton9')
            ok_button.click()
        except TimeoutException:
            print('неверный логин')
            return None

        print('логин введен')

        try:
            # ввод пароля
            password_field = wait_element_to_be_present(By.NAME, 'passwd')
            password_field.send_keys(password)

            login_button = wait_element_to_be_clickable(By.ID, 'idSIButton9')
            login_button.click()
        except TimeoutException:
            print('неверный пароль')
            return None

        print('пароль введен')
        print(f'ждем результат проверки учетки (max = {max_wait_time} сек)...')

        try:
            # это вопрос оставаться ли в системе. он должен быть, но если его нет - пофиг
            login_button = wait_element_to_be_clickable(By.ID, 'idSIButton9')
            login_button.click()
        except:
            pass

        try:
            wait_element_to_be_present(By.ID, 'blendContainer')
        except TimeoutException:
            print('введены неверные учетные данные')
            return None

        print('учетные данные верны, парсим токен...')

        browser_log = driver.get_log("performance")

        def process_browser_log_entry(entry):
            response = json.loads(entry['message'])['message']
            return response

        events = [process_browser_log_entry(entry) for entry in browser_log]
        events = [event for event in events if 'Network.response' in event['method']]

        for event in events:
            if 'redeem.microsoft.com/webblendredeem' in str(event):
                request_id = event['params']['requestId']
                dict = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id})
                token = json.loads(dict['body'])['metadata']['mscomct']
                print("токен получен!")
                return token

        # если по какой-то причине не получили токен, отправляем None
        return None

    except Exception as e:
        print("Ошибка:")
        print(e)
    finally:
        driver.close()
        driver.quit()


if __name__ == '__main__':
    login = ''
    password = ''
    token = get_token_for_account(login, password)
    print(token)
