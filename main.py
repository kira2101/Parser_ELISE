
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from XLS import *
import time
from alive_progress import alive_bar
import socket
from selenium.common.exceptions import WebDriverException
from tqdm import tqdm

#try:
options = webdriver.FirefoxOptions()
print('Статус: WebDriver запущен...')

flag_opt =input('Запустить браузер в фоновом режиме?  Y/N?:')

match flag_opt:
    case 'Y'|'y':
        options.headless = True
        print("Статус: запуск FireFox в фоновом режиме")
        browser = webdriver.Firefox(options=options)
    case 'N'|'n':
        options.headless = False
        browser = webdriver.Firefox(options=options)
        print("Статус: запуск FireFox в оконном режиме режиме")
    case _:
        options.headless = True
        print("Статус: запуск FireFox по умолчанию в фоновом режиме")
        browser = webdriver.Firefox(options=options)

flag_log =input('Включить режим логирования?  Y/N?:')
match flag_log:
    case 'Y'|'y':
        log_on_flag = True
        print("Режим логирования включен")

    case 'N'|'n':
        log_on_flag = False
        print("Режим логирования выключен")
    case _:
        log_on_flag = True
        print("Режим логирования включен по умолчанию")

#options.headless = True
#browser = webdriver.Firefox(options=options)


def LoginMemberZone(password, email):
    try:
        password = password
        email = email
        browser.get("http://elsie.ua/ukr/login.html")
        input_email_XPATH = '//*[@id="input_login_inpt"]'
        input_pass_XPATH = '//*[@id="input_passwd_inpt"]'
        btn_login_XPATH = '/html/body/div[1]/div[8]/div/form/div[3]/input'
        browser.find_element(By.XPATH, input_email_XPATH).send_keys(email)
        browser.find_element(By.XPATH, input_pass_XPATH).send_keys(password)
        browser.find_element(By.XPATH, btn_login_XPATH).click()
        print("Статус: вход выполнен успешно...")
        return True
    except:
        return False

def InsertELCICode(code):
    try:
        input_code_XPATH = '/html/body/div/div[8]/form[1]/table/tbody/tr[2]/td[1]/input'
        browser.find_element(By.XPATH, input_code_XPATH).clear()
        browser.find_element(By.XPATH, input_code_XPATH).send_keys(code)
        browser.find_element(By.XPATH, input_code_XPATH).send_keys(Keys.ENTER)
    except:
        time.sleep(1)
        input_code_XPATH = '/html/body/div/div[8]/form[1]/table/tbody/tr[2]/td[1]/input'
        browser.find_element(By.XPATH, input_code_XPATH).clear()
        browser.find_element(By.XPATH, input_code_XPATH).send_keys(code)
        browser.find_element(By.XPATH, input_code_XPATH).send_keys(Keys.ENTER)



def CheckInternetStatus():
    try:
        socket.gethostbyaddr('www.google.com')
    except socket.gaierror:
        return False
    return True


def StatusOfCode(code):
    flag = True
    while flag == True:
        try:
            time.sleep(1)
            msgXPATH = '/html/body/div/div[8]/form[2]/table/tbody/tr/td'  #
            msg = browser.find_element(By.XPATH, msgXPATH).text
            flag = False

        except:


            InsertELCICode(code)
            time.sleep(1)
            flag = True

    if flag == False:
        return msg
    else:
        pass

def GrabPrice(code, row_to_exel ):
    flag = True
    while flag == True:
        try:
            time.sleep(1)
            price_text_XPATH = '/html/body/div/div[8]/form[2]/table/tbody/tr[1]/td[4]'
            price = browser.find_element(By.XPATH, price_text_XPATH)

            print('grab ok')
            flag = False
        except:

            InsertELCICode(code)
            print('Grab is fail!')

            flag = True

    if flag == False:

        row_to_exel.append(price.text)
        SaveRowToExcel(row_to_exel, excel_file_name)
    else:
        pass





path = input('Ввудите путь к каталогу: ')
LoginMemberZone('034e0637','Duhin.av@gmail.com')
exel_file = GetExcel(path)
elci_code = exel_file['Код Элси'].tolist()

print('Найдено ЕЛСИ кодов: ' + str(len(elci_code)))
row_to_exel =[]

excel_file_name =  CreatExelFile()
i=0
if elci_code:


    for code in tqdm(elci_code):

        msg = ""
        row_to_exel = exel_file.iloc[i].tolist()


        InsertELCICode(code)



        msg = StatusOfCode(code)

        match msg:
            case "Не знайдено!":

                row_to_exel.append('NONE')
                SaveRowToExcel(row_to_exel, excel_file_name)
                text_msg = f"Товар по коду {code} не найден"


            case 'NULL':
                row_to_exel.append('ERROR')
                SaveRowToExcel(row_to_exel, excel_file_name)
                print('Статус ошибка чтения')
                print(row_to_exel)
            case _:
                GrabPrice(code, row_to_exel)
                text_msg = f'По коду {code} найден товар по цене {row_to_exel[-1]}'




        i=i+1
        if log_on_flag:
            print(f'\r{text_msg}')
        else:
            pass
    else:
        print("Error: No Elci code")



#except WebDriverException:
    #print('WebDriver остановлен...')
    #print('https://github.com/mozilla/geckodriver/releases')









