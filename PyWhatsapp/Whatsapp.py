import os
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time
import openpyxl as excel
import urllib.parse

try:
    import autoit
except ModuleNotFoundError:
    pass
browser = None
Contact = None
message = None
Link = "https://web.whatsapp.com/"
wait = None
choice = None

unsaved_Contacts = None


def input_contacts():
    global Contact, unsaved_Contacts

    Contact = []
    unsaved_Contacts = []
    option = True
    while option:
        print("PLEASE CHOOSE ONE OF THE OPTIONS:\n")
        print("1.Message to Saved Contact number")
        print("2.Message to Unsaved Contact number\n")

        x = input("Enter your choice(1 or 2):")
        if x == '1':
            option = False
            try:
                n = int(input('Enter number of Contacts to add(count)->'))
            except:
                print("Wrong Input")
                n = int(input('Enter number of Contacts to add(A number)->'))
            for i in range(0, n):
                inp = str(input("Enter contact name(text)->"))
            inp = '"' + inp + '"'
            Contact.append(inp)
        if x == '2':
            option = False
            n = int(input('Enter number of unsaved Contacts to add(count)->'))
            print()
            for i in range(0, n):
                inp = str(input(
                    "Enter unsaved contact number with country:\n\nValid input: 92313xxxxx59\nInvalid input: "
                    "+92313xxxxx59\n\n"))

                unsaved_Contacts.append(inp)
            continue

    if len(Contact) != 0:
        print("\nSaved contacts entered list->", Contact)
    if len(unsaved_Contacts) != 0:
        print("Unsaved numbers entered list->", unsaved_Contacts)
    input("\nPress ENTER to continue...")


def input_message():
    global message
    print(
        "Enter the message and use the symbol '-' to end the message:\nFor example: Hi, this is a test message-\n\nYour message: ")
    message = []
    temp = ""
    done = False

    while not done:
        temp = input()
        if len(temp) != 0 and temp[-1] == "-":
            done = True
            message.append(temp[:-1])
        else:
            message.append(temp)
    message = "\n".join(message)
    print()
    print(message)


def whatsapp_login():
    global wait, browser, Link
    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    browser = webdriver.Chrome("/Users/Jawad/Desktop/whatsapp/chromedriver.exe")
    wait = WebDriverWait(browser, 600)
    browser.get(Link)
    browser.maximize_window()
    print("QR scanned")


def send_message(target):
    global message, wait, browser
    try:
        x_arg = '//span[contains(@title,' + target + ')]'
        ct = 0
        while ct != 10:
            try:
                group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
                group_title.click()
                break
            except:
                ct += 1
                time.sleep(3)
        input_box = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        for ch in message:
            if ch == "\n":
                ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(
                    Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            else:
                input_box.send_keys(ch)
        input_box.send_keys(Keys.ENTER)
        print("Message sent successfuly")
        time.sleep(1)
    except NoSuchElementException:
        return


def send_unsaved_contact_message():
    global message
    try:
        time.sleep(7)
        input_box = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        for ch in message:
            if ch == "\n":
                ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(
                    Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            else:
                input_box.send_keys(ch)
        input_box.send_keys(Keys.ENTER)
        print("Message sent successfuly")
    except NoSuchElementException:
        print("Failed to send message")
        return


def send_files():
    global doc_filename
    # Attachment Drop Down Menu
    clipButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div/span')
    clipButton.click()
    time.sleep(1)

    # To send a Document(PDF, Word file, PPT)
    docButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[3]/button')
    docButton.click()
    time.sleep(1)

    docPath = os.getcwd() + "\\Documents\\" + doc_filename

    autoit.control_focus("Open", "Edit1")
    autoit.control_set_text("Open", "Edit1", (docPath))
    autoit.control_click("Open", "Button1")

    time.sleep(3)
    whatsapp_send_button = browser.find_element_by_xpath(
        '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div')
    whatsapp_send_button.click()


def sender():
    global Contact, choice, docChoice, unsaved_Contacts
    for i in Contact:
        send_message(i)
        print("Message sent to ", i)
        if docChoice == "yes":
            try:
                send_files()
                print("File sent")
            except:
                print('Files not sent')
    time.sleep(5)

    if len(unsaved_Contacts) > 0:
        for i in unsaved_Contacts:
            link = "https://wa.me/" + i
            browser.get(link)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="action-button"]').click()
            time.sleep(2)
            browser.find_element_by_xpath('//*[@id="fallback_block"]/div/div/a').click()
            time.sleep(4)
            print("Sending message to", i)
            send_unsaved_contact_message()
            if docChoice == "yes":
                try:
                    send_files()
                    print("File sent")
                except:
                    print('Files not sent')
            time.sleep(7)


def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    secondCol = sheet['B']
    driver = webdriver.Chrome("/Users/Jawad/Desktop/whatsapp/chromedriver.exe")
    driver.get('https://web.whatsapp.com')
    time.sleep(60)

    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        message = str(secondCol[cell].value)
        print(contact)
        print("Your message: '" + message + "'")
        link = "https://web.whatsapp.com/send?phone=" + contact + "&amp;text=" + urllib.parse.quote_plus(
            message) + "&amp;source=&amp;data="
        # print(link)
        driver.get(link)
        time.sleep(4)
        print("Sending message to", contact)

        try:
            time.sleep(7)
            input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
            for ch in message:
                if ch == "\n":
                    ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(
                        Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
                else:
                    input_box.send_keys(ch)
            input_box.send_keys(Keys.ENTER)
            print("Message sent successfuly")
            time.sleep(5)
            send_files()
            time.sleep(5)
        except:
            print("User is not on Whatsapp")

    driver.quit()


# main program
action = input("What do you want to perform? \n"
               "Press 1: Message numbers of Excel Sheet\n"
               "Press 2: Message Other way\n")
if action == "1":
    readContacts("./contacts-message.xlsx")
if action == "2":
    input_contacts()
    input_message()
    docChoice = input("Would you file to send Attachment(yes/no): ")
    if docChoice == "yes":
        print(
            'Note the document file should be present in the Document Folder\nAdd document extension with the name i.e jawad.jpg')
        doc_filename = input("Enter the Document file name you want to send: ")
    print("SCAN YOUR QR CODE FOR WHATSAPP WEB")
    whatsapp_login()
    sender()
