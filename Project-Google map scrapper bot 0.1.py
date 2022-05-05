from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook
import time


# searcher function search different elements form the map results and extract them
def searcher(list1, list2, list3, list4, list5, list6, list7, list8):
    title = driver.find_element(By.XPATH, '//div[@class="lMbq3e"]/div[1]')
    try:
        title = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="lMbq3e"]/div[1]')))
        # print(title.text)  #only to check the output is correct or not
        driver.execute_script('arguments[0].scrollIntoView(true);', title)
        list1.append(title.text)
    except Exception:
        # print('No name')
        time.sleep(2)
        list1.append('No name')
    try:

        ratings = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="jANrlb"]/div[1]')))
        driver.execute_script('arguments[0].scrollIntoView(true);', ratings)
        # print(ratings.text)
        list2.append(ratings.text)
    except Exception:
        # print('No ratings')
        list2.append('No ratings')
    try:
        reviews = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="jANrlb"]/button')))
        driver.execute_script('arguments[0].scrollIntoView(true);', reviews)
        # print(reviews.text)
        list3.append(reviews.text)
    except Exception:
        # print('No reviews')
        time.sleep(2)
        list3.append('No reviews')
    try:
        typ = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="tAiQdd"]//span[1]/span[1]/button')))
        # print(typ.text)
        driver.execute_script('arguments[0].scrollIntoView(true);', typ)
        list4.append(typ.text)
    except Exception:
        # print('Unknown')
        list4.append('Unknown')
    try:
        locat = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//body/div[@id='app-container']/div[@id='content-container']/div[@id='pane']/div[1]/div[1]/div[1]/div[1]/div[7]//button[@data-item-id='address']/div[1]")))
        # print(locat.text)
        driver.execute_script('arguments[0].scrollIntoView(true);', locat)
        list5.append(locat.text)
    except Exception:
        # print('Unknown')
        list5.append('Unknown')
    try:
        opening = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//body/div[@id='app-container']/div[@id='content-container']/div[@id='pane']/div[1]/div[1]/div[1]/div[1]/div[7]//div[@role='button']/div[1]/div[1]/span[1]")))
        # print(opening.text)
        driver.execute_script('arguments[0].scrollIntoView(true);', opening)
        list6.append(opening.text)
    except Exception:
        # print("Not scheduled")
        list6.append("Not scheduled")
    try:
        web = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//body/div[@id='app-container']/div[@id='content-container']/div[@id='pane']/div[1]/div[1]/div[1]/div[1]/div[7]//button[@data-tooltip='Open website']/div[1]/div[2]")))
        driver.execute_script('arguments[0].scrollIntoView(true);', web)  # scrolling used for the dynamic location directly
        # print(web.text)
        list7.append(web.text)
    except Exception:
        # print("No website")
        list7.append('No website')
    try:
        phn = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//body/div[@id='app-container']/div[@id='content-container']/div[@id='pane']/div[1]/div[1]/div[1]/div[1]/div[7]//button[@data-tooltip='Copy phone number']/div[1]/div[2]")))
        driver.execute_script('arguments[0].scrollIntoView(true);', phn)
        # print(phn.text)
        list8.append(phn.text)
    except Exception:
        # print('No contact')
        list8.append('No contact')


# only for creating a new file name Details in desktop if not exist to contain the excel file
def folder():
    fold = r'C:\Users\Tuhin\Desktop\Details'
    if not os.path.exists(fold):
        os.makedirs(fold)


# scroll the main page of map and take it to the last search result
def scroll(inp):
    for scroll_item in range(5):
        inp.send_keys(Keys.END)
        time.sleep(10)

# sheet function used to contain the extracted value to a excel file
def sheet(inp1, inp2, inp3, inp4, inp5, inp6, inp7, inp8):
    details = {
        'Name': inp1,
        'Ratings': inp2,
        'Review': inp3,
        'Type': inp4,
        'Location': inp5,
        'Open-close': inp6,
        'Website': inp7,
        'Contact': inp8,
    }
    folder()
    df = pd.DataFrame(details)
    writer = pd.ExcelWriter('Details.xlsx', engine='xlsxwriter')
    writer.save()
    df.to_excel(r'C:\Users\Tuhin\Desktop\Details\Details.xlsx', index=False)
    df.to_excel(writer, sheet_name='Details.xlsx', index=False)


def excel_designer():
    os.chdir(r'C:\Users\Tuhin\Desktop\Details')  #set the folfer path
    wb = load_workbook('Details.xlsx')  # for fomatting the excel sheet
    # worksheet1 = wb['Sheet1']   #for specific sheet open
    sheet1 = wb.active
    sheet1.column_dimensions['A'].width = 30  # name of the rows or column for formatting
    sheet1.column_dimensions['B'].width = 15
    sheet1.column_dimensions['C'].width = 15
    sheet1.column_dimensions['D'].width = 15
    sheet1.column_dimensions['E'].width = 90
    sheet1.column_dimensions['F'].width = 25
    sheet1.column_dimensions['G'].width = 25
    sheet1.column_dimensions['H'].width = 25
    wb.save('Details.xlsx')


s = input("Enter your search string in map:")
print("Enter your chromedriver path:....>")  #This is my path you can use your own where the chrome driver you saved or located ....> C:\Users\Tuhin\Desktop\chromedriver_win32
os.environ['PATH'] = input(r"")   #this is the path where the chorome driver exists
driver = webdriver.Chrome()
driver.get("https://www.google.com/maps/")

search1 = driver.find_element(By.ID, 'searchboxinput')     #find the map search-box and search your query
time.sleep(5)
search1.send_keys(s)
time.sleep(5)
driver.find_element(By.XPATH, value='//button[@aria-label="Search"]').click()
time.sleep(10)
elem = driver.find_element(By.XPATH, value=f'//div[@aria-label="Results for {s}"]')
href_list = []

#extracting the website values to access different search results elements
while True:
    scroll(elem)
    time.sleep(5)
    search_results = WebDriverWait(driver, 60).until(
        EC.visibility_of_all_elements_located((By.XPATH, f'//div[@aria-label="Results for {s}"]//a')))
    for item in search_results:
        a = item.get_attribute('href')
        href_list.append(a)
        time.sleep(5)
    button_ = driver.find_element(By.ID, 'ppdPk-Ej1Yeb-LgbsSe-tJiF1e')
    if button_.is_enabled() == False:
        break
    else:
        button_.click()
driver.quit()
print(href_list)
Name = []
Ratings = []
Reviews = []
Type = []
Location = []
Open_close = []
Website = []
Contact = []
for item in href_list:   #access each search results elements one by one
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    driver.get(item)
    # driver.maximize_window()
    searcher(Name, Ratings, Reviews, Type, Location, Open_close, Website, Contact)
    WebDriverWait(driver, 15).until(EC.url_changes(item))
    # time.sleep(3)
    # driver.back()
    sheet(Name, Ratings, Reviews, Type, Location, Open_close, Website, Contact)
driver.close()
excel_designer()

