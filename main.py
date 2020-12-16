from selenium import webdriver
import openpyxl
import user_file
import time
from bs4 import BeautifulSoup

file = 'TA - RPA Challenge Shopping List.xlsx'
updated_file = 'RPA-Updatedfile.xlsx'
excel_file = openpyxl.load_workbook(file)

def setup_driver():
    global driver
    AMAZON_URL = 'https://www.amazon.com/'
    driver = webdriver.Chrome('chromedriver.exe')
    driver.get(AMAZON_URL)
    driver.maximize_window()

def get_elements():
    item1 = user_file.readdatafile(file,"Sheet1",2,1)
    iroobt = driver.find_element_by_id('twotabsearchtextbox')
    iroobt.send_keys(item1)
    driver.find_element_by_xpath('//*[@id="nav-search-submit-text"]/input').click()
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    prices = soup.find_all('span',{'class' : 'a-price-whole'})
    href_links = soup.find_all('a',{'class':'a-link-normal a-text-normal'})
    item1_hreflink = []
    for link in href_links:
        item1_hreflink.append(link.get('href'))
    item1_price = [price.get_text() for price in prices]
    for i in range(0, len(item1_price)):
        item1_price[i]= item1_price[i].replace('.', '')
        item1_price[i] = int(item1_price[i])
    item1_cheap_price_href = 'https://www.amazon.com/'+str(min(list(zip(item1_price,item1_hreflink)))[1])
    item1_cheap_price= '$'+str(min(list(zip(item1_price,item1_hreflink)))[0])
    user_file.writedatafile(item1_cheap_price,updated_file,'Sheet1',2,2)
    user_file.writedatafile(item1_cheap_price_href,updated_file,'Sheet1',2,3)
    driver.find_element_by_id('nav-logo-sprites').click()
    item2 = user_file.readdatafile(file, "Sheet1", 3, 1)
    iroobt = driver.find_element_by_id('twotabsearchtextbox')
    iroobt.send_keys(item2)
    driver.find_element_by_xpath('//*[@id="nav-search-submit-text"]/input').click()
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    prices = soup.find_all('span', {'class': 'a-price-whole'})
    href_links = soup.find_all('a', {'class': 'a-link-normal a-text-normal'})
    item2_hreflink = []
    for link in href_links:
        item2_hreflink.append(link.get('href'))
    item2_price = [price.get_text() for price in prices]
    for i in range(0, len(item2_price)):
        item2_price[i] = item2_price[i].replace('.', '')
        item2_price[i] = int(item2_price[i])
    item2_cheap_price_href = 'https://www.amazon.com/' + str(min(list(zip(item2_price, item2_hreflink)))[1])
    item2_cheap_price = '$' + str(min(list(zip(item2_price, item2_hreflink)))[0])
    user_file.writedatafile(item2_cheap_price, updated_file, 'Sheet1', 3, 2)
    user_file.writedatafile(item2_cheap_price_href, updated_file, 'Sheet1', 3, 3)
    driver.find_element_by_id('nav-logo-sprites').click()
    item3 = user_file.readdatafile(file, "Sheet1", 4, 1)
    iroobt = driver.find_element_by_id('twotabsearchtextbox')
    iroobt.send_keys(item3)
    driver.find_element_by_xpath('//*[@id="nav-search-submit-text"]/input').click()
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    prices = soup.find_all('span', {'class': 'a-price-whole'})
    href_links = soup.find_all('a', {'class': 'a-link-normal a-text-normal'})
    item3_hreflink = []
    for link in href_links:
        item3_hreflink.append(link.get('href'))
    item3_price = [price.get_text() for price in prices]
    for i in range(0, len(item3_price)):
        item3_price[i] = item3_price[i].replace('.', '')
        item3_price[i] = int(item3_price[i])
    item3_cheap_price_href = 'https://www.amazon.com/' + str(min(list(zip(item3_price, item3_hreflink)))[1])
    item3_cheap_price = '$' + str(min(list(zip(item3_price, item3_hreflink)))[0])
    user_file.writedatafile(item3_cheap_price, updated_file, 'Sheet1', 4, 2)
    user_file.writedatafile(item3_cheap_price_href, updated_file, 'Sheet1', 4, 3)
    driver.find_element_by_id('nav-logo-sprites').click()
    item4 = user_file.readdatafile(file, "Sheet1", 5, 1)
    iroobt = driver.find_element_by_id('twotabsearchtextbox')
    iroobt.send_keys(item4)
    driver.find_element_by_xpath('//*[@id="nav-search-submit-text"]/input').click()
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    prices = soup.find_all('span', {'class': 'a-price-whole'})
    href_links = soup.find_all('a', {'class': 'a-link-normal a-text-normal'})
    item4_hreflink = []
    for link in href_links:
        item4_hreflink.append(link.get('href'))
    item4_price = [price.get_text() for price in prices]
    for i in range(0, len(item4_price)):
        item4_price[i] = item4_price[i].replace('.', '')
        item4_price[i] = int(item4_price[i])
    item4_cheap_price_href = 'https://www.amazon.com/' + str(min(list(zip(item4_price, item4_hreflink)))[1])
    item4_cheap_price = '$' + str(min(list(zip(item4_price, item4_hreflink)))[0])
    user_file.writedatafile(item4_cheap_price, updated_file, 'Sheet1', 5, 2)
    user_file.writedatafile(item4_cheap_price_href, updated_file, 'Sheet1', 5, 3)
    driver.find_element_by_id('nav-logo-sprites').click()
    item5 = user_file.readdatafile(file, "Sheet1", 6, 1)
    iroobt = driver.find_element_by_id('twotabsearchtextbox')
    iroobt.send_keys(item5)
    driver.find_element_by_xpath('//*[@id="nav-search-submit-text"]/input').click()
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    prices = soup.find_all('span', {'class': 'a-price-whole'})
    href_links = soup.find_all('a', {'class': 'a-link-normal a-text-normal'})
    item5_hreflink = []
    for link in href_links:
        item5_hreflink.append(link.get('href'))
    item5_price = [price.get_text() for price in prices]
    for i in range(0, len(item5_price)):
        item5_price[i] = item5_price[i].replace('.', '')
        item5_price[i] = int(item5_price[i])
    item5_cheap_price_href = 'https://www.amazon.com/' + str(min(list(zip(item5_price, item5_hreflink)))[1])
    item5_cheap_price = '$' + str(min(list(zip(item5_price, item5_hreflink)))[0])
    user_file.writedatafile(item5_cheap_price, updated_file, 'Sheet1', 6, 2)
    user_file.writedatafile(item5_cheap_price_href, updated_file, 'Sheet1', 6, 3)


def tear_down():
    driver.close()
    

if __name__=='__main__':
    setup_driver()
    get_elements()
    tear_down()