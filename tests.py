from selenium import webdriver


def main():
    chromedriver = 'C:\\Users\\Гусюся\\chromedriver\\chromedriver.exe'
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    browser = webdriver.Chrome(executable_path=chromedriver, options=options)
    browser.implicitly_wait(5)
    browser.get('https://bo.nalog.ru/')
    browser.implicitly_wait(5)
    button = browser.find_element_by_xpath('//main[@class="search-page"]/div/button')
    browser.execute_script("arguments[0].click();", button)
    browser.implicitly_wait(2)
    browser.find_element_by_xpath('//form[@class="extended-search-form"]/form/div[0]/div[0]/div[0]/input').send_keys('ООО "СБЕРБАНК"')
    browser.find_element_by_xpath('//form[@class="extended-search-form"]/form/div[0]/div[0]/div[1]/input').send_keys('Коломенская ул., 23, корп. 2')
    browser.implicitly_wait(5)
    button_find = browser.find_element_by_xpath('//form[@class="extended-form"]/div[2]/button[0]')
    browser.execute_script("arguments[0].click();", button_find)
    # browser.find_element_by_xpath('//button[@class="button button_search"]')
    # button = browser.find_element_by_xpath('//div[@class="header-wrapper"]/div[2]/button')
    # browser.execute_script("arguments[0].click();", button)
    browser.implicitly_wait(5)
    # browser.find_element_by_xpath('//form[@class="extended-form"]/div[0]/div[0]/div[0]/input').send_keys('ООО "СБЕРБАНК"')
    # button_clear = browser.find_element_by_xpath('//form[@class="extended-form"]/div[2]/button[1]')
    # browser.execute_script("arguments[0].click();", button_clear)
    # browser.implicitly_wait(5)


if __name__ == '__main__':
    main()
