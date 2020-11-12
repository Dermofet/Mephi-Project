import requests
import selenium
import xlwt
from selenium import webdriver
from xlrd import open_workbook
from xlutils import copy


URL = 'https://maps.googleapis.com/maps/api/place/textsearch/json'
key = 'AIzaSyA5xIulXbroyFR7-O8vChDJYeMYRCML8jA'
language = 'ru'


def find_places(response_json):
    dict_address = []
    dict_name = []
    dict_place_id = []
    dict_rating = []
    dict_user_ratings_total = []
    data = {'dict_name': dict_name, 'dict_address': dict_address, 'dict_place_id': dict_place_id, 'dict_rating': dict_rating, 'dict_user_ratings_total': dict_user_ratings_total}
    # 55.677972, 37.664917
    # 55.654067, 37.648722
    for i in response_json['results']:
        _address_ = i['formatted_address']
        print('_address_ = ', _address_)
        dict_address.append(_address_)
    for i in response_json['results']:
        _name_ = i['name']
        print('_name_ = ', _name_)
        dict_name.append(_name_)
    for i in response_json['results']:
        _place_id_ = i['place_id']
        print('_place_id_ = ', _place_id_)
        dict_place_id.append(_place_id_)
    for i in response_json['results']:
        _rating_ = i['rating']
        print('_rating_', _rating_)
        dict_rating.append(_rating_)
    for i in response_json['results']:
        _user_ratings_total_ = i['user_ratings_total']
        print('_user_ratings_total_', _user_ratings_total_)
        dict_user_ratings_total.append(_user_ratings_total_)
    print('data = ', data)
    print('dict_name = ', dict_name)
    print('data[dict_name] = ', data['dict_name'])
    return data


def make_file_excel():
    wb = xlwt.Workbook()
    wb.add_sheet('Технический лист')
    wb.save('ПП.xls')


def open_excel_file(data, inn, sheet_name):
    wb = open_workbook('ПП.xls')
    wbcpy =copy.copy(wb)
    sheet = wbcpy.add_sheet(sheet_name)
    sheet.write(0, 0, label='Name')
    sheet.write(0, 1, label='Address')
    sheet.write(0, 2, label='Place ID')
    sheet.write(0, 3, label='Rating')
    sheet.write(0, 4, label='User Ratings Total')
    sheet.write(0, 5, label='INN')
    for t, i in enumerate(data['dict_name']):
        sheet.write(t + 1, 0, label=i)
    for t, i in enumerate(data['dict_address']):
        sheet.write(t + 1, 1, label=i)
    for t, i in enumerate(data['dict_place_id']):
        sheet.write(t + 1, 2, label=i)
    for t, i in enumerate(data['dict_rating']):
        sheet.write(t + 1, 3, label=i)
    for t, i in enumerate(data['dict_user_ratings_total']):
        sheet.write(t + 1, 4, label=i)
    for t, i in enumerate(inn):
        sheet.write(t + 1, 5, label=i)
    wbcpy.save('ПП.xls')


def write_data_txt(data, inn):
    file = open('ПП.txt', 'w', encoding='utf-8')
    for i in data['dict_name']:
        file.write("%s\n" % i)
    file.write("\n")
    for i in data['dict_address']:
        file.write("%s\n" % i)
    file.write("\n")
    for i in data['dict_place_id']:
        file.write("%s\n" % i)
    file.write("\n")
    for i in data['dict_rating']:
        file.write("%s\n" % i)
    file.write("\n")
    for i in data['dict_user_ratings_total']:
        file.write("%s\n" % i)
    for i in inn:
        file.write("%s\n" % i)
    file.close()


def parsing_GM(location, radius, keyword, sheet_name):
    payload = {'query': keyword, 'location': location, 'radius': radius, 'language': language, 'key': key}
    response = requests.get(URL, params=payload)
    response_json = response.json()
    data = find_places(response_json)
    chromedriver = 'C:\\Users\\Гусюся\\chromedriver\\chromedriver.exe'
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    browser = webdriver.Chrome(executable_path=chromedriver, options=options)
    inn = []
    browser.implicitly_wait(5)
    browser.get('https://bo.nalog.ru/search?query=%D0%9E%D0%9E%D0%9E+%22%D0%93%D0%90%D0%97%D0%9F%D0%A0%D0%9E%D0%9C%22&page=1')
    browser.implicitly_wait(10)
    for i, t in zip(data['dict_name'], data['dict_address']):
        # button_clear = browser.find_element_by_class_name('button button_md button_secondary')
        # browser.execute_script("arguments[0].click();", button_clear)
        # button = browser.find_element_by_class_name('button extended-search__button button_none')
        # browser.execute_script("arguments[0].click();", button)
        browser.implicitly_wait(10)
        print(browser.find_element_by_id('name'))
        browser.find_element_by_id('name').clear()
        browser.find_element_by_id('name').send_keys(i)
        browser.find_element_by_id('address').clear()
        browser.find_element_by_id('address').send_keys(t)
        browser.find_element_by_id('allFieldsMatch').click()
        browser.find_element_by_class_name('button button_md button_primary').click()
        browser.implicitly_wait(15)
        _inn_ = browser.find_elements_by_class_name('results-search-table-item').text()
        # _inn_ = browser.find_elements_by_xpath('//div[@class="results-research-tbody"]/a[0]/div[1]/div[0]/div[1]/text()')
        print(_inn_)
        inn.append(_inn_)
    browser.close()
    print('inn = ', inn)
    open_excel_file(data, inn, sheet_name)
    write_data_txt(data, inn)


def main():
    print('Введите через запятую координаты точки, относительно которой будет произведен поиск')
    location = input()
    print('Введите радиус поиска')
    radius = input()
    print('Введите ключевое слово')
    keyword = input()
    print('Напишите название листа Excel файла, в котором будет сохранены данные')
    sheet_name = input()
    make_file_excel()
    parsing_GM(location, radius, keyword, sheet_name)
    while True:
        print('Хотите продолжить? Yes/No')
        answer = input()
        if answer == 'Yes':
            answer = None
            print('Хотите изменить только ключевое слово? Yes/No')
            answer = input()
            if answer == 'Yes':
                keyword = None
                print('Введите ключевое слово')
                keyword = input()
                print('Напишите название листа Excel файла, в котором будет сохранены данные')
                sheet_name = input()
                parsing_GM(location, radius, keyword, sheet_name)
            else:
                answer = None
                print('Если хотите изменить значений всех переменных - напишите "Yes" (без кавычек). '
                      'Если хотите изменить только координаты точки - напишите "location". '
                      'Если хотите изменить радиус - напишите "radius".')
                answer = input()
                if answer == 'Yes':
                    location = None
                    radius = None
                    keyword = None
                    print('Введите через запятую координаты точки, относительно которой будет произведен поиск')
                    location = input()
                    print('Введите радиус поиска')
                    radius = input()
                    print('Введите ключевое слово')
                    keyword = input()
                    print('Напишите название листа Excel файла, в котором будет сохранены данные')
                    sheet_name = input()
                    parsing_GM(location, radius, keyword, sheet_name)
                if answer == 'location':
                    location = None
                    print('Введите через запятую координаты точки, относительно которой будет произведен поиск')
                    location = input()
                    print('Напишите название листа Excel файла, в котором будет сохранены данные')
                    sheet_name = input()
                    parsing_GM(location, radius, keyword, sheet_name)
                if answer == 'radius':
                    radius = None
                    print('Введите радиус поиска')
                    radius = input()
                    print('Напишите название листа Excel файла, в котором будет сохранены данные')
                    sheet_name = input()
                    parsing_GM(location, radius, keyword, sheet_name)
        else:
            break
    print('Программа закончила работу')


if __name__ == '__main__':
    main()
