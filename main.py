import time
import requests
import xlwt
from selenium import webdriver
import re


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


def open_excel_file(data, commercial_inf, name_excel_file):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet('GM_Parser')
    sheet.write(0, 0, label='Название')
    sheet.write(0, 1, label='Адрес')
    sheet.write(0, 2, label='Place ID')
    sheet.write(0, 3, label='Рейтинг')
    sheet.write(0, 4, label='User Ratings Total')
    sheet.write(0, 5, label='ИНН')
    sheet.write(0, 6, label='КПП')
    sheet.write(0, 7, label='ОГРН')
    sheet.write(0, 8, label='Уставной капитал')
    sheet.write(0, 9, label='Активы')
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
    for t, i in enumerate(commercial_inf['inn']):
        sheet.write(t + 1, 5, label=i)
    for t, i in enumerate(commercial_inf['kpp']):
        sheet.write(t + 1, 6, label=i)
    for t, i in enumerate(commercial_inf['ogrn']):
        sheet.write(t + 1, 7, label=i)
    for t, i in enumerate(commercial_inf['capital']):
        sheet.write(t + 1, 8, label=i)
    for t, i in enumerate(commercial_inf['assets']):
        sheet.write(t + 1, 9, label=i)
    wb.save(name_excel_file)


# def write_data_txt(data, inn):
#     file = open('ПП.txt', 'w', encoding='utf-8')
#     for i in data['dict_name']:
#         file.write("%s\n" % i)
#     file.write("\n")
#     for i in data['dict_address']:
#         file.write("%s\n" % i)
#     file.write("\n")
#     for i in data['dict_place_id']:
#         file.write("%s\n" % i)
#     file.write("\n")
#     for i in data['dict_rating']:
#         file.write("%s\n" % i)
#     file.write("\n")
#     for i in data['dict_user_ratings_total']:
#         file.write("%s\n" % i)
#     for i in inn:
#         file.write("%s\n" % i)
#     file.close()


def check(name):
    en_let = 0
    if re.search(r'[a-z]', name):
        en_let = en_let + 1
    print('en_let = ', en_let)
    return en_let


def parsing_GM(location, radius, keyword, name_excel_file):
    payload = {'query': keyword, 'location': location, 'radius': radius, 'language': language, 'key': key}
    response = requests.get(URL, params=payload)
    response_json = response.json()
    data = find_places(response_json)
    driver = webdriver.Chrome()
    inn = []
    kpp = []
    ogrn = []
    capital = []
    assets = []
    driver.get('https://bo.nalog.ru/')
    for i, t in zip(data['dict_name'], data['dict_address']):
        time.sleep(5)
        button_extended_search = driver.find_elements_by_tag_name('button')[3]
        driver.execute_script("arguments[0].click();", button_extended_search)
        button_search = driver.find_elements_by_tag_name('button')[6]
        driver.execute_script("arguments[0].click();", button_search)
        driver.find_element_by_id('name').send_keys(i)
        driver.find_element_by_id('address').send_keys(t)
        button_search_all_fields = driver.find_element_by_class_name('form-item-checkbox_checkbox')
        driver.execute_script("arguments[0].click();", button_search_all_fields)
        button_search = driver.find_elements_by_tag_name('button')[5]
        driver.execute_script("arguments[0].click();", button_search)
        time.sleep(5)
        if check(i) == 0:
            elem = driver.find_element_by_xpath('//*[@id="root"]/main/div/div/div[2]/div[2]/a[1]').get_attribute('href')
            print(elem)
            driver.get(elem)
            time.sleep(5)
            inn_elem = driver.find_elements_by_class_name('header-card-content-item__text')[3].text
            print(inn_elem)
            inn.append(inn_elem)
            kpp_elem = driver.find_elements_by_class_name('header-card-content-item__text')[4].text
            kpp.append(kpp_elem)
            ogrn_elem = driver.find_elements_by_class_name('header-card-content-item__text')[5].text
            ogrn.append(ogrn_elem)
            capital_elem = driver.find_elements_by_class_name('item-number__main')[0].text
            capital.append(capital_elem)
            assets_elem = driver.find_elements_by_class_name('item-number__main')[1].text
            assets.append(assets_elem)
        else:
            inn.append('')
            kpp.append('')
            ogrn.append('')
            capital.append('')
            assets.append('')
    driver.close()
    print('inn = ', inn)
    commercial_inf = {'inn': inn, 'kpp': kpp, 'ogrn': ogrn, 'capital': capital, 'assets': assets}
    open_excel_file(data, commercial_inf, name_excel_file)
    # write_data_txt(data, commercial_inf)


def main():
    print('Введите через запятую координаты точки, относительно которой будет произведен поиск')
    location = input()
    print('Введите радиус поиска')
    radius = input()
    print('Введите ключевое слово')
    keyword = input()
    print('Введите названия Excel-файла вместе с расширением файла (.xls). Например "GM_Parser.xls"')
    name_excel_file = input()
    # make_file_excel()
    parsing_GM(location, radius, keyword, name_excel_file)
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
                print('Введите названия Excel-файла вместе с расширением файла (.xls). Например "GM_Parser.xls"')
                name_excel_file = input()
                parsing_GM(location, radius, keyword, name_excel_file)
            else:
                answer = None
                print('Если хотите изменить значений всех переменных - напишите "Yes" (без кавычек).')
                print('Если хотите изменить только координаты точки - напишите "location" (без кавычек).')
                print('Если хотите изменить радиус - напишите "radius" (без кавычек).')
                print('ЕСли хотите изменить и координаты точки и радиус поиска - напишите "location and radius" (без кавычек).')
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
                    print('Введите названия Excel-файла вместе с расширением файла (.xls). Например "GM_Parser.xls"')
                    name_excel_file = input()
                    parsing_GM(location, radius, keyword, name_excel_file)
                if answer == 'location':
                    location = None
                    print('Введите через запятую координаты точки, относительно которой будет произведен поиск')
                    location = input()
                    print('Введите названия Excel-файла вместе с расширением файла (.xls). Например "GM_Parser.xls"')
                    name_excel_file = input()
                    parsing_GM(location, radius, keyword, name_excel_file)
                if answer == 'radius':
                    radius = None
                    print('Введите радиус поиска')
                    radius = input()
                    print('Введите названия Excel-файла вместе с расширением файла (.xls). Например "GM_Parser.xls"')
                    name_excel_file = input()
                    parsing_GM(location, radius, keyword, name_excel_file)
        else:
            break
    print('Программа закончила работу')


if __name__ == '__main__':
    main()
