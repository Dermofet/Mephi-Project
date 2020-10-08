import requests
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy

URL = 'https://maps.googleapis.com/maps/api/place/textsearch/json'
key = 'AIzaSyA5xIulXbroyFR7-O8vChDJYeMYRCML8jA'


def find_places(response_json):
    name = response_json.get('results').get('name').split(',')
    address = response_json.get('results').get('formatted_address').split(',')
    place_id = response_json.get('results').get('place_id').split(',')
    rating = response_json.get('results').get('rating').split(',')
    user_rating_total = response_json.get('results').get('user_rating_total').split(',')
    url = response_json.get('results').get('url').split(',')
    data = {'name': name,
            'address': address,
            'place_id': place_id,
            'rating': rating,
            'user_rating_total': user_rating_total,
            'url': url}
    return data


def make_excel_file():
    book = xlwt.Workbook('utf8')
    font = xlwt.easyxf('font: height 240,name Arial,colour_index black, bold off,\
        italic off; align: wrap on, vert top, horiz left;\
        pattern: pattern solid, fore_colour white;')
    sheet = book.add_sheet('GM_parser')
    sheet.write(0, 0, '', font)
    sheet.row(1).height = 2500
    sheet.col(0).width = 20000
    sheet.portrait = False
    sheet.set_print_scaling(85)
    book.save('ПП.xls')


def open_excel_file(x):
    rb = open_workbook("ПП.xls")
    wb = copy(rb)
    s = wb.get_sheet('GM_parser')
    s.write(0, 0, x)
    wb.save('ПП.xls')


def main():
    print('Введите через запятую координаты точки, относительно которой будет произведен поиск')
    location_ = input()
    print('Введите радиус поиска')
    radius_ = input()
    print('Введите ключевое слово')
    keyword_: str = input()
    payload = {'query': keyword_, 'location': location_, 'radius': radius_, 'key': key}
    response = requests.get(URL, params=payload)
    response_json = response.json()
    data = find_places(response_json)
    make_excel_file()
    x = data['name']
    open_excel_file(x)
    x = data['address']
    open_excel_file(x)
    x = data['place_id']
    open_excel_file(x)
    x = data['rating']
    open_excel_file(x)
    x = data['users_rating_total']
    open_excel_file(x)
    x = data['url']
    open_excel_file(x)


if __name__ == '__main__':
    main()
