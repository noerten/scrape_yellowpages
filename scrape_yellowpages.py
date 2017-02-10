import math
import os
import pickle
import requests
from bs4 import BeautifulSoup

from openpyxl import Workbook


SEARCH_LINK = 'http://www.yellowpages.com/search'
DOMAIN = 'http://www.yellowpages.com'


def load_pickle(filepath):
    if not os.path.exists(filepath):
        return None
    with open(filepath, 'rb') as f:
        return pickle.load(f)


def save_pickle(data, filepath):
    with open(filepath, 'wb') as f:
        pickle.dump(data, f)


def show_progress(current, total, decimals=2):
    number = current / total*100
    print('{0:,.{1}f}%'.format(number, decimals))


def make_soup(html):
    return BeautifulSoup(html, 'html.parser')


def get_html(url, with_payload=True, search_terms='Architects',
             geo_location_terms='Los%20Angeles%2C%20CA', page=1):
    if with_payload:
        payload = {'search_terms': search_terms,
                   'geo_location_terms': geo_location_terms,
                   'page': page}
    else:
        payload = None
    return requests.get(url, params=payload).text


def get_number_of_pages(html, items_per_page=30):
    soup = make_soup(html)
    pagination_text = soup.find('div', class_='pagination').p.get_text()
    number_of_items = int(pagination_text.split()[-1].split('r')[0])
    return math.ceil(number_of_items/items_per_page)


def get_one_page_info(html, page):
    soup = make_soup(html)
    one_page_info = []
    companies_soup = soup.find('div', class_='search-results organic')
    for company_soup in companies_soup.find_all('div', recursive=False):
        company_info = {}
        name_soup = company_soup.find('h3', class_='n').a
        company_info['name'] = name_soup.get_text()
        company_info['link'] = name_soup.get('href')
        try:
            maybe_site = company_soup.find('div', class_='links').a.get('href')
            if maybe_site.startswith('http'):
                company_info['website'] = maybe_site
            else:
                print('page', page, 'not website', company_info['name'],
                      maybe_site)
                company_info['website'] = None
        except AttributeError:
            print('page', page, 'no website', company_info['name'])
            company_info['website'] = None
        try:
            company_info['phone'] = company_soup.find(
                'div', class_='phones phone primary').get_text()
        except AttributeError:
            print('page', page, 'no phone', company_info['name'])
            company_info['phone'] = None
        one_page_info.append(company_info)
    return one_page_info


def get_email(html, company_name):
    soup = make_soup(html)
    try:
        raw_email = soup.find('a', class_='email-business').get('href')
        return raw_email.split('mailto:')[-1]
    except AttributeError:
        print(company_name, 'no email')
        return None


def output_info_to_xlsx(info, filepath):
    header = ['Название', 'Телефон', 'Сайт', 'E-mail']
    wbook = Workbook()
    wsheet = wbook.active
    wsheet.append(header)
    for company in info:
        row = (company['name'], company['phone'], company['website'],
               company['email'])
        wsheet.append(row)
    wbook.save(filepath)


def picklize(item, value):
    return str(item) + '_' + str(value) + '.pickle'


def main():
    info = []
    number_of_pages = get_number_of_pages(get_html(SEARCH_LINK))
    print('number of pages:', number_of_pages)
    last_page_pickle = picklize('page', number_of_pages)
    if not load_pickle(last_page_pickle):
        print('parsing search pages')
        for page in range(1, number_of_pages+1):
            info.extend(get_one_page_info(get_html(SEARCH_LINK, page=page),
                                          page))
            save_pickle(info, picklize('page', page))
            show_progress(page, number_of_pages)
    else:
        info = load_pickle(last_page_pickle)
        print('loaded info list without emails from pickle')
    info_length = len(info)
    last_company_pickle = picklize('company', info_length-1)
    print(last_company_pickle)
    if not load_pickle(last_company_pickle):
        print('parsing emails')
        for position, company in enumerate(info):
            link = DOMAIN + company['link']
            company['email'] = get_email(get_html(link, with_payload=False),
                                         company['name'])
            save_pickle(info, picklize('company', position))
            show_progress(position, info_length)
    else:
        info = load_pickle(last_company_pickle)
        print('loaded full info list')
    output_info_to_xlsx(info, 'yellowpages_architects.xlsx')
    print('saved!')


if __name__ == '__main__':
    main()
