#!/usr/bin/env python3

import requests
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def fetch_content_from_url(url):
    return requests.get(url).content


def get_courses_url_list(xml, amount):
    tree = etree.fromstring(xml)
    courses_list = [element[0].text for element in tree]
    return courses_list[:amount]


def get_course_info(course_url):
    plain_html = fetch_content_from_url(course_url)
    soup = BeautifulSoup(plain_html, 'html.parser')
    course_name = soup.find_all('h2')[0].get_text()
    course_language = soup.find_all('div', 'rc-Language')[0].get_text()
    course_start_date = soup.find_all('div', 'startdate')[0].get_text()
    course_length = len(soup.find_all('div', 'week'))
    try:
        course_ratings = soup.find_all('div', 'ratings-text')[0].get_text()
    except IndexError:
        course_ratings = None
    course_info = {
        'course_name': course_name,
        'course_language': course_language,
        'course_start_date': course_start_date,
        'course_length': course_length,
        'course_ratings': course_ratings
    }
    return course_info


def fill_courses_info_to_xlsx(courses_info):
    wb = Workbook()
    ws1 = wb.active
    table_title = [
        'Course name',
        'Language',
        'Start date',
        'Duration (week)',
        'Rating'
    ]
    ws1.append(table_title)
    for course in courses_info:
        course_row = [
            course['course_name'],
            course['course_language'],
            course['course_start_date'],
            course['course_length'],
            course['course_ratings'] or 'No ratings yet'
        ]
        ws1.append(course_row)
    return wb


def save_to_file(filled_workbook, filepath='./cources.xlsx'):
    filled_workbook.save(filepath)


def main():
    print('Collecting data....')
    xml_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    course_quantity = 10
    course_xml = fetch_content_from_url(xml_url)
    courses_url_list = get_courses_url_list(course_xml, course_quantity)
    courses_info = [get_course_info(course_url) for course_url in courses_url_list]
    filled_workbook = fill_courses_info_to_xlsx(courses_info)
    save_to_file(filled_workbook)
    print('Complete! Check courses.xlsx')
    

if __name__ == '__main__':
    main()
