import requests
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_xml_from_sitemap():
    response = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    return response.content


def get_courses_list(xml, amount):
    tree = etree.fromstring(xml)
    courses_list = [element[0].text for element in tree]
    return courses_list[:amount]


def get_course_info(course_url):
    dirty_html = requests.get(course_url).content
    soup = BeautifulSoup(dirty_html, 'html.parser')
    course_name = soup.find_all('h2')[0].get_text()
    course_language = soup.find_all('div', 'rc-Language')[0].get_text()
    course_start_date = soup.find_all('div', 'startdate')[0].get_text()
    course_length = len(soup.find_all('div', 'week'))
    try:
        course_ratings = soup.find_all('div', 'ratings-text')[0].get_text()
    except IndexError:
        course_ratings = 'No ratings yet' # None
    course_info = {
        'course_name': course_name,
        'course_language': course_language,
        'course_start_date': course_start_date,
        'course_length': course_length,
        'course_ratings': course_ratings
    }
    return course_info


def fill_courses_info_to_xlsx(courses_info, filepath='./courses.xlsx'):
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
            course['course_ratings']
        ]
        ws1.append(course_row)
    wb.save(filepath)


def main():
    print('Collecting data....')
    course_xml = get_xml_from_sitemap()
    course_quantity = 5
    courses_list = get_courses_list(course_xml, course_quantity)
    courses_info = [get_course_info(course_url) for course_url in courses_list]
    fill_courses_info_to_xlsx(courses_info)
    print('Complete! Check courses.xlsx')
    

if __name__ == '__main__':
    main()
