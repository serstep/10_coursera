import requests, json, re, openpyxl, sys
from lxml import etree
from bs4 import BeautifulSoup
from random import sample


def get_courses_links_list():
    response = requests.get("https://www.coursera.org/sitemap~www~courses.xml")
    root = etree.fromstring(response.content)
    return [url.text for url in root.iter("{*}loc")]


def get_course_start_data(course_page):
    tag_with_json_data = course_page.find("script", attrs={"type":"application/ld+json"})
    if tag_with_json_data is None:
        return None
    json_data = json.loads(tag_with_json_data.text)
    try:
        return json_data["hasCourseInstance"][0]["startDate"]
    except KeyError:
        return None


def get_course_rate(course_page):
    course_rate_tag = course_page.find("div", class_="ratings-text")

    if course_rate_tag is None:
        return None

    rate = re.search(r"\d([.]\d)*", course_rate_tag.text)
    if rate is None:
        return None

    return rate.group()


def get_course_info(course_url):
    response = requests.get(course_url)
    course_page = BeautifulSoup(response.content, "html.parser")

    course_name = course_page.find("div", class_="title").text
    course_language = course_page.find("div", class_="language-info").text
    course_start_date = get_course_start_data(course_page)
    number_of_weeks = len(course_page.find_all("div", class_="week"))
    course_rate = get_course_rate(course_page)
    
    return [course_name, course_language, course_start_date, number_of_weeks, course_rate]


def get_all_courses_info(required_number):
    courses_links_list = get_courses_links_list()
    courses_info_list = []

    for url in sample(courses_links_list, required_number):
        course_info = get_course_info(url)
        courses_info_list.append(course_info)

    return courses_info_list


def output_courses_info_to_xlsx(filepath, courses_info):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.cell(row=1, column=1, value="Name")
    worksheet.cell(row=1, column=2, value="Language")
    worksheet.cell(row=1, column=3, value="Start date")
    worksheet.cell(row=1, column=4, value="Number of weeks")
    worksheet.cell(row=1, column=5, value="Rate")

    for course_number, course_info in enumerate(courses_info, start=2):
        for property_number, course_property in enumerate(course_info, start=1):
            worksheet.cell(row=course_number, column=property_number, value=course_property)

    workbook.save(filepath)


if __name__ == '__main__':
    required_courses_number = 20

    if len(sys.argv) < 2:
        print("Set filepath as first parameter please")
        exit(1)

    courses_info_list = get_all_courses_info(required_courses_number)
    print("Courses info obtained.")

    try:
        output_courses_info_to_xlsx(sys.argv[1], courses_info_list)
    except openpyxl.utils.exceptions.InvalidFileException:
        print("File writing error.")
        exit(1)

    print("Done.")
    
    
