# -*- coding: utf-8 -*-
"""
Created on Sun Sep 22 18:41:10 2019

@author: Amirhossein Ghaffari
"""

from selenium.webdriver import Firefox
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from time import sleep
import xlsxwriter
import os
import time


############################################################################
def get_all_data(browser):
    count = 1
    all_data = [["city", "station", "co", "o3", "no2", "so2", "pm10", "pm2.5", "aqi", "max polutant", "datetime"]]
    sleep(0.1)
    row_table = browser.find_element_by_css_selector("#grid > table:nth-child(1)")
    all_total_rows = row_table.find_elements_by_tag_name('tr')
    for row in all_total_rows:
        row_data = []
        rows_of_row = row.find_elements_by_tag_name('td')
        [row_data.append(d.text) for d in rows_of_row]
        len_row = len(row_data)
        if len_row > 11:
            temp = row_data[-11:]
            if temp[0] == "تهران" or temp[0] == "البرز" or temp[0] == "قزوین" or temp[0] == "قم" or temp[
                0] == "سمنان" or temp[0] == "مركزی":
                all_data.append(temp)
        count = count + 1

    return all_data


############################################################################

year = 10
month = 12
day = 31
hour = 24
minute = 0
second = 0
delay = 1
excel_table = []

# make output folders
output_path = r'C:\Users\Amirhossein Ghaffari\Desktop\polution\Iran_polution'
if not os.path.exists(output_path):
    os.makedirs(output_path)

# open web page in browser
browser = Firefox()
browser.get('http://aqms.doe.ir/Home/AQI')
browser.minimize_window()
# browser.maximize_window()

# find search button
while True:
    try:
        myElem = WebDriverWait(browser, delay).until(
            EC.element_to_be_clickable((By.ID, "btnSearch")))
        print("Page is ready!")
        break
    except TimeoutException:
        print("Loading took too much time!")

for y in range(1400 - year, 1400):
    output_path_y = output_path + "/" + str(y)
    if not os.path.exists(output_path_y):
        os.makedirs(output_path_y)
    for m in range(1, month + 1):
        output_path_m = output_path_y + "/" + str(m)
        if not os.path.exists(output_path_m):
            os.makedirs(output_path_m)
        if m > 6 and day == 31:
            day_e = 30
        else:
            day_e = day
        for d in range(1, day_e + 1):
            excel_table = []
            for h in range(hour):
                # import date and time
                # date_time = datetime.datetime(y, m, d, h, minute, second)
                # date_time_str = date_time.strftime("%Y/%m/%d %H:%M:%S")
                date_time_str = "{0}/{1}/{2} {3}:00:00".format(y, m, d, h)
                try:
                    user_field = browser.find_element_by_name("DateStr")
                    user_field.clear()
                    user_field.send_keys(date_time_str)

                    # click on search button
                    myElem.click()
                    sleep(2)

                    # click on istgah tab
                    browser.find_element_by_css_selector("a[href*='#tab_2']").click()
                    sleep(5)

                    # download table in list file
                    excel_table.append(get_all_data(browser))
                    # numpy_table = np.asarray(excel_table[h])
                except:
                    browser.close()

                    # open web page in browser
                    browser = Firefox()
                    browser.get('http://aqms.doe.ir/Home/AQI')
                    browser.minimize_window()

                    while True:
                        try:
                            myElem = WebDriverWait(browser, delay).until(
                                EC.element_to_be_clickable((By.ID, "btnSearch")))
                            print("Page is ready!")
                            break
                        except TimeoutException:
                            print("Loading took too much time!")

                    # import date and time
                    #                    date_time = datetime.datetime(y, m, d, h, minute, second)
                    #                    date_time_str = date_time.strftime("%Y/%m/%d %H:%M:%S")
                    date_time_str = "{0}/{1}/{2} {3}:00:00".format(y, m, d, h)
                    user_field = browser.find_element_by_name("DateStr")
                    user_field.clear()
                    user_field.send_keys(date_time_str)

                    # click on search button
                    myElem.click()
                    sleep(2)

                    # click on istgah tab
                    browser.find_element_by_css_selector("a[href*='#tab_2']").click()
                    sleep(5)

                    # download table in list file
                    excel_table.append(get_all_data(browser))
                    # numpy_table = np.asarray(excel_table[h])

            # write data to excel file
            with xlsxwriter.Workbook(output_path_m + "/" + str(d) + '.xlsx') as workbook:
                for h in range(hour):
                    worksheet = workbook.add_worksheet(str(h))
                    for row_num, data in enumerate(excel_table[h]):
                        worksheet.write_row(row_num, 0, data)
