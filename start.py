import os
import openpyxl as px
import time
import MySQLdb

from urllib.parse import urljoin
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select


ROOT_PATH = os.getcwd()
EVIDENCE_ROOT_PATH = os.path.join(ROOT_PATH, 'evidence')
HOST_NAME = 'http://127.0.0.1:8000'
POS_TEST_CASE_START_ROW = 5
POS_INPUT_START_ROW = 3
DB_USER = 'root'
DB_PWD = 'root'
DB_HOST = 'localhost'


def main():
    driver = webdriver.Chrome(os.path.join(ROOT_PATH, 'chromedriver.exe'))
    driver.maximize_window()
    driver.get(HOST_NAME)
    try:
        driver.find_element_by_id('id_username').send_keys('admin')
        driver.find_element_by_id('id_password').send_keys('admin')
        driver.find_element_by_xpath('//*[@id="login-form"]/div[3]/button').click()
        print('adminでログインしました。')
    except NoSuchElementException:
        pass

    for path in collect_test_files():
        test_xlsx_file(path, driver)

    return driver


def collect_test_files():
    file_list = []
    for root, dirs, files in os.walk(ROOT_PATH):
        for name in files:
            if name.endswith('.xlsx') and name.startswith('test'):
                file_list.append(os.path.join(root, name))
    return file_list


def test_xlsx_file(path, driver):
    book = px.load_workbook(path)
    sheet_case = book['テストケース']
    for i in range(POS_TEST_CASE_START_ROW, sheet_case.max_row + 1):
        case_no = sheet_case['B{}'.format(i)].value
        input = sheet_case['C{}'.format(i)].value
        expect = sheet_case['D{}'.format(i)].value
        if input:
            input_data(book[input], driver)
        if expect:
            expect_data(book[expect], driver, case_no)


def input_data(sheet, driver):
    url = sheet['B1'].value
    if not url:
        return False
    driver.get(urljoin(HOST_NAME, url))
    form_name = None
    for i in range(POS_INPUT_START_ROW, sheet.max_row + 1):
        if sheet['A{}'.format(i)].value == "FORM ID":
            form_name = sheet['B{}'.format(i)].value
        elif sheet['A{}'.format(i)].value == "FIELD":
            name = sheet['B{}'.format(i)].value
            value = sheet['C{}'.format(i)].value

            if form_name and name and value:
                element = driver.find_element_by_xpath('//form[@id="{}"]//*[@name="{}"]'.format(form_name, name))
                if element.tag_name == 'input':
                    input_type = element.get_attribute('type')
                    if input_type == "checkbox":
                        label = driver.find_element_by_xpath('//form[@id="{}"]//*[@for="{}"]'.format(form_name, 'id_' + name))
                        if value is True:
                            if not element.is_selected():
                                label.click()
                        elif value is False:
                            if element.is_selected():
                                label.click()
                    else:
                        element.send_keys(value)
                elif element.tag_name == 'select':
                    data_select_id = element.get_attribute('data-select-id')
                    if data_select_id:
                        select_option_id = 'select-options-{}'.format(data_select_id)
                        # ドロップダウンリストを展開する
                        driver.find_element_by_css_selector('[data-activates={}]'.format(select_option_id)).click()
                        driver.find_element_by_css_selector('[data-activates={}]'.format(select_option_id)).click()
                        time.sleep(1)
                        # 指定項目を選択する。
                        xpath = '//ul[@id="{}"]//span[contains(text(), "{}")]'.format(select_option_id, value)
                        list_element = driver.find_element_by_xpath(xpath)
                        list_element.click()
                    else:
                        select_element = Select(element)
                        select_element.select_by_visible_text(value)
        elif sheet['A{}'.format(i)].value == "CLICK":
            xpath = sheet['B{}'.format(i)].value
            driver.find_element_by_xpath(xpath).click()


def expect_data(sheet, driver, case_no):
    for i in range(1, sheet.max_row + 1):
        table_name = sheet['A{}'.format(i)].value
        sql = sheet['B{}'.format(i + 1)].value
        if table_name and sql:
            results = select_data(sql)
            print(results[0])


def select_data(sql):
    con = MySQLdb.connect(user=DB_USER, passwd=DB_PWD, db='areaparking', host=DB_HOST, charset='utf8')
    cursor = con.cursor()
    results = []
    try:
        sql = sql.rstrip().rstrip(';') + " limit 100;"
        cursor.execute(sql)
        print('SELECT:', sql)
        for row in cursor:
            results.append(row)
    except Exception as e:
        print(e)
    finally:
        cursor.close()
        con.close()
    return results



if __name__ == '__main__':
    driver = main()
    # driver.quit()
