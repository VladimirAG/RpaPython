import pandas as pd
import logging
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime


class FsspSearch:
    fssp_adress = r"http://fssprus.ru/"

    def fillfield(self, driver, dataframe):
        driver.get(FsspSearch.fssp_adress)
        sleep(2)
        try:
            for i, row in dataframe.iterrows():
                driver.find_element_by_class_name("fix-gosbar").send_keys(Keys.ESCAPE)
                driver.find_element_by_class_name("main-form__btn").click()
                FsspSearch.clearfields(driver)

                driver.find_element_by_name("is[last_name]").send_keys(row[0])
                driver.find_element_by_name("is[first_name]").send_keys(row[1])
                driver.find_element_by_name("is[patronymic]").send_keys(row[2])
                driver.find_element_by_name("is[date]").send_keys(str(row[3]))
                driver.find_element_by_name("is[date]").send_keys(Keys.ENTER)
                sleep(2)
                if driver.find_element_by_class_name("f-popupopen"):
                    driver.find_element_by_class_name("f-popupopen").send_keys(Keys.ESCAPE)
                else:
                    logging.info("No such element")
                FsspSearch.readdata()
                driver.back()
                sleep(2)
        finally:
            driver.quit()

    @staticmethod
    def clearfields(driver):
        driver.find_element_by_name("is[last_name]").clear()
        driver.find_element_by_name("is[first_name]").clear()
        driver.find_element_by_name("is[patronymic]").clear()
        driver.find_element_by_name("is[date]").clear()

    @staticmethod
    def readdata():
        pass


try:
    df = pd.read_excel("Test.xlsx", sheet_name='Лист1',
                       usecols=['Фамилия', 'Имя', 'Отчество', 'Дата рождения'])
    df['Дата рождения'] = [datetime.date(d).strftime("%d.%m.%Y") for d in df['Дата рождения']]
    wdriver = webdriver.Chrome()
    fssp_reader = FsspSearch()
    fssp_reader.fillfield(wdriver, df)

except IOError:
    logging.error("IOError occurred")

logging.info("Process finished")
