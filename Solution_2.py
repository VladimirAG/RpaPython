import pandas as pd
import logging
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select


class CourtSearch:
    court_adress = r"https://sudrf.ru/index.php?id=300#sp"

    def fillfield(self, driver, dataframe):
        try:
            driver.get(CourtSearch.court_adress)
            sleep(2)
            for i, row in dataframe.iterrows():
                #TODO Здесь необходимо добавить выбор региона из выпадающего меню
                #region_select = Select(driver.find_element_by_id('court_subj'))
                #region_select.select_by_value('77')
                for j in range(len(df.columns)):
                    driver.find_element_by_name("f_name").send_keys(row[j])
                    driver.find_element_by_name("f_name").send_keys(Keys.SPACE)
                sleep(2)
                CourtSearch.readdata()
                sleep(1)
                CourtSearch.clearfields(driver)
        finally:
            driver.quit()

    @staticmethod
    def clearfields(driver):
        driver.find_element_by_name("f_name").clear()

    @staticmethod
    def readdata():
        #TODO считать информацию со страницы
        pass


try:
    df = pd.read_excel("Test.xlsx", sheet_name='Лист1', usecols=['Фамилия', 'Имя', 'Отчество'])
    wdriver = webdriver.Chrome()
    fssp_reader = CourtSearch()
    fssp_reader.fillfield(wdriver, df)
except IOError:
    logging.error("IOError occurred")


logging.info("Process finished")
