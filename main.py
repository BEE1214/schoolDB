##=================================##
##=-Project:           schoolDB -=##
##=-Author:         Adam Dvorsky -=##
##=-Date:             2021-03-08 -=##
##=================================##

from logging import exception
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from time import sleep
import pandas
import openpyxl


def Web():
    options = Options()
    options.binary_location = '/usr/bin/brave-browser'
    options.add_argument('load-extension=' + '/home/adamd/.config/BraveSoftware/Brave-Browser/Default/Extensions/cfhdojbkjhnklbpkdaibdccddilifddb/3.10.2_0')
    driver_path = '/home/adamd/.chromedriver/chromedriver'
    driver = webdriver.Chrome(options=options, executable_path=driver_path)
    # wait = WebDriverWait(driver, 10)

    driver.get('https://www.stredneskoly.sk/')
    sleep(1)
    driver.switch_to.window(driver.window_handles[1]) 
    driver.close()
    driver.switch_to.window(driver.window_handles[0]) 

    # driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 'w')

    # List2
    # /home/adamd/Documents/PRIGO/Databaze/databaze_sk_python.xlsx

    iKraj = 10
    iSchool = 2
    iKrajName = 'Košický'
    # while
    iKrajend = False
    iSchoolend = False
    iName = ''
    iAdressstr = ''
    iAdressar = []
    iEmail = ''
    iActRow = 590
    iNameCol = 0
    iAdrCol = 1
    iCityCol = 2
    iEmailCol = 4

    book = openpyxl.load_workbook('/home/adamd/Gitrepos/schoolDB/databaze_sk_python.xlsx')
    writer = pandas.ExcelWriter('/home/adamd/Gitrepos/schoolDB/databaze_sk_python.xlsx', engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)


    while not iKrajend:
        try:
            driver.find_element_by_xpath(f'//*[@id="av_section_1"]/div/div/div/div/div[{iKraj}]/div[4]/a')\
                .click()
            sleep(2)
            iKraj += 1
        except NoSuchElementException:
            iKrajend = True
        while not iSchoolend:
            iClickable = False
            try:
                driver.find_element_by_xpath(f'/html/body/div[1]/div/div[2]/div/main/div/div/div[2]/div/div[2]/a[{iSchool}]')\
                    .click()
                sleep(1)
            except NoSuchElementException:
                iSchoolend = True
            except ElementNotInteractableException:
                print('Need more time to load page!!!')
                iClickable = True
                # writer.save()
                # driver.quit()
            if iSchoolend:
                print("There is no more schools!!")
            else:
                try:
                    if iClickable:
                        if timer < 15:
                            sleep(4)
                        else:
                            sleep(timer/2)
                        driver.find_element_by_xpath(f'/html/body/div[1]/div/div[2]/div/main/div/div/div[2]/div/div[2]/a[{iSchool}]')\
                            .click()
                    else:
                        iName = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/main/div/div/div[1]/div/h1').text
                        iNameData = pandas.DataFrame([iName])
                        iAdressstr = driver.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/main/div/div/div[2]/section/div/table/tbody/tr[1]/td').text
                        # iEmail = driver.findElement(By.xpath("MY_XPATH_TO_FIND_THAT_SPAN").getAttribute("innerHTML")
                        iEmailcorrect = False
                        iEmNum = 1
                        while not iEmailcorrect:
                            iEmail = driver.find_element_by_xpath(f'//*[@id="udaje"]/tbody/tr[{iEmNum}]/th').text
                            if iEmail.find("EMAILOVÁ ADRESA:") != -1 or iEmail.find("EMAILOVÉ ADRESY:") != -1:
                                iEmail = driver.find_element_by_xpath(f'//*[@id="udaje"]/tbody/tr[{iEmNum}]/td').get_attribute('innerHTML')
                                iEmailcorrect = True
                            else:
                                iEmNum += 1
                                

                        iEmail = iEmail.replace('<img title="(zavináč)" src="https://www.stredneskoly.sk/image/index/ants.png" alt="(zavináč)" width="13" height="16">', '@')
                        iEmail = pandas.DataFrame([iEmail])
                        # iEmail = driver.find_element_by_xpath('//*[@id="udaje"]/tbody/tr[3]/td/text()[1]') + '@' + driver.find_element_by_xpath('//*[@id="udaje"]/tbody/tr[3]/td/text()[2]')
                        iAdressar = iAdressstr.split(', ')
                        if len(iAdressar) > 2:
                            iAdressst = pandas.DataFrame([iAdressar[0]])
                            iAdresscity = pandas.DataFrame([iAdressar[2] + ' ' + iAdressar[1]])
                        elif len(iAdressar) == 2:
                            iAdressst = pandas.DataFrame([iAdressar[0]])
                            iAdresscity = pandas.DataFrame([iAdressar[1]])
                        else:
                            print('Wrong address')

                        iKraj = iKrajName
                        iKraj = pandas.DataFrame([iKraj])
                        iKraj.to_excel(writer, sheet_name='List2', startrow=iActRow, startcol=3, index=False, header=False)
                        iNameData.to_excel(writer, sheet_name='List2', startrow=iActRow, startcol=iNameCol, index=False, header=False)
                        iAdressst.to_excel(writer, sheet_name='List2', startrow=iActRow, startcol=iAdrCol, index=False, header=False)
                        iAdresscity.to_excel(writer, sheet_name='List2', startrow=iActRow, startcol=iCityCol, index=False, header=False)
                        iEmail.to_excel(writer, sheet_name='List2', startrow=iActRow, startcol=iEmailCol, index=False, header=False)
                        driver.back()
                        timer = 10
                        if(iSchool%2 == 0):
                            timer += 3
                        elif iSchool == 70:
                            timer = timer + 10
                        print(iSchool)
                        sleep(timer)
                        iActRow += 1
                        iSchool += 1
                except NoSuchElementException:
                    driver.back()
                    sleep(timer)
                    iSchool += 1



    

    writer.save()
    driver.quit()

def main():
    Web()
    return 0

if __name__ == '__main__':
    main()