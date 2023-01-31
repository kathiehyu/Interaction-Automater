from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


import xlrd2 as xlrd

# import os

def main():
        loc = "roster.xlsx"

        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        
        driver = webdriver.Chrome()
        driver.get("https://gatech.co1.qualtrics.com/jfe/form/SV_1KSYHuch7lUC6pw")
        driver.implicitly_wait(100)
        assert "Online" in driver.title
        ac = ActionChains(driver)        
        for i in range(1, sheet.nrows):
                resident_i = {}

                name = sheet.cell_value(i, 0)
                if name == "":
                        break
                resident_i["name"] = name
                print(name)
                room = sheet.cell_value(i, 1)[0:3]
                resident_i["room"] = room
                print(room)
                date = sheet.cell_value(i, 2)
                if date == "":
                        break
                resident_i["date"] = date
                print(date)
                
                community = driver.find_element(By.ID, "QR~QID36")
                select = Select(community)
                select.select_by_value("11")
                for opt in select.options:
                        if opt.is_selected():
                                print(opt.get_attribute('innerText'))
                                # print(opt.get_attribute('value'))

                raname = driver.find_element(By.ID, "QR~QID45")
                select = Select(raname)
                select.select_by_visible_text("Kathie Huynh")
                for opt in select.options:
                        if opt.is_selected():
                                print(opt.get_attribute('innerText'))

                nxt = driver.find_element(By.ID, "NextButton")
                nxt.click()
                driver.implicitly_wait(10)

                resname = driver.find_element(By.ID, "QR~QID2")
                resname.send_keys(name)
                # building = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "QID132-9-label")))
                building = driver.find_element(By.ID, "QID132-9-label") # for WDS: change to "~6"
                driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                # select = Select(building)
                # select.select_by_visible_text("WDN - Woodruff North")
                # building.click()
                # building.send_keys(Keys.ENTER)
                print(building.size)
                # ac.move_to_element(building).click().perform()
                #nxt = driver.find_element(By.ID, "NextButton")
                #nxt.click()
                driver.implicitly_wait(10)
                nxt = driver.find_element(By.ID, "NextButton")
                nxt.click()
                print(driver.title)
                break

##        with open("interaction automation - Sheet1.csv", newline = '') as csvfile:
##                # fields=['inperson','videocall','phonecall','text','inhall','campus',
##                # 'social','academic','involvement','community','current','crisis','intent']
##                reader = csv.DictReader(csvfile)
##                for name in reader.fieldnames:
##                        print(name)
##                # for row in reader:
                        
                

        driver.close()
        driver.quit()

if __name__ == "__main__":
        main()
