from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions.wheel_actions import WheelActions as WA
import time


import xlrd2 as xlrd

# import os

halfMonth = 0

def main():
	
	completed = 0

	# access spreadsheet
	loc = "roster.xlsx"
	try:
		wb = xlrd.open_workbook(loc)
		sheet = wb.sheet_by_index(0)
	except:
		print("Can't find/open the excel. Is it in the same folder as this script?" +
			" Is it blank? Is it correctly named \"roster.xlsx\"?")
		closeDriver(completed)

	# create webdriver, access browser
	try:
		driver = webdriver.Chrome()
		driver.get("https://gatech.co1.qualtrics.com/jfe/form/SV_1KSYHuch7lUC6pw")
		driver.implicitly_wait(100)
	except:
		print("Unable to open browser/link. You must have Google " +
			"Chrome for this to work.\nIf you do, but it still isn't working, " +
			"the link may be outdated. You can try to replace the link in the " +
			"code with the correct link, but it might not work. Then you have " +
			"no choice but to manually input. Good luck.")
		closeDriver(completed)
	
	# print("Online" in driver.title)
	# print("numRows: ")
	# print(sheet.nrows)


	# getting input values common across all forms
	RA = sheet.cell_value(0, 0)
	if len(RA) == 0:
		print("Please put your name in cell (1, A)")
		closeDriver(completed)

	raEmail = sheet.cell_value(0, 25)

	# getting building
	if (sheet.cell_value(0, 1) == "WDN"):
		buildingName = "WDN - Woodruff North"
	elif (sheet.cell_value(0, 1) == "WDS"):
		buildingName = "WDS -Woodruff South"
	if len(buildingName) == 0:
		print("Please put either WDN or WDS in cell (1, B)")
		closeDriver(completed)
	print("Your building: " + buildingName)

	# getting floor
	if len(sheet.cell_value(1, 1)) != 0:
		floorNum = sheet.cell_value(1, 1)[0:1]
	if len(floorNum) == 0:
		print("every resident must have a room number. " +
			"Follow the format on the provided excel.")
		closeDriver(completed)
	print("Your floor: " + floorNum)

	print("starting row: " + str(sheet.cell_value(0, 24)))

	# fill out form for each row (resident)
	for i in range(int(sheet.cell_value(0, 24)), sheet.nrows):
		print(i)

		# print("row: " + str(i))

		# getting resident name
		name = sheet.cell_value(i, 0)
		if name == "":
			if halfMonth == 0:
				halfMonth == 1
				main()
			else:
				print("terminating due to blank resident name. " + 
					"Last row completed: " + completed)
			break
		print("Filling out form for " + name)

		# getting resident room
		room = sheet.cell_value(i, 1)[0:3]
		if room == "":
			print("terminating due to blank resident room. " +
				"Last row completed: " + completed)
			break
		# print("room: " + room)

		# getting date of interaction
		if halfMonth == 0:
			date = sheet.cell_value(i, 2)
			if date == "":
				print("CAUTION: skipping " + name + " due to no date")
				continue
			print("date: " + str(date))
		else:
			date = sheet.cell_value(i, 3)
			if date == "":
				print("CAUTION: skipping " + name + " due to no date")
				continue
			print("date: " + str(date))
		
		# getting building area (community) (Woodruff)
		community = driver.find_element(By.ID, "QR~QID36")
		select = Select(community)
		select.select_by_value("11")
		time.sleep(0.5)

		# print selected community
		for opt in select.options:
				if opt.is_selected():
						print("Your community: " + opt.get_attribute('innerText'))

		try:
			# enter ra name
			raName = driver.find_element(By.ID, "QR~QID45")
			select = Select(raName)
			select.select_by_visible_text(RA)

			selected = False
			# print selected ra name
			for opt in select.options:
					if opt.is_selected():
						selected = True
						print(opt.get_attribute('innerText'))
			if not selected:
				print("Terminating. Can't find ra name. Is it formatted/spelled " +
					"the same way in the form as what you entered in cell (0, 0)?")
				closeDriver()
			time.sleep(0.5)

			# enter ra email
			if len(raEmail) != 0:
				enterEmail = driver.find_element(By.ID, "QR~QID38")
				enterEmail.click()
				time.sleep(1)
				enterEmail.send_keys(raEmail)
				time.sleep(1)

			# next page
			nxt = driver.find_element(By.ID, "NextButton")
			nxt.click()
			time.sleep(1.5)

			# enter resident name, sometimes takes a bit of time, why?
			# i think it is the send_keys method
			resName = driver.find_element(By.ID, "QR~QID2")
			resName.click()
			time.sleep(1.5)
			resName.send_keys(name)
			print("name inputted: " + resName.get_attribute("value"))
			time.sleep(1.5)

			# select building
			# building = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='WDN - Woodruff North']")))
			building = driver.find_element(By.XPATH, "//span[text()='" +
				buildingName + "']")
			# print("Building found: " + buildingName)
			# ActionChains(driver).scroll_to_element(building).perform()
			ActionChains(driver).click(building).perform()
			time.sleep(1)

			# next page
			nxt = driver.find_element(By.ID, "NextButton")
			nxt.click()
			time.sleep(1)

			# from now on I think the element ID's are different for WDS vs WDN

			# select floor
			floor = driver.find_element(By.ID, "QR~QID" +
				("137" if sheet.cell_value(0, 1) == "WDN" else "77"))
			select = Select(floor)
			select.select_by_visible_text(floorNum)
			time.sleep(0.75)

			# next page
			nxt = driver.find_element(By.ID, "NextButton")
			nxt.click()
			time.sleep(1)

			# select room
			# floorIDNorth = {"1": "143", "2": "139", "3": "140", "4": "141", "5": "142"}
			# floorIDSouth = {"1": "147", "2": "134", "3": "144", "4": "145", "5": "146"}
			if (sheet.cell_value(0, 1) == "WDN"):
				roomNum = driver.find_element(By.ID, "QR~QID1" +
					(str(int(floorNum) + 37) if floorNum != "1" else "43"))
			else:
				idAppend = "47" if floorNum == "1" else "34"
				roomNum = driver.find_element(By.ID, "QR~QID1" +
					(str(int(floorNum) + 41) if (floorNum != "1" and floorNum != "2")
						else idAppend))
			
			select = Select(roomNum)
			print(room[0:3])
			if floorNum == "4" and room[0:3] == "418":
				# because for some reason 418 is 4118???
				select.select_by_visible_text("4118")
			else:	
				select.select_by_visible_text(room[0:3])
			time.sleep(1)

			# select date
			dateNum = driver.find_element(By.XPATH, "//*[@id='QID135_cal_t']/tbody")

			# the id starts from 0, but the day of the month may not start at 1
			# get to the first day of the month
			dateCount = 0
			while dateCount < 7:
				print(dateCount)
				# print("actual date: " + driver.find_element(By.ID, "QID135_cal_t_cell" +
				# str(dateCount)).text)
				if (driver.find_element(By.ID, "QID135_cal_t_cell" +
					str(dateCount)).text == "1"):
					break
				dateCount += 1
			
			# add desired date value to the dateCount to get the appropriate ID
			IDsuffix = dateCount - 1 + int(date);
			print("IDsuffix: " + str(IDsuffix))
			dateNum = driver.find_element(By.ID, "QID135_cal_t_cell" + str(IDsuffix))
			print("actual date: " + dateNum.text)
			dateNum.click()
			# ActionChains.click(dateNum).perform()
			time.sleep(1)

			# choosing method of interaction
			for col in range(4, 10):
				if len(sheet.cell_value(i, col)) != 0:
					method = driver.find_element(By.ID, "QID150-" + str(col - 3) +
						"-label")
					print(method.text)
					ActionChains(driver).click(method).perform()
					time.sleep(0.25)

			intent = False

			# choosing topic of interaction
			for col in range(10, 17):
				if len(sheet.cell_value(i, col)) != 0:
					if (col == 16):
						intent = True
						method = driver.find_element(By.ID, "QID41-8-label")
						ActionChains(driver).click(method).perform()
						break
					method = driver.find_element(By.ID, "QID41-" + str(col - 9) + "-label")
					print(method.text)
					ActionChains(driver).click(method).perform()
					time.sleep(0.25)

			time.sleep(0.5)

			#next page
			nxt = driver.find_element(By.ID, "NextButton")
			nxt.click()
			time.sleep(1.5)

			if intent:
				for col in range(17, 23):
					if len(sheet.cell_value(i, col)) != 0:
						method = driver.find_element(By.ID, "QID109-" +
							str(col - 16) + "-label")
						ActionChains(driver).click(method).perform()
						time.sleep(0.25)
			else:
				#next page
				nxt = driver.find_element(By.ID, "NextButton")
				nxt.click()

			time.sleep(1.5)
			print("desired comment: " + sheet.cell_value(i, 23))

			# adding comment
			if len(sheet.cell_value(i, 23)) != 0:
				comment = driver.find_element(By.ID, "QR~QID49")
				comment.click()
				time.sleep(0.75)
				comment.send_keys(sheet.cell_value(i, 23))
				time.sleep(0.75)
				print(comment.get_attribute("value"))

			#submits
			nxt = driver.find_element(By.ID, "NextButton")
			nxt.click()
			completed = i

			time.sleep(4)
		except Exception as e:
			print(e)
			print("Error encountered.", end = "")
			closeDriver(completed)		
		# break
	closeDriver(completed)   

def closeDriver(completed):
	print("Last row completed: " + str(int(completed)) +
		".\nIf it has not gone through your whole roster,\n1. check the error " +
		"message (if any) and fix\n2. Change the value " +
		"in the cell at (0, Y) or (0, 25) on roster.xlsx to the row after this one\n" + 
		"3. Rerun the program. Don't forget to set it back to 1 for next time.\n" + 
		"If there is no error message, rerun it one more time. If it still doesn't"
		" finish, good luck, you're on your own)")

	driver.close()
	driver.quit()

if __name__ == "__main__":
	main()
