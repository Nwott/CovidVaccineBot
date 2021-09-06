from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import time
from threading import Timer
import winrt.windows.ui.notifications as notifications
import winrt.windows.data.xml.dom as dom
import json
from datetime import date
import winsound

filename = "notification.wav"

# open json file containing personal information
with open("files/info.json", "r") as f:
    info = json.load(f)

app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe'

# create notifier
nManager = notifications.ToastNotificationManager
notifier = nManager.create_toast_notifier(app)

# define your notification as string
tString = """
  <toast>
    <visual>
      <binding template='ToastGeneric'>
        <text>COVID Vaccine Available</text>
        <text>There is a COVID vaccine cancellation.</text>
      </binding>
    </visual>       
  </toast>
"""

# convert notification to an XmlDocument
xDoc = dom.XmlDocument()
xDoc.load_xml(tString)

# initialize selenium
driver = webdriver.Chrome("C:/Users/Nwott/Documents/Development/GIVECOVIDVACCINERN/files/chromedriver.exe")
driver.get("https://covid19.ontariohealth.ca")

def is_between(time, time_range):
    if time_range[1] < time_range[0]:
        return time >= time_range[0] or time <= time_range[1]
    return time_range[0] <= time <= time_range[1]

def check_loop(counter):
  if counter == 0:
    # first vaccine location
    driver.find_element_by_xpath("/html/body/div[1]/div/main/div[1]/div[2]/div[1]/div[2]/button").click()
    xpath = '//*[@id="root"]/div/main/div/div[2]/div[3]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/table/tbody'
  elif counter == 1:
    # second vaccine location
    time.sleep(5)
    driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[1]/p[1]/a").click()
    time.sleep(2)
    driver.find_element_by_xpath("/html/body/div[1]/div/main/div[1]/div[2]/div[2]/div[2]/button").click()
    xpath = '//*[@id="root"]/div/main/div/div[2]/div[3]/div/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/table/tbody'
    time.sleep(3)
  elif counter == 2:
    # third vaccine location
    time.sleep(5)
    driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[1]/p[1]/a").click()
    time.sleep(2)
    driver.find_element_by_xpath("/html/body/div[1]/div/main/div[1]/div[2]/div[3]/div[2]/button").click()
    xpath = '//*[@id="root"]/div/main/div/div[2]/div[3]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/table/tbody'
    time.sleep(3)
  today = date.today()
  d1 = today.strftime("%m")
  time.sleep(5)
  # check if it's on may i forgot to do for june ill do it later
  if d1 == "05":
    # find table for location
    table = driver.find_element_by_xpath(xpath)
    # get all rows
    rows = table.find_elements_by_tag_name("tr")
    # make a list for all days with no vaccine cancellations
    blocked_days = driver.find_elements_by_css_selector(".CalendarDay__blocked_out_of_range")
    d2 = today.strftime("%d")
    available_days = []

    for i in range(0, len(rows)):
      # find individual table cells
      columns = rows[i].find_elements_by_tag_name("td")
      
      for i in range(0, len(columns)):
        # if the table cell is not a blocked day, add it to available days list
        if not columns[i] in blocked_days and not columns[i].text == d2:
          try:
            # remove all days that are shown as empty
            val = int(columns[i].text)
            available_days.append(columns[i])
          except ValueError:
            continue
    # if there are no available days        
    if len(available_days) <= 0:
      # click the next arrow on the calendar (https://prnt.sc/13d35e1)
      driver.find_element_by_class_name("calendar__next").click()
      # find table for june
      table = driver.find_element_by_xpath('//*[@id="root"]/div/main/div/div[2]/div[3]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/table/tbody')
      # does the same thing as above but for june
      rows = table.find_elements_by_tag_name("tr")
      blocked_days = driver.find_elements_by_css_selector(".CalendarDay__blocked_out_of_range")
      available_days = []

      for i in range(0, len(rows)):
        columns = rows[i].find_elements_by_tag_name("td")
        
        for i in range(0, len(columns)):
          if not columns[i] in blocked_days:
            try:
              val = int(columns[i].text)
              available_days.append(columns[i])
            except ValueError:
              continue
      # repeat everything above but for the other locations
      if len(available_days) <= 0 and counter != 2:
        check_loop(counter + 1)
      elif counter == 2:
        time.sleep(5)
        driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[1]/p[1]/a").click()
        time.sleep(10)
        check_loop(0)
      # if there are more than 0 available days  
      elif len(available_days) > 0:
        # why did i name it early days idk it doesnt make sense whatever
        early_days = []
        # check each available day to see if it's before june 14 yeah my dad changed it from june 6th to 14th yay cool
        for day in available_days:
          if int(day.text) <= 14:
            # add to list of days before june 14th
            early_days.append(day)
            
        # if there is a day before june 14th available, send windows notification
        if len(early_days) > 0:
          winsound.PlaySound(filename, winsound.SND_FILENAME)
          notifier.show(notifications.ToastNotification(xDoc))
        # if not then go to other locations
        else:
          if counter != 2:
            check_loop(counter + 1)
          elif counter == 2:
            time.sleep(5)
            driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[1]/p[1]/a").click()
            time.sleep(10)
            check_loop(0)
    # if vaccine date available in may send windows notification
    else:
      winsound.PlaySound(filename, winsound.SND_FILENAME)
      notifier.show(notifications.ToastNotification(xDoc))

      for day in available_days:
        available_times = driver.find_elements_by_class_name("tw-p-1 tw-w-1/3 tw-inline-block")

        for times in available_times:
          timeslots = driver.find_elements_by_xpath(".//*")

          if day.text == "29" or day.text == "30":
            timeslots.click()

  elif d1 == "06":
    # find table for june
      table = driver.find_element_by_xpath('//*[@id="root"]/div/main/div/div[2]/div[3]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/table/tbody')
      # does the same thing as above but for june
      rows = table.find_elements_by_tag_name("tr")
      blocked_days = driver.find_elements_by_css_selector(".CalendarDay__blocked_out_of_range")
      available_days = []
      d2 = today.strftime("%d")

      for i in range(0, len(rows)):
        columns = rows[i].find_elements_by_tag_name("td")
        
        for i in range(0, len(columns)):
          if not columns[i] in blocked_days and not columns[i].text == d2:
            try:
              val = int(columns[i].text)
              available_days.append(columns[i])
            except ValueError:
              continue
      # repeat everything above but for the other locations
      if len(available_days) <= 0 and counter != 2:
        check_loop(counter + 1)
      elif counter == 2:
        time.sleep(5)
        driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[1]/p[1]/a").click()
        time.sleep(10)
        check_loop(0)
      # if there are more than 0 available days  
      elif len(available_days) > 0:
        # why did i name it early days idk it doesnt make sense whatever
        early_days = []
        # check each available day to see if it's before june 14 yeah my dad changed it from june 6th to 14th yay cool
        for day in available_days:
          if int(day.text) <= 14:
            # add to list of days before june 14th
            early_days.append(day)
            
        # if there is a day before june 14th available, send windows notification
        if len(early_days) > 0:
          winsound.PlaySound(filename, winsound.SND_FILENAME)
          notifier.show(notifications.ToastNotification(xDoc))
        # if not then go to other locations
        else:
          if counter != 2:
            check_loop(counter + 1)
          elif counter == 2:
            time.sleep(5)
            driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[1]/p[1]/a").click()
            time.sleep(10)
            check_loop(0)
      
      

def booking():
  # enter personal information and then put in postal code and click next
    driver.find_element_by_id("booking_button").click()
    driver.find_element_by_id("email").send_keys(info["email"])
    driver.find_element_by_id("emailx2").send_keys(info["email"])
    driver.find_element_by_id("mobile").send_keys(info["phone"])
    driver.find_element_by_id("schedule_button").click()
    time.sleep(5)
    driver.find_element_by_id("location-search-input").send_keys(info["postal_code"])
    driver.find_element_by_xpath("/html/body/div[1]/div/main/div/div[5]/button").click()
    time.sleep(5)
    # start check loop
    check_loop(0)

def enter_info():
  # enter personal information and go next
    health_card_number = driver.find_element_by_id("hcn")
    health_card_number.send_keys(info["health_card_number"])
    driver.find_element_by_id("vcode").send_keys(info["letter_code"])
    driver.find_element_by_id("scn").send_keys(info["nine-character"])
    driver.find_element_by_id("dob").send_keys(info["birthday"])
    driver.find_element_by_id("postal").send_keys(info["postal_code"])
    driver.find_element_by_id("continue_button").click()
    booking()

def start_process():
  # click accept terms and conditions and go next
    driver.find_element_by_xpath("/html/body/div[4]/div/div/form/div[1]/div/div/label").click()
    driver.find_element_by_id("continue_button").click()
    enter_info()

# start program
start_process()
