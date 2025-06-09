import os
import time
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from ics import Calendar

# Set download directory (modify for your system)
DOWNLOAD_DIR = os.path.expanduser("C:\\Automation\\CalendarUpdate")

# Set up Chrome options for auto download
chrome_options = webdriver.ChromeOptions()
prefs = {"download.default_directory": DOWNLOAD_DIR}
chrome_options.add_experimental_option("prefs", prefs)

# Open browser
driver = webdriver.Chrome(options=chrome_options)

# Navigate to the website
driver.get("https://ecampus.sejong.ac.kr/login.php")

# Log in (Modify the selectors accordingly)
username = driver.find_element(By.NAME, "username")
password = driver.find_element(By.NAME, "password")
login_button = driver.find_element(By.NAME, "loginbutton")

username.send_keys("21011712")
password.send_keys("5547Dec69@")
login_button.click()

time.sleep(5)  # Wait for the page to load

# Navigate to the calendar download page
driver.get("https://ecampus.sejong.ac.kr/calendar/export.php?course=1")

checkbox = driver.find_element(By.ID, "id_events_exportevents_courses")  # Modify ID as needed
if not checkbox.is_selected():
    checkbox.click()  # Click to check it

time.sleep(2)  # Ensure the checkbox is registered

checkbox = driver.find_element(By.ID, "id_period_timeperiod_monthnow")  # Modify ID as needed
if not checkbox.is_selected():
    checkbox.click()  # Click to check it

time.sleep(2)  # Ensure the checkbox is registered

# Find and click the download button
download_button = driver.find_element(By.ID, "id_export")
download_button.click()

time.sleep(5)  # Wait for download to complete

driver.quit()

ics_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith(".ics")]
ics_files.sort(key=lambda x: os.path.getctime(os.path.join(DOWNLOAD_DIR, x)), reverse=True)

if ics_files:
    latest_ics = os.path.join(DOWNLOAD_DIR, ics_files[0])

    # Read the .ics file
    with open(latest_ics, "r", encoding="utf-8") as file:
        calendar = Calendar(file.read())

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar_folder = namespace.GetDefaultFolder(9)  # 9 = Outlook Calendar

    # Add events to Outlook Calendar
    for event in calendar.events:
        appointment = calendar_folder.Items.Add(1)  # 1 = Outlook Appointment Item
        appointment.Subject = event.name
        appointment.Start = event.begin.datetime
        appointment.End = event.end.datetime

        if hasattr(event, 'location') and event.location:
            appointment.Location = event.location
        else:
            appointment.Location = "No location specified"

        appointment.Body += f"\nLocation: {event.location}"  # Append location to the body instead
        appointment.ReminderMinutesBeforeStart = 15  # Set reminder
        appointment.Save()

    print("✅ Calendar successfully updated!")

# Close the browser
driver.quit()
