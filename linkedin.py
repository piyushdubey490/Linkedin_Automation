import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# LinkedIn Login Credentials
linkedin_email = os.getenv("LINKEDIN_EMAIL")
linkedin_password = os.getenv("LINKEDIN_PASSWORD")

# Gmail Credentials
sender_email = os.getenv("SENDER_EMAIL")
sender_password = os.getenv("SENDER_PASSWORD")
recipient_email = os.getenv("RECIPIENT_EMAIL")

# Load Excel file
excel_file = "/Users/piyushdubey/Desktop/ABInBev/Book1.xlsx"  
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active

# Get previous occurrence data
previous_unread_messages = sheet['A1'].value if sheet['A1'].value else 0
previous_unread_notifications = sheet['B1'].value if sheet['B1'].value else 0

# LinkedIn Login and Data Extraction
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.get("https://www.linkedin.com/login")
time.sleep(1)  # Wait for page to load

# Fill in login details
driver.find_element(By.ID, "username").send_keys(linkedin_email)
driver.find_element(By.ID, "password").send_keys(linkedin_password)
driver.find_element(By.XPATH, "//button[text()='Sign in']").click()
time.sleep(0)  # Wait for login

# Extract unread messages and notifications
unread_messages_element = driver.find_element(By.XPATH, '//*[@id="global-nav"]/div/nav/ul/li[4]/a')
unread_notifications_element = driver.find_element(By.XPATH, '//*[@id="global-nav"]/div/nav/ul/li[5]/a')

unread_messages = 0
if unread_messages_element.text != 'Messaging':
    unread_messages = int(unread_messages_element.text)

unread_notifications_text = unread_notifications_element.text
unread_notifications = int(unread_notifications_text.split()[0])

# Compare with previous data
messages_difference = unread_messages - previous_unread_messages
notifications_difference = unread_notifications - previous_unread_notifications

# Update Excel with current data
sheet['A1'].value = unread_messages
sheet['B1'].value = unread_notifications
wb.save(excel_file)

# Create Email Body
email_body = f"""
    <html>
        <body>
            <h2>LinkedIn Unread Notifications and Messages Update</h2>

            <p>Number of Unread Messages: {unread_messages}</p>
            <p>Number of Unread Notifications: {unread_notifications}</p>
            <p>Comparison with Previous Data:</p>
            <p>Messages Difference: {messages_difference}</p>
            <p>Notifications Difference: {notifications_difference}</p>
            <p>Profile Link: <a href="https://www.linkedin.com/in/piyushdubey490/">Piyush Dubey</a></p>
            
        </body>
    </html>
"""

# Send Email Notification
msg = MIMEMultipart()
msg['From'] = os.getenv("SENDER_EMAIL")
msg['To'] = os.getenv("RECIPIENT_EMAIL")
msg['Subject'] = "LinkedIn's Unread Notifications and Messages Update"
msg.attach(MIMEText(email_body, 'html'))

try:
    with smtplib.SMTP('smtp-mail.outlook.com', 587) as smtp: 
        smtp.starttls()
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)
except Exception as e:
    print("An error occurred while sending the email:", str(e))

# Quit WebDriver
driver.quit()
