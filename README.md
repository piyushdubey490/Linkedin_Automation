# LinkedIn Automation with Email for Unread Notification and Messages

This script uses Selenium with Chrome WebDriver to automate LinkedIn tasks, such as tracking the number of unread notifications and messages. It also sends an email notification with the updated information.



YT Demo Link : https://youtu.be/WuD7dw39N40



## Prerequisites

Make sure you have the following dependencies installed:

- Python (version 3.7 or higher)
- Selenium
- OpenPyXL
- smtplib (built-in module)
- email (built-in module)
- Chrome WebDriver
- ChromeDriverManager
- dotenv

## Instructions

1. Install the required dependencies using pip:

```shell
   pip install selenium openpyxl webdriver_manager python-dotenv
```
2. Create a .env file in the same directory as the script and add the following environment variables:

```shell
LINKEDIN_EMAIL=your_linkedin_email
LINKEDIN_PASSWORD=your_linkedin_password
SENDER_EMAIL=your_sender_email
SENDER_PASSWORD=your_sender_password
RECIPIENT_EMAIL=your_recipient_email
```
Replace the placeholder values with your actual credentials.

3. Update the excel_file variable with the path to your Excel file that will store the data.

4. Run the script:
 ```shell
python3 linkedin.py
 ```
5. The script will open LinkedIn, log in using your credentials, extract the number of unread messages and notifications, compare them with the previous data stored in the Excel file, update the Excel file, and send an email with the updated information.

#Script Explanation

The script performs the following steps:

Import necessary libraries.
Load environment variables from the .env file.
Set up LinkedIn login credentials and Gmail credentials.
Load the Excel file and get previous occurrence data.
Set up Chrome WebDriver and login to LinkedIn.
Extract the number of unread messages and notifications.
Compare with previous data and update the Excel file.
Create the email body.
Send the email notification.
Quit the WebDriver.


