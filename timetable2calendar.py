import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, date

import os.path
from dotenv import load_dotenv

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import smtplib
import ssl
from email.message import EmailMessage

load_dotenv('cred.env')

class Email:
    def __init__(self):
        self.sender = f"{os.getenv('SENDER')}"
        self.password = f"{os.getenv('PASSWORD')}"
        self.receiver = f"{os.getenv('RECEIVER')}"
        self.subject = 'Your Weekly APU Timetable is Ready!'

    def send_email(self, html_content):
        context = ssl.create_default_context()
        em = EmailMessage()
        em['From'] = self.sender
        em['To'] = self.receiver
        em['Subject'] = self.subject
        em.set_content(html_content, subtype = 'html')

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context = context) as smtp:
            smtp.login(self.sender, self.password)
            smtp.sendmail(self.sender, self.receiver, em.as_string())



intake = "APU3F2311CS(AI)"
# week = "2024-01-29"
intake_group = "G1"
SCOPES = ["https://www.googleapis.com/auth/calendar"]

today = date.today()
week = today-timedelta(days=today.weekday())
print("This is week:",week)

calendar_link = f"https://calendar.google.com/calendar/u/0/r/week/{week.year}/{week.month}/{week.day}"

creds = None

if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

service = build('calendar', 'v3', credentials = creds)

response = requests.get(f"https://api.apiit.edu.my/timetable-print/index.php?Week={week}&Intake={intake}&Intake_Group={intake_group}&print_request=print_tt")

soup = BeautifulSoup(response.text, 'html.parser')

timetable_table = soup.find('table', class_='table')

if timetable_table:
    timetable_data = []

    rows = timetable_table.find_all('tr')[2:]

    for row in rows:
        cells = row.find_all('td')

        date = cells[0].text.strip()
        time = cells[1].text.strip()
        classroom = cells[2].text.strip()
        location = cells[3].text.strip()
        subject = cells[4].text.strip()
        lecturer = cells[5].text.strip()

        timetable_data.append({
            'Date': date,
            'Time': time,
            'Classroom': classroom,
            'Location': location,
            'Subject/Module': subject,
            'Lecturer': lecturer
        })

    for entry in timetable_data:
        print(entry['Subject/Module'])
        date_str = entry['Date']
        time_str = entry['Time']

        date = datetime.strptime(date_str, '%a, %d-%b-%Y')
        start_time_str , end_time_str = time_str.split(' - ')
        start_time = datetime.strptime(start_time_str, '%H:%M')
        end_time = datetime.strptime(end_time_str, '%H:%M')
        start_datetime = date.replace(hour=start_time.hour, minute=start_time.minute)
        end_datetime = date.replace(hour=end_time.hour, minute=end_time.minute)

        start_isoformat = start_datetime.isoformat()
        end_isoformat = end_datetime.isoformat()
        time_zone = 'Asia/Kuala_Lumpur'

        event = {
            'summary' : f"{entry['Subject/Module']}",
            'location' : f"{entry['Classroom']}",
            'start' : {
                'dateTime': f'{start_isoformat}',
                'timeZone' : f'{time_zone}',
            },
            'end' : {
                'dateTime': f'{end_isoformat}',
                'timeZone' : f'{time_zone}',
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                # {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
                ],
            },
        }
        event = service.events().insert(calendarId='primary', body=event).execute()
        print ('Event created: %s' % (event.get('htmlLink')))
    
    html_content = f"""
        <html>
            <body>
                <p>Hi, your APU timetable for week of {week} is ready.</p>
                <p>Please find calendar with link below :)</p>
                <p>Link: {calendar_link}
                </p>
            </body>
        </html>
    """
    email_client = Email()
    email_client.send_email(html_content)
    print("Export Completed")
else:
    print("not found")