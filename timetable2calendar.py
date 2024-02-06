import os.path
import requests
import smtplib
import ssl
from datetime import datetime, timedelta, date
from email.message import EmailMessage
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv('cred.env')

class Email:
    def __init__(self):
        self.sender = os.getenv('SENDER')
        self.password = os.getenv('PASSWORD')
        self.receiver = os.getenv('RECEIVER')
        self.subject = 'Your Weekly APU Timetable is Ready!'

    def send_email(self, html_content):
        context = ssl.create_default_context()
        em = EmailMessage()
        em['From'] = self.sender
        em['To'] = self.receiver
        em['Subject'] = self.subject
        em.set_content(html_content, subtype='html')

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(self.sender, self.password)
            smtp.sendmail(self.sender, self.receiver, em.as_string())

def get_credentials(SCOPES):
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
    return creds

def get_week_start():
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    return week_start

def fetch_timetable(week, intake, intake_group):
    response = requests.get(f"https://api.apiit.edu.my/timetable-print/index.php?Week={week}&Intake={intake}&Intake_Group={intake_group}&print_request=print_tt")
    soup = BeautifulSoup(response.text, 'html.parser')
    return soup.find('table', class_='table')

def create_event(service, entry):
    date_str = entry['Date']
    time_str = entry['Time']

    date = datetime.strptime(date_str, '%a, %d-%b-%Y')
    start_time_str, end_time_str = time_str.split(' - ')
    start_time = datetime.strptime(start_time_str, '%H:%M')
    end_time = datetime.strptime(end_time_str, '%H:%M')
    start_datetime = date.replace(hour=start_time.hour, minute=start_time.minute)
    end_datetime = date.replace(hour=end_time.hour, minute=end_time.minute)

    start_isoformat = start_datetime.isoformat()
    end_isoformat = end_datetime.isoformat()
    time_zone = 'Asia/Kuala_Lumpur'

    event = {
        'summary': f"{entry['Subject/Module']}",
        'location': f"{entry['Classroom']}",
        'start': {
            'dateTime': f'{start_isoformat}',
            'timeZone': f'{time_zone}',
        },
        'end': {
            'dateTime': f'{end_isoformat}',
            'timeZone': f'{time_zone}',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }
    return service.events().insert(calendarId='primary', body=event).execute()

def main():
    intake = "APU3F2311CS(AI)"
    intake_group = "G1"
    SCOPES = ["https://www.googleapis.com/auth/calendar"]
    remove_list = ['ALG']

    week_start = get_week_start()
    calendar_link = f"https://calendar.google.com/calendar/u/0/r/week/{week_start.year}/{week_start.month}/{week_start.day}"

    creds = get_credentials(SCOPES)
    service = build('calendar', 'v3', credentials=creds)

    timetable_table = fetch_timetable(week_start, intake, intake_group)

    if timetable_table:
        timetable_data = []
        rows = timetable_table.find_all('tr')[2:]

        if rows:
            for row in rows:
                cells = row.find_all('td')

                date = cells[0].text.strip()
                time = cells[1].text.strip()
                classroom = cells[2].text.strip()
                location = cells[3].text.strip()
                subject = cells[4].text.strip()
                lecturer = cells[5].text.strip()

                module_name = subject.split('-')[3]
                if module_name not in remove_list:
                    timetable_data.append({
                        'Date': date,
                        'Time': time,
                        'Classroom': classroom,
                        'Location': location,
                        'Subject/Module': subject,
                        'Lecturer': lecturer
                    })

            for entry in timetable_data:
                create_event(service, entry)

            html_content = f"""
                <html>
                    <body>
                        <p>Hi, your APU timetable for the week of {week_start} is ready.</p>
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
            print("No class this week!")
            html_content = f"""
                <html>
                    <body>
                        <p>Hi, you have no class for the week of {week_start}.</p>
                        <p>Enjoy your holiday:)</p>
                    </body>
                </html>
            """
            email_client = Email()
            email_client.send_email(html_content)
    else:
        print("Timetable not found")


if __name__ == "__main__":
    main()
