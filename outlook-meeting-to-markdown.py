""" This script connects to Outlook and retrieves today's calendar events. """
import re
from datetime import datetime, timedelta
import win32com.client
from art import text2art

class OutlookMeetingToMarkDown:
    """ This class connects to Outlook and retrieves calendar events."""
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.calendar_folder = self.namespace.GetDefaultFolder(9)

    def sanitize_for_obsidian_link(self, note_title):
        # Define a pattern for invalid characters
        invalid_characters = r'[<>:"/\\|?*]'

        # Replace invalid characters with '-'
        sanitized_title = re.sub(invalid_characters, '-', note_title)

        return sanitized_title

    def run(self, start_date, end_date):
        """ This method retrieves calendar events for a specified date range."""

        # Restrict the calendar items to the specified date range
        items = self.calendar_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        restriction = f"[Start] >= '{start_date}' AND [End] < '{end_date}'"
        restricted_items = items.Restrict(restriction)

        # Loop through the restricted items and print meeting details
        subjects = []
        for item in restricted_items:
            if item.Class == 26:  # 26 corresponds to AppointmentItem
                start = item.Start
                subject = item.Subject
                location = item.Location
                try:
                    attendees = [f"[[{recipient.Name}]]" for recipient in item.Recipients]
                except Exception as _:
                    attendees = ["Error accessing recipients"]

                _ = item.Body # Might use Body later

                # Print the meeting details
                formated_date = start.strftime("%Y-%m-%d")
                subject = f"[[{formated_date} {subject}]]"
                subject = self.sanitize_for_obsidian_link(subject)
                subjects.append(subject)
                print(f"Subject: {subject}")
                print(f"Start: {start}")
                print(f"Location: {location}")
                print("Attendees:")
                for attendee in attendees:
                    print(f"  - \"{attendee}\"")
                print("-" * 40)

        print("Meeting Summary:")
        for subject in subjects:
            print(f"  - {subject}")
        print("-" * 40)

om2md = OutlookMeetingToMarkDown()

print(text2art("Today's Meetings"))
today = datetime.today().date()
start_date = today.strftime("%m/%d/%Y")
end_date = (today + timedelta(days=1)).strftime("%m/%d/%Y")
om2md.run(start_date, end_date)

print(text2art("Tomorrow's Meetings"))
today = today + timedelta(days=1)
start_date = today.strftime("%m/%d/%Y")
end_date = (today + timedelta(days=1)).strftime("%m/%d/%Y")
om2md.run(start_date, end_date)
