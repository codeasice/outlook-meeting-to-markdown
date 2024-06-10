import win32com.client
from datetime import datetime, timedelta

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Access the calendar folder (9 corresponds to the Calendar folder)
calendar_folder = namespace.GetDefaultFolder(9)

# Get today's date
today = datetime.today().date()

# Define the date range (e.g., today's meetings)
start_date = today.strftime("%m/%d/%Y")
end_date = (today + timedelta(days=1)).strftime("%m/%d/%Y")

# Restrict the calendar items to the specified date range
items = calendar_folder.Items
items.IncludeRecurrences = True
items.Sort("[Start]")

# Restrict the items to those starting today
restriction = f"[Start] >= '{start_date}' AND [End] < '{end_date}'"
restricted_items = items.Restrict(restriction)

# Loop through the restricted items and print meeting details
for item in restricted_items:
    if item.Class == 26:  # 26 corresponds to AppointmentItem
        start = item.Start
        subject = item.Subject
        location = item.Location
        try:
            attendees = [f"[[{recipient.Name}]]" for recipient in item.Recipients]
        except Exception as e:
            attendees = ["Error accessing recipients"]
        body = item.Body

        # Print the meeting details
        formated_date = start.strftime("%Y-%m-%d")
        print(f"Subject: [[{formated_date} {subject}]]")
        print(f"Start: {start}")
        print(f"Location: {location}")
        print("Attendees:")
        for attendee in attendees:
            print(f"  - \"{attendee}\"")
        # print(f"Body: {body}")
        print("-" * 40)
    else:
        print("*" * 40)
        print("*********Skipping non-appointment item:" + item.Subject)
        print("*" * 40)
