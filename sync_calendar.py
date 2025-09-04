import os
import requests
from icalendar import Calendar
from O365 import Account

# === CONFIGURATION ===
ICS_URL = "webcal://api.eventlink.com/?m=Calendar&a=iCalFeedByCalendarID&id=59b5241e-991e-45e3-95f7-34cfd6a49e4f&token=3134fb70-9f9d-4142-b1cd-1dcdb0a266be"

# Outlook app credentials (stored as environment variables for security)
CLIENT_ID = os.getenv("OUTLOOK_CLIENT_ID")
CLIENT_SECRET = os.getenv("OUTLOOK_CLIENT_SECRET")
TENANT_ID = os.getenv("OUTLOOK_TENANT_ID")

# === AUTHENTICATE WITH OUTLOOK ===
credentials = (CLIENT_ID, CLIENT_SECRET)
account = Account(credentials, auth_flow_type="credentials", tenant_id=TENANT_ID)

if not account.is_authenticated:
    account.authenticate()

schedule = account.schedule()
calendar = schedule.get_default_calendar()

# === DOWNLOAD EVENTLINK CALENDAR ===
print("Downloading Eventlink calendar...")
ics_url = ICS_URL.replace("webcal://", "https://")  # Convert webcal:// to https://
resp = requests.get(ics_url)

if resp.status_code != 200:
    raise Exception(f"Failed to fetch ICS feed. Status: {resp.status_code}")

cal = Calendar.from_ical(resp.text)

# === SYNC EVENTS ===
print("Parsing events and syncing to Outlook...")

for component in cal.walk():
    if component.name == "VEVENT":
        summary = str(component.get("summary"))
        start = component.get("dtstart").dt
        end = component.get("dtend").dt
        location = str(component.get("location", ""))

        # Create new Outlook event
        event = calendar.new_event()
        event.subject = summary
        event.start = start
        event.end = end
        if location:
            event.location = location

        try:
            event.save()
            print(f"✅ Added: {summary} ({start} - {end})")
        except Exception as e:
            print(f"❌ Failed to add event {summary}: {e}")

print("🎉 Sync complete!")
