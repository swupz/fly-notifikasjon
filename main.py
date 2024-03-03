IMAGE_FILE = "fly.png"

from win32api import GetSystemMetrics
import tkinter as tk
from datetime import datetime
import os.path
import schedule
import time
from tzlocal import get_localzone

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]
calendarEvents = []
notificatedCalendarEvents = []

def getCalendarEvents():
  """Shows basic usage of the Google Calendar API.
  Prints the start and name of the next 10 events on the user's calendar.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
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

  try:
    service = build("calendar", "v3", credentials=creds)

    # Call the Calendar API
    now = datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
    print("Getting the upcoming 10 events")
    events_result = (
        service.events()
        .list(
            calendarId="primary",
            timeMin=now,
            maxResults=10,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events = events_result.get("items", [])

    if not events:
      print("No upcoming events found.")
      return
    else:
      # Prints the start and name of the next 10 events
      for event in events:
        start = event["start"].get("dateTime", event["start"].get("date"))
        print(start, event["summary"])

      global calendarEvents
      calendarEvents = events

  except HttpError as error:
    print(f"An error occurred: {error}")


class Party():
    def __init__(self, message):
        # Set dimensions of the window
        self.window_width = GetSystemMetrics(0)
        self.window_height = GetSystemMetrics(1)

        # Set dimensions of the canvas
        self.canvas_width = self.window_width + 10
        self.canvas_height = self.window_height + 10

        # Set the speed of the plane and rectangle
        self.speed = 1

        # Create the main window
        self.root = tk.Tk()
        self.root.title("Moving Plane")

        # Remove window decorations
        self.root.config(highlightbackground='black')
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.wm_attributes('-transparentcolor', 'black')

        # Set window position
        self.root.geometry(f"{self.canvas_width}x{self.canvas_height}+{-5}+{-5}")

        # Create a canvas
        self.canvas = tk.Canvas(self.root, width=self.canvas_width, height=self.canvas_height, bg='black')
        self.canvas.pack()

        # Load the plane image
        self.plane_image = tk.PhotoImage(file=IMAGE_FILE) # Replace 'plane_image.png' with the path to your plane image

        # Create the plane on the canvas
        self.plane = self.canvas.create_image(self.window_width, 60, image=self.plane_image)

        self.lengthOfBanner = getLengthOfBanner(message)

        # Create a rectangle on the canvas
        self.rectangle = self.canvas.create_rectangle(self.canvas_width + 70, 30, self.canvas_width + self.lengthOfBanner, 90, fill='white')
        
        # Text
        self.text = self.canvas.create_text(self.canvas_width + 35 + self.lengthOfBanner * 0.5, 60, text=message, font=("Helvetica", 18), fill="#010101")

        # Start moving the objects
        self.move_objects()

        # Run the Tkinter event loop
        self.root.mainloop()

    # Function to move the plane and rectangle
    def move_objects(self):
        if self.canvas.coords(self.plane)[0] < - 500 - self.lengthOfBanner:
            self.stop()

        self.canvas.move(self.plane, -self.speed, 0)
        self.canvas.move(self.rectangle, -self.speed, 0)
        self.canvas.move(self.text, -self.speed, 0)
        self.root.after(20, self.move_objects)

    def stop(self):
       self.root.destroy()


def seconds_between_dates(date_str1, date_str2):
    date1 = datetime.fromisoformat(date_str1)
    date2 = datetime.fromisoformat(date_str2)
    delta = date2 - date1
    seconds = delta.total_seconds()
    print(delta, seconds)
    return seconds

def getLengthOfBanner(text):
   length = len(text)
   multiplier = 20

   if length > 40:
    multiplier = 13
   if length > 30:
    multiplier = 16
   elif length > 20:
    multiplier = 18

   return length * multiplier # 20 set to match font size to get correct amount of pixels

def getTimezoneNumber():
    # Get local timezone
    local_timezone = get_localzone()

    # Get the current UTC offset in hours and minutes
    utc_offset = local_timezone.utcoffset(datetime.now())

    # Convert the UTC offset to hours as a float
    utc_offset_hours = utc_offset.total_seconds() / 3600

    if utc_offset_hours == 1.0:
       return "+01:00"
    elif utc_offset_hours == 2.0:
        return "+02:00"
    else:
       return "+02:00"


def isNotInThePast(iso_date_str):
    # Parse the ISO format date string into a datetime object
    date_obj = datetime.fromisoformat(iso_date_str)

    # Get the current date and time
    timezonelll = getTimezoneNumber()
    current_time = datetime.fromisoformat(datetime.now().replace(microsecond=0).isoformat() + timezonelll)

    # Compare the parsed date with the current date
    if date_obj < current_time:
        return False  # Date is in the past
    else:
        return True  # Date is not in the past


def main():
    def fystEventMedTid():
        for event in calendarEvents:
            if "dateTime" in event["start"]:
                if isNotInThePast(event["start"]["dateTime"]):
                    return event

        return False

    firstEvent = fystEventMedTid()

    if firstEvent is not False:
        timezonelll = getTimezoneNumber()
        start = firstEvent["start"]["dateTime"]
        now = datetime.now().replace(microsecond=0).isoformat() + timezonelll
        secondsToNextEvent = seconds_between_dates(now, start)

        if (firstEvent["id"] not in notificatedCalendarEvents):

            if secondsToNextEvent < 150 and secondsToNextEvent > 0:
                message = "2 min - " + firstEvent["summary"]
                Party(message)
                notificatedCalendarEvents.append(firstEvent["id"])

   
# Schedule the function to run every 5 minutes
schedule.every(1).minutes.do(main)
schedule.every(15).minutes.do(getCalendarEvents)

# Start med events
getCalendarEvents()
main()

# Infinite loop to keep the script running
while True:
    schedule.run_pending()
    time.sleep(1)  # Sleep for 1 second to avoid high CPU usage
