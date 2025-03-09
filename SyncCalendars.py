############################################
# Uses win32com.client to interact with Outlook.
# Reads events from the source calendar.
# Checks for overlapping events in the target calendar.
# Creates "Busy" appointments in the target calendar.
# Handles errors and logs messages.
############################################

import win32com.client
from datetime import datetime, timedelta

# Configuration: Update these with your Outlook accounts and calendar names
SOURCE_ACCOUNT = "user@workemail.com"  # Source calendar account
TARGET_ACCOUNT = "user@personalemail.com"  # Target calendar account
CALENDAR_NAME = "Calendar"  # Calendar name in both accounts

# Set date range for the next month
START_DATE = datetime.now()
END_DATE = START_DATE + timedelta(days=30)

# Initialize Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")


def get_calendar(account_name, calendar_name):
    """Retrieve the calendar folder for a given Outlook account."""
    for account in namespace.Folders:
        if account.Name == account_name:
            try:
                return account.Folders(calendar_name)
            except Exception as e:
                print(f"Error accessing calendar '{calendar_name}' for account '{account_name}': {e}")
    return None


def get_events(calendar_folder, start_date, end_date):
    """Fetch events from the given calendar within the specified date range."""
    items = calendar_folder.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    # Outlook filter format: "[Start] >= 'MM/DD/YYYY HH:MM AM/PM' AND [End] <= 'MM/DD/YYYY HH:MM AM/PM'"
    filter_str = f"[Start] >= '{start_date.strftime('%m/%d/%Y %I:%M %p')}' AND [End] <= '{end_date.strftime('%m/%d/%Y %I:%M %p')}'"
    events = items.Restrict(filter_str)

    return list(events)


def event_overlaps(target_calendar, start, end):
    """Check if an event overlaps in the target calendar."""
    items = target_calendar.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    # Overlapping filter: Check if any event falls within the given time range
    filter_str = f"([Start] < '{end.strftime('%m/%d/%Y %I:%M %p')}') AND ([End] > '{start.strftime('%m/%d/%Y %I:%M %p')}')"
    conflicts = items.Restrict(filter_str)

    # Allow ±1 minute buffer to avoid millisecond mismatches
    buffer = timedelta(minutes=1)
    for conflict in conflicts:
        conflict_start = conflict.Start.replace(tzinfo=None)
        conflict_end = conflict.End.replace(tzinfo=None)

        if (conflict_start - buffer <= start <= conflict_end + buffer) or (conflict_start <= start and conflict_end >= end):
            print(f"Overlap Found: [{conflict.Subject}] ({conflict_start} - {conflict_end})")
            return True
    return False


def book_busy_time(target_calendar, event):
    """Create a 'Busy' event in the target calendar."""
    try:
        new_event = target_calendar.Items.Add(1)  # 1 = olAppointmentItem
        new_event.Subject = "Busy"
        new_event.Start = event.Start
        new_event.End = event.End
        new_event.Location = "Reserved by Automation"
        new_event.BusyStatus = 2  # 2 = olBusy
        new_event.Body = f"Blocked due to event: {event.Subject}"
        new_event.Save()
        print(f"Booked Busy Time: {event.Start} - {event.End}")
    except Exception as e:
        print(f"Error creating event: {e}")


# Main Execution
print(f"\nSyncing events from {SOURCE_ACCOUNT} -> {TARGET_ACCOUNT} ({CALENDAR_NAME})")

source_calendar = get_calendar(SOURCE_ACCOUNT, CALENDAR_NAME)
target_calendar = get_calendar(TARGET_ACCOUNT, CALENDAR_NAME)

if not source_calendar:
    print(f"Source calendar not found for {SOURCE_ACCOUNT}. Exiting.")
    exit()
if not target_calendar:
    print(f"Target calendar not found for {TARGET_ACCOUNT}. Exiting.")
    exit()

# Fetch source events
source_events = get_events(source_calendar, START_DATE, END_DATE)

# Process each event
for event in source_events:
    try:
        event_start = event.Start.replace(tzinfo=None)
        event_end = event.End.replace(tzinfo=None)

        # Check for overlapping events
        if not event_overlaps(target_calendar, event_start, event_end):
            book_busy_time(target_calendar, event)
        else:
            print(f"Conflicting event already exists for: {event.Start} - {event.End}")
    except Exception as e:
        print(f"Error processing event: {event.Subject}, Error: {e}")

print("\nSyncing complete.")
