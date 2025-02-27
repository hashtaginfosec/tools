##############################################################
# Syncs Outlook calendars between two accounts. 
# Update lines 10 - 12 with the account and calendar names. 
# You must have Outlook installed and be logged in with both accounts.
##############################################################
# Create an Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Configuration: Source and Target Account and Calendar Names
$SourceAccount = "user@workemail.com"     # Account to read events from
$TargetAccount = "user@personalemail.com"  # Account to block "Busy" time on
$CalendarName = "Calendar"           # Calendar folder name in both accounts

# Set the date range for the next month
$StartDate = (Get-Date).Date
$EndDate = $StartDate.AddMonths(1)

# Function to get calendar folder by account and calendar name
function Get-CalendarFolder($AccountName, $CalendarName) {
    foreach ($Account in $Namespace.Folders) {
        if ($Account.Name -eq $AccountName) {
            try {
                $CalendarFolder = $Account.Folders.Item($CalendarName)
                if ($CalendarFolder) {
                    return $CalendarFolder
                }
            } catch {
                Write-Host "Error accessing calendar: $CalendarName for account: $AccountName" -ForegroundColor Red
            }
        }
    }
    return $null
}

# Get Source and Target Calendars
$SourceCalendar = Get-CalendarFolder -AccountName $SourceAccount -CalendarName $CalendarName
$TargetCalendar = Get-CalendarFolder -AccountName $TargetAccount -CalendarName $CalendarName

if (-not $SourceCalendar) {
    Write-Host "Source calendar not found for $SourceAccount. Exiting." -ForegroundColor Red
    exit
}
if (-not $TargetCalendar) {
    Write-Host "Target calendar not found for $TargetAccount. Exiting." -ForegroundColor Red
    exit
}

# Get events from Source Calendar
function Get-EventsFromCalendar($CalendarFolder) {
    $Items = $CalendarFolder.Items
    $Items.Sort("[Start]")
    $Items.IncludeRecurrences = $true
    
    # Filter appointments within the specified date range
    $Filter = "[Start] >= '" + $StartDate.ToString("g") + "' AND [End] <= '" + $EndDate.ToString("g") + "'"
    $Appointments = $Items.Restrict($Filter)
    
    return $Appointments
}

# Normalize DateTime for accurate comparison
function Normalize-DateTime($DateTime) {
    return [datetime]::SpecifyKind($DateTime, [datetimekind]::Utc)
}

# Check if an event overlaps in the target calendar
function EventOverlaps($TargetFolder, $Start, $End) {
    $Items = $TargetFolder.Items
    $Items.Sort("[Start]")
    $Items.IncludeRecurrences = $true

    # Adjusted filter to check for overlapping events
    $Filter = "([Start] < '" + $End.ToString("g") + "') AND ([End] > '" + $Start.ToString("g") + "')"
    $MatchingItems = $Items.Restrict($Filter)
    
    # Compare with a buffer of ±1 minute to avoid millisecond mismatches
    foreach ($Match in $MatchingItems) {
        $MatchStart = Normalize-DateTime $Match.Start
        $MatchEnd = Normalize-DateTime $Match.End
        
        $Buffer = [TimeSpan]::FromMinutes(1)
        if (
            ($MatchStart - $Buffer -le $Start -and $MatchEnd + $Buffer -ge $End) -or
            ($MatchStart -le $Start -and $MatchEnd -ge $End) -or
            ($MatchStart -ge $Start -and $MatchStart -le $End)
        ) {
            Write-Host "Overlap Found: [$($Match.Subject)] ($($Match.Start) - $($Match.End))" -ForegroundColor Yellow
            return $true
        }
    }
    return $false
}

# Book busy time on Target Calendar
function Book-BusyTime($TargetFolder, $Event) {
    $NewAppointment = $TargetFolder.Items.Add([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)
    $NewAppointment.Subject = "Busy"
    $NewAppointment.Start = $Event.Start
    $NewAppointment.End = $Event.End
    $NewAppointment.Location = "Reserved by Automation"
    $NewAppointment.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olBusy
    $NewAppointment.Body = "Blocked to avoid conflict with event: $($Event.Subject)"
    $NewAppointment.Save()
    Write-Host "Booked Busy Time: $($Event.Start) - $($Event.End)" -ForegroundColor Green
}

# Get Events from Source Calendar
Write-Host "`nReading events from Source Calendar: $SourceAccount -> $CalendarName" -ForegroundColor Yellow
$SourceEvents = Get-EventsFromCalendar -CalendarFolder $SourceCalendar

# Loop through each event and book busy time on Target Calendar
foreach ($Event in $SourceEvents) {
    try {
        # Normalize start and end times to UTC
        $EventStart = Normalize-DateTime $Event.Start
        $EventEnd = Normalize-DateTime $Event.End
        
        # Check for overlapping events with ±1 minute buffer
        if (-not (EventOverlaps -TargetFolder $TargetCalendar -Start $EventStart -End $EventEnd)) {
            # Book the busy time
            Book-BusyTime -TargetFolder $TargetCalendar -Event $Event
            # Write-Host "Event found for: $($Event.Start) - $($Event.End)" -ForegroundColor Yellow
        } else {
            Write-Host "Conflicting event already exists for: $($Event.Start) - $($Event.End)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error processing event: $($Event.Subject)" -ForegroundColor Red
    }
}

Write-Host "`nCompleted booking Busy times on $TargetAccount -> $CalendarName calendar." -ForegroundColor Cyan
