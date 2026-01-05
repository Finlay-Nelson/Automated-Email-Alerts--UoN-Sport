# -*- coding: utf-8 -*-
"""
Created on Wed Jan 29 12:28:12 2025

Library dependencies:
  - pip install pandas
  - pip install openpyxl
  - pip install smtplib
  - pip install schedule

Notes:
    Usage
    -  There may be adjustments required should the template on which the script was based is changed (out of my control)
    -  A properly configured email account is required that allows programmed sending. Currently configured for GMail and requires a unique activation key (not listed here)
    -  Must have appropriate permissions on the WiFi network to which the PC running the script is connected
    -  The full address of the spreadsheet needs to specified (ideally a local access route to prevent permission issues)
    -  Recipient email addresses should be specified in __main__ and are currently configured to complete for a University of Nottingham address
    -  It's assumed all recipients have a University of Nottingham email address
    -  The script scales easily enough to all clubs etc. but until rolled out fully it is in its single-recipient format
    -  The email sign off needs updated as appropriate before executing

    Proposed additions
    - Add a clause that sends an email to unaffected parties rather than not providing any update at all (reassuring that something hasn't been missed or forgotten)
    - Automated rather than manual execution at pre-defined intervals 
    - Regular comparison to most recent file for "immediate" email updates (say daily) in addition to the regular distribution (currently weekly)
    - Condense extract_start_time() and extract_end_time() into a single function with arguments "start" and "end" that define the characters to isolate
  
"""
# Handles spreadsheets
import pandas as pd
# Actual sending of the email
import smtplib
# Let's you sort by date and format as need
from datetime import datetime, timedelta
# Email formatting 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# General utility functions
import os

def extract_end_time(value):
    """
    Reformat the "session time" column to identify the end time 
    Add this as a separate column in the modified BBT dataframe
    The existing NoPlay dates spreadsheet doesn't separate start and end times
    This simplifies creating time objects. 
    NOTE: It is assumed 24 time has been used in the original
    """
    # Remove anything that isn't a number
    digits = ''.join(filter(str.isdigit, value))
    # Check there are enough digits to carry out the task
    if len(digits) >= 4:
        # Format as HH:MM 
        return f"{digits[-4:-2]}:{digits[-2:]}"
      
    return None

def extract_start_time(value):
    """
    Reformat the "session time" column to identify the start time 
    Add this as a separate column in the modified BBT dataframe
    The existing NoPlay dates spreadsheet doesn't separate start and end times
    This simplifies creating time objects. 
    NOTE: It is assumed 24 time has been used in the original
    """
    # Remove anything that isn't a number
    digits = ''.join(filter(str.isdigit, value))
    # Checks there's enough numbers to do what it wants
    if len(digits) >= 4: 
        # Format as HH:MM
        return f"{digits[:2]}:{digits[2:4]}"  
    
    return None

def send_email(to_address, body, subject="Upcoming Disruption Notification"):
    """
    Sends the NoPlay information to a club's designated recipient'
    """
    # The email address sending the email (update as required)
    from_address = 'email_address'
    # This code allows emails to be sent from the above address programatically
    # It should be stored securely to prevent misuse (update as required)
    password = 'password'
    msg = MIMEMultipart("alternative")
    # General admin. Recipient, sender, and subject information
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    # Adds the emails contents
    msg.attach(MIMEText(body, 'html'))
    # Specifics of sending the email through Google. 
    # Alternatives methods exist but are subject to network restrictions etc.
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    try:        
        # Establish a connection to the appropriate server
        # Authenticate using tusername and password for the senders email
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=120)
        # Authenticate
        server.login(from_address, password)  
        # Compose the email
        text = msg.as_string()
        # Send the emaol
        server.sendmail(from_address, to_address, text)
        # Close connection to the server
        server.quit()  
        print(f'Email sent to {to_address}')
    except Exception as e:
        print(f'Failed to send email: {e}')
        
def download_NoPlay(filename = r"filename",notice = 2)
    """ 
    The file path here needs to not be hardcoded.
    I've parsed it into the degrees of severity
    - my local address followed by the path within the UoN Sport OneDrive.
    Compiles the relevant NoPlay date data within a given timeframe
    """
    # File path to the NoPlay dates spreadsheet. 
    # This is a hardcoded local route on my PC
    # Be careful of any additional whitespace at the end of the filename
    path1 = r"path1"
    path2 = r"path2"
    filepath = os.path.join(path1, path2, filename)
    # Advance notice to be given to recipients. Define the window period
    # Can be specified in days if desired (weeks is the default)
    now = datetime.today()
    then = now + timedelta(weeks=notice)
    # Load the NoPlayDate spreadsheet
    df = pd.read_excel(filepath)
    # Separate columns so we can look at them individually
    df.columns = df.columns.str.strip()
    # Remove the "Term" column as it's not really relevant
    df.drop(columns=['Term'], inplace=True)
    # Convert "Event date" to a consistent format
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    # Convert "Session time" to a string
    df['Session time affected'] = df['Session time affected'].astype(str).str.strip()
    # List any instance where "Event date" is within our disruption window
    upcoming_disruptions = df[(df['Date'] >= now) & (df['Date'] <= then)].copy()
  
    return now, then, df, upcoming_disruptions
    
def sort_chronologically(disruptions):
    """
    Takes the list of disruptions and formats for usability and ease of access
    Sorts chronologically to more easily identify the window of interest
    Returns a list of affected clubs and, unsorted, their disruptions
    """
    # Format event date. %Y %m %d makes it easier to sort chronologically
    disruptions['Date'] = disruptions['Date'].dt.strftime("%Y %m %d")
    # Remove any none numeric characters from the "Session time affected"
    tmp = disruptions['Session time affected'].str.replace(r'[ :.\-]', '', regex=True)
    disruptions['Session time affected'] = tmp
    # Replace the original "Time" column with "Start" and "End" time columns
    disruptions['Start time'] = disruptions['Session time affected'].apply(extract_start_time)
    disruptions['End time'] = disruptions['Session time affected'].apply(extract_end_time)
    disruptions.drop(columns=['Session time affected'], inplace=True)
    # Rearrange columns with club use in mind (i.e. club, date, time, ...)
    disruptions = disruptions.reindex(columns=[
        'Club / Programme Name', 
        'Date',
        'Day affected',
        'Start time',
        'End time',
        'Location', 
        'Hours Lost',
        'Alternative time offered / available?',
        'Clash Event'])
    # Sort disruptions in chronological order
    disruptions = disruptions.sort_values(by='Date').reset_index(drop=True)
    # List of affected clubs making sure there are no repeats
    # There are ways of accounting for capitalisation 
    # Naming/spelling needs to be consistent
    # Clean the 'Club / Programme Name' column first
    disruptions['Club / Programme Name'] = disruptions['Club / Programme Name'].str.lower().str.strip()
    affected_clubs = disruptions['Club / Programme Name'].unique()
    # Print terminal summary of the time period being assessed
    then_str = then.strftime('%d %m %Y')
    print(f"No Play dates checked until {then_str}")
    
    return affected_clubs, disruptions
    
def process_noplay(disruptions, club, coach, then):
    """
    Works through all disruptions and attributes them to their clubs
    Composes and returns an email for each club with their NoPlay information
    """
    date, start, end, location, day, change = [], [], [], [], [], []
    # Go through each disruption checking if it applies to this club
    for _, row in disruptions.iterrows():
        if club == row['Club / Programme Name']:
            # Add to the "Day affected" column (Mon, Tue, etc.)
            day.append(row['Day affected'])
            # Add to the "Date affected" column (dd/mm for readability)
            date_str = row['Date']
            date_obj = datetime.strptime(date_str, "%Y %m %d")
            formatted_date = date_obj.strftime("%d/%m")
            date.append(formatted_date)
            # Add to the "Location", "Start", and "End" time columns
            location.append(row['Location'])
            start.append(row['Start time'])
            end.append(row['End time'])
            # Add to the "Change" column (NA if nothing is in place)
            change.append('None' if pd.isna(row['Alternative time offered / available?']) else row['Alternative time offered / available?'])
    # Create an object for each club with its NoPlay information
    combined = list(zip(date, start, end, location, day, change))
    combined_sorted = sorted(combined, key=lambda x: (x[0], x[1]))
    date, start, end, location, day, remedy = zip(*combined_sorted)
    # Print the club name to the terminal
    print(f"{club}")
    # Create an html table of the clubs disruptions - just formatting
    table_html = """
    <table border="1" 
        cellpadding="4" 
        cellspacing="0" 
        style="border-collapse: collapse;">
    <tr>
    <th>Date</th>
    <th>Day</th>
    <th>Time</th>
    <th>Location</th>
    <th>Change
    </tr>
    """
    # Populate the table with NoPlay information
    for date,start,end,location,day,change in combined_sorted:
        table_html += f"""
        <tr>
        <td>{date}</td>
        <td>{day}</td>
        <td>{start}-{end}</td>
        <td>{location}</td>
        <td>{change}</tr>"""
    table_html += "</table>"
    # Some contextual stuff before the NoPlay information
    then = then.strftime('%d/%m')
    intro_html = f"""
    <h2>Schedule Disruptions for {club}</h2>
    <p>Hi {coach},</p>
    <br>
    <p>Hope everything is well on your end. Here's the no play information for {club} until {then}:</p>
    <br>
    """
    # Some niceties after the NoPlay information
    outro_html = """
    <br>
    <p>Have a nice weekend!</p>
    <br> Thanks,
    <br> Insert Name
    """
    full_html = intro_html + table_html + outro_html
  
    return full_html
##############################################################################
        
if __name__ == "__main__":   
  
    """Load the file to find and list upcoming disruptions"""
    # We have the time window, original spreadsheet and filtered version
    now, then, df, disruptions = download_NoPlay()
    """Look at the disruptions, reformat, and create a chronological list"""
    # Make more club friendly and make a list of clubs affected
    affected_clubs, disruptions = sort_chronologically(disruptions)
    """Collect each clubs information and email to the designated recipient"""
    # Iterate through each club over the specified window
    if affected_clubs.size != 0:
        for club in affected_clubs:
            """Just applying this to one club for the time being"""
            if club == "club":
                coach = "coach"
                matID = "institutionID"
                # This could all get un-indented when scaled up to more clubs
                to_address = f"{matID}@exmail.nottingham.ac.uk"
                # Create a table for the club's NoPlay information
                table_html = process_noplay(disruptions, club, coach, then)
                # Send an email with the table included
                send_email(to_address, table_html)
              
