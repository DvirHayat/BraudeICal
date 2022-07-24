from icalendar import Calendar,Event ,Timezone,TimezoneDaylight
import pandas as pd
from re import sub
from datetime import datetime
from pathlib import Path
import os
import pytz

#Semester Begin and Ends Dates in [y,m,d] format
#Enter the day before the semester starts 
sem_begin = [2022,10,22]
sem_end = [2023,1,22]

#Enter path to xlsx file
braude_path=r"C:\Users\Dvir\Downloads\YedionXlsFile_07691_01398.xlsx"

#Get day For Reaccurance format
def getDay(char):
    if(char=="א"):
        return "SU"
    elif(char=="ב"):
        return "MO"
    elif(char=="ג"):
        return "TU"
    elif(char=="ד"):
        return "WE"
    elif(char=="ה"):
        return "TH"
    elif(char=="ו"):
        return "FR"

#initiating Data Frame
braude_df=pd.read_excel(braude_path,skiprows=[0])

#Creating ical file
cal = Calendar()

# Calendar properties required to be compliant
cal.add('version', '1.0')
cal.add('prodid', '//My courses schadule//')
tz = Timezone(TZID='Israel')
cal.add_component(tz)
tzdl=TimezoneDaylight(TZOFFSETFROM="+0200",TZOFFSETTO="+0300",TZNAME="IDT",DTSTART="19700327T020000")
cal.add_component(tzdl)

#Loop of data extracting and event creation
for i in braude_df.index:
    #Getting course data from Data Frame
    lect=braude_df["מרצה"].loc[i]
    day=braude_df["יום"].loc[i]
    course_name=braude_df.loc[i,"שם נושא"]
    course_type=braude_df.loc[i,"סוג מקצוע"]
    course_code=braude_df.loc[i,"קוד נושא"]
    room = braude_df.loc[i,"חדר"]
    
    #Creating the course name event headline 
    if ("הערה") in course_name:
        corr_course=sub(r" ?\([^)]+\)", "",course_name)
    else:
        corr_course=course_name
    subject_line=(f"{corr_course} - {course_type} ({course_code})")
    room=braude_df["חדר"].loc[i]
    
    #Extracting time in str from DF file
    start=datetime.strptime(str(braude_df.loc[i,"מ-שעה"]),"%H:%M")
    start=start.strftime("%H:%M")
    end=datetime.strptime(str(braude_df.loc[i,"עד-שעה"]),"%H:%M")
    end=end.strftime("%H:%M")
    
    #Creating events for each course in ICal File 
    ev = Event()
    ev.add('summary',subject_line)
    ev.add('dtstart', datetime(sem_begin[0], sem_begin[1], sem_begin[2], int(start[0:2]), int(start[3:6]), 0, tzinfo=pytz.timezone("Israel")))
    ev.add('dtend', datetime(sem_begin[0], sem_begin[1], sem_begin[2], int(end[0:2]), int(end[3:6]), 0, tzinfo=pytz.timezone("Israel")))
    ev.add('dtstamp',datetime(sem_begin[0], sem_begin[1], sem_begin[2], int(start[0:2]), int(start[3:6]), 0, tzinfo=pytz.timezone("Israel")))
    ev.add('description',lect)
    ev.add('location',room)
    until_date = datetime(sem_end[0], sem_end[1], sem_end[2], tzinfo=pytz.timezone("Israel")) 
    ev.add('rrule',{'freq':'weekly','interval':1,'byday':getDay(day),'until':until_date})
    cal.add_component(ev)

# Write to disk
directory = Path.cwd() / 'MyCalendar'
try:
   directory.mkdir(parents=True, exist_ok=False)
except FileExistsError:
   print("Folder already exists")
else:
   print("Folder was created")
 
f = open(os.path.join(directory, 'מערכת שעות - בראודה.ics'), 'wb')
f.write(cal.to_ical())
f.close()


