# BraudeICal
Python script that translates excel file from my collage website to an ical file  

The only way a schadule is made downlodable from my collage website with days, hours, lecturer name, room name, etc. is through an Excel spreadsheet. This lead me to create a 'translator' that will help me and hopefully other students, to manage their time more efficently. 

This script provides more suitable way to make a course schadule through creating an Ical file that includes the courses as repeated events throughout the semester. 

In order to run the script there are 3 Python libraries that are needed to be installed:

-icalander

-pandas

-pytz

These can be installed using the pip installer : "pip install LibName" 

After that, your Excel spreadsheet should be in the correct format provided by the college website (https://info.braude.ac.il/) , as shown in schadule_format.xlxs


1)Enter the dates in the list in which the semesters are starting and ending.

2)Change the file path to your Excel spreadsheet file.

The script will make the Ical file in local repository under:"My Calander"
