# department-course-scheduling
This python project creates course schedules for deparments based on deparment's need, time requirements, and faculty's preferences. The goal is to find the optimal schedules that both fulfill the requirements and maximize the preferences. It finds the optimal schedule for course assignment (match professors with courses), then using the schedule to find time assignment (scheduled class with time slots).

Installation
-
Download the files and install packages using requirements.txt.

Dependencies
-
* pandas 1.0.4
* gspread 3.6.0
* oauth2client 4.1.3
* ortools 7.6.7691


Date Input
-
The input can be either a google sheets or a local Excel file.
* Google sheets are access via their links. 
To set up the access to a google sheets, read the instruction to get a google credential [here](https://cloud.google.com/docs/authentication/getting-started) and download a .json file. Put the .json file under the project directory and rename it to 'client_secret.json'. Make sure to change the share setting of the sheets to be either 'Anyone with the link', or shared it with the 'client email' in the google credential.
* The Excel file need to be put under the project directory and named into 'scheduling_info.xlsx'.


Input Format
-
The Excel file and the google sheets need to follow an exact format since that's how the read_input method process information. The format is explained below, ane the Excel file in this repo can be used as an example.

The boolean values are written as 1 or 0. 
The integers values must be intergers and can not be floats.
Times are written in 24 hours system (e.g. 14:05:00)
The time frames Morning, Afternoon, and Evening, are separabted by 12 pm and 5 pm using start time of a time slot.

The input file should contain five sheets.
* CanTeach: Booleans for whether a professor is capable of teaching a class. Course names as columns names and Professor names as row names. 
* Prefer: Booleans for whether a professor prefer to teach a class. Same format as CanTeach.
* Course: Courses names as row names. The 'Unit' column lists the number of units of the course. Semester names are the following column names. Each lists how many sections of that course are wanted for the semester. Each semester must have a corresponding 'must offer' column listing how many sections must be offered, with the names being the semester names followed by'_MustOffer'. (i.e. a 2 under 'must offer' and a 4 under the semester column mean the course must have 2 sections, but up to 4 sections can be offered.) The 'Lab' column contains boolean of whether this course is a lab course.
* Prof: The 'MaxUnit' column lists the maximum units a professor would teach for the school year. Each combination of weekdays from the time slots have a column (MWF, TR, etc.), with list of string values of 'Morning, Afternoon, Evening' representing the professor's prefor time frames.
* Time: Five columns for every weekdays, containing booleans values of whether the time slot is on the days. The 'START TIME' and 'END TIME' columns lists its start time and end time. The 'Lab' column contains boolean of whether this is a lab time slot.
