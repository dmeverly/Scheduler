# Scheduler

## Description
Creates and exports a .xlsx work calendar based on provided .xlsx template

Project sponser is tasked with creating an annual employee schedule based on a rotating template
Manually transcribing employee shift assignments from template to calendar was time-consuming and a waste of company resources
This solution automates the process of employee shift assignment transcription, saving time, money, and mental frustration

## Usage
Scheduler.py is a standalone script which depends upon Template.xlsx (template). On execution, the template queries the user
to input the desired starting week number on the rotating template, starting month, and starting year.  The script creates
file Schedule.xlsx which is the transcribed employee shift calendar by month for 12 consecutive months.
