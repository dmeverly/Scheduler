# Scheduler

**Author**: David Everly  
**Language**: Python  
**Version**: 1 

---

## Description  
Script reads a schedule from template.xlsx and writes an employee schedule based using the template on a standard monthly calendar. Client requested this project to support 3 shifts per day with 2 on dayshift and 1 on nightshift.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Features](#features)
- [Configuration](#configuration)
- [Examples](#examples)
- [Future Work and Extension](#future-work-and-extension)
- [References](#references)
- [Contributing](#contributing)
- [Licenses](#licenses)

## Installation
Dependencies:   
pandas
openpyxl

# Usage
Program is intended to be run using Unix-like terminal such as Linux, macOS Terminal (untested), or MINGW64 (Git Bash) on Windows.  

Run the script with: 
```bash 
python scheduler.py
```    
Or use the provided shell script:    
```bash
./run  
```

## Features  
At runtime, queries the user on command line for starting template week number, starting month number, and starting year.  Output to Schedule.xlsx a formatted schedule

## Configuration  
Pre-processing details are hardcoded and specific to the format defined in Template.xlsx.  Future modification must either keep with the current format, or extend this script to support alternative formats.

## Examples  

```bash
$ ./run
Template Week Number: 1
Starting Month Number: 7
Starting Year: 2025
Schedule saved as Schedule.xlsx
```

<iframe src="https://docs.google.com/spreadsheets/d/e/2PACX-1vRhSfov48lHD9mZk3m05FcPiqS9fAAVw-penkR9oDgX4RjbmHX2TzpdzqAl9daO_F8v2RHPXxekbIQP/pubhtml?widget=true&amp;headers=false"></iframe>

## Future Work and Extension  
This project is part of a larger initiative, which was to generate a rotating template and then automate transcription of the template onto a monthly calendar.  See templater.py for more details.  The natural next step is to integrate the programs for a complete, end-to-end program which generates a template meeting employee constraints and produces a schedule with minimal user input.  This is best achieved by extending the scheduler.py script to take dataframe input, rather than .xlsx input, and map the calendar from there.

## References  
No external sources were used. However, LLM queries assisted with architectural design and debugging.  

## Contributing  
No external parties contributed to this project.  

## Licenses  
None

<a href="https://www.dmeverly.com/completedprojects/Scheduler/" style="display: block; text-align:right;" target = "_blank">  Project Overview -> </a> 