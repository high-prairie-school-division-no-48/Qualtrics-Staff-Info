# Qualtrics-Staff-Info

This project will handle API calls to import essential staff information from Active Directory into Qualtrics.

# Main Components
1. Job Scheduler - Service that currently acts as a user agent that is used to run scripts on a user configured schedule
2. AD_Pull.ps1 - PowerShell script that retrieves all the essential staff data from Active Directory
3. qualtricsStaffInfo.py - Python script that retrieves all contacts from a specified Qualtrics directory and compares their details to their entry in the Active Directory list 

# Basic Walkthrough
1. Configure Job Scheduler to schedule two tasks to run on a daily basis. 
2. Job Scheduler runs AD_Pull.ps1 at 6:10 am to retrieve the most up-to-date staff information from Active Directory. This export file is stored onto a shared network drive.
3. Job Scheduler then runs qualtricsStaffInfo.py at 6:15 am to see which contacts need to be added, removed, or modified.

Fields:
- First name
- Last name
- Email
- Employee ID (used as unique identifier when cross-referencing between Qualtrics and Active Directory)
- Primary location
- Job title
- Description (any additional info extracted from Active Directory)
![image](https://user-images.githubusercontent.com/87395701/151998201-25346edf-d2fc-47d8-b274-9c8eb53dfc4f.png)

# Notes
Job Scheduler
- To access network drive, need to initially sign in locally with an Active Directory account with proper permissions. This is so the job agent can then act as that user and have proper access to the network drive and the required files/directories.
- Requires Java to run
- Can be interacted through a web interface
- Currently installed on remote server
