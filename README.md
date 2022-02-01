# Qualtrics-Staff-Info

This project will handle API calls to import essential staff information from Active Directory into Qualtrics.

# Main Components
1. Job Scheduler - service that currently acts as a user agent on remote server that can be interacted through a web interface
2. AD_Pull.ps1 - PowerShell script that retrieves all the essential staff data from Active Directory
3. qualtricsStaffInfo.py - Python script that pulls all the contacts from a specified Qualtrics directory and compares it to the Active Directory list 

# Basic Walkthrough
1. Configure Job Scheduler to schedule two tasks on a daily basis. 
2. Job Scheduler runs AD_Pull.ps1 at 6:10 am to retrieve the most up-to-date staff information from Active Directory. This export file is stored onto a shared network drive.
3. Job Scheduler then runs qualtricsStaffInfo.py at 6:15 am to see which contacts need to be added, removed, or modified.
![image](https://user-images.githubusercontent.com/87395701/151998201-25346edf-d2fc-47d8-b274-9c8eb53dfc4f.png)
