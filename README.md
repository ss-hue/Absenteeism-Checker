# Absenteeism-Checker


This file contains the main and only code for this tool used in a professional context to check the agents that are absent comparing two excel/csv files downloaded from the main project tool.

This application aims to automate and save precious time for the Real Time Analyst (RTA), that previous to this was used to perform this task manually.

The two main input files that receive this applications are:

1. File containing all agents by their IDs and each respectively schedule for the week. It is required to download and set this file in a weekly basis.

2. File containing the timestamp and ID of each and every agent that is successfully logged in at the time of the downloading. 

This two files then are compared to each other using each agent ID to check whether the agent is logged in or not (absent). 
