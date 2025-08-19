# Google Apps Script Project
EK Klinkman

## Overview

This repository contains the code for a Google Apps Script (GAS) project. The project is designed to merge with the user's Google Calendar and pull information from calendar blocks to generate a time report. 

In its current form, the script identifies relevant calendar blocks/categories if they are labeled with colors found in the GCal "Time Breakdown" menu. You will need to hard-code your time category labels in the Apps Script code according to GCal color indexing. The code is currently configured to identify 8 different subcategories of task (`colorIdMap`) and bin these 8 subcategories into 3 main groups (`majorGroups`). 

The script accepts a custom YYYY-MM-DD time range (`timeMin`, `timeMax`) and sums the hours for each sub- and main category on separate Google sheet tabs. The project is designed to be linked to a Google Sheet and the user's email address (permission popups will likely occur) and be run within the Google Apps Script environment to generate the report.


## Troubleshooting

1. You will likely need to authorize the project by providing permission for the script to access your data.

![authorization](https://github.com/user-attachments/assets/db536c78-593d-4427-8ae8-4b80d9e1981a)

Allow the project to access your Google Account: press ‘Allow’.

## Contact
   If you have any questions or need further assistance, feel free to reach out:
   * Name: Emily Klinkman, MS.
   * Email: emilykk@umich.edu
   * GitHub: https://github.com/emklinkman
