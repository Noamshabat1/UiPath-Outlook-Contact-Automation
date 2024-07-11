
# Automated Process Using UiPath Studio

This project involves building a workflow in UiPath Studio that automatically adds contacts to Outlook based on an attached Excel file. The development is done in the .NET environment, specifically using VB.NET.

## Process Overview

The workflow includes two solutions:
1. The Excel file is open during the automation process.
2. The Excel file is closed during the automation process.

For each contact, the process will:
- Save their full name, email address, and phone number in Outlook.
- Perform a search for the contact on Google (on the website itself). The search results are not significant for this process.
- Update the contact card in Outlook in the Job title field with one of the following options:
  - "Does not exist in Google"
  - "The phone in Google is compatible"
  - "The phone in Google is not compatible"

The Flow will select the correct option based on the comparison between the data in columns D and E in the Excel file. At the end of handling each contact, the current date and time will be entered in the "Treated B" cell.

## Commands Used

The following UiPath commands are used to accomplish the tasks:
- `Use Application/Browser`
- `Click`
- `Keyboard Shortcuts`
- `Type Into`
- `Get From Clipboard`
- `If`
- `Assign`
- `Terminate Workflow`
- `Excel Application Scope`
- `Read Cell`
