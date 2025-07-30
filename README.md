# Microsoft Outlook Automated

## Description

A python tool helps me automating the daily boring stuff like saving emails' attachments and archiving emails
from different senders
> [!Note]
> This tool is still under development and testing so that in some lines there is a hard coded conditions and folder names, there will be a user inputs to enter and select different things

## Pywin32 Package

Using **win32com** lib from the [pywin32](https://pypi.org/project/pywin32/) package to dispatch Outlook using MAPI (Messaging Application Programming Interface) over HTTP protocol so that we can exchange connections between our tool and Outlook app

## Future Features

- [ ] Save emails' attachments dependant on user inputs
- [ ] Read each attachment, parse its content and gather useful data
- [ ] Use gathered data to make weekly useful Excel sheets
