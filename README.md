# BotGoogleSheetsAPI

![Example Image](User/Desktop/Google sheets.jpg)

Telegram bot for uploading xlsx files to Google Sheets via Google Sheets API
## Functioning
[Token_API.Py](https://github.com/IvanZaitsevSPb/BotGoogleSheetsAPI/blob/main/Token_API.py)

>Token_API - take from father bot

[Google_func.py](https://github.com/IvanZaitsevSPb/BotGoogleSheetsAPI/blob/main/google_func.py)

>SPREADSHEET_ID - ID sheets where to insert

[credentials]()

>Google Cloud > API & Services > Credentials > Download  JSON file OAuth 2.0

[requirements.txt](https://github.com/IvanZaitsevSPb/BotGoogleSheetsAPI/blob/main/requirements.txt)

>Libraries for installation are specified

To install from a file, use the command
>pip install -r requirements.txt

**Launching the bot via a file [main.py](https://github.com/IvanZaitsevSPb/BotGoogleSheetsAPI/blob/main/main.py)**

## Description

The bot performs the function of converting and filling the Google Sheets table with a file or xlsx files.

### Input data

You can send one or more xlsx files to the bot, check the API connection in your personal account, add or remove a column or row in the table via the bot

### Output data

Output of a link to the table to be filled in, log output after loading "file .xlsx loaded successfully"

### Commands

>*/start* - getting to know the bot

>*/push_table* - Sends data to the table (functional after the scent documents)

>*/delete_xlsx* - deletes all xlsx files in the root folder

>*/check_API* - displays the site for checking on/off the Google Sheets API

### Interaction algorithm

1) You send the documents, the bot notifies you with a reply letter that the files have been saved
2) Check */check_API*, the bot sends in response a link to the developer's personal account in which you check whether the API is enabled
3) */push_table* sending data to the table, waiting for a reply letter with a link to the table
4) After successful download, delete all xlsx files from the root folder */delete_xlsx*