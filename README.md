# tado_smart_thermostat
PowerShell script that uses the Google Sheets v4 API and Tado v2 API to set your tado° Smart AC Control

1) Set up a Google Sheets document (https://docs.google.com/spreadsheets/) like this: https://i.imgur.com/bLD7IHz.png

2) Make Google Sheet public (only required because I'm not familiar with setting up oauth authentication. I'm sure it can be done with a private sheet)

3) Set up a Google Developer project (https://console.developers.google.com/flows/enableapi?apiid=sheets.googleapis.com) to create an api credential (https://console.developers.google.com/apis/credentials)

4) Update the information in the 'script_setup' function in the .ps1 file with the updated information

5) Set up a scheduled task to run every minute (or as often as you'd like) to fetch/update your desired settings
