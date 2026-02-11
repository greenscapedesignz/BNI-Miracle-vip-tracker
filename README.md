# BNI-Miracle-vip-tracker

A Google Apps Script web application for tracking VIP referrals for BNI presenters.

## Features
- Automatically shows presenters for upcoming Wednesday meetings
- Tracks VIP information and referrals
- Stores data in Google Sheets

## Setup Instructions

1. Create a new Google Spreadsheet
2. Name a sheet "FP Log" with columns:
   - A: Meeting Months
   - B: Date of Meeting
   - C: On/Off Line
   - D: Presenter Name
   - E: Business Type
   - F: Contact Info

3. Go to Extensions > Apps Script
4. Delete the default Code.gs content
5. Copy the contents of `Code.js` into Code.gs
6. Create a new HTML file called `index.html`
7. Copy the contents of `index.html` into it
8. Deploy as Web App (see Deployment section)

## Deployment

1. Click Deploy > New Deployment
2. Select "Web app" as deployment type
3. Set "Execute as" to "Me"
4. Set "Who has access" to "Anyone"
5. Click Deploy
6. Copy the web app URL and share with your team

## Files
- `Code.js` - Server-side Google Apps Script code
- `index.html` - Frontend HTML/CSS/JavaScript
