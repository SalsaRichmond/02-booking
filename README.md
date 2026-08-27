# Questionnaire New Generation for 2026 - Salsa Guy Richmond, LLC

**Project Directory:** `C:\Users\Angel A Rodriguez\.gemini\antigravity-ide\scratch\questionnaire-new-generation-2026`  
**Revision:** REV 20.71 (Custom HTML Form & Direct Web App Integration)  
**Date:** August 27, 2026  

## Project Overview
This project contains the **Next-Generation 2026 Event Booking Suite** for **Salsa Guy Richmond, LLC**. It combines a state-of-the-art responsive HTML5 booking form (`index.html`) with an automated Google Apps Script Web App backend (`Code.gs`).

## Key Components

1. **`index.html`**:
   - Modern, responsive HTML5 booking form designed with rich Latin-inspired aesthetics (crimson red, gold accents, glassmorphic cards, Google Fonts `Outfit` & `Inter`).
   - Dynamic form validation, instant submission, and interactive success modal with generated **`Event ID`**.

2. **`Code.gs`**:
   - Google Apps Script automation suite connecting to Google Sheets, Google Calendar, and Google Drive.
   - `doGet(e)` & `doPost(e)` handlers for serving the Web App and processing HTTP POST submissions in < 2 seconds.
   - `updateMasterTemplates()`: Automated tool to clean up, update, and standardize placeholders across the 3 Master Google Docs (Proposal, Contract, Performance Info).
   - Multi-tiered document generator matching exact headers, short tags, and legacy question phrasing.

## Master Document Templates

- **Proposal Master Doc**: [`1plCZvjBJijgJrGduzXrTMjbtDopMgCfB5Vo7MLWxspo`](https://docs.google.com/document/d/1plCZvjBJijgJrGduzXrTMjbtDopMgCfB5Vo7MLWxspo/edit)
- **Contract Master Doc**: [`1BuEv7BF6wsHutvEWwVZOVb3m8J3zkYujKheti8X871g`](https://docs.google.com/document/d/1BuEv7BF6wsHutvEWwVZOVb3m8J3zkYujKheti8X871g/edit)
- **Performance Info Master Doc**: [`1eyXzMdmYiV0CDmxvNZfNO3dRYQ83hvJHHy_eseWPqIE`](https://docs.google.com/document/d/1eyXzMdmYiV0CDmxvNZfNO3dRYQ83hvJHHy_eseWPqIE/edit)

## Recommended Workspace Action
Copy `Code.gs` into your Google Apps Script project bound to the Google Spreadsheet, and run `📄 UPDATE MASTER DOC TEMPLATES` from the `💃 Salsa App` custom UI menu!

