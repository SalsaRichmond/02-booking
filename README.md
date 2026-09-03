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

## Recommended Workspace Actions

1. **Deploy `Code.gs` in Google Apps Script**:
   - Open your Google Sheet linked to the questionnaire.
   - Go to **Extensions > Apps Script**, paste the updated [`Code.gs`](file:///c:/Users/Angel%20A%20Rodriguez/.gemini/antigravity-ide/scratch/questionnaire-new-generation-2026/Code.gs), and click **Save**.
   - Click **Deploy > Manage Deployments**, edit your active Web App deployment (or create a **New Deployment**), ensure access is set to **"Anyone"**, and click **Deploy**.

2. **Verify Notifications & Email Permissions**:
   - In Google Sheets, refresh the page to load the `💃 Salsa App` menu.
   - Click **`📧 SEND TEST NOTIFICATION EMAIL`** to verify instant email delivery to `salsaguyrichmond@gmail.com` and inspect remaining daily quota.
   - Click **`🔔 SETUP AUTO-SUBMISSION TRIGGER`** if you also accept submissions directly through Google Forms or Google Sheets.
   - Click **`🔍 DIAGNOSE NOTIFICATION SYSTEM`** to run a comprehensive connectivity check across Gmail, Drive, Calendar, and Sheets.


