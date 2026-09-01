# SALSA GUY RICHMOND LLC — SYSTEM AUDIT & RECOMMENDATIONS
### Strategic Improvements for Booking Questionnaire, Google Apps Script Pipeline & Master Google Docs

* **Target Audience:** Angel A. Rodriguez (The Salsa Guy / Profe)
* **Subject Codebase:** Event Booking 2026 (`public/index.html`, `Code.gs` v20.71, Google Drive)
* **Current Live URL:** [https://sgr-booking-2026.tradicion.workers.dev](https://sgr-booking-2026.tradicion.workers.dev)
* **Audit Scope:** Frontend form flow, backend JSON field mapping (`Code.gs` lines 1015–1080), header synchronization, template document population, and mobile conversion.

---

## 1. Fix Missing Backend Mappings in `Code.gs` (Critical Data Loss Prevention)
**Priority:** High (Bug Fix) — `[PRIORITY: HIGH]`

When examining lines 1015 to 1080 of `Code.gs` (`handleFormSubmitJson`) alongside the frontend submission payload in `public/index.html`, we identified multiple fields where the customer enters valuable information on the website, but the backend either lacks a matching field map or maps to the wrong key:

* **Missing `outOfStateLogistics` Mapping:** The frontend collects detailed out-of-state flight/hotel logistics in `formData.outOfStateLogistics`, but `fieldHeaderMap` in `Code.gs` does NOT contain an entry for `outOfStateLogistics`. When out-of-state clients submit lodging and per diem requirements, this data is never saved into the Google Sheet columns.
* **Duplicate `eventPurpose` vs `eventDescription`:** In `public/index.html` (line 3763), `eventPurpose` is assigned to `document.getElementById('eventDescription').value`. Meanwhile, the actual event purpose field is bypassed. `Code.gs` expects `eventPurpose` and `eventDescription` separately.
* **Missing `serviceTypeRequested` Mapping:** The customer picks *"Performance Only"*, *"Dance Instruction Only"*, or *"Both"*, yet `fieldHeaderMap` lacks an explicit alias mapping for `serviceTypeRequested` to save the primary service choice in its own dedicated sheet column.
* **Repertoire Individual Category Storage:** The form collects all 21 repertoire checkboxes into a single concatenated string in `performanceServices`. However, `Code.gs` `setupMasterHeaders` defines separate columns for *Repertoire: Mexico*, *Repertoire: Caribbean*, *Repertoire: Central America*, *Repertoire: South America*, and *Repertoire: Theatrical*. Splitting them into their regional categories during submission ensures the Google Sheet and Review Doc reflect clean categorical grouping.

---

## 2. Add LocalStorage Auto-Save & Draft Recovery
**Priority:** High (Conversion & User Experience) — `[PRIORITY: HIGH]`

The 2026 booking questionnaire is comprehensive (covering 6 thorough sections from logistics to hospitality). Prospective event hosts often need time to look up their venue address, budget confirmation, or schedule before completing the form.

* **Real-Time Input Caching:** Automatically save client inputs to browser `localStorage` on every keystroke or selection. If a client accidentally closes their mobile tab or navigates away to check venue details, their answers are instantly restored upon return.
* **Draft Recovery Banner:** Display a subtle notification: *"We restored your previous event draft. [Clear Draft]"* so users know their progress was saved.
* **Submission Clearance:** Clear the local cache only after `handleFormSubmitJson` returns a confirmed Request ID.

---

## 3. Pre-Submission Date Availability Indicator (Connected to Google Calendar)
**Priority:** Medium (User Experience & Scheduling) — `[PRIORITY: MEDIUM]`

Currently, the website mentions the general availability rule (Fridays–Sundays for shows, Mondays–Thursdays for dance lessons). However, clients can still submit requests for dates where Salsa Guy already has a booked performance in InfoCalendar.

* **Real-time Availability Status:** Expose a lightweight Apps Script endpoint (`doGet` with `action=checkAvailability&date=YYYY-MM-DD`) that queries InfoCalendar without revealing private event details.
* **Instant Visual Feedback:** When a date is selected, show:
  * `✅ Date Appears Available`
  * `⚠️ Date Has Existing Bookings (Inquire for Alternative Hours)`
  * `ℹ️ Weekday: Available for Instruction/Workshops Only`
* **Reduces Back-and-Forth:** Saves hours of email correspondence by setting clear scheduling expectations right at the moment of questionnaire fill-in.

---

## 4. Dynamic Price Range Calculator (Instant Budget Transparency)
**Priority:** Medium (Sales Velocity & Lead Qualification) — `[PRIORITY: MEDIUM]`

Event organizers frequently hesitate when selecting "Confirmed Budget Amount" because they do not know standard rates for Latin dance performances, live Parranda caroling, or sound system provisioning.

* **Dynamic Estimate Widget:** Based on service type (Solo instruction vs. 4-person troupe vs. full Parranda ensemble) and duration (1 hour vs. half-day), provide an estimated ballpark investment range right above the budget field.
* **Filters Unqualified Leads:** Educates clients on professional performance rates before they submit unrealistic numbers (e.g. $100 for a 6-performer Mexican folkloric production).
* **Custom Package Tiering:** Add quick-select tier buttons: *"Standard Cultural Showcase"*, *"Festival Headliner Experience"*, or *"VIP Masterclass + Performance"*.

---

## 5. Automated SMS / WhatsApp Notifications for Profe & Client
**Priority:** Medium (Response Speed) — `[PRIORITY: MEDIUM]`

Currently, `Code.gs` sends automated HTML emails via `sendAdminSubmittalNotification` and `sendClientReceiptNotification`. While email works well, booking inquiries are time-sensitive.

* **Instant SMS / WhatsApp Ping for Angel:** Integrate a webhook (e.g., Twilio or WhatsApp Business API via Google Apps Script `UrlFetchApp`) sending a 1-sentence notification:
  > *"💃 New Booking Request! Maria Santos on Oct 14 for $1,500 at Richmond Center. Tap to view Proposal Doc."*
* **Fast 15-Minute Response Time:** Event planners often book the first vendor that responds. Rapid notification gives Salsa Guy a massive competitive advantage.

---

## 6. Master Google Docs Template Modernization (Proposal, Contract, Perf Info)
**Priority:** High (Branding & Close Rate) — `[PRIORITY: HIGH]`

The automation suite generates three customized Google Docs in Drive for every lead using templates: Proposal, Contract, and Performance Information. Our inspection of `Code.gs` lines 1980–2015 revealed opportunities to elevate these documents:

* **Automated Dynamic Repertoire Table:** Instead of dumping raw checkbox text, generate a styled, formatted table in the Proposal and Performance Info Doc listing each selected dance (e.g., *El Baile de los Viejitos*, *Bomba*, *Alma Ranchera*) along with performer count and required stage dimensions.
* **Google Doc PDF Auto-Export & Direct Download Link:** In addition to generating the editable Google Doc, have `processRow` generate a finalized read-only PDF using `doc.getAs('application/pdf')` and save it to the event's Drive folder. This link can be immediately sent to the client in their receipt email.
* **Digital Signature Readiness:** Add an interactive signing block placeholder with date and IP stamp to prepare the contract document for e-signature workflows.

---

## 7. Complete Multi-Language Field Synchronization (Spanish, French, German, etc.)
**Priority:** Medium (Global Reach & Inclusivity) — `[PRIORITY: MEDIUM]`

The frontend has an impressive language switcher supporting English, Spanish, French, German, etc. However, following the reorganization into 6 sequential sections, some translation strings in `langDict` still reference legacy section numbers (e.g., *"sec6: 7. Technical..."*, *"sec5: 8. Budget..."*). Synchronizing these dictionary keys guarantees seamless multi-lingual presentation for international and cultural embassy clients.

---

## Executive Summary & Action Roadmap

| # | Recommendation | Impact Area | Target Timeline |
| :--- | :--- | :--- | :--- |
| **1** | **Fix `Code.gs` Field Mappings & Aliases** | Data Integrity & Sheets Logging | **Phase 1 (Immediate)** |
| **2** | **LocalStorage Draft Auto-Save & Recovery** | User Experience & Lead Recovery | **Phase 1 (Immediate)** |
| **3** | **Calendar Availability Real-Time Check** | Client Scheduling Accuracy | **Phase 2 (Next Sprint)** |
| **4** | **Dynamic Price Range Ballpark Widget** | Lead Qualification & Conversion | **Phase 2 (Next Sprint)** |
| **5** | **SMS / WhatsApp Instant Host Notifications** | Lead Response Time (< 15 mins) | **Phase 3 (Expansion)** |
| **6** | **Master Doc PDF Auto-Export & Styled Tables** | Proposal Presentation & Close Rate | **Phase 2 (Next Sprint)** |
| **7** | **Complete 6-Section i18n Translation Sync** | Global & Cultural Client Reach | **Phase 2 (Next Sprint)** |

---

*Prepared with dedication to the continued excellence and growth of Salsa Guy Richmond LLC & Tradición Puerto Rican Folk Dancing.*
