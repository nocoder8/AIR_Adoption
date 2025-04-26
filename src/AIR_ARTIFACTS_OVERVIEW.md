## AI Reporting and Alerting Artifacts Overview

This document provides an overview of the Google Apps Scripts and Looker Studio dashboards used for AI screening reporting, alerting, and adoption tracking.

### 1. `AIR_DailySummary.js` (Google Apps Script)

-   **Purpose:** Automatically generate and distribute a daily email report summarizing company-wide AI screening activity and recruiter adoption metrics.
-   **Key Metrics Calculated:** Total invitations sent, completion rates (overall and adjusted for recent invites < 48hrs old), average time from invitation to completion (using schedule date as proxy), average match stars (for completed interviews).
-   **Breakdowns Provided:** Metrics are broken down by Recruiter, Job Function, and Location Country.
-   **Adoption Metrics:** Calculates recruiter-specific AI adoption rates based on eligible candidates (post-launch, >= configured match score) using data from a separate Application sheet (if available).
-   **Recruiter Activity Tracking:** Shows the last date each recruiter sent an invitation and provides a 10-day trend of their daily invitation counts.
-   **Data Processing:** Filters out specified test positions, deduplicates candidate records per position (prioritizing most recent status), uses data from `Log_Enhanced` and potentially an `Active+Rejected` sheet.
-   **Output:** Generates an HTML email report with tables and KPI boxes.
-   **Distribution:** Sends the email daily (around 10 AM) to configured recipients and sends error notifications if issues occur.
-   **Time Range:** Currently configured to analyze data across all time (`VS_REPORT_TIME_RANGE_DAYS_RB = 99999`).

### 2. `Recruiter_alerts_AIR.js` (Google Apps Script)

-   **Purpose:** Proactively notify recruiters about specific candidates requiring timely action in the AI screening process and provide an administrative overview of pending actions.
-   **Primary Alert Trigger:** Identifies candidates whose screening is 'COMPLETED' but feedback status is 'AI_RECOMMENDED' for more than 1 day and less than or equal to 7 days.
-   **Recruiter Alert Email:** Sends a personalized email to the responsible recruiter (`Creator_user_id`) listing these candidates, highlighting those pending > 3 days as urgent, and prompting review within 24 business hours.
-   **Contextual Information (Recruiter Email):** Includes counts of 'NEW' candidates with >=4 match score awaiting review (from Application sheet) and candidates 'PENDING' > 2 days who may need a nudge.
-   **Admin Digest Email:** Compiles and sends a separate email to an administrator summarizing pending 'AI_RECOMMENDED' items per recruiter and listing all candidates currently meeting the alert criteria across the organization.
-   **Data Processing:** Uses data from the `Log_Enhanced` sheet, deduplicates records per position (prioritizing most recent status), filters out specified test positions and specific candidate names ('Erica Thomas'), and potentially uses the `Active+Rejected` sheet for contextual counts.
-   **Output & Links:** Generates HTML emails with candidate details and includes a link to the relevant Looker Studio dashboard for a broader view.
-   **Distribution:** Runs automatically Tuesday-Friday (around 1 PM), sending individual emails to recruiters (CC'ing admin) and the digest email to the admin. Includes error reporting.

### 3. AI Recruiter: EF4EF Dashboard (Looker Studio Dashboard - Page 1)

-   **Purpose:** Provide a high-level, real-time (updated every 30 mins) overview of the AI screening funnel performance and status.
-   **Key KPIs Displayed:** Total AI Invitations sent, average time (in days) from Invitation to Completion.
-   **Invitations vs. Completions Trend:** Line chart showing daily invitations sent alongside the number of those specific invitations that eventually resulted in a completed screening (bar chart).
-   **Status Distribution:** Donut chart and table showing the current count and percentage breakdown of interviews across various statuses (PENDING, COMPLETED, SCHEDULED, etc.).
-   **Daily Completion Trend:** Line chart tracking the number of AI screenings actually *completed* each calendar day, highlighting weekends.

### 4. AIR Adoption Opportunities - Candidates List (Looker Studio Dashboard - Multi-Page)

-   **Overall Purpose:** Offer recruiters detailed, actionable lists to manage candidates at specific, critical stages of the AI screening pipeline.
-   **View: Pending Recruiter Review/Submission:**
    -   Lists candidates whose AI screening is 'COMPLETED' and feedback is 'AI_RECOMMENDED'.
    -   Prompts recruiters to review AI feedback and action the candidate (positive/negative) within 24 business hours.
    -   Shows candidate details, recruiter, HM, completion date, and a summary count per recruiter.
-   **View: Follow-up List (Invited but yet-to-complete):**
    -   Lists candidates who received an invitation but haven't completed the screening.
    -   Displays 'Days Pending' to help recruiters identify who may need a follow-up nudge.
-   **View: Candidates to Invite (New Candidates):**
    -   Lists 'NEW' candidates with a high match score (>=4) who haven't been invited for AI screening yet.
    -   Prompts recruiters to initiate the screening for these potential fits.
-   **View: Summary:**
    -   Provides a pivot table summarizing candidate counts across different interview statuses (PENDING, COMPLETED, SCHEDULED, etc.) broken down by individual recruiter (`Creator_user_id`).
    -   Offers a quick overview of each recruiter's pipeline status. 