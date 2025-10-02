# UOB Assignment Exporter

This unpacked extension fetches user group and submission data from the UOB CMS and exports the submission details as a formatted report (HTML or PDF).

## Install

1. Open Chrome and navigate to `chrome://extensions/`.
2. Enable **Developer mode** (top right).
3. Click **Load unpacked**, then choose the `uob-assignment-exporter` folder from this repository.

## Use

1. Ensure you are logged in to `https://cms.uobmydigitalspace.com` in the same browser profile.
2. Open the extension popup in Chrome.
3. The extension automatically detects the group/structure serials from the active CMS tab and fetches data when the popup opens. If detection fails, set fallback values in `DEFAULT_SERIALS` inside `popup.js`.
4. Choose the desired export format (HTML or PDF) and click **Export Report**:
   - **HTML** downloads a standalone `.html` file with one row per student, showing latest score/feedback alongside inline previews, clickable links, and collapsible accordions when multiple submissions exist.
   - **PDF** opens the report in a print-friendly tab so you can use the browser print dialog (`⌘P`/`Ctrl+P`) to save it as PDF (allow pop-ups for the extension).

If fetch requests fail, confirm the serial values and that your login session is still active.
