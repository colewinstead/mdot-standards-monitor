# MDOT Standards Monitor

This Windows utility checks the MDOT Engineering Standards/Guides/Manuals page every weekday. It renders the JavaScript page in Chrome, compares meaningful page content and links, and SHA-256 hashes every directly linked MDOT document. When something changes, it sends an HTML summary through the signed-in Classic Outlook profile.

For modified files, the report also identifies PDF page numbers and text excerpts, Word paragraph changes, Excel sheet/cell changes, and text-file line changes. PDF page snapshots include a low-resolution visual fingerprint so changes to drawings or scanned pages can still be assigned to a page even when no text can be extracted.

## Install

1. Confirm Classic Outlook has a working signed-in profile. Chrome and Outlook may remain open during checks.
2. Open PowerShell in this folder.
3. Run:

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\install.ps1
   ```

4. Enter team email addresses or a distribution-list address separated by semicolons.
5. Choose whether to send a test email. No test is sent by default.

For unattended installation, recipients can be supplied directly. A test email is still omitted unless `-SendTest` is added:

```powershell
.\install.ps1 -Recipients 'team@example.com;other@example.com'
```

The installer creates the first baseline without sending an update alert and registers **MDOT Standards Daily Monitor** for 12:00 PM Monday through Friday. The task runs only while the installing user is logged on, starts after a missed schedule, and asks Windows to wake the computer when possible.

## Commands

```powershell
py -3.14 .\monitor.py check
py -3.14 .\monitor.py check --dry-run
py -3.14 .\monitor.py initialize
py -3.14 .\monitor.py send-test
py -3.14 .\monitor.py send-preview
.\install.ps1 -Remove
```

- `check` performs the normal comparison and sends an alert only when needed.
- `check --dry-run` performs a live comparison but changes no state and sends no email.
- `initialize` deliberately replaces the baseline without sending an update alert.
- `send-test` sends a clearly labeled test message.
- `send-preview` sends a clearly labeled fictional example of a change report.
- `install.ps1 -Remove` removes the scheduled task but retains configuration and history.

## Data and troubleshooting

Configuration, the current baseline (including detailed page/content snapshots), failure state, and rotating logs are stored in:

```text
%LOCALAPPDATA%\MDOTStandardsMonitor
```

The configuration contains only the page URL and recipient addresses—never an Outlook password. Review `monitor.log` there if a run fails. Change alerts go to the team recipient list. A check failure or immediate Outlook send error generates a separate failure notice only to the configured failure recipient; successful unchanged checks and recoveries remain silent.

The monitor follows changes on the MDOT page and hashes directly linked files hosted by MDOT. It records the presence, title, and URL of third-party links but does not crawl or hash external websites.

## Tests

```powershell
py -3.14 -m unittest discover -s tests -v
```
