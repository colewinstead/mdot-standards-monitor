# MDOT Standards Monitor

This Windows utility checks the MDOT Engineering Standards/Guides/Manuals page every weekday. It renders the JavaScript page in Chrome, compares meaningful page content and links, and SHA-256 hashes every directly linked MDOT document. When something changes, it sends an HTML summary through the signed-in Classic Outlook profile.

For modified files, the report also identifies PDF page numbers and text excerpts, Word paragraph changes, Excel sheet/cell changes, and text-file line changes. Modified PDF pages include low-resolution before/after previews, including drawings and scanned pages with no extractable text.

Temporary rendering or download failures are retried with exponential backoff. A detected change is checked a second time after two minutes before anyone is notified. The monitor also keeps a local change-history dashboard and sends a quiet weekly health summary so the team knows monitoring is still operational.

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
py -3.14 .\monitor.py local-test --open
py -3.14 .\monitor.py status
py -3.14 .\monitor.py history --open
py -3.14 .\monitor.py config show
py -3.14 .\monitor.py update --check-only
py -3.14 .\monitor.py update
.\install.ps1 -Remove
```

- `check` performs the normal comparison and sends an alert only when needed.
- `check --dry-run` performs a live comparison but changes no state and sends no email.
- `initialize` deliberately replaces the baseline without sending an update alert.
- `send-test` sends a clearly labeled test message.
- `send-preview` sends a clearly labeled fictional example of a change report.
- `local-test --open` creates a temporary MDOT-like page and documents, changes them, and
  opens the resulting HTML alert. It exercises Chrome rendering, downloads, hashes, PDF
  page comparison, change detection, and report formatting without Outlook, Task Scheduler,
  network access, configured recipients, or changes to the saved baseline.
- `status` reports the baseline, last success/change/error, run duration, Chrome and Outlook
  availability, history count, and Windows scheduled-task status. Add `--json` for automation.
- `history --open` builds and opens the local event dashboard. Change reports contain embedded
  previews and remain readable even after old page-image cache entries are cleaned up.
- `update --check-only` safely checks GitHub. `update` refuses a dirty or diverged repository,
  tests the remote commit in an isolated worktree, and only then fast-forwards the current branch.
- `install.ps1 -Remove` removes the scheduled task but retains configuration and history.

## Configuration commands

The common settings can be changed without manually editing JSON:

```powershell
py -3.14 .\monitor.py config validate
py -3.14 .\monitor.py config add-recipient engineer@example.com
py -3.14 .\monitor.py config remove-recipient engineer@example.com
py -3.14 .\monitor.py config set-failure-recipient owner@example.com
py -3.14 .\monitor.py config set-heartbeat-days 7
py -3.14 .\monitor.py config set-confirmation-delay 120
```

Global filters limit which documents are downloaded and monitored. Values are case-insensitive;
section and title values are substring matches:

```powershell
py -3.14 .\monitor.py config set-filter --section Roadway --extension pdf --extension xlsx
py -3.14 .\monitor.py config set-filter --title "Design Manual"
py -3.14 .\monitor.py config clear-filters
```

Recipient routes send a filtered report to an additional group. Global recipients still receive
the complete report. To prevent duplicate messages, an address already in the global recipient
list is omitted from routed batches:

```powershell
py -3.14 .\monitor.py config add-route "Bridge team" `
  --recipients "bridge1@example.com;bridge2@example.com" --section Bridge --extension pdf
py -3.14 .\monitor.py config remove-route "Bridge team"
```

Use `--include-page-text` on `add-route` if that routed group should also receive general page-text
changes. Reusing a route name replaces that route.

## Data and troubleshooting

Configuration, the current baseline, low-resolution current-page previews, runtime health, local
history reports, and rotating logs are stored in:

```text
%LOCALAPPDATA%\MDOTStandardsMonitor
```

The configuration never contains an Outlook password. Review `monitor.log` there if a run fails.
Change alerts go to the applicable recipient lists. Check failures, immediate Outlook send errors,
and the weekly health summary go only to the configured failure recipient. The team does not receive
routine health messages. Unchanged daily checks remain silent. Set `heartbeat_days` to zero through
the configuration command to disable the weekly summary.

Existing installations can pull updates in place. The scheduled task, configuration, baseline, and
history remain intact; do not rerun `install.ps1` merely to update the program.

The monitor follows changes on the MDOT page and hashes directly linked files hosted by MDOT. It records the presence, title, and URL of third-party links but does not crawl or hash external websites.

One ProjectWise-hosted PDF, `RWD Workflow Training ORD`, is dynamically resolved from the RWD CADD Manual on every run. The resolver accepts exactly one link matching the trusted ProjectWise host and download path, so a changed ProjectWise key is followed automatically while every other embedded link remains ignored. Multiple matching destinations cause a safe failure instead of an ambiguous selection.

## Tests

```powershell
py -3.14 -m unittest discover -s tests -v
```
