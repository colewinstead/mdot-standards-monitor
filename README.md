# MDOT Standards Monitor

This Windows utility checks the MDOT Engineering Standards/Guides/Manuals page every weekday. It renders the JavaScript page in Chrome, recursively expands the website's collapsed folder grids, compares the folder inventory, meaningful page content, and links, and SHA-256 hashes every discovered MDOT document. When something changes, it sends an HTML summary through the signed-in Classic Outlook profile.

The website sections **Construction** and **Construction Materials** (the latter is represented by the top-level `Materials` category in the site's data grid) are intentionally excluded, including all of their descendant folders and documents. All other current and newly added folders are included automatically. Folder additions and removals are reported separately; document additions, removals, renames, moves, and content modifications use the existing detailed change report.

For modified files, the report also identifies PDF page numbers and text excerpts, Word paragraph changes, Excel sheet/cell changes, and text-file line changes. Modified PDF pages include low-resolution before/after previews, including drawings and scanned pages with no extractable text.

Temporary rendering or download failures are retried with exponential backoff. Retries are scoped to the failed page render, dynamic link resolution, or individual document, so one unavailable file does not restart every completed download. A detected change is checked a second time after two minutes before anyone is notified. The monitor also keeps a local change-history dashboard and sends a quiet weekly health summary so the team knows monitoring is still operational.

If one document remains unavailable after all retries, the monitor retains its last good snapshot and continues checking and reporting changes in every other document. The configured failure recipient receives a separate private warning; document errors are never included in the team change email.

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

The monitor follows changes on the MDOT page and hashes files exposed by the recursively expanded MDOT folder grids. It records the presence, title, and URL of third-party links but does not crawl or hash external websites.

One ProjectWise-hosted PDF, `RWD Workflow Training ORD`, is dynamically resolved from the RWD CADD Manual on every run. The resolver accepts exactly one link matching the trusted ProjectWise host and download path, so a changed ProjectWise key is followed automatically while every other embedded link remains ignored. Multiple matching destinations cause a safe failure instead of an ambiguous selection.

## ProjectWise Explorer folder monitoring (planned)

This feature is **not implemented yet**. The existing ProjectWise resolver above follows one public
download link; it cannot sign in to a private ProjectWise datasource or enumerate a ProjectWise
Explorer folder.

The intended feature is read-only monitoring of one or more ProjectWise Explorer folders. It should
use the signed-in employee's approved Windows/ProjectWise SSO session and recursively inventory the
configured folders. For each visible document, retain enough information to detect additions,
removals, renames, new versions, metadata changes, and file-content changes. Feed confirmed changes
through the existing email formatting, recipient routing, history dashboard, retry, and failure-alert
systems. Creating the first ProjectWise baseline must not send a change alert.

Do not monitor ProjectWise's local `pwworkdir` cache. That directory contains only documents copied
to the computer and can therefore miss new or untouched documents in the datasource. Do not store a
ProjectWise username, password, access token, or other secret in this repository or in `config.json`.
The integration must not check out, check in, upload, rename, move, delete, or change the workflow
state of any ProjectWise document.

### Work-computer discovery checklist

Perform these read-only checks on the work computer before changing the program:

1. Pull this repository with `git pull --ff-only` and run the existing unit tests.
2. Record the installed ProjectWise Explorer version, Python version, PowerShell version, datasource
   display name, and target folder path as shown in ProjectWise Explorer.
3. Determine whether an organization-approved ProjectWise SDK, API, or PowerShell module is already
   installed. In PowerShell, `Get-Module -ListAvailable pwps_dab` is a harmless initial check. Do not
   install a module or request elevated rights without confirming company policy.
4. Determine whether the user's normal SSO session can perform a read-only folder/document query.
   Never pass a password on a command line. ProjectWise Explorer's `pwc.exe` arguments can open a
   datasource or select a folder, but that alone is not proof that folder contents can be enumerated.
5. Confirm that the query works while connected through the network or VPN conditions under which
   the Windows scheduled task will run.

### Implementation requirements

- Add a separate `projectwise_sources` configuration collection. Each source should have a friendly
  name, datasource identifier, ProjectWise folder path or stable folder ID, recursive flag, and
  optional title/extension filters.
- Add configuration commands equivalent to `config add-projectwise-folder`,
  `config remove-projectwise-folder`, and `config list-projectwise-folders`. Final arguments should
  be based on what the installed ProjectWise tooling actually supports; the command names are
  proposals, not currently available commands.
- Keep the MDOT website baseline and every ProjectWise source baseline independent so adding or
  repairing a ProjectWise source cannot reset MDOT history.
- Prefer stable ProjectWise document and folder IDs over display paths when available. Preserve the
  display path in reports so people can understand where a change occurred.
- Use read-only metadata first. Download or copy out a file only when a content hash or detailed
  comparison is required, place temporary copies under the monitor's local application-data folder,
  and clean them safely after use.
- Re-query the ProjectWise source during change confirmation. If a source is unavailable, preserve
  its last good baseline and notify only the configured failure recipient rather than reporting all
  documents as removed.
- Support `check --dry-run`, `status`, history reports, global filters, and recipient routes for
  ProjectWise changes without changing their current MDOT behavior.
- Add mocked unit tests that run without ProjectWise, plus an explicit read-only integration-test
  command for the work computer. The integration test must never send email or modify the baseline.
- Verify Task Scheduler execution under the same Windows user and SSO context. If ProjectWise access
  requires an interactive logged-on session, retain the installer's current logged-on-user task
  behavior and document that limitation.

### Prompt for tomorrow's Codex session

After opening this repository in Codex on the work computer, use this prompt:

> Continue the planned ProjectWise Explorer folder-monitoring work described in README.md. First run
> only the read-only discovery checklist and inspect the installed ProjectWise tooling. Do not install
> software, request elevation, store credentials, or perform any ProjectWise write operation. Report
> what datasource-query method is available and propose the exact configuration shape and integration
> test. If a safe read-only query works, implement the adapter with mocked tests and a no-email,
> no-baseline-write integration test. Preserve all existing MDOT monitoring behavior.

## Tests

```powershell
py -3.14 -m unittest discover -s tests -v
```
