# Print Request App — Design Document

Status: **Draft for review — no implementation yet.**

## 1. Goal & Scope

Replace the current Outlook/Excel automation script — which parses free-form
requester emails — with a Windows desktop app that **is** the submission
tool. The app collects a structured print request from the requester,
computes/validates page counts per file, and on submit:

1. Sends a single Outlook email (with the original files attached) to the
   person who does the printing.
2. Writes a row per attached file to a shared Excel log.
3. If the Excel write fails for any reason, sends a fallback email with the
   full data so a human can reconcile it into Excel manually — without
   blocking or failing the user's submission.

### Out of scope (for this version)

- No request history/dashboard inside the app — the shared Excel file *is*
  the system of record.
- No authentication — the app trusts the identity of the Windows/Outlook
  user it runs as.
- No editing of previously submitted requests.
- No installer/deployment tooling design (covered separately later).
- No support for legacy binary `.ppt`/`.doc` (pre-2007) formats beyond what
  is noted explicitly in §6.3.

## 2. Tech Stack

| Concern | Choice |
|---|---|
| UI framework | WPF, **.NET 8 (net8.0-windows)** |
| MVVM | CommunityToolkit.Mvvm (source-generator based, MIT license, no heavy dependency) |
| Outlook automation | `Microsoft.Office.Interop.Outlook` (COM interop) |
| Word page counting | `Microsoft.Office.Interop.Word` (COM interop) |
| PDF page counting | PdfPig (see §6.1) |
| PPTX slide counting | DocumentFormat.OpenXml (Open XML SDK) |
| Excel writing | ClosedXML (see §6.4) |
| Config | JSON file (`appsettings.json`) next to the executable |
| Logging | Serilog, rolling file sink under `%LOCALAPPDATA%` |

.NET 8 is chosen over .NET Framework 4.8 for current tooling/support; Office
COM interop works the same way under .NET 8 via the
`Microsoft.Office.Interop.*` PIAs (registered as COM references, embedded
interop types). Target machines are assumed to have the .NET 8 Desktop
Runtime (or the app is published self-contained — a deployment detail to
confirm later, not blocking this design).

## 3. Project Structure

```
PrintRequestApp.sln
src/
  PrintRequestApp/                    WPF UI project (net8.0-windows)
    App.xaml / App.xaml.cs            FlowDirection, culture, DI bootstrap
    Views/
      MainWindow.xaml                 Request form + attachment list
      ConfirmationDialog.xaml         Review → Sending → Result states
    ViewModels/
      MainViewModel.cs
      AttachmentItemViewModel.cs
      ConfirmationViewModel.cs
    Controls/                         Reusable RTL-aware controls (labeled field, yes/no toggle)
    Resources/
      Styles.xaml                     Fonts, colors, control templates
      Strings.xaml                    Centralized Hebrew UI strings (see §7)
    appsettings.json                  Placeholder config (see §5)

  PrintRequestApp.Core/                Class library (net8.0), no WPF dependency
    Models/
      PrintRequest.cs
      AttachmentItem.cs
      ColorMode.cs / FileKind.cs (enums)
      SubmissionResult.cs
    Services/
      PageCounting/
        IPageCounter.cs
        PdfPageCounter.cs
        WordPageCounter.cs
        PptxSlideCounter.cs
        PageCounterFactory.cs
      Outlook/
        IOutlookService.cs
        OutlookEmailService.cs
      Excel/
        IExcelRequestWriter.cs
        ExcelRequestWriter.cs
      Submission/
        RequestSubmissionService.cs   Orchestrates send → excel → fallback
      Configuration/
        AppSettings.cs
        AppSettingsProvider.cs

  PrintRequestApp.Tests/               xUnit tests for Core (models, counters
                                        against sample files, submission
                                        orchestration with faked services)
```

**Why a Core class library separate from the WPF project:** all COM
interop, page counting, Outlook, and Excel logic is UI-agnostic and needs to
be unit-testable without spinning up WPF. The WPF project only holds
Views/ViewModels and wires Core services in via simple constructor
injection (a lightweight DI container, e.g. `Microsoft.Extensions.
DependencyInjection`, set up in `App.xaml.cs`).

### Threading note for Office COM

Office COM objects are single-threaded apartment (STA). The WPF UI thread
is already STA, so short COM calls can run there, but page-counting a batch
of Word documents or sending mail should not block the UI. The design runs
these operations on a dedicated background thread that is explicitly
created with `ApartmentState.STA` (not the default MTA thread pool), with
the ViewModel awaiting a `TaskCompletionSource` bridged from that thread.
The UI shows a busy indicator during these calls (see §8).

## 4. Data Model

### `PrintRequest` (one per submission, request-level fields)

| Field | Type | Hebrew label | Notes |
|---|---|---|---|
| `ProgramName` | string | שם התכנית | required |
| `BudgetLine` | string | סעיף תקציבי | required |
| `CopiesCount` | int | מספר עותקים | required, ≥ 1 |
| `HolePunch` | bool | חירור | default false |
| `DoubleSided` | bool | דו צדדי | default false |
| `Stapling` | bool | הידוק | default false |
| `HasCoverPage` | bool | יש/אין דף פתיח | default false |
| `BlankPageBetweenCoverAndContent` | bool | האם להשאיר עמוד ריק בין דף פתיחה לתוכן | only meaningful/enabled when `HasCoverPage == true` |
| `SlidesPerPage` | int? | (אם נשלח כ‑PPT) כמה שקפים בעמוד | only required/shown when at least one attachment is a PPT/PPTX |
| `Attachments` | `List<AttachmentItem>` | קבצים מצורפים | required, ≥ 1 item |
| `SubmittedByDisplayName` / `SubmittedByEmail` | string | — | **not a form field.** Read automatically from the current Outlook session (`NameSpace.CurrentUser`) at submit time, used only for the Excel log / fallback email "requested by" column. No manual entry needed since Outlook already knows who is sending. |
| `SubmittedAtUtc` | DateTime | — | set at submission time |

Color/B&W is intentionally **not** on this object — it lives per file on
`AttachmentItem`, per the closed design decision in the brief.

### `AttachmentItem` (one per attached file)

| Field | Type | Notes |
|---|---|---|
| `FilePath` | string | full path to the original file on disk |
| `FileName` | string | derived, shown in UI |
| `FileKind` | enum `{ Pdf, Word, PowerPoint }` | derived from extension |
| `DetectedPageCount` | int? | result of auto-detection; `null` if detection failed or isn't supported for the file type |
| `DetectionSucceeded` | bool | drives the "לא זוהה אוטומטית" UI hint |
| `ManualPageCount` | int? | set when the user edits the count; `null` until touched |
| `EffectivePageCount` | int (computed) | `ManualPageCount ?? DetectedPageCount`; submission is blocked if this is null (user must supply a number when auto-detection fails) |
| `ColorMode` | enum `{ Color, BlackAndWhite }?` | **required, no default.** Starts unset so the user must make an explicit, visible choice for every file (backed by the "mark all" buttons for speed) rather than silently inheriting a default that could be wrong |

### `SubmissionResult` (returned by the submit flow, drives the result screen)

| Field | Type |
|---|---|
| `PrimaryEmailSent` | bool |
| `ExcelWriteSucceeded` | bool |
| `FallbackEmailSent` | bool |
| `ErrorMessage` | string? (populated on failures, for logging/user display) |

## 5. Configuration (`appsettings.json`)

Placed next to the executable, human-editable by IT without a rebuild:

```json
{
  "Recipient": {
    "PrimaryEmail": "PLACEHOLDER_PRINT_RECIPIENT@example.com"
  },
  "Excel": {
    "SharedFilePath": "PLACEHOLDER_UNC_PATH\\PrintRequests.xlsx",
    "WriteRetryAttempts": 3,
    "WriteRetryDelayMs": 500
  },
  "FallbackEmail": {
    "Address": "PLACEHOLDER_FALLBACK_RECIPIENT@example.com"
  }
}
```

Loaded once at startup by `AppSettingsProvider` into a strongly-typed
`AppSettings` object, injected wherever needed. All three placeholder
values are listed again in §11 (Open Questions) — nothing else in this
design depends on their real values.

## 6. Library Choices & Tradeoffs

### 6.1 PDF page counting — **PdfPig**

| Option | Verdict |
|---|---|
| **PdfPig** (chosen) | Pure managed C#, Apache 2.0, no native/COM dependency, actively maintained, reads page count cheaply without rendering. Read-only (fine — we only need page count). |
| PdfSharp | Primarily a PDF *creation* library; its reading support is weaker for malformed/encrypted/scanned PDFs, which real-world requester files will include. |
| iText7 | Very capable, but AGPL-licensed unless a commercial license is purchased — a licensing complication this internal tool doesn't need. |

### 6.2 Word page counting — Word COM Interop (required by brief)

Uses `Microsoft.Office.Interop.Word`: open the document, call
`Document.Repaginate()` then read
`Document.ComputedStatistics(WdStatistic.wdStatisticPages)` for a live,
accurate count (the cached `Document.ComputedStatistics` / built-in
properties can be stale for documents that haven't been repaginated since
last edit). Documents are opened with `Visible = false`,
`ReadOnly = true`, and are always closed and COM-released in a `finally`
block to avoid orphaned `WINWORD.EXE` processes accumulating on the
requester's machine.

### 6.3 PPTX slide counting — Open XML SDK (required by brief)

`DocumentFormat.OpenXml`: open the `.pptx` as a `PresentationDocument` and
count `PresentationPart.SlideParts`. No PowerPoint installation or COM
instance needed at all — fast and lightweight compared to the Word path.

**Caveat:** Open XML SDK only reads the modern `.pptx` (OOXML) format, not
legacy binary `.ppt`. Since Office is confirmed present on every target
machine, if legacy `.ppt` support turns out to be needed, the fallback is a
PowerPoint COM Interop path (`Application.Presentations.Open` +
`Slides.Count`), structurally parallel to the Word counter. Not built now
— flagged so it isn't a surprise later if a requester attaches an old
`.ppt`. In v1, `.ppt` attachment triggers the same "not auto-detected,
please enter manually" path as any unsupported file, per §4.

### 6.4 Excel writing — **ClosedXML**

| Option | Verdict |
|---|---|
| **ClosedXML** (chosen) | MIT license, no commercial-use restriction, straightforward high-level API for opening an existing workbook, appending rows, and saving. |
| EPPlus | Very capable but v5+ requires a paid commercial license (Polyform Noncommercial only free tier) — avoided to keep this dependency-free of licensing tracking. |
| NPOI | Free/Apache, viable alternative, but a more verbose, lower-level API for this simple append-rows use case. |
| Excel COM Interop | Would guarantee perfect fidelity with a hand-maintained template, but is slower, requires Excel installed and a visible/hidden instance per write, and is materially more fragile against a file that's locked or open by another user — exactly the failure mode this design needs to handle gracefully. ClosedXML reading/writing the file directly at the filesystem level is simpler to make resilient (see §8.2). |

### 6.5 Outlook automation — COM Interop (required by brief)

`Microsoft.Office.Interop.Outlook`. The service first tries to attach to an
**already-running** Outlook instance via
`Marshal.GetActiveObject("Outlook.Application")` (most users have Outlook
open), falling back to `new Application()` if none is running. This avoids
spawning a second competing Outlook process against the same mail profile.

**Deployment note (not a blocker for this design):** some environments
show Outlook's "a program is trying to send an email on your behalf"
security prompt for programmatic `.Send()` calls, depending on antivirus/
Trust Center configuration. This is an IT/environment configuration
concern to resolve at rollout, not something the app's design can control.

## 7. RTL / Hebrew Implementation

- **Global `FlowDirection = RightToLeft`** set once on `App.xaml`'s
  application-level style targeting `Window` (and explicitly on
  `MainWindow` and `ConfirmationDialog` roots as a belt-and-braces
  measure). WPF mirrors layout automatically under RTL: `Grid` columns
  render right-to-left, horizontal `StackPanel`s flow right-to-left, and
  `HorizontalAlignment`/`Dock` values are mirrored — so standard layout
  panels do not need manual RTL-specific positioning.
- **Culture**: `Thread.CurrentThread.CurrentCulture` /
  `CurrentUICulture` set to `he-IL` at startup (`App.xaml.cs`), and
  `FrameworkElement.LanguageProperty` default overridden to
  `XmlLanguage.GetLanguage("he-IL")` app-wide via a style, so WPF's
  built-in bidi text engine, date/number formatting, and spell-check (if
  ever enabled) all treat content as Hebrew rather than falling back to
  `en-US` defaults.
- **Font**: a font with solid Hebrew glyph coverage and correct bidi
  shaping — **Segoe UI** (ships with Windows, has full Hebrew coverage and
  is Microsoft's own UI font, so it renders identically to how the text
  would look in Outlook/Word on the same machine). Set once in
  `Styles.xaml` as the default `FontFamily` for `Window`/`Control`.
- **Mixed Hebrew/numeric fields**: pure RTL mirroring is not quite enough
  for fields like מספר עותקים (copies) or page-count inputs, where the
  *content* is a number but the *label* is Hebrew. Numeric `TextBox`
  controls are given `FlowDirection=LeftToRight` with
  `TextAlignment=Right` (frame stays right-aligned in the RTL layout like
  every other input, but the digits inside type and display in their
  natural left-to-right order — this is the standard fix for the classic
  "reversed numbers" bug in RTL number fields). Free-text fields (program
  name, budget line) keep the inherited RTL flow direction since they're
  Hebrew content.
- **Strings are centralized**, not scattered as inline XAML literals — a
  single `Resources/Strings.xaml` resource dictionary holds every label,
  button caption, and message. The UI only ever needs Hebrew for this
  version, so this isn't built as a multi-language resx/localization
  system; it's centralized purely so every piece of Hebrew copy (labels,
  validation messages, confirmation text, email subject/body templates)
  lives in one reviewable place instead of being duplicated across
  Views/ViewModels/services.
- **Testing for correctness**: manual verification pass on real Hebrew
  content mixed with numbers/punctuation (e.g. copy counts, page counts,
  budget line codes that may contain Latin letters/digits) in the main
  form, the attachment grid, the confirmation dialog, and the composed
  Outlook email body/subject — checking specifically for reversed digits,
  misplaced punctuation, and correct alignment, since these are the
  classic bidi bugs and automated tests won't catch rendering issues.

## 8. Screens

### 8.1 Main Window

Single window, `FlowDirection=RightToLeft`, roughly:

```
┌───────────────────────────────────────────────┐
│                בקשת הדפסה חדשה                  │
├───────────────────────────────────────────────┤
│  פרטי הבקשה                                     │
│  שם התכנית:        [___________]                │
│  סעיף תקציבי:       [___________]                │
│  מספר עותקים:       [___]                        │
│  חירור:            ( ) כן  ( ) לא                 │
│  דו צדדי:           ( ) כן  ( ) לא                 │
│  הידוק:             ( ) כן  ( ) לא                 │
│  יש דף פתיח:        ( ) כן  ( ) לא                 │
│    ↳ עמוד ריק בין דף פתיחה לתוכן: ( ) כן ( ) לא    │  ← enabled only if the row above is כן
│  שקפים בעמוד (PPT בלבד): [___]                    │  ← shown only once a PPT/PPTX is attached
├───────────────────────────────────────────────┤
│  קבצים מצורפים                [הוסף קובץ]        │
│  [סמן הכל כצבעוני]  [סמן הכל כשחור-לבן]           │
│  ┌─────────────────────────────────────────┐   │
│  │ שם קובץ │ עמודים (זוהה) │ עריכה │ צבע │ הסר │   │
│  │ ...     │ ...           │ ...   │ ... │ ... │   │
│  └─────────────────────────────────────────┘   │
├───────────────────────────────────────────────┤
│  [הודעות אימות, אם יש]                           │
│                                     [שלח בקשה]   │
└───────────────────────────────────────────────┘
```

- The attachment grid: each row shows file name, the auto-detected page
  count (or "לא זוהה אוטומטית — נא להזין ידנית" if detection failed), an
  editable numeric field for the manual override, a required Color/B&W
  toggle (two mutually exclusive buttons, unset by default — visually
  flagged, e.g. red outline, until the user picks one), and a remove
  button.
- "סמן הכל כצבעוני" / "סמן הכל כשחור-לבן" set `ColorMode` on every current
  attachment in one click; still overridable per row afterward.
- "שלח בקשה" is disabled (with inline validation messages explaining why)
  until: all required request-level fields are filled, at least one file
  is attached, every attachment has an explicit `ColorMode`, and every
  attachment has a resolvable `EffectivePageCount`.

### 8.2 Confirmation Dialog

Modal, opened when "שלח בקשה" is clicked and validation passes. Three
internal states within the same dialog (avoids stacking multiple popups):

1. **Review** — read-only rendering of the entire `PrintRequest`: every
   request-level field with its Hebrew label and current value, then a
   table of every attachment (file name, effective page count, whether it
   was auto-detected or manually entered, color/B&W). Two buttons:
   "חזרה לעריכה" (back to the form, no data lost) and "אישור ושליחה"
   (confirm & send).
2. **Sending** — shown after confirm is clicked; buttons disabled, a busy
   indicator and status text ("שולח הודעת דוא"ל…", then "מעדכן את קובץ
   האקסל…") reflecting the orchestration steps in §9 as they happen.
3. **Result** — final state, message driven by `SubmissionResult` (see
   §9.3 for the exact message matrix), with a "סגור" (close) button that
   returns to a blank form for the next request on success, or stays open
   on the filled form on hard failure so the user can retry.

## 9. Submit Flow

### 9.1 Client-side validation (before the confirmation dialog opens)

- `ProgramName`, `BudgetLine` non-empty.
- `CopiesCount ≥ 1`.
- If `HasCoverPage == false`, `BlankPageBetweenCoverAndContent` is forced
  to `false`/disabled in the UI (not user-editable).
- If no attachment is a PowerPoint file, `SlidesPerPage` is hidden/ignored;
  if at least one is, it's required and must be ≥ 1.
- At least one attachment present.
- Every attachment has `ColorMode` set and a non-null `EffectivePageCount`.

### 9.2 Orchestration (`RequestSubmissionService.SubmitAsync`)

Runs after the user confirms in the dialog:

```
1. Read current Outlook user (NameSpace.CurrentUser) → SubmittedByDisplayName/Email
2. Compose & send primary email via OutlookEmailService
   - To: config.Recipient.PrimaryEmail
   - Subject: "בקשת הדפסה - {ProgramName}"
   - HTML body, RTL (dir="rtl"), containing every request-level field and,
     per attached file, its name and color/B&W setting.
     Page counts are NOT included in this email — the person printing acts
     on file + color/B&W + copies/finishing options; page count is
     tracking/reporting data for the Excel log, not a printing instruction.
   - Attachments: the original files themselves (FilePath for each
     AttachmentItem), so the recipient has everything needed to print
     without going back to the requester.
   - If this step throws (Outlook not installed/running and cannot be
     started, COM failure, etc.) → this is a hard failure. Primary email
     IS the actual submission mechanism, so there's no silent fallback for
     it. Show a clear error in the Result state, keep the filled form
     intact so nothing is lost, and offer a "נסה שוב" (retry) action.
     PrimaryEmailSent = false; the flow stops here.
3. PrimaryEmailSent = true. Continue — steps 4-5 must not be able to fail
   the user's submission, since the email that matters already went out.
4. Try: write one row per attachment to the shared Excel file
   (ExcelRequestWriter, see §9.4 for resilience). On any exception
   (unreachable path, locked file, write conflict, etc.):
     ExcelWriteSucceeded = false, exception logged, continue to step 5.
   On success: ExcelWriteSucceeded = true, skip step 5.
5. (Only if step 4 failed) Send fallback email to
   config.FallbackEmail.Address containing the full data that would have
   gone to Excel — one line/row per attachment, all request-level fields,
   requester identity, timestamp — formatted for a human to manually type
   into the Excel file. Wrapped in its own try/catch; failure here does
   NOT fail the submission either — see §9.3 for how this is surfaced.
     FallbackEmailSent = true/false accordingly.
6. Return SubmissionResult { PrimaryEmailSent, ExcelWriteSucceeded, FallbackEmailSent, ErrorMessage }
```

### 9.3 Result messaging (Result state of the confirmation dialog)

| PrimaryEmailSent | ExcelWriteSucceeded | FallbackEmailSent | Message shown |
|---|---|---|---|
| true | true | — | "הבקשה נשלחה ונרשמה בהצלחה." (full success) |
| true | false | true | "הבקשה נשלחה בהצלחה. עדכון קובץ האקסל נכשל אוטומטית, ולכן נשלחה הודעת גיבוי לעדכון ידני." (submitted; recorded via fallback) |
| true | false | false | "הבקשה נשלחה בהצלחה, אך גם עדכון האקסל וגם הודעת הגיבוי נכשלו." — plus the full request data rendered directly in the dialog with a copy-to-clipboard action, so the user can report it manually themselves. This double-failure edge case is rare but must never silently lose data. |
| false | — | — | Hard failure — see §9.2 step 2. Submission did not happen; user can retry. |

### 9.4 Excel write resilience (§9.2 step 4 detail)

The shared file is a plain UNC/network path (§5), so it can be unreachable
(drive not mounted), locked (someone has it open, or another submission is
writing at the same instant), or simply not yet configured (placeholder
path). `ExcelRequestWriter`:

- Checks `File.Exists(path)` first; missing/unreachable path is treated as
  an immediate failure (no point retrying a path that isn't there).
- Wraps the open-append-save sequence in a small retry loop
  (`Excel.WriteRetryAttempts` / `WriteRetryDelayMs` from config — short
  delay, e.g. a few hundred ms, a few attempts) to ride out a transient
  lock from a near-simultaneous write by another user.
- Any exception surviving the retries (still locked, permissions, corrupt
  file, disk full, etc.) is caught, logged, and reported as
  `ExcelWriteSucceeded = false` — never thrown up to the UI.
- **Known limitation, called out rather than solved here:** ClosedXML
  read-modify-write against a shared file is not a true multi-writer-safe
  transaction — two submissions landing in the same short window could
  still race. The retry loop reduces but doesn't eliminate this. If
  concurrent submissions turn out to be frequent in practice, a follow-up
  design (e.g. a tiny local write-queue service, or moving the log to
  something transactional) may be worth revisiting — not needed for v1.

### 9.5 Excel schema (one row per attachment, per requirements)

| Column | Source |
|---|---|
| תאריך ושעה | `SubmittedAtUtc` (local time) |
| נשלח על ידי | `SubmittedByDisplayName` / email |
| שם התכנית | `PrintRequest.ProgramName` |
| סעיף תקציבי | `PrintRequest.BudgetLine` |
| מספר עותקים | `PrintRequest.CopiesCount` |
| חירור / דו צדדי / הידוק / דף פתיח / עמוד ריק | request-level flags, repeated on every row for that request |
| שקפים בעמוד | `PrintRequest.SlidesPerPage` (blank if not applicable) |
| שם קובץ | `AttachmentItem.FileName` |
| מספר עמודים | `AttachmentItem.EffectivePageCount` |
| צבעוני / שחור-לבן | `AttachmentItem.ColorMode` |

The fallback email (§9.2 step 5) carries the same row-per-attachment data,
formatted as a readable table in the email body, so reconciling it into
Excel by hand is a direct copy.

## 10. Failure Handling Summary

| Failure | Handling |
|---|---|
| Word not installed / COM fails when counting a Word file | `DetectionSucceeded = false`, `DetectedPageCount = null`; UI shows "לא זוהה אוטומטית", user must type a page count manually. Does not block submission. |
| PDF unreadable/corrupt/encrypted | Same pattern — detection failure, manual entry required. |
| PPTX slide count fails (corrupt file) | Same pattern. |
| Legacy `.ppt` attached | Not auto-detected (§6.3); same manual-entry pattern. |
| Outlook not installed/running and can't be started | Hard failure at submit time (§9.2 step 2) — this is the one failure that must block, since it's the actual submission channel. Form stays filled, user can retry. |
| Excel file unreachable/locked | Non-blocking; triggers fallback email (§9.2–9.4). |
| Fallback email also fails | Non-blocking; data shown in-app for manual copy (§9.3 last row). |

## 11. Open Questions

Everything else in this document is a settled design decision. Only these
three values are placeholders pending real answers:

1. **Excel file path** — the real shared/network path for
   `PrintRequests.xlsx` (currently `PLACEHOLDER_UNC_PATH\PrintRequests.xlsx`
   in `appsettings.json`).
2. **Primary recipient email** — the real address of the person who prints
   requests (currently `PLACEHOLDER_PRINT_RECIPIENT@example.com`).
3. **Fallback recipient email** — the real address for the manual-
   reconciliation fallback (currently
   `PLACEHOLDER_FALLBACK_RECIPIENT@example.com`; may end up being the same
   person/mailbox as #2, or someone else — to confirm).
