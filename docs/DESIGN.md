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
| MVVM | Hand-written `INotifyPropertyChanged` view models for now (e.g. `AttachmentItemViewModel`) - simpler to get right without a build environment to verify source-generator output against; CommunityToolkit.Mvvm remains an option if hand-written boilerplate grows unwieldy |
| Outlook automation | COM interop, likely late-bound (`Type.GetTypeFromProgID` + `dynamic`) given the PIA-version fragility found while wiring Word - to be confirmed when this phase is built |
| Word page counting | Late-bound COM interop (`Type.GetTypeFromProgID("Word.Application")` + `dynamic`/`IDispatch`), not a compiled `Microsoft.Office.Interop.Word` PIA reference - see §6.2 |
| PDF page counting | PdfSharp (see §6.1) |
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
is already STA, so short COM calls can run there, but page-counting (Word
in particular, which spins up a real `WINWORD.EXE` process) should not
block the UI. **Implemented**: attachment page counting runs on a
dedicated background thread explicitly created with `ApartmentState.STA`
(not the default MTA thread pool), bridged back via a
`TaskCompletionSource` that the UI `await`s. A full-window overlay (a
rotating ring + status text naming the file currently being processed) is
shown for the duration and blocks input to the rest of the form — it can
only animate at all because the UI thread stays free during the COM call.

## 4. Data Model

### `PrintRequest` (one per submission, request-level fields)

| Field | Type | Hebrew label | Notes |
|---|---|---|---|
| `ProgramName` | string | שם התכנית | required |
| `BudgetLine` | string | סעיף תקציבי | required |
| `CopiesCount` | int | מספר עותקים | required, ≥ 1 |
| `HolePunch` | bool | חירור | default false |
| `DoubleSided` | bool | דו"צ | default false |
| `Stapling` | bool | הידוק | default false |
| `PageType` | enum (paper size) | סוג דף | closed dropdown of paper sizes. Most-used sizes (A4, A3, A5) pinned at the top, A4 selected by default; remaining ISO sizes (A2, A1, A0, A6) listed below for edge cases. Replaces the earlier cover-page fields (`HasCoverPage` / `BlankPageBetweenCoverAndContent`) — the cover-page concept is legacy and no longer used by requesters |
| `SlidesPerPage` | int? | (אם נשלח כ‑PPT) כמה שקפים בעמוד | only required/shown when at least one attachment is a PPT/PPTX |
| `Notes` | string? | הערות נוספות | free-text, optional |
| `Attachments` | `List<AttachmentItem>` | קבצים מצורפים | required, ≥ 1 item |
| `SubmittedByDisplayName` / `SubmittedByEmail` | string | — | **not a form field, not yet implemented.** Will be read automatically from the current Outlook session (`NameSpace.CurrentUser`) at submit time once the Outlook phase is built, used only for the Excel log / fallback email "requested by" column. No manual entry needed since Outlook already knows who is sending. |
| `SubmittedAt` | DateTime | — | **implemented** as local time (`DateTime.Now`), not UTC — simpler for a single-timezone internal tool, and it's what both a human reading the Excel log and the month/year sheet grouping (§9.5) need. Set at submission time |

Color/B&W is intentionally **not** on this object — it lives per file on
`AttachmentItem`, per the closed design decision in the brief.

### `AttachmentItem` — two forms, mutable while editing vs. immutable at submit

**`AttachmentItemViewModel`** (WPF project, `INotifyPropertyChanged`) — the
mutable, bindable form backing the attachment grid while the user is still
editing:

| Field | Type | Notes |
|---|---|---|
| `FilePath` | string | full path to the original file on disk |
| `FileName` | string | derived, shown in UI |
| `FileKind` | enum `{ Pdf, Word, PowerPoint }` | derived from extension |
| `DetectedPageCount` | int? | result of auto-detection; `null` if detection failed or isn't supported for the file type. Kept internally to compare against, not shown as its own column |
| `PageCount` | int? | **the single field the user sees and edits**, pre-filled with `DetectedPageCount` when detection succeeds (empty otherwise). Revised from an earlier two-column "detected display + separate manual override" design, which read as confusing/unintuitive in testing - one editable value is simpler to reason about than two related-but-separate numbers. Submission is blocked if this is null (user must supply a number when auto-detection fails) |
| `PageCountStatus` | string (computed) | read-only companion label next to the editable field: "זוהה אוטומטית" if `PageCount == DetectedPageCount` and non-null, "הוזן ידנית" if the user changed/entered it, "⚠ יש להזין מספר עמודים" (shown in red/bold) if still null |
| `ColorMode` | enum `{ Color, BlackAndWhite }` | **Defaults to `BlackAndWhite`** for every newly-attached file (revised from the original "no default" decision after hands-on UI testing showed an empty selector reads as broken rather than deliberate). Still overridable per file, and via the "mark all" buttons |

**`Core.Models.AttachmentItem`** (Core project, immutable) — a plain
submission-time snapshot built via `AttachmentItemViewModel.ToAttachmentItem()`
once `Send` validation has confirmed `PageCount` and `ColorMode` are both set:
`FilePath`, `FileName`, `FileKind`, `PageCount` (non-nullable `int` here — no
longer optional once snapshotted), `ColorMode` (non-nullable). This is the
form `PrintRequest.Attachments` holds and what the confirmation dialog and
(eventually) the Excel/email writers consume — keeping it a plain, immutable
POCO in Core rather than reusing the WPF ViewModel directly.

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

### 6.1 PDF page counting — **PdfSharp**

| Option | Verdict |
|---|---|
| **PdfSharp** (chosen) | MIT-licensed, mature, actively-versioned (6.x targets .NET 6/8), cross-platform. `PdfReader.Open(path, PdfDocumentOpenMode.InformationOnly)` + `.PageCount` reads the page count cheaply without parsing content streams. |
| PdfPig | Originally chosen (pure managed, no COM), but **dropped**: it currently has no stable NuGet release at all, and the only available prerelease build (`0.1.9-alpha001-patch1`) threw `FileNotFoundException` from inside its own DLL on every real PDF tested. Not usable as of this writing. |
| iText7 | Very capable, but AGPL-licensed unless a commercial license is purchased — a licensing complication this internal tool doesn't need. |

### 6.2 Word page counting — Word COM Interop (required by brief)

Opens the document, calls `Document.Repaginate()` for an accurate, live
count (the cached statistics/built-in properties can be stale for documents
that haven't been repaginated since last edit), then reads
`Document.ActiveWindow.Selection.Information(wdNumberOfPagesInDocument)`.
Documents are opened `ReadOnly = true` with `Application.Visible = false`
(suppresses on-screen rendering) — but *not* also passing `Visible:false` to
`Documents.Open` itself, since that appears to suppress creation of the
document's own window entirely, leaving `Selection` null (confirmed by a
real "cannot perform runtime binding on a null reference" failure). Always
closed and COM-released in a `finally` block to avoid orphaned
`WINWORD.EXE` processes accumulating on the requester's machine.

Originally attempted via `Document.ComputedStatistics(wdStatisticPages)`
(the documented VBA API for this), but a real test machine showed COM
automation doesn't expose that member at all —
`'System.__ComObject' does not contain a definition for
'ComputedStatistics'` — despite it being real, documented VBA surface.
`Selection.Information` is the more reliably-automatable path.

**Late-bound (no `Microsoft.Office.Interop.Word` PIA reference):** the
compiled PIA pins a specific Office-version assembly dependency (`office,
Version=15.0.0.0`) that isn't guaranteed to be resolvable at runtime — this
surfaced as an unhandled `FileNotFoundException` on a real test machine,
and is a fragile approach in general for an internal tool that may run
against different installed Office versions across different users'
machines. Instead, `WordPageCounter` gets the COM type via
`Type.GetTypeFromProgID("Word.Application")`, creates it with
`Activator.CreateInstance`, and drives it entirely through `dynamic`
(late-bound `IDispatch` calls) — this works against whatever Word version
is actually installed, with zero compile-time assembly-version dependency.

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

### 6.5 Outlook automation — COM Interop, late-bound — **implemented**

Late-bound, same approach and reasoning as `WordPageCounter` (§6.2): gets
the COM type via `Type.GetTypeFromProgID("Outlook.Application")` and drives
it through `dynamic`, no compiled `Microsoft.Office.Interop.Outlook` PIA
reference and its Office-version-pinned assembly dependency risk. Creating
a new `Outlook.Application` when Outlook is already running attaches to
that existing instance automatically — this is Outlook's own COM server
behavior, unlike Word (which spawns a separate instance per
`Activator.CreateInstance` call), so no separate "attach to running
instance" step is needed.

Composes an HTML email (`dir="rtl"`) with every request-level field
**except page counts** (page counts are tracking data for the Excel log,
not a printing instruction — the person printing acts on file name +
color/B&W + finishing options) and the color/B&W setting per attached
file, attaches the original files themselves, and sends via `MailItem.
Send()`. Any failure (Outlook not installed, send rejected, etc.) comes
back as a clear error shown to the user rather than a silent fallback,
since this is the actual submission mechanism (§9.2).

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
│                בקשת שכפול חדשה                  │
├───────────────────────────────────────────────┤
│  פרטי הבקשה                                     │
│  שם התכנית:        [___________]                │
│  סעיף תקציבי:       [___________]                │
│  מספר עותקים:       [___]                        │
│  חירור:            ( ) כן  ( ) לא                 │
│  דו"צ:              ( ) כן  ( ) לא                 │
│  הידוק:             ( ) כן  ( ) לא                 │
│  סוג דף:            [A4 ▾]                         │
│  שקפים בעמוד (PPT בלבד): [___]                    │  ← shown only once a PPT/PPTX is attached
│  הערות נוספות:      [___________]                 │
├───────────────────────────────────────────────┤
│  קבצים מצורפים                [הוסף קובץ]        │
│  [סמן הכל כצבעוני]  [סמן הכל כשחור-לבן]           │
│  ┌─────────────────────────────────────────┐   │
│  │ שם קובץ │ עמודים (זוהה) │ עריכה │ צבע │ הסר │   │
│  │ ...     │ ...           │ ...   │ ... │ ... │   │
│  └─────────────────────────────────────────┘   │
├───────────────────────────────────────────────┤
│  [הודעות אימות, אם יש]                           │
│              שלח אל: [dany@lahav.ac.il ▾]  [שלח בקשה]   │
└───────────────────────────────────────────────┘
```

The recipient dropdown is a hardcoded placeholder (currently one entry)
rather than something users are expected to touch — a dropdown instead of
a fixed label only so more recipients can be added later without
redesigning this control.

- The attachment grid: each row shows file name, a single editable page-count
  field (pre-filled from auto-detection when it succeeds), a read-only status
  label next to it ("זוהה אוטומטית" / "הוזן ידנית" / a red "⚠ יש להזין מספר
  עמודים" when still empty), a Color/B&W selector defaulting to B&W for every
  newly-attached file, and a remove button.
- "סמן הכל כצבעוני" / "סמן הכל כשחור-לבן" set `ColorMode` on every current
  attachment in one click; still overridable per row afterward.
- "שלח בקשה" validates before submitting: at least one file attached, and
  every attachment has a resolvable `PageCount` (detected or manually
  entered) — otherwise it shows an error listing which files still need a
  page count, rather than sending an incomplete request.
- While a newly-added file's page count is being detected, a full-window
  overlay (rotating ring + "מזהה מספר עמודים: {file name}") covers the form
  and blocks input until it's done — added specifically because Word
  counting can take several real seconds and an unresponsive-looking window
  otherwise reads as frozen.

### 8.2 Confirmation Dialog

Modal, opened when "שלח בקשה" is clicked and validation passes, built from
the `PrintRequest` object (not a pre-formatted string) so every field and
attachment renders as its own set of elements rather than one concatenated
block of text — mixing Hebrew, numbers, and Latin file names in a single
text run is what caused visibly jumbled rendering in an earlier version.
Request-level fields are a label/value grid (one row per field); attachments
are a list, one row each, with file name / page count / color mode as
separate `TextBlock`s in a `Grid` row (not a `StackPanel`, so a long file
name wraps in its own bounded column instead of silently clipping).

**Currently implemented: only the Review state below** — there's no
Outlook/Excel backend yet (that's a later phase), so confirming just shows
a plain "sent" acknowledgement. The two additional states are the intended
design for once that backend exists:

1. **Review** *(implemented)* — read-only rendering of the entire
   `PrintRequest`: every request-level field with its Hebrew label and
   current value, then a row per attachment (file name, page count,
   color/B&W). Two buttons: "חזרה לעריכה" (back to the form, no data lost)
   and "אישור ושליחה" (confirm & send).
2. **Sending** *(not yet built)* — shown after confirm is clicked; buttons
   disabled, a busy indicator and status text ("שולח הודעת דוא"ל…", then
   "מעדכן את קובץ האקסל…") reflecting the orchestration steps in §9 as they
   happen.
3. **Result** *(not yet built)* — final state, message driven by `SubmissionResult` (see
   §9.3 for the exact message matrix), with a "סגור" (close) button that
   returns to a blank form for the next request on success, or stays open
   on the filled form on hard failure so the user can retry.

## 9. Submit Flow

### 9.1 Client-side validation (before the confirmation dialog opens)

- `ProgramName`, `BudgetLine` non-empty.
- `CopiesCount ≥ 1`.
- `PageType` always has a value (defaults to A4).
- If no attachment is a PowerPoint file, `SlidesPerPage` is hidden/ignored;
  if at least one is, it's required and must be ≥ 1.
- `Notes` is optional, no validation.
- At least one attachment present.
- Every attachment has a non-null `PageCount` (blocks submission with
  an error listing the offending files otherwise); `ColorMode` always has a
  value since it defaults to `BlackAndWhite`.

### 9.2 Orchestration (`RequestSubmissionService.SubmitAsync`)

**Currently implemented directly in `MainWindow.Send_Click`** (steps 2-4
below, minus reading the Outlook user identity and the fallback email) -
not yet its own `RequestSubmissionService`/`SubmissionResult` in Core, and
the recipient is a hardcoded placeholder dropdown (one entry,
`dany@lahav.ac.il`) rather than config-driven (§5/§11) - a real "no silent
fallback" hard failure on email error and a "doesn't block success on
Excel failure" soft failure on the write are both in place and match the
plan below. Runs after the user confirms in the dialog:

```
1. Read current Outlook user (NameSpace.CurrentUser) → SubmittedByDisplayName/Email
2. Compose & send primary email via OutlookEmailService
   - To: config.Recipient.PrimaryEmail
   - Subject: "בקשת שכפול - {ProgramName}"
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

### 9.4 Excel write resilience (§9.2 step 4 detail) — **implemented**

`ExcelRequestWriter` (`Core/Services/Excel`) takes a `filePath` directly
(not yet read from `appsettings.json` — that wiring happens once this is
connected into the live `Send` flow, not needed to build/test the writer
in isolation):

- Checks `File.Exists(path)` first; missing/unreachable path is treated as
  an immediate failure (no point retrying a path that isn't there) —
  returns `false`, never throws.
- Wraps the open-append-save sequence in a retry loop (3 attempts, ~500ms
  apart) that only retries on `IOException` specifically (the shape a
  locked/in-use file throws) — other exceptions (corrupt file, permissions,
  etc.) fail immediately without retrying.
- Every failure path returns `false` rather than throwing, so the caller
  (eventually `RequestSubmissionService`) can fall back to the
  reconciliation email without the exception ever reaching the UI.
- **Known limitation, called out rather than solved here:** ClosedXML
  read-modify-write against a shared file is not a true multi-writer-safe
  transaction — two submissions landing in the same short window could
  still race. The retry loop reduces but doesn't eliminate this. If
  concurrent submissions turn out to be frequent in practice, a follow-up
  design (e.g. a tiny local write-queue service, or moving the log to
  something transactional) may be worth revisiting — not needed for v1.
- Tested (`PrintRequestApp.Tests`) against a freshly-created dummy
  workbook per test (not a real production file): correct sheet/row
  creation, the `כללי` default, appending without clobbering existing
  rows, a missing path, and a locked file (opened with `FileShare.None` in
  the test to force the failure).

### 9.5 Excel schema — **implemented, provisional pending a walkthrough with
the actual spreadsheet owner**

Scoped down from an earlier, richer draft of this section to exactly what's
confirmed needed right now; revisit once that conversation happens. One
worksheet per calendar month **and year** the request was submitted in
(sheet name `"MM-yyyy"`, e.g. `"07-2026"`, created with a header row the
first time a request lands in that month), one row per attachment:

| Column | Source |
|---|---|
| תאריך | `PrintRequest.SubmittedAt` |
| שם תכנית | `PrintRequest.ProgramName`, defaulting to `"כללי"` when blank — not every copy job is for a named program |
| סעיף תקציבי | `PrintRequest.BudgetLine` |
| שם קובץ | `AttachmentItem.FileName` |
| סוג צבע | `AttachmentItem.ColorMode` (צבעוני / שחור-לבן) |
| סה"כ עמודים | `AttachmentItem.PageCount × PrintRequest.CopiesCount` — the actual print/copy volume for that file, not just its page count |

The fallback email (§9.2 step 5) carries the same row-per-attachment data,
formatted as a readable table in the email body, so reconciling it into
Excel by hand is a direct copy.

## 10. Failure Handling Summary

| Failure | Handling |
|---|---|
| Word not installed / COM fails when counting a Word file | `DetectedPageCount = null`; `PageCount` starts empty, `PageCountStatus` shows the red "⚠ יש להזין מספר עמודים" hint, user must type a page count manually. Does not block submission by itself (only an actually-empty `PageCount` at send time blocks). |
| PDF unreadable/corrupt/encrypted | Same pattern — detection failure, manual entry required. |
| PPTX slide count fails (corrupt file) | Same pattern. |
| Legacy `.ppt` attached | Not auto-detected (§6.3); same manual-entry pattern. |
| Outlook not installed/running and can't be started | Hard failure at submit time (§9.2 step 2) — this is the one failure that must block, since it's the actual submission channel. Form stays filled, user can retry. |
| Excel file unreachable/locked | Non-blocking; triggers fallback email (§9.2–9.4). |
| Fallback email also fails | Non-blocking; data shown in-app for manual copy (§9.3 last row). |

## 11. Open Questions

Everything else in this document is a settled design decision. Only these
three values are placeholders pending real answers:

1. **Excel file path** — the real shared/network path for the log
   (currently a dummy file on the Desktop for manual testing, not
   config-driven yet).
2. **Primary recipient email** — the real address of the person who prints
   requests (currently hardcoded to `dany@lahav.ac.il` for testing, as the
   only entry in the "שלח אל" dropdown — explicitly described as something
   that "would never be touched" day-to-day, so a dropdown rather than a
   settings file for now).
3. **Fallback recipient email** — the real address for the manual-
   reconciliation fallback (currently
   `PLACEHOLDER_FALLBACK_RECIPIENT@example.com`; may end up being the same
   person/mailbox as #2, or someone else — to confirm).
