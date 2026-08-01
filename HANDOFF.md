# Lab Label Printer — Handoff & Maintenance Guide

**Project:** Lab Label Printer for the Saturday Clinic for the Uninsured
**File:** `Lab Label Printer.xlsm`
**Printer:** Brother QL‑1100 (primary); portable to other Brother QL models such as the QL‑820NWBc — see §3
**Label stock:** Brother DK‑1201 die-cut address labels — 29 × 90 mm (≈ 1⅛" × 3½")
**Last updated:** July 26, 2026

This document explains everything the next person needs to run, maintain, or modify the
tool. The [README](./README.md) is the short version; this is the full picture.

---

## 1. What this is (and isn't)

It is a single Excel workbook that prints patient lab labels one patient at a time. A
volunteer types five fields and clicks a button.

It intentionally does **not** log anything, store patient data, or keep a database. When
the file is closed, nothing about the patients who were printed is retained. This keeps
it simple and avoids holding any PHI. (The clinic's separate dispensary workbook is the
tool that does full logging — this lab-label tool is deliberately stripped down.)

## 2. Volunteer workflow

1. Open **Lab Label Printer.xlsm**. If a yellow security bar appears, click
   **Enable Content** so the buttons work.
2. Fill in the yellow input cells:
   - **Last Name**
   - **First Name**
   - **DOB**
   - **Sex** — choose M, F, or Other from the dropdown
   - **Date** — already filled with today's date; change only if needed
3. Confirm the **Label preview** on the right looks correct.
4. Click **PRINT LABELS**.
   - Enter the number of copies (default 1) and click OK.
   - A confirmation shows the exact label text and the printer name — click **OK** to
     print or **Cancel** to stop.
5. Click **Clear** to reset the form for the next patient.

If a required field is empty, the tool lists what's missing and does **not** print.

## 2b. Batch printing (Batch tab)

For printing many patients at once, use the **Batch** tab:

1. Enter or paste one patient per row under the headers **Last Name, First Name, DOB,
   Sex, Date, # Labels** (rows 6 onward; the list area is rows 6–105).
2. Date blank = today; # Labels blank = 1.
3. Click **PRINT ALL LABELS** — it validates every row, shows a "print X labels for Y
   patients?" confirmation (listing any incomplete rows it will skip), then prints them
   all. Each label is produced by copying the row's values into the Labels sheet and
   printing it, so batch output is identical to single-label output. The Labels sheet is
   restored to its previous contents when the batch finishes.
4. **Clear List** empties the table.

The Batch-tab code lives in the **BatchPrint** module (also exported to `BatchPrint.bas`
in the repo). To rebuild the Batch tab from scratch, run `SetupBatch`.

## 3. Printer configuration — the important part

### The problem
The Brother QL‑1100 driver on the clinic PC tends to "stick" on the larger
**2.4" × 3.9"** shipping-label size (DK‑1202) even when the small **DK‑1201** address
roll is physically loaded. When Excel's page size doesn't match what the driver thinks
is loaded, you get a **label size mismatch** error and it won't print.

### The solution baked into this file
Instead of forcing the driver to the small size (which has historically been unreliable),
the workbook is set to the size the driver *expects* and the content is placed so it
lands on the small label:

| Setting | Value |
|---|---|
| Paper size | **2.4" × 3.9"** (matches what the driver reports) |
| Orientation | **Landscape** |
| Scaling | **No Scaling** (Adjust to **100%**, not "fit to page") |
| Center on page | **Off** — both Horizontally and Vertically unchecked |
| Print area | **E3:E7** |
| Margins | Top/Left/Right/Bottom = **0.06"**, Header/Footer = 0 |

Because centering is off and scaling is 100%, the label prints in the **top-left corner**
of the 2.4" × 3.9" area — which is exactly where the smaller DK‑1201 label sits. Result:
no mismatch error, and the text prints correctly on the small labels.

### The placement toggle (buttons on the Labels sheet)
Because the clinic uses **non‑Brother rolls** (which have no spool sensor), the printer
can't reliably auto‑detect the loaded label, so the driver sometimes stays stuck on the
big 2.4" × 3.9" size and sometimes reads the small roll correctly. To handle both without
editing code, the Labels sheet has two buttons that only change **where the content is
placed** — they never change the driver's paper size, and they **always print onto the
small lab roll**:

| Button | Use when… | What it does |
|---|---|---|
| **Small label** | the printer correctly detects the small roll | centers the content to fill the small label (no size mismatch) |
| **Large (corner)** | the driver is stuck thinking a big 2.4 × 3.9" label is loaded | keeps the big page and drops the content in the **top‑left corner**, so it still lands cleanly on the small roll (avoids the mismatch error) |

A status line just above the buttons shows which mode is active. The code lives in the
**LabelSize** module (`LabelSize.bas`), and the size/margin numbers for each mode are
plain constants at the top of that module — edit them to tune for a different roll. To
(re)create the buttons, run `SetupLabelToggle`.

> Note: this is a **placement** toggle only. It does not — and cannot — change the Brother
> driver's paper size or its black‑only vs. black/red media type; those are set once in the
> printer driver's **Printing preferences** (see §3b).

### If the mismatch error comes back
It usually means the roll was swapped and the driver re-detected a different size.
1. Make sure the **DK‑1201** roll is loaded and the printer is on/connected.
2. **Open and close the label cover** (or power-cycle the printer) to force it to
   re-read the roll.
3. In Excel, confirm **File → Print** still shows paper size **2.4" × 3.9"** and that the
   preview shows the text in the top-left corner.
4. If the text lands slightly off the physical label, nudge the **Top** and **Left**
   margins in **File → Print → Page Setup → Margins** by a few hundredths of an inch.

### Printer selection
Printing goes to whatever printer Windows currently has active. The confirmation dialog
always shows the printer name (e.g. `Brother QL‑1100 on Ne02:`) so the volunteer can
cancel if it's pointed at the wrong device. For fewest surprises, set the **Brother
QL‑1100 as the default printer** on the clinic PC.

### Switching printers or using another Brother QL model (e.g. QL‑820NWBc)
Nothing in the tool is tied to a specific printer. It prints to whatever printer is set as
active/default in Windows, and the print-error message is **model-neutral** ("Make sure the
label printer is on and selected as the default printer"). So the same workbook runs on any
Brother QL model that takes the DK‑1201 roll, including the **QL‑820NWBc** (it supports
DK‑1201 and label widths up to 62 mm). To move the tool to another QL printer:

1. In **File → Print**, pick the printer (or set it as the Windows default). Make sure it
   is actually powered on and connected — the QL‑820NWBc is a wireless/network model, so
   "plugged in" isn't enough; it must be joined to the network (or connected by USB) or
   Excel will hang on "Connecting to printer… / Loading page sizes."
2. **Re-select the paper size** — page size is stored per driver, so the QL‑1100's
   2.4" × 3.9" setting won't carry over.
   - If the new printer's driver correctly detects the **DK‑1201** roll (29 × 90 mm),
     select the true DK‑1201 size and turn centering back **on** (Page Setup → Margins →
     Center on page: Horizontally + Vertically). This fills the label and looks better than
     the top-left trick.
   - If that driver "sticks" on the larger 2.4" × 3.9" size the way the QL‑1100 does, keep
     the **2.4" × 3.9" + top-left** workaround from the table in Section 3.
3. Confirm the **File → Print** preview: one page, text on the left, SCU emblem on the
   right next to the top lines.

## 3b. Setting up the QL‑820NWBc over Bluetooth (Windows)

The 820NWBc can print over Bluetooth from Windows, but **only** if you install Brother's own
driver with the Bluetooth option — Windows' generic pairing creates a device with **no
working print queue** ("Driver unavailable"), which never prints. What worked:

1. On the printer, turn Bluetooth on: **Menu → Bluetooth → Bluetooth (On/Off) → On**. Its
   Bluetooth name is the model + 4 digits, e.g. **QL‑820NWB3827**.
2. If Windows already auto‑paired it (Settings → Bluetooth & devices → Devices shows it),
   **remove that pairing first** — Brother's installer skips already‑paired devices, so it
   won't show in the installer's list until you remove it.
3. Run Brother's **Software/Document Installer**, choose **Bluetooth Connection**, pick the
   printer from the list, and confirm the pass‑key **on the printer's LCD** when prompted.
4. **Restart the PC and power‑cycle the printer** if the queue flaps between Idle/Error.
5. Confirm **Settings → Printers & scanners → Brother QL‑820NWB** shows **"Idle."**

Windows often lists the 820 more than once (one entry per connection it has seen — Bluetooth
vs. network). Keep the one with a real driver that shows "Idle"; the extras are harmless
clutter you can remove.

### Non‑Brother rolls & the black/red media type
The clinic uses **non‑Brother rolls**, which have no spool sensor, so the printer can't
reliably auto‑detect the label. Two consequences:

- **Set the media manually** in the driver's **Printing preferences** — a fixed paper size
  and **black‑only** (not the 2‑color black/red mode). If the driver guesses "black/red"
  while a black roll is loaded, you get a media‑mismatch error.
- The **placement toggle** (Section 3, buttons on the **Developer** sheet) exists precisely
  because detection is unreliable: use **Small roll** when the driver reads the small label,
  or **Large / corner** when it's stuck on the big size. Both print onto the small roll.

### If a label prints blank
The QL‑820NWBc is **direct thermal** — no ink or ribbon; it marks heat‑sensitive paper. A
blank label almost always means the **roll is loaded upside‑down** (the heat‑sensitive side
must face the print head) or the labels **aren't direct‑thermal stock**. Flip/reload the
roll, or raise the **print density** in Printing preferences.

## 4. How the workbook is built (cell map)

Everything lives on one sheet named **Labels**.

| Cell(s) | Contents |
|---|---|
| B1 | Title: "Lab Label Printer" |
| B2 | Instruction line |
| B3:B7 | Field labels: Last Name, First Name, DOB, Sex, Date |
| C3:C7 | **Input cells** (yellow). C6 = Sex dropdown (M/F/Other). C7 = Date. |
| E2 | "Label preview:" |
| E3:E7 | **Live preview** = the print area (see formulas below) |
| ~B9 | The **PRINT LABELS** and **Clear** buttons |

The preview cells are formulas, so the preview always matches what prints:

```
E3  =IF(C4="",C3,C3&", "&C4)     (name only, no "Name:" prefix)
E4  ="DOB: "&IF(ISNUMBER(C5),TEXT(C5,"mm/dd/yyyy"),C5)
E5  ="Sex: "&C6
E6  ="Date: "&IF(ISNUMBER(C7),TEXT(C7,"m/d/yyyy"),C7)     (bold)
E7  Saturday Clinic for the Uninsured
```

## 5. The VBA (macros)

Open with **Developer → Visual Basic** or `Alt+F11`. Code lives in **Module1** and the
**ThisWorkbook** object.

| Macro | Location | What it does |
|---|---|---|
| `PrintLabels` | Module1 | Checks all fields are filled, prompts for quantity, shows a confirm dialog with the label + printer, then prints N copies. Handles cancel, non-numbers, and print errors gracefully. |
| `ClearForm` | Module1 | Clears C3:C6 and resets C7 to today's date. |
| `LabelSheet` | Module1 | Helper that returns the "Labels" sheet safely. |
| `ComposeLabel` | Module1 | Writes the preview formulas into E3:E7. |
| `SetupSheet` | Module1 | One-time builder — recreates the entire layout: column widths, labels, input styling, the Sex dropdown, the preview, the two buttons, and the page setup. Run it only if you need to rebuild the sheet from scratch. |
| `Workbook_Open` | ThisWorkbook | Sets the Date cell (C7) to today each time the file opens. |
| `PrintBatch` | BatchPrint | Reads the Batch tab (rows 6–105), validates each row, confirms the total, then prints every patient by feeding each row into the Labels sheet and printing. Restores the Labels inputs afterward. |
| `ClearBatch` | BatchPrint | Clears the Batch tab list (B6:G105). |
| `SetupBatch` | BatchPrint | One-time builder for the Batch tab (table, Sex dropdown, buttons). |
| `UseSmallLabel` | LabelSize | Sets the Labels page layout for when the printer detects the small roll (centered, fit to one label). Wired to the green **Small roll** button on the Developer sheet. |
| `UseLargeLabel` | LabelSize | Sets the top‑left "corner" layout for when the driver is stuck on the big 2.4×3.9 size. Wired to the blue **Large / corner** button. |
| `SetupLabelToggle` | LabelSize | One-time builder for the **Developer** sheet (the two placement buttons + mode readout). Also clears the old buttons/status off the Labels sheet. |
| `AddEmblem` | EmblemSetup | One-time: (re)places the SCU emblem on the right of the label and extends the print area to include it. |

> Note: `SetupSheet` resets page setup to landscape/top-left/100% but does **not** set the
> paper size (Excel can't pick the Brother custom size by name in code). If you ever
> re-run `SetupSheet`, re-select paper size **2.4" × 3.9"** in File → Print afterward.

## 6. Making common changes

- **Change the clinic name or any label line:** edit the formula in the matching E-cell
  (E3–E7). For example, change E7's text to update the clinic name line.
- **Change how the date/DOB is formatted:** edit the `TEXT(...)` format in E4/E6
  (e.g. `"mm/dd/yyyy"` vs `"m/d/yyyy"`).
- **Add or change Sex options:** Data → Data Validation on cell C6, edit the list.
- **Cap the max labels per print / warning threshold:** in `PrintLabels`, the code warns
  above 50 — change that number.
- **Use a different label size later:** set the new size in File → Print → paper-size
  dropdown, turn centering back on (or adjust margins), and re-preview.

## 7. Troubleshooting

| Symptom | Fix |
|---|---|
| "Label size mismatch" / won't print | Reload DK‑1201, open/close cover to re-detect, confirm page size 2.4"×3.9" + top-left placement (Section 3). |
| Buttons do nothing | Macros are disabled — click **Enable Content**, or lower macro security in File → Options → Trust Center. |
| Prints to the wrong printer | Set Brother QL‑1100 as the default printer; the confirm dialog shows the current printer. |
| Text prints off the edge of the label | Nudge Top/Left margins in Page Setup by a few hundredths of an inch. |
| Date is wrong | It auto-sets to today on open; just type over cell C7 (Date) if needed. |
| Preview shows blank name/DOB | Those input cells are empty — fill them in. |

## 8. Maintenance & backup

- **Source of truth:** the GitHub repo `lab-label-printer`. Commit any edited version of
  the `.xlsm` with a short message describing what changed.
- **Because `.xlsm` is a binary,** GitHub won't show line-by-line diffs — keep commit
  messages descriptive (e.g. "Change clinic name line", "Adjust top margin 0.06→0.08").
- Keep a known-good copy on the clinic PC's Desktop for day-to-day use; treat the repo as
  the backup/version history.

## 9. Ownership / contact

- Maintained by: _(add name/role)_
- Questions or changes: _(add contact)_

## 10. Planned enhancement — automatic printer selection (not yet built)

**Goal:** print correctly whether the **QL‑820NWBc (Bluetooth)** or the **QL‑1100 (USB)** is
connected, without the volunteer picking a printer in File → Print:

- Only one connected → use it automatically.
- Both connected → ask which one.
- Neither → a clear "no label printer connected" message.

Today the tool prints to Windows' active/default printer and the confirm dialog shows which
one; this enhancement would automate the choice.

### Suggested design
Add a module **`PrinterSelect.bas`** with one entry point the print code calls:

- `SetChosenPrinter() As Boolean` — picks the target printer, sets
  `Application.ActivePrinter`, and returns `False` if none is available or the user cancels.

Internally:

1. **List candidates** by name match — a printer name containing `"QL-820"` → label
   "Brother QL‑820NWBc (Bluetooth)"; `"QL-1100"` → "Brother QL‑1100 (USB)".
2. **Get each one's ActivePrinter string.** Excel needs the exact `"Name on Port:"` form
   (e.g. `Brother QL-1100 on USB001:`). Read it from the registry key
   `HKCU\Software\Microsoft\Windows NT\CurrentVersion\Devices` — each value name is a printer
   name, its data is `winspool,<port>`. Enumerate with WMI `StdRegProv` (`EnumValues` +
   `GetStringValue`) and build `name & " on " & port`.
3. **Check availability** via WMI `Win32_Printer`: treat a printer as available when
   `WorkOffline = False` **and** `PrinterStatus <> 7` (7 = offline). That reflects the
   Bluetooth 820 being in range / the USB 1100 being plugged in.
4. **Decide:**
   - 0 available → `MsgBox` "No label printer is connected. Turn on the 820 (Bluetooth in
     range) or connect the 1100 by USB, then try again." → return `False`.
   - 1 available → set `Application.ActivePrinter` to it → return `True`.
   - 2 available → `MsgBox …, vbYesNoCancel` "Which printer?  Yes = 820 (Bluetooth),
     No = QL‑1100 (USB)"; map Yes/No to the two, Cancel → return `False`.

### Where to wire it in
- **`PrintLabels` (Module1):** insert `If Not SetChosenPrinter Then Exit Sub` just before the
  confirmation dialog (the block beginning `Dim msg As String`). Then the confirm dialog's
  "Printer: …" line already shows the chosen printer.
- **`PrintBatch` (BatchPrint):** call `SetChosenPrinter` **once before** the print loop (ask
  once for the whole batch); `Exit Sub` if it returns `False`.

### Caveats / test notes
- `WorkOffline` is the best signal VBA can read but isn't perfect; if it ever picks an
  unreachable printer, the existing `PrintErr` handler still catches the failure with a clear
  message, so it degrades safely.
- Duplicate/stale printer entries (see §3b) can create more than one `"QL-820…"` match — take
  the first available one, and clean up dead duplicates in Windows.
- This is plain VBA (WMI + a registry read); it does **not** need "Trust access to the VBA
  project object model" at runtime.
- After building, export `PrinterSelect.bas` and re-export `Module1.bas` / `BatchPrint.bas`
  so the repo's text sources stay in sync with the `.xlsm`.

---

*This tool stores no patient data. Any values in the input cells are cleared with the
Clear button and are never written to disk beyond the open workbook.*
