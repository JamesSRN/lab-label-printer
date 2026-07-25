# Lab Label Printer

A simple, self-contained Microsoft Excel workbook (with macros) for printing patient
lab labels at the **Saturday Clinic for the Uninsured** on a **Brother QL‑1100** label
printer.

It is intentionally lightweight — no logging, no database, no patient records are
stored. A volunteer fills in five fields, clicks a button, chooses how many labels to
print, and confirms. That's it.

---

## What a label looks like

```
Name: Last, First
DOB: 00/00/0000
Sex: M/F/Other
Date: 7/25/2026        <- bold
Saturday Clinic for the Uninsured
```

## Features

- **Five simple inputs:** Last Name, First Name, DOB, Sex (M / F / Other dropdown),
  and Date (auto-fills to today, editable).
- **Live preview:** the label updates as you type so you see exactly what will print.
- **PRINT LABELS button:** validates that nothing is missing, asks how many copies,
  shows a confirmation with the full label and the target printer, then prints.
- **Clear button:** wipes the fields and resets the date to today.
- **Error-resistant:** blocks printing until required fields are filled, rejects
  invalid quantities, and reports a friendly message if the printer can't be reached.

## Requirements

- Windows + Microsoft Excel (desktop) with macros enabled.
- A **Brother QL‑1100** installed, with a **DK‑1201** die-cut address roll
  (29 × 90 mm ≈ 1⅛" × 3½") loaded.

## How to use

1. Open **Lab Label Printer.xlsm**. If Excel shows a security bar, click
   **Enable Content** (needed for the buttons to work).
2. Type the patient's Last Name, First Name, and DOB; pick Sex from the dropdown.
   The Date defaults to today.
3. Click **PRINT LABELS**, enter the number of copies, and click **OK** on the
   confirmation.
4. Click **Clear** to reset for the next patient.

## Important: printer / label-size note

The clinic's Brother QL‑1100 driver can get "stuck" reporting the larger
**2.4" × 3.9"** (DK‑1202 shipping) label even when the small **DK‑1201** address roll
is loaded — which causes a *label size mismatch* error and blocks printing.

To avoid fighting the driver, this workbook is deliberately set up as follows:

- **Page size:** 2.4" × 3.9" (the size the driver expects) — so **no mismatch error**.
- **Content placement:** the label text is pinned to the **top-left corner** at 100%
  scale (no centering, no shrink-to-fit), so it lands correctly on the smaller
  DK‑1201 label that is physically loaded.

If the mismatch error ever returns, it usually means the roll was changed and the
driver re-detected a different size — reload the DK‑1201 roll (open/close the cover to
force re-detection) and confirm the page size is still 2.4" × 3.9" with top-left
placement.

## Under the hood (VBA macros)

| Macro | Purpose |
|-------|---------|
| `PrintLabels` | Validates inputs, prompts for quantity, confirms, prints N copies. |
| `ClearForm` | Clears the input cells and resets the date to today. |
| `SetupSheet` | One-time builder: lays out the sheet, dropdown, buttons, and page setup. |
| `Workbook_Open` | Sets the Date field to today's date whenever the file is opened. |

The label preview cells (E3:E7) are live formulas that reference the input cells, so
the preview and the printout always match.

## Making common changes

- **Clinic name / label text:** edit the preview cell formulas (E3:E7) on the sheet, or
  the `ComposeLabel` routine in the VBA (Developer → Visual Basic, `Alt+F11`).
- **Switch to a different label size:** change the page size in
  **File → Print → paper-size dropdown**, then adjust margins/placement to suit.
- **Add a field:** add an input cell and extend the preview formulas + the validation
  list in `PrintLabels`.

## Files

- `Lab Label Printer.xlsm` — the workbook (this is the whole tool).
- `README.md` — this file.
- `HANDOFF.md` — a fuller handoff/maintenance guide.

---

*Built for the Saturday Clinic for the Uninsured. Contains no patient data.*
