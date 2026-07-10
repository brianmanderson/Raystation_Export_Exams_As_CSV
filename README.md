# Raystation_Export_Exams_As_CSV

RayStation scripting utilities for batch-exporting DICOM images and RT structure sets
for a list of patients given in a CSV file. Instead of exporting exams one patient at a
time through the RayStation UI, you provide a CSV of MRNs and the scripts walk the
patient database, load each patient, and write the DICOM files to disk under
`export_path/MRN/CaseName/ExamName`. One-off clinical scripts from 2020, not maintained.

## Scripts

- `Export_Images_RT_from_CSV.py` — expects a CSV with rows of `MRN, Exam1, Exam2, ...`.
  For each patient it exports the DICOM images and RT structure set of each named exam.
  Set `export_path` in `main()` before running.
- `Export_Everything_from_CSV.py` — expects a CSV whose first column is the MRN. Exports
  images and RT structures for every exam that has at least one contoured ROI. (Contains
  additional, currently disabled code for beam-set dose, DRR, and spatial-registration
  export.)

## How it works

- Runs inside RayStation's scripting environment (`from connect import *`); it is not a
  standalone Python program. A Windows Forms dialog prompts for the CSV file.
- Patients are looked up by MRN via `PatientDB.QueryPatientInfo`, falling back to the
  index service, and skipping copy patients (last name suffixed `_0`–`_4`).
- Exports use `case.ScriptableDicomExport`. A folder that already exists (or contains
  `Completed.txt`) is skipped, so interrupted runs can be resumed.

## Requirements

- RayStation with scripting enabled (IronPython/.NET; uses `clr` and
  `System.Windows.Forms`).
- Write access to the export destination path.
