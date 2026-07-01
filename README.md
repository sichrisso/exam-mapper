# Exam Form Mapper

Maps multiple-choice exam questions across scrambled forms (A / B / C / D / …),
detects the correct answer, and produces a structured Excel file plus a
plain-text issues report for chemistry-education research.

---

## What it produces

For each exam it writes two files, named descriptively from the detected
semester and exam:

1. **`<Semester>_<Exam>_Result.xlsx`** — the mapping. One row per master
   question, with each form's question number, the choice-order permutation,
   the detected correct answer, and an OMIT column.
2. **`<Semester>_<Exam>_Result_ISSUES.txt`** — a review report listing
   everything that might need a human check before the mapping is trusted
   (missing questions, number jumps, undetected answers, likely image
   questions, and omit notes that could not be auto-attached). When an exam is
   clean it still writes a short "No issues detected" confirmation.

---

## Requirements

- Python 3.10+
- Dependencies:

  ```
  pip install pypdf pandas openpyxl rapidfuzz
  ```

- **pypdf version matters.** The tool is tested against **pypdf 6.7.2**.
  Different pypdf versions extract superscripts (e.g. unit exponents like
  `mg L⁻¹`) slightly differently. The parser is written to handle both the
  5.x and 6.x extraction styles, but if you see questions going missing,
  first check your version with `python -c "import pypdf; print(pypdf.__version__)"`.

---

## Usage

### Command line

Process every exam under a folder tree (`data/<semester>/<exam>/<form PDFs>`):

```
python exam_mapper.py data/
```

Process a single exam's forms directly:

```
python exam_mapper.py FormA.pdf FormB.pdf FormC.pdf FormD.pdf
```

Results are written to a `Result/` folder beside the input, plus an
`extraction_log.txt` summarising the whole run.

### Web app

```
python app.py
```

Then open `http://localhost:5000`, drop in the PDF forms for one exam, and
download the mapping, the issues report, and (if any) the diagnostic log.

The web app imports the parser as `exam_mapper`, so the parser file must be
named `exam_mapper.py` alongside `app.py`, with `index.html` in a `templates/`
folder.

---

## File naming for forms

The form letter is read from the filename, or from the PDF header if the
filename is ambiguous. All of these are recognised:

- `ExamA.pdf`, `Form_B.pdf`, `KEY Exam1 vC.pdf`
- `v001.pdf`, `v002.pdf` (numeric versions)
- `vEarly A.pdf`, `vLATE B.pdf` (early/late sittings)

Forms are assigned letters in order (early → numbered → late), so an exam with
seven versions still maps cleanly to A–G.

---

## What the tool detects

**Cross-form question matching.** Uses a global optimal one-to-one match so two
questions competing for the same counterpart never cause one to be wrongly
reported as missing.

**Choice-order scrambling.** For each form, records how the master choices were
reordered, e.g. `B -> D -> C -> A -> E`.

**Correct answer.** Detected from checkmarks/ticks, bold text, math-bold
glyphs, underlines, or an explicit "Answer: X" note. Multi-answer questions
(where two or more choices are marked correct) are captured as e.g. `A,C`.

**Variants.** Two forms may share a question stem but offer different answer
options (or ask the reverse of the same concept). These are split into separate
variant rows rather than being silently merged.

**OMIT (not-graded / credit-to-all).** When a form's key marks a question
"not graded", "omit", "credit to all", "everyone gets credit", etc., that
question's row is flagged `OMIT = Yes`. See the important caveat below.

---

## Column reference (Excel)

| Column                   | Meaning                                                                      |
| ------------------------ | ---------------------------------------------------------------------------- |
| `OMIT`                   | `Yes` if the question was confidently detected as not graded / credit-to-all |
| `Master Question Number` | The canonical question id (follows the first form's numbering)               |
| `Variant`                | Blank, or `V1 (A/C)` style label when a question has form-specific variants  |
| `Question Prompt`        | The question text                                                            |
| `Choice A`–`Choice E`    | The master answer options                                                    |
| `Form X #`               | That form's number for this question (`MISSING` / `N/A` when absent)         |
| `Form X Choice Order`    | How the master choices were reordered on that form                           |
| `Form X Correct Choice`  | The detected answer on that form (`Unknown` if undetected)                   |

---

## Known limitations

- **Scanned / OCR files** often have garbled text;
  the tool flags them rather than crashing, but their questions may not parse.
  Re-OCR is the fix.
- **Image-answer questions** (structures/diagrams as options) cannot have their
  correct answer read from text; they are flagged for manual review.
- **Floating omit notes** are flagged, not auto-applied.

---
