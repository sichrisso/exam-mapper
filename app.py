import os
import uuid
import json
import logging
import tempfile
import shutil
import io
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    # Import exactly what the current script exposes. The current exam_mapper
    # additionally provides descriptive naming and the issues-report writer.
    from exam_mapper import (
        process_pdfs, assign_form_letters,
        _descriptive_basename, write_issues_report,
    )

    uploaded_files = request.files.getlist('pdfs')
    if not uploaded_files or all(f.filename == '' for f in uploaded_files):
        return jsonify({'error': 'No files uploaded'}), 400

    job_dir = tempfile.mkdtemp(prefix='exam_job_')

    # Capture this request's exam_mapper diagnostic log into an in-memory
    # buffer so the JSON response can carry the full trace for this run.
    _log_buffer = io.StringIO()
    _capture_handler = logging.StreamHandler(_log_buffer)
    _capture_handler.setFormatter(
        logging.Formatter("[%(asctime)s] %(levelname)-7s %(message)s",
                          datefmt="%H:%M:%S")
    )
    _exam_logger = logging.getLogger("exam_mapper")
    _exam_logger.addHandler(_capture_handler)

    try:
        # Save all uploaded PDFs, keeping original filenames
        saved_names = []
        for f in uploaded_files:
            if not f or f.filename == '':
                continue
            if not f.filename.lower().endswith('.pdf'):
                return jsonify({'error': f'"{f.filename}" is not a PDF'}), 400
            safe_name = secure_filename(f.filename)
            f.save(os.path.join(job_dir, safe_name))
            saved_names.append(safe_name)

        # Assign form letters across the WHOLE batch at once.
        letter_map, unrecognised = assign_form_letters(saved_names)

        if unrecognised:
            return jsonify({
                'error': (
                    f'Could not determine the form letter for: {", ".join(unrecognised)}. '
                    'The form letter is read from the filename (e.g. ExamA.pdf, '
                    'Form_B.pdf, v001.pdf, vEARLY.pdf) or from inside the PDF header.'
                )
            }), 400

        pdf_map = {
            letter: os.path.join(job_dir, fname)
            for fname, letter in letter_map.items()
        }

        if not pdf_map:
            return jsonify({'error': 'No valid PDF files were uploaded'}), 400

        out_path = os.path.join(job_dir, 'Exam_Mapping.xlsx')
        log_entries = []

        # process_pdfs now RETURNS the mapping dataframe — needed for the
        # issues report (which scans it for MISSING / Unknown cells).
        df_out = process_pdfs(
            pdf_map=pdf_map,
            meta_override={},
            answer_key=None,
            out_path=out_path,
            log=log_entries,
        )

        if not os.path.exists(out_path):
            return jsonify({'error': 'Processing completed but no output was generated'}), 500

        # Build a descriptive base name from detected metadata, e.g.
        # "Spring-2016_Exam-3_Result" (falls back gracefully if unknown).
        meta = {'semester': log_entries[0].get('semester', '')} if log_entries else None
        exam_label = (log_entries[0].get('exam') or 'Exam') if log_entries else 'Exam'
        sem_label = (log_entries[0].get('semester') or 'Result') if log_entries else 'Result'
        base = _descriptive_basename(sem_label, exam_label, meta)

        token = str(uuid.uuid4())
        result_dir = os.path.join(tempfile.gettempdir(), f'exam_result_{token}')
        os.makedirs(result_dir, exist_ok=True)

        xlsx_name = f'{base}.xlsx'
        report_name = f'{base}_ISSUES.txt'
        shutil.copy(out_path, os.path.join(result_dir, xlsx_name))

        # Write the downloadable issues report next to the Excel, scanning the
        # generated mapping for every MISSING / Unknown cell plus all per-form
        # parse warnings, jumps, and errors.
        had_issues = False
        try:
            had_issues = write_issues_report(
                os.path.join(result_dir, report_name),
                sem_label, exam_label, log_entries,
                df=df_out, available_forms=sorted(pdf_map.keys()),
            )
        except Exception as _re:
            # Report writing must never block the result; note it and move on.
            with open(os.path.join(result_dir, report_name), 'w') as rf:
                rf.write(f"(could not generate issues report: {_re})\n")

        # Also persist the raw diagnostic log as its own downloadable file.
        diag_text = _log_buffer.getvalue()
        with open(os.path.join(result_dir, 'diagnostic_log.txt'), 'w') as lf:
            lf.write(diag_text)

        # Remember the real filenames for the download routes.
        with open(os.path.join(result_dir, 'names.json'), 'w') as nf:
            json.dump({'xlsx': xlsx_name, 'report': report_name}, nf)

        return jsonify({
            'token': token,
            'summary': _build_summary(log_entries, had_issues),
            'log': log_entries,
            'diagnostic_log': diag_text,
            'has_issues_report': True,
            'xlsx_name': xlsx_name,
            'report_name': report_name,
        })

    except Exception as e:
        return jsonify({
            'error': f'Processing failed: {str(e)}',
            'diagnostic_log': _log_buffer.getvalue(),
        }), 500
    finally:
        _exam_logger.removeHandler(_capture_handler)
        shutil.rmtree(job_dir, ignore_errors=True)


def _result_dir(token):
    uuid.UUID(token)  # raises ValueError on bad token
    return os.path.join(tempfile.gettempdir(), f'exam_result_{token}')


@app.route('/download/<token>')
def download(token):
    """Download the Excel mapping (uses the descriptive name)."""
    try:
        rd = _result_dir(token)
    except ValueError:
        return jsonify({'error': 'Invalid token'}), 400

    names_path = os.path.join(rd, 'names.json')
    xlsx_name = 'Exam_Mapping.xlsx'
    if os.path.exists(names_path):
        try:
            xlsx_name = json.load(open(names_path)).get('xlsx', xlsx_name)
        except Exception:
            pass

    out_file = os.path.join(rd, xlsx_name)
    if not os.path.exists(out_file):
        return jsonify({'error': 'File not found or expired'}), 404

    return send_file(
        out_file, as_attachment=True, download_name=xlsx_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/download/<token>/report')
def download_report(token):
    """Download the issues report (.txt) to keep alongside the Excel."""
    try:
        rd = _result_dir(token)
    except ValueError:
        return jsonify({'error': 'Invalid token'}), 400

    names_path = os.path.join(rd, 'names.json')
    report_name = 'Exam_ISSUES.txt'
    if os.path.exists(names_path):
        try:
            report_name = json.load(open(names_path)).get('report', report_name)
        except Exception:
            pass

    out_file = os.path.join(rd, report_name)
    if not os.path.exists(out_file):
        return jsonify({'error': 'Report not found or expired'}), 404

    return send_file(out_file, as_attachment=True,
                     download_name=report_name, mimetype='text/plain')


@app.route('/download/<token>/log')
def download_log(token):
    """Download the raw diagnostic log (.txt) for this run."""
    try:
        rd = _result_dir(token)
    except ValueError:
        return jsonify({'error': 'Invalid token'}), 400

    out_file = os.path.join(rd, 'diagnostic_log.txt')
    if not os.path.exists(out_file):
        return jsonify({'error': 'Log not found or expired'}), 404

    return send_file(out_file, as_attachment=True,
                     download_name='diagnostic_log.txt', mimetype='text/plain')


def _build_summary(log_entries, had_issues=False):
    statuses = [e.get('status', 'OK') for e in log_entries]
    if any(s == 'ERROR' for s in statuses):
        overall = 'error'
    elif any(s in ('CHECK', 'EMPTY') for s in statuses):
        overall = 'warning'
    else:
        overall = 'success'

    # Surface jumps as their own count so the UI can call them out.
    total_jumps = sum(len(e.get('jumps', [])) for e in log_entries)

    return {
        'overall':         overall,
        'forms_parsed':    [e['form'] for e in log_entries],
        'total_questions': sum(e.get('questions', 0) for e in log_entries),
        'total_unknown':   sum(e.get('unknown', 0) for e in log_entries),
        'total_warnings':  sum(len(e.get('warnings', [])) for e in log_entries),
        'total_errors':    sum(len(e.get('errors', [])) for e in log_entries),
        'total_jumps':     total_jumps,
        'has_issues':      had_issues,
        'semester':        log_entries[0].get('semester', '') if log_entries else '',
        'exam':            log_entries[0].get('exam', '') if log_entries else '',
    }


if __name__ == '__main__':
    app.run(debug=True, port=5000)