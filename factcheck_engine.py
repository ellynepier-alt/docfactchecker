import io
import os
import re
import json
import tempfile
from docx import Document
from docx.shared import RGBColor

SUPPORTED_EXTS = {'.txt', '.md', '.docx', '.pdf'}


def load_kb(path='doc_guidelines_kb.json'):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_text_from_uploaded(uploaded_file):
    name = uploaded_file.name
    ext = os.path.splitext(name)[1].lower()
    data = uploaded_file.getvalue()

    if ext in ['.txt', '.md']:
        return data.decode('utf-8', errors='ignore')

    if ext == '.docx':
        doc = Document(io.BytesIO(data))
        chunks = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        chunks.append(cell.text)
        return '\n'.join(chunks)

    if ext == '.pdf':
        import fitz
        doc = fitz.open(stream=data, filetype='pdf')
        return '\n'.join(page.get_text() for page in doc)

    raise ValueError('Unsupported file type. Use PDF, DOCX, TXT, or MD.')


def make_context(text, start, end, width=170):
    lo = max(0, start - width)
    hi = min(len(text), end + width)
    snippet = re.sub(r'\s+', ' ', text[lo:hi]).strip()
    return ('...' if lo > 0 else '') + snippet + ('...' if hi < len(text) else '')


def add_flag(flags, kind, severity, matched, issue, rec, ctx):
    flags.append({
        'kind': kind,
        'severity': severity,
        'matched': matched,
        'issue': issue,
        'rec': rec or '',
        'context': ctx
    })


def run_factcheck(filename, text, kb):
    flags = []
    low = text.lower()
    norm = re.sub(r'\s+', ' ', low)

    # Terminology flags
    for item in kb.get('terminology_flags', []):
        pattern = item['pattern'].lower()
        for m in re.finditer(re.escape(pattern), low):
            add_flag(
                flags,
                'Terminology',
                item['severity'],
                text[m.start():m.end()],
                item['issue'],
                item.get('rec'),
                make_context(text, m.start(), m.end())
            )

    # Contradiction flags
    for item in kb.get('contradiction_flags', []):
        pattern = item['pattern'].lower()
        for m in re.finditer(re.escape(pattern), low):
            add_flag(
                flags,
                'Possible contradiction',
                item['severity'],
                text[m.start():m.end()],
                item['issue'],
                item.get('rec'),
                make_context(text, m.start(), m.end())
            )

    # Amantadine checks
    if 'amantadine' in norm:
        add_flag(
            flags,
            'Key fact to verify',
            'review',
            'amantadine',
            'Confirm that the material states amantadine applies to traumatic VS/UWS or MCS patients 4-16 weeks post injury, at 100-200 mg twice daily, after checking contraindications or case-specific risks.',
            '14',
            'Material mentions amantadine.'
        )
        for m in re.finditer(r'amantadine[^.]{0,90}?(\d{2,4})\s*mg', norm):
            dose = m.group(1)
            if dose not in ['100', '200']:
                add_flag(
                    flags,
                    'Key fact mismatch',
                    'high',
                    f'amantadine {dose} mg',
                    'Guideline dose is amantadine 100-200 mg twice daily.',
                    '14',
                    make_context(norm, m.start(), m.end())
                )

    # 3-month / 12-month chronic threshold checks
    for m in re.finditer(r'(?<!non)(?<!non-)traumatic[^.]{0,100}?3\s*month', norm):
        add_flag(
            flags,
            'Key fact mismatch',
            'high',
            'traumatic ... 3 months',
            'Three months applies to nontraumatic VS/UWS. Traumatic VS/UWS uses 12 months.',
            '7',
            make_context(norm, m.start(), m.end())
        )
    for m in re.finditer(r'non-?traumatic[^.]{0,100}?12\s*month', norm):
        add_flag(
            flags,
            'Key fact mismatch',
            'high',
            'nontraumatic ... 12 months',
            'Twelve months applies to traumatic VS/UWS. Nontraumatic VS/UWS uses 3 months.',
            '7',
            make_context(norm, m.start(), m.end())
        )

    # Evidence-level checks near topic keywords
    topic_levels = {
        'amantadine': ('14', ['B']),
        'crs-r': ('6', ['B']),
        'coma recovery': ('6', ['B']),
        'disability rating': ('5', ['B']),
        'multidisciplinary rehabilitation': ('1', ['B']),
        'patient and family preferences': ('11', ['A']),
        'goals of care': ('9', ['A']),
        'pain': ('13', ['B'])
    }
    for m in re.finditer(r'Level\s+([ABCU])\b', text, re.IGNORECASE):
        cited = m.group(1).upper()
        win_lo = max(0, m.start() - 240)
        win = low[win_lo:m.end() + 70]
        nearest = None
        for keyword, (rec_id, levels) in topic_levels.items():
            pos = win.rfind(keyword)
            if pos != -1:
                dist = abs((win_lo + pos) - m.start())
                if nearest is None or dist < nearest[0]:
                    nearest = (dist, keyword, rec_id, levels)
        if nearest:
            _, keyword, rec_id, levels = nearest
            if cited not in levels:
                add_flag(
                    flags,
                    'Evidence-level mismatch',
                    'high',
                    f'Level {cited} near {keyword}',
                    f'Material cites Level {cited}, but guideline recommendation {rec_id} is Level {"/".join(levels)}.',
                    rec_id,
                    make_context(text, m.start(), m.end())
                )

    # Coverage map
    coverage = []
    for rec in kb.get('recommendations', []):
        keywords = [w.lower() for w in rec['topic'].replace('/', ' ').split() if len(w) > 4]
        if any(k in low for k in keywords):
            coverage.append(rec['id'])

    severity_order = {'high': 0, 'medium': 1, 'low': 2, 'review': 3}
    flags.sort(key=lambda f: severity_order.get(f['severity'], 9))

    return {
        'filename': filename,
        'text_length': len(text),
        'flags': flags,
        'coverage': sorted(set(coverage), key=lambda x: (int(re.sub(r'\D', '', x) or 0), x))
    }


def build_word_report(result, kb):
    doc = Document()
    doc.add_heading('DoC Guideline Fact-Check Report', level=0)
    doc.add_paragraph(f"Material checked: {result['filename']}")
    doc.add_paragraph(f"Guideline: {kb['meta']['title']} ({kb['meta']['year']})")
    doc.add_paragraph(kb['meta']['citation'])
    doc.add_paragraph('This is a human-in-the-loop educational review aid. Flags identify candidates for review, not final clinical judgments.')

    counts = {'high': 0, 'medium': 0, 'low': 0, 'review': 0}
    for flag in result['flags']:
        counts[flag['severity']] = counts.get(flag['severity'], 0) + 1

    doc.add_heading('Summary', level=1)
    doc.add_paragraph(f"High: {counts['high']} | Medium: {counts['medium']} | Low: {counts['low']} | Review: {counts['review']}")

    doc.add_heading('Flags', level=1)
    if not result['flags']:
        doc.add_paragraph('No flags raised. Human review is still recommended.')

    colors = {
        'high': RGBColor(192, 0, 0),
        'medium': RGBColor(199, 106, 0),
        'low': RGBColor(127, 106, 0),
        'review': RGBColor(31, 78, 121)
    }

    for flag in result['flags']:
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(f"[{flag['severity'].upper()}] {flag['kind']}")
        run.bold = True
        run.font.color.rgb = colors.get(flag['severity'], RGBColor(0, 0, 0))
        p.add_run(f" — matched: {flag['matched']}")
        if flag['rec']:
            p.add_run(f" (Rec {flag['rec']})")
        doc.add_paragraph('Issue: ' + flag['issue'])
        if flag['context']:
            context_para = doc.add_paragraph('Context: ' + flag['context'])
            context_para.runs[0].italic = True

    doc.add_heading('Recommendation Coverage Map', level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Light Grid Accent 1'
    header = table.rows[0].cells
    header[0].text = 'Rec'
    header[1].text = 'Topic'
    header[2].text = 'Level'
    header[3].text = 'Touched?'

    for rec in kb.get('recommendations', []):
        row = table.add_row().cells
        row[0].text = rec['id']
        row[1].text = rec['topic']
        row[2].text = '/'.join(rec['level'])
        row[3].text = 'Yes' if rec['id'] in result['coverage'] else ''

    doc.add_heading('Appendix: Guideline Recommendation Wording', level=1)
    for rec in kb.get('recommendations', []):
        p = doc.add_paragraph()
        p.add_run(f"Recommendation {rec['id']} (Level {'/'.join(rec['level'])}): ").bold = True
        p.add_run(rec['text'])

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output
