import os, re, json, tempfile
from docx import Document
from docx.shared import RGBColor
from werkzeug.utils import secure_filename

SUPPORTED_EXTS = {'.txt', '.md', '.docx', '.pdf'}


def load_kb(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ['.txt', '.md']:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    if ext == '.docx':
        doc = Document(path)
        chunks = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        chunks.append(cell.text)
        return '\n'.join(chunks)
    if ext == '.pdf':
        import fitz
        doc = fitz.open(path)
        return '\n'.join(page.get_text() for page in doc)
    raise ValueError('Unsupported file type')


def context(text, start, end, width=160):
    lo = max(0, start - width); hi = min(len(text), end + width)
    out = re.sub(r'\s+', ' ', text[lo:hi]).strip()
    return ('...' if lo else '') + out + ('...' if hi < len(text) else '')


def add_flag(flags, kind, severity, matched, issue, rec, ctx):
    flags.append({'kind': kind, 'severity': severity, 'matched': matched, 'issue': issue, 'rec': rec or '', 'context': ctx})


def run_checks(filepath, kb):
    text = extract_text(filepath)
    flags = []
    low = text.lower()
    norm = re.sub(r'\s+', ' ', low)

    for f in kb['terminology_flags']:
        for m in re.finditer(re.escape(f['pattern'].lower()), low):
            add_flag(flags, 'Terminology', f['severity'], text[m.start():m.end()], f['issue'], f.get('rec'), context(text, m.start(), m.end()))

    for f in kb['contradiction_flags']:
        for m in re.finditer(re.escape(f['pattern'].lower()), low):
            add_flag(flags, 'Possible contradiction', f['severity'], text[m.start():m.end()], f['issue'], f.get('rec'), context(text, m.start(), m.end()))

    if 'amantadine' in norm:
        for m in re.finditer(r'amantadine[^.]{0,80}?(\d{2,4})\s*mg', norm):
            dose = m.group(1)
            if dose not in ['100', '200']:
                add_flag(flags, 'Key fact mismatch', 'high', f'amantadine {dose} mg', 'Guideline dose is amantadine 100-200 mg twice daily.', '14', context(norm, m.start(), m.end()))
        add_flag(flags, 'Key fact to verify', 'review', 'amantadine', 'Confirm the material states: amantadine applies to traumatic VS/UWS or MCS patients 4-16 weeks post injury, 100-200 mg twice daily.', '14', 'Material mentions amantadine.')

    for m in re.finditer(r'(?<!non)(?<!non-)traumatic[^.]{0,90}?3\s*month', norm):
        add_flag(flags, 'Key fact mismatch', 'high', 'traumatic ... 3 months', 'Three months applies to nontraumatic VS/UWS; traumatic VS/UWS uses 12 months.', '7', context(norm, m.start(), m.end()))
    for m in re.finditer(r'non-?traumatic[^.]{0,90}?12\s*month', norm):
        add_flag(flags, 'Key fact mismatch', 'high', 'nontraumatic ... 12 months', 'Twelve months applies to traumatic VS/UWS; nontraumatic VS/UWS uses 3 months.', '7', context(norm, m.start(), m.end()))

    topic_levels = {
        'amantadine': ('14', ['B']),
        'crs-r': ('6', ['B']),
        'coma recovery': ('6', ['B']),
        'multidisciplinary rehabilitation': ('1', ['B']),
        'patient and family preferences': ('11', ['A']),
        'goals of care': ('9', ['A']),
        'pain': ('13', ['B'])
    }
    for m in re.finditer(r'Level\s+([ABCU])\b', text, re.IGNORECASE):
        cited = m.group(1).upper()
        win_lo = max(0, m.start() - 220)
        win = low[win_lo:m.end()+50]
        best = None
        for kw, (rid, levels) in topic_levels.items():
            pos = win.rfind(kw)
            if pos != -1:
                dist = abs((win_lo + pos) - m.start())
                if best is None or dist < best[0]:
                    best = (dist, kw, rid, levels)
        if best and cited not in best[3]:
            add_flag(flags, 'Evidence-level mismatch', 'high', f'Level {cited} near {best[1]}', f'Material cites Level {cited}, but guideline recommendation {best[2]} is Level {"/".join(best[3])}.', best[2], context(text, m.start(), m.end()))

    coverage = []
    for rec in kb['recommendations']:
        topic_words = [w.lower() for w in rec['topic'].replace('/', ' ').split() if len(w) > 4]
        if any(w in low for w in topic_words):
            coverage.append(rec['id'])

    order = {'high':0, 'medium':1, 'low':2, 'review':3}
    flags.sort(key=lambda x: order.get(x['severity'], 9))
    return {'filename': os.path.basename(filepath), 'text_length': len(text), 'flags': flags, 'coverage': sorted(set(coverage))}


def make_report(result, kb, out_path):
    doc = Document()
    doc.add_heading('DoC Guideline Fact-Check Report', level=0)
    doc.add_paragraph(f"Material checked: {result['filename']}")
    doc.add_paragraph(f"Guideline: {kb['meta']['title']} ({kb['meta']['year']})")
    doc.add_paragraph(kb['meta']['citation'])
    doc.add_paragraph('This is a human-in-the-loop review aid. Flags identify candidates for review, not final clinical judgments.')

    counts = {s: 0 for s in ['high', 'medium', 'low', 'review']}
    for f in result['flags']:
        counts[f['severity']] = counts.get(f['severity'], 0) + 1
    doc.add_heading('Summary', level=1)
    doc.add_paragraph(f"High: {counts['high']} | Medium: {counts['medium']} | Low: {counts['low']} | Review: {counts['review']}")

    doc.add_heading('Flags', level=1)
    if not result['flags']:
        doc.add_paragraph('No flags raised. Human review is still recommended.')
    colors = {'high': RGBColor(192,0,0), 'medium': RGBColor(199,106,0), 'low': RGBColor(127,106,0), 'review': RGBColor(31,78,121)}
    for f in result['flags']:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(f"[{f['severity'].upper()}] {f['kind']}")
        r.bold = True; r.font.color.rgb = colors.get(f['severity'], RGBColor(0,0,0))
        p.add_run(f" — matched: {f['matched']}")
        if f['rec']:
            p.add_run(f" (Rec {f['rec']})")
        doc.add_paragraph('Issue: ' + f['issue'])
        if f['context']:
            c = doc.add_paragraph('Context: ' + f['context'])
            c.runs[0].italic = True

    doc.add_heading('Recommendation Coverage Map', level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0].cells
    hdr[0].text = 'Rec'; hdr[1].text = 'Topic'; hdr[2].text = 'Level'; hdr[3].text = 'Touched?'
    for rec in kb['recommendations']:
        row = table.add_row().cells
        row[0].text = rec['id']; row[1].text = rec['topic']; row[2].text = '/'.join(rec['level']); row[3].text = 'Yes' if rec['id'] in result['coverage'] else ''

    doc.add_heading('Appendix: Guideline Recommendation Wording', level=1)
    for rec in kb['recommendations']:
        p = doc.add_paragraph()
        p.add_run(f"Recommendation {rec['id']} (Level {'/'.join(rec['level'])}): ").bold = True
        p.add_run(rec['text'])
    doc.save(out_path)
