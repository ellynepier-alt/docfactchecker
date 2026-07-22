import os, re, json, tempfile
from docx import Document
from docx.shared import RGBColor
from werkzeug.utils import secure_filename

SUPPORTED_EXTS = {'.txt', '.md', '.docx', '.pdf', '.pptx'}


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
    if ext == '.pptx':
        from pptx import Presentation
        prs = Presentation(path)
        chunks = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        text = ''.join(run.text for run in para.runs)
                        if text.strip():
                            chunks.append(text)
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text.strip():
                                chunks.append(cell.text)
                if shape.has_chart:
                    continue
        return '\n'.join(chunks)
    raise ValueError('Unsupported file type')


def context(text, start, end, width=160):
    lo = max(0, start - width); hi = min(len(text), end + width)
    out = re.sub(r'\s+', ' ', text[lo:hi]).strip()
    return ('...' if lo else '') + out + ('...' if hi < len(text) else '')


def add_flag(flags, kind, severity, matched, issue, rec, ctx):
    flags.append({'kind': kind, 'severity': severity, 'matched': matched, 'issue': issue, 'rec': rec or '', 'context': ctx})


def count_syllables(word):
    word = word.lower()
    vowels = 'aeiouy'
    count = 0
    prev_vowel = False
    for ch in word:
        is_vowel = ch in vowels
        if is_vowel and not prev_vowel:
            count += 1
        prev_vowel = is_vowel
    if word.endswith('e') and count > 1:
        count -= 1
    return max(1, count)


ACRONYM_GLOSSARY = {
    'VS/UWS': 'Vegetative State/Unresponsive Wakefulness Syndrome',
    'MCS': 'Minimally Conscious State',
    'DoC': 'Disorders of Consciousness',
    'CRS-R': 'Coma Recovery Scale-Revised',
    'DRS': 'Disability Rating Scale',
    'TBI': 'Traumatic Brain Injury',
    'SPECT': 'Single Photon Emission Computed Tomography',
    'PET': 'Positron Emission Tomography',
    'fMRI': 'functional Magnetic Resonance Imaging',
    'EEG': 'Electroencephalography',
    'ERP': 'Event-Related Potential',
    'SEP': 'Somatosensory Evoked Potential',
    'TMS': 'Transcranial Magnetic Stimulation',
    'PCI': 'Perturbational Complexity Index',
    'EMG': 'Electromyography',
    'MOLST': 'Medical Orders for Life-Sustaining Treatment',
}


def analyze_clarity(text):
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', text) if s.strip()]
    words = re.findall(r"[A-Za-z']+", text)
    num_sentences = max(1, len(sentences))
    num_words = max(1, len(words))
    syllables = sum(count_syllables(w) for w in words)

    avg_sentence_len = num_words / num_sentences
    avg_syllables_per_word = syllables / num_words

    flesch_reading_ease = 206.835 - 1.015 * avg_sentence_len - 84.6 * avg_syllables_per_word
    flesch_kincaid_grade = 0.39 * avg_sentence_len + 11.8 * avg_syllables_per_word - 15.59

    long_sentences = []
    for s in sentences:
        wc = len(re.findall(r"[A-Za-z']+", s))
        if wc > 30:
            snippet = s.strip()
            if len(snippet) > 220:
                snippet = snippet[:220] + '...'
            long_sentences.append({'text': snippet, 'word_count': wc})
    long_sentences = long_sentences[:5]

    undefined_acronyms = []
    for acr, expansion in ACRONYM_GLOSSARY.items():
        if re.search(r'\b' + re.escape(acr) + r'\b', text):
            key_words = [w for w in re.findall(r'[A-Za-z]+', expansion) if len(w) > 3]
            found_expansion = any(re.search(re.escape(w), text, re.IGNORECASE) for w in key_words) if key_words else False
            if not found_expansion:
                undefined_acronyms.append({'acronym': acr, 'expansion': expansion})

    passive_matches = re.findall(r'\b(?:is|are|was|were|be|been|being)\s+\w+ed\b', text, re.IGNORECASE)
    passive_count = len(passive_matches)
    passive_ratio = passive_count / num_sentences

    suggestions = []
    if flesch_kincaid_grade > 14:
        suggestions.append(
            f"This material reads at roughly a {flesch_kincaid_grade:.1f} grade level (college and above). "
            "Consider simplifying sentence structure and terminology, especially if patients or families will read it."
        )
    elif flesch_kincaid_grade > 10:
        suggestions.append(
            f"This material reads at roughly a {flesch_kincaid_grade:.1f} grade level (high school). "
            "Reasonable for a clinician audience, but may still be dense for family-facing materials."
        )

    if long_sentences:
        suggestions.append(
            f"{len(long_sentences)} sentence(s) exceed 30 words. Consider breaking these into shorter sentences for clarity."
        )

    if undefined_acronyms:
        names = ', '.join(a['acronym'] for a in undefined_acronyms[:6])
        suggestions.append(
            f"These clinical acronyms appear without being spelled out: {names}. "
            "Consider defining them on first use for readers unfamiliar with DoC terminology."
        )

    if passive_ratio > 0.5:
        suggestions.append(
            "This material relies heavily on passive voice (e.g., 'was performed' rather than 'the clinician performed'). "
            "Active voice is often clearer, especially for family-facing materials."
        )

    if not suggestions:
        suggestions.append('No major clarity issues detected. Language and structure appear reasonably accessible.')

    return {
        'flesch_reading_ease': round(flesch_reading_ease, 1),
        'flesch_kincaid_grade': round(flesch_kincaid_grade, 1),
        'avg_sentence_length': round(avg_sentence_len, 1),
        'long_sentences': long_sentences,
        'undefined_acronyms': undefined_acronyms,
        'passive_voice_count': passive_count,
        'suggestions': suggestions,
    }


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

    if 'crs-r' in norm:
        for m in re.finditer(r'crs-r[^.]{0,60}?score[^.]{0,20}?(?:of\s*)?(\d{1,2})', norm):
            score = int(m.group(1))
            if score < 6:
                add_flag(flags, 'Key fact mismatch', 'high', f'CRS-R score {score}', 'Guideline associates CRS-R scores of 6 or higher (>1 month after onset) with increased likelihood of recovery in nontraumatic post-anoxic VS/UWS.', '6', context(norm, m.start(), m.end()))

    for m in re.finditer(r'\bdrs\b[^.]{0,80}?(\d{1,2})\s*(?:-|to)\s*(\d{1,2})\s*month', norm):
        lo, hi = int(m.group(1)), int(m.group(2))
        if (lo, hi) != (2, 3):
            add_flag(flags, 'Key fact mismatch', 'medium', f'DRS at {lo}-{hi} months', 'Guideline specifies the DRS should be performed at 2-3 months post injury for traumatic VS/UWS.', '5', context(norm, m.start(), m.end()))

    for m in re.finditer(r'\bmri\b[^.]{0,80}?(\d{1,2})\s*(?:-|to)\s*(\d{1,2})\s*week', norm):
        lo, hi = int(m.group(1)), int(m.group(2))
        if (lo, hi) != (6, 8):
            add_flag(flags, 'Key fact mismatch', 'medium', f'MRI at {lo}-{hi} weeks', 'Guideline specifies MRI should be performed 6-8 weeks post injury in traumatic VS/UWS.', '5', context(norm, m.start(), m.end()))

    for m in re.finditer(r'\bspect\b[^.]{0,80}?(\d{1,2})\s*(?:-|to)\s*(\d{1,2})\s*month', norm):
        lo, hi = int(m.group(1)), int(m.group(2))
        if (lo, hi) != (1, 2):
            add_flag(flags, 'Key fact mismatch', 'medium', f'SPECT at {lo}-{hi} months', 'Guideline specifies SPECT should be performed 1-2 months post injury in traumatic VS/UWS.', '5', context(norm, m.start(), m.end()))

    topic_levels = {
        'amantadine': ('14', ['B']),
        'crs-r': ('6', ['B']),
        'coma recovery': ('6', ['B']),
        'sep': ('6', ['C']),
        'multidisciplinary rehabilitation': ('1', ['B']),
        'patient and family preferences': ('11', ['A']),
        'goals of care': ('9', ['A']),
        'pain': ('13', ['B']),
        'serial standardized': ('4', ['B']),
        'drs': ('5', ['B']),
        'spect': ('5', ['B']),
        'chronic phase': ('10', ['B'])
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
    return {
        'filename': os.path.basename(filepath),
        'text_length': len(text),
        'flags': flags,
        'coverage': sorted(set(coverage)),
        'clarity': analyze_clarity(text),
    }


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
