import os, re, json, tempfile
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn

SUPPORTED_EXTS = {'.txt', '.md', '.docx', '.pdf', '.pptx'}


def load_kb(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def ocr_image_bytes(image_bytes):
    """Run OCR on raw image bytes; return empty string on any failure (corrupt/unsupported image, no OCR engine, etc.)."""
    try:
        import pytesseract
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(image_bytes))
        return pytesseract.image_to_string(img).strip()
    except Exception:
        return ''


def get_docx_images(doc):
    """Return [{'alt': str|None, 'ocr_text': str}] for every image in the document,
    including images inside floating text boxes/diagrams (which live under
    wp:anchor rather than wp:inline)."""
    images = []
    for blip in doc.element.body.iter(qn('a:blip')):
        rId = blip.get(qn('r:embed'))
        if not rId:
            continue
        # Walk up to the wp:inline or wp:anchor container, which holds the docPr (alt text) sibling.
        container = blip
        alt = None
        for _ in range(10):
            container = container.getparent()
            if container is None:
                break
            tag = container.tag.split('}')[-1]
            if tag in ('inline', 'anchor'):
                alt = _xml_descendant_attr(container, 'docPr', 'descr') or _xml_descendant_attr(container, 'docPr', 'title')
                break
        ocr_text = ''
        try:
            part = doc.part.related_parts[rId]
            ocr_text = ocr_image_bytes(part.blob)
        except Exception:
            pass
        images.append({'alt': alt, 'ocr_text': ocr_text})
    return images


def get_pptx_images(prs):
    """Return [{'alt': str|None, 'ocr_text': str}] for every picture shape, recursing into groups."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    images = []

    def walk(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                alt = _xml_descendant_attr(shape._element, 'cNvPr', 'descr')
                ocr_text = ''
                try:
                    ocr_text = ocr_image_bytes(shape.image.blob)
                except Exception:
                    pass
                images.append({'alt': alt, 'ocr_text': ocr_text})
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(shape.shapes)

    for slide in prs.slides:
        walk(slide.shapes)
    return images


def get_pdf_images(path):
    """Return [str] of OCR'd text for every embedded image in a PDF (no alt-text concept in plain PDF)."""
    import fitz
    texts = []
    doc = fitz.open(path)
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                base = doc.extract_image(xref)
                ocr_text = ocr_image_bytes(base['image'])
                if ocr_text:
                    texts.append(ocr_text)
            except Exception:
                continue
    return texts


def extract_docx_textboxes(doc):
    """Text inside floating text boxes (e.g., diagram/decision-tree shapes) lives in
    <w:txbxContent> elements nested inside drawings, which doc.paragraphs/doc.tables
    never reach. Walk the XML directly to pick these up."""
    chunks = []
    for node in doc.element.body.iter():
        tag = node.tag.split('}')[-1] if isinstance(node.tag, str) else ''
        if tag != 'txbxContent':
            continue
        for p in node.findall('.//' + qn('w:p')):
            line = ''.join(t.text or '' for t in p.iter(qn('w:t')))
            if line.strip():
                chunks.append(line)
    return chunks


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
        chunks.extend(extract_docx_textboxes(doc))
        chunks.extend(img['ocr_text'] for img in get_docx_images(doc) if img['ocr_text'])
        return '\n'.join(chunks)
    if ext == '.pdf':
        import fitz
        doc = fitz.open(path)
        chunks = [page.get_text() for page in doc]
        chunks.extend(get_pdf_images(path))
        return '\n'.join(chunks)
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
        chunks.extend(img['ocr_text'] for img in get_pptx_images(prs) if img['ocr_text'])
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


def _xml_descendant_attr(element, local_tag, attr):
    """Find first descendant with the given local (namespace-stripped) tag and return an attribute."""
    for node in element.iter():
        tag = node.tag.split('}')[-1] if isinstance(node.tag, str) else ''
        if tag == local_tag:
            return node.get(attr)
    return None


def assess_image_alt_quality(images):
    """images: list of {'alt': str|None, 'ocr_text': str}. Returns (total, missing, poor, feedback[])."""
    generic_terms = {'image', 'picture', 'photo', 'graphic', 'img', 'untitled', 'diagram', 'chart', 'picture1', 'graphic1'}
    total = len(images)
    missing = 0
    poor = 0
    feedback = []
    for im in images:
        alt = (im.get('alt') or '').strip()
        ocr = (im.get('ocr_text') or '').strip()
        if not alt:
            missing += 1
            if ocr:
                snippet = ocr[:150].replace('\n', ' ')
                feedback.append(f'Missing alt text on an image that contains readable text: "{snippet}" — consider using this (or a summary of it) as the alt text.')
            else:
                feedback.append('Missing alt text on an image with no machine-readable text detected — add a description of what the image depicts.')
        elif alt.lower().strip('.') in generic_terms or len(alt) < 4:
            poor += 1
            feedback.append(f'Alt text "{alt}" is too generic to convey meaning to screen-reader users.')
        elif ocr and len(ocr) > 15:
            ocr_words = set(re.findall(r'[a-z]{4,}', ocr.lower()))
            alt_words = set(re.findall(r'[a-z]{4,}', alt.lower()))
            if ocr_words and not (ocr_words & alt_words):
                poor += 1
                preview = ocr[:100].replace('\n', ' ')
                feedback.append(f'Alt text "{alt}" does not appear to reflect the visible text in the image ("{preview}..."). Screen-reader users may miss this information entirely.')
    return total, missing, poor, feedback


def check_docx_link_text(doc):
    generic_texts = {'click here', 'here', 'more', 'read more', 'link', 'this link', 'click', 'learn more', 'more info', 'info'}
    total_links = 0
    bad_links = []
    for hyperlink in doc.element.body.iter(qn('w:hyperlink')):
        total_links += 1
        text = ''.join(t.text or '' for t in hyperlink.iter(qn('w:t'))).strip()
        if text.lower() in generic_texts:
            bad_links.append(text)
    return total_links, bad_links


def check_accessibility_docx(path):
    findings = []
    doc = Document(path)

    images = get_docx_images(doc)
    total_images, missing_alt, poor_alt, alt_feedback = assess_image_alt_quality(images)
    if total_images:
        compliant_images = total_images - missing_alt - poor_alt
        pct = round(100 * compliant_images / total_images)
        findings.append({
            'check': 'Image alternative text (OCR-verified)', 'wcag': '1.1.1 Non-text Content (Level A)',
            'category': 'Images', 'percent': pct,
            'status': 'fail' if (missing_alt or poor_alt) else 'pass',
            'detail': f'{compliant_images} of {total_images} image(s) ({pct}%) have adequate alt text; {missing_alt} missing, {poor_alt} generic or mismatched with the image\'s actual visible content (checked via OCR).',
        })
        for fb in alt_feedback[:8]:
            findings.append({'check': 'Alt text feedback', 'wcag': '1.1.1 Non-text Content (Level A)', 'status': 'warn', 'detail': fb})
    else:
        findings.append({'check': 'Image alternative text', 'wcag': '1.1.1 Non-text Content (Level A)', 'category': 'Images', 'percent': None, 'status': 'na', 'detail': 'No images found in this document.'})

    heading_used = any(p.style and p.style.name.startswith('Heading') and p.text.strip() for p in doc.paragraphs)
    real_headings = sum(1 for p in doc.paragraphs if p.style and p.style.name.startswith('Heading') and p.text.strip())
    fake_headings = 0
    for p in doc.paragraphs:
        if p.style and p.style.name.startswith('Heading'):
            continue
        txt = p.text.strip()
        if txt and len(txt) < 80 and p.runs and all(r.bold for r in p.runs if r.text.strip()):
            fake_headings += 1
    heading_total = real_headings + fake_headings
    heading_pct = round(100 * real_headings / heading_total) if heading_total else (100 if heading_used else None)
    findings.append({
        'check': 'Heading styles used for structure', 'wcag': '1.3.1 Info and Relationships (A) / 2.4.6 Headings and Labels (AA)',
        'category': 'Headings & structure', 'percent': heading_pct,
        'status': 'pass' if heading_used else 'fail',
        'detail': ('Document uses Word Heading styles, which screen readers rely on for section navigation.' if heading_used
                   else 'No paragraphs use Word Heading styles, so screen-reader users cannot navigate by section.'),
    })
    if fake_headings:
        findings.append({
            'check': 'Bold text used in place of headings', 'wcag': '1.3.1 Info and Relationships (Level A)',
            'status': 'warn', 'detail': f'{fake_headings} short bold line(s) look like section titles but are not tagged with a Heading style, so they are invisible to screen-reader navigation.',
        })

    tables_total = len(doc.tables)
    tables_missing_header = 0
    for t in doc.tables:
        has_header_flag = False
        if t.rows:
            tr = t.rows[0]._tr
            trPr = tr.find(qn('w:trPr'))
            if trPr is not None and trPr.find(qn('w:tblHeader')) is not None:
                has_header_flag = True
        if not has_header_flag:
            tables_missing_header += 1
    if tables_total:
        tables_compliant = tables_total - tables_missing_header
        pct = round(100 * tables_compliant / tables_total)
        findings.append({
            'check': 'Table header rows', 'wcag': '1.3.1 Info and Relationships (Level A)',
            'category': 'Tables', 'percent': pct,
            'status': 'fail' if tables_missing_header else 'pass',
            'detail': f'{tables_compliant} of {tables_total} table(s) ({pct}%) have a designated header row; {tables_missing_header} do not, so screen readers cannot announce column context for those data cells.',
        })
    else:
        findings.append({'check': 'Table header rows', 'wcag': '1.3.1 Info and Relationships (Level A)', 'category': 'Tables', 'percent': None, 'status': 'na', 'detail': 'No tables found in this document.'})

    total_links, bad_links = check_docx_link_text(doc)
    if total_links:
        good_links = total_links - len(bad_links)
        pct = round(100 * good_links / total_links)
        findings.append({
            'check': 'Meaningful link text', 'wcag': '2.4.4 Link Purpose in Context (Level A)',
            'category': 'Links', 'percent': pct,
            'status': 'fail' if bad_links else 'pass',
            'detail': (f'{good_links} of {total_links} link(s) ({pct}%) use descriptive text; {len(bad_links)} use generic text like "{bad_links[0]}" that gives no context out of place — screen-reader users often navigate by a list of links alone.' if bad_links
                       else f'All {total_links} link(s) ({pct}%) use descriptive text.'),
        })
    else:
        findings.append({'check': 'Meaningful link text', 'wcag': '2.4.4 Link Purpose in Context (Level A)', 'category': 'Links', 'percent': None, 'status': 'na', 'detail': 'No hyperlinks found in this document.'})

    return findings


def check_accessibility_pptx(path):
    from pptx import Presentation

    findings = []
    prs = Presentation(path)
    slides = list(prs.slides)

    slides_without_title = 0
    for slide in slides:
        has_title = slide.shapes.title is not None and slide.shapes.title.has_text_frame and slide.shapes.title.text_frame.text.strip()
        if not has_title:
            slides_without_title += 1

    images = get_pptx_images(prs)
    total_images, missing_alt, poor_alt, alt_feedback = assess_image_alt_quality(images)
    if total_images:
        compliant_images = total_images - missing_alt - poor_alt
        pct = round(100 * compliant_images / total_images)
        findings.append({
            'check': 'Image alternative text (OCR-verified)', 'wcag': '1.1.1 Non-text Content (Level A)',
            'category': 'Images', 'percent': pct,
            'status': 'fail' if (missing_alt or poor_alt) else 'pass',
            'detail': f'{compliant_images} of {total_images} image(s) ({pct}%) have adequate alt text; {missing_alt} missing, {poor_alt} generic or mismatched with the image\'s actual visible content (checked via OCR).',
        })
        for fb in alt_feedback[:8]:
            findings.append({'check': 'Alt text feedback', 'wcag': '1.1.1 Non-text Content (Level A)', 'status': 'warn', 'detail': fb})
    else:
        findings.append({'check': 'Image alternative text', 'wcag': '1.1.1 Non-text Content (Level A)', 'category': 'Images', 'percent': None, 'status': 'na', 'detail': 'No images found in this presentation.'})

    slides_with_title = len(slides) - slides_without_title
    slide_pct = round(100 * slides_with_title / len(slides)) if slides else None
    findings.append({
        'check': 'Slide titles', 'wcag': '2.4.6 Headings and Labels (AA) / 1.3.1 Info and Relationships (A)',
        'category': 'Slide titles', 'percent': slide_pct,
        'status': 'fail' if slides_without_title else 'pass',
        'detail': f'{slides_with_title} of {len(slides)} slide(s) ({slide_pct}%) have a title placeholder; {slides_without_title} do not, so screen readers cannot announce the topic of those slides.',
    })

    return findings


def check_accessibility_pdf(path):
    findings = [
        {
            'check': 'Tagged PDF structure', 'wcag': '1.3.1 Info and Relationships (A) / 4.1.2 Name, Role, Value (A)',
            'status': 'manual',
            'detail': "Automatic tag detection isn't available in this tool. Verify this PDF is tagged (e.g., with Acrobat's Accessibility Checker) so screen readers can interpret headings, tables, and reading order.",
        },
    ]
    image_texts = get_pdf_images(path)
    if image_texts:
        preview = image_texts[0][:150].replace('\n', ' ')
        findings.append({
            'check': 'Image alternative text', 'wcag': '1.1.1 Non-text Content (Level A)',
            'status': 'manual',
            'detail': f'{len(image_texts)} image(s) contain machine-readable text (e.g., "{preview}..."). This text has been included in the fact-check, but PDF alt-text tagging still needs manual verification (e.g., with Acrobat\'s Accessibility Checker).',
        })
    else:
        findings.append({
            'check': 'Image alternative text', 'wcag': '1.1.1 Non-text Content (Level A)',
            'status': 'manual',
            'detail': "No machine-readable text was detected in this PDF's images. Alt-text tagging still can't be reliably verified automatically — check with Acrobat's Accessibility Checker.",
        })
    return findings


def check_accessibility(path, text):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        findings = check_accessibility_docx(path)
    elif ext == '.pptx':
        findings = check_accessibility_pptx(path)
    elif ext == '.pdf':
        findings = check_accessibility_pdf(path)
    else:
        findings = [{
            'check': 'Format limitations', 'wcag': 'N/A', 'status': 'na',
            'detail': 'Plain text/Markdown files have no document-structure accessibility concerns beyond readability (see Clarity section above).',
        }]
    return findings


# Only these "primary" checks count toward the compliance score. Supplementary
# per-item detail (e.g. individual "Alt text feedback" entries, "Bold text used
# in place of headings") is excluded so one bad image doesn't get double-penalized.
SCORED_ACCESSIBILITY_CHECKS = {
    'Image alternative text', 'Image alternative text (OCR-verified)',
    'Heading styles used for structure', 'Table header rows',
    'Meaningful link text', 'Slide titles', 'Tagged PDF structure',
}


def compute_accessibility_score(findings):
    categories = []
    for f in findings:
        if f['check'] not in SCORED_ACCESSIBILITY_CHECKS:
            continue
        pct = f.get('percent')
        if pct is None:
            # Fallback for checks without a granular percent (e.g. tagged-PDF, which is
            # inherently binary/manual) — treat pass/fail as 100/0, skip na/manual entirely.
            if f['status'] == 'pass':
                pct = 100
            elif f['status'] == 'fail':
                pct = 0
            elif f['status'] == 'warn':
                pct = 50
            else:
                continue
        categories.append({'category': f.get('category', f['check']), 'percent': pct})

    if not categories:
        return {'score': None, 'zone': None, 'breakdown': []}

    score = round(sum(c['percent'] for c in categories) / len(categories))
    if score >= 90:
        zone = 'green'
    elif score >= 70:
        zone = 'yellow'
    else:
        zone = 'red'
    return {'score': score, 'zone': zone, 'breakdown': categories}


def run_checks(filepath, kb):
    text = extract_text(filepath)
    flags = []
    low = text.lower()
    norm = re.sub(r'[ \t]*\n+[ \t]*', '. ', low)
    norm = re.sub(r'\s+', ' ', norm)
    norm = re.sub(r'\.\s*\.', '.', norm)

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

    # --- Precise terminology/definition checks (Appendix A-style glossary) ---
    for m in re.finditer(r'\bcoma\b[^.]{0,100}?eyes?\s+(?:are\s+|remain(?:ing|ed)?\s+)?open', norm):
        add_flag(flags, 'Possible contradiction', 'high', 'coma ... eyes open', 'Coma is defined by no evidence of wakefulness, including eyes remaining continuously closed. Eyes opening is inconsistent with a coma diagnosis and suggests VS/UWS or a higher level of consciousness.', '7', context(norm, m.start(), m.end()))
    for m in re.finditer(r'eyes?\s+(?:are\s+|remain(?:ing|ed)?\s+)?open[^.]{0,100}?\bcoma\b', norm):
        add_flag(flags, 'Possible contradiction', 'high', 'eyes open ... coma', 'Coma is defined by no evidence of wakefulness, including eyes remaining continuously closed. Eyes opening is inconsistent with a coma diagnosis and suggests VS/UWS or a higher level of consciousness.', '7', context(norm, m.start(), m.end()))

    for m in re.finditer(r'(?:vegetative state|vs/uws|unresponsive wakefulness)[^.]{0,150}?(purposeful behavior|follow(?:s|ing)?\s+commands|command[- ]following)', norm):
        add_flag(flags, 'Possible contradiction', 'high', 'VS/UWS ... purposeful behavior/commands', 'VS/UWS is defined by NO evidence of purposeful behavior. Command following or purposeful behavior indicates at least MCS, not VS/UWS.', '7', context(norm, m.start(), m.end()))

    def sentence_window(t, start, end):
        lo = t.rfind('.', 0, start)
        lo = 0 if lo == -1 else lo + 1
        hi = t.find('.', end)
        hi = len(t) if hi == -1 else hi + 1
        return t[lo:hi]

    def sign_present(window, patterns):
        for pat in patterns:
            for pm in re.finditer(pat, window):
                preceding = window[max(0, pm.start() - 25):pm.start()]
                if re.search(r'\b(?:no|not|without|absence of|lack(?:ing)? of|negative for)\s*$', preceding):
                    continue
                return True
        return False

    mcs_plus_patterns = [r'command[- ]following', r'follow(?:s|ing)?\s+commands?', r'intelligible speech']
    mcs_minus_patterns = [r'automatic movements?', r'object manipulation', r'localizing', r'visual pursuit', r'visual fixation', r'affective behaviors?']

    for m in re.finditer(r'mcs\+', norm):
        win = sentence_window(norm, m.start(), m.end())
        has_plus_sign = sign_present(win, mcs_plus_patterns)
        has_minus_sign = sign_present(win, mcs_minus_patterns)
        if has_minus_sign and not has_plus_sign:
            add_flag(flags, 'Possible contradiction', 'medium', 'MCS+ near MCS- behaviors', 'MCS+ requires behavioral evidence of preserved receptive language (e.g., command following, intelligible speech). The nearby behaviors described (e.g., automatic movements, object manipulation, visual pursuit) define MCS-, not MCS+.', '7', context(norm, m.start(), m.end()))

    for m in re.finditer(r'mcs-(?!\w)', norm):
        win = sentence_window(norm, m.start(), m.end())
        has_minus_sign = sign_present(win, mcs_minus_patterns)
        has_plus_sign = sign_present(win, mcs_plus_patterns)
        if has_plus_sign and not has_minus_sign:
            add_flag(flags, 'Possible contradiction', 'medium', 'MCS- near MCS+ behaviors', 'MCS- is defined by nonlinguistic signs only (automatic movements, object manipulation, visual pursuit/fixation, affective behaviors). Command following or intelligible speech nearby indicates MCS+, not MCS-.', '7', context(norm, m.start(), m.end()))

    for m in re.finditer(r'persistent vegetative state[^.]{0,80}?(irreversible|permanent)', norm):
        add_flag(flags, 'Terminology', 'medium', 'persistent vegetative state ... permanent/irreversible', '"Persistent vegetative state" (PVS) denotes VS/UWS lasting more than 1 month and does not itself imply irreversibility. "Permanent vegetative state" is a distinct prognostic term applied at 3 months (nontraumatic) or 12 months (traumatic) indicating high probability of irreversibility.', '7', context(norm, m.start(), m.end()))

    for m in re.finditer(r'locked-in syndrome[^.]{0,150}?vegetative state', norm):
        add_flag(flags, 'Terminology', 'high', 'locked-in syndrome ... vegetative state', 'Locked-in syndrome (tetraplegia, anarthria, near-normal cognition) is a distinct condition that can be misdiagnosed as VS/UWS. It is not itself a disorder of consciousness and should not be equated with vegetative state.', None, context(norm, m.start(), m.end()))
    for m in re.finditer(r'vegetative state[^.]{0,150}?locked-in syndrome', norm):
        add_flag(flags, 'Terminology', 'high', 'vegetative state ... locked-in syndrome', 'Locked-in syndrome (tetraplegia, anarthria, near-normal cognition) is a distinct condition that can be misdiagnosed as VS/UWS. It is not itself a disorder of consciousness and should not be equated with vegetative state.', None, context(norm, m.start(), m.end()))

    for m in re.finditer(r'emerg(?:ed|ence)\s+from\s+mcs[^.]{0,150}?(?:used|using|use of)\s+(?:a|one|1)\s+(?:single\s+)?(?:familiar\s+)?object', norm):
        add_flag(flags, 'Key fact mismatch', 'medium', 'emergence from MCS ... one object', 'Emergence from MCS (EMCS) via functional object use requires demonstrated use of at least 2 different familiar objects, not a single object.', '7', context(norm, m.start(), m.end()))

    for m in re.finditer(r'(\d{1,2})\s*hours?\s+of\s+(?:multidisciplinary\s+)?therapy\s+(?:daily|per\s+day)', norm):
        hours = int(m.group(1))
        if hours != 3:
            add_flag(flags, 'Key fact mismatch', 'high', f'{hours} hours of therapy daily',
                      f'Standard IRF admission criteria require patients be able to tolerate approximately 3 hours of multidisciplinary therapy daily (the "3-hour rule"), not {hours} hours.',
                      '1', context(norm, m.start(), m.end()))

    for m in re.finditer(r'physicians?\s*\([^)]*intensivist[^)]*\)\.?\s*(?:the\s+)?overall leaders?', norm):
        add_flag(flags, 'Key fact mismatch', 'medium', 'Physicians ... overall leaders',
                  'The physician role (Intensivist, Neurologist, Physiatrist) on a multidisciplinary DoC team is typically described as team oversight, medical management, and disposition planning — not simply "the overall leaders."',
                  '1', context(norm, m.start(), m.end()))

    # --- Negated / inverted recommendation detector ---
    # Catches statements that flip the polarity of a guideline recommendation
    # (e.g., "referral ... is not critical" when the guideline recommends it).
    negation_phrase = r'not\s+(?:critical|necessary|essential|important|required|needed|recommended|beneficial|effective)|unnecessary|no\s+(?:need|benefit|role)|does\s+not\s+(?:improve|help|benefit)|isn.t\s+(?:critical|necessary|essential|important|required|needed)'
    negated_rec_topics = [
        {'keywords': [r'referral to a specialized rehabilitation', r'referral to (?:a |an )?multidisciplinary', r'multidisciplinary rehabilitation'], 'rec': '1',
         'note': 'Guideline Recommendation 1 states clinicians SHOULD refer medically stable patients with DoC to specialized multidisciplinary rehabilitation settings (Level B) to optimize diagnosis, prognostication, and management.'},
        {'keywords': [r'\bamantadine\b'], 'rec': '14',
         'note': 'Guideline Recommendation 14 supports amantadine (100-200 mg twice daily) to hasten functional recovery in appropriate traumatic VS/UWS or MCS patients (Level B).'},
        {'keywords': [r'goals of care'], 'rec': '9',
         'note': 'Guideline Recommendation 9 states clinicians MUST counsel families to establish goals of care once prognosis indicates likely severe long-term disability (Level A).'},
        {'keywords': [r'patient and family preferences', r'family preferences'], 'rec': '11',
         'note': 'Guideline Recommendation 11 requires incorporating patient and family preferences into care decisions (Level A).'},
        {'keywords': [r'pain assessment', r'pain management'], 'rec': '13',
         'note': 'The guideline supports routine pain assessment and management in patients with DoC (Level B).'},
        {'keywords': [r'serial standardized behavioral evaluations', r'serial behavioral evaluations'], 'rec': '4',
         'note': 'Guideline Recommendation 4 states clinicians SHOULD perform serial standardized behavioral evaluations to establish prognosis (Level B).'},
    ]
    for topic in negated_rec_topics:
        for kw in topic['keywords']:
            for m in re.finditer(kw + r'[^.]{0,150}?(?:' + negation_phrase + r')', norm):
                add_flag(flags, 'Possible inverted recommendation', 'high', m.group(0)[:120], f"This appears to invert the guideline's actual recommendation. {topic['note']}", topic['rec'], context(norm, m.start(), m.end()))
            for m in re.finditer(r'(?:' + negation_phrase + r')[^.]{0,150}?' + kw, norm):
                add_flag(flags, 'Possible inverted recommendation', 'high', m.group(0)[:120], f"This appears to invert the guideline's actual recommendation. {topic['note']}", topic['rec'], context(norm, m.start(), m.end()))

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
    seen = set()
    deduped_flags = []
    for f in flags:
        key = (f['kind'], f['severity'], f['matched'], f['issue'])
        if key in seen:
            continue
        seen.add(key)
        deduped_flags.append(f)
    flags = deduped_flags
    flags.sort(key=lambda x: order.get(x['severity'], 9))
    accessibility_findings = check_accessibility(filepath, text)
    return {
        'filename': os.path.basename(filepath),
        'text_length': len(text),
        'flags': flags,
        'coverage': sorted(set(coverage)),
        'clarity': analyze_clarity(text),
        'accessibility': accessibility_findings,
        'accessibility_score': compute_accessibility_score(accessibility_findings),
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
