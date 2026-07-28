import json
import os
import tempfile
import uuid
from datetime import datetime

import streamlit as st

from factcheck_engine import SUPPORTED_EXTS, load_kb, run_checks

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KB_PATH = os.path.join(BASE_DIR, 'doc_guidelines_kb.json')

HISTORY_DIR = os.path.join(BASE_DIR, 'data', 'history')
INDEX_PATH = os.path.join(HISTORY_DIR, 'index.json')

os.makedirs(HISTORY_DIR, exist_ok=True)

APP_NAME = 'DoC Kit Check'

st.set_page_config(page_title=APP_NAME, layout='wide')

MGB_DEEP_BLUE = '#003A96'
MGB_TEAL = '#009CA6'
MGB_TEXT = '#202020'
MGB_MUTED = '#5F6B7A'
MGB_BORDER = 'rgba(0, 58, 150, 0.12)'

# Status/severity colors used by the badge() helper below — kept in one place so the
# table styling (render_wrapped_table) and the badges always agree with each other.
BADGE_COLORS = {
    'pass':   ('#E3F4E1', '#1E7A34'),
    'fail':   ('#FBE4E4', '#B3261E'),
    'warn':   ('#FFF4DE', '#8A5A00'),
    'na':     ('#EEF0F2', '#5F6B7A'),
    'manual': ('#E3EEFB', '#0B5FA5'),
    'high':   ('#FBE4E4', '#B3261E'),
    'medium': ('#FFF4DE', '#8A5A00'),
    'low':    ('#EEF0F2', '#5F6B7A'),
    'covered': ('#E3F4E1', '#1E7A34'),
}

st.markdown(f"""
<style>
    h1 {{ color: {MGB_DEEP_BLUE} !important; font-weight: 700; }}
    h2 {{ color: {MGB_DEEP_BLUE} !important; border-bottom: 2px solid {MGB_TEAL}; padding-bottom: 0.4rem; margin-top: 0.5rem; }}
    h3, h4 {{ color: {MGB_TEXT} !important; }}
    p, li, span, label {{ color: {MGB_TEXT}; }}
    [data-testid="stMetricValue"] {{ color: {MGB_DEEP_BLUE} !important; }}
    [data-testid="stMetricLabel"] {{ color: {MGB_MUTED} !important; font-weight: 600; text-transform: uppercase; font-size: 0.75rem; letter-spacing: 0.03em; }}
    .stProgress > div > div > div {{ background-color: {MGB_TEAL} !important; }}
    hr {{ border-top: 1px solid {MGB_TEAL}; opacity: 0.4; margin: 1.75rem 0; }}
    [data-testid="stTabs"] button[aria-selected="true"] {{ color: {MGB_DEEP_BLUE} !important; border-bottom-color: {MGB_DEEP_BLUE} !important; }}
    [data-testid="stTabs"] button p {{ font-weight: 600; }}

    .app-header {{ padding-bottom: 0.25rem; }}
    .app-tagline {{ color: {MGB_MUTED}; font-size: 1rem; margin-top: -0.75rem; }}

    .section-eyebrow {{
        color: {MGB_TEAL}; font-size: 0.72rem; font-weight: 700;
        letter-spacing: 0.09em; text-transform: uppercase; margin-bottom: -0.6rem;
    }}

    .badge {{
        display: inline-block; padding: 2px 10px; border-radius: 999px;
        font-size: 0.75rem; font-weight: 600; white-space: nowrap;
    }}

    div[data-testid="stExpander"] {{
        border: 1px solid {MGB_BORDER} !important; border-radius: 8px !important;
    }}

    /* Multiselect tags (e.g. "Include these materials in the cumulative view") render
       with a dark blue background — force the label and the x/close icon to white so
       they're actually readable against it. */
    span[data-baseweb="tag"] {{ color: #FFFFFF !important; }}
    span[data-baseweb="tag"] span {{ color: #FFFFFF !important; }}
    span[data-baseweb="tag"] svg {{ fill: #FFFFFF !important; }}

    /* Primary button ("Run fact-check") — force white label text/icon regardless of
       theme background color. */
    button[kind="primary"] {{ color: #FFFFFF !important; }}
    button[kind="primary"] p {{ color: #FFFFFF !important; }}
    button[kind="primary"] div {{ color: #FFFFFF !important; }}
</style>
""", unsafe_allow_html=True)


def badge(text, kind):
    """Render a small colored status pill instead of an emoji/icon — used for pass/fail/
    warn/manual accessibility statuses, high/medium/low severity, and covered/not-covered
    markers. Must be used inside render_wrapped_table's raw_columns, since it returns HTML."""
    bg, fg = BADGE_COLORS.get(kind, ('#EEF0F2', MGB_TEXT))
    label = html_module.escape(str(text))
    return f'<span class="badge" style="background:{bg}; color:{fg};">{label}</span>'


import html as html_module


def render_wrapped_table(rows, columns=None, raw_columns=None):
    """Render a list-of-dicts as a static HTML table with wrapped text (no truncation/click-to-expand).
    Columns listed in raw_columns are inserted as-is (use for badge() HTML); every other
    column is escaped normally."""
    if not rows:
        st.write('None.')
        return
    if columns is None:
        columns = list(rows[0].keys())
    raw_columns = raw_columns or set()

    def esc(v):
        return html_module.escape(str(v)).replace('\n', '<br>')

    parts = ['<table style="width:100%; border-collapse:collapse; font-size:0.9rem;">']
    parts.append('<thead><tr>')
    for c in columns:
        parts.append(f'<th style="text-align:left; padding:8px 10px; border-bottom:2px solid {MGB_TEAL}; color:{MGB_DEEP_BLUE};">{esc(c)}</th>')
    parts.append('</tr></thead><tbody>')
    for row in rows:
        parts.append('<tr>')
        for c in columns:
            val = row.get(c, '')
            cell = val if c in raw_columns else esc(val)
            parts.append(f'<td style="padding:8px 10px; border-bottom:1px solid rgba(0,58,150,0.08); white-space:normal; word-wrap:break-word; vertical-align:top;">{cell}</td>')
        parts.append('</tr>')
    parts.append('</tbody></table>')
    st.markdown(''.join(parts), unsafe_allow_html=True)


@st.cache_data
def get_kb():
    return load_kb(KB_PATH)


# ---------------------------------------------------------------------------
# History storage helpers
# ---------------------------------------------------------------------------
def load_history():
    if not os.path.exists(INDEX_PATH):
        return []
    with open(INDEX_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_history(entries):
    with open(INDEX_PATH, 'w', encoding='utf-8') as f:
        json.dump(entries, f, indent=2)


def add_history_entry(filename, result, counts, nickname=None):
    entries = load_history()
    entry = {
        'id': uuid.uuid4().hex[:10],
        'filename': filename,
        'nickname': (nickname or '').strip(),
        'checked_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'counts': counts,
        'flags': result['flags'],
        'coverage': result['coverage'],
        'clarity': result['clarity'],
        'accessibility': result['accessibility'],
        'accessibility_score': result.get('accessibility_score', default_accessibility_score()),
        'proofreading': result.get('proofreading', default_proofreading()),
        'ocr_warning': result.get('ocr_warning'),
    }
    entries.insert(0, entry)
    save_history(entries)
    return entry


def delete_history_entry(entry_id):
    entries = load_history()
    entries = [e for e in entries if e['id'] != entry_id]
    save_history(entries)


def default_clarity():
    return {
        'flesch_reading_ease': None,
        'flesch_kincaid_grade': None,
        'avg_sentence_length': None,
        'long_sentences': [],
        'undefined_acronyms': [],
        'passive_voice_count': 0,
        'suggestions': ['Clarity analysis was not available when this material was originally checked. Re-run the check to generate it.'],
    }


def default_accessibility():
    return [{
        'check': 'Not available', 'wcag': 'N/A', 'status': 'na',
        'detail': 'Section 508 accessibility analysis was not available when this material was originally checked. Re-run the check to generate it.',
    }]


def default_accessibility_score():
    return {'score': None, 'zone': None, 'breakdown': []}


def default_proofreading():
    return {'spelling': [], 'grammar': [], 'plagiarism': [], 'internal_duplication': [], 'missing_references': [], 'conciseness': [], '_unavailable': True}


def counts_from_flags(flags):
    counts = {s: 0 for s in ['high', 'medium', 'low']}
    for flag in flags:
        counts[flag['severity']] = counts.get(flag['severity'], 0) + 1
    return counts


# ---------------------------------------------------------------------------
# Shared rendering for a result (used for both fresh checks and history)
# ---------------------------------------------------------------------------
def render_result(filename, counts, flags, coverage, clarity, accessibility, accessibility_score, proofreading, kb, nickname=None, checked_at=None, key_suffix=None, ocr_warning=None):
    display_name = nickname.strip() if nickname and nickname.strip() else filename
    caption = display_name if not checked_at else f'{display_name} — checked {checked_at}'
    st.subheader(caption)
    if nickname and nickname.strip():
        st.caption(f"Original file: {filename}")

    if ocr_warning:
        st.error(f"OCR issue: {ocr_warning}")

    st.caption(f"Guideline: {kb['meta']['title']} ({kb['meta']['year']}) — {kb['meta']['citation']}")

    rec_by_id = {rec['id']: rec for rec in kb['recommendations']}

    # =======================================================================
    # SECTION 1: Guideline fact-check
    # =======================================================================
    st.markdown('<div class="section-eyebrow">Section 1 of 4</div>', unsafe_allow_html=True)
    st.header('Guideline fact-check')
    st.caption('Checks the content of your document against the DoC clinical guideline\'s recommendations and key facts.')

    total_errors = len(flags)
    st.metric('Errors flagged', total_errors)

    st.markdown('#### Flags')
    if not flags:
        st.info('No flags raised. Human review is still recommended.')
    else:
        for i, flag in enumerate(flags):
            with st.expander(f"[{flag['severity'].upper()}] {flag['kind']} — {flag['matched']}", key=f"flag_{key_suffix}_{i}"):
                st.write(f"**Issue:** {flag['issue']}")
                if flag['rec']:
                    rec = rec_by_id.get(flag['rec'])
                    if rec:
                        level = '/'.join(rec['level'])
                        st.write(f"**Related recommendation:** Recommendation {rec['id']} — {rec['topic']} (Level {level})")
                        st.write(f"**Guideline text:** {rec['text']}")
                        pages = []
                        if rec.get('summary_page'):
                            pages.append(f"p. {rec['summary_page']} (recommendations summary)")
                        if rec.get('detail_page'):
                            pages.append(f"p. {rec['detail_page']} (detailed rationale)")
                        page_str = '; '.join(pages)
                        st.caption(f"Source: {kb['meta']['citation']}" + (f" — {page_str}" if page_str else ""))
                    else:
                        st.write(f"**Related recommendation:** {flag['rec']}")
                if flag['context']:
                    st.caption(f"Context: {flag['context']}")
                ai_review = flag.get('ai_review')
                if ai_review:
                    verdict = ai_review.get('verdict', '')
                    if verdict == 'unavailable':
                        st.caption(f"AI review: unavailable — {ai_review.get('reason', '')}")
                    else:
                        verdict_kind = 'covered' if 'false positive' in verdict else ('fail' if 'real issue' in verdict else 'warn')
                        st.markdown(f"AI review: {badge(verdict, verdict_kind)}", unsafe_allow_html=True)
                        if ai_review.get('reason'):
                            st.caption(ai_review['reason'])

    st.markdown('#### Recommendation coverage map')
    st.caption('Which guideline recommendations your document actually touches on.')
    coverage_rows = [
        {
            'Rec': rec['id'],
            'Topic': rec['topic'],
            'Level': '/'.join(rec['level']),
            'Touched?': badge('Covered', 'covered') if rec['id'] in coverage else '—',
        }
        for rec in kb['recommendations']
    ]
    render_wrapped_table(coverage_rows, raw_columns={'Touched?'})

    st.divider()

    # =======================================================================
    # SECTION 2: Section 508 / accessibility check
    # =======================================================================
    st.markdown('<div class="section-eyebrow">Section 2 of 4</div>', unsafe_allow_html=True)
    st.header('Section 508 / accessibility check')
    st.caption('Checks the document\'s structure (images, headings, tables, links) against Section 508 / WCAG requirements — separate from the fact-check above.')

    score = accessibility_score.get('score') if accessibility_score else None
    breakdown = accessibility_score.get('breakdown', []) if accessibility_score else []
    if score is not None:
        sc1, sc2 = st.columns([1, 3])
        sc1.metric('Overall Section 508 compliance', f'{score}%')
        with sc2:
            st.progress(score / 100)

        if breakdown:
            st.caption('Breakdown by category — the percentage of that material (images, tables, etc.) that meets the relevant WCAG criterion:')
            cat_cols = st.columns(len(breakdown))
            for col, cat in zip(cat_cols, breakdown):
                with col:
                    st.metric(cat['category'], f"{cat['percent']}%")
                    st.progress(cat['percent'] / 100)

        st.caption('Scores reflect only the automatically-checkable items below — not the reference checklist further down, which requires human judgment.')
    else:
        st.caption('No automatically-scoreable accessibility checks apply to this file type.')

    status_labels = {'pass': 'Pass', 'fail': 'Fail', 'warn': 'Warning', 'na': 'N/A', 'manual': 'Manual check'}
    a11y_rows = [
        {
            'Status': badge(status_labels.get(f['status'], f['status']), f['status']),
            'Check': f['check'],
            'WCAG criterion': f['wcag'],
            'Finding': f['detail'],
        }
        for f in accessibility
    ]
    render_wrapped_table(a11y_rows, raw_columns={'Status'})
    if any(f['status'] == 'manual' for f in accessibility):
        st.caption('"Manual check" = requires verification with a dedicated accessibility checker (e.g., Acrobat).')

    fixable = [f for f in accessibility if f.get('fix') and f['status'] in ('fail', 'warn', 'manual')]
    if fixable:
        st.markdown('#### How can I fix this?')
        for fi, f in enumerate(fixable):
            with st.expander(f"Fix: {f['check']}" + (f" — {f['detail'][:60]}..." if len(f['detail']) > 60 else ''), key=f"fix_{key_suffix}_{fi}"):
                st.write(f['fix'])

    with st.expander('General accessibility checklist (WCAG POUR principles)', key=f"checklist_{key_suffix}"):
        st.caption(
            "Covers what software can't check automatically — needs a human judgment call. "
            "Organized by WCAG's four principles (Perceivable, Operable, Understandable, Robust)."
        )
        checklist_rows = [
            {'Principle': c['principle'], 'Guidance': c['item'], 'WCAG': c['wcag']}
            for c in kb.get('accessibility_checklist', [])
        ]
        render_wrapped_table(checklist_rows)

    st.divider()

    # =======================================================================
    # SECTION 3: Clarity & readability
    # =======================================================================
    st.markdown('<div class="section-eyebrow">Section 3 of 4</div>', unsafe_allow_html=True)
    st.header('Clarity & understandability')
    st.caption('How readable your document is — separate from both the fact-check and the 508 check above.')

    c1, c2, c3 = st.columns(3)
    c1.metric('Reading grade level', clarity['flesch_kincaid_grade'] if clarity['flesch_kincaid_grade'] is not None else 'N/A')
    c2.metric('Reading ease', clarity['flesch_reading_ease'] if clarity['flesch_reading_ease'] is not None else 'N/A')
    c3.metric('Avg. sentence length', f"{clarity['avg_sentence_length']} words" if clarity['avg_sentence_length'] is not None else 'N/A')

    for suggestion in clarity['suggestions']:
        if 'No major clarity issues' in suggestion:
            st.success(suggestion)
        else:
            st.warning(suggestion)

    if clarity['long_sentences']:
        with st.expander(f"Long sentences ({len(clarity['long_sentences'])})", key=f"longsent_{key_suffix}"):
            for s in clarity['long_sentences']:
                st.write(f"**{s['word_count']} words:** {s['text']}")

    if clarity['undefined_acronyms']:
        with st.expander(f"Undefined acronyms ({len(clarity['undefined_acronyms'])})", key=f"acronyms_{key_suffix}"):
            for a in clarity['undefined_acronyms']:
                st.write(f"**{a['acronym']}** — {a['expansion']}")

    with st.expander('Appendix: guideline recommendation wording', key=f"appendix_{key_suffix}"):
        st.caption(
            "What this is: a condensed summary of every recommendation in the guideline listed above, "
            "checked against the source PDF for factual accuracy (numbers, timeframes, and evidence levels all match). "
            "The wording itself is paraphrased for brevity, not a verbatim quotation — see the full guideline "
            "for the guideline's own exact phrasing of each recommendation statement."
        )
        for rec in kb['recommendations']:
            st.markdown(f"**Recommendation {rec['id']} (Level {'/'.join(rec['level'])})**")
            st.write(rec['text'])
            st.divider()

    st.divider()

    # =======================================================================
    # SECTION 4: Proofreading
    # =======================================================================
    st.markdown('<div class="section-eyebrow">Section 4 of 4</div>', unsafe_allow_html=True)
    st.header('Proofreading')
    st.caption('Spelling, grammar, possible plagiarism, missing citations, and wordiness — separate from the guideline fact-check above.')

    if proofreading.get('_unavailable'):
        st.info('Proofreading analysis was not available when this material was originally checked. Re-run the check to generate it.')
    else:
        pf_col1, pf_col2, pf_col3, pf_col4, pf_col5 = st.columns(5)
        pf_col1.metric('Spelling', len(proofreading['spelling']))
        pf_col2.metric('Grammar', len(proofreading['grammar']))
        pf_col3.metric('Possible plagiarism', len(proofreading['plagiarism']) + len(proofreading['internal_duplication']))
        pf_col4.metric('Missing citations', len(proofreading['missing_references']))
        pf_col5.metric('Wordy phrases', sum(c['count'] for c in proofreading['conciseness']))

        st.markdown('#### Spelling')
        if proofreading['spelling']:
            spelling_rows = [{'Word': s['word'], 'Occurrences': s['count'], 'Suggested correction': s['suggestion'] or '—'} for s in proofreading['spelling']]
            render_wrapped_table(spelling_rows)
        else:
            st.success('No likely misspellings detected.')

        st.markdown('#### Grammar (basic checks)')
        st.caption('Catches mechanical issues (repeated words, spacing, capitalization, unmatched punctuation) — not a full grammar model.')
        if proofreading['grammar']:
            grammar_rows = [{'Issue': g['issue'], 'Context': g['context']} for g in proofreading['grammar']]
            render_wrapped_table(grammar_rows)
        else:
            st.success('No basic grammar issues detected.')

        st.markdown('#### Possible plagiarism')
        st.caption(
            'Checks for text copied near-verbatim from the guideline itself without quotation marks, and for duplicated '
            'passages within your own document. This does NOT search the internet or other external sources — for that, '
            'use a dedicated plagiarism-detection service.'
        )
        if proofreading['plagiarism']:
            st.write('**Unattributed guideline text:**')
            plag_rows = [
                {'Matches source': p['matched_source'], 'Sentence in your document': p['sentence'], 'Has quote marks?': 'Yes' if p['quoted'] else 'No'}
                for p in proofreading['plagiarism']
            ]
            render_wrapped_table(plag_rows)
        if proofreading['internal_duplication']:
            st.write('**Duplicated within your own document:**')
            for d in proofreading['internal_duplication']:
                st.write(f"- {d}")
        if not proofreading['plagiarism'] and not proofreading['internal_duplication']:
            st.success('No unattributed guideline text or internal duplication detected.')

        st.markdown('#### Missing citations')
        st.caption('Sentences that make a claim or cite a statistic but have no nearby citation marker (e.g., "(Smith, 2020)", "[1]", a URL).')
        if proofreading['missing_references']:
            for r in proofreading['missing_references']:
                st.warning(r)
        else:
            st.success('No unsupported claims/statistics detected.')

        st.markdown('#### Clarity & conciseness suggestions')
        if proofreading['conciseness']:
            conc_rows = [{'Wordy phrase': c['phrase'], 'Occurrences': c['count'], 'Suggested alternative': c['suggestion']} for c in proofreading['conciseness']]
            render_wrapped_table(conc_rows)
        else:
            st.success('No common wordy phrases detected.')
        st.caption('See the Clarity & understandability section above for reading-level, sentence-length, and passive-voice suggestions.')


# ---------------------------------------------------------------------------
# Layout: Check a document | Previously reviewed
# ---------------------------------------------------------------------------
st.markdown(f'<div class="app-header"><h1>{APP_NAME}</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="app-tagline">Guideline compliance, accessibility, clarity, and proofreading checks for DoC materials.</div>', unsafe_allow_html=True)
st.write('')

kb = get_kb()

tab_check, tab_history, tab_coverage = st.tabs(['Check a document', 'Previously reviewed', 'Cumulative coverage'])

with tab_check:
    st.write(
        'Upload a document (PDF, DOCX, PPTX, PNG, JPEG, TXT, or MD) and this tool will flag '
        'places where the content may conflict with the DoC clinical guideline.'
    )

    uploaded_file = st.file_uploader(
        'Choose a file to check',
        type=[ext.lstrip('.') for ext in SUPPORTED_EXTS],
    )
    nickname_input = st.text_input('Nickname for this document (optional)', placeholder='e.g., "Family handout draft 2"')

    try:
        gemini_api_key = st.secrets.get('GEMINI_API_KEY', os.environ.get('GEMINI_API_KEY', ''))
    except Exception:
        gemini_api_key = os.environ.get('GEMINI_API_KEY', '')
    ai_review_enabled = st.checkbox(
        'Double-check with AI (Gemini)',
        value=False,
        help='Sends every flag\'s sentence and surrounding context (not the whole document) to '
             'Google\'s Gemini API to sanity-check it. Do not enable this on '
             'documents containing PHI or other confidential/regulated content — see the note below.',
    )
    if ai_review_enabled:
        if not gemini_api_key:
            st.warning(
                'No Gemini API key configured. Add GEMINI_API_KEY to your Streamlit secrets or environment '
                'to use this — get a free key at aistudio.google.com. AI double-checking will be skipped '
                'for this run.'
            )
        st.caption(
            'On Google\'s free API tier, submitted text may be used to improve their models and is not '
            'HIPAA/BAA-eligible. Only use this on non-PHI, non-confidential material — same policy as the rest of this tool.'
        )
        st.caption(
            'Every flag gets checked — Terminology, Worth double-checking, Key fact mismatch, and quiz-answer flags '
            'alike — paced to stay within the free tier\'s rate limit (~9/minute) rather than capped at a fixed '
            'number, so documents with many flags will take longer to finish.'
        )

    if uploaded_file is not None and st.button('Run fact-check', type='primary'):
        spinner_text = 'Checking document against guideline...'
        if ai_review_enabled and gemini_api_key:
            spinner_text += ' AI double-checking is on, so this may take a while for documents with many flags.'
        with st.spinner(spinner_text):
            temp_dir = tempfile.mkdtemp(prefix='doc_factcheck_')
            in_path = os.path.join(temp_dir, uploaded_file.name)
            with open(in_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            try:
                result = run_checks(in_path, kb, verify_with_llm=ai_review_enabled, llm_api_key=gemini_api_key)
                counts = counts_from_flags(result['flags'])
                entry = add_history_entry(uploaded_file.name, result, counts, nickname=nickname_input)

                st.success('Fact-check complete — saved to "Previously reviewed."')
                render_result(
                    entry['filename'], entry['counts'], entry['flags'], entry['coverage'],
                    entry['clarity'], entry['accessibility'], entry['accessibility_score'], entry['proofreading'], kb,
                    nickname=entry.get('nickname'), checked_at=entry['checked_at'], key_suffix=f"check_{entry['id']}",
                    ocr_warning=entry.get('ocr_warning'),
                )

            except Exception as e:
                st.error(f'Could not process this file: {e}')

def label_for(e):
    name = e.get('nickname', '').strip() or e['filename']
    return f"{name} — {e['checked_at']}"


with tab_history:
    history = load_history()

    if not history:
        st.info('No materials reviewed yet. Check a document in the first tab to get started.')
    else:
        st.write(f'{len(history)} material(s) reviewed so far.')

        options = [label_for(e) for e in history]
        selected_label = st.selectbox('Select a previously reviewed material', options)
        selected_entry = history[options.index(selected_label)]

        del_col, _ = st.columns([1, 4])
        with del_col:
            confirm_delete = st.checkbox('Confirm delete', key=f"confirm_delete_{selected_entry['id']}")
            if st.button('Delete this entry', disabled=not confirm_delete, key=f"delete_{selected_entry['id']}"):
                delete_history_entry(selected_entry['id'])
                st.success('Deleted.')
                st.rerun()

        render_result(
            selected_entry['filename'], selected_entry['counts'], selected_entry['flags'],
            selected_entry['coverage'], selected_entry.get('clarity', default_clarity()),
            selected_entry.get('accessibility', default_accessibility()),
            selected_entry.get('accessibility_score', default_accessibility_score()),
            selected_entry.get('proofreading', default_proofreading()), kb,
            nickname=selected_entry.get('nickname'), checked_at=selected_entry['checked_at'], key_suffix=f"history_{selected_entry['id']}",
            ocr_warning=selected_entry.get('ocr_warning'),
        )

import io
import pandas as pd


def build_excel_bytes(rows, sheet_name='Cumulative Coverage'):
    df = pd.DataFrame(rows)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        for col_idx, col in enumerate(df.columns, start=1):
            max_len = max(df[col].astype(str).map(len).max() if len(df) else 0, len(col))
            worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 60)
    return buffer.getvalue()


with tab_coverage:
    st.write('A running record of which guideline recommendations have ever been addressed, across every material checked so far, and by which document(s).')

    history = load_history()

    if not history:
        st.info('No materials reviewed yet. Check a document in the first tab to get started.')
    else:
        all_labels = [label_for(e) for e in history]
        selected_labels = st.multiselect(
            'Include these materials in the cumulative view',
            options=all_labels,
            default=all_labels,
        )
        included = [e for e, label in zip(history, all_labels) if label in selected_labels]

        if not included:
            st.warning('No materials selected — pick at least one above to see coverage.')
        else:
            st.caption(f'Showing cumulative coverage across {len(included)} of {len(history)} reviewed material(s).')

            # Build: rec_id -> list of {name, checked_at}
            rec_hits = {rec['id']: [] for rec in kb['recommendations']}
            for e in included:
                name = e.get('nickname', '').strip() or e['filename']
                for rid in e.get('coverage', []):
                    if rid in rec_hits:
                        rec_hits[rid].append({'name': name, 'checked_at': e['checked_at']})

            covered_count = sum(1 for rid, hits in rec_hits.items() if hits)
            total_count = len(kb['recommendations'])
            st.metric('Recommendations covered', f'{covered_count} of {total_count}')
            st.progress(covered_count / total_count if total_count else 0)

            rows = []
            for rec in kb['recommendations']:
                hits = rec_hits[rec['id']]
                if hits:
                    docs_str = '; '.join(f"{h['name']} ({h['checked_at']})" for h in hits)
                else:
                    docs_str = '—'
                rows.append({
                    'Rec': rec['id'],
                    'Topic': rec['topic'],
                    'Level': '/'.join(rec['level']),
                    'Covered?': 'Yes' if hits else '',
                    'References': len(hits),
                    'Covered by (material — date)': docs_str,
                })
            display_rows = [dict(r, **{'Covered?': badge('Covered', 'covered') if r['Covered?'] else '—'}) for r in rows]
            render_wrapped_table(display_rows, raw_columns={'Covered?'})

            excel_bytes = build_excel_bytes(rows)
            st.download_button(
                label='Download cumulative coverage (Excel)',
                data=excel_bytes,
                file_name=f'cumulative_coverage_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
