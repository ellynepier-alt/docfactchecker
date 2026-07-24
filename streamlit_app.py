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

st.set_page_config(page_title='DoC Guideline Fact-Checker', page_icon='📋', layout='wide')


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
        'accessibility_score': result.get('accessibility_score', {'score': None, 'zone': None}),
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


def counts_from_flags(flags):
    counts = {s: 0 for s in ['high', 'medium', 'low', 'review']}
    for flag in flags:
        counts[flag['severity']] = counts.get(flag['severity'], 0) + 1
    return counts


# ---------------------------------------------------------------------------
# Shared rendering for a result (used for both fresh checks and history)
# ---------------------------------------------------------------------------
def render_result(filename, counts, flags, coverage, clarity, accessibility, accessibility_score, kb, nickname=None, checked_at=None, key_suffix=None):
    display_name = nickname.strip() if nickname and nickname.strip() else filename
    caption = display_name if not checked_at else f'{display_name} — checked {checked_at}'
    st.subheader(caption)
    if nickname and nickname.strip():
        st.caption(f"Original file: {filename}")

    st.caption(f"Guideline: {kb['meta']['title']} ({kb['meta']['year']}) — {kb['meta']['citation']}")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric('High', counts['high'])
    col2.metric('Medium', counts['medium'])
    col3.metric('Low', counts['low'])
    col4.metric('Review', counts['review'])

    st.markdown('#### Flags')
    if not flags:
        st.info('No flags raised. Human review is still recommended.')
    else:
        for i, flag in enumerate(flags):
            with st.expander(f"[{flag['severity'].upper()}] {flag['kind']} — {flag['matched']}"):
                st.write(f"**Issue:** {flag['issue']}")
                if flag['rec']:
                    st.write(f"**Related recommendation:** {flag['rec']}")
                if flag['context']:
                    st.caption(f"Context: {flag['context']}")

    st.markdown('#### Recommendation coverage map')
    coverage_rows = [
        {
            'Rec': rec['id'],
            'Topic': rec['topic'],
            'Level': '/'.join(rec['level']),
            'Touched?': '✅' if rec['id'] in coverage else '',
        }
        for rec in kb['recommendations']
    ]
    st.dataframe(coverage_rows, use_container_width=True, hide_index=True)

    st.markdown('#### Section 508 / accessibility check')

    score = accessibility_score.get('score') if accessibility_score else None
    zone = accessibility_score.get('zone') if accessibility_score else None
    breakdown = accessibility_score.get('breakdown', []) if accessibility_score else []
    if score is not None:
        zone_display = {'green': '🟢 Green — looks solid', 'yellow': '🟡 Yellow — needs some attention', 'red': '🔴 Red — significant issues'}
        sc1, sc2 = st.columns([1, 3])
        sc1.metric('Overall Section 508 compliance', f'{score}%')
        with sc2:
            st.write(zone_display.get(zone, ''))
            st.progress(score / 100)

        if breakdown:
            st.caption('Breakdown by category — the percentage of that material (images, tables, etc.) that meets the relevant WCAG criterion:')
            cat_cols = st.columns(len(breakdown))
            for col, cat in zip(cat_cols, breakdown):
                with col:
                    st.metric(cat['category'], f"{cat['percent']}%")
                    st.progress(cat['percent'] / 100)

        st.caption('Scores reflect only the automatically-checkable items below (alt text, headings, table headers, link text, slide titles, or tagged-PDF status) — not the reference checklist further down, which requires human judgment.')
    else:
        st.caption('No automatically-scoreable accessibility checks apply to this file type.')

    status_icons = {'pass': '✅', 'fail': '❌', 'warn': '⚠️', 'na': 'ℹ️', 'manual': '🔍'}
    a11y_rows = [
        {
            '': status_icons.get(f['status'], ''),
            'Check': f['check'],
            'WCAG criterion': f['wcag'],
            'Finding': f['detail'],
        }
        for f in accessibility
    ]
    st.dataframe(a11y_rows, use_container_width=True, hide_index=True)
    if any(f['status'] == 'manual' for f in accessibility):
        st.caption('🔍 = requires manual verification with a dedicated accessibility checker (e.g., Acrobat).')

    with st.expander('General accessibility checklist (WCAG POUR principles)'):
        st.markdown(
            "The table above only covers what this tool can verify automatically for this file type "
            "(e.g., alt text, headings, table headers). Many real accessibility requirements can't be "
            "checked by software at all — they need a human to look and judge. The checklist below is "
            "that missing half: a plain-language reference list of the WCAG requirements automated tools "
            "typically can't verify, organized under the four WCAG principles (**P**erceivable, **O**perable, "
            "**U**nderstandable, **R**obust — \"POUR\"). Use it as a guide when writing feedback to a document's "
            "author, or as a manual checklist to walk through yourself."
        )
        checklist_rows = [
            {'Principle': c['principle'], 'Guidance': c['item'], 'WCAG': c['wcag']}
            for c in kb.get('accessibility_checklist', [])
        ]
        st.dataframe(checklist_rows, use_container_width=True, hide_index=True)

    st.markdown('#### Clarity & understandability suggestions')
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
        with st.expander(f"Long sentences ({len(clarity['long_sentences'])})"):
            for s in clarity['long_sentences']:
                st.write(f"**{s['word_count']} words:** {s['text']}")

    if clarity['undefined_acronyms']:
        with st.expander(f"Undefined acronyms ({len(clarity['undefined_acronyms'])})"):
            for a in clarity['undefined_acronyms']:
                st.write(f"**{a['acronym']}** — {a['expansion']}")

    with st.expander('Appendix: guideline recommendation wording'):
        for rec in kb['recommendations']:
            st.markdown(f"**Recommendation {rec['id']} (Level {'/'.join(rec['level'])})**")
            st.write(rec['text'])
            st.divider()


# ---------------------------------------------------------------------------
# Layout: Check a document | Previously reviewed
# ---------------------------------------------------------------------------
st.title('📋 DoC Guideline Fact-Checker')

kb = get_kb()

tab_check, tab_history = st.tabs(['Check a document', 'Previously reviewed'])

with tab_check:
    st.write(
        'Upload a document (PDF, DOCX, PPTX, TXT, or MD) and this tool will flag '
        'places where the content may conflict with the DoC clinical guideline.'
    )

    uploaded_file = st.file_uploader(
        'Choose a file to check',
        type=[ext.lstrip('.') for ext in SUPPORTED_EXTS],
    )
    nickname_input = st.text_input('Nickname for this document (optional)', placeholder='e.g., "Family handout draft 2"')

    if uploaded_file is not None and st.button('Run fact-check', type='primary'):
        with st.spinner('Checking document against guideline...'):
            temp_dir = tempfile.mkdtemp(prefix='doc_factcheck_')
            in_path = os.path.join(temp_dir, uploaded_file.name)
            with open(in_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            try:
                result = run_checks(in_path, kb)
                counts = counts_from_flags(result['flags'])
                entry = add_history_entry(uploaded_file.name, result, counts, nickname=nickname_input)

                st.success('Fact-check complete — saved to "Previously reviewed."')
                render_result(
                    entry['filename'], entry['counts'], entry['flags'], entry['coverage'],
                    entry['clarity'], entry['accessibility'], entry['accessibility_score'], kb,
                    nickname=entry.get('nickname'), checked_at=entry['checked_at'], key_suffix=f"check_{entry['id']}",
                )

            except Exception as e:
                st.error(f'Could not process this file: {e}')

with tab_history:
    history = load_history()

    if not history:
        st.info('No materials reviewed yet. Check a document in the first tab to get started.')
    else:
        st.write(f'{len(history)} material(s) reviewed so far.')

        def label_for(e):
            name = e.get('nickname', '').strip() or e['filename']
            return f"{name} — {e['checked_at']}"

        options = [label_for(e) for e in history]
        selected_label = st.selectbox('Select a previously reviewed material', options)
        selected_entry = history[options.index(selected_label)]

        del_col, _ = st.columns([1, 4])
        with del_col:
            confirm_delete = st.checkbox('Confirm delete', key=f"confirm_delete_{selected_entry['id']}")
            if st.button('🗑️ Delete this entry', disabled=not confirm_delete, key=f"delete_{selected_entry['id']}"):
                delete_history_entry(selected_entry['id'])
                st.success('Deleted.')
                st.rerun()

        render_result(
            selected_entry['filename'], selected_entry['counts'], selected_entry['flags'],
            selected_entry['coverage'], selected_entry.get('clarity', default_clarity()),
            selected_entry.get('accessibility', default_accessibility()),
            selected_entry.get('accessibility_score', default_accessibility_score()), kb,
            nickname=selected_entry.get('nickname'), checked_at=selected_entry['checked_at'], key_suffix=f"history_{selected_entry['id']}",
        )
