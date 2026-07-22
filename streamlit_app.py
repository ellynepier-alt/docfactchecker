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


def add_history_entry(filename, result, counts):
    entries = load_history()
    entry = {
        'id': uuid.uuid4().hex[:10],
        'filename': filename,
        'checked_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'counts': counts,
        'flags': result['flags'],
        'coverage': result['coverage'],
    }
    entries.insert(0, entry)
    save_history(entries)
    return entry


def counts_from_flags(flags):
    counts = {s: 0 for s in ['high', 'medium', 'low', 'review']}
    for flag in flags:
        counts[flag['severity']] = counts.get(flag['severity'], 0) + 1
    return counts


# ---------------------------------------------------------------------------
# Shared rendering for a result (used for both fresh checks and history)
# ---------------------------------------------------------------------------
def render_result(filename, counts, flags, coverage, kb, checked_at=None, key_suffix=None):
    caption = filename if not checked_at else f'{filename} — checked {checked_at}'
    st.subheader(caption)

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
        'Upload a document (PDF, DOCX, TXT, or MD) and this tool will flag '
        'places where the content may conflict with the DoC clinical guideline.'
    )

    uploaded_file = st.file_uploader(
        'Choose a file to check',
        type=[ext.lstrip('.') for ext in SUPPORTED_EXTS],
    )

    if uploaded_file is not None and st.button('Run fact-check', type='primary'):
        with st.spinner('Checking document against guideline...'):
            temp_dir = tempfile.mkdtemp(prefix='doc_factcheck_')
            in_path = os.path.join(temp_dir, uploaded_file.name)
            with open(in_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            try:
                result = run_checks(in_path, kb)
                counts = counts_from_flags(result['flags'])
                entry = add_history_entry(uploaded_file.name, result, counts)

                st.success('Fact-check complete — saved to "Previously reviewed."')
                render_result(
                    entry['filename'], entry['counts'], entry['flags'], entry['coverage'], kb,
                    checked_at=entry['checked_at'], key_suffix=f"check_{entry['id']}",
                )

            except Exception as e:
                st.error(f'Could not process this file: {e}')

with tab_history:
    history = load_history()

    if not history:
        st.info('No materials reviewed yet. Check a document in the first tab to get started.')
    else:
        st.write(f'{len(history)} material(s) reviewed so far.')

        options = [f"{e['filename']} — {e['checked_at']}" for e in history]
        selected_label = st.selectbox('Select a previously reviewed material', options)
        selected_entry = history[options.index(selected_label)]

        render_result(
            selected_entry['filename'], selected_entry['counts'], selected_entry['flags'],
            selected_entry['coverage'], kb,
            checked_at=selected_entry['checked_at'], key_suffix=f"history_{selected_entry['id']}",
        )
