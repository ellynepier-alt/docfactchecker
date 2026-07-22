import json
import os
import tempfile
import uuid
from datetime import datetime

import streamlit as st

from factcheck_engine import SUPPORTED_EXTS, load_kb, run_checks, make_report

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KB_PATH = os.path.join(BASE_DIR, 'doc_guidelines_kb.json')

HISTORY_DIR = os.path.join(BASE_DIR, 'data', 'history')
REPORTS_DIR = os.path.join(HISTORY_DIR, 'reports')
INDEX_PATH = os.path.join(HISTORY_DIR, 'index.json')

os.makedirs(REPORTS_DIR, exist_ok=True)

st.set_page_config(page_title='DoC Guideline Fact-Checker', page_icon='📋', layout='wide')


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


def add_history_entry(filename, result, counts, docx_path):
    entries = load_history()
    entry = {
        'id': uuid.uuid4().hex[:10],
        'filename': filename,
        'checked_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'counts': counts,
        'flags': result['flags'],
        'coverage': result['coverage'],
        'docx_path': docx_path,
    }
    entries.insert(0, entry)
    save_history(entries)
    return entry


# ---------------------------------------------------------------------------
# Shared rendering for a result (used for both fresh checks and history)
# ---------------------------------------------------------------------------
def render_result(filename, counts, flags, docx_path=None, checked_at=None, key_suffix=None):
    caption = filename if not checked_at else f'{filename} — checked {checked_at}'
    st.subheader(caption)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric('High', counts['high'])
    col2.metric('Medium', counts['medium'])
    col3.metric('Low', counts['low'])
    col4.metric('Review', counts['review'])

    st.markdown('#### Flags')
    if not flags:
        st.info('No flags raised. Human review is still recommended.')
    else:
        for flag in flags:
            with st.expander(f"[{flag['severity'].upper()}] {flag['kind']} — {flag['matched']}"):
                st.write(f"**Issue:** {flag['issue']}")
                if flag['rec']:
                    st.write(f"**Related recommendation:** {flag['rec']}")
                if flag['context']:
                    st.caption(f"Context: {flag['context']}")

    if docx_path and os.path.exists(docx_path):
        with open(docx_path, 'rb') as f:
            st.download_button(
                label='Download full report (.docx)',
                data=f.read(),
                file_name=f'DoC_factcheck_{filename}.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                key=f'download_{key_suffix}',
            )


def counts_from_flags(flags):
    counts = {s: 0 for s in ['high', 'medium', 'low', 'review']}
    for flag in flags:
        counts[flag['severity']] = counts.get(flag['severity'], 0) + 1
    return counts


# ---------------------------------------------------------------------------
# Layout: Check a document | Previously reviewed
# ---------------------------------------------------------------------------
st.title('📋 DoC Guideline Fact-Checker')

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
                kb = load_kb(KB_PATH)
                result = run_checks(in_path, kb)
                counts = counts_from_flags(result['flags'])

                entry_id = uuid.uuid4().hex[:10]
                stored_docx_path = os.path.join(REPORTS_DIR, f'{entry_id}.docx')
                make_report(result, kb, stored_docx_path)

                entry = add_history_entry(uploaded_file.name, result, counts, stored_docx_path)

                st.success('Fact-check complete — saved to "Previously reviewed."')
                render_result(
                    entry['filename'], entry['counts'], entry['flags'],
                    docx_path=entry['docx_path'], checked_at=entry['checked_at'],
                    key_suffix=f"check_{entry['id']}",
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
            docx_path=selected_entry['docx_path'], checked_at=selected_entry['checked_at'],
            key_suffix=f"history_{selected_entry['id']}",
        )
