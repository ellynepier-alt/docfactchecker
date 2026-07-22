import os
import tempfile
from pathlib import Path

import streamlit as st

from factcheck_engine import SUPPORTED_EXTS, load_kb, run_checks, make_report

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KB_PATH = os.path.join(BASE_DIR, 'doc_guidelines_kb.json')

st.set_page_config(page_title='DoC Guideline Fact-Checker', page_icon='📋')

st.title('📋 DoC Guideline Fact-Checker')
st.write(
    'Upload a document (PDF, DOCX, TXT, or MD) and this tool will flag '
    'places where the content may conflict with the DoC clinical guideline.'
)

uploaded_file = st.file_uploader(
    'Choose a file to check',
    type=[ext.lstrip('.') for ext in SUPPORTED_EXTS],
)

if uploaded_file is not None:
    if st.button('Run fact-check', type='primary'):
        with st.spinner('Checking document against guideline...'):
            temp_dir = tempfile.mkdtemp(prefix='doc_factcheck_')
            in_path = os.path.join(temp_dir, uploaded_file.name)
            with open(in_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            try:
                kb = load_kb(KB_PATH)
                result = run_checks(in_path, kb)

                out_path = os.path.join(temp_dir, 'DoC_factcheck_report.docx')
                make_report(result, kb, out_path)

                counts = {s: 0 for s in ['high', 'medium', 'low', 'review']}
                for flag in result['flags']:
                    counts[flag['severity']] = counts.get(flag['severity'], 0) + 1

                st.success('Fact-check complete.')

                col1, col2, col3, col4 = st.columns(4)
                col1.metric('High', counts['high'])
                col2.metric('Medium', counts['medium'])
                col3.metric('Low', counts['low'])
                col4.metric('Review', counts['review'])

                st.subheader('Flags')
                if not result['flags']:
                    st.info('No flags raised. Human review is still recommended.')
                else:
                    for flag in result['flags']:
                        with st.expander(f"[{flag['severity'].upper()}] {flag['kind']} — {flag['matched']}"):
                            st.write(f"**Issue:** {flag['issue']}")
                            if flag['rec']:
                                st.write(f"**Related recommendation:** {flag['rec']}")
                            if flag['context']:
                                st.caption(f"Context: {flag['context']}")

                with open(out_path, 'rb') as f:
                    st.download_button(
                        label='Download full report (.docx)',
                        data=f.read(),
                        file_name='DoC_factcheck_report.docx',
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    )

            except Exception as e:
                st.error(f'Could not process this file: {e}')
