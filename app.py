import streamlit as st
from factcheck_engine import load_kb, extract_text_from_uploaded, run_factcheck, build_word_report, SUPPORTED_EXTS

st.set_page_config(page_title='DoC Guideline Fact-Checker', page_icon='🧠', layout='centered')

st.title('DoC Guideline Fact-Checker')
st.caption('A human-in-the-loop review aid for AAN-ACRM-NIDILRR Disorders of Consciousness guideline alignment')

st.markdown('''
Upload a toolkit material and generate a Word report that flags candidate issues against the Disorders of Consciousness guideline.

**Supported files:** PDF, DOCX, TXT, MD  
**Important:** Do not upload PHI or identifiable patient data unless the deployed environment is approved for that use.
''')

with st.expander('What this checks'):
    st.markdown('''
- Deprecated or discouraged terminology
- Potential contradictions of guideline recommendations
- Key facts such as amantadine dose and chronic VS/UWS timeframes
- Evidence-level mismatches
- Coverage map of guideline recommendations
    ''')

uploaded = st.file_uploader('Choose a toolkit material', type=['pdf', 'docx', 'txt', 'md'])

if uploaded is not None:
    kb = load_kb('doc_guidelines_kb.json')
    try:
        text = extract_text_from_uploaded(uploaded)
        st.success(f'Extracted {len(text):,} characters from {uploaded.name}.')

        if st.button('Run fact check', type='primary'):
            result = run_factcheck(uploaded.name, text, kb)
            counts = {'high': 0, 'medium': 0, 'low': 0, 'review': 0}
            for flag in result['flags']:
                counts[flag['severity']] = counts.get(flag['severity'], 0) + 1

            st.subheader('Summary')
            c1, c2, c3, c4 = st.columns(4)
            c1.metric('High', counts['high'])
            c2.metric('Medium', counts['medium'])
            c3.metric('Low', counts['low'])
            c4.metric('Review', counts['review'])

            if result['flags']:
                st.subheader('Flags')
                for flag in result['flags']:
                    if flag['severity'] == 'high':
                        st.error(f"[{flag['severity'].upper()}] {flag['kind']}: {flag['matched']}")
                    elif flag['severity'] == 'medium':
                        st.warning(f"[{flag['severity'].upper()}] {flag['kind']}: {flag['matched']}")
                    else:
                        st.info(f"[{flag['severity'].upper()}] {flag['kind']}: {flag['matched']}")
                    st.write(flag['issue'])
                    if flag['rec']:
                        st.caption(f"Recommendation: {flag['rec']}")
                    if flag['context']:
                        st.caption(f"Context: {flag['context']}")
            else:
                st.success('No flags raised. Human review is still recommended.')

            st.subheader('Recommendation coverage')
            st.write(', '.join(result['coverage']) if result['coverage'] else 'No recommendation topics detected by keyword matching.')

            report = build_word_report(result, kb)
            st.download_button(
                label='Download Word report',
                data=report,
                file_name='DoC_factcheck_report.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
    except Exception as e:
        st.error(f'Could not process this file: {e}')

st.divider()
st.markdown('''
**Governance note:** This tool is for educational material review and should be hosted only in an approved environment. It is not a clinical decision support system and should not be used with PHI unless the deployment has been approved for that purpose.
''')
