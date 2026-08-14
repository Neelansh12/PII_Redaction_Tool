import streamlit as st
import docx
import io
from redact_pii import PIIRedactor

st.set_page_config(page_title="PII Redaction Engine", page_icon="🔒", layout="centered")

st.title("🔒 PII Redaction Tool")
st.write("Upload a `.docx` document to automatically detect and replace Personally Identifiable Information (PII) with synthetic data.")

uploaded_file = st.file_uploader("Choose a DOCX file", type=["docx"])

if uploaded_file is not None:
    redactor = PIIRedactor()
    
    with st.spinner("Processing document and redacting PII..."):
        bio = io.BytesIO()
        counts, mapping = redactor.redact_document(uploaded_file, bio)
        bio.seek(0)

    st.success("✅ Redaction complete!")
    
    st.subheader("📊 Detection Summary")
    if counts:
        cols = st.columns(3)
        sorted_counts = sorted(counts.items())
        for idx, (kind, count) in enumerate(sorted_counts):
            with cols[idx % 3]:
                st.metric(label=kind.replace("_", " ").title(), value=count)
        st.caption(f"**Total unique replacements generated:** {len(mapping)}")
    else:
        st.info("No PII entities detected in the uploaded document.")

    st.download_button(
        label="📥 Download Redacted Document",
        data=bio,
        file_name="Redacted_Output.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )