import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
import re
from io import BytesIO

def extract_text_from_pdf(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    
    if not text.strip():  # If no text was extracted, use OCR
        images = convert_from_bytes(pdf_file.read())
        text = "\n".join(pytesseract.image_to_string(img) for img in images)
    
    return text

def extract_proposals(text):
    proposals = []
    
    # Debugging output
    print("Extracted Text:", text[:1000])  # Print the first 1000 characters to check the text format
    
    proposal_pattern = re.compile(
        r'Proposal\s*(?:No\.\s*)?(\d+)\s*[–-]\s*(.*?)\nFor\s+([\d,]+)?\s+Against\s+([\d,]+)?\s+Abstain\s+([\d,]+)?(?:\s+Withheld\s+([\d,]+))?\s+Broker Non-Votes\s+([\d,]+)?', re.S
    )
    matches = proposal_pattern.findall(text)
    
    if not matches:
        print("No proposals matched. Check the formatting of the document.")
    
    for match in matches:
        proposal_number, proposal_text, votes_for, votes_against, votes_abstain, votes_withheld, votes_broker = match
        votes_for = votes_for.replace(',', '') if votes_for else ""
        votes_against = votes_against.replace(',', '') if votes_against else ""
        votes_abstain = votes_abstain.replace(',', '') if votes_abstain else ""
        votes_withheld = votes_withheld.replace(',', '') if votes_withheld else ""
        votes_broker = votes_broker.replace(',', '') if votes_broker else ""
        
        resolution_outcome = "Approved" if votes_for and votes_against and int(votes_for) > int(votes_against) else "Not Approved"
        proposal_data = [
            ["Proposal Proxy Year", "2024"],
            ["Resolution Outcome", f"{resolution_outcome} ({votes_for} > {votes_against})"],
            ["Proposal Text", proposal_text.strip()],
            ["Mgmt Proposal Category", ""],
            ["Vote Results - For", votes_for],
            ["Vote Results - Against", votes_against],
            ["Vote Results - Abstained", votes_abstain],
            ["Vote Results - Withheld", votes_withheld],
            ["Vote Results - Broker Non-Votes", votes_broker],
            ["Proposal Vote Results Total", ""],
            ["", ""]  # Blank row for spacing
        ]
        proposals.extend(proposal_data)
    
    return proposals

def process_pdf(pdf_file):
    text = extract_text_from_pdf(pdf_file)
    proposals_data = extract_proposals(text)
    proposals_df = pd.DataFrame(proposals_data, columns=["Field", "Value"])
    
    return proposals_df

def generate_excel(proposals_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not proposals_df.empty:
            proposals_df.to_excel(writer, sheet_name='Proposal Data', index=False)
    
    output.seek(0)
    return output

def main():
    st.title("AGM Data Extractor & Formatter")
    uploaded_file = st.file_uploader("Upload AGM results PDF", type=["pdf"])
    
    if uploaded_file is not None:
        proposals_df = process_pdf(uploaded_file)
        
        if not proposals_df.empty:
            st.write("### Extracted Proposal Data")
            st.dataframe(proposals_df)
        else:
            st.write("No Proposal Data Found - Please check if the document format is correct.")
        
        excel_file = generate_excel(proposals_df)
        st.download_button(
            label="Download Extracted Data as Excel",
            data=excel_file,
            file_name="AGM_Extracted_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
