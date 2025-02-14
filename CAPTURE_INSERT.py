import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

def extract_proposals(text):
    proposals = []
    proposal_pattern = re.compile(r'Proposal (\d+):\s*(.*?)\nFor -- ([\d,]+) Against -- ([\d,]+) Abstain -- ([\d,]+) BrokerNon-Votes -- ([\d,]+)', re.S)
    matches = proposal_pattern.findall(text)
    
    for match in matches:
        proposal_number, proposal_text, votes_for, votes_against, votes_abstain, votes_broker = match
        votes_for, votes_against, votes_abstain, votes_broker = map(lambda x: x.replace(',', ''), [votes_for, votes_against, votes_abstain, votes_broker])
        resolution_outcome = "Approved" if int(votes_for) > int(votes_against) else "Not Approved"
        proposals.extend([
            ["Proposal Proxy Year", "2024"],
            ["Resolution Outcome", f"{resolution_outcome} ({votes_for} > {votes_against})"],
            ["Proposal Text", proposal_text.strip()],
            ["Mgmt Proposal Category", ""],
            ["Vote Results - For", votes_for],
            ["Vote Results - Against", votes_against],
            ["Vote Results - Abstained", votes_abstain],
            ["Vote Results - Withheld", ""],
            ["Vote Results - Broker Non-Votes", votes_broker],
            ["Proposal Vote Results Total", ""],
            ["", ""]  # Blank row for spacing
        ])
    
    return proposals

def extract_director_votes(text):
    directors = []
    director_pattern = re.compile(r'([A-Za-z\s]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)')
    matches = director_pattern.findall(text)
    
    for match in matches:
        name, votes_for, votes_withheld, votes_broker = match
        votes_for, votes_withheld, votes_broker = map(lambda x: x.replace(',', ''), [votes_for, votes_withheld, votes_broker])
        directors.extend([
            ["Director Election Year", "2024"],
            ["Individual", name.strip()],
            ["Director Votes For", votes_for],
            ["Director Votes Against", ""],
            ["Director Votes Abstained", ""],
            ["Director Votes Withheld", votes_withheld],
            ["Director Votes Broker-Non-Votes", votes_broker],
            ["", ""]  # Blank row for spacing
        ])
    
    return directors

def process_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
    
    proposals_data = extract_proposals(text)
    directors_data = extract_director_votes(text)
    
    proposals_df = pd.DataFrame(proposals_data, columns=["Field", "Value"])
    directors_df = pd.DataFrame(directors_data, columns=["Field", "Value"])
    
    return proposals_df, directors_df

def generate_excel(proposals_df, directors_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        proposals_df.to_excel(writer, sheet_name='Proposal Data', index=False)
        directors_df.to_excel(writer, sheet_name='Non-Proposal Data', index=False)
    
    output.seek(0)
    return output

def main():
    st.title("AGM Data Extractor & Formatter")
    uploaded_file = st.file_uploader("Upload AGM results PDF", type=["pdf"])
    
    if uploaded_file is not None:
        proposals_df, directors_df = process_pdf(uploaded_file)
        st.write("### Extracted Proposal Data")
        st.dataframe(proposals_df)
        
        st.write("### Extracted Non-Proposal Data")
        st.dataframe(directors_df)
        
        excel_file = generate_excel(proposals_df, directors_df)
        st.download_button(
            label="Download Extracted Data as Excel",
            data=excel_file,
            file_name="AGM_Extracted_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
