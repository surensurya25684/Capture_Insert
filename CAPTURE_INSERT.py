import streamlit as st
import re
import pandas as pd
from io import BytesIO
from PyPDF2 import PdfReader

# --------- Helper Functions ---------
def extract_text_from_pdf(pdf_file):
    """Extracts text from a PDF file."""
    pdf_reader = PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def parse_proposals(text):
    """
    Parses proposals from the PDF text.
    Looks for sections starting with "Proposal <number>:" and then extracts:
      - Proposal Proxy Year (from fiscal year text if available)
      - Proposal Text (all text before the votes line)
      - Vote counts from the line containing For, Against, Abstain, Broker Non-Votes.
      - Computes Resolution Outcome (Approved if For > Against)
      - Leaves Mgmt. Proposal Category, Vote Results - Withheld and Proposal Vote Results Total blank.
    """
    proposals = []
    # Split by proposal blocks (each block starts with "Proposal" followed by a number)
    proposal_blocks = re.split(r'(?=Proposal\s*\d+:)', text)
    for block in proposal_blocks:
        if block.strip().startswith("Proposal"):
            # Search for the vote counts line using a regex
            votes_match = re.search(
                r'For\s*--\s*([\d,]+)\s*Against\s*--\s*([\d,]+)\s*Abstain\s*--\s*([\d,]+)\s*Broker Non-Votes\s*--\s*([\d,]+)',
                block
            )
            if votes_match:
                for_votes_str = votes_match.group(1).replace(',', '')
                against_votes_str = votes_match.group(2).replace(',', '')
                abstain_votes_str = votes_match.group(3).replace(',', '')
                broker_votes_str = votes_match.group(4).replace(',', '')
                try:
                    for_votes = int(for_votes_str)
                except:
                    for_votes = 0
                try:
                    against_votes = int(against_votes_str)
                except:
                    against_votes = 0
                try:
                    abstain_votes = int(abstain_votes_str)
                except:
                    abstain_votes = 0
                try:
                    broker_votes = int(broker_votes_str)
                except:
                    broker_votes = 0

                # Compute Resolution Outcome
                if for_votes > against_votes:
                    resolution = f"Approved ({for_votes} > {against_votes})"
                else:
                    resolution = f"Not Approved ({for_votes} <= {against_votes})"

                # Remove the votes part to get the proposal text
                proposal_text = re.split(r'For\s*--', block)[0].strip()
                # Optionally remove the "Proposal X:" label from the beginning.
                proposal_text = re.sub(r'^Proposal\s*\d+:\s*', '', proposal_text, flags=re.IGNORECASE)

                # Extract year from a phrase like "fiscal year ending ... 2024"
                year_match = re.search(r'fiscal year ending.*?(\d{4})', proposal_text, re.IGNORECASE)
                if year_match:
                    proxy_year = year_match.group(1)
                else:
                    # Fallback: look for any standalone 4-digit number (e.g. 2024)
                    year_match = re.search(r'\b(20\d{2})\b', proposal_text)
                    proxy_year = year_match.group(1) if year_match else ""

                proposal = {
                    "Proposal Proxy Year": proxy_year,
                    "Resolution Outcome": resolution,
                    "Proposal Text": proposal_text,
                    "Mgmt. Proposal Category": "",
                    "Vote Results - For": for_votes,
                    "Vote Results - Against": against_votes,
                    "Vote Results - Abstained": abstain_votes,
                    "Vote Results - Withheld": "",
                    "Vote Results - Broker Non-Votes": broker_votes,
                    "Proposal Vote Results Total": ""
                }
                proposals.append(proposal)
    return proposals

def parse_directors(text):
    """
    Parses director election data from the PDF text.
    Expects to find a table starting with a header that includes:
      "Nominee", "For", "Withheld", and "Broker Non-Votes".
    For each director row, the following fields are captured:
      - Director Election Year (set as 2024)
      - Individual (director's name)
      - Director Votes For
      - Director Votes Against (left blank if not given)
      - Director Votes Abstained (left blank if not given)
      - Director Votes Withheld
      - Director Votes Broker-Non-Votes
    """
    directors = []
    # Find the table section starting at "Nominee" (case-insensitive)
    table_match = re.search(r'Nominee\s+For\s+Withheld\s+Broker Non-Votes(.*)', text, re.DOTALL | re.IGNORECASE)
    if table_match:
        table_text = table_match.group(1)
        # Split into lines and iterate over each line
        lines = table_text.splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue
            # Split columns by two or more spaces (this should separate name and the numbers)
            parts = re.split(r'\s{2,}', line)
            # Expect at least 4 parts: [Individual, For, Withheld, Broker Non-Votes]
            if len(parts) >= 4:
                name = parts[0]
                votes_for_str = parts[1].replace(',', '')
                withheld_str = parts[2].replace(',', '')
                broker_str = parts[3].replace(',', '')
                try:
                    votes_for = int(votes_for_str)
                except:
                    votes_for = 0
                try:
                    withheld = int(withheld_str)
                except:
                    withheld = 0
                try:
                    broker = int(broker_str)
                except:
                    broker = 0
                director = {
                    "Director Election Year": "2024",
                    "Individual": name,
                    "Director Votes For": votes_for,
                    "Director Votes Against": "",
                    "Director Votes Abstained": "",
                    "Director Votes Withheld": withheld,
                    "Director Votes Broker-Non-Votes": broker
                }
                directors.append(director)
    return directors

# --------- Streamlit App UI ---------
st.title("AGM Result Extractor")
st.write("Upload your AGM results PDF file to extract proposals and director election data.")

# PDF file uploader
uploaded_file = st.file_uploader("Upload AGM PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Extracting text from PDF..."):
        pdf_text = extract_text_from_pdf(uploaded_file)
    st.success("PDF text extraction complete!")
    
    # Parse proposals and director election data
    proposals = parse_proposals(pdf_text)
    directors = parse_directors(pdf_text)
    
    # Preview Proposals
    st.header("Proposals Preview")
    if proposals:
        df_proposals = pd.DataFrame(proposals)
        st.dataframe(df_proposals)
    else:
        st.warning("No proposals found in the document.")
    
    # Preview Director Election Data (Non-Proposal)
    st.header("Director Elections Preview")
    if directors:
        df_directors = pd.DataFrame(directors)
        st.dataframe(df_directors)
    else:
        st.warning("No director election data found in the document.")
    
    # Create Excel file with two sheets
    if proposals or directors:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            if proposals:
                df_proposals.to_excel(writer, sheet_name="Proposal sheet", index=False)
            if directors:
                df_directors.to_excel(writer, sheet_name="Non-proposal sheet", index=False)
            writer.save()
        processed_data = output.getvalue()
        
        st.download_button(
            label="Download Excel File",
            data=processed_data,
            file_name="AGM_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
