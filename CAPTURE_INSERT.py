import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.title("AGM Results Extractor & Formatter")

st.markdown("""
This app extracts proposal data and director election results from an AGM results PDF,
formats the data into two sheets, and then provides an Excel file for download.
""")

# --- Function to extract text from the uploaded PDF ---
def extract_text_from_pdf(pdf_file):
    text = ""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
    return text

# --- Function to extract proposal data using regex ---
def extract_proposals(text):
    proposals = []
    # Regex pattern to match a proposal block.
    # This assumes the proposal block starts with "Proposal <number>:" and then proposal text followed by a line with votes.
    proposal_pattern = re.compile(
        r"Proposal\s*\d+:\s*(?P<text>.*?)(?=\nFor\s*--)\nFor\s*--\s*(?P<for>[\d,]+)\s*Against\s*--\s*(?P<against>[\d,]+)\s*Abstain\s*--\s*(?P<abstain>[\d,]+).*?(?:Broker\s*Non-Votes|BrokerNon-Votes)\s*--\s*(?P<broker>[\d,]+)",
        re.DOTALL | re.IGNORECASE
    )
    for match in proposal_pattern.finditer(text):
        proposal_text = match.group("text").strip().replace("\n", " ")
        vote_for = match.group("for").strip()
        vote_against = match.group("against").strip()
        vote_abstain = match.group("abstain").strip()
        vote_broker = match.group("broker").strip()
        
        # Attempt to extract a year from the proposal text (e.g., fiscal year). Default to blank if not found.
        year_match = re.search(r'\b(20\d{2})\b', proposal_text)
        proposal_year = year_match.group(1) if year_match else ""
        
        # Calculate Resolution Outcome
        try:
            for_val = int(vote_for.replace(",", ""))
            against_val = int(vote_against.replace(",", ""))
            if for_val > against_val:
                resolution = f"Approved ({vote_for}>{vote_against})"
            else:
                resolution = f"Not Approved ({vote_for}>{vote_against})"
        except Exception:
            resolution = ""
        
        proposals.append({
            "Proposal Proxy Year": proposal_year,
            "Resolution Outcome": resolution,
            "Proposal Text": proposal_text,
            "Mgmt Proposal Category": "",  # Leave blank
            "Vote Results - For": vote_for,
            "Vote Results - Against": vote_against,
            "Vote Results - Abstained": vote_abstain,
            "Vote Results - Withheld": "",  # Since not given
            "Vote Results - Broker Non-Votes": vote_broker,
            "Proposal Vote Results Total": ""  # Leave blank
        })
    return proposals

# --- Function to extract director election data ---
def extract_director_elections(text):
    directors = []
    lines = text.splitlines()
    header_index = None
    
    # Find the header row that contains director election info (assumes header includes "Nominee" and "For")
    for idx, line in enumerate(lines):
        if "Nominee" in line and "For" in line:
            header_index = idx
            break

    if header_index is None:
        return directors  # No director data found

    # Process each subsequent line (until an empty line or until data stops)
    for line in lines[header_index+1:]:
        if not line.strip():
            continue  # skip blank lines
        # Split line by 2 or more spaces (to account for table-like formatting)
        parts = re.split(r'\s{2,}', line.strip())
        if len(parts) < 3:
            continue  # not enough data
        # Expecting at least: [Individual, For, Withheld, Broker Non-Votes] 
        # Sometimes the Broker Non-Votes column may be the 3rd or 4th element.
        # We'll try to handle both cases.
        name = parts[0]
        vote_for = parts[1]
        # Assume Withheld is always provided as the next column.
        vote_withheld = parts[2] if len(parts) >= 3 else ""
        vote_broker = parts[3] if len(parts) >= 4 else ""
        
        directors.append({
            "Director Election Year": "2024",  # Fixed as per instructions
            "Individual": name,
            "Director Votes For": vote_for,
            "Director Votes Against": "",  # Leave blank if not available
            "Director Votes Abstained": "",  # Leave blank if not available
            "Director Votes Withheld": vote_withheld,
            "Director Votes Broker-Non-Votes": vote_broker
        })
    return directors

# --- Function to create vertical layout for proposals ---
def build_proposals_sheet(proposals):
    rows = []
    for proposal in proposals:
        rows.append(["Proposal Proxy Year:", proposal["Proposal Proxy Year"]])
        rows.append(["Resolution Outcome:", proposal["Resolution Outcome"]])
        rows.append(["Proposal Text:", f'"{proposal["Proposal Text"]}"'])
        rows.append(["Mgmt Proposal Category:", proposal["Mgmt Proposal Category"]])
        rows.append(["Vote Results - For:", proposal["Vote Results - For"]])
        rows.append(["Vote Results - Against:", proposal["Vote Results - Against"]])
        rows.append(["Vote Results - Abstained:", proposal["Vote Results - Abstained"]])
        rows.append(["Vote Results - Withheld:", proposal["Vote Results - Withheld"]])
        rows.append(["Vote Results - Broker Non-Votes:", proposal["Vote Results - Broker Non-Votes"]])
        rows.append(["Proposal Vote Results Total:", proposal["Proposal Vote Results Total"]])
        # Add a blank row as a separator
        rows.append(["", ""])
    df = pd.DataFrame(rows, columns=["Field", "Value"])
    return df

# --- Function to create vertical layout for director elections ---
def build_directors_sheet(directors):
    rows = []
    for director in directors:
        rows.append(["Director Election Year:", director["Director Election Year"]])
        rows.append(["Individual:", director["Individual"]])
        rows.append(["Director Votes For:", director["Director Votes For"]])
        rows.append(["Director Votes Against:", director["Director Votes Against"]])
        rows.append(["Director Votes Abstained:", director["Director Votes Abstained"]])
        rows.append(["Director Votes Withheld:", director["Director Votes Withheld"]])
        rows.append(["Director Votes Broker-Non-Votes:", director["Director Votes Broker-Non-Votes"]])
        # Add a blank row as a separator
        rows.append(["", ""])
    df = pd.DataFrame(rows, columns=["Field", "Value"])
    return df

# --- File uploader ---
uploaded_file = st.file_uploader("Upload AGM Results PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Extracting text from PDF..."):
        pdf_text = extract_text_from_pdf(uploaded_file)
    
    st.subheader("Extracted PDF Text Preview")
    st.text_area("", pdf_text, height=200)

    # --- Extract proposal and director data ---
    proposals = extract_proposals(pdf_text)
    directors = extract_director_elections(pdf_text)
    
    st.write(f"Found **{len(proposals)}** proposal(s) and **{len(directors)}** director record(s).")

    # --- Build the sheets ---
    proposals_df = build_proposals_sheet(proposals)
    directors_df = build_directors_sheet(directors)

    # --- Create Excel file with two sheets in memory ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        proposals_df.to_excel(writer, sheet_name="Proposal sheet", index=False)
        directors_df.to_excel(writer, sheet_name="non-proposal sheet", index=False)
        writer.close()
    processed_data = output.getvalue()

    st.download_button(
        label="Download Extracted Data as Excel",
        data=processed_data,
        file_name="AGM_Extracted_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Please upload an AGM results PDF to begin extraction.")
