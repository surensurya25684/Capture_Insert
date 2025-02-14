import streamlit as st
import pdfplumber
import re
import io
from openpyxl import Workbook

st.set_page_config(page_title="AGM Results Extractor", layout="wide")

st.title("AGM Results Extractor and Formatter")

st.markdown("""
This app extracts proposal-related details and director election results from an AGM PDF document.
The proposals will be formatted in a vertical layout in one Excel sheet, and the director election results in another.
""")

uploaded_file = st.file_uploader("Upload AGM PDF Document", type=["pdf"])

def extract_pdf_text(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def parse_proposals(text):
    """
    Parses proposals from the text.
    Expected format (example):
    Proposal 2: The Company’s stockholders ratified the selection of PricewaterhouseCoopers LLP as the Company’s independent
    registered accounting firm for the fiscal year ending December 31, 2024.
    For -- 100,286,297 Against -- 41,446 Abstain -- 653  Broker Non-Votes --0
    """
    proposals = []
    # Regex pattern to capture:
    # 1: Proposal number (optional, for ordering)
    # 2: Proposal text (could be multi-line, non-greedy)
    # 3: For votes, 4: Against votes, 5: Abstained, 6: Broker Non-Votes
    pattern = re.compile(
        r"Proposal\s*\d+:\s*(.*?)\s*\n\s*For\s*--\s*([\d,]+)\s*Against\s*--\s*([\d,]+)\s*Abstain\s*--\s*([\d,]+)\s*Broker Non-Votes\s*--\s*([\d,]+)",
        re.DOTALL | re.IGNORECASE,
    )
    matches = pattern.findall(text)
    for match in matches:
        proposal_text, vote_for, vote_against, vote_abstained, vote_broker = match

        # Clean extra spaces and newlines from proposal text
        proposal_text = " ".join(proposal_text.strip().split())
        # Determine resolution outcome
        try:
            for_val = int(vote_for.replace(",", ""))
            against_val = int(vote_against.replace(",", ""))
        except:
            for_val = against_val = 0

        if for_val > against_val:
            outcome = f"Approved ({vote_for} > {vote_against})"
        else:
            outcome = f"Not Approved ({vote_for} > {vote_against})"
        # Use a fixed Proposal Proxy Year (could also be extracted from text if needed)
        proxy_year = "2024"
        proposals.append({
            "Proposal Proxy Year": proxy_year,
            "Resolution Outcome": outcome,
            "Proposal Text": f'"{proposal_text}"',
            "Mgmt Proposal Category": "",
            "Vote Results - For": vote_for,
            "Vote Results - Against": vote_against,
            "Vote Results - Abstained": vote_abstained,
            "Vote Results - Withheld": "",  # not provided
            "Vote Results - Broker Non-Votes": vote_broker,
            "Proposal Vote Results Total": ""
        })
    return proposals

def parse_director_elections(text):
    """
    Parses director election data from the text.
    Expected table format example:
    
    Nominee                                                        For                    Withheld                              Broker Non-Votes
    Richard Mack                                             96,399,316             2,877,317                             1,051,763 
    Michael McGillis                                        93,192,011              6,084,622                             1,051,763
    """
    directors = []
    lines = text.splitlines()
    header_found = False
    for line in lines:
        if "Nominee" in line and "For" in line:
            header_found = True
            continue
        if header_found:
            if line.strip() == "":
                continue
            # Use two or more spaces to split columns
            parts = re.split(r'\s{2,}', line.strip())
            # Expecting at least 4 parts: [Name, For, Withheld, Broker Non-Votes]
            if len(parts) >= 4:
                name = parts[0]
                votes_for = parts[1]
                votes_withheld = parts[2]
                votes_broker = parts[3]
                # Director Votes Against and Abstained are not given so leave blank.
                directors.append({
                    "Director Election Year": "2024",
                    "Individual": name,
                    "Director Votes For": votes_for,
                    "Director Votes Against": "",
                    "Director Votes Abstained": "",
                    "Director Votes Withheld": votes_withheld,
                    "Director Votes Broker-Non-Votes": votes_broker
                })
            else:
                # If the row does not match the expected format, skip.
                continue
    return directors

def create_excel(proposals, directors):
    wb = Workbook()
    
    # Proposal Sheet (first sheet)
    ws1 = wb.active
    ws1.title = "Proposal"
    
    # For each proposal, write the vertical block of rows.
    # We will use two columns: C1 (field label) and C2 (value)
    row_idx = 1
    for proposal in proposals:
        ws1.cell(row=row_idx, column=1, value="Proposal Proxy Year:")
        ws1.cell(row=row_idx, column=2, value=proposal["Proposal Proxy Year"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Resolution Outcome:")
        ws1.cell(row=row_idx, column=2, value=proposal["Resolution Outcome"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Proposal Text:")
        ws1.cell(row=row_idx, column=2, value=proposal["Proposal Text"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Mgmt Proposal Category:")
        ws1.cell(row=row_idx, column=2, value=proposal["Mgmt Proposal Category"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Vote Results - For:")
        ws1.cell(row=row_idx, column=2, value=proposal["Vote Results - For"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Vote Results - Against:")
        ws1.cell(row=row_idx, column=2, value=proposal["Vote Results - Against"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Vote Results - Abstained:")
        ws1.cell(row=row_idx, column=2, value=proposal["Vote Results - Abstained"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Vote Results - Withheld:")
        ws1.cell(row=row_idx, column=2, value=proposal["Vote Results - Withheld"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Vote Results - Broker Non-Votes:")
        ws1.cell(row=row_idx, column=2, value=proposal["Vote Results - Broker Non-Votes"])
        row_idx += 1
        
        ws1.cell(row=row_idx, column=1, value="Proposal Vote Results Total:")
        ws1.cell(row=row_idx, column=2, value=proposal["Proposal Vote Results Total"])
        row_idx += 1
        
        # Add a blank separator row
        ws1.cell(row=row_idx, column=1, value="--------------------------------------------------")
        row_idx += 1

    # Non-Proposal Sheet (Director Election)
    ws2 = wb.create_sheet(title="Non-Proposal")
    row_idx = 1
    for director in directors:
        ws2.cell(row=row_idx, column=1, value="Director Election Year:")
        ws2.cell(row=row_idx, column=2, value=director["Director Election Year"])
        row_idx += 1
        
        ws2.cell(row=row_idx, column=1, value="Individual:")
        ws2.cell(row=row_idx, column=2, value=director["Individual"])
        row_idx += 1
        
        ws2.cell(row=row_idx, column=1, value="Director Votes For:")
        ws2.cell(row=row_idx, column=2, value=director["Director Votes For"])
        row_idx += 1
        
        ws2.cell(row=row_idx, column=1, value="Director Votes Against:")
        ws2.cell(row=row_idx, column=2, value=director["Director Votes Against"])
        row_idx += 1
        
        ws2.cell(row=row_idx, column=1, value="Director Votes Abstained:")
        ws2.cell(row=row_idx, column=2, value=director["Director Votes Abstained"])
        row_idx += 1
        
        ws2.cell(row=row_idx, column=1, value="Director Votes Withheld:")
        ws2.cell(row=row_idx, column=2, value=director["Director Votes Withheld"])
        row_idx += 1
        
        ws2.cell(row=row_idx, column=1, value="Director Votes Broker-Non-Votes:")
        ws2.cell(row=row_idx, column=2, value=director["Director Votes Broker-Non-Votes"])
        row_idx += 1
        
        # Blank separator row
        ws2.cell(row=row_idx, column=1, value="--------------------------------------------------")
        row_idx += 1

    # Save workbook to a BytesIO stream
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if uploaded_file is not None:
    # Extract text from the PDF
    pdf_text = extract_pdf_text(uploaded_file)
    
    # Parse the proposals and director election data
    proposals = parse_proposals(pdf_text)
    directors = parse_director_elections(pdf_text)
    
    st.write("### Extracted Proposal Data")
    if proposals:
        for idx, prop in enumerate(proposals, start=1):
            st.write(f"**Proposal {idx}:** {prop['Proposal Text']}")
            st.write(f"Votes For: {prop['Vote Results - For']}, Votes Against: {prop['Vote Results - Against']}, Abstained: {prop['Vote Results - Abstained']}, Broker Non-Votes: {prop['Vote Results - Broker Non-Votes']}")
            st.write("---")
    else:
        st.warning("No proposals found.")
    
    st.write("### Extracted Director Election Data")
    if directors:
        for idx, dir_data in enumerate(directors, start=1):
            st.write(f"**Director {idx}:** {dir_data['Individual']} - Votes For: {dir_data['Director Votes For']}, Withheld: {dir_data['Director Votes Withheld']}, Broker Non-Votes: {dir_data['Director Votes Broker-Non-Votes']}")
            st.write("---")
    else:
        st.warning("No director election data found.")
    
    # Create Excel file with two sheets
    excel_data = create_excel(proposals, directors)
    
    st.download_button(
        label="Download Excel File",
        data=excel_data,
        file_name="AGM_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
