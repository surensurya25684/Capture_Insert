import streamlit as st
import pandas as pd
import io
import PyPDF2
import re

# Function to extract text from the uploaded PDF
def extract_text_from_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def parse_agm_data(text):
    """
    Parses the AGM PDF text to extract proposal data and director election data.
    Proposals are items that include vote counts (e.g. Executive Compensation, Stock Plan, Auditor Ratification).
    The director election block is identified by 'Election of Directors'.
    """
    proposals = []
    directors = []
    
    # Extract the meeting year from a reference like "2024 Annual Meeting"
    meeting_year_match = re.search(r'(\d{4})\s+Annual Meeting', text)
    meeting_year = meeting_year_match.group(1) if meeting_year_match else "2024"
    
    # Split the text into blocks that start with a number and a period (e.g., "1. Election of Directors", "2. Approval ...")
    blocks = re.split(r'\n(?=\d+\.\s)', text)
    
    for block in blocks:
        block = block.strip()
        # Process director election block (usually item 1)
        if re.match(r'\d+\.\s+Election of Directors', block, re.IGNORECASE):
            # Split the block into lines and ignore header lines (like "Nominee For Withheld Broker Non-Votes")
            lines = block.splitlines()
            for line in lines:
                if "Nominee" in line or "Election of Directors" in line:
                    continue
                # Match a line that contains a name and three vote numbers
                match = re.match(r'(.+?)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)', line)
                if match:
                    name = match.group(1).strip()
                    votes_for = match.group(2).strip()
                    votes_withheld = match.group(3).strip()
                    votes_broker = match.group(4).strip()
                    directors.append({
                        "Director Election Year": meeting_year,
                        "Individual": name,
                        "Director Votes For": votes_for,
                        "Director Votes Against": "",   # Not provided in the document
                        "Director Votes Abstained": "",   # Not provided in the document
                        "Director Votes Withheld": votes_withheld,
                        "Director Votes Broker-Non-Votes": votes_broker
                    })
        else:
            # Process proposals (look for blocks that include "Votes Cast")
            if "Votes Cast" in block:
                # Extract the header text (proposal text) by taking the part before "Votes Cast"
                header_part = block.split("Votes Cast")[0].strip()
                # Remove the leading numbering (e.g., "2. ") from the header text
                header_part = re.sub(r'^\d+\.\s*', '', header_part).strip()
                
                # Use regex to extract vote counts
                for_match = re.search(r'For\s+([\d,]+)', block)
                against_match = re.search(r'Against\s+([\d,]+)', block)
                abstentions_match = re.search(r'Abstentions\s+([\d,]+)', block)
                broker_match = re.search(r'Broker Non-Votes\s+([\d,]+)', block)
                
                votes_for = for_match.group(1).strip() if for_match else ""
                votes_against = against_match.group(1).strip() if against_match else ""
                votes_abstained = abstentions_match.group(1).strip() if abstentions_match else ""
                votes_broker = broker_match.group(1).strip() if broker_match else ""
                
                # Determine the resolution outcome based on vote counts (Approved if For > Against)
                try:
                    if votes_for and votes_against:
                        if int(votes_for.replace(',', '')) > int(votes_against.replace(',', '')):
                            resolution_outcome = f"Approved ({votes_for} For > {votes_against} Against)"
                        else:
                            resolution_outcome = f"Not Approved ({votes_for} For > {votes_against} Against)"
                    else:
                        resolution_outcome = ""
                except:
                    resolution_outcome = ""
                
                proposals.append({
                    "Proposal Proxy Year": meeting_year,
                    "Resolution Outcome": resolution_outcome,
                    "Proposal Text": header_part,
                    "Mgmt. Proposal Category": "",
                    "Vote Results - For": votes_for,
                    "Vote Results - Against": votes_against,
                    "Vote Results - Abstained": votes_abstained,
                    "Vote Results - Broker Non-Votes": votes_broker,
                    "Proposal Vote Results Total": ""
                })
    return proposals, directors

def create_vertical_format_data_proposals(proposals):
    """
    Creates a vertical formatted list of strings for proposals.
    Each proposal's fields are numbered sequentially with a blank row between proposals.
    """
    rows = []
    row_num = 1
    for proposal in proposals:
        rows.append(f"R{row_num}. Proposal Proxy Year: {proposal['Proposal Proxy Year']}")
        row_num += 1
        rows.append(f"R{row_num}. Resolution Outcome: {proposal['Resolution Outcome']}")
        row_num += 1
        rows.append(f"R{row_num}. Proposal Text: \"{proposal['Proposal Text']}\"")
        row_num += 1
        rows.append(f"R{row_num}. Mgmt. Proposal Category: {proposal['Mgmt. Proposal Category']}")
        row_num += 1
        rows.append(f"R{row_num}. Vote Results - For: {proposal['Vote Results - For']}")
        row_num += 1
        rows.append(f"R{row_num}. Vote Results - Against: {proposal['Vote Results - Against']}")
        row_num += 1
        rows.append(f"R{row_num}. Vote Results - Abstained: {proposal['Vote Results - Abstained']}")
        row_num += 1
        rows.append(f"R{row_num}. Vote Results - Broker Non-Votes: {proposal['Vote Results - Broker Non-Votes']}")
        row_num += 1
        rows.append(f"R{row_num}. Proposal Vote Results Total: {proposal['Proposal Vote Results Total']}")
        row_num += 1
        # Add a blank row for clarity
        rows.append("")
        row_num += 1
    return rows

def create_vertical_format_data_directors(directors):
    """
    Creates a vertical formatted list of strings for director election data.
    Each director's fields are numbered sequentially with a blank row between each director.
    """
    rows = []
    row_num = 1
    for director in directors:
        rows.append(f"R{row_num}. Director Election Year: {director['Director Election Year']}")
        row_num += 1
        rows.append(f"R{row_num}. Individual: {director['Individual']}")
        row_num += 1
        rows.append(f"R{row_num}. Director Votes For: {director['Director Votes For']}")
        row_num += 1
        rows.append(f"R{row_num}. Director Votes Against: {director['Director Votes Against'] if director['Director Votes Against'] else '(Leave blank)'}")
        row_num += 1
        rows.append(f"R{row_num}. Director Votes Abstained: {director['Director Votes Abstained'] if director['Director Votes Abstained'] else '(Leave blank)'}")
        row_num += 1
        rows.append(f"R{row_num}. Director Votes Withheld: {director['Director Votes Withheld']}")
        row_num += 1
        rows.append(f"R{row_num}. Director Votes Broker-Non-Votes: {director['Director Votes Broker-Non-Votes']}")
        row_num += 1
        # Add a blank row for clarity
        rows.append("")
        row_num += 1
    return rows

def create_excel_file(proposal_rows, director_rows):
    """
    Creates an Excel file with two sheets:
      - "Proposal sheet" containing vertical proposal data
      - "non-proposal sheet" containing vertical director election data
    """
    # Create DataFrames where each row is a string line
    df_proposals = pd.DataFrame(proposal_rows, columns=["Proposal Data"])
    df_directors = pd.DataFrame(director_rows, columns=["Director Election Data"])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_proposals.to_excel(writer, sheet_name="Proposal sheet", index=False)
        df_directors.to_excel(writer, sheet_name="non-proposal sheet", index=False)
    return output.getvalue()

# --- STREAMLIT APP ---

def main():
    st.title("AGM Result Extractor")
    st.write("Upload the AGM results PDF to extract proposals and director election results.")
    
    uploaded_file = st.file_uploader("Upload AGM PDF", type=["pdf"])
    
    if uploaded_file is not None:
        try:
            # Extract text from the PDF
            pdf_text = extract_text_from_pdf(uploaded_file)
            # Parse the PDF text into proposals and director election data
            proposals, directors = parse_agm_data(pdf_text)
            
            # Create vertical formatted data for each sheet
            proposal_rows = create_vertical_format_data_proposals(proposals)
            director_rows = create_vertical_format_data_directors(directors)
            
            # Create the Excel file with two sheets
            excel_data = create_excel_file(proposal_rows, director_rows)
            
            st.success("Data extracted successfully!")
            st.download_button(
                label="Download Excel File",
                data=excel_data,
                file_name="agm_extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error processing file: {e}")

if __name__ == "__main__":
    main()
