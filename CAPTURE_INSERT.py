import streamlit as st
import re
import pandas as pd
import xlsxwriter
from io import BytesIO
from PyPDF2 import PdfReader

st.title("AGM Result Extractor")
st.write("Upload your AGM results PDF file to extract proposals and director election data.")

def extract_text_from_pdf(pdf_file):
    try:
        pdf_reader = PdfReader(pdf_file)
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return ""
    text = ""
    for i, page in enumerate(pdf_reader.pages):
        try:
            page_text = page.extract_text()
        except Exception as e:
            st.warning(f"Error extracting text from page {i+1}: {e}")
            continue
        if page_text:
            text += page_text + "\n"
    return text

def get_item507_section(text):
    # Look for "Item 5.07." and capture until the next "Item" with a number followed by a period.
    match = re.search(r'(Item\s+5\.07\..*?)(?=Item\s+\d+\.)', text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1)
    else:
        st.warning("Could not isolate Item 5.07 section; using entire document for parsing.")
        return text

def parse_directors(section_text):
    directors = []
    # Updated regex: allow optional leading number, and allow a colon or period after "Election of Directors"
    proposal1_pattern = r'(?:\d+\.\s*)?Proposal\s+1\s*[-–—]\s*Election of Directors[.:]?(.*?)(?=(?:\d+\.\s*)?Proposal\s+\d+\s*[-–—]|$)'
    match = re.search(proposal1_pattern, section_text, re.DOTALL | re.IGNORECASE)
    if match:
        proposal1_content = match.group(1).strip()
        st.write("Debug: Proposal 1 content snippet:", proposal1_content[:500])
        # Look for the director table header
        table_match = re.search(r'Nominee\s+For\s+Against\s+Abstain\s+Broker Non[-\s]?Votes(.*)', proposal1_content, re.DOTALL | re.IGNORECASE)
        if table_match:
            table_text = table_match.group(1)
            lines = table_text.splitlines()
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # Split by two or more spaces
                parts = re.split(r'\s{2,}', line)
                if len(parts) >= 5:
                    name = parts[0]
                    try:
                        votes_for = int(parts[1].replace(',', ''))
                    except:
                        votes_for = ""
                    try:
                        votes_against = int(parts[2].replace(',', ''))
                    except:
                        votes_against = ""
                    try:
                        votes_abstain = int(parts[3].replace(',', ''))
                    except:
                        votes_abstain = ""
                    try:
                        votes_broker = int(parts[4].replace(',', ''))
                    except:
                        votes_broker = ""
                    
                    director = {
                        "Director Election Year": "2024",
                        "Individual": name,
                        "Director Votes For": votes_for,
                        "Director Votes Against": votes_against,
                        "Director Votes Abstained": votes_abstain,
                        "Director Votes Withheld": "",
                        "Director Votes Broker-Non-Votes": votes_broker
                    }
                    directors.append(director)
        else:
            st.warning("Director table header not found in Proposal 1.")
    else:
        st.warning("Proposal 1 (Election of Directors) not found.")
    return directors

def parse_proposals(section_text):
    proposals = []
    # Regex for Proposals 2, 3, and 4 with optional leading numbers.
    proposal_pattern = r'(?:\d+\.\s*)?Proposal\s+([2-4])\s*[-–—]\s*(.*?)(?=(?:\d+\.\s*)?Proposal\s+[2-4]\s*[-–—]|$)'
    proposal_blocks = re.findall(proposal_pattern, section_text, re.DOTALL | re.IGNORECASE)
    
    if not proposal_blocks:
        st.warning("No proposals (2, 3, 4) found using the expected pattern.")
    
    for proposal_number, content in proposal_blocks:
        lines = [line.strip() for line in content.splitlines() if line.strip()]
        if not lines:
            continue
        # First line is the proposal title; the rest is narrative and vote info.
        proposal_title = lines[0]
        narrative_lines = []
        vote_header = ""
        vote_numbers_line = ""
        for i, line in enumerate(lines[1:], start=1):
            # Look for a vote header that contains "For" and either "Against" or "One Year"
            if re.search(r'\bFor\b', line, re.IGNORECASE) and (re.search(r'\bAgainst\b', line, re.IGNORECASE) or re.search(r'\bOne Year\b', line, re.IGNORECASE)):
                vote_header = line
                for j in range(i+1, len(lines)):
                    if re.search(r'\d', lines[j]):
                        vote_numbers_line = lines[j]
                        break
                break
            else:
                narrative_lines.append(line)
        narrative_text = " ".join([proposal_title] + narrative_lines)
        year_match = re.search(r'\b(20\d{2})\b', narrative_text)
        proxy_year = year_match.group(1) if year_match else "2024"
        
        vote_numbers = []
        if vote_numbers_line:
            vote_numbers = re.findall(r'[\d,]+', vote_numbers_line)
            vote_numbers = [int(num.replace(',', '')) for num in vote_numbers]
        
        vote_results_for = ""
        vote_results_against = ""
        vote_results_abstained = ""
        vote_results_broker = ""
        
        if vote_header:
            if re.search(r'\bAgainst\b', vote_header, re.IGNORECASE):
                if len(vote_numbers) >= 3:
                    vote_results_for = vote_numbers[0]
                    vote_results_against = vote_numbers[1]
                    vote_results_abstained = vote_numbers[2]
                    if len(vote_numbers) >= 4:
                        vote_results_broker = vote_numbers[3]
                try:
                    if int(vote_results_for) > int(vote_results_against):
                        resolution = f"Approved ({vote_results_for} > {vote_results_against})"
                    else:
                        resolution = f"Not Approved ({vote_results_for} <= {vote_results_against})"
                except:
                    resolution = ""
            elif re.search(r'\bOne Year\b', vote_header, re.IGNORECASE):
                if len(vote_numbers) >= 5:
                    vote_results_for = vote_numbers[0]
                    vote_results_abstained = vote_numbers[3]
                    vote_results_broker = vote_numbers[4]
                resolution = "Approved"
            else:
                resolution = ""
        else:
            resolution = ""
        
        proposal = {
            "Proposal Proxy Year": proxy_year,
            "Resolution Outcome": resolution,
            "Proposal Text": narrative_text,
            "Mgmt. Proposal Category": "",
            "Vote Results - For": vote_results_for,
            "Vote Results - Against": vote_results_against,
            "Vote Results - Abstained": vote_results_abstained,
            "Vote Results - Withheld": "",
            "Vote Results - Broker Non-Votes": vote_results_broker,
            "Proposal Vote Results Total": ""
        }
        proposals.append(proposal)
    return proposals

def format_proposals_for_excel(proposals):
    rows = []
    for proposal in proposals:
        rows.append(["Proposal Proxy Year:", proposal.get("Proposal Proxy Year", "")])
        rows.append(["Resolution Outcome:", proposal.get("Resolution Outcome", "")])
        rows.append(["Proposal Text:", proposal.get("Proposal Text", "")])
        rows.append(["Mgmt. Proposal Category:", proposal.get("Mgmt. Proposal Category", "")])
        rows.append(["Vote Results - For:", proposal.get("Vote Results - For", "")])
        rows.append(["Vote Results - Against:", proposal.get("Vote Results - Against", "")])
        rows.append(["Vote Results - Abstained:", proposal.get("Vote Results - Abstained", "")])
        rows.append(["Vote Results - Withheld:", proposal.get("Vote Results - Withheld", "")])
        rows.append(["Vote Results - Broker Non-Votes:", proposal.get("Vote Results - Broker Non-Votes", "")])
        rows.append(["Proposal Vote Results Total:", proposal.get("Proposal Vote Results Total", "")])
        rows.append([])  # blank row as separator
    return rows

def format_directors_for_excel(directors):
    rows = []
    for director in directors:
        rows.append(["Director Election Year:", director.get("Director Election Year", "")])
        rows.append(["Individual:", director.get("Individual", "")])
        rows.append(["Director Votes For:", director.get("Director Votes For", "")])
        rows.append(["Director Votes Against:", director.get("Director Votes Against", "")])
        rows.append(["Director Votes Abstained:", director.get("Director Votes Abstained", "")])
        rows.append(["Director Votes Withheld:", director.get("Director Votes Withheld", "")])
        rows.append(["Director Votes Broker-Non-Votes:", director.get("Director Votes Broker-Non-Votes", "")])
        rows.append([])  # blank row as separator
    return rows

uploaded_file = st.file_uploader("Upload AGM PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Extracting text from PDF..."):
        pdf_text = extract_text_from_pdf(uploaded_file)
    if not pdf_text:
        st.error("No text extracted from PDF.")
    else:
        st.success("PDF text extraction complete!")
        # Uncomment below to inspect the extracted text:
        # st.text_area("Extracted PDF Text", pdf_text, height=300)
    
    # Isolate the Item 5.07 section
    section_text = get_item507_section(pdf_text)
    
    # Parse director election data (Proposal 1) and proposals 2-4
    directors = parse_directors(section_text)
    proposals = parse_proposals(section_text)
    
    st.header("Proposals Preview")
    if proposals:
        df_proposals = pd.DataFrame(proposals)
        st.dataframe(df_proposals)
    else:
        st.warning("No proposals found in the document.")
    
    st.header("Director Elections Preview")
    if directors:
        df_directors = pd.DataFrame(directors)
        st.dataframe(df_directors)
    else:
        st.warning("No director election data found in the document.")
    
    # Format data for Excel (vertical layout)
    proposals_rows = format_proposals_for_excel(proposals)
    directors_rows = format_directors_for_excel(directors)
    
    # Write the Excel file with two sheets using xlsxwriter
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    proposal_sheet = workbook.add_worksheet("Proposal sheet")
    director_sheet = workbook.add_worksheet("Non-proposal sheet")
    
    row_idx = 0
    for row in proposals_rows:
        col_idx = 0
        for cell in row:
            proposal_sheet.write(row_idx, col_idx, cell)
            col_idx += 1
        row_idx += 1

    row_idx = 0
    for row in directors_rows:
        col_idx = 0
        for cell in row:
            director_sheet.write(row_idx, col_idx, cell)
            col_idx += 1
        row_idx += 1

    workbook.close()
    processed_data = output.getvalue()
    
    st.download_button(
        label="Download Excel File",
        data=processed_data,
        file_name="AGM_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
