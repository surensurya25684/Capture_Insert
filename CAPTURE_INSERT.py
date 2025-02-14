import streamlit as st
import PyPDF2
import re
import pandas as pd

def extract_text_from_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def parse_nutanix_8k_data(text):
    """
    This function parses the Nutanix 8‑K text to extract:
      • Director election data (from Proposal 1 – Election of Directors)
      • Proposal data (from Proposals 2, 3, and 4)
    
    It assumes that the proposals are listed under Item 5.07.
    """
    # Look for the block starting at "Item 5.07"
    match = re.search(r'Item\s+5\.07\.(.*)', text, re.DOTALL)
    if not match:
        st.error("Could not find the 'Item 5.07' block in the document.")
        return [], []
    block = match.group(1).strip()
    
    # Use a regex to capture proposals.
    # This pattern looks for lines like "1. Proposal 1 – <Title>." and then captures all text until the next proposal.
    proposal_pattern = re.compile(r'(\d+)\.\s+Proposal\s+\d+\s+–\s+(.*?)\.(.*?)(?=\n\d+\.\s+Proposal|\Z)', re.DOTALL)
    proposals_matches = proposal_pattern.findall(block)
    
    proposals_list = []
    directors_list = []
    
    # In this filing, the meeting year is 2024.
    meeting_year = "2024"
    
    for prop in proposals_matches:
        order, title, details = prop
        title = title.strip()
        details = details.strip()
        
        # Check if this is the director election proposal
        if "Election of Directors" in title:
            # Parse the director election table.
            # The table header is expected to be:
            # "Nominee For Against Abstain Broker Non-Votes"
            lines = details.splitlines()
            director_data_started = False
            for line in lines:
                line = line.strip()
                # Look for the header line and then start processing subsequent lines
                if re.search(r'Nominee\s+For\s+Against\s+Abstain\s+Broker\s+Non-Votes', line):
                    director_data_started = True
                    continue
                if director_data_started and line:
                    # Use a regex to capture the director row.
                    m = re.match(r'(.+?)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)', line)
                    if m:
                        name = m.group(1).strip()
                        votes_for = m.group(2).strip()
                        votes_against = m.group(3).strip()
                        votes_abstain = m.group(4).strip()
                        votes_broker = m.group(5).strip()
                        directors_list.append({
                            "Director Election Year": meeting_year,
                            "Individual": name,
                            "Director Votes For": votes_for,
                            "Director Votes Against": votes_against,
                            "Director Votes Abstained": votes_abstain,
                            "Director Votes Withheld": "",  # Not provided in the filing
                            "Director Votes Broker-Non-Votes": votes_broker
                        })
        else:
            # This is a proposal (non‑director)
            # We now try to extract the vote numbers.
            # For proposals 2 and 3, the expected header is "For Against Abstain" (optionally with "Broker Non-Votes").
            # For proposal 4, the header is "One Year Two Years Three Years Abstain Broker Non-Votes".
            numbers = re.findall(r'([\d,]+)', details)
            if "One Year" in details:
                # Proposal 4 – assume: 
                # • Vote Results - For = first number (One Year)
                # • Vote Results - Against = (Leave blank)
                # • Vote Results - Abstained = fourth number (Abstain)
                # • Vote Results - Broker Non-Votes = fifth number
                if len(numbers) >= 5:
                    vote_for = numbers[0]
                    vote_abstain = numbers[3]
                    vote_broker = numbers[4]
                    vote_against = ""
                else:
                    vote_for = vote_abstain = vote_broker = vote_against = ""
            else:
                # Proposals 2 and 3 – assume the first three numbers are For, Against, and Abstain;
                # if a fourth number exists, it is Broker Non-Votes.
                if len(numbers) >= 3:
                    vote_for = numbers[0]
                    vote_against = numbers[1]
                    vote_abstain = numbers[2]
                    vote_broker = numbers[3] if len(numbers) >= 4 else ""
                else:
                    vote_for = vote_against = vote_abstain = vote_broker = ""
            
            # Compute resolution outcome if possible.
            resolution_outcome = ""
            try:
                if vote_for and vote_against:
                    if int(vote_for.replace(',', '')) > int(vote_against.replace(',', '')):
                        resolution_outcome = f"Approved ({vote_for} For > {vote_against} Against)"
                    else:
                        resolution_outcome = f"Not Approved ({vote_for} For > {vote_against} Against)"
            except:
                resolution_outcome = ""
            
            proposals_list.append({
                "Proposal Proxy Year": meeting_year,
                "Resolution Outcome": resolution_outcome,
                "Proposal Text": title,
                "Mgmt. Proposal Category": "",
                "Vote Results - For": vote_for,
                "Vote Results - Against": vote_against,
                "Vote Results - Abstained": vote_abstain,
                "Vote Results - Broker Non-Votes": vote_broker,
                "Proposal Vote Results Total": ""
            })
    return proposals_list, directors_list

def main():
    st.title("Nutanix 8-K Data Extraction Preview")
    
    uploaded_file = st.file_uploader("Upload Nutanix 8-K PDF", type=["pdf"])
    if uploaded_file is not None:
        text = extract_text_from_pdf(uploaded_file)
        proposals, directors = parse_nutanix_8k_data(text)
        
        st.subheader("Director Election Data (non-proposal sheet)")
        if directors:
            df_directors = pd.DataFrame(directors)
            st.table(df_directors)
        else:
            st.write("No director election data found.")
        
        st.subheader("Proposal Data (Proposal sheet)")
        if proposals:
            df_proposals = pd.DataFrame(proposals)
            st.table(df_proposals)
        else:
            st.write("No proposal data found.")

if __name__ == "__main__":
    main()

