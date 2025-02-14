import streamlit as st
import pandas as pd
import re
from io import BytesIO
import PyPDF2

st.title("AGM Results Extractor")

st.markdown("""
This tool allows you to upload an Annual General Meeting (AGM) results PDF.
It will extract:
- **Proposals** (with proxy year, vote results, and resolution outcome)
- **Director Election Results** (with fixed election year 2024 and vote details)

The results will be output in an Excel file with two sheets:
- **Proposal** sheet for proposals  
- **non-proposal** sheet for director elections  
""")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload AGM Results PDF", type="pdf")

if uploaded_file:
    # --- Extract Text from PDF ---
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        full_text = ""
        for page in pdf_reader.pages:
            full_text += page.extract_text() + "\n"
    except Exception as e:
        st.error(f"Error reading PDF file: {e}")
        st.stop()
    
    st.subheader("Extracted PDF Text (for reference)")
    st.text_area("PDF Text", full_text, height=200)

    # --- Extract Proposals ---
    proposals = []
    # This regex captures:
    # - The proposal text (which may span multiple lines) up to the vote results line
    # - The vote numbers: For, Against, Abstain, Broker Non-Votes
    proposal_pattern = re.compile(
        r"Proposal\s*\d+:\s*(.*?)\s*\n\s*For\s*--\s*([\d,]+)\s*Against\s*--\s*([\d,]+)\s*Abstain\s*--\s*([\d,]+)\s*Broker\s*Non-Votes\s*--\s*([\d,]+)",
        re.DOTALL | re.IGNORECASE
    )
    
    for match in proposal_pattern.finditer(full_text):
        # Clean up the proposal text (remove newlines inside the text)
        prop_text = " ".join(match.group(1).strip().split())
        votes_for_raw = match.group(2).strip()
        votes_against_raw = match.group(3).strip()
        votes_abstain_raw = match.group(4).strip()
        votes_broker_raw = match.group(5).strip()
        
        # Remove commas and convert vote counts to integers for comparison
        try:
            votes_for = int(votes_for_raw.replace(",", ""))
            votes_against = int(votes_against_raw.replace(",", ""))
        except:
            votes_for = 0
            votes_against = 0

        # Determine resolution outcome
        resolution = "Approved" if votes_for > votes_against else "Rejected"
        resolution_outcome = f"{resolution} ({votes_for_raw}>{votes_against_raw})"
        
        # Try to extract the proposal proxy year (looking for a 4-digit number like 2024)
        year_match = re.search(r'\b(20\d{2})\b', prop_text)
        proposal_year = year_match.group(1) if year_match else ""
        
        proposals.append({
            "Proposal Proxy Year": proposal_year,
            "Resolution Outcome": resolution_outcome,
            "Proposal Text": prop_text,
            "Mgmt. Proposal Category": "",
            "Vote Results - For": votes_for_raw,
            "Vote Results - Against": votes_against_raw,
            "Vote Results - Abstained": match.group(4).strip(),
            "Vote Results - Withheld": "",  # Not provided
            "Vote Results - Broker Non-Votes": votes_broker_raw,
            "Proposal Vote Results Total": ""
        })
    
    st.subheader("Extracted Proposals")
    if proposals:
        st.dataframe(pd.DataFrame(proposals))
    else:
        st.warning("No proposals found using the current pattern.")
    
    # --- Extract Director Election Results ---
    directors = []
    # Look for the section with director election results using the "Nominee" header.
    # The assumption is that the section header contains "Nominee" followed by column headings.
    director_section = ""
    director_section_match = re.search(r"Nominee\s+For\s+.*Broker Non-Votes\s*\n(.*)", full_text, re.DOTALL | re.IGNORECASE)
    if director_section_match:
        director_section = director_section_match.group(1)
    else:
        # Fallback: if "Nominee" is found anywhere, split from that point onward.
        if "Nominee" in full_text:
            director_section = full_text.split("Nominee", 1)[1]
    
    director_lines = director_section.splitlines()
    # Remove potential header line (if it contains 'For')
    if director_lines and "For" in director_lines[0]:
        director_lines = director_lines[1:]
    
    # Process each non-empty line assuming each line contains:
    # Name, Votes For, Votes Withheld, Broker Non-Votes (separated by multiple spaces)
    for line in director_lines:
        line = line.strip()
        if not line:
            continue
        parts = re.split(r'\s{2,}', line)
        if len(parts) >= 4:
            name = parts[0].strip()
            votes_for = parts[1].strip()
            votes_withheld = parts[2].strip()
            votes_broker = parts[3].strip()
            directors.append({
                "Director Election Year": "2024",
                "Individual": name,
                "Director Votes For": votes_for,
                "Director Votes Against": "",   # Not available
                "Director Votes Abstained": "",   # Not available
                "Director Votes Withheld": votes_withheld,
                "Director Votes Broker-Non-Votes": votes_broker
            })
    
    st.subheader("Extracted Director Election Results")
    if directors:
        st.dataframe(pd.DataFrame(directors))
    else:
        st.warning("No director election results found using the current pattern.")
    
    # --- Create Excel with Two Sheets ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        if proposals:
            df_proposals = pd.DataFrame(proposals)
            # Write proposals data to the "Proposal" sheet
            df_proposals.to_excel(writer, sheet_name="Proposal", index=False)
        if directors:
            df_directors = pd.DataFrame(directors)
            # Write director election data to the "non-proposal" sheet
            df_directors.to_excel(writer, sheet_name="non-proposal", index=False)
        writer.save()
    
    processed_data = output.getvalue()
    
    st.download_button(
        label="Download Excel file",
        data=processed_data,
        file_name="AGM_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
