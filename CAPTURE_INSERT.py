import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.title("AGM Results Extractor")

st.markdown("""
This app allows you to upload an Annual General Meeting PDF document.
It will extract the proposals (including director election votes) from Item 5.07 and then output an Excel file.
""")

uploaded_file = st.file_uploader("Upload AGM PDF file", type=["pdf"])

if uploaded_file is not None:
    # Extract text from the PDF file
    full_text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # Try to locate the Item 5.07 section (from "Item 5.07" to a marker such as "Item 9.01" or "SIGNATURES")
    section_match = re.search(r"Item\s+5\.07(.*?)(Item\s+9\.01|SIGNATURES)", full_text, re.DOTALL | re.IGNORECASE)
    if section_match:
        section_text = section_match.group(1)
    else:
        st.error("Could not find the Item 5.07 section in the uploaded document.")
        st.stop()

    # Split the section into proposals based on numbering (e.g., 1. 2. etc.)
    proposals = re.split(r"\n\s*\d+\.\s*", section_text)
    # Remove empty strings and strip extra spaces
    proposals = [p.strip() for p in proposals if p.strip()]

    # DataFrames for director election results and for other proposals
    director_df = pd.DataFrame(columns=["Nominee", "For Votes", "Withheld Votes", "Broker Non-Votes"])
    proposals_list = []

    # Process each proposal in the Item 5.07 section
    for proposal in proposals:
        # Get the proposal title (assumed to be before the first period) and its content
        title_match = re.match(r"([^\.]+)\.\s*(.*)", proposal, re.DOTALL)
        if title_match:
            title = title_match.group(1).strip()
            content = title_match.group(2).strip()
        else:
            title = ""
            content = proposal

        # If it's the Election of Directors proposal, extract candidate rows
        if "Election of Directors" in title:
            lines = content.splitlines()
            # Look for the header line containing "Nominee", "For", "Withheld" (assuming these columns appear)
            header_index = None
            for i, line in enumerate(lines):
                if "Nominee" in line and "For" in line and "Withheld" in line:
                    header_index = i
                    break
            if header_index is not None:
                # Process each subsequent line as a candidate row
                for line in lines[header_index+1:]:
                    # Use regex to capture candidate name and three numbers (For, Withheld, Broker Non-Votes)
                    candidate_match = re.match(r"(.+?)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", line)
                    if candidate_match:
                        nominee = candidate_match.group(1).strip()
                        for_votes = candidate_match.group(2).replace(',', '')
                        withheld_votes = candidate_match.group(3).replace(',', '')
                        broker_non_votes = candidate_match.group(4).replace(',', '')
                        director_df = director_df.append({
                            "Nominee": nominee,
                            "For Votes": int(for_votes),
                            "Withheld Votes": int(withheld_votes),
                            "Broker Non-Votes": int(broker_non_votes)
                        }, ignore_index=True)
        else:
            # For other proposals, extract vote counts. We expect lines like "For <number>", "Against <number>", etc.
            votes = {}
            vote_labels = ["For", "Against", "Abstentions", "Broker Non-Votes"]
            for label in vote_labels:
                pattern = label + r"\s+([\d,]+)"
                match = re.search(pattern, content)
                if match:
                    votes[label] = int(match.group(1).replace(',', ''))
                else:
                    votes[label] = None
            proposals_list.append({
                "Proposal": title,
                "For": votes["For"],
                "Against": votes["Against"],
                "Abstentions": votes["Abstentions"],
                "Broker Non-Votes": votes["Broker Non-Votes"]
            })

    proposals_df = pd.DataFrame(proposals_list)

    # Create an Excel file with two sheets: one for Directors and one for Proposals
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not director_df.empty:
            director_df.to_excel(writer, index=False, sheet_name="Directors")
        if not proposals_df.empty:
            proposals_df.to_excel(writer, index=False, sheet_name="Proposals")
    excel_data = output.getvalue()

    st.download_button("Download Excel File", data=excel_data, file_name="AGM_results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Director Election Results")
    st.dataframe(director_df)

    st.subheader("Other Proposals")
    st.dataframe(proposals_df)
