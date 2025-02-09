import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

# Function to extract AGM data from PDF dynamically
def extract_pdf_data(pdf_file):
    agm_details = {}
    director_elections = []
    proposals = []

    # Read the PDF
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    # Extract AGM date
    match = re.search(r"Date of report.*?: (\w+ \d{1,2}, \d{4})", full_text)
    agm_details["Annual Meeting Date"] = match.group(1) if match else "Unknown"

    # Extract signing date (last date in the document)
    date_matches = re.findall(r"(\w+ \d{1,2}, \d{4})", full_text)
    agm_signing_date = date_matches[-1] if date_matches else "Unknown"

    # Extract total votes outstanding
    match = re.search(r"(\d{2,},\d{3,}) shares present", full_text)
    agm_details["Default Total Votes Outstanding"] = match.group(1).replace(",", "") if match else "0"

    # Extract director elections dynamically
    director_section = re.search(r"I\..*?directors.*?\n\n(.*?)\n\nII\.", full_text, re.DOTALL)
    if director_section:
        director_lines = director_section.group(1).strip().split("\n")
        for line in director_lines:
            parts = re.findall(r"([\w\s.-]+?)\s+(\d{2,},\d{3,})\s+(\d{2,},\d{3,})\s+(\d{2,},\d{3,})", line)
            if parts:
                name, votes_for, votes_against, votes_abstained = parts[0]
                director_elections.append({
                    "Individual": name.strip(),
                    "Director Votes For": int(votes_for.replace(",", "")),
                    "Director Votes Against": int(votes_against.replace(",", "")),
                    "Director Votes Abstained": int(votes_abstained.replace(",", "")),
                })

    # Extract proposals dynamically
    proposal_section = re.search(r"III\..*?vote.*?\n\n(.*?)\n\nIV\.", full_text, re.DOTALL)
    if proposal_section:
        proposal_lines = proposal_section.group(1).strip().split("\n")
        for line in proposal_lines:
            parts = re.findall(r"([\w\s.-]+?)\s+(\d{2,},\d{3,})\s+(\d{2,},\d{3,})\s+(\d{2,},\d{3,})?", line)
            if parts:
                text, votes_for, votes_against, votes_abstained = parts[0]
                resolution_outcome = "Approved" if int(votes_for.replace(",", "")) > int(votes_against.replace(",", "")) else "Rejected"
                proposals.append({
                    "Proposal Proxy Year": "2024",
                    "Resolution Outcome": resolution_outcome,
                    "Proposal Text": text.strip(),
                    "Mgmt Proposal Category": "",
                    "Vote Results - For": int(votes_for.replace(",", "")),
                    "Vote Results - Against": int(votes_against.replace(",", "")),
                    "Vote Results - Abstained": int(votes_abstained.replace(",", "")) if votes_abstained else 0,
                    "Vote Results - Withheld": "",
                    "Vote Results - Broker Non-Votes": "",
                    "Proposal Vote Results Total": "",
                })

    return agm_details, director_elections, proposals, agm_signing_date

# Streamlit UI
st.title("📊 AGM Data Extractor & Formatter")

st.write("Upload your **VR Template (Excel)** and **AGM Results (PDF)** to format the data.")

# File Uploaders
uploaded_template = st.file_uploader("Upload VR Template (Excel)", type=["xlsx"])
uploaded_pdf = st.file_uploader("Upload AGM Result PDF", type=["pdf"])

if uploaded_template and uploaded_pdf:
    with st.spinner("Processing... ⏳"):
        # Load the template
        template_xl = pd.ExcelFile(uploaded_template)
        scalar_df = template_xl.parse("Scalar datapoint")
        proposals_df = template_xl.parse("Proposals")
        non_proposals_df = template_xl.parse("Non-Proposals")

        # Extract data from the uploaded PDF
        agm_details, director_elections, proposals, agm_signing_date = extract_pdf_data(uploaded_pdf)

        # Update Scalar DataFrame
        for key, value in agm_details.items():
            scalar_df.loc[scalar_df["Datapoint name"] == key, "Data Entry"] = value

        # Set fixed values
        scalar_df.loc[2, "Data Entry"] = "Votes Added"
        scalar_df.loc[3, "Data Entry"] = "Management"
        
        # Clear specified row
        scalar_df.iloc[4, 1] = ""

        # Update the signing date
        scalar_df.loc[scalar_df["Datapoint name"].str.contains("Date", case=False, na=False), "Data Entry"] = agm_signing_date

        # Insert proposals data **vertically** in "Data Entry" column
        for i, proposal in proposals_df.iterrows():
            if i < len(proposals):
                proposals_df.at[i, "Data Entry"] = "\n".join(f"{k} - {v}" for k, v in proposals[i].items())

        # Insert director elections data **vertically** in "Data Entry" column
        for i, director in non_proposals_df.iterrows():
            if i < len(director_elections):
                non_proposals_df.at[i, "Data Entry"] = "\n".join(f"{k} - {v}" for k, v in director_elections[i].items())

        # Save results to an Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            scalar_df.to_excel(writer, sheet_name="Scalar datapoint", index=False)
            proposals_df.to_excel(writer, sheet_name="Proposals", index=False)
            non_proposals_df.to_excel(writer, sheet_name="Non-Proposals", index=False)

        output.seek(0)

        st.success("✅ Data successfully processed! Download your formatted file below.")

        # Provide a download button
        st.download_button(
            label="📥 Download Formatted Excel",
            data=output,
            file_name="formatted_agm_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
