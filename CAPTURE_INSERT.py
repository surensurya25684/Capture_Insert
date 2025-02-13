import streamlit as st
import pandas as pd
import re
import fitz  # PyMuPDF
import io

# Function to extract Section 5.07 from PDF
def extract_section_507(pdf_file):
    try:
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        extracted_text = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text")

            if "5.07" in text:  # Look for Section 5.07 in the text
                extracted_text.append(text)

        if not extracted_text:
            st.warning("⚠️ Section 5.07 not found in the PDF. Please check the document.")
            return ""

        return "\n".join(extracted_text)

    except Exception as e:
        st.error(f"Error extracting text from PDF: {e}")
        return ""

# Function to extract Director Election Results
def parse_director_elections(text):
    directors = []
    
    director_pattern = re.compile(r"([\w\s]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)")

    matches = director_pattern.findall(text)

    if not matches:
        st.warning("⚠️ No Director Election data found in Section 5.07.")

    for match in matches:
        director_name, votes_for, votes_against, votes_abstain, votes_broker_non_votes = match

        directors.append([
            ("Director Election Year", "2024"),
            ("Individual", director_name.strip()),
            ("Director Votes For", votes_for),
            ("Director Votes Against", votes_against),
            ("Director Votes Abstained", votes_abstain),
            ("Director Votes Broker-Non-Votes", votes_broker_non_votes),
            ("---", "---")  # Separator row
        ])
    
    return directors

# Function to extract AGM Proposals
def parse_agm_proposals(text):
    proposals = []

    # Proposal pattern to capture proposal number, vote results
    proposal_pattern = re.compile(
        r"Proposal No\.\s*(\d+)\s*–\s*(.*?)\nFor\s*([\d,]+)\s*Against\s*([\d,]+)\s*Abstain\s*([\d,]+)\s*Broker Non-Votes\s*([\d,]+)",
        re.DOTALL
    )

    matches = proposal_pattern.findall(text)

    if not matches:
        st.warning("⚠️ No AGM Proposals found in Section 5.07.")

    for match in matches:
        proposal_number, proposal_text, votes_for, votes_against, votes_abstain, votes_broker_non_votes = match
        resolution_outcome = "Approved" if int(votes_for.replace(',', '')) > int(votes_against.replace(',', '')) else "Rejected"

        proposals.append([
            ("Proposal Proxy Year", "2024"),
            ("Resolution Outcome", f"{resolution_outcome} ({votes_for} > {votes_against})"),
            ("Proposal Text", f'"{proposal_text.strip()}"'),
            ("Mgmt Proposal Category", ""),
            ("Vote Results - For", votes_for),
            ("Vote Results - Against", votes_against),
            ("Vote Results - Abstained", votes_abstain),
            ("Vote Results - Withheld", ""),
            ("Vote Results - Broker Non-Votes", votes_broker_non_votes),
            ("Proposal Vote Results Total", ""),
            ("---", "---")  # Separator row
        ])
    
    return proposals

# Function to save extracted data into an Excel file using openpyxl
def save_to_excel(proposals, directors):
    try:
        proposal_data = [item for sublist in proposals for item in sublist]
        director_data = [item for sublist in directors for item in sublist]

        proposal_df = pd.DataFrame(proposal_data, columns=["Field", "Value"])
        director_df = pd.DataFrame(director_data, columns=["Field", "Value"])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:  # Use openpyxl instead of xlsxwriter
            proposal_df.to_excel(writer, sheet_name="Proposal Sheet", index=False)
            director_df.to_excel(writer, sheet_name="Non-Proposal Sheet", index=False)
        
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error generating Excel file: {e}")
        return None

# Streamlit UI
st.title("📄 AGM Proposal & Director Election Data Extractor (Section 5.07 Only)")
st.write("Upload an AGM results **PDF** document, extract **Section 5.07**, and save structured data in **Excel format**.")

uploaded_file = st.file_uploader("Upload AGM Result File (PDF)", type=["pdf"])

if uploaded_file is not None:
    st.info("📄 Extracting Section 5.07 from the PDF...")

    pdf_text = extract_section_507(uploaded_file)

    # Display Extracted Text for Debugging
    st.subheader("Extracted Section 5.07 Preview")
    st.text_area("Extracted Text:", pdf_text[:5000], height=300)  # Show first 5000 characters

    # Extract and process data
    proposals = parse_agm_proposals(pdf_text)
    directors = parse_director_elections(pdf_text)

    if not proposals and not directors:
        st.warning("⚠️ No proposals or director election data found in Section 5.07.")
    else:
        st.success("✅ Data extracted successfully! You can now download the Excel file.")

        # Generate Excel file
        excel_file = save_to_excel(proposals, directors)

        if excel_file:
            st.download_button(
                label="📥 Download Extracted Data (Excel)",
                data=excel_file,
                file_name="AGM_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Failed to generate Excel file.")
