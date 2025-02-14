import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.title("AGM Results Extractor")
st.markdown("""
Upload an Annual General Meeting PDF file. The app will:
• Extract the "Item 5.07" section (which includes proposals and director election votes).
• Parse director election details and other proposals.
• Output an Excel file with two sheets: one for director results and one for other proposals.
""")

uploaded_file = st.file_uploader("Upload AGM PDF file", type=["pdf"])

if uploaded_file is not None:
    # Extract text from the PDF
    full_text = ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
    except Exception as e:
        st.error("Error reading PDF file: " + str(e))
        st.stop()

    # Debug: show the first portion of extracted text
    st.expander("Show Extracted Full Text").write(full_text[:1000])

    # Extract the Item 5.07 section.
    # This regex looks for text starting at "Item 5.07" until one of the common stopping points.
    section_match = re.search(r"Item\s+5\.07(.*?)(Item\s+9\.01|SIGNATURES|$)", full_text, re.DOTALL | re.IGNORECASE)
    if section_match:
        section_text = section_match.group(1)
        st.expander("Show Extracted Item 5.07 Section").write(section_text[:1000])
    else:
        st.error("Could not find the Item 5.07 section in the document.")
        st.stop()

    # Split the section into proposals using numbered points (e.g., "1. ", "2. ", etc.)
    proposals = re.split(r"\n?\s*\d+\.\s+", section_text)
    proposals = [p.strip() for p in proposals if p.strip()]
    st.write("Found", len(proposals), "proposal(s).")
    st.expander("Show Proposals List").write(proposals)

    # Prepare DataFrames
    director_df = pd.DataFrame(columns=["Nominee", "For Votes", "Withheld Votes", "Broker Non-Votes"])
    proposals_list = []

    # Process each proposal
    for proposal in proposals:
        # Attempt to split the proposal into a title and content.
        title_match = re.match(r"([^\.]+)\.\s*(.*)", proposal, re.DOTALL)
        if title_match:
            title = title_match.group(1).strip()
            content = title_match.group(2).strip()
        else:
            # Fallback: Use the first line as the title
            lines = proposal.splitlines()
            title = lines[0].strip() if lines else ""
            content = proposal

        # If the proposal relates to the "Election of Directors", process candidate rows.
        if "Election of Directors" in title:
            lines = content.splitlines()
            header_index = None
            # Find a header line containing "Nominee", "For" and "Withheld"
            for i, line in enumerate(lines):
                if "Nominee" in line and "For" in line and "Withheld" in line:
                    header_index = i
                    break

            if header_index is not None:
                for line in lines[header_index+1:]:
                    candidate_match = re.match(r"(.+?)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)", line)
                    if candidate_match:
                        nominee = candidate_match.group(1).strip()
                        for_votes = candidate_match.group(2).replace(',', '')
                        withheld_votes = candidate_match.group(3).replace(',', '')
                        broker_non_votes = candidate_match.group(4).replace(',', '')
                        try:
                            director_df.loc[len(director_df)] = [nominee, int(for_votes), int(withheld_votes), int(broker_non_votes)]
                        except Exception as e:
                            st.write("Error processing candidate row:", line, str(e))
            else:
                st.write("Election of Directors header not found in proposal:", title)
        else:
            # For other proposals, extract vote counts (For, Against, Abstentions, Broker Non-Votes)
            votes = {}
            for label in ["For", "Against", "Abstentions", "Broker Non-Votes"]:
                pattern = label + r"\s+([\d,]+)"
                m = re.search(pattern, content)
                if m:
                    votes[label] = int(m.group(1).replace(',', ''))
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

    st.download_button("Download Excel File", data=excel_data,
                       file_name="AGM_results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Director Election Results")
    st.dataframe(director_df)

    st.subheader("Other Proposals")
    st.dataframe(proposals_df)
