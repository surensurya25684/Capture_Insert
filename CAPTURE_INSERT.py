import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.title("AGM Results Extractor")
st.markdown("""
**Instructions:**
1. Upload an 8-K (or similar) PDF containing AGM voting results.
2. The app will extract text, look for "Item 5.07," split out each proposal, and parse director election tables and vote counts.
3. Finally, you can download an Excel file with two sheets: "Directors" and "Proposals."
""")

uploaded_file = st.file_uploader("Upload AGM PDF file", type=["pdf"])

if uploaded_file is not None:
    # Extract text from the PDF
    full_text = ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
    except Exception as e:
        st.error(f"Error reading PDF file: {e}")
        st.stop()

    # Debug: show the first portion of extracted text
    st.expander("Show Extracted PDF Text").write(full_text[:2000])

    # Locate the Item 5.07 section using a flexible regex
    # This attempts to capture everything after "Item 5.07" up until "Item 9.01" or "SIGNATURES" or the end of file.
    section_507 = re.search(
        r"(?s)Item\s+5\.07(.*?)(?=Item\s+9\.01|SIGNATURES|$)",
        full_text,
        re.IGNORECASE
    )

    if not section_507:
        st.error("Could not find the Item 5.07 section in the document.")
        st.stop()

    section_text = section_507.group(1)
    st.expander("Show Item 5.07 Extracted Text").write(section_text[:2000])

    # Split the text into separate proposals by looking for lines like "1." or "2." or "3." etc.
    # We allow optional newlines/spaces before the digit, then a period, then some space.
    # Note: Adjust if your PDF lumps everything on one line.
    proposals = re.split(r"(?:^|\n)\s*\d+\.\s+", section_text)
    proposals = [p.strip() for p in proposals if p.strip()]

    # If you suspect proposals might be all in one chunk, try removing (?:^|\n) so it splits on any "X. " pattern.
    # proposals = re.split(r"\s*\d+\.\s+", section_text)

    st.write(f"**Found {len(proposals)} proposal section(s)**")
    st.expander("Show Proposal Splits").write(proposals)

    # Prepare DataFrames
    director_df = pd.DataFrame(columns=["Nominee", "For Votes", "Withheld Votes", "Broker Non-Votes"])
    proposals_list = []

    for proposal_text in proposals:
        # Separate the first sentence or line as the "title" if possible
        title_match = re.match(r"([^\.]+)\.\s*(.*)", proposal_text, re.DOTALL)
        if title_match:
            proposal_title = title_match.group(1).strip()
            proposal_body = title_match.group(2).strip()
        else:
            # Fallback: use the entire chunk as the body
            proposal_title = ""
            proposal_body = proposal_text

        # Look for director election language
        # Some 8-Ks say "Election of Directors," others might say "The following nominees..."
        # We'll check if the text includes "nominee(s)" or "The following nominees" or "Election of Directors".
        # Adjust this condition to match your files.
        if ("election of directors" in proposal_title.lower()
            or "the following nominees" in proposal_title.lower()
            or "the following nominees" in proposal_body.lower()):
            
            # We'll attempt to find a line that has the header "For", "Withheld", and "Broker Non-Votes"
            lines = proposal_body.splitlines()
            header_index = None
            for i, line in enumerate(lines):
                # If the line contains all three columns, treat it as a header
                if ("For" in line and "Withheld" in line and "Broker Non-Votes" in line):
                    header_index = i
                    break

            # If we found a header, parse subsequent lines
            if header_index is not None:
                for line in lines[header_index+1:]:
                    # Typical line format: "Name 123,456 78,910 11,121"
                    # We'll parse that with a regex capturing 1) name, 2) for votes, 3) withheld, 4) broker non-votes
                    # If your lines have more columns or different formatting, you must adjust.
                    candidate_match = re.match(r"(.+?)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)$", line.strip())
                    if candidate_match:
                        nominee = candidate_match.group(1).strip()
                        for_votes = candidate_match.group(2).replace(',', '')
                        withheld_votes = candidate_match.group(3).replace(',', '')
                        broker_non_votes = candidate_match.group(4).replace(',', '')
                        # Add row to DataFrame
                        director_df.loc[len(director_df)] = [
                            nominee,
                            int(for_votes),
                            int(withheld_votes),
                            int(broker_non_votes)
                        ]
            else:
                # If we didn't find a header line, let's see if the PDF lines up differently
                st.write("**Could not detect director table header** in this proposal:\n", proposal_text[:300])

        else:
            # Parse standard proposals with "For", "Against", "Abstain", "Broker Non-Votes"
            # If your PDF uses different words (e.g., "Abstentions"), adapt accordingly.
            # We'll gather them from the text with simple regex patterns.
            votes = {}
            for label in ["For", "Against", "Abstain", "Abstention", "Abstentions", "Broker Non-Votes"]:
                pattern = rf"{label}\s*:\s*([\d,]+)|{label}\s+([\d,]+)"
                match = re.search(pattern, proposal_body, re.IGNORECASE)
                if match:
                    # This pattern can match in group(1) or group(2) so we check which group is not None
                    number_str = match.group(1) if match.group(1) else match.group(2)
                    votes[label.lower()] = int(number_str.replace(',', ''))
                else:
                    votes[label.lower()] = None

            # We also store the raw text so you know which proposal it came from
            proposals_list.append({
                "Proposal": proposal_title or "Proposal",
                "For": votes.get("for"),
                "Against": votes.get("against"),
                "Abstain": votes.get("abstain") or votes.get("abstention") or votes.get("abstentions"),
                "Broker Non-Votes": votes.get("broker non-votes")
            })

    proposals_df = pd.DataFrame(proposals_list)

    # Write to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not director_df.empty:
            director_df.to_excel(writer, index=False, sheet_name="Directors")
        if not proposals_df.empty:
            proposals_df.to_excel(writer, index=False, sheet_name="Proposals")
    excel_data = output.getvalue()

    # Provide download button
    st.download_button(
        "Download Excel File",
        data=excel_data,
        file_name="AGM_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Show dataframes in the UI
    st.subheader("Director Election Results")
    st.dataframe(director_df)

    st.subheader("Other Proposals")
    st.dataframe(proposals_df)
