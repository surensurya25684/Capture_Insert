import streamlit as st
import pandas as pd
import re
import fitz  # PyMuPDF
import io

# Function to extract text from PDF
def extract_text_from_pdf(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = "\n".join([page.get_text("text") for page in doc])
    return text

# Function to extract AGM Proposals
def parse_agm_proposals(text):
    proposals = []
    
    proposal_pattern = re.compile(
        r"Proposal (\d+): (.*?)\nFor -- ([\d,]+) Against -- ([\d,]+) Abstain -- ([\d,]+) BrokerNon-Votes -- (\d+)", re.DOTALL
    )

    matches = proposal_pattern.findall(text)
    
    for match in matches:
        proposal_number, proposal_text, votes_for, votes_against, votes_abstain, votes_broker_non_votes = match
        resolution_outcome = "Approved" if int(votes_for.replace(',', '')) > int(votes_against.replace(',', '')) else "Rejected"

        proposals.append([
            ("Proposal Proxy Year", "2024"),
            ("Resolution Outcome", f"{resolution_outcome} ({votes_for} > {votes_against})"),
            ("Proposal Text", f'"{proposal_text.strip()
