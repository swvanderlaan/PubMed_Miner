#!/usr/bin/env python3

# Import necessary libraries
import os
import re
import argparse
import logging
import subprocess
import importlib
from datetime import datetime
from Bio import Entrez
from collections import defaultdict
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Change log:
# * v1.0.6, 2024-11-15: Added top 10 journals plot. Fixed issue with JID extraction. Fixed issue with open access extraction. Added more logging. Added --debug flag. 
# * v1.0.5, 2024-11-15: Fixed an issue where the logo was not properly referenced.
# * v1.0.4, 2024-11-15: Added logo to Word document header.
# * v1.0.3, 2024-11-15: Expanded Word-document information.
# * v1.0.2, 2024-11-15: Added retry logic for PubMed API, better logging for aliases, results directory customization, improved input validation, enhanced plotting, and bar annotations.
# * v1.0.1, 2024-11-15: Added alias handling for authors, improved deduplication of YearCount and PubTypeYearCount in Excel, added Authors column.
# * v1.0.0, 2024-11-14: Initial version. Added --year flag for filtering by year range, adjusted tables and figures by author. Stratified tables and figures for each author in DEFAULT_NAMES. Summarized results in Word and Excel files. Added support for "Access Type" in Publications sheet.

# Version and License Information
VERSION_NAME = 'PubMed Miner'
VERSION = '1.0.6'
VERSION_DATE = '2024-11-15'
COPYRIGHT = 'Copyright 1979-2024. Sander W. van der Laan | s.w.vanderlaan [at] gmail [dot] com | https://vanderlaanand.science.'
COPYRIGHT_TEXT = '''
Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License

By exercising the Licensed Rights (defined below), You accept and agree to be bound by the terms and conditions of this Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License ("Public License"). To the extent this Public License may be interpreted as a contract, You are granted the Licensed Rights in consideration of Your acceptance of these terms and conditions, and the Licensor grants You such rights in consideration of benefits the Licensor receives from making the Licensed Material available under these terms and conditions.

Full license text available at https://creativecommons.org/licenses/by-nc-nd/4.0/legalcode.

This software is provided "as is" without warranties or guarantees of any kind.
'''

# Alias mapping for handling multiple author names
ALIAS_MAPPING = {
    "Schiffelers R": "Schiffelers RM",
    "Hofer I": "Hoefer IE",
    "Hoefer I.E.": "Hoefer IE",
    "Hoefer I.": "Hoefer IE",
    "Imo E. Hofer": "Hoefer IE",
    "Imo Hofer": "Hoefer IE",
    "Imo E. Hoefer": "Hoefer IE",
    "Imo E Hoefer": "Hoefer IE",
    "Imo Hoefer": "Hoefer IE",
    "Schoneveld AH": "Schoneveld AH",
    "Schoneveld A.H.": "Schoneveld AH",
    "Schoneveld A.": "Schoneveld AH",
    "Arjen H. Schoneveld": "Schoneveld AH",
    "Arjen H Schoneveld": "Schoneveld AH",
    "Arjen Schoneveld": "Schoneveld AH",
    "Hester M. den Ruijter": "den Ruijter HM",
    "Hester M den Ruijter": "den Ruijter HM",
    "Hester den Ruijter": "den Ruijter HM",
    "van der Laan S": "van der Laan SW",
    "van der Laan S.W.": "van der Laan SW",
    "van der Laan Sander W.": "van der Laan SW",
    "Sander W. van der Laan": "van der Laan SW",
    "van Solinge WW": "van Solinge W",
    "van Solinge W.W.": "van Solinge W",
    # Add other aliases as needed
}

# Set some defaults
DEFAULT_ORGANIZATION = "University Medical Center Utrecht"
DEFAULT_NAMES = ["van der Laan SW", "Pasterkamp G", "Mokry M", "Schiffelers RM", 
"van Solinge W", "Haitjema S", "den Ruijter HM", 
"Hoefer IE",
"Schoneveld AH",
"Vader P"]
DEFAULT_DEPARTMENTS = ["Central Diagnostic Laboratory"]

# Setup Logging
def setup_logger(results_dir, verbose, debug):
    """
    Setup the logger to log to a file and console.
    """
    date_str = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(results_dir, f"{date_str}_CDL_UMCU_Publications.log")
    os.makedirs(results_dir, exist_ok=True)
    
    logger = logging.getLogger('pubmed_miner')
    # Determine logging level
    if debug:
        level = logging.DEBUG
    elif verbose:
        level = logging.INFO
    else:
        level = logging.WARNING
    logger.setLevel(level)
    
    fh = logging.FileHandler(log_file)
    fh.setLevel(logging.DEBUG)  # File should always capture all messages
    ch = logging.StreamHandler()
    ch.setLevel(level)  # Console output based on verbosity/debug
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger

# Check and install necessary packages
def check_install_package(package_name, logger):
    """
    Check if a package is installed and install it if not.
    """
    try:
        importlib.import_module(package_name)
    except ImportError:
        logger.info(f'{package_name} is not installed. Installing it now...')
        subprocess.check_call(['pip', 'install', package_name])

# Retry logic for API calls
def fetch_with_retry(db, term, retries=3, backoff=2):
    """
    Fetch PubMed data with retry logic to handle rate limits or network issues.
    """
    for attempt in range(retries):
        try:
            handle = Entrez.esearch(db=db, term=term, retmax=100)
            return Entrez.read(handle)
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(backoff ** attempt)
                continue
            raise e

# Parse year or year range
def parse_year_range(year_range_str):
    """
    Parse year or year range string and return start and end year.
    """
    if '-' in year_range_str:
        start_year, end_year = map(int, year_range_str.split('-'))
    else:
        start_year = end_year = int(year_range_str)
    return start_year, end_year

# Parse command-line arguments
def parse_arguments():
    """
    Parse command-line arguments.
    """
    parser = argparse.ArgumentParser(description=f"""
{VERSION_NAME} v{VERSION} ({VERSION_DATE})
Retrieve PubMed publications for a list of authors.""",
        epilog=f"""
This script retrieves PubMed publications for a list of authors and departments from UMC Utrecht.
It then analyzes the publication data and saves the results to a Word document and an Excel file.

Required arguments:
    -e, --email <email-address>  Email address for PubMed API access.

Optional arguments:
    -n, --names <names>          List of author names to search for. 
                                    Default: {DEFAULT_NAMES} with these aliases: {ALIAS_MAPPING}.
    -dep, --departments <depts>    List of departments to search for. Default: {DEFAULT_DEPARTMENTS}.
    -org, --organization <org>   Organization name for filtering results. Default: {DEFAULT_ORGANIZATION}.
    -y, --year <year>            Filter publications by year or year range (e.g., 2024 or 2017-2024).
    -o, --output-file <file>     Output base name for the Word and Excel files. Default: date_CDL_UMCU_Publications.
    -r, --results-dir <dir>      Directory to save results. Default: results.
    -v, --verbose                Enable verbose output.
    -d, --debug                  Enable debug output.
    -V, --version                Show program's version number and exit.

Example:
    python pubmed_miner.py --email <email-address> --year 2017-2024 --verbose
+ {VERSION_NAME} v{VERSION}. {COPYRIGHT} +
{COPYRIGHT_TEXT}""",
        formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("-e", "--email", required=True, help="Email for PubMed API access.")
    parser.add_argument("-n", "--names", nargs='+', default=DEFAULT_NAMES, help="List of author names to search for.")
    parser.add_argument("-dep", "--departments", nargs='+', default=DEFAULT_DEPARTMENTS, help="List of departments to search for.")
    parser.add_argument("-org", "--organization", default=DEFAULT_ORGANIZATION, help="Organization name for filtering results.")
    parser.add_argument("-y", "--year", help="Filter publications by year or year range (e.g., 2024 or 2017-2024).")
    parser.add_argument("-o", "--output-file", default="CDL_UMCU_Publications", help="Output base name for the Word and Excel files.")
    parser.add_argument("-r", "--results-dir", default="results", help="Directory to save results.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output.")
    parser.add_argument("-d", "--debug", action="store_true", help="Enable debug output.")
    parser.add_argument('-V', '--version', action='version', version=f'{VERSION_NAME} {VERSION} ({VERSION_DATE})')
    return parser.parse_args()

# Fetch publication detailss
def fetch_publication_details(pubmed_ids, logger, main_author, start_year=None, end_year=None):
    """
    Fetch detailed information for each PubMed ID, filter by year, and identify preprints.
    """
    canonical_author = ALIAS_MAPPING.get(main_author, main_author)
    publications = []
    preprints = []

    for pub_id in pubmed_ids:
        record = None

        # Retry logic for Entrez.efetch
        for attempt in range(3):
            try:
                handle = Entrez.efetch(db="pubmed", id=pub_id, rettype="medline", retmode="text")
                record = handle.read()
                handle.close()
                break  # Exit loop on success
            except Exception as e:
                if attempt < 2:  # Retry for first two attempts
                    time.sleep(2 ** attempt)  # Exponential backoff
                    logger.warning(f"Retrying PubMed fetch for ID {pub_id} (attempt {attempt + 2})...")
                else:  # Final failure
                    logger.error(f"Failed to fetch PubMed details for ID {pub_id}: {e}")
                    continue

        if not record:
            continue  # Skip this ID if all retries failed

        # Extract publication details
        authors = re.findall(r"AU  - (.+)", record)
        authors = [ALIAS_MAPPING.get(author, author) for author in authors]  # Replace aliases with canonical names
        title = re.search(r"TI  - (.+)", record).group(1) if re.search(r"TI  - (.+)", record) else "No title found"
        journal_abbr = re.search(r"TA  - (.+)", record).group(1) if re.search(r"TA  - (.+)", record) else "No journal abbreviation found"
        # journal_id = re.search(r"JID  - (.+)", record).group(1) if re.search(r"JID  - (.+)", record) else "No journal ID"
        # Extract Journal ID with debugging print
        jid_match = re.search(r"JID\s*-\s*(\d+)", record)
        if jid_match:
            journal_id = jid_match.group(1)
            # debug
            # logger.debug(f"Extracted Journal ID: {journal_id}")
        else:
            journal_id = "No journal ID"
            # debug
            # logger.debug(f"No Journal ID found in record.")
        pub_date = re.search(r"DP  - (.+)", record).group(1)[:4] if re.search(r"DP  - (.+)", record) else "Unknown Year"
        doi_match = re.search(r"AID - (10\..+?)(?: \[doi\])", record)
        doi_link = f"https://doi.org/{doi_match.group(1)}" if doi_match else "No DOI found"
        pub_type = re.findall(r"PT  - (.+)", record)
        source = re.search(r"SO  - (.+)", record).group(1) if re.search(r"SO  - (.+)", record) else "No source found"

        # Extract UOF (Update of) field
        uof_match = re.search(r"UOF\s*-\s*(.+)", record)
        uof = uof_match.group(1) if uof_match else "No UOF"

        # Determine publication type based on PT and SO
        publication_type = "Other"
        if "Review" in pub_type or "Review" in source:
            publication_type = "Review"
        elif "Book" in pub_type or "Book" in source:
            publication_type = "Book"
        elif "Journal Article" in pub_type or "Journal Article" in source:
            publication_type = "Journal Article"

        # Determine access type based on the presence of a PMC ID
        pmc_match = re.search(r"PMC\s+-\s+PMC\d+", record)
        if pmc_match:
            access_type = "open access"
            # debug
            # logger.debug(f"Detected PMC ID: {pmc_match.group()}")
        else:
            access_type = "closed access"
            # debug
            # logger.debug(f"No PMC ID detected; setting access type to 'closed access'.")

        # Skip errata or corrections
        if "ERRATUM" in title.upper() or "AUTHOR CORRECTION" in title.upper():
            logger.info(f"Skipping 'erratum' or 'author correction' for [{title}]")
            continue

        # Filter by year if specified
        if start_year and end_year:
            if not (start_year <= int(pub_date) <= end_year):
                logger.info(f"Skipping [{title}] as it falls outside the year range.")
                continue
        
        # Check if the UOF references preprint servers and extract DOI
        uof_doi = None
        if any(preprint_source in uof for preprint_source in ["medRxiv", "bioRxiv", "arXiv"]):
            logger.info(f"PMID {pub_id} references a preprint ({uof}). Adding to preprints.")
            uof_parts = uof.split()
            for part in uof_parts:
                if part.startswith("10."):
                    uof_doi = f"https://doi.org/{part}"
                    break
            uof_citation = " ".join(uof_parts[:uof_parts.index(part)]) if uof_doi else uof
            preprints.append((pub_id, canonical_author if canonical_author in authors else f"{canonical_author} et al.", pub_date, journal_abbr, journal_id, title, uof_doi, uof_citation, publication_type))

        # Separate preprints
        if "Preprint" in pub_type:
            preprints.append((pub_id, canonical_author if canonical_author in authors else f"{canonical_author} et al.", pub_date, journal_abbr, journal_id, title, doi_link, source, publication_type, access_type))
        else:
            publications.append((pub_id, canonical_author if canonical_author in authors else f"{canonical_author} et al.", pub_date, journal_abbr, journal_id, title, doi_link, source, publication_type, access_type))
        # debug
        logger.debug(f"Found MEDLINE record for PubMed ID {pub_id} matching criteria (within year range).")
        logger.debug(f"Authors: {authors}")
        logger.debug(f"Title: {title}")
        logger.debug(f"Journal: {journal_abbr}")
        logger.debug(f"Journal ID: {journal_id}")
        logger.debug(f"Publication Date: {pub_date}")
        logger.debug(f"DOI: {doi_link}")
        logger.debug(f"Publication Type: {publication_type}")
        logger.debug(f"Citation: {source}")
        logger.debug(f"Access Type: {access_type}")
        logger.debug(f"UOF: {uof}")
        # debug - this produces a lot of output
        logger.debug(f"This was the full record:\n{record}")

    return publications, preprints

# Analyze publication data
def analyze_publications(publications_data, main_author):
    """
    Analyze the publication data and return counts for authors, years, and journals.
    Also return counts of publications by author per year and by year per journal.
    """
    canonical_author = ALIAS_MAPPING.get(main_author, main_author)
    author_count = defaultdict(int)
    year_count = defaultdict(int)
    author_year_count = defaultdict(lambda: defaultdict(int))
    year_journal_count = defaultdict(lambda: defaultdict(int))
    pub_type_count = defaultdict(lambda: defaultdict(int))

    for pub in publications_data:
        (
            pub_id,
            author,
            year,
            journal_abbr,
            journal_id,
            title,
            doi_link,
            citation,
            pub_type,
            access_type,
        ) = pub  # Adjusted unpacking to match the expanded structure
        
        if canonical_author in author:
            author_count[canonical_author] += 1
            author_year_count[canonical_author][year] += 1
            year_count[year] += 1
            year_journal_count[year][journal_abbr] += 1
            pub_type_count[pub_type][year] += 1
    
    return author_count, year_count, author_year_count, year_journal_count, pub_type_count

# Write results to Excel
def write_to_excel(author_data, output_file, results_dir, logger):
    """
    Write the combined results for all authors into six sheets of an Excel file.
    """
    logger.info("Writing results to Excel file.")
    output_file = f"{datetime.now().strftime('%Y%m%d')}_{output_file}"  # Add date to the filename
    writer = pd.ExcelWriter(os.path.join(results_dir, f"{output_file}.xlsx"), engine='xlsxwriter')

    # Combine all publications into a single DataFrame
    logger.info("> Combining all publications into a single DataFrame.")
    publications_data = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = ALIAS_MAPPING.get(main_author, main_author)
        for pub in publications:
            publications_data.append(list(pub) + [canonical_author])
    logger.debug(f"Publications data: {publications_data[:3]}")  # Log first 3 publications
    publications_df = pd.DataFrame(
        publications_data,
        columns=["PubMed ID", "Author", "Year", "Journal", "JID", "Title", "DOI Link", "Citation", "Publication Type", "Access Type", "Main Author"]
    )

    # Deduplicate based on PubMed ID and combine authors for duplicates in publications
    publications_df = (
        publications_df.groupby("PubMed ID")
        .agg({
            "Author": lambda x: ", ".join(sorted(set(x))),  # Combine full unique author names
            "Year": "first",
            "Journal": "first",
            "JID": "first",
            "Title": "first",
            "DOI Link": "first",
            "Citation": "first",
            "Publication Type": "first",
            "Access Type": "first",
        })
        .reset_index()
    )

    publications_df["Year"] = pd.to_numeric(publications_df["Year"], errors='coerce')
    publications_df.to_excel(writer, sheet_name="Publications", index=False)

    # Combine all preprints into a single DataFrame with deduplication
    logger.info("> Combining all preprints into a single DataFrame with deduplication.")
    preprints_data = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = ALIAS_MAPPING.get(main_author, main_author)
        for preprint in preprints:
            preprints_data.append(list(preprint) + [canonical_author])
    logger.debug(f"Preprints data: {preprints_data[:3]}")  # Log first 3 preprints
    preprints_df = pd.DataFrame(
        preprints_data,
        columns=["PubMed ID", "Author", "Year", "Journal", "JID", "Title", "DOI Link", "Citation", "Publication Type", "Access Type", "Main Author"]
    )

    # Deduplicate based on PubMed ID and combine authors for duplicates in preprints
    preprints_df = (
        preprints_df.groupby("PubMed ID")
        .agg({
            "Author": lambda x: ", ".join(sorted(set(x))),  # Combine full unique author names
            "Year": "first",
            "Journal": "first",
            "JID": "first",
            "Title": "first",
            "DOI Link": "first",
            "Citation": "first",
            "Publication Type": "first",
        })
        .reset_index()
    )

    preprints_df["Year"] = pd.to_numeric(preprints_df["Year"], errors='coerce')
    preprints_df.to_excel(writer, sheet_name="Preprints", index=False)

    # Combine author counts
    logger.info("> Combining author counts.")
    author_counts = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = ALIAS_MAPPING.get(main_author, main_author)
        for author, count in author_count.items():
            author_counts.append([author, count, canonical_author])
    author_counts_df = pd.DataFrame(author_counts, columns=["Author", "Number of Publications", "Main Author"])
    author_counts_df.drop(columns=["Main Author"], inplace=True)
    author_counts_df.to_excel(writer, sheet_name="AuthorCount", index=False)

    # Combine and deduplicate year counts
    logger.info("> Combining and deduplicating year counts.")
    year_counts = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = ALIAS_MAPPING.get(main_author, main_author)
        for year, count in year_count.items():
            year_counts.append([year, count, canonical_author])
    year_counts_df = pd.DataFrame(year_counts, columns=["Year", "Number of Publications", "Authors"])

    year_counts_df = (
        year_counts_df.groupby("Year")
        .agg({
            "Number of Publications": "sum",
            "Authors": lambda x: ", ".join(sorted(set(x)))
        })
        .reset_index()
    )

    year_counts_df["Year"] = pd.to_numeric(year_counts_df["Year"], errors='coerce')
    year_counts_df.to_excel(writer, sheet_name="YearCount", index=False)

    # Combine and deduplicate publication type by year counts
    logger.info("> Combining and deduplicating publication type by year counts.")
    pub_type_year_counts = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = ALIAS_MAPPING.get(main_author, main_author)
        for pub_type, years in pub_type_count.items():
            for year, count in years.items():
                pub_type_year_counts.append([pub_type, year, count, canonical_author])
    pub_type_year_df = pd.DataFrame(pub_type_year_counts, columns=["Publication Type", "Year", "Number of Publications", "Authors"])

    pub_type_year_df = (
        pub_type_year_df.groupby(["Publication Type", "Year"])
        .agg({
            "Number of Publications": "sum",
            "Authors": lambda x: ", ".join(sorted(set(x)))
        })
        .reset_index()
    )

    pub_type_year_df["Year"] = pd.to_numeric(pub_type_year_df["Year"], errors='coerce')
    pub_type_year_df.to_excel(writer, sheet_name="PubTypeYearCount", index=False)

    # Save the Excel file
    logger.info(f"Excel file saved to [{os.path.join(results_dir, f'{output_file}.xlsx')}].\n")
    writer.close()

# Word document table creation
def add_table_to_doc(doc, data, headers, title):
    """
    Add a table to a Word document.
    Safeguard against empty or malformed rows.
    data: List of rows, each row is a list of items.
    headers: List of column headers.
    title: Title for the table.
    """
    doc.add_heading(title, level=2)
    table = doc.add_table(rows=1, cols=len(headers))
    for idx, header in enumerate(headers):
        table.rows[0].cells[idx].text = header

    for row in data:
        # Skip empty or malformed rows
        if not row or len(row) != len(headers):
            continue
        row_cells = table.add_row().cells
        for idx, item in enumerate(row):
            row_cells[idx].text = str(item)
    doc.add_paragraph()

# Write results to Word
def write_to_word(author_data, output_file, results_dir, logger, args):
    """
    Write the combined results for all authors to a Word document.
    """
    logger.info("Writing results to Word document.")
    output_file = f"{datetime.now().strftime('%Y%m%d')}_{output_file}"  # Add date to the filename
    query_date = datetime.now().strftime('%Y-%m-%d')
    query_quarter = (datetime.now().month - 1) // 3 + 1
    document = Document()

    # Add a header for the logo
    section = document.sections[0]
    header = section.header
    header_paragraph = header.paragraphs[0]

    # Add the logo to the header
    logo_path = os.path.join("images/FullLogo_Transparent.png")  # Ensure path is correct
    try:
        run = header_paragraph.add_run()
        run.add_picture(logo_path, width=Inches(1.5))  # Adjust size as needed
        header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    except Exception as e:
        logger.error(f"Could not add logo to Word document: {e}")

    # Add main content to the document
    document.add_heading(f"Publications for {query_quarter} at the Central Diagnostics Laboratory", level=1)
    document.add_paragraph(f"This document summarizes the publications linked to the Central Diagnostics Laboratory (CDL) at the University Medical Center Utrecht (UMCU).")
    document.add_paragraph()
    document.add_paragraph(f"The following settings are used:")
    document.add_paragraph(f"* Query date: {query_date}.")
    document.add_paragraph(f"* Authors: {', '.join(author_data.keys())}.")
    document.add_paragraph(f"* Aliases used: {', '.join(ALIAS_MAPPING.keys())}.")
    document.add_paragraph(f"* Year range: {args.year}." if args.year else "* No year filter used.")
    document.add_paragraph(f"* Department(s): {', '.join(DEFAULT_DEPARTMENTS)}.")
    document.add_paragraph(f"* Organization: {DEFAULT_ORGANIZATION}.")
    document.add_paragraph()
    document.add_paragraph(f"Results saved on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.")
    document.add_paragraph(f"Log file saved to {os.path.join(results_dir, f'{query_date}_CDL_UMCU_Publications.log')}.")
    document.add_paragraph()
    document.add_paragraph(f"{VERSION_NAME} v{VERSION} ({VERSION_DATE}).")
    document.add_paragraph(f"{COPYRIGHT}")
    document.add_paragraph()
    document.add_paragraph(f"GitHub repository: https://github.com/swvanderlaan/PubMed_Miner. \nAny issues or requests? Create one here: https://github.com/swvanderlaan/PubMed_Miner/issues.")

    # Add results for each canonical author
    logger.info(f"> Adding results for {len(author_data)} author(s).")
    for canonical_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        document.add_heading(f"Author: {canonical_author}", level=1)

        # Main publications and preprints for this author
        if publications:
            logger.info(f"Adding {len(publications)} publications for {canonical_author}.")
            
            # Debugging: Check structure of publications
            logger.debug(f"Publications for {canonical_author}: {publications[:3]}")  # Log first 3 publications
            
            # Ensure valid publications align with expected structure
            valid_publications = [
                pub for pub in publications if len(pub) >= 10
            ]  # Adjust for expected columns
            
            if not valid_publications:
                logger.warning(f"No valid publications found for {canonical_author}. Check data structure.")
            else:
                add_table_to_doc(
                    document,
                    valid_publications,
                    headers=["PubMed ID", "Author", "Year", "Journal", "JID", "Title", "DOI Link", "Citation", "Publication Type", "Access Type"],
                    title="Main Publications",
                )
        else:
            logger.warning(f"No publications found for {canonical_author}.")

        if preprints:
            logger.info(f"Adding {len(preprints)} preprints for {canonical_author}.")
            
            # Debugging: Check structure of preprints
            logger.debug(f"Raw Preprints for {canonical_author}: {preprints}")
            
            # Adjust for expected structure
            valid_preprints = [
                preprint for preprint in preprints if len(preprint) >= 9
            ]
            
            if not valid_preprints:
                logger.warning(f"No valid preprints found for {canonical_author}. Check data structure.")
            else:
                logger.info(f"Adding {len(valid_preprints)} valid preprints for {canonical_author}.")
                add_table_to_doc(
                    document,
                    valid_preprints,
                    headers=["PubMed ID", "Author", "Year", "Journal", "JID", "Title", "DOI Link", "Citation", "Publication Type"],
                    title="Preprints",
                )
        else:
            logger.warning(f"No preprints found for {canonical_author}.")

        # Summary tables for this author
        add_table_to_doc(
            document,
            [(author, count) for author, count in author_count.items()],
            headers=["Author", "Number of Publications"],
            title="Number of Publications per Author",
        )
        add_table_to_doc(
            document,
            [(year, count) for year, count in year_count.items()],
            headers=["Year", "Number of Publications"],
            title="Number of Publications per Year",
        )
        add_table_to_doc(
            document,
            [(year, journal, count) for year, journals in year_journal_count.items() for journal, count in journals.items()],
            headers=["Year", "Journal", "Number of Publications"],
            title="Number of Publications per Year per Journal",
        )
        add_table_to_doc(
            document,
            [(ptype, year, count) for ptype, years in pub_type_count.items() for year, count in years.items()],
            headers=["Publication Type", "Year", "Number of Publications"],
            title="Publication Type by Year",
        )

    # Add combined graphs after individual sections
    logger.info("> Adding graphs for all authors.")
    document.add_heading("Graphs", level=1)
    for plot_name in [
        f"{datetime.now().strftime('%Y%m%d')}_publications_per_author.png",
        f"{datetime.now().strftime('%Y%m%d')}_publications_per_year.png",
        f"{datetime.now().strftime('%Y%m%d')}_publications_per_author_and_year.png",
        f"{datetime.now().strftime('%Y%m%d')}_top10_journals_grouped.png",
        f"{datetime.now().strftime('%Y%m%d')}_total_publications_preprints_by_author.png",
        f"{datetime.now().strftime('%Y%m%d')}_publications_by_access_type.png",
    ]:
        plot_path = os.path.join(results_dir, plot_name)
        if os.path.exists(plot_path):
            try:
                document.add_picture(plot_path, width=Inches(6))
                document.add_paragraph()
            except Exception as e:
                logger.error(f"Error adding plot {plot_path}: {e}")
        else:
            logger.warning(f"Plot file not found: {plot_path}")

    # Save the document
    output_path = os.path.join(results_dir, f"{output_file}.docx")
    document.save(output_path)
    logger.info(f"Word document saved to [{output_path}].")

# Plot the results
def plot_results(author_data, results_dir, logger):
    """
    Plot the results for each author and save the plots.
    """
    date_str = datetime.now().strftime('%Y%m%d')

    # Consistent color mapping for authors
    canonical_authors = list(author_data.keys())
    colors = plt.colormaps["tab10"](np.linspace(0, 1, len(canonical_authors)))
    color_map = {canonical_author: colors[idx] for idx, canonical_author in enumerate(canonical_authors)}

    # Access type color mapping
    access_color_map = {"open access": "green", "closed access": "red"}

    # Plot publications per author
    logger.info("> Plotting publications per author.")
    fig, ax = plt.subplots()
    for canonical_author, (_, _, author_count, _, _, _, _) in author_data.items():
        ax.bar(canonical_author, author_count[canonical_author], color=color_map[canonical_author], label=canonical_author)
    ax.set_xlabel("Authors")
    ax.set_ylabel("Number of Publications")
    ax.set_title("Publications per Author")
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, f"{date_str}_publications_per_author.png"))

    # Total number of publications and preprints grouped by author and year (two panels)
    logger.info("> Plotting total publications and preprints grouped by author and year (two panels).")
    fig, axes = plt.subplots(1, 2, figsize=(16, 7), sharey=True)
    access_types = ["open access", "closed access"]

    # Process data for plotting
    for idx, access_type in enumerate(access_types):
        ax = axes[idx]
        for canonical_author, (publications, preprints, _, _, _, _, _) in author_data.items():
            yearly_totals = defaultdict(int)
            for pub in publications + preprints:
                pub_access_type = pub[-1]  # Access type is the last field
                year = int(pub[2])  # Year is the third field
                if pub_access_type == access_type:
                    yearly_totals[year] += 1
            years = sorted(yearly_totals.keys())
            counts = [yearly_totals[year] for year in years]
            ax.bar(
                years,
                counts,
                label=canonical_author,
                color=color_map[canonical_author],
                alpha=0.8,
            )
        ax.set_title(f"Total Publications ({access_type.capitalize()})")
        ax.set_xlabel("Year")
        ax.set_xticks(years)
        ax.set_xticklabels(years, rotation=45)
        if idx == 0:
            ax.set_ylabel("Total Publications and Preprints")
        ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")

    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, f"{date_str}_total_publications_preprints_by_author.png"))

    # Bar plot for publications per year (stacked by author)
    fig, ax = plt.subplots()
    width = 0.8 / len(author_data)

    # Gather all unique years across authors
    all_years = sorted(set(year for _, (_, _, _, year_count, _, _, _) in author_data.items() for year in year_count))

    # Plot each author's publications per year with unique color
    logger.info("> Plotting publications per year.")
    for idx, (canonical_author, (_, _, _, year_count, _, _, _)) in enumerate(author_data.items()):
        counts = [year_count.get(year, 0) for year in all_years]
        ax.bar(
            [y + idx * width for y in range(len(all_years))],
            counts,
            width=width,
            color=color_map[canonical_author],
            label=canonical_author,
        )

    # Set x-ticks to display years in the middle of each bar cluster
    ax.set_xticks([y + (width * len(author_data)) / 2 for y in range(len(all_years))])
    ax.set_xticklabels(all_years, rotation=45)
    ax.set_xlabel("Year")
    ax.set_ylabel("Number of Publications")
    ax.set_title("Publications per Year (Colored by Author)")
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, f"{date_str}_publications_per_year.png"))

    # Plot publications per author per year (stacked bar plot)
    logger.info("> Plotting publications per author and year.")
    fig, ax = plt.subplots()
    width = 0.2
    x = sorted({year for _, (_, _, _, year_count, _, _, _) in author_data.items() for year in year_count})
    x_indices = np.arange(len(x))
    for idx, (canonical_author, (_, _, _, year_count, author_year_count, _, _)) in enumerate(author_data.items()):
        counts = [author_year_count[canonical_author].get(year, 0) for year in x]
        ax.bar(x_indices + idx * width, counts, width, color=color_map[canonical_author], label=canonical_author)
    ax.set_xticks(x_indices + width / 2 * (len(canonical_authors) - 1))
    ax.set_xticklabels(x, rotation=45)
    ax.set_xlabel("Year")
    ax.set_ylabel("Number of Publications")
    ax.set_title("Publications per Author and Year")
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, f"{date_str}_publications_per_author_and_year.png"))

    # Top 10 journals by number of publications (grouped by year)
    logger.info("> Plotting top 10 journals.")
    journal_counts = defaultdict(lambda: defaultdict(int))
    for _, (_, _, _, _, _, year_journal_count, _) in author_data.items():
        for year, journals in year_journal_count.items():
            for journal, count in journals.items():
                journal_counts[journal][year] += count

    # Flatten journal counts and sort by total
    total_journal_counts = {journal: sum(counts.values()) for journal, counts in journal_counts.items()}
    sorted_journals = sorted(total_journal_counts.items(), key=lambda x: x[1], reverse=True)
    top_10_journals = [journal for journal, _ in sorted_journals[:10]]  # Keep only top 10

    # Prepare data for plotting
    all_years = sorted({year for journal in journal_counts for year in journal_counts[journal]})
    grouped_counts = defaultdict(lambda: defaultdict(int))
    for journal in top_10_journals:
        for year in all_years:
            grouped_counts[journal][year] = journal_counts[journal].get(year, 0)

    # Plot data
    fig, ax = plt.subplots(figsize=(12, 7))
    bar_width = 0.8 / len(top_10_journals)  # Adjust bar width based on the number of journals
    colors = plt.cm.tab10(np.linspace(0, 1, len(top_10_journals)))  # Generate a color palette

    # Create x positions for each year
    x_positions = np.arange(len(all_years))

    # Plot each journal as a separate group
    for idx, journal in enumerate(top_10_journals):
        counts = [grouped_counts[journal][year] for year in all_years]
        offset = idx * bar_width  # Offset each journal's bars
        ax.bar(
            x_positions + offset,
            counts,
            width=bar_width,
            label=journal,
            color=colors[idx],
        )

    # Customize x-axis and labels
    ax.set_xticks(x_positions + bar_width * len(top_10_journals) / 2)  # Center x-ticks
    ax.set_xticklabels(all_years, rotation=45)
    ax.set_xlabel("Year")
    ax.set_ylabel("Number of Publications")
    ax.set_title("Top 10 Journals (Grouped by Year)")
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")

    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, f"{date_str}_top10_journals_grouped.png"))

    # Publications grouped by access type and year (not stacked)
    logger.info("> Plotting publications by access type and year (grouped).")
    fig, ax = plt.subplots(figsize=(12, 7))
    access_counts = {"open access": defaultdict(int), "closed access": defaultdict(int)}

    # Process data for grouped bar plot
    for _, (publications, _, _, _, _, _, _) in author_data.items():
        for pub in publications:
            access_type = pub[-1]
            if access_type in access_counts:
                access_counts[access_type][int(pub[2])] += 1

    all_years = sorted(set(year for totals in access_counts.values() for year in totals.keys()))
    x_indices = np.arange(len(all_years))
    width = 0.35
    for idx, (access_type, yearly_counts) in enumerate(access_counts.items()):
        counts = [yearly_counts[year] for year in all_years]
        ax.bar(
            x_indices + idx * width,
            counts,
            width=width,
            label=access_type,
            color=access_color_map[access_type],
            alpha=0.8,
        )
    ax.set_xticks(x_indices + width / 2)
    ax.set_xticklabels(all_years, rotation=45)
    ax.set_xlabel("Year")
    ax.set_ylabel("Number of Publications")
    ax.set_title("Publications by Access Type and Year (Grouped)")
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, f"{date_str}_publications_by_access_type.png"))

# Main function
def main():
    args = parse_arguments()
    results_dir = "results"
    os.makedirs(results_dir, exist_ok=True)
    today = datetime.now().strftime('%Y%m%d')
    output_file = os.path.join(results_dir, f"{today}_{args.output_file}")
    logger = setup_logger(results_dir, args.verbose, args.debug)

    start_year, end_year = None, None
    if args.year:
        start_year, end_year = parse_year_range(args.year)

    # Ensure all required packages are installed
    for package in ['Bio', 'docx', 'matplotlib', 'numpy', 'pandas']:
        check_install_package(package, logger)

    Entrez.email = args.email

    # Print some information
    logger.info(f"Running {VERSION_NAME} v{VERSION} ({VERSION_DATE})\n")
    logger.info(f"Settings:")
    logger.info(f"> Search parameters given:")
    logger.info(f"  - authors: {args.names}")
    logger.info(f"  - department(s): {args.departments}")
    logger.info(f"  - organization: ['{args.organization}']")
    logger.info(f"  - filtering by year (range) [{args.year}]" if args.year else "No year filter used.")
    logger.info(f"  - output file(s): [{output_file}]")
    logger.info(f"> PubMed email used: {args.email}.")
    logger.info(f"> Debug mode: {'On' if args.debug else 'Off'}.")
    logger.info(f"> Verbose mode: {'On' if args.verbose else 'Off'}.\n")

    author_data = {}
    logger.info(f"Querying PubMed for publications and preprints.\n")

    for main_author in args.names:
        canonical_author = ALIAS_MAPPING.get(main_author, main_author)
        all_pubmed_ids = []
        for department in args.departments:
            search_query = f"{main_author} {department} {args.organization}"
            logger.info(f"Querying PubMed for [ {search_query} ]\n")
            try:
                record = fetch_with_retry(db="pubmed", term=search_query)
                all_pubmed_ids.extend(record["IdList"])
            except Exception as e:
                logger.error(f"Failed to fetch PubMed IDs for query [{search_query}]: {e}")
                continue

        publications, preprints = fetch_publication_details(all_pubmed_ids, logger, canonical_author, start_year, end_year)
        author_count, year_count, author_year_count, year_journal_count, pub_type_count = analyze_publications(publications, canonical_author)
        author_data[canonical_author] = (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count)

        logger.info(f"Found {len(publications)} publications and {len(preprints)} preprints for [{canonical_author}].\n")

    logger.info(f"Done. Summarizing and saving results.\n")
    logger.info(f"Saving plots to [{results_dir}].")
    plot_results(author_data, results_dir, logger)

    # Save results to Word and Excel
    write_to_word(author_data, args.output_file, results_dir, logger, args)
    write_to_excel(author_data, args.output_file, results_dir, logger)

    logger.info(f"Saved the following results:")
    logger.info(f"> Data summarized and saved to {os.path.join(results_dir, f'{today}_{args.output_file}.docx')}.")
    logger.info(f"> Excel concatenated and saved to {os.path.join(results_dir, f'{today}_{args.output_file}.xlsx')}.")
    logger.info(f"> Plots saved to {results_dir}/.")
    logger.info(f"> Log file saved to {os.path.join(results_dir, f'{today}_CDL_UMCU_Publications.log')}.\n")
    logger.info(f"Thank you for using {VERSION_NAME} v{VERSION} ({VERSION_DATE}).")
    logger.info(f"{COPYRIGHT}")
    logger.info(f"Script completed successfully on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.")

if __name__ == "__main__":
    main()
