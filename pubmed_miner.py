#!/usr/bin/env python3

# Import necessary libraries
import os
import re
import argparse
import logging
import subprocess
import importlib
import time
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
# * v1.1.0, 2024-11-18: Fixed an issue where not all the aliases for --names, --departments and --organization were properly queried in conjunction with --organization. Added an option to include ORCID in the author alias list. Fixed issue where the moving average plot might not handle edge years (with fewer than moving_avg_window data points) gracefully.
# * v1.0.10, 2024-11-15: Fixed an issue with consistency of filenaming. Added moving average per author per year to barplot.
# * v1.0.9, 2024-11-15: Fixed an issue with the dimensions of the preprint and publication tables.
# * v1.0.8, 2024-11-15: Fixed issue with aliases for departments. Also edit option to include more departments to search for. Fixed handling authors. Fixed output of preprint citation. 
# * v1.0.7, 2024-11-15: Fixed an issue where the plot for total_publications_preprints_by_author was not displaying the years correctly.
# * v1.0.6, 2024-11-15: Added top 10 journals plot. Fixed issue with JID extraction. Fixed issue with open access extraction. Added more logging. Added --debug flag. 
# * v1.0.5, 2024-11-15: Fixed an issue where the logo was not properly referenced.
# * v1.0.4, 2024-11-15: Added logo to Word document header.
# * v1.0.3, 2024-11-15: Expanded Word-document information.
# * v1.0.2, 2024-11-15: Added retry logic for PubMed API, better logging for aliases, results directory customization, improved input validation, enhanced plotting, and bar annotations.
# * v1.0.1, 2024-11-15: Added alias handling for authors, improved deduplication of YearCount and PubTypeYearCount in Excel, added Authors column.
# * v1.0.0, 2024-11-14: Initial version. Added --year flag for filtering by year range, adjusted tables and figures by author. Stratified tables and figures for each author in DEFAULT_NAMES. Summarized results in Word and Excel files. Added support for "Access Type" in Publications sheet.

# Version and License Information
VERSION_NAME = 'PubMed Miner'
VERSION = '1.1.0'
VERSION_DATE = '2024-11-18'
COPYRIGHT = 'Copyright 1979-2024. Sander W. van der Laan | s.w.vanderlaan [at] gmail [dot] com | https://vanderlaanand.science.'
COPYRIGHT_TEXT = '''
Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License

By exercising the Licensed Rights (defined below), You accept and agree to be bound by the terms and conditions of this Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License ("Public License"). To the extent this Public License may be interpreted as a contract, You are granted the Licensed Rights in consideration of Your acceptance of these terms and conditions, and the Licensor grants You such rights in consideration of benefits the Licensor receives from making the Licensed Material available under these terms and conditions.

Full license text available at https://creativecommons.org/licenses/by-nc-nd/4.0/legalcode.

This software is provided "as is" without warranties or guarantees of any kind.
'''

# Alias mapping for handling multiple author names and ORCID
ALIAS_MAPPING = {
    "van der Laan SW": [
        "van der Laan SW",
        "van der Laan S", 
        "van der Laan, Sander W",
        "van der Laan Sander W",
        "van der Laan, Sander",
        "van der Laan Sander",
        "Sander van der Laan", 
        "0000-0001-6888-1404",
        # Add other aliases as needed
    ],
    "Pasterkamp G": [
        "Pasterkamp G",
        "Gerard Pasterkamp",
        # Add other aliases as needed
    ],
    "Mokry M": [
        "Mokry M",
        "Michal Mokry",
        # Add other aliases as needed
    ],
    "Schiffelers RM": [
        "Schiffelers RM",
        "Schiffelers R", 
        "Raymond M. Schiffelers", 
        "Raymond Schiffelers", 
        "R. Schiffelers", 
        "R Schiffelers", 
        "Raymond M Schiffelers",
        # Add other aliases as needed
    ],
    "van Solinge W": [
        "van Solinge W",
        "van Solinge WW", 
        "van Solinge W.W.", 
        "Wouter W. van Solinge", 
        "Wouter van Solinge",
        # Add other aliases as needed
    ],
    "Haitjema S": [
        "Haitjema S",
        "Saskia Haitjema",
        "S Haitjema",
        # Add other aliases as needed
    ],
    "den Ruijter HM": [
        "den Ruijter HM",
        "Hester M den Ruijter", 
        "Hester den Ruijter",
        # Add other aliases as needed
    ],
    "Hoefer IE": [
        "Hoefer IE",
        "Imo E Hoefer", 
        "I Hoefer",
        "IE Hoefer",
        "Imo Hoefer",
        "Hofer I", 
        "I Hofer",
        "Imo Hofer", 
        # Add other aliases as needed
    ],
    "Schoneveld AH": [
        "Schoneveld AH", 
        "Schoneveld A", 
        "Arjen H Schoneveld", 
        "Arjen Schoneveld",
        "Schoneveld Arjen",
        # Add other aliases as needed
    ],
    "Vader P": [
        "Vader P",
        "P Vader", 
        "Pieter Vader",
        "Vader Pieter"
        # Add other aliases as needed
    ],
    # Add other aliases as needed
}

# Departement mapping for handling multiple department names
DEPARTMENT_ALIAS_MAPPING = {
    "Central Diagnostic Laboratory": [
        "Central Diagnostic Laboratory",
        "CDL",
        "CDL Research",
        "Central Diagnostics Laboratory",
        "Central Diagnostics Laboratory Research",
        "Central Diagnostic Laboratory, Division Laboratories, Pharmacy, and Biomedical genetics",
        "Central Diagnostics Laboratory, Division Laboratories, Pharmacy, and Biomedical genetics"
        "Central Diagnostic Laboratory, Division Laboratory, Pharmacy, and Biomedical genetics",
        "Laboratory of Clinical Chemistry and Hematology, Division Laboratories and Pharmacy",
        "Laboratory of Clinical Chemistry and Hematology, Division Laboratories & Pharmacy",
        "Laboratory of Clinical Chemistry and Hematology",
        "Laboratory Clinical Chemistry and Hematology",
        # Add other aliases as needed
    ],
    # Add other departments as needed
    "Laboratory of Experimental Cardiology": [
        "Laboratory of Experimental Cardiology",
        "Laboratory Experimental Cardiology",
        "Lab of Experimental Cardiology",
        "Experimental Cardiology",
        "Experimental Cardiology Lab",
        "Experimental Cardiology Laboratory",
        "Laboratory of Experimental Cardiology, Division Heart and Lungs"
        "Laboratory of Experimental Cardiology, Division of Heart and Lungs",
        # Add other aliases as needed
    ]
}

# Organization mapping for handling multiple organization names
ORGANIZATION_ALIAS_MAPPING = {
    "University Medical Center Utrecht": [
        "University Medical Center Utrecht",
        "University Medical Center Utrecht, Utrecht University",
        "UMCU",
        "UMC Utrecht",
        "University Medical Centre Utrecht",
        "Universitair Medisch Centrum Utrecht",
        # Add other aliases as needed
    ]
}

# Set some defaults
DEFAULT_ORGANIZATION = "University Medical Center Utrecht"
DEFAULT_NAMES = ["van der Laan SW", 
"Pasterkamp G", 
"Mokry M", 
"Schiffelers RM", 
"van Solinge W", 
"Haitjema S", 
"den Ruijter HM", 
"Hoefer IE",
"Schoneveld AH",
"Vader P"]
DEFAULT_DEPARTMENTS = ["Central Diagnostic Laboratory", "Laboratory of Experimental Cardiology"]

# Setup Logging
def setup_logger(results_dir, output_base_name, verbose, debug):
    """
    Setup the logger to log to a file and console.
    """
    log_file = os.path.join(results_dir, f"{output_base_name}.log")
    
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

# Normalize text -- needed for validation_results (could be used elsewhere too)
def normalize_text(text):
    """
    Normalize text for comparison: lowercase, remove punctuation, and normalize spaces.
    """
    import re
    text = text.lower()  # Lowercase
    text = re.sub(r'[^\w\s,\.]', '', text)  # Remove punctuation except commas/periods
    text = re.sub(r'\s+', ' ', text).strip()  # Normalize spaces
    return text.strip()

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
# Retry logic for API calls -- needed in validation_results using medline format
def fetch_validation_with_retry(db, term, retries=3, backoff=2):
    """
    Fetch PubMed data with retry logic to handle rate limits or network issues.
    """
    for attempt in range(retries):
        try:
            handle = Entrez.efetch(db=db, id=term, rettype="medline", retmode="text")
            record = handle.read()  # Retrieve the MEDLINE format text
            handle.close()
            if not record:  # Check if the record is empty
                logging.warning(f"No record returned for term [{term}] on attempt {attempt + 1}.")
                continue
            return record  # Return the record as a string
        except Exception as e:
            logging.error(f"Attempt {attempt + 1} failed for term [{term}]: {e}")
            if attempt < retries - 1:
                time.sleep(backoff ** attempt)  # Exponential backoff
    logging.error(f"Failed to fetch data for term [{term}] after {retries} attempts.")
    return None  # Return None if all retries fail

# Extract affiliations -- needed in validation_results
def extract_affiliations(record):
    """
    Extract all affiliations (AD fields) from a MEDLINE record, accounting for multi-line entries.
    """
    affiliations = []
    lines = record.splitlines()

    # Flag to track if we're in an "AD" block
    in_ad_block = False
    current_affiliation = ""

    for line in lines:
        if line.startswith("AD  - "):  # Start of a new affiliation
            if current_affiliation:  # If there's an ongoing affiliation, save it
                affiliations.append(current_affiliation.strip())
            current_affiliation = line[6:]  # Strip "AD  - " prefix
            in_ad_block = True
        elif in_ad_block and line.startswith("      "):  # Continuation line for "AD"
            current_affiliation += " " + line.strip()  # Append the continuation
        else:
            if in_ad_block:  # End of "AD" block
                affiliations.append(current_affiliation.strip())
                current_affiliation = ""
                in_ad_block = False

    # Catch any remaining affiliation
    if current_affiliation:
        affiliations.append(current_affiliation.strip())

    return " ".join(affiliations)  # Combine into a single string

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
    -n, --names <names>          List of (main) author names to search for. Could also be an ORCID.
                                 Default: {DEFAULT_NAMES} with these aliases: {ALIAS_MAPPING}.
    -dep, --departments <depts>  List of departments to search for. Default: {DEFAULT_DEPARTMENTS}.
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
    parser.add_argument("-n", "--names", nargs='+', default=DEFAULT_NAMES, help="List of (main) author names or ORCIDs to search for.")
    parser.add_argument("-dep", "--departments", nargs='+', default=DEFAULT_DEPARTMENTS, help="List of departments to search for.")
    parser.add_argument("-org", "--organization", default=DEFAULT_ORGANIZATION, help="Organization name for filtering results.")
    parser.add_argument("-y", "--year", help="Filter publications by year or year range (e.g., 2024 or 2017-2024).")
    parser.add_argument("-o", "--output-file", default="CDL_UMCU_Publications", help="Output base name for the Word and Excel files.")
    parser.add_argument("-r", "--results-dir", default="results", help="Directory to save results.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output.")
    parser.add_argument("-d", "--debug", action="store_true", help="Enable debug output.")
    parser.add_argument('-V', '--version', action='version', version=f'{VERSION_NAME} {VERSION} ({VERSION_DATE})')
    return parser.parse_args()

# Get canonical author name
def get_canonical_author(author, logger=None):
    """
    Get the canonical author name or ORCID from the ALIAS_MAPPING.
    """
    if re.match(r"^\d{4}-\d{4}-\d{4}-\d{4}$", author):  # Check for ORCID format
        if logger:
            logger.info(f"Processing ORCID: {author}")
        return author  # Return ORCID directly

    for canonical, aliases in ALIAS_MAPPING.items():
        if logger:
            logger.debug(f"Checking aliases for canonical '{canonical}': {aliases}")
        if author in aliases:
            if logger:
                logger.info(f"Matched alias '{author}' to canonical author '{canonical}'.")
            return canonical
        if author == canonical:
            if logger:
                logger.info(f"Matched canonical author '{author}'.")
            return canonical

    if logger:
        logger.warning(f"No match found for author '{author}'. Returning as is.")
    return author


# Fetch publication detailss
def fetch_publication_details(pubmed_ids, logger, main_author, start_year=None, end_year=None):
    """
    Fetch detailed information for each PubMed ID, filter by year, and identify preprints.
    """
    canonical_author = get_canonical_author(main_author, logger)
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
        authors = re.findall(r"AU  - (.+)", record) or []  # Short author list
        full_authors = re.findall(r"FAU - (.+)", record) or []  # Full author list
        author_ids = re.findall(r"AUID- (.+)", record) or []  # Author identifiers (e.g., ORCID)

        # Replace aliases with canonical names for `AU` and `FAU` fields
        authors = [get_canonical_author(author) for author in authors]
        full_authors = [get_canonical_author(author) for author in full_authors]

        # Combine `AU`, `FAU`, and `AUID` into a single list for alias matching
        all_author_data = set(authors + full_authors + author_ids)

        # Check if the canonical author or any alias matches
        author_match = any(alias in all_author_data for alias in all_author_data)

        title = re.search(r"TI  - (.+)", record).group(1) if re.search(r"TI  - (.+)", record) else "No title found"
        journal_abbr = re.search(r"TA  - (.+)", record).group(1) if re.search(r"TA  - (.+)", record) else "No journal abbreviation found"
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
        # Remove "PMID:" from the uof field, if it exists
        uof = uof.replace("PMID:", "").strip()

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
            uof_doi = None
            for part in uof_parts:
                if part.startswith("10."):
                    uof_doi = f"https://doi.org/{part}"
                    break
            uof_citation = " ".join(uof_parts[:uof_parts.index(part)]) if uof_doi else uof

            # Remove "doi:" from the uof_citation if present
            if "doi:" in uof_citation.lower():
                uof_citation = re.sub(r"\bdoi:\b", "", uof_citation, flags=re.IGNORECASE).strip()

            preprints.append(
                (
                    pub_id,
                    canonical_author if canonical_author in authors else f"{canonical_author} et al.",
                    pub_date,
                    journal_abbr,
                    journal_id,
                    title,
                    uof_doi,
                    uof_citation,
                    publication_type,
                )
            )
        # Separate preprints
        if "Preprint" in pub_type:
            preprints.append((pub_id, canonical_author if canonical_author in authors else f"{canonical_author} et al.", pub_date, journal_abbr, journal_id, title, doi_link, source, publication_type))
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
        # logger.debug(f"This was the full record:\n{record}")

    return publications, preprints

# Filter results to ensure all criteria are met
def validate_results(pubmed_ids, logger, author_aliases, department_aliases, organization_aliases):
    """
    Filter the PubMed IDs based on author aliases, department aliases, and organization aliases.
    """
    filtered_ids = set()
    for pub_id in pubmed_ids:
        logger.debug(f"Processing PubMed ID {pub_id}.")
        record = fetch_validation_with_retry(db="pubmed", term=pub_id)  # Fetch detailed record
        
        if not record:  # Check if the record is None or invalid
            logger.warning(f"PubMed ID {pub_id} returned no data. Skipping.")
            continue

        if not isinstance(record, str):
            logger.error(f"Unexpected record format for PubMed ID {pub_id}: {type(record)}. Skipping.")
            continue

        # Extract authors and affiliations
        authors = re.findall(r"AU  - (.+)", record) or []  # Short author list
        logger.debug(f"> Authors: {authors}")
        # Normalize author names
        authors = [normalize_text(author) for author in authors]
        # Normalize author aliases for comparison
        normalized_author_aliases = [normalize_text(alias) for alias in author_aliases]
        logger.debug(f"> Author aliases (normalized): {normalized_author_aliases}")

        # Checking for affiliation match        
        affiliations = extract_affiliations(record)
        logger.debug(f"> Affiliations: {affiliations}")
        # Normalize affiliations
        affiliations = normalize_text(" ".join(re.findall(r"AD  - (.+)", record)))
        logger.debug(f"> Affiliations (normalized): {affiliations}")

        # Normalize department aliases
        logger.debug(f"> Department aliases: {department_aliases}")
        normalized_department_aliases = [normalize_text(alias) for alias in department_aliases]
        logger.debug(f"> Department aliases (normalized): {normalized_department_aliases}")

        # Normalize organization aliases
        logger.debug(f"> Organization aliases: {organization_aliases}")
        normalized_organization_aliases = [normalize_text(alias) for alias in organization_aliases]
        logger.debug(f"> Organization aliases (normalized): {normalized_organization_aliases}")

        # Check for matching author alias
        author_match = any(alias in authors for alias in normalized_author_aliases)
        logger.debug(f"> Author match: {author_match}")

        # Check for matches
        dept_match = any(alias in affiliations for alias in normalized_department_aliases)
        logger.debug(f"> Department match: {dept_match}")

        org_match = any(alias in affiliations for alias in normalized_organization_aliases)
        logger.debug(f"> Organization match: {org_match}")

        if author_match and dept_match and org_match:
            filtered_ids.add(pub_id)
        else:
            logger.debug(f"PubMed ID {pub_id} did not meet all criteria and was excluded.")
    
    return filtered_ids

# Analyze publication data
def analyze_publications(publications_data, main_author, logger):
    """
    Analyze the publication data and return counts for authors, years, and journals.
    Also return counts of publications by author per year and by year per journal.
    """
    canonical_author = get_canonical_author(main_author, logger)
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
def write_to_excel(author_data, output_base_name, results_dir, logger):
    """
    Write the combined results for all authors into six sheets of an Excel file.
    """
    logger.info("Writing results to Excel file.")
    # output_file = f"{datetime.now().strftime('%Y%m%d')}_{output_file}"  # Add date to the filename
    # writer = pd.ExcelWriter(os.path.join(results_dir, f"{output_file}.xlsx"), engine='xlsxwriter')
    output_path = os.path.join(results_dir, f"{output_base_name}.xlsx")
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    # Combine all publications into a single DataFrame
    logger.info("> Combining all publications into a single DataFrame.")
    publications_data = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = get_canonical_author(main_author, logger)
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
        canonical_author = get_canonical_author(main_author, logger)
        for preprint in preprints:
            preprints_data.append(list(preprint) + [canonical_author])
    logger.debug(f"Preprints data: {preprints_data[:3]}")  # Log first 3 preprints
    preprints_df = pd.DataFrame(
        preprints_data,
        columns=["PubMed ID", "Author", "Year", "Journal", "JID", "Title", "DOI Link", "Citation", "Publication Type", "Main Author"]
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
        canonical_author = get_canonical_author(main_author, logger)
        for author, count in author_count.items():
            author_counts.append([author, count, canonical_author])
    author_counts_df = pd.DataFrame(author_counts, columns=["Author", "Number of Publications", "Main Author"])
    author_counts_df.drop(columns=["Main Author"], inplace=True)
    author_counts_df.to_excel(writer, sheet_name="AuthorCount", index=False)

    # Combine and deduplicate year counts
    logger.info("> Combining and deduplicating year counts.")
    year_counts = []
    for main_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        canonical_author = get_canonical_author(main_author, logger)
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
        canonical_author = get_canonical_author(main_author, logger)
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
    logger.info(f"Excel file saved to [{output_path})].\n")
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
def write_to_word(author_data, output_base_name, results_dir, logger, args):
    """
    Write the combined results for all authors to a Word document.
    """
    logger.info("Writing results to Word document.")
    # output_file = f"{datetime.now().strftime('%Y%m%d')}_{output_file}"  # Add date to the filename
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
        f"{output_base_name}_publications_per_author.png",
        f"{output_base_name}_total_publications_preprints_by_author.png",
        # f"{output_base_name}_publications_per_year.png",
        f"{output_base_name}_publications_per_year_with_moving_avg.png",
        f"{output_base_name}_publications_per_author_and_year.png",
        f"{output_base_name}_top10_journals_grouped.png",
        f"{output_base_name}_publications_by_access_type.png",
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
    output_path = os.path.join(results_dir, f"{output_base_name}.docx")
    document.save(output_path)
    logger.info(f"Word document saved to [{output_path}].")

# Plot the results
def plot_results(author_data, results_dir, logger, output_base_name):
    """
    Plot the results for each author and save the plots.
    """
    date_str = datetime.now().strftime('%Y%m%d')
    plot_filenames = {
        "publications_per_author": f"{output_base_name}_publications_per_author.png",
        "total_publications_preprints_by_author": f"{output_base_name}_total_publications_preprints_by_author.png",
        # "publications_per_year": f"{output_base_name}_publications_per_year.png",
        "publications_per_year_with_moving_avg": f"{output_base_name}_publications_per_year_with_moving_avg.png",
        "publications_per_author_and_year": f"{output_base_name}_publications_per_author_and_year.png",
        "top10_journals_grouped": f"{output_base_name}_top10_journals_grouped.png",
        "publications_by_access_type": f"{output_base_name}_publications_by_access_type.png",
    }

    # Consistent color mapping for authors
    canonical_authors = list(author_data.keys())
    colors = plt.colormaps["tab10"](np.linspace(0, 1, len(canonical_authors)))
    color_map = {canonical_author: colors[idx] for idx, canonical_author in enumerate(canonical_authors)}

    # Access type color mapping
    access_color_map = {"open access": "green", "closed access": "red"}

    # PLOT publications per author
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
    plt.savefig(os.path.join(results_dir, plot_filenames["publications_per_author"]))

    # PLOT Total number of publications and preprints grouped by author and year (two panels)
    logger.info("> Plotting total publications and preprints grouped by author and year (two panels).")
    fig, axes = plt.subplots(1, 2, figsize=(16, 7), sharey=True)
    access_types = ["open access", "closed access"]

    # Gather all unique years across both panels
    all_years = sorted(set(
        int(pub[2])  # Extract the year
        for _, (publications, preprints, _, _, _, _, _) in author_data.items()
        for pub in publications + preprints
    ))

    # Width of bars for each author in a grouped bar chart
    bar_width = 0.8 / len(author_data)  # Divide the total width by the number of authors

    # Process data for plotting
    for idx, access_type in enumerate(access_types):
        ax = axes[idx]
        for author_idx, (canonical_author, (publications, preprints, _, _, _, _, _)) in enumerate(author_data.items()):
            yearly_totals = defaultdict(int)
            for pub in publications + preprints:
                pub_access_type = pub[-1]  # Access type is the last field
                year = int(pub[2])  # Year is the third field
                if pub_access_type == access_type:
                    yearly_totals[year] += 1
            
            # Compute bar positions for the author within each year
            positions = [x + author_idx * bar_width for x in range(len(all_years))]
            counts = [yearly_totals.get(year, 0) for year in all_years]  # Ensure all years are represented
            
            # Plot the bars for the current author
            ax.bar(
                positions,
                counts,
                bar_width,
                label=canonical_author,
                color=color_map[canonical_author],
                alpha=0.8,
            )
        
        # Configure the x-axis
        ax.set_title(f"Total Publications ({access_type.capitalize()})")
        ax.set_xlabel("Year")
        ax.set_xticks([x + (bar_width * len(author_data)) / 2 for x in range(len(all_years))])  # Center ticks for years
        ax.set_xticklabels(all_years, rotation=45)  # Use unified years across panels
        
        if idx == 0:
            ax.set_ylabel("Total Publications and Preprints")
        ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")

    plt.tight_layout()
    plt.savefig(os.path.join(results_dir, plot_filenames["total_publications_preprints_by_author"]))

    # PLOT publications per year colored and grouped by author with moving average
    logger.info("> Plotting publications per year with moving average.")
    fig, ax = plt.subplots(figsize=(10, 6))
    width = 0.8 / len(author_data)  # Divide bar width by number of authors for grouped bars

    # Gather all unique years across authors
    all_years = sorted(set(year for _, (_, _, _, year_count, _, _, _) in author_data.items() for year in year_count))

    # Define the moving average window size
    moving_avg_window = 3

    # Plot each author's publications per year with bars and moving average
    for idx, (canonical_author, (_, _, _, year_count, _, _, _)) in enumerate(author_data.items()):
        counts = [year_count.get(year, 0) for year in all_years]  # Publications count for each year
        bar_positions = [year + idx * width for year in range(len(all_years))]  # Bar positions for this author

        # Plot bars for this author
        ax.bar(
            bar_positions,
            counts,
            width=width,
            color=color_map[canonical_author],
            label=canonical_author,
            alpha=0.8,
        )

        # Calculate the moving average (adjust at edges)
        moving_avg = []
        for i in range(len(counts)):
            window_start = max(0, i - moving_avg_window + 1)
            window = counts[window_start:i + 1]
            moving_avg.append(sum(window) / len(window))

        # Plot the moving average line
        ax.plot(
            range(len(moving_avg)),
            moving_avg,
            marker='o',
            label=f"{canonical_author} (Moving Avg)",
            linestyle='--',
            color=color_map[canonical_author],
        )

    # Customize the x-axis
    ax.set_xticks(range(len(all_years)))
    ax.set_xticklabels(all_years, rotation=45)
    ax.set_xlabel("Year")
    ax.set_ylabel("Number of Publications")
    ax.set_title("Publications per Year with Moving Average")
    ax.legend(bbox_to_anchor=(1.05, 1), loc="upper left")
    plt.tight_layout()

    # Save the plot
    plt.savefig(os.path.join(results_dir, plot_filenames["publications_per_year_with_moving_avg"]))

    # PLOT publications per author per year (stacked bar plot)
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
    plt.savefig(os.path.join(results_dir, plot_filenames["publications_per_author_and_year"]))

    # PLOT Top 10 journals by number of publications (grouped by year)
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
    plt.savefig(os.path.join(results_dir, plot_filenames["top10_journals_grouped"]))

    # PLOT Publications grouped by access type and year (not stacked)
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
    plt.savefig(os.path.join(results_dir, plot_filenames["publications_by_access_type"]))

# Main function
def main():
    args = parse_arguments()

    # Make sure the results directory exists
    results_dir = "results"
    os.makedirs(results_dir, exist_ok=True)

    # Get today's date
    today = datetime.now().strftime('%Y%m%d')

    # File base naming convention
    base_name = args.output_file if args.output_file else "CDL_UMCU_Publications"
    output_base_name = f"{today}_{base_name}"
    
    # Set up logging
    logger = setup_logger(results_dir, output_base_name, args.verbose, args.debug)

    # Set year range if provided
    start_year, end_year = None, None
    if args.year:
        start_year, end_year = parse_year_range(args.year)

    # Ensure all required packages are installed
    for package in ['Bio', 'docx', 'matplotlib', 'numpy', 'pandas']:
        check_install_package(package, logger)

    # Set the email for Entrez
    Entrez.email = args.email

    # Print some information
    logger.info(f"Running {VERSION_NAME} v{VERSION} ({VERSION_DATE})\n")
    logger.info(f"Settings:")
    logger.info(f"> Search parameters given:")
    logger.info(f"  - authors: {args.names}")
    logger.info(f"  - department(s): {args.departments}")
    logger.info(f"  - organization: ['{args.organization}']")
    logger.info(f"  - filtering by year (range) [{args.year}]" if args.year else "  - no year filter used.")
    logger.info(f"  - output file(s): [{output_base_name}]")
    logger.info(f"> PubMed email used: {args.email}.")
    logger.info(f"> Debug mode: {'On' if args.debug else 'Off'}.")
    logger.info(f"> Verbose mode: {'On' if args.verbose else 'Off'}.\n")

    author_data = {}
    logger.info(f"Querying PubMed for publications and preprints.\n")

    # Collect all PubMed IDs for the given author(s)
    for main_author in args.names:
        all_pubmed_ids = set()  # Use a set to collect unique IDs
        # Retrieve the canonical author and their aliases
        canonical_author = get_canonical_author(main_author, logger)
        canonical_author_aliases = ALIAS_MAPPING.get(canonical_author, [canonical_author])

        # Construct the author query based on whether it's an ORCID
        author_query = " OR ".join(
            f'({alias}[Author - Identifier])' if re.match(r"^\d{4}-\d{4}-\d{4}-\d{4}$", alias)
            else f'({alias}[Author])'
            for alias in canonical_author_aliases
        )

        logger.info(f"Searching PubMed for canonical author '{canonical_author}' with aliases: {canonical_author_aliases}.")

        for department in args.departments:
            # Retrieve department aliases
            department_aliases = DEPARTMENT_ALIAS_MAPPING.get(department, [department])
            department_query = " OR ".join(f'({alias}[Affiliation])' for alias in department_aliases)

            logger.info(f"> Using department '{department}' with aliases: {department_aliases}.")

            # Retrieve organization aliases
            organization_aliases = ORGANIZATION_ALIAS_MAPPING.get(args.organization, [args.organization])
            organization_query = " OR ".join(f'({alias})' for alias in organization_aliases)

            logger.info(f"> Using organization '{args.organization}' with aliases: {organization_aliases}.")

            # Construct the full search query
            search_query = f"(({author_query}) AND ({department_query})) AND ({organization_query})"
            logger.info(f"Constructed PubMed search query: {search_query}")

            try:
                record = fetch_with_retry(db="pubmed", term=search_query)
                if not record:
                    logger.error(f"No data returned for query [{search_query}]. Skipping.")
                    continue

                all_pubmed_ids.update(record["IdList"])  # Add unique PubMed IDs to the set
            except Exception as e:
                logger.error(f"Failed to fetch PubMed IDs for query [{search_query}]: {e}")
        
        # Validate the results to ensure all criteria are met 
        # INFORMATION -- this is not fully implemented yet, because the validation yields very low results
        # for now, we will skip this step when processing all_pubmed_ids, 
        # logger.info(f"Validating {len(all_pubmed_ids)} unique publications for [{canonical_author}].")
        # # Make sure we list all the department and organization aliases -- not just the one we're previously iterating over
        # validation_department_aliases = DEPARTMENT_ALIAS_MAPPING.get(department, [department])
        # validation_organization_aliases = ORGANIZATION_ALIAS_MAPPING.get(args.organization, [args.organization])
        # validated_pubmed_ids = validate_results(
        #     pubmed_ids=all_pubmed_ids,
        #     logger=logger,
        #     author_aliases=canonical_author_aliases,
        #     department_aliases=validation_department_aliases,
        #     organization_aliases=validation_organization_aliases
        # )
        # logger.info(f"Validated PubMed IDs: {len(validated_pubmed_ids)} unique publications remain after validation.")

        # Log the number of unique IDs for this author
        logger.info(f"Found {len(all_pubmed_ids)} unique publications for author '{canonical_author}'.")

        # Fetch detailed publication data for the current author
        publications, preprints = fetch_publication_details(
            sorted(all_pubmed_ids), logger, canonical_author, start_year, end_year
        )
        author_count, year_count, author_year_count, year_journal_count, pub_type_count = analyze_publications(
            publications, canonical_author, logger
        )
        author_data[canonical_author] = (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count)

        logger.info(f"Found {len(publications)} publications and {len(preprints)} preprints for [{canonical_author}].\n")

    # Summarizing and saving results
    logger.info(f"Done. Summarizing and saving results.\n")
    logger.info(f"Saving plots to [{results_dir}].")
    plot_results(author_data, results_dir, logger, output_base_name)

    # Save results to Word and Excel
    write_to_word(author_data, output_base_name, results_dir, logger, args)
    write_to_excel(author_data, output_base_name, results_dir, logger)

    logger.info(f"Saved the following results:")
    logger.info(f"> Data summarized and saved to {os.path.join(results_dir, f'{output_base_name}.docx')}.")
    logger.info(f"> Excel concatenated and saved to {os.path.join(results_dir, f'{output_base_name}.xlsx')}.")
    logger.info(f"> Plots saved to {results_dir}/.")
    logger.info(f"> Log file saved to {os.path.join(results_dir, f'{output_base_name}.log')}.\n")
    logger.info(f"Thank you for using {VERSION_NAME} v{VERSION} ({VERSION_DATE}).")
    logger.info(f"{COPYRIGHT}")
    logger.info(f"Script completed successfully on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.")

if __name__ == "__main__":
    main()
