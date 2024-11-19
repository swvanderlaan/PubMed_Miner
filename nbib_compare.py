#!/usr/bin/env python3

# Import necessary libraries
import pandas as pd
import re
import os
import logging
import argparse
import importlib
from datetime import datetime

# Change log
# * v1.0.0, 2024-11-14: Initial version. 

# Version Information
VERSION_NAME = "PubMed NBIB Comparator"
VERSION = "1.0.0"
VERSION_DATE = "2024-11-18"
COPYRIGHT = 'Copyright 1979-2024. Sander W. van der Laan | s.w.vanderlaan [at] gmail [dot] com | https://vanderlaanand.science.'
COPYRIGHT_TEXT = '''
Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License

By exercising the Licensed Rights (defined below), You accept and agree to be bound by the terms and conditions of this Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License ("Public License"). To the extent this Public License may be interpreted as a contract, You are granted the Licensed Rights in consideration of Your acceptance of these terms and conditions, and the Licensor grants You such rights in consideration of benefits the Licensor receives from making the Licensed Material available under these terms and conditions.

Full license text available at https://creativecommons.org/licenses/by-nc-nd/4.0/legalcode.

This software is provided "as is" without warranties or guarantees of any kind.
'''

def setup_logger(results_dir, output_file, verbose, debug):
    """
    Setup the logger to log to a file and console.
    """
    # Define the log file
    log_file_basename = output_file.replace('.xlsx', os.path.join(f'_{datetime.now().strftime("%Y%m%d")}.log'))

    # Create the results directory if it does not exist
    os.makedirs(results_dir, exist_ok=True)

    log_file = os.path.join(results_dir, log_file_basename)
    
    logger = logging.getLogger('pubmed_miner')

    # Determine the logging level
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

def parse_nbib_file(nbib_file, logger=None, debug=False):
    """Parses an NBIB file and extracts metadata for each publication."""
    if logger:
        logger.info(f"Parsing NBIB file: {nbib_file}")
    
    with open(nbib_file, 'r') as f:
        records = f.read().split("\n\n")
    
    metadata = []
    for record in records:
        if debug and logger:
            logger.debug(f"Processing record: {record}")
        if not record.strip():
            continue
        # Extract fields
        pmid = re.search(r"PMID- (\d+)", record)
        authors = re.findall(r"FAU - (.+)", record)
        affiliations = re.findall(r"AD  - (.+)", record)
        title = re.search(r"TI  - (.+)", record)
        journal = re.search(r"TA  - (.+)", record)
        journal_id = re.search(r"JID - (\d+)", record)
        year = re.search(r"DP  - (\d{4})", record)
        doi = re.search(r"AID - (10\..+?)(?: \[doi\])", record)
        citation = re.search(r"SO  - (.+)", record)
        pub_type = re.findall(r"PT  - (.+)", record)

        if pmid:
            metadata.append({
                "PubMed ID": pmid.group(1),
                "Authors": "; ".join(authors) if authors else "N/A",
                "Author Affiliations": "; ".join(affiliations) if affiliations else "N/A",
                "Title": title.group(1) if title else "N/A",
                "Journal": journal.group(1) if journal else "N/A",
                "JID": journal_id.group(1) if journal_id else "N/A",
                "Year": year.group(1) if year else "N/A",
                "DOI Link": f"https://doi.org/{doi.group(1)}" if doi else "N/A",
                "Citation": citation.group(1) if citation else "N/A",
                "Publication Type": ", ".join(pub_type) if pub_type else "N/A",
            })
    return pd.DataFrame(metadata)

# Match PMIDs and find unmatched ones
# def match_pmids(excel_file, nbib_data, logger=None, debug=False):
#     """Matches PMIDs from Excel file with NBIB data and identifies unmatched PMIDs."""
#     if logger:
#         logger.info(f"Reading Excel file: {excel_file}")
    
#     # Load Excel data
#     if debug and logger:
#         logger.debug(f"> Loading Excel...")
#     excel_data = pd.read_excel(excel_file, sheet_name="Publications")
    
#     # Ensure column names match expectations
#     if debug and logger:
#         logger.debug(f"> Renaming columns...")
#     excel_data.rename(columns={"PubMed ID": "PubMed ID"}, inplace=True)
    
#     # Find unmatched PMIDs
#     if debug and logger:
#         logger.debug(f"> Finding unmatched PMIDs...")
#     unmatched_pmids = excel_data[~excel_data["PubMed ID"].astype(str).isin(nbib_data["PubMed ID"].astype(str))]
    
#     if logger:
#         logger.info(f"Found {len(unmatched_pmids)} unmatched PMIDs.")
#     return unmatched_pmids
def match_pmids(excel_file, nbib_data, logger=None, debug=False):
    """Matches PMIDs between Excel file and NBIB data, identifies unmatched PMIDs in both directions."""
    if logger:
        logger.info(f"Reading Excel file: {excel_file}")
    
    # Load Excel data
    if debug and logger:
        logger.debug(f"> Loading Excel...")
    excel_data = pd.read_excel(excel_file, sheet_name="Publications")
    
    # Ensure column names match expectations
    if debug and logger:
        logger.debug(f"> Renaming columns...")
    excel_data.rename(columns={"PubMed ID": "PubMed ID"}, inplace=True)
    
    # Find unmatched PMIDs (Excel -> NBIB)
    if debug and logger:
        logger.debug(f"> Finding PMIDs in Excel not found in NBIB...")
    unmatched_pmids_excel = excel_data[~excel_data["PubMed ID"].astype(str).isin(nbib_data["PubMed ID"].astype(str))]
    
    # Find unmatched PMIDs (NBIB -> Excel)
    if debug and logger:
        logger.debug(f"> Finding PMIDs in NBIB not found in Excel...")
    unmatched_pmids_nbib = nbib_data[~nbib_data["PubMed ID"].astype(str).isin(excel_data["PubMed ID"].astype(str))]
    
    if logger:
        logger.info(f"Found {len(unmatched_pmids_excel)} unmatched PMIDs in Excel.")
        logger.info(f"Found {len(unmatched_pmids_nbib)} unmatched PMIDs in NBIB.")
    
    return unmatched_pmids_excel, unmatched_pmids_nbib

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description=f"""
{VERSION_NAME} v{VERSION} ({VERSION_DATE})
Compares PubMed records from an NBIB file with an Excel output to find unmatched PMIDs.""",
epilog=f"""
This script compares PubMed records from an NBIB file with an Excel output generated with `pubmed_miner.py` 
to find unmatched PMIDs.

The output file will contain the unmatched PMIDs from the Excel file including the following metadata:
- PubMed ID
- Title
- Journal
- JID
- Year
- DOI Link
- Citation
- Publication Type

Required arguments:
    --input-nbib       Path to the input NBIB file.
    --input-excel      Path to the input Excel file (e.g., results from pubmed_miner.py).

Optional arguments:
    --output-file      Path to the output file. Defaults to the input Excel basename with '_unmatched_nbib_pmids.xlsx'.
    -v, --verbose      Enable verbose output.
    -d, --debug        Enable debug output.
    -V, --version      Show the version information.
""")
    parser.add_argument("-in", "--input-nbib", required=True, help="Path to the input NBIB file. Requires a text file with PubMed records in NBIB format. Required.")
    parser.add_argument("-ie", "--input-excel", required=True, help="Path to the input Excel file (e.g., results from pubmed_miner.py). Required.")
    parser.add_argument("-o", "--output-file", help="Path to the output file. Defaults to the input Excel basename with '_unmatched_nbib_pmids.xlsx'. Optional.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output.")
    parser.add_argument("-d", "--debug", action="store_true", help="Enable debug output.")
    parser.add_argument('-V', '--version', action='version', version=f'{VERSION_NAME} {VERSION} ({VERSION_DATE})')
    return parser.parse_args()

def main():
    # Parse arguments
    args = parse_arguments()
    
    # Make sure the results directory exists
    results_dir = "results"
    os.makedirs(results_dir, exist_ok=True)

    # Get today's date
    today = datetime.now().strftime('%Y%m%d')

    # File base naming convention
    if args.output_file:
        output_file = args.output_file
    else:
        output_basename = os.path.splitext(os.path.basename(args.input_excel))[0]
        output_file = f"{today}_{output_basename}_unmatched_nbib_pmids.xlsx"
    
    # Setup logger
    logger = setup_logger(results_dir, output_file, args.verbose, args.debug)

    # Ensure all required packages are installed
    for package in ['pandas']:
        check_install_package(package, logger)
    
    # Print some information
    logger.info(f"Running {VERSION_NAME} v{VERSION} ({VERSION_DATE})\n")
    logger.info(f"Settings:")
    logger.info(f"Input NBIB file: {args.input_nbib}")
    logger.info(f"Input Excel file: {args.input_excel}")
    logger.info(f"Output file: {output_file}")
    logger.info(f"> Debug mode: {'On' if args.debug else 'Off'}.")
    logger.info(f"> Verbose mode: {'On' if args.verbose else 'Off'}.\n")

    # Parse NBIB file
    if args.verbose:
        logger.info(f"Parsing NBIB file: {args.input_nbib}")
    nbib_data = parse_nbib_file(args.input_nbib, debug=args.debug)
    
    # # Match PMIDs and find unmatched ones
    # if args.verbose:
    #     logger.info(f"Matching PMIDs from Excel file with NBIB data...")
    # unmatched_pmids = match_pmids(args.input_excel, nbib_data, debug=args.debug)
    
    # # Save unmatched PMIDs to an Excel file
    # logger.info(f"Saving unmatched PMIDs...")
    # unmatched_pmids.to_excel(os.path.join(results_dir, output_file), index=False)

    # Match PMIDs and find unmatched ones in both directions
    if args.verbose:
        logger.info(f"Matching PMIDs between Excel file and NBIB data...")
    unmatched_pmids_excel, unmatched_pmids_nbib = match_pmids(args.input_excel, nbib_data, logger=logger, debug=args.debug)

    # Save unmatched PMIDs from Excel to an Excel file
    unmatched_excel_file = os.path.join(results_dir, f"{today}_{output_basename}_unmatched_excel_pmids.xlsx")
    logger.info(f"Saving unmatched PMIDs from Excel to NBIB...")
    unmatched_pmids_excel.to_excel(unmatched_excel_file, index=False)

    # Save unmatched PMIDs from NBIB to an Excel file
    unmatched_nbib_file = os.path.join(results_dir, output_file)
    logger.info(f"Saving unmatched PMIDs from NBIB to Excel...")
    unmatched_pmids_nbib.to_excel(unmatched_nbib_file, index=False)

    logger.info(f"Saved the following results:")
    # logger.info(f"> Data summarized and saved to {output_file}.")
    logger.info(f"> Unmatched PMIDs from Excel saved to {unmatched_excel_file}.")
    logger.info(f"> Unmatched PMIDs from NBIB saved to {unmatched_nbib_file}.")
    logger.info(f"> Log file saved to {f'{today}_{output_basename}_unmatched_nbib_pmids.log'}.\n")
    logger.info(f"Thank you for using {VERSION_NAME} v{VERSION} ({VERSION_DATE}).")
    logger.info(f"{COPYRIGHT}")
    logger.info(f"Script completed successfully on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.")

if __name__ == "__main__":
    main()
