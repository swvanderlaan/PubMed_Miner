#!/usr/bin/env python3

import argparse
import logging
import os
from datetime import datetime
from Bio import Entrez

# Change log:
# * v1.0.0, 2024-11-15: Initial version. Fetches and displays PubMed article metadata.

# Version and License Information
VERSION_NAME = 'Article Metadata Viewer'
VERSION = '1.0.0'
VERSION_DATE = '2024-11-15'
COPYRIGHT = 'Copyright 1979-2024. Sander W. van der Laan | s.w.vanderlaan [at] gmail [dot] com | https://vanderlaanand.science.'
COPYRIGHT_TEXT = '''
Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License

By exercising the Licensed Rights (defined below), You accept and agree to be bound by the terms and conditions of this Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License ("Public License"). To the extent this Public License may be interpreted as a contract, You are granted the Licensed Rights in consideration of Your acceptance of these terms and conditions, and the Licensor grants You such rights in consideration of benefits the Licensor receives from making the Licensed Material available under these terms and conditions.

Full license text available at https://creativecommons.org/licenses/by-nc-nd/4.0/legalcode.

This software is provided "as is" without warranties or guarantees of any kind.
'''

# Setup Logging
def setup_logger(verbose):
    """
    Setup logger to log to console.
    """
    logger = logging.getLogger('article_meta')
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG if verbose else logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger

# Fetch metadata for a given PubMed ID
def fetch_pubmed_metadata(pubmed_id, email, logger):
    """
    Fetch and display metadata for a PubMed article.
    :param pubmed_id: PubMed ID of the article.
    :param email: Email address to use for PubMed API.
    :param logger: Logger object for logging.
    """
    Entrez.email = email
    try:
        # Fetch metadata in MEDLINE format
        logger.info(f"Fetching metadata for PubMed ID {pubmed_id}.")
        with Entrez.efetch(db="pubmed", id=pubmed_id, rettype="medline", retmode="text") as handle:
            record = handle.read()
        
        # Display metadata fields
        logger.info(f"Metadata for PubMed ID {pubmed_id}:")
        fields = record.split("\n")
        for field in fields:
            if field.strip():  # Ignore empty lines
                print(field)
    except Exception as e:
        logger.error(f"Error fetching metadata for PubMed ID {pubmed_id}: {e}")

# Argument Parsing
def parse_arguments():
    """
    Parse command-line arguments.
    """
    parser = argparse.ArgumentParser(description=f"""
{VERSION_NAME} v{VERSION} ({VERSION_DATE})
Fetch and display metadata for a PubMed article.
    """,
    epilog=f"""
This script requires a PubMed ID and an email address to access the PubMed API.
It will fetch metadata for the specified article and display it in the console.

Required arguments:
    -e, --email <email-address>          Email address for PubMed API access.
    -p, --pubmedid PUBMEDID    PubMed ID of the article to fetch metadata for.

Optional arguments:
    -v, --verbose              Enable verbose output.
    -V, --version              Show program's version number and exit.

Example:
    python article_meta.py --email <email-address> --pubmedid 38698167 --verbose
{COPYRIGHT_TEXT}
    """,
    formatter_class=argparse.RawTextHelpFormatter)
    
    parser.add_argument("-e", "--email", required=True, help="Email for PubMed API access.")
    parser.add_argument("-p", "--pubmedid", required=True, help="PubMed ID of the article to fetch metadata for.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output.")
    parser.add_argument('-V', '--version', action='version', version=f'{VERSION_NAME} {VERSION} ({VERSION_DATE})')
    return parser.parse_args()

# Main function
def main():
    args = parse_arguments()
    logger = setup_logger(args.verbose)

    logger.info(f"Running {VERSION_NAME} v{VERSION} ({VERSION_DATE})")
    logger.info(f"Settings:")
    logger.info(f"> PubMed ID: {args.pubmedid}")
    logger.info(f"> Email: {args.email}")
    logger.info(f"> Verbose mode: {'On' if args.verbose else 'Off'}\n")

    fetch_pubmed_metadata(args.pubmedid, args.email, logger)
    logger.info(f"Script completed successfully on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.")

if __name__ == "__main__":
    main()
