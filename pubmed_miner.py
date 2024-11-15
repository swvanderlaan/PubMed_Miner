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
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Change log:
# * v1.0.3, 2024-11-15: Expanded Word-document information.
# * v1.0.2, 2024-11-15: Added retry logic for PubMed API, better logging for aliases, results directory customization, improved input validation, enhanced plotting, and bar annotations.
# * v1.0.1, 2024-11-15: Added alias handling for authors, improved deduplication of YearCount and PubTypeYearCount in Excel, added Authors column.
# * v1.0.0, 2024-11-14: Initial version. Added --year flag for filtering by year range, adjusted tables and figures by author. Stratified tables and figures for each author in DEFAULT_NAMES. Summarized results in Word and Excel files. Added support for "Access Type" in Publications sheet.

# Version and License Information
VERSION_NAME = 'PubMed Miner'
VERSION = '1.0.3'
VERSION_DATE = '2024-11-15'
COPYRIGHT = 'Copyright 1979-2024. Sander W. van der Laan | s.w.vanderlaan [at] gmail [dot] com | https://vanderlaanand.science.'
COPYRIGHT_TEXT = '''
The MIT License (MIT).
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
associated documentation files (the "Software"), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, 
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is 
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies 
or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR 
PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS 
BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, 
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE 
OR OTHER DEALINGS IN THE SOFTWARE.

Reference: http://opensource.org.
'''

# Alias mapping for handling multiple author names
ALIAS_MAPPING = {
    "Schiffelers R": "Schiffelers RM",
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
DEFAULT_NAMES = ["van der Laan SW", "Pasterkamp G", "Mokry M", "Schiffelers RM", "van Solinge W", "Haitjema S"]
DEFAULT_DEPARTMENTS = ["Central Diagnostic Laboratory"]

# Setup Logging
def setup_logger(results_dir, verbose):
    """
    Setup the logger to log to a file and console.
    """
    date_str = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(results_dir, f"{date_str}_CDL_UMCU_Publications.log")
    os.makedirs(results_dir, exist_ok=True)
    
    logger = logging.getLogger('pubmed_miner')
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    
    fh = logging.FileHandler(log_file)
    fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG if verbose else logging.INFO)
    
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
    -d, --departments <depts>    List of departments to search for. Default: {DEFAULT_DEPARTMENTS}.
    -org, --organization <org>   Organization name for filtering results. Default: {DEFAULT_ORGANIZATION}.
    -y, --year <year>            Filter publications by year or year range (e.g., 2024 or 2017-2024).
    -o, --output-file <file>     Output base name for the Word and Excel files. Default: date_CDL_UMCU_Publications.
    -r, --results-dir <dir>      Directory to save results. Default: results.
    -v, --verbose                Enable verbose output.
    -V, --version                Show program's version number and exit.

Example:
    python pubmed_miner.py --email <email-address> --year 2017-2024 --verbose
+ {VERSION_NAME} v{VERSION}. {COPYRIGHT} +
{COPYRIGHT_TEXT}""",
        formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("-e", "--email", required=True, help="Email for PubMed API access.")
    parser.add_argument("-n", "--names", nargs='+', default=DEFAULT_NAMES, help="List of author names to search for.")
    parser.add_argument("-d", "--departments", nargs='+', default=DEFAULT_DEPARTMENTS, help="List of departments to search for.")
    parser.add_argument("-org", "--organization", default=DEFAULT_ORGANIZATION, help="Organization name for filtering results.")
    parser.add_argument("-y", "--year", help="Filter publications by year or year range (e.g., 2024 or 2017-2024).")
    parser.add_argument("-o", "--output-file", default="CDL_UMCU_Publications", help="Output base name for the Word and Excel files.")
    parser.add_argument("-r", "--results-dir", default="results", help="Directory to save results.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output.")
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
        journal = re.search(r"JT  - (.+)", record).group(1) if re.search(r"JT  - (.+)", record) else "No journal found"
        pub_date = re.search(r"DP  - (.+)", record).group(1)[:4] if re.search(r"DP  - (.+)", record) else "Unknown Year"
        doi_match = re.search(r"AID - (10\..+?)(?: \[doi\])", record)
        doi_link = f"https://doi.org/{doi_match.group(1)}" if doi_match else "No DOI found"
        pub_type = re.findall(r"PT  - (.+)", record)
        publication_type = 'Other'
        if "Journal Article" in pub_type:
            publication_type = "Journal Article"
        elif "Review" in pub_type:
            publication_type = "Review"
        elif "Book" in pub_type:
            publication_type = "Book"

        # Determine access type
        access_type = "open access" if "PMC - Free Full Text" in record else "closed access"

        # Skip errata or corrections
        if "ERRATUM" in title.upper() or "AUTHOR CORRECTION" in title.upper():
            logger.info(f"Skipping 'erratum' or 'author correction' for [{title}]")
            continue

        # Filter by year if specified
        if start_year and end_year:
            if not (start_year <= int(pub_date) <= end_year):
                logger.info(f"Skipping [{title}] as it falls outside the year range.")
                continue

        # Separate preprints
        if "Preprint" in pub_type:
            preprints.append((pub_id, canonical_author if canonical_author in authors else f"{canonical_author} et al.", pub_date, journal, title, publication_type, doi_link, access_type))
        else:
            publications.append((pub_id, canonical_author if canonical_author in authors else f"{canonical_author} et al.", pub_date, journal, title, publication_type, doi_link, access_type))

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
        pub_id, author, year, journal, title, pub_type, doi_link, access_type = pub
        if canonical_author in author:
            author_count[canonical_author] += 1
            author_year_count[canonical_author][year] += 1
            year_count[year] += 1
            year_journal_count[year][journal] += 1
            pub_type_count[pub_type][year] += 1
    
    return author_count, year_count, author_year_count, year_journal_count, pub_type_count

# Write results to Excel
def write_to_excel(author_data, output_file, results_dir, logger):
    """
    Write the combined results for all authors into six sheets of an Excel file.
    Each sheet represents one of the following tables:
    1. Publications (deduplicated by PubMed ID with combined authors, includes access type)
    2. Preprints (deduplicated by PubMed ID with combined authors)
    3. Number of Publications per Author
    4. Number of Publications per Year (deduplicated with Authors column)
    5. Number of Publications per Year per Journal (deduplicated by PubMed ID with combined authors)
    6. Publication Type by Year (deduplicated with Authors column)
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
    publications_df = pd.DataFrame(
        publications_data,
        columns=["PubMed ID", "Author(s)", "Year", "Journal", "Title", "Publication Type", "DOI Link", "Access Type", "Main Author"]
    )

    # Deduplicate based on PubMed ID and combine authors for duplicates in publications
    publications_df = (
        publications_df.groupby("PubMed ID")
        .agg({
            "Author(s)": lambda x: ", ".join(sorted(set(x))),  # Combine full unique author names
            "Year": "first",
            "Journal": "first",
            "Title": "first",
            "Publication Type": "first",
            "DOI Link": "first",
            "Access Type": "first",
            "Main Author": lambda x: ", ".join(sorted(set(x)))  # Combine main authors
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
    preprints_df = pd.DataFrame(
        preprints_data,
        columns=["PubMed ID", "Author(s)", "Year", "Journal", "Title", "Publication Type", "DOI Link", "Access Type", "Main Author"]
    )

    # Deduplicate based on PubMed ID and combine authors for duplicates in preprints
    preprints_df = (
        preprints_df.groupby("PubMed ID")
        .agg({
            "Author(s)": lambda x: ", ".join(sorted(set(x))),  # Combine full unique author names
            "Year": "first",
            "Journal": "first",
            "Title": "first",
            "Publication Type": "first",
            "DOI Link": "first",
            "Access Type": "first",
            "Main Author": lambda x: ", ".join(sorted(set(x)))  # Combine main authors
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

    document.add_heading(f"Publications for {query_quarter} at the Central Diagnostics Laboratory", level=1)
    document.add_paragraph(f"This document summarizes the publications linked to the Central Diagnostics Laboratory (CDL) at the University Medical Center Utrecht (UMCU).\n")
    document.add_paragraph()
    document.add_paragraph(f"The following settings are used:")
    document.add_paragraph(f"* Query date: {query_date}.")
    document.add_paragraph(f"* Authors: {', '.join(author_data.keys())}.\n")
    document.add_paragraph(f"* Aliases used: {', '.join(ALIAS_MAPPING.keys())}.\n")
    document.add_paragraph(f"* Year range: {args.year}.\n" if args.year else "No year filter used.\n")
    document.add_paragraph(f"* Department(s): {', '.join(DEFAULT_DEPARTMENTS)}.\n")
    document.add_paragraph(f"* Organization: {DEFAULT_ORGANIZATION}.\n")
    document.add_paragraph()
    document.add_paragraph(f"Results saved on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.\n")
    document.add_paragraph(f"Log file saved to {os.path.join(results_dir, f'{query_date}_CDL_UMCU_Publications.log')}.\n")
    document.add_paragraph()
    document.add_paragraph(f"{VERSION_NAME} v{VERSION} ({VERSION_DATE}).\n")
    document.add_paragraph(f"{COPYRIGHT}.\n")
    document.add_paragraph()
    document.add_paragraph(f"GitHub repository: https://github.com/swvanderlaan/PubMed_Miner. Any issues or requests? Create one here: https://github.com/swvanderlaan/PubMed_Miner/issues.\n")

    # Add results for each canonical author
    logger.info(f"> Adding results for {len(author_data)} author(s).")
    for canonical_author, (publications, preprints, author_count, year_count, author_year_count, year_journal_count, pub_type_count) in author_data.items():
        document.add_heading(f"Author: {canonical_author}", level=1)

        # Main publications and preprints for this author
        if publications:
            logger.info(f"Adding {len(publications)} publications for {canonical_author}.")
            valid_publications = [row for row in publications if len(row) == 8]
            add_table_to_doc(
                document,
                valid_publications,
                headers=["PubMed ID", "Author", "Year", "Journal", "Title", "Publication Type", "DOI Link", "Access Type"],
                title="Main Publications",
            )
        else:
            logger.warning(f"No publications found for {canonical_author}.")

        if preprints:
            logger.info(f"Adding {len(preprints)} preprints for {canonical_author}.")
            valid_preprints = [row for row in preprints if len(row) == 8]
            add_table_to_doc(
                document,
                valid_preprints,
                headers=["PubMed ID", "Author", "Year", "Journal", "Title", "Publication Type", "DOI Link", "Access Type"],
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
def plot_results(author_data, results_dir):
    """
    Plot the results for each author and save the plots.
    """
    date_str = datetime.now().strftime('%Y%m%d')

    # Consistent color mapping for authors
    canonical_authors = list(author_data.keys())
    colors = plt.colormaps["tab10"](np.linspace(0, 1, len(canonical_authors)))
    color_map = {canonical_author: colors[idx] for idx, canonical_author in enumerate(canonical_authors)}

    # Plot publications per author
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

    # Bar plot for publications per year (stacked by author)
    fig, ax = plt.subplots()
    width = 0.8 / len(author_data)

    # Gather all unique years across authors
    all_years = sorted(set(year for _, (_, _, _, year_count, _, _, _) in author_data.items() for year in year_count))

    # Plot each author's publications per year with unique color
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

# Main function
def main():
    args = parse_arguments()
    results_dir = "results"
    os.makedirs(results_dir, exist_ok=True)
    today = datetime.now().strftime('%Y%m%d')
    output_file = os.path.join(results_dir, f"{today}_{args.output_file}")
    logger = setup_logger(results_dir, args.verbose)

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
    plot_results(author_data, results_dir)

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
