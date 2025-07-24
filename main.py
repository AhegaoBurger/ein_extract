#!/usr/bin/env python3
"""
EIN Presswire Distribution Report PDF to Excel Converter

This script extracts media outlet names and their corresponding URLs
from EIN Presswire distribution report PDFs, including embedded hyperlinks.
"""

import pdfplumber
import pandas as pd
import re
import sys
from pathlib import Path
import PyPDF2
from collections import defaultdict

def extract_hyperlinks_pypdf2(pdf_path):
    """
    Extract hyperlinks and their page numbers using PyPDF2

    Returns:
        Dict mapping page numbers to lists of URLs
    """
    page_urls = defaultdict(list)

    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)

            for page_num, page in enumerate(pdf_reader.pages):
                if '/Annots' in page:
                    annotations = page['/Annots']
                    for annotation_ref in annotations:
                        annotation = annotation_ref.get_object()
                        if annotation.get('/Subtype') == '/Link':
                            if '/A' in annotation and '/URI' in annotation['/A']:
                                url = annotation['/A']['/URI']
                                page_urls[page_num].append(url)
    except Exception as e:
        print(f"Warning: Could not extract hyperlinks with PyPDF2: {e}")

    return page_urls

def extract_hyperlinks_pdfplumber(pdf_path):
    """
    Extract hyperlinks using pdfplumber's hyperlinks feature

    Returns:
        Dict mapping page numbers to lists of (url, bbox) tuples
    """
    page_hyperlinks = defaultdict(list)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                hyperlinks = page.hyperlinks
                for link in hyperlinks:
                    if 'uri' in link:
                        # Store URL with its bounding box for position info
                        page_hyperlinks[page_num].append({
                            'url': link['uri'],
                            'bbox': link.get('bbox', None),
                            'top': link.get('top', 0)
                        })
    except Exception as e:
        print(f"Warning: Could not extract hyperlinks with pdfplumber: {e}")

    return page_hyperlinks

def extract_media_reprints_with_urls(pdf_path):
    """
    Extract media outlet names and their corresponding URLs from EIN Presswire PDF

    Args:
        pdf_path: Path to the PDF file

    Returns:
        List of tuples (agency_name, url)
    """
    results = []
    seen_agencies = set()

    # First, extract all hyperlinks from the PDF
    print("Extracting hyperlinks from PDF...")
    hyperlinks_pypdf2 = extract_hyperlinks_pypdf2(pdf_path)
    hyperlinks_pdfplumber = extract_hyperlinks_pdfplumber(pdf_path)

    # Common headers to skip
    skip_phrases = [
        'Media Reprints:',
        'Click on the logos to view the reprints.',
        'Here is a list of independent news publishers',
        'Publishers are not required to report back to us',
        'You could very well appear on others.',
        'News Databases:',
        'Your press release is now also available',
        'Note: Paid access is often required',
        'Industry Newswire Highlights:',
        'Complete Newswire Distribution:',
        'World Media Directory:',
        'Report Summary',
        'View tracking links',
        'Download PDF Report',
        'Boost your reach',
        'Share now on',
        'Click below to',
        'Click on the boxes below'
    ]

    with pdfplumber.open(pdf_path) as pdf:
        in_media_section = False

        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue

            # Get hyperlinks for this page
            page_links_pypdf2 = hyperlinks_pypdf2.get(page_num, [])
            page_links_pdfplumber = hyperlinks_pdfplumber.get(page_num, [])

            # Combine all URLs from both methods
            all_page_urls = list(page_links_pypdf2)
            all_page_urls.extend([link['url'] for link in page_links_pdfplumber])

            # Remove duplicates while preserving order
            seen_urls = set()
            unique_urls = []
            for url in all_page_urls:
                if url not in seen_urls and not url.startswith('mailto:'):
                    seen_urls.add(url)
                    unique_urls.append(url)

            # Check if we're in the Media Reprints section
            if 'Media Reprints:' in text:
                in_media_section = True
                print(f"Found Media Reprints section on page {page_num + 1}")

            # Stop when we reach other sections
            if any(marker in text for marker in ['News Databases:', 'Industry Newswire Highlights:',
                                                  'World Media Directory:', 'Report Summary']):
                if in_media_section and 'News Databases:' not in text:
                    break

            if not in_media_section:
                continue

            # Extract text with layout preservation for better matching
            layout_text = page.extract_text(layout=True)
            if layout_text:
                lines = layout_text.split('\n')
            else:
                lines = text.split('\n')

            # Track which URLs we've used
            used_urls = set()

            # Process lines to find media outlets
            for i, line in enumerate(lines):
                line = line.strip()

                # Skip empty lines and known headers
                if not line or any(skip.lower() in line.lower() for skip in skip_phrases):
                    continue

                # Additional filters
                if (len(line) > 60 or  # Too long to be an outlet name
                    line.count(' ') > 8 or  # Too many spaces
                    any(char in line for char in ['@', '%', '&amp;', '...', '|']) or  # Special chars
                    line.startswith('(') or  # Parenthetical info
                    line.startswith('http')):  # URLs
                    continue

                # Check if this looks like a media outlet
                upper_count = sum(1 for c in line if c.isupper())
                if upper_count < 2:
                    continue

                # Common patterns for media outlets
                tv_pattern = r'^[A-Z]{3,5}\s+(?:ABC|NBC|CBS|FOX|CW|MyNetworkTV)\s*\d*'
                media_pattern = r'^[A-Z][A-Za-z\s\-&\.\']+(?:\s+(?:News|Times|Journal|Post|Daily|Weekly|Today|Online|Report|Observer|Press|Media|Review|Update|Herald|Gazette|Tribune|Monitor|Entertainment|Tech|Finance|Business))*$'
                caps_pattern = r'^[A-Z\s\-]+$'

                if (re.match(tv_pattern, line) or
                    re.match(media_pattern, line) or
                    (re.match(caps_pattern, line) and len(line) < 40)):

                    agency = line.strip()

                    # Skip if already seen
                    if agency in seen_agencies:
                        continue

                    # Try to find the best matching URL
                    url = "URL_TBD"

                    # Strategy: Find unused URL that hasn't been assigned yet
                    for candidate_url in unique_urls:
                        if candidate_url not in used_urls:
                            url = candidate_url
                            used_urls.add(candidate_url)
                            break

                    seen_agencies.add(agency)
                    results.append((agency, url))

                    if url != "URL_TBD":
                        print(f"Matched: {agency} -> {url[:50]}...")

    return results

def save_to_excel(data, output_path):
    """
    Save the extracted data to an Excel file

    Args:
        data: List of tuples (agency_name, url)
        output_path: Path to save the Excel file
    """
    # Create DataFrame
    df = pd.DataFrame(data, columns=['Agency', 'URL'])

    # Save to Excel with formatting
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Media Outlets')

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Media Outlets']

        # Adjust column widths
        worksheet.column_dimensions['A'].width = 40
        worksheet.column_dimensions['B'].width = 80

        # Add hyperlinks for actual URLs
        for idx, (_, url) in enumerate(df.values, start=2):
            if url and url.startswith('http'):
                cell = worksheet.cell(row=idx, column=2)
                cell.hyperlink = url
                cell.style = 'Hyperlink'

    print(f"Excel file saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) != 2:
        print("Usage: python ein_pdf_to_excel.py <pdf_file>")
        print("\nExample:")
        print("  python ein_pdf_to_excel.py EINPresswire-Report.pdf")
        sys.exit(1)

    pdf_path = Path(sys.argv[1])

    if not pdf_path.exists():
        print(f"Error: File '{pdf_path}' not found")
        sys.exit(1)

    if not pdf_path.suffix.lower() == '.pdf':
        print("Error: Input file must be a PDF")
        sys.exit(1)

    # Generate output filename
    output_path = pdf_path.with_suffix('.xlsx')

    print(f"Processing: {pdf_path}")
    print("Extracting media outlets and hyperlinks from PDF...")

    try:
        # Extract data
        data = extract_media_reprints_with_urls(pdf_path)

        if not data:
            print("Warning: No media outlets found in the PDF")
            print("Make sure the PDF contains a 'Media Reprints:' section")
        else:
            print(f"\nFound {len(data)} unique media outlets")

            # Count how many have actual URLs
            url_count = sum(1 for _, url in data if url != "URL_TBD")
            print(f"Successfully extracted {url_count} hyperlinks")

        # Save to Excel
        save_to_excel(data, output_path)

        # Print summary
        if data:
            print("\nFirst 10 entries:")
            for i, (agency, url) in enumerate(data[:10], 1):
                if url == "URL_TBD":
                    url_display = "URL_TBD"
                else:
                    url_display = url[:50] + "..." if len(url) > 50 else url
                print(f"  {i:2d}. {agency:<35} -> {url_display}")

            if len(data) > 10:
                print(f"  ... and {len(data) - 10} more entries")

    except Exception as e:
        print(f"Error processing PDF: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
