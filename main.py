#!/usr/bin/env python3
"""
EIN Presswire PDF to Excel Converter - Streamlit Web App
A user-friendly web interface for converting EIN Presswire PDFs to Excel
"""

import streamlit as st
import pdfplumber
import pandas as pd
import PyPDF2
from collections import defaultdict
import re
from pathlib import Path
import io

# Only run the app when using streamlit run
if __name__ == "__main__":
    # Check if running with streamlit
    try:
        # This will only work when running with streamlit
        _ = st.runtime.exists()
    except:
        print("\n⚠️  Please run this app using streamlit:")
        print("    uv run streamlit run main.py")
        print("\nOr if you have streamlit installed globally:")
        print("    streamlit run main.py\n")
        exit(1)

# Page config
st.set_page_config(
    page_title="EIN Presswire PDF to Excel Converter",
    page_icon="📄",
    layout="centered"
)

# Title and description
st.title("📄 EIN Presswire PDF to Excel Converter")
st.markdown("""
This tool extracts media outlet names and their URLs from EIN Presswire Distribution Report PDFs
and saves them to an Excel file.

### How to use:
1. Upload your EIN Presswire PDF file below
2. Click "Convert to Excel"
3. Download the resulting Excel file
""")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

def extract_hyperlinks_pypdf2(pdf_file):
    """Extract hyperlinks using PyPDF2"""
    page_urls = defaultdict(list)

    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        for page_num, page in enumerate(pdf_reader.pages):
            if '/Annots' in page:
                annotations = page['/Annots']
                # Handle the case where annotations might not be iterable
                if hasattr(annotations, '__iter__'):
                    for annotation_ref in annotations:
                        try:
                            annotation = annotation_ref.get_object()
                            if annotation.get('/Subtype') == '/Link':
                                if '/A' in annotation and '/URI' in annotation['/A']:
                                    url = annotation['/A']['/URI']
                                    page_urls[page_num].append(url)
                        except:
                            continue
    except Exception as e:
        st.warning(f"Could not extract hyperlinks with PyPDF2: {e}")

    return page_urls

def extract_hyperlinks_pdfplumber(pdf_file):
    """Extract hyperlinks using pdfplumber"""
    page_hyperlinks = defaultdict(list)

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                hyperlinks = page.hyperlinks
                for link in hyperlinks:
                    if 'uri' in link:
                        page_hyperlinks[page_num].append({
                            'url': link['uri'],
                            'bbox': link.get('bbox', None),
                            'top': link.get('top', 0)
                        })
    except Exception as e:
        st.warning(f"Could not extract hyperlinks with pdfplumber: {e}")

    return page_hyperlinks

def extract_media_reprints_with_urls(pdf_file):
    """Extract media outlet names and URLs"""
    results = []
    seen_agencies = set()

    # Reset file position
    pdf_file.seek(0)

    # Extract hyperlinks
    with st.spinner("Extracting hyperlinks from PDF..."):
        hyperlinks_pypdf2 = extract_hyperlinks_pypdf2(pdf_file)
        pdf_file.seek(0)  # Reset for pdfplumber
        hyperlinks_pdfplumber = extract_hyperlinks_pdfplumber(pdf_file)
        pdf_file.seek(0)  # Reset for main processing

    skip_phrases = [
        'Media Reprints:', 'Click on the logos', 'Here is a list',
        'Publishers are not', 'News Databases:', 'Your press release',
        'Note: Paid access', 'Industry Newswire', 'Complete Newswire',
        'World Media Directory', 'Report Summary', 'View tracking',
        'Download PDF', 'Boost your reach', 'Share now on', 'Click below'
    ]

    with st.spinner("Processing PDF content..."):
        with pdfplumber.open(pdf_file) as pdf:
            in_media_section = False

            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)

            for page_num, page in enumerate(pdf.pages):
                progress_bar.progress((page_num + 1) / total_pages)

                text = page.extract_text()
                if not text:
                    continue

                # Get hyperlinks for this page
                page_links_pypdf2 = hyperlinks_pypdf2.get(page_num, [])
                page_links_pdfplumber = hyperlinks_pdfplumber.get(page_num, [])

                # Combine URLs
                all_page_urls = list(page_links_pypdf2)
                all_page_urls.extend([link['url'] for link in page_links_pdfplumber])

                # Remove duplicates
                seen_urls = set()
                unique_urls = []
                for url in all_page_urls:
                    if url not in seen_urls and not url.startswith('mailto:'):
                        seen_urls.add(url)
                        unique_urls.append(url)

                # Check section
                if 'Media Reprints:' in text:
                    in_media_section = True
                    st.info(f"Found Media Reprints section on page {page_num + 1}")

                if any(marker in text for marker in ['News Databases:', 'Industry Newswire']):
                    if in_media_section and 'News Databases:' not in text:
                        break

                if not in_media_section:
                    continue

                # Process lines
                lines = text.split('\n')
                used_urls = set()

                for line in lines:
                    line = line.strip()

                    if not line or any(skip.lower() in line.lower() for skip in skip_phrases):
                        continue

                    if (len(line) > 60 or line.count(' ') > 8 or
                        any(char in line for char in ['@', '%', '&amp;', '...', '|']) or
                        line.startswith('(') or line.startswith('http')):
                        continue

                    upper_count = sum(1 for c in line if c.isupper())
                    if upper_count < 2:
                        continue

                    # Media outlet patterns
                    tv_pattern = r'^[A-Z]{3,5}\s+(?:ABC|NBC|CBS|FOX|CW|MyNetworkTV)\s*\d*'
                    media_pattern = r'^[A-Z][A-Za-z\s\-&\.\']+(?:\s+(?:News|Times|Journal|Post|Daily|Weekly|Today|Online|Report|Observer|Press|Media|Review|Update|Herald|Gazette|Tribune|Monitor))*$'
                    caps_pattern = r'^[A-Z\s\-]+$'

                    if (re.match(tv_pattern, line) or
                        re.match(media_pattern, line) or
                        (re.match(caps_pattern, line) and len(line) < 40)):

                        agency = line.strip()

                        if agency in seen_agencies:
                            continue

                        # Find URL
                        url = "URL_TBD"
                        for candidate_url in unique_urls:
                            if candidate_url not in used_urls:
                                url = candidate_url
                                used_urls.add(candidate_url)
                                break

                        seen_agencies.add(agency)
                        results.append((agency, url))

            progress_bar.empty()

    return results

def create_excel_download(data):
    """Create Excel file and return as bytes"""
    output = io.BytesIO()

    # Create DataFrame from list of tuples
    df = pd.DataFrame(data, columns=['Agency', 'URL'])

    # Create Excel writer with the BytesIO buffer
    with pd.ExcelWriter(output, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, sheet_name='Media Outlets')

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Media Outlets']

        # Adjust column widths
        worksheet.column_dimensions['A'].width = 40
        worksheet.column_dimensions['B'].width = 80

        # Add hyperlinks
        for idx, row in enumerate(df.values, start=2):
            agency, url = row
            if url and url.startswith('http'):
                cell = worksheet.cell(row=idx, column=2)
                cell.hyperlink = url
                cell.style = 'Hyperlink'

    # Get the Excel file bytes
    output.seek(0)
    return output.getvalue()

# Main conversion logic
if uploaded_file is not None:
    st.success(f"File uploaded: {uploaded_file.name}")

    col1, col2 = st.columns([1, 2])

    with col1:
        if st.button("🚀 Convert to Excel", type="primary", use_container_width=True):
            try:
                # Extract data
                data = extract_media_reprints_with_urls(uploaded_file)

                if not data:
                    st.error("❌ No media outlets found in the PDF. Please make sure this is an EIN Presswire Distribution Report.")
                else:
                    # Store results in session state
                    st.session_state['conversion_results'] = data
                    st.session_state['file_name'] = uploaded_file.name

            except Exception as e:
                st.error(f"❌ Error processing PDF: {str(e)}")

# Display results if available
if 'conversion_results' in st.session_state:
    data = st.session_state['conversion_results']
    original_name = st.session_state['file_name']

    # Statistics
    url_count = sum(1 for _, url in data if url != "URL_TBD")

    st.markdown("---")
    st.success(f"✅ Conversion complete!")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Media Outlets", len(data))
    with col2:
        st.metric("URLs Extracted", url_count)
    with col3:
        st.metric("Missing URLs", len(data) - url_count)

    # Preview
    st.markdown("### Preview (first 10 entries)")
    preview_data = []
    for agency, url in data[:10]:
        if url == "URL_TBD":
            preview_data.append({"Agency": agency, "URL": "URL_TBD"})
        else:
            # Truncate long URLs for display
            display_url = url[:50] + "..." if len(url) > 50 else url
            preview_data.append({"Agency": agency, "URL": display_url})

    st.dataframe(preview_data, use_container_width=True)

    if len(data) > 10:
        st.info(f"... and {len(data) - 10} more entries")

    # Download button
    excel_data = create_excel_download(data)
    output_name = Path(original_name).stem + ".xlsx"

    st.download_button(
        label="📥 Download Excel File",
        data=excel_data,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray; font-size: 0.8em;'>
EIN Presswire PDF to Excel Converter<br>
Extracts media outlet names and URLs from distribution reports
</div>
""", unsafe_allow_html=True)
