import re, pdfplumber, pandas as pd, tldextract, os, sys

def host_to_agency(host):
    """Very small map for the big names that always appear"""
    MAP = {
        'apnews.com':'AP News',
        'benzinga.com':'Benzinga',
        'menafn.com':'MENAFN',
        'businessdaily.ch':'Business Daily Switzerland',
        'switzerlandtechreview.com':'Switzerland Tech Review',
        'switzerlanddailymonitor.com':'Switzerland Daily Monitor',
        'berndailypress.com':'Bern Daily Press',
        'switzerlandweekly.com':'Switzerland Weekly',
        'swissentertainmentupdate.com':'Swiss Entertainment Update',
        'internationaltechtimes.com':'International Tech Times',
    }
    return MAP.get(host, host.split('.')[-2].upper())

def extract_agency_url_pairs(pdf_path):
    rows = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    # 1) Harvest every raw line that looks like an outlet name
    #    EIN reports always put the list under “Media Reprints”
    start = re.search(r'Media Reprints:', full_text, re.I)
    if not start:
        return rows
    block = full_text[start.end():]

    # Split on figure / span blocks
    chunks = re.split(r'<figure>.*?</figure>', block, flags=re.S)
    for chunk in chunks:
        # remove tags
        chunk = re.sub(r'<[^>]+>', ' ', chunk)
        # split by white-space heavy tokens
        names = re.findall(r'[A-Z][A-Z0-9&\s\-/]+', chunk)
        for raw in names:
            raw = raw.strip()
            if len(raw) > 3 and raw not in {'PDF', 'RSS', 'CSV'}:
                rows.append({'Agency': raw, 'URL': 'URL_TBD'})

    # 2) Try to harvest any *actual* hyperlinks that sometimes survive
    #    (this catches the few that EIN embeds as naked https://...)
    for url in re.findall(r'https?://[^\s<>\)\]]+', full_text):
        host = tldextract.extract(url).registered_domain
        rows.append({'Agency': host_to_agency(host), 'URL': url})

    # 3) De-duplicate while keeping the first URL found per Agency
    df = pd.DataFrame(rows).drop_duplicates(subset='Agency')
    return df

# ------------------- CLI --------------------
if __name__ == '__main__':
    pdf_file = sys.argv[1] if len(sys.argv) > 1 else input('Drag PDF here: ').strip('"')
    out_file = os.path.splitext(pdf_file)[0] + '.xlsx'
    df = extract_agency_url_pairs(pdf_file)
    df.to_excel(out_file, index=False)
    print(f'Done → {out_file}  ({len(df)} rows)')
