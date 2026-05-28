# ==========================================
# Word Citation Extractor 
# Version: 1.15.0
# Citation: Pundir, V. (2026, May 28). Word Metadata Extractor Version (1.15.0). Retrieved from https://github.com/accidentalscholar/word-citation-extractor. 
# Citation: RIS and BibTeX files included for referencing software.
# Tested in: Python 3.10.9 64 bit packaged by Anaconda, Inc.
# Reporsitory: https://github.com/accidentalscholar/word-citation-extractor
# Provided under: GNU AFFERO GENERAL PUBLIC LICENSE (see accompanying license file)
# ==========================================

import sys
import subprocess
import os
import re
import urllib.parse
from datetime import datetime
import xml.etree.ElementTree as ET

# --- 1. FAULT TOLERANCE: Auto-Install Missing Libraries ---
def install_and_import():
    required_packages = {
        'docx': 'python-docx',
        'pandas': 'pandas',
        'requests': 'requests',
        'xlsxwriter': 'xlsxwriter',
        'tkinter': 'tk'
    }
    
    for module_name, pip_name in required_packages.items():
        try:
            __import__(module_name)
        except ImportError:
            print(f"[{module_name}] missing. Installing {pip_name}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            except Exception as e:
                print(f"Failed to install {pip_name}. Error: {e}")

install_and_import()

import docx
import pandas as pd
import requests
import xlsxwriter
import tkinter as tk
from tkinter import filedialog

# --- 2. GLOBALS, CACHES & COMPILED REGEX ---
SESSION = requests.Session()
SESSION.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) CitationBot/1.15'})
URL_CACHE = {}
CROSSREF_CACHE = {}

PAREN_PATTERN = re.compile(r'\(([A-Za-z][^()]{1,100})[, ]\s*(\d{4}[a-z]?)[^()]*\)') 
AUTHOR_DATE_PATTERN = re.compile(r'\b([A-Z][A-Za-z\-\'\s\.]{1,100}?)\s\((\d{4}[a-z]?)[^()]*\)') 
BRACKET_PATTERN = re.compile(r'\[(\d+)\]')
URL_PATTERN = re.compile(r'(https?://\S+)')
TITLE_PATTERN = re.compile(r'<title>(.*?)</title>', re.IGNORECASE)

def safe_print(text):
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode('ascii', 'replace').decode('ascii'))

# --- 3. HEURISTICS & HELPERS ---

def get_notes_mapping(doc):
    """Maps footnote/endnote IDs to their text within the Word XML structure."""
    notes = {}
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    try:
        for part in doc.part.package.parts:
            if 'footnotes' in part.partname or 'endnotes' in part.partname:
                root = ET.fromstring(part.blob)
                for note in root.findall('.//w:footnote', nsmap) + root.findall('.//w:endnote', nsmap):
                    note_id = note.get(f'{{{nsmap["w"]}}}id')
                    if note_id in ['0', '-1']: continue 
                    texts = [t.text for t in note.findall('.//w:t', nsmap) if t.text]
                    if texts:
                        notes[note_id] = "".join(texts).strip()
    except Exception:
        pass 
    return notes

def find_best_match(author_part, year_part, refs_data):
    """Scores references to find the best match for the in-text citation."""
    if not refs_data: return ("", "", None)
    
    words = [w.lower() for w in re.findall(r'[A-Za-z]{3,}', author_part)]
    if not words: return ("", "", None)
    
    best_ref = None
    best_score = 0
    clean_author_phrase = re.sub(r'[^a-zA-Z\s]', '', author_part).strip().lower()
    
    for ref_dict in refs_data:
        ref_lower = ref_dict['clean']
        score = 0
        
        if year_part and year_part in ref_lower:
            score += 50
            
        if clean_author_phrase and clean_author_phrase in ref_lower:
            score += 30
            
        word_matches = sum(1 for w in words if w in ref_lower)
        score += (word_matches * 5)
        
        if score > best_score and score >= 55: 
            best_score = score
            best_ref = ref_dict
            
    if best_ref:
        best_ref['matched'] = True
        return best_ref['text'], best_ref['page'], best_ref
        
    return ("", "", None)

def find_best_bracket_match(num_str, refs_data):
    """Matches bracket numbers to the bibliography list."""
    if not refs_data: return ("", "", None)
    patterns_to_check = [f"[{num_str}]", f"{num_str}.", f"{num_str} "]
    for ref_dict in refs_data:
        for p in patterns_to_check:
            if ref_dict['text'].startswith(p):
                ref_dict['matched'] = True
                return ref_dict['text'], ref_dict['page'], ref_dict
    return ("", "", None)

def validate_url(url):
    """Evaluates URLs for direct linking behaviour, 200 OK, and Soft 404s."""
    if url in URL_CACHE: return URL_CACHE[url]
    
    parsed = urllib.parse.urlparse(url)
    netloc_lower = parsed.netloc.lower()
    
    # Direct Link Logic with Search Engine Filter
    path_parts = [p for p in parsed.path.split('/') if p]
    is_direct = "base domain"
    
    search_engines = ['google.', 'bing.com', 'yahoo.', 'duckduckgo.com', 'baidu.com']
    is_search_engine = any(se in netloc_lower for se in search_engines)
    
    if not is_search_engine:
        if parsed.query: 
            is_direct = "direct link"
        elif len(path_parts) > 1:
            is_direct = "direct link"
        elif len(path_parts) == 1:
            if path_parts[0].lower() not in ['index.html', 'index.php', 'index.htm', 'default.aspx', 'home']:
                is_direct = "direct link"
            
    # Status check with DOI specific logic to bypass bot blockers
    status_200, is_404 = "no", "no"
    is_doi = 'doi.org' in netloc_lower
    
    try:
        with SESSION.get(url, timeout=5, stream=True, allow_redirects=not is_doi) as response:
            if response.status_code == 200 or (is_doi and response.status_code in [301, 302, 303, 307, 308]):
                status_200 = "yes"
                
                if response.status_code == 200 and not is_doi:
                    head_content = response.raw.read(4096).decode('utf-8', errors='ignore').lower()
                    title_match = TITLE_PATTERN.search(head_content)
                    if title_match and any(x in title_match.group(1) for x in ["404", "not found", "error"]):
                        is_404 = "yes"
                        
            elif response.status_code == 404:
                is_404 = "yes"
    except requests.RequestException:
        pass
        
    result = (is_direct, status_200, is_404)
    URL_CACHE[url] = result
    return result

def check_real_source(title_or_author):
    """Verifies against Crossref API to see if the academic source exists."""
    query_str = title_or_author[:120].strip()
    if len(query_str) < 10: return "no"
    if query_str in CROSSREF_CACHE: return CROSSREF_CACHE[query_str]
    try:
        quoted_query = urllib.parse.quote(query_str.encode('utf-8'))
        url = f"https://api.crossref.org/works?query={quoted_query}&rows=1&select=title"
        res = SESSION.get(url, timeout=5).json()
        result = "yes" if res.get('message', {}).get('total-results', 0) > 0 else "no"
    except Exception:
        result = "no"
    CROSSREF_CACHE[query_str] = result
    return result

def infer_source_type(reference_text):
    """Categorises the source type using an ordered hierarchy of logic with strict word boundaries."""
    text_lower = reference_text.lower()
    
    # 1. Academic Journal (Now including SSRN, ScienceDirect, and arXiv)
    academic_domains = ['sciencedirect.com', 'ssrn.com', 'arxiv.org']
    if re.search(r'\b(journal|doi\.org|vol\.?)\b', text_lower) or re.search(r'\b\d+\(\d+\)', text_lower) or any(d in text_lower for d in academic_domains):
        return "Academic Journal"
        
    # 2. Report
    if '.pdf' in text_lower or re.search(r'\breport\b', text_lower):
        return "Report"
        
    # 3. Government Document
    if re.search(r'\b(gov|department|ministry|commission|bureau)\b', text_lower) or '.gov.' in text_lower or '.gov/' in text_lower:
        return "Government Document"
        
    # 4. Social Media
    socials = ['youtube', 'twitter', 'x\.com', 'facebook', 'instagram', 'tiktok', 'linkedin', 'reddit', 'weibo', 'wechat', 'vk\.com', 'telegram', 'pinterest']
    social_pattern = r'\b(' + '|'.join(socials) + r')\b'
    if re.search(social_pattern, text_lower):
        return "Social Media"
        
    # 5. Book (Strict word boundaries to prevent URL bleeding)
    book_keywords = ['press', 'publisher', 'books', 'routledge', 'cambridge', 'oxford', 'wiley', 'springer', 'sage', 'macmillan', 'penguin', 'bloomsbury', 'harvard', 'yale', 'mit', 'elsevier', 'harpercollins', 'norton', 'ed\.', 'eds\.', 'isbn']
    book_pattern = r'\b(' + '|'.join(book_keywords) + r')\b'
    if re.search(book_pattern, text_lower):
        return "Book"
        
    # 6. Website
    if 'http' in text_lower or 'www.' in text_lower:
        return "Website"
        
    # 7. Book (Structural Fallback - Demoted below website to protect blog titles)
    if re.search(r'\b[A-Z][a-zA-Z\s]+:\s[A-Z][a-zA-Z\s]+(?:\.|\,|$)', reference_text):
        return "Book"
        
    return "Other/Unknown"

# --- 4. MAIN PARSER ---

def process_docx(filepath):
    doc = docx.Document(filepath)
    filename = os.path.basename(filepath)
    
    page_num = 1
    para_page_map = []
    for p in doc.paragraphs:
        xml = p._element.xml
        breaks = xml.count('w:lastRenderedPageBreak') + xml.count('w:type="page"')
        page_num += breaks
        para_page_map.append(page_num)
        
    para_count = len(doc.paragraphs)
    
    # Locate Bibliography Section
    bib_start_idx = -1
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().lower()
        if txt in ["references", "bibliography", "works cited", "reference list"] and len(txt) < 30:
            bib_start_idx = i
            break
            
    if bib_start_idx == -1:
        for i, p in enumerate(doc.paragraphs):
            style_name = p.style.name.lower() if p.style else ""
            if "bibliography" in style_name or "works cited" in style_name:
                bib_start_idx = i
                break
                
    if bib_start_idx == -1:
        search_start = int(para_count * 0.75)
        for i in range(search_start, para_count):
            txt = doc.paragraphs[i].text.strip()
            if len(txt) > 40 and (re.search(r'\b(19|20)\d{2}\b', txt) or 'http' in txt or re.match(r'^\[\d+\]', txt)):
                bib_start_idx = i
                break
                
    if bib_start_idx == -1:
        bib_start_idx = para_count 
        
    refs_data = []
    for i in range(bib_start_idx, para_count):
        txt = doc.paragraphs[i].text.strip()
        if len(txt) > 10:
            clean_txt = re.sub(r'^\[\d+\]\s*', '', txt).strip().lower()
            refs_data.append({'text': txt, 'page': para_page_map[i], 'clean': clean_txt, 'matched': False})
            
    is_alpha = "N/A"
    if len(refs_data) > 1:
        clean_list = [r['clean'] for r in refs_data]
        is_alpha = "Yes" if clean_list == sorted(clean_list) else "No"
            
    citations_data = []
    footnote_counter = 1
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    notes_dict = get_notes_mapping(doc)
    
    def add_citation_entry(in_text, in_text_page, matched_ref, style, ref_page):
        url_match = URL_PATTERN.search(matched_ref)
        url = url_match.group(1).rstrip('.') if url_match else ""
        
        is_direct, status_200, is_404 = validate_url(url) if url else ("N/A", "N/A", "N/A")
        real_source = check_real_source(matched_ref) if matched_ref else "N/A"
        source_type = infer_source_type(matched_ref) if matched_ref else "Unknown"
        
        # Extract Accessed Date
        date_accessed = ""
        if matched_ref:
            date_match = re.search(r'(?i)(?:accessed|retrieved|viewed)\s*(?:on\s*)?(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4}|\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})', matched_ref)
            if date_match:
                date_accessed = date_match.group(1).strip()
        
        citations_data.append({
            'Column A': filename,
            'Column B': in_text,
            'Column C': in_text_page,
            'Column D': matched_ref,
            'Column E': ref_page,
            'Column F': style,
            'Column G': source_type,
            'Column H': url,
            'Column I': date_accessed,
            'Column J': is_direct,
            'Column K': status_200,
            'Column L': is_404,
            'Column M': real_source
        })

    for i in range(0, bib_start_idx):
        para = doc.paragraphs[i]
        txt = para.text.strip()
        current_page = para_page_map[i]
        if not txt: continue

        for match in PAREN_PATTERN.finditer(txt):
            author, year = match.groups()
            matched_ref, ref_page, _ = find_best_match(author.strip(), year, refs_data)
            display_text = match.group(0)
            add_citation_entry(display_text, current_page, matched_ref, "Harvard/APA", ref_page)

        for match in AUTHOR_DATE_PATTERN.finditer(txt):
            author, year = match.groups()
            matched_ref, ref_page, _ = find_best_match(author.strip(), year, refs_data)
            display_text = match.group(0)
            add_citation_entry(display_text, current_page, matched_ref, "Harvard/APA", ref_page)

        for match in BRACKET_PATTERN.finditer(txt):
            num_str = match.group(1)
            matched_ref, ref_page, _ = find_best_bracket_match(num_str, refs_data)
            add_citation_entry(f"[{num_str}]", current_page, matched_ref, "Numbered", ref_page)

        for run in para.runs:
            refs = run._r.findall('.//w:footnoteReference', nsmap) + run._r.findall('.//w:endnoteReference', nsmap)
            for ref in refs:
                note_id = ref.get(f'{{{nsmap["w"]}}}id')
                if note_id in notes_dict:
                    note_text = notes_dict[note_id]
                    add_citation_entry(f"Footnote [{footnote_counter}]", current_page, note_text, "Footnote/Note Style", current_page)
                    footnote_counter += 1
                    
    for ref in refs_data:
        if not ref['matched']:
            add_citation_entry("None found", "N/A", ref['text'], "Unknown", ref['page'])
            
    doc_stats = {
        'Filename': filename,
        'Style': citations_data[0]['Column F'] if citations_data and citations_data[0]['Column F'] != "Unknown" else "Unknown",
        'Unique_References': len(refs_data), 
        'Data_Rows': len(citations_data),
        'Alphabetical': is_alpha
    }
            
    return citations_data, doc_stats

# --- 5. EXCEL EXPORT ---

def generate_excel(folder_path):
    safe_print("Scanning folder...")
    docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.docx') and not f.startswith('~')]
    
    all_citations = []
    summary_data = []
    
    for f in docx_files:
        safe_print(f"Processing {os.path.basename(f)}...")
        cit_data, stats = process_docx(f)
        all_citations.extend(cit_data)
        summary_data.append(stats)
        
    output_path = os.path.join(folder_path, f"Citation_Analysis_v1.15.0_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book
    
    # Colour format for warnings
    format_orange = workbook.add_format({'bg_color': '#FFD8B2', 'font_color': '#9C0006'})
    format_percent = workbook.add_format({'num_format': '0.0%'})
    format_bold = workbook.add_format({'bold': True})
    
    # 1. Summary Sheet
    worksheet_summary = workbook.add_worksheet("Summary")
    headers_summary = ["Filename", "Inferred referencing style", "Total unique references", 
                       "% In-text matched to ref", "% Ref matched to in-text", "% Real sources (API)", 
                       "% Refs with URL", "% Direct Links", "% HTTP 200", "% 200 and NOT 404", 
                       "Alphabetical References?"]
                       
    worksheet_summary.write_row('A1', headers_summary)
    
    for row_num, stats in enumerate(summary_data, 1):
        fname = stats['Filename']
        sheet_name = fname[:31] 
        s_ref = f"'{sheet_name}'!"
        max_row = max(2, stats['Data_Rows'] + 1)
        
        # Bounded Ranges updated for the new column layout
        rng_B = f"{s_ref}B2:B{max_row}" # In-text
        rng_D = f"{s_ref}D2:D{max_row}" # Full ref
        rng_H = f"{s_ref}H2:H{max_row}" # URL
        rng_J = f"{s_ref}J2:J{max_row}" # Direct link
        rng_K = f"{s_ref}K2:K{max_row}" # HTTP 200
        rng_L = f"{s_ref}L2:L{max_row}" # 404 Page
        rng_M = f"{s_ref}M2:M{max_row}" # Real API Source
        
        worksheet_summary.write_string(row_num, 0, fname)
        worksheet_summary.write_string(row_num, 1, stats['Style'])
        worksheet_summary.write_number(row_num, 2, stats['Unique_References'])
        
        worksheet_summary.write_formula(row_num, 3, f'=IFERROR(COUNTIFS({rng_B}, "<>None found", {rng_D}, "?*") / COUNTIF({rng_B}, "<>None found"), 0)', format_percent)
        worksheet_summary.write_formula(row_num, 4, f'=IFERROR(ROWS(UNIQUE(FILTER({rng_D}, ({rng_B}<>"None found")*({rng_D}<>"")))) / MAX(1, {stats["Unique_References"]}), 0)', format_percent)
        
        worksheet_summary.write_formula(row_num, 5, f'=IFERROR(COUNTIF({rng_M},"yes")/COUNTIFS({rng_M},"?*"),0)', format_percent) 
        worksheet_summary.write_formula(row_num, 6, f'=IFERROR(COUNTIF({rng_H},"?*")/COUNTIFS({rng_D},"?*"),0)', format_percent) 
        worksheet_summary.write_formula(row_num, 7, f'=IFERROR(COUNTIF({rng_J},"direct link")/COUNTIFS({rng_H},"?*"),0)', format_percent) 
        worksheet_summary.write_formula(row_num, 8, f'=IFERROR(COUNTIF({rng_K},"yes")/COUNTIFS({rng_H},"?*"),0)', format_percent) 
        worksheet_summary.write_formula(row_num, 9, f'=IFERROR(COUNTIFS({rng_K},"yes",{rng_L},"no")/COUNTIFS({rng_H},"?*"),0)', format_percent) 
        
        worksheet_summary.write_string(row_num, 10, stats['Alphabetical'])

    worksheet_summary.add_table(0, 0, len(summary_data) if summary_data else 1, len(headers_summary)-1, 
                                {'columns': [{'header': h} for h in headers_summary], 'style': 'Table Style Medium 9'})

    # 2. Individual File Sheets & Excel Native Charts
    df_all = pd.DataFrame(all_citations)
    headers_file = ["Filename", "In text citation", "Estimated page in-text", "Matched full reference", "Estimated page reference", 
                    "Inferred style", "Source type", "Source URL", "Date accessed", "Direct link?", "200 OK?", "404 Page?", "Real Source (API)?"]

    for f in docx_files:
        fname = os.path.basename(f)
        sheet_name = fname[:31] 
        
        if not df_all.empty:
            df_file = df_all[df_all['Column A'] == fname]
        else:
            df_file = pd.DataFrame(columns=df_all.columns if not df_all.empty else headers_file)
            
        df_file.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=1)
        worksheet = writer.sheets[sheet_name]
        
        num_rows = len(df_file) if len(df_file) > 0 else 1
        worksheet.add_table(0, 0, num_rows, len(headers_file)-1, 
                            {'columns': [{'header': h} for h in headers_file], 'style': 'Table Style Medium 2'})
                            
        # Conditional formatting logic indices applied
        worksheet.conditional_format(1, 9, num_rows, 9, {'type': 'cell', 'criteria': '!=', 'value': '"direct link"', 'format': format_orange})
        worksheet.conditional_format(1, 10, num_rows, 10, {'type': 'cell', 'criteria': '==', 'value': '"no"', 'format': format_orange})
        worksheet.conditional_format(1, 11, num_rows, 11, {'type': 'cell', 'criteria': '==', 'value': '"yes"', 'format': format_orange})
        worksheet.conditional_format(1, 12, num_rows, 12, {'type': 'cell', 'criteria': '==', 'value': '"no"', 'format': format_orange})

        # Generate Native Excel Chart
        source_types = df_file['Column G'].tolist() if 'Column G' in df_file else []
        if source_types:
            counts = pd.Series(source_types).value_counts()
            if not counts.empty:
                summary_start_row = 1
                cat_col = 14 # Column O
                val_col = 15 # Column P
                
                worksheet.write_string(0, cat_col, "Source Type", format_bold)
                worksheet.write_string(0, val_col, "Count", format_bold)
                
                for row_offset, (cat, count) in enumerate(counts.items()):
                    worksheet.write_string(summary_start_row + row_offset, cat_col, str(cat))
                    worksheet.write_number(summary_start_row + row_offset, val_col, count)
                    
                chart = workbook.add_chart({'type': 'column'})
                
                chart.add_series({
                    'name': 'Source Types',
                    'categories': [sheet_name, summary_start_row, cat_col, summary_start_row + len(counts) - 1, cat_col],
                    'values':     [sheet_name, summary_start_row, val_col, summary_start_row + len(counts) - 1, val_col],
                    'fill':       {'color': '#87CEEB'},
                    'border':     {'color': '#5F9EA0'}
                })
                
                chart.set_title({'name': f'Source Types Distribution'})
                chart.set_x_axis({'name': 'Source Type'})
                chart.set_y_axis({'name': 'Frequency', 'major_gridlines': {'visible': True}})
                chart.set_legend({'none': True})
                
                worksheet.insert_chart('R2', chart)

    writer.close()
    safe_print(f"\nDone! Excel file saved as:\n{output_path}")

# --- 6. EXECUTION ---
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() 
    
    safe_print("Please select the folder containing your .docx files...")
    folder_selected = filedialog.askdirectory(title="Select Folder with .docx Files")
    
    if folder_selected:
        generate_excel(folder_selected)
    else:
        safe_print("No folder selected. Exiting.")