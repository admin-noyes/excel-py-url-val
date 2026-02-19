import pandas as pd
import requests
from urllib.parse import urlparse
import re
from pathlib import Path
from datetime import datetime

def is_valid_url(string):
    """Check if a string is a valid URL"""
    url_pattern = r'https?://[^\s]+'
    return re.match(url_pattern, string) is not None

def check_url_status(url, timeout=5):
    """Check the HTTP status of a URL"""
    try:
        response = requests.head(url, timeout=timeout, allow_redirects=True)
        return response.status_code
    except requests.ConnectionError:
        return "Connection Error"
    except requests.Timeout:
        return "Timeout"
    except requests.RequestException as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error: {str(e)}"

def is_broken_url(status_code):
    """Determine if a URL is broken based on status code"""
    if isinstance(status_code, int):
        return status_code >= 400
    return True

def scan_excel_for_urls(excel_file):
    """Scan Excel file for URLs and return list of found URLs with their locations"""
    try:
        xls = pd.ExcelFile(excel_file)
        urls_found = []
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
            
            for row_idx, row in df.iterrows():
                for col_idx, cell_value in enumerate(row):
                    if pd.notna(cell_value) and is_valid_url(str(cell_value)):
                        url = str(cell_value).strip()
                        urls_found.append({
                            'URL': url,
                            'Sheet': sheet_name,
                            'Row': row_idx + 2,  # +2 because pandas adds 1 and Excel headers start at row 1
                            'Column': df.columns[col_idx]
                        })
        
        return urls_found
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def main():
    # Specify your Excel file path here
    excel_file = input("Enter the path to your Excel file: ").strip()
    
    if not Path(excel_file).exists():
        print(f"File not found: {excel_file}")
        return
    
    print(f"\nScanning {excel_file} for URLs...")
    urls_found = scan_excel_for_urls(excel_file)
    
    if not urls_found:
        print("No URLs found in the Excel file.")
        return
    
    print(f"Found {len(urls_found)} URL(s). Checking status...")
    
    broken_urls = []
    
    for i, url_info in enumerate(urls_found, 1):
        url = url_info['URL']
        status = check_url_status(url)
        
        print(f"[{i}/{len(urls_found)}] {url} - Status: {status}")
        
        if is_broken_url(status):
            url_info['Status Code'] = status
            broken_urls.append(url_info)
    
    if not broken_urls:
        print("\nNo broken URLs found!")
        return
    
    # Create a DataFrame with broken URLs
    broken_df = pd.DataFrame(broken_urls)
    
    # Generate output file name
    output_file = excel_file.replace('.xlsx', '_broken_urls.xlsx').replace('.xls', '_broken_urls.xlsx')
    
    # Write to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        broken_df.to_excel(writer, sheet_name='Broken URLs', index=False)
    
    print(f"\n✓ Report generated successfully!")
    print(f"Found {len(broken_urls)} broken URL(s)")
    print(f"Output file: {output_file}")

if __name__ == "__main__":
    main()
