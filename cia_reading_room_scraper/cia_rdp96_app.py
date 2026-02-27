from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
import time
import re
import json
from datetime import datetime

def setup_directories():
    """Create directory structure"""
    directories = [
        "output/layer_0",
        "output/layer_1", 
        "output/page_urls"
    ]
    for directory in directories:
        os.makedirs(directory, exist_ok=True)

def scrape_search_pages():
    """Layer 0: Scrape search result pages"""
    driver = webdriver.Chrome(options=Options())
    
    page_no = 4092
    while True:
        url = f"https://www.cia.gov/readingroom/search/site/cia%20rdp96?page={page_no}"
        
        try:
            driver.get(url)
            time.sleep(2)
            
            with open(f"output/layer_0/page_{page_no}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            
            print(f"Layer 0: Saved page {page_no}")
            page_no += 1
            
        except:
            print(f"Layer 0: Error on page {page_no}, stopping")
            break
    
    driver.quit()

def extract_document_urls():
    """Extract document URLs from Layer 0 files"""
    urls = []
    
    for filename in os.listdir("output/layer_0"):
        if filename.endswith(".html"):
            filepath = os.path.join("output/layer_0", filename)
            
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()
            
            matches = re.findall(r'a href="(https://www\.cia\.gov/readingroom/document/cia-rdp96-.*?)"', content)
            urls.extend(matches)
    
    urls = list(set(urls))
    
    # Save URLs list
    with open("output/page_urls/layer_1_urls.txt", "w", encoding="utf-8") as f:
        for url in urls:
            f.write(url + "\n")
    
    print(f"Extracted {len(urls)} unique document URLs")
    return urls

def scrape_document_pages(urls):
    """Layer 1: Scrape individual document pages"""
    driver = webdriver.Chrome(options=Options())
    
    # Check what's already downloaded
    downloaded = set()
    for filename in os.listdir("output/layer_1"):
        if filename.endswith(".html"):
            downloaded.add(filename.replace(".html", ""))
    
    print(f"Found {len(downloaded)} already downloaded documents")
    
    # Filter out already downloaded URLs
    urls_to_scrape = []
    for url in urls:
        doc_id = url.split('/')[-1]
        if doc_id not in downloaded:
            urls_to_scrape.append(url)
    
    print(f"Scraping {len(urls_to_scrape)} new documents (skipping {len(urls) - len(urls_to_scrape)})")
    
    for i, url in enumerate(urls_to_scrape):
        try:
            driver.get(url)
            time.sleep(2)
            
            doc_id = url.split('/')[-1]
            
            with open(f"output/layer_1/{doc_id}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            
            print(f"Layer 1: Saved {doc_id} ({i+1}/{len(urls_to_scrape)})")
            
        except Exception as e:
            print(f"Layer 1: Error scraping {url}: {e}")
    
    driver.quit()


def create_cia_json():
    """Create cia_rdp96.json from Layer 1 pages"""
    
    documents = {}
    
    html_files = [f for f in os.listdir("output/layer_1") if f.endswith(".html")]
    total_files = len(html_files)
    
    print(f"Processing {total_files} documents...")
    
    for idx, filename in enumerate(html_files, 1):
        filepath = os.path.join("output/layer_1", filename)
        doc_id = filename.replace(".html", "")
        
        if idx % 100 == 0 or idx == total_files:
            print(f"  Processed {idx}/{total_files} documents")
        
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        
        # Extract both dates
        dates = []
        
        # 1. Document Creation Date
        creation_match = re.search(r'Document Creation Date:[^>]*>.*?content="([^"]+?)"', content, re.DOTALL)
        if creation_match:
            dates.append(creation_match.group(1))
        
        # 2. Document Release Date  
        release_match = re.search(r'Document Release Date:[^>]*>.*?content="([^"]+?)"', content, re.DOTALL)
        if release_match:
            dates.append(release_match.group(1))
        
        # Find earliest date
        earliest_date = ""
        earliest_dt = None
        
        for date_str in dates:
            try:
                # Extract just YYYY-MM-DD part
                date_part = date_str.split('T')[0]
                dt = datetime.strptime(date_part, "%Y-%m-%d")
                
                if earliest_dt is None or dt < earliest_dt:
                    earliest_dt = dt
                    earliest_date = date_part
            except:
                continue
        
        doc_date = earliest_date
        
        if not doc_date:
            print(f"  Warning: No date found for {doc_id}")
        
        # Extract Title
        title_match = re.search(r'<title>(.*?)</title>', content)
        title = title_match.group(1).strip() if title_match else ""
        
        # Extract Body
        body_sections = re.findall(r'Body:[^>]*>(.*?)</div>', content, re.DOTALL)
        body_match = max(body_sections, key=len) if body_sections else ""
        
        if not body_match:
            body_match_search = re.search(r'<div class="field-name-body"[^>]*>.*?<div class="field-item[^>]*>(.*?)</div>', content, re.DOTALL)
            body_match = body_match_search.group(1) if body_match_search else ""
        
        body = body_match.strip() if body_match else ""
        body = re.sub(r'<[^>]+>', '', body)
        body = re.sub(r'\s+', ' ', body)
        
        # Add to documents
        if doc_date not in documents:
            documents[doc_date] = {}
        
        documents[doc_date][doc_id] = {
            "title": title,
            "body": body
        }
    
    with open("output/cia_rdp96.json", "w", encoding="utf-8") as f:
        json.dump(documents, f, indent=2, ensure_ascii=False)
    
    print(f"\nCreated cia_rdp96.json with {len(documents)} dates")
    
    return documents


if __name__ == "__main__":
    # Setup directory structure
    # setup_directories()
    # scrape_search_pages()
    
    document_urls = extract_document_urls()
    
    if document_urls:
        # Layer 1: Scrape document pages
        scrape_document_pages(document_urls)
        print(f"\nLayer 1 complete: {len(document_urls)} documents scraped")
        
        # Create JSON
        create_cia_json()
    else:
        print("\nNo document URLs found to scrape")