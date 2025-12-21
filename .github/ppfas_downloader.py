import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import datetime

# Configuration
BASE_URL = "https://amc.ppfas.com/downloads/portfolio-disclosure/"
DOWNLOAD_DIR = r"D:\Bhardwaj\Antigravity_Projects\MF_Data_Compiler\PPFAS_2025_Disclosures"
TARGET_SCHEAM_NAME = "Parag Parikh Flexi Cap Fund"

def get_monthly_links(year):
    """
    Scrapes the PPFAS portfolio disclosure page and returns a list of dictionaries 
    containing the month, year, and download URL for the target scheme.
    """
    print(f"Fetching {BASE_URL}...")
    try:
        response = requests.get(BASE_URL)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Error fetching URL: {e}")
        return []

    soup = BeautifulSoup(response.content, 'html.parser')
    
    # The page uses tabs for years. The current year (or recent years) are usually available.
    # We look for the "Monthly Portfolio" section.
    # Based on observation, the links are inside accordions for each month.
    
    # We can search for all links that match the scheme name and are within the requested year's context.
    # However, the structure is: Tab(Year) -> Accordion(Month) -> Link(Scheme)
    
    # PPFAS website structure often uses IDs like 'twentyfive' for 2025.
    # Let's try to identify the year tab content.
    
    year_ids = {
        2025: "twentyfive",
        2024: "twentyfour",
        2023: "twentythree"
        # Add more if needed or logic to guess
    }
    
    year_id = year_ids.get(year)
    if not year_id:
        # Fallback: Try to find a div that might correspond to the year or just search globally if structure allows
        print(f"Year ID mapping for {year} not found. Searching globally...")
        year_container = soup # Search everywhere
    else:
        year_container = soup.find(id=year_id)
        if not year_container:
            print(f"Container for year {year} (id='{year_id}') not found.")
            return []

    results = []
    
    # Find all card headers or collapse divs
    # The structure observed:
    # <div id="collapseJanuary2025" ...>
    #   <div class="card-body">
    #      <a title="Parag Parikh Flexi Cap Fund" href="...">...</a>
    
    cards = year_container.find_all(class_="card")
    
    for card in cards:
        # Try to find the month name first
        header = card.find(class_="card-header")
        if not header:
            continue
            
        header_text = header.get_text(strip=True)
        # expected text like "January 2025"
        
        if str(year) not in header_text:
            continue
            
        # Parse month
        month = None
        for m in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']:
            if m in header_text:
                month = m
                break
        
        if not month:
            continue
            
        # Now look for the link inside the collapse body
        # The body usually follows the header
        body = card.find(class_="collapse")
        if not body:
            # Maybe it's not collapsed or structure is slightly different
            body = card
            
        link_tag = body.find('a', title=TARGET_SCHEAM_NAME)
        if not link_tag:
             # Try Partial match on text if title attribute is missing
             link_tag = body.find('a', string=lambda t: t and TARGET_SCHEAM_NAME in t)
             
        if link_tag and link_tag.get('href'):
            full_url = urljoin(BASE_URL, link_tag['href'])
            results.append({
                "month": month,
                "year": year,
                "url": full_url
            })
            
    return results

def download_file(url, filename, folder):
    if not os.path.exists(folder):
        os.makedirs(folder)
        
    filepath = os.path.join(folder, filename)
    
    # Optional: Skip if exists
    # if os.path.exists(filepath):
    #     print(f"Skipping {filename}, already exists.")
    #     return

    print(f"Downloading {filename}...")
    try:
        r = requests.get(url, stream=True)
        r.raise_for_status()
        with open(filepath, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"Saved to {filepath}")
    except Exception as e:
        print(f"Failed to download {filename}: {e}")

def main():
    current_year = 2025 # Or datetime.date.today().year
    print(f"Starting downloader for Year {current_year}...")
    
    links = get_monthly_links(current_year)
    
    if not links:
        print("No links found. Please check the website structure or year.")
        # Fallback debug: print all links in that year container?
        return

    print(f"Found {len(links)} disclosure files.")
    
    for item in links:
        filename = f"{item['month']}_{item['year']}.xls"
        download_file(item['url'], filename, DOWNLOAD_DIR)

if __name__ == "__main__":
    main()
