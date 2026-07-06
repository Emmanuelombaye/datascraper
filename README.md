# Child Care Centers USA - Web Scraping Project

This repository contains a professional sample project demonstrating the architecture required to scrape child care center data across all 50 USA states and export the data into state-separated Excel sheets.

## Features Demonstrated
1. **State-Separated Excel Export**: Utilizes `pandas` and `openpyxl` to dynamically create a new Excel sheet for every state processed.
2. **Regex Email Extraction**: Includes a robust regular expression method to find and extract email addresses hidden within blocks of HTML text, as requested.
3. **Structured Data Parsing**: Uses a class-based approach `ChildCareScraper` to keep data strictly formatted to the required fields.
4. **Professional Logging**: Uses Python's `logging` module to track progress instead of basic print statements.
5. **Anti-Bot Considerations**: Includes custom User-Agent headers setup.

## Required Fields Handled
- Center Name
- Director/Owner/Principal
- Phone Number
- Email Address (Must)
- Full Address
- City
- State
- ZIP Code
- Capacity

## Setup Instructions

1. Ensure you have Python 3.8+ installed.
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the scraper:
   ```bash
   python scraper.py
   ```

## Output
The script will generate an Excel file named `ChildCare_Centers_USA.xlsx`. When you open it, you will see a separate worksheet tab at the bottom for Texas, California, Illinois, Ohio, and New York, complete with auto-formatted column widths.

*Note: The current `scraper.py` is configured with placeholder dummy data to demonstrate the architecture safely without hammering a real target server. To adapt it to a specific directory, the `parse_listing` logic would be implemented with the specific CSS selectors of the target website using the included BeautifulSoup library.*
