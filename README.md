# LinkedIn Applied Jobs Scraper (Excel + Monthly PDFs)

Scrape your **LinkedIn “Applied Jobs”** history, export a clean **Excel proof log**, and generate **PDF snapshots** organized into **monthly folders** - useful as supporting documentation for **MOE Tuition Grant Bond proof of job applications (2025–2026)**.

> **What this script does**
- Opens LinkedIn in Chrome and asks you to **log in manually** (supports 2FA).
- Iterates through your **Applied Jobs pages** in **descending order** (e.g., page 37 → page 1).
- For each applied job, extracts:
  - **Company**
  - **Role**
  - **Application date** (computed from LinkedIn’s relative text like “11 months ago”)
  - **About the job** / job description (short preview for Excel + full text saved separately)
  - **Job URL**
- Saves:
  - `output/applications_2025_2026.xlsx` (clean columns for MOE proof)
  - `output/applications_full_descriptions.xlsx` (same rows + full description text)
  - `output/pdfs/YYYY-MM/*.pdf` (PDF snapshots grouped by application month)

## Why this exists (MOE Tuition Grant proof)

LinkedIn doesn’t show exact application dates — it shows relative strings like:

- `Application submitted 11 months ago`
- `3w ago`, `2d ago`, `1yr ago`

This script converts those relative strings into an **exact YYYY-MM-DD date** using a reference “current date” (configured inside the script), then uses that to:
- **log the application date in Excel**
- **place PDFs inside monthly folders**, e.g. `output/pdfs/2025-02/`

## Folder output structure

After running, your project will look like:

project/
├── scraper.py (main script)
├── requirements.txt
├── output/
│   ├── applications_2025_2026.xlsx
│   └── pdfs/
│       ├── 2025-01/
│       ├── 2025-02/
│       ├── ... (auto-created per month)

## Requirements

### 1) Python
Recommended: **Python 3.10+**

### 2) Chrome Browser
You need Google Chrome installed.

### 3) ChromeDriver
Selenium needs a ChromeDriver that matches your Chrome version.
- Easiest path: use Selenium Manager (newer Selenium versions can auto-handle this)
- If it fails, install ChromeDriver manually and ensure it’s on your PATH.

### 4) Python packages
Install dependencies:

 - pip install selenium pandas openpyxl python-dateutil

## How to run

1.  Save your script as:

linkedin_job_scraper.py

2.  Run:

python linkedin_job_scraper.py

3.  A Chrome window opens → **log into LinkedIn manually**.
    
4.  Return to your terminal and press **ENTER** to start scraping.

## Configuration

At the bottom of the script:

    START_PAGE = 37
    END_PAGE = 1
    scraper = LinkedInJobScraper(start_page=START_PAGE, end_page=END_PAGE)
    scraper.run()

-   If LinkedIn changes the number of applied-job pages, update `START_PAGE`.
-   Default logic scrapes pages in descending order: `37 → 1`.

### Current date used for calculations

Inside `__init__`:
`self.current_date = datetime(2026, 1, 3, 16, 0) # +08 reference time`

This “reference now” is what converts `11 months ago` into a real date.

> If you want it to always use _real current time_ instead of a fixed reference, you can replace it with:

    self.current_date = datetime.now()

(Keeping it fixed can be useful for audit consistency.)

## What “Application Date” means in the Excel

LinkedIn shows relative time (“11 months ago”), not the true application timestamp.

This script:

1.  Extracts the relative string

2.  Converts it using:
    -   months → `relativedelta(months=n)`
    -   years → `relativedelta(years=n)`
    -   weeks → `relativedelta(weeks=n)`
    -   days → `relativedelta(days=n)`
        
3.  Saves as:

-   **Excel Date:** `YYYY-MM-DD`
-   **Month Folder:** `YYYY-MM`

✅ This is sufficient for “proof-of-effort” documentation, but it is still an approximation based on LinkedIn’s relative display.

## Notes on PDFs

The script uses Chrome DevTools Protocol:

    self.driver.execute_cdp_cmd("Page.printToPDF", pdf_settings)

It saves the **currently open job details page** as a PDF and files it under the computed month.
Filename format:

    {Company}_{Role}.pdf

Special characters are stripped to avoid Windows filename issues.

## Common issues & fixes

### “No job cards found”:

-   LinkedIn sometimes loads slowly or changes the DOM.
-   Try increasing wait time in: self.wait = WebDriverWait(self.driver, 10). Change `10` to `15` or `20`.

### LinkedIn layout/CSS selectors changed:

This script relies on selectors like:
-   `li.jobs-item-card-list__item`
-   `h1.t-24`
-   `.jobs-description-content__text`
  
If LinkedIn updates their UI, some fields may show as `Unknown`. Update selectors accordingly.

### Getting blocked/rate limits:

LinkedIn can detect automation. This script tries to reduce risk by:
-  manual login 
- random waits (`wait_random`)  
- no headless mode by default

If you get CAPTCHA or restrictions:

-   slow down delays (increase `min_sec/max_sec`)
-   scrape fewer pages at a time

## Ethical & account safety disclaimer

This tool is for **personal record-keeping** and documentation. LinkedIn’s Terms may restrict automated scraping. Use responsibly, slowly, and only on your own account.

## What to submit to MOE (suggested)

Typically, you’ll want:

-   The Excel file as a summary log (`applications_2025_2026.xlsx`)
-   Monthly PDFs as supporting “screenshots” of the application records
    

Consider also adding:

-   A short cover note explaining that LinkedIn shows relative dates, and you computed exact dates using a fixed reference timestamp.


## Roadmap improvements (optional)

If you want to level this up:

-   Add a **“resume mode”** to skip jobs already captured
    
-   Add a **CSV export** option
    
-   Add screenshot PNG export (lighter than PDFs)
    
-   Add “applied job ID” deduplication using the URL job ID
    
-   Make `current_date` automatically use your timezone (+08 Singapore)

## License

Personal-use script. Adjust licensing text if you plan to share publicly.