"""
LinkedIn Applied Jobs Scraper (Robust DOM Version)
Scrapes applied jobs from LinkedIn "My Jobs" → Applied, saves to Excel, and generates PDFs organized by month.

Author: Created for MOE Tuition Grant Bond Proof
Date: Jan 2026
"""

import os
import time
import re
import random
import base64
from dataclasses import dataclass
from datetime import datetime
from dateutil.relativedelta import relativedelta
from zoneinfo import ZoneInfo
import sys
import io
import shutil

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.options import Options

import json
import traceback
import tkinter as tk
from tkinter import messagebox

from selenium.common.exceptions import WebDriverException
from urllib3.exceptions import ReadTimeoutError

# Fix Windows console encoding issues
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class RestartSession(Exception):
    """Raised when the Selenium session is hung/unresponsive and must be restarted."""
    pass

# This is a small container object that represents one job entry found on the Applied Jobs list page (the page that shows ~10 jobs).
@dataclass
class JobListItem:
    url: str # Job detail link to open 
    role_hint: str | None # A “best-effort” role/title captured from the list page (hint because we will use job title from detail page) although both are same
    company_hint: str | None # Same idea as role_hint, but for company
    applied_relative: str | None  # LinkedIn sometimes shows Applied XYZ months … ago” on the list page clearly.
    job_id: str # The stable job ID extracted from the URL (extracted so duplicate jobs not scraped and for unique pdf filenames)


# Defining blueprint for scraper object
class LinkedInJobScraper:
    def __init__(self, start_page=37, end_page=1, chrome_user_data_dir=None, chrome_profile_dir=None):
        self.start_page = start_page
        self.end_page = end_page
        self.base_url = "https://www.linkedin.com/my-items/saved-jobs/?cardType=APPLIED" # This is the base URL and we append &start=<index> for pagination

        # Use "now" in Singapore timezone (matches your requirement)
        self.current_date = datetime.now(ZoneInfo("Asia/Singapore"))

        # Storage for extracted results
        self.data = [] # A list of dictionaries, one per job application
        self.seen_job_ids: set[str] = set() # A set of job IDs you’ve already processed (to avoid duplicates)

        # Output dirs
        self.output_dir = "output"

        # Resume/checkpoint file
        self.state_path = os.path.join(self.output_dir, "scrape_state.json")

        # Selenium profile folder (the "cookie jar" that can get poisoned)
        self.selenium_profile_dir = os.path.abspath("chrome_selenium_profile")

        # Profile health / poisoning strikes:
        # 0 = normal, 1 = failed once (reuse profile once), 2 = failed twice (reset profile)
        self.profile_poison_strikes = 0

        # For resume
        self.last_page = None
        self.last_url = None

        self.pdf_dir = os.path.join(self.output_dir, "pdfs")
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.pdf_dir, exist_ok=True)

        # Store profile config
        self.chrome_user_data_dir = chrome_user_data_dir
        self.chrome_profile_dir = chrome_profile_dir

        # These will hold Selenium objects later
        self.driver = None # Chrome browser automation controller
        self.wait = None # WebDriverWait helper used for reliable waiting

    # -----------------------
    # Browser
    # -----------------------

    # launching Chrome (Selenium WebDriver)
    def setup_driver(self):

        # Create Chrome options object
        chrome_options = Options()

        # Configure Chrome to behave better for scraping (Opens Chrome maximized and reduces weird responsive layouts)
        chrome_options.add_argument("--start-maximized")

        # Reduces some obvious automation signals in Chrome (not a magic bypass, but helps)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")

        # Removes some Selenium automation flags/banners
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)

        # Optional: reduce noisy WebRTC DNS logs
        chrome_options.add_argument("--disable-features=WebRtcHideLocalIpsWithMdns,WebRtcEnableChromeMdns")

        chrome_options.add_argument(f"--user-data-dir={self.selenium_profile_dir}")
        chrome_options.add_argument("--profile-directory=Default") # Pin Selenium to one profile directory

        # Optional: reduce random background noise
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-popup-blocking")

        # Good if your machine is heavy / Chrome unstable:
        # chrome_options.add_argument("--disable-dev-shm-usage")
        # chrome_options.add_argument("--no-sandbox")  # optional; mostly for Linux, usually not needed on Windows


        # launches Chrome and gives Selenium control of it
        self.driver = webdriver.Chrome(options=chrome_options)

        # Creating a smart wait driver. Keep checking for elements up to 15 seconds before timing out.
        self.wait = WebDriverWait(self.driver, 15)

        # Hard timeouts so we don't hang forever on "infinite loading"
        self.driver.set_page_load_timeout(45) # navigation timeout
        self.driver.set_script_timeout(30)  # JS execution timeout




        print("✓ Browser initialized")

    def manual_login(self):

        print("\n" + "=" * 60)
        print("MANUAL LOGIN REQUIRED")
        print("=" * 60)

        # Just navigate to the applied page (may redirect to login).
        # Do NOT require any post_wait element here.
        try:
            self.safe_get(self.base_url, context="manual_login base_url", post_wait_xpath=None)
        except RestartSession as e:
            raise

        input("\nLog in fully (including OTP), then press ENTER here...")

        # After ENTER, now we expect the applied jobs list to exist.
        try:

            self._wait_for_list_page()
            print("✓ Login confirmed (Applied jobs list detected).")
        
        except TimeoutException:

            if self.is_blocked_or_checkpoint():
                self.guard_not_blocked("manual_login: list not detected after ENTER")

            else:
                print("⚠ Applied jobs list not detected after login ENTER.")
                print("   You might still be on a loading screen / wrong page.")
                print("   Current URL:", self.driver.current_url)
                raise


    # -----------------------
    # Helpers
    # -----------------------

    def profile_healthy(self) -> bool:
        """
        Health check right after login:
        - Can we load linkedin feed?
        - Is body present?
        - Not checkpointed?
        If this fails, profile may be poisoned / linkedin stuck.
        """
        try:
            self.safe_get(
                "https://www.linkedin.com/feed/",
                context="profile_healthy(feed)",
                post_wait_xpath="//body",
                timeout_sec=25
            )
            return not self.is_blocked_or_checkpoint()
        except Exception:
            return False


    def reset_profile(self):
        """
        Emergency reset: delete the selenium profile folder.
        Use this ONLY after failure repeats (poison strike #2).
        """
        print("\n🧹 Resetting Selenium profile folder (poisoned profile suspected).")
        print("   Path:", self.selenium_profile_dir)

        # Ensure driver is closed before deleting
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.driver = None
        self.wait = None

        # Delete the whole profile folder
        try:
            shutil.rmtree(self.selenium_profile_dir, ignore_errors=True)
            print("✅ Profile folder removed.")
        except Exception as e:
            print("⚠ Could not remove profile folder:", e)

    def notify_topmost(self, title: str, msg: str):
        """
        Pops a top-most window above all other windows (works on Windows).
        """
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showwarning(title, msg, parent=root)
            root.destroy()
        except Exception:
            # fallback to console if tkinter fails
            print(f"[NOTIFY] {title}: {msg}")


    def save_state(self, reason: str = ""):
        """
        Persist the minimal resume info to disk.
        """
        state = {
            "saved_at_sg": datetime.now(ZoneInfo("Asia/Singapore")).strftime("%Y-%m-%d %H:%M:%S"),
            "reason": reason,
            "last_page": self.last_page,
            "last_url": self.last_url,
            "start_page": self.start_page,
            "end_page": self.end_page,
            "seen_job_ids": sorted(list(self.seen_job_ids)),
            "rows_collected": len(self.data),
        }
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            with open(self.state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, indent=2)
            print(f"💾 State saved: {os.path.abspath(self.state_path)} ({reason})")
        except Exception as e:
            print(f"⚠ Failed to save state: {e}")


    def load_state(self):
        """
        Load resume info if it exists.
        """
        if not os.path.exists(self.state_path):
            return False

        try:
            with open(self.state_path, "r", encoding="utf-8") as f:
                state = json.load(f)

            self.last_page = state.get("last_page")
            self.last_url = state.get("last_url")

            seen = state.get("seen_job_ids", [])
            self.seen_job_ids = set(seen)

            print("↩ Resume state loaded:")
            print("   last_page:", self.last_page)
            print("   rows_collected:", state.get("rows_collected"))
            print("   seen_job_ids:", len(self.seen_job_ids))
            return True
        except Exception as e:
            print(f"⚠ Failed to load state: {e}")
            return False

    def get_resume_start_page(self) -> int:
        """
        If state says we were last working on page N, resume from that page.
        Otherwise resume from self.start_page.
        """
        if isinstance(self.last_page, int) and self.end_page <= self.last_page <= self.start_page:
            return self.last_page
        return self.start_page

    # Create variable wait timings to make it look human-like
    def wait_random(self, min_sec=1.5, max_sec=3.5):
        time.sleep(random.uniform(min_sec, max_sec))

    # Some actions need tiny pauses, not multi-second waits like scrolling page, expanding buttons etc. 
    def micro_pause(self):
        time.sleep(random.uniform(0.15, 0.45))


    def is_blocked_or_checkpoint(self) -> bool:
        """
        Conservative detection to avoid false positives.
        Only triggers on strong signals:
        - URL contains checkpoint/challenge
        - known checkpoint DOM elements exist
        """
        url = (self.driver.current_url or "").lower()
        title = (self.driver.title or "").lower()

        # Strong URL signals
        if "/checkpoint/" in url or "/challenge/" in url:
            return True

        # Strong title signals (keep tight)
        if "linkedin security verification" in title or "checkpoint" in title:
            return True

        # Strong DOM signals (avoid scanning whole page text)
        strong_selectors = [
            (By.CSS_SELECTOR, "input[name='pin']"),                 # OTP / pin
            (By.CSS_SELECTOR, "input[name='challengeId']"),
            (By.CSS_SELECTOR, "div#captcha-internal"),              # sometimes used
            (By.CSS_SELECTOR, "div.recaptcha"),                     # recaptcha containers
            (By.XPATH, "//*[contains(., 'unusual activity')]"),
            (By.XPATH, "//*[contains(., 'security verification')]"),
            (By.XPATH, "//*[contains(., 'prove you are')]"),
        ]

        for by, sel in strong_selectors:
            try:
                els = self.driver.find_elements(by, sel)
                if els:
                    return True
            except Exception:
                pass

        return False
    
    def guard_not_blocked(self, context: str = ""):
        if not self.is_blocked_or_checkpoint():
            return

        print("\n🛑 LinkedIn checkpoint/blocked detected.")
        if context:
            print("   Context:", context)
        print("   Current URL:", self.driver.current_url)

        self.last_url = self.driver.current_url

        try:
            self.save_to_excel()
        except Exception as e:
            print("   ⚠ Could not save Excel while blocked:", e)

        self.save_state(reason=f"blocked: {context}")

        self.notify_topmost(
            "LinkedIn login needed",
            "LinkedIn triggered a checkpoint / refresh loop.\n\n"
            "I saved your Excel + resume state.\n"
            "Please log in again in the Selenium Chrome window, then rerun/restart."
        )

        raise RuntimeError("LinkedIn session blocked/checkpointed")

    # LinkedIn paginates jobs like this:
    # Page 1: ...?cardType=APPLIED
    # Page 2: ...?cardType=APPLIED&start=10
    # Page 3: ...?cardType=APPLIED&start=20
    # Function to get full LinkedIn URL for a given page number
    def get_page_url(self, page_number: int) -> str:
        if page_number == 1:
            return self.base_url
        start_index = (page_number - 1) * 10
        return f"{self.base_url}&start={start_index}"


    # Extract stable job ID from job URL
    # Example URL: https://www.linkedin.com/jobs/view/3928472391/
    # Job ID: 3928472391
    def extract_job_id_from_url(self, url: str) -> str:
        m = re.search(r"/jobs/view/(\d+)", url)
        return m.group(1) if m else "unknown"


    # Parse applied-relative time strings into actual dates
    # LinkedIn never shows exact dates for applications
    # Converts strings like "Applied 11mo ago" into actual date strings
    def parse_application_date(self, text: str):
        """
        Accepts strings like:
        - "Application submitted 11 months ago"
        - "Applied 11mo ago"
        - "Applied 2w ago"
        - "1 yr ago"
        """
        if not text:
            return None, None

        t = text.lower().strip()

        # normalize common variants
        t = t.replace("months", "mo").replace("month", "mo")
        t = t.replace("years", "yr").replace("year", "yr")
        t = t.replace("weeks", "w").replace("week", "w")
        t = t.replace("days", "d").replace("day", "d")

        m = re.search(r"(\d+)\s*(mo|yr|w|d)", t)
        if not m:
            return None, None

        amount = int(m.group(1))
        unit = m.group(2)

        if unit == "mo":
            dt = self.current_date - relativedelta(months=amount)
        elif unit == "yr":
            dt = self.current_date - relativedelta(years=amount)
        elif unit == "w":
            dt = self.current_date - relativedelta(weeks=amount)
        elif unit == "d":
            dt = self.current_date - relativedelta(days=amount)
        else:
            return None, None

        return dt.strftime("%Y-%m-%d"), dt.strftime("%m-%Y")

    # -----------------------
    # List-page scraping
    # -----------------------
    def _wait_for_list_page(self):
        self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class,'linked-area')]//a[contains(@href, '/jobs/view/')]")
            )
        )

    def get_jobs_from_list_page(self) -> list[JobListItem]:
        """
        Collect job links and (best-effort) role/company/applied relative time from the list page.
        We dedupe by job_id to avoid repeats.
        """

        # Make sure the page is ready
        self._wait_for_list_page()

        # Add a human-like delay
        self.wait_random(1.5, 2.5)  

        # Find if link’s href contains /jobs/view/ inside a div with class linked-area
        anchors = self.driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'linked-area')]//a[contains(@href, '/jobs/view/')]"
        )

        # Store each JobListItem in an output list
        items: list[JobListItem] = []


        for a in anchors:
            try:
                # Get the URL from the hrefs
                url = a.get_attribute("href")
                if not url:
                    continue

                job_id = self.extract_job_id_from_url(url)

                # Skip the job id if it has been parsed through already (no duplicates)
                if job_id == "unknown" or job_id in self.seen_job_ids:
                    continue

                # Go upward until you find the nearest div whose class contains linked-area
                try:
                    card = a.find_element(By.XPATH, "ancestor::div[contains(@class,'linked-area')][1]")

                except NoSuchElementException:
                    card = a
                    for _ in range(6):
                        card = card.find_element(By.XPATH, "..")

                # Extract Role Hint by finding the div with classes t-roman t-sans inside the card and reading the text
                role_hint = None
                try:
                    title_div = card.find_element(By.CSS_SELECTOR, "div.t-roman.t-sans")
                    role_hint = (title_div.text or "").strip() or None
                except NoSuchElementException:
                    role_hint = (a.text or "").strip() or None

                # Extract company hint using “next sibling after title div”
                company_hint = None
                try:
                    # locate the title block first
                    title_div = card.find_element(By.CSS_SELECTOR, "div.t-roman.t-sans")

                    # company is typically the very next sibling with t-14 t-black t-normal so we extract using that pattern
                    company_el = title_div.find_element(
                        By.XPATH,
                        "following-sibling::div[contains(@class,'t-14') and contains(@class,'t-black') and contains(@class,'t-normal')][1]"
                    )
                    company_hint = (company_el.text or "").strip() or None

                except NoSuchElementException:
                    company_hint = None

                # Extract applied relative time from the card text
                card_text = (card.text or "").strip()
                applied_rel = None
                m = re.search(
                    r"(\d+\s*(?:mo|yr|w|d)\s*ago|\d+\s*(?:months|years|weeks|days)\s*ago)",
                    card_text,
                    re.I
                )
                if m:
                    applied_rel = m.group(1)

                # Fallback mechanism & safety net: Company extraction if DOM-based company isn't found
                if not company_hint:
                    lines = [ln.strip() for ln in card_text.split("\n") if ln.strip()]
                    for ln in lines:
                        if role_hint and ln == role_hint:
                            continue
                        if re.search(r"\bago\b", ln, re.I):
                            continue
                        if re.search(r"\b(singapore|remote|hybrid|full-time|part-time|contract)\b", ln, re.I):
                            continue
                        company_hint = ln
                        break
                            
                # Create JobListItem object and store it
                items.append(JobListItem(
                    url=url,
                    role_hint=role_hint,
                    company_hint=company_hint,
                    applied_relative=applied_rel,
                    job_id=job_id
                ))

                # Mark job as processed and added to seen set   
                self.seen_job_ids.add(job_id)

            except (StaleElementReferenceException, NoSuchElementException):
                continue

        return items

    # -----------------------
    # Job detail scraping
    # -----------------------
    def _expand_description_if_possible(self):
        """
        Newer LinkedIn layout often uses jobs-description__footer-button to expand/collapse.
        """
        try:
            # Looks for the see more button in the about job section
            btn = self.driver.find_element(By.CSS_SELECTOR, "button.jobs-description__footer-button")
            aria = btn.get_attribute("aria-expanded")

            # If collapsed, clicking expands. If already expanded, do nothing.
            if aria == "false":
                btn.click()
                self.wait_random(0.8, 1.4)

        except NoSuchElementException:
            # Sometimes there is no see more button but instead it literally contains the text, "See More" 
            try:
                btn2 = self.driver.find_element(By.XPATH, "//button[contains(., 'See more')]")
                btn2.click()
                self.wait_random(0.8, 1.4)
            except Exception:
                pass
                
    # Main function to scrape job details from job detail page
    def scrape_job_details(self, job: JobListItem):
            
        # Open the job detail page
        self.safe_get(
            job.url,
            context=f"scrape_job_details(job_id={job.job_id})",
            post_wait_xpath="//h1"
        )

        # NEW
        self.guard_not_blocked(f"scrape_job_details(job_id={job.job_id}) after driver.get")

        # Wait for the title to load (but don’t die if it doesn’t)
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.t-24.t-bold.inline")))

        except TimeoutException:
            pass

        self.wait_random(0.6, 1.2)


        # Extract the job title (“Role”)
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.t-24.t-bold.inline")))
            role = self.driver.find_element(By.CSS_SELECTOR, "h1.t-24.t-bold.inline").text.strip()
        except TimeoutException:
            role = job.role_hint or "Unknown Role"

        # Extract the company name (using multiple selectors)
        company = job.company_hint or "Unknown Company"
        for sel in [
            "div.job-details-jobs-unified-top-card__company-name a",
            "a.app-aware-link[href*='/company/']",
            "span.job-details-jobs-unified-top-card__company-name"
        ]:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                txt = el.text.strip()
                if txt:
                    company = txt
                    break
            except NoSuchElementException:
                continue

        # Applied-relative time: prefer list hint (more consistent)
        applied_text = job.applied_relative or ""

        # If not found, attempt on detail page
        if not applied_text:
            try:
                # look for text with "Application submitted"
                el = self.driver.find_element(By.XPATH, "//*[contains(., 'Application submitted')]")
                applied_text = el.text.strip()
            except NoSuchElementException:
                applied_text = "Unknown"

        # Convert relative time to actual date + folder month (02-2025)
        applied_date, folder_month = self.parse_application_date(applied_text)

        # First expand description
        self._expand_description_if_possible()
        job_description = "Description not available"

        # Try multiple likely containers
        description_selectors = [
            (By.CSS_SELECTOR, "div.jobs-description__content"),
            (By.CSS_SELECTOR, "div.jobs-description-content__text"),
            (By.CSS_SELECTOR, "div.jobs-box__html-content"),
            (By.XPATH, "//article[contains(@class,'jobs-description')]"),
        ]


        for by, sel in description_selectors:
            try:
                el = self.driver.find_element(by, sel)
                txt = (el.text or "").strip()
                # Make sure full job description is selected and it's not some tiny snippet
                if txt and len(txt) > 50:
                    job_description = txt
                    break
            except NoSuchElementException:
                continue

        pdf_filename = None

        # Only save the PDF if we successfully parsed the application month.
        if folder_month:
            pdf_filename = self.save_as_pdf(company, role, folder_month, job.job_id)

        return {
            "Company": company,
            "Role": role,
            "Application Date": applied_date if applied_date else applied_text,
            "About the Job": (job_description[:500] + "...") if len(job_description) > 500 else job_description,
            "Full Description": job_description,
            "URL": job.url,
            "Job ID": job.job_id,
            "Page Number": None,  # filled by caller
            "PDF Filename": pdf_filename
        }

    def save_as_pdf(self, company, role, folder_month, job_id):
        try:

            # NEW
            self.guard_not_blocked(f"save_as_pdf(job_id={job_id}) before Page.printToPDF")

            # Create folder for that month
            month_folder = os.path.join(self.pdf_dir, folder_month)
            os.makedirs(month_folder, exist_ok=True)

            # Make a safe filename
            safe_company = re.sub(r'[<>:"/\\|?*]', '', company)[:50]
            safe_role = re.sub(r'[<>:"/\\|?*]', '', role)[:60]
            filename = f"{safe_company}_{safe_role}_{job_id}.pdf"
            filepath = os.path.join(month_folder, filename)

            # Print to PDF using Chrome DevTools Protocol command
            pdf_settings = {
                "landscape": False,
                "displayHeaderFooter": False,
                "printBackground": True,
                "preferCSSPageSize": True,
            }
            result = self.driver.execute_cdp_cmd("Page.printToPDF", pdf_settings)
            
            # Decode and write file
            with open(filepath, "wb") as f:
                f.write(base64.b64decode(result["data"]))

            # Return relative PDF path
            print(f"PDF saved: {folder_month}/{filename}")
            return os.path.join(folder_month, filename)

        except Exception as e:
            print(f"Could not save PDF: {e}")
            return None

    # -----------------------
    # Excel
    # -----------------------

    # If no job data has been collected yet, avoid creating empty files.
    def save_to_excel(self):

        # Always show where output is going
        abs_out = os.path.abspath(self.output_dir)
        print(f"\n💾 Saving Excel... output_dir = {abs_out}")

        if not self.data:
            print("\n⚠ No data to save")
            return

        # Each element in self.data is a dictionary returned from scrape_job_details().
        df = pd.DataFrame(self.data)

        # This file is intended to be human-readable and suitable for submission
        columns_order = ["Company", "Role", "Application Date", "About the Job", "URL", "Job ID", "Page Number", "PDF Filename"]

        # Keep only columns that actually exist (defensive against schema changes)
        df = df[[c for c in columns_order if c in df.columns]]

        excel_path = os.path.join(self.output_dir, "applications_2025_2026.xlsx")
        excel_path_abs = os.path.abspath(excel_path)
        
        try:
            df.to_excel(excel_path, index=False, engine="openpyxl")
            print(f"✅ Excel saved: {excel_path_abs}  (rows={len(df)})")
        except Exception as e:
            print(f"❌ Failed writing Excel: {e}")
            raise

        # -----------------------
        # Save full dataset (including complete job descriptions)
        # -----------------------
        # This version keeps every field captured during scraping and is useful for:
        # - Auditing
        # - Cross-checking
        # - Reprocessing data without re-scraping LinkedIn


        # full descriptions
        try:
            df_full = pd.DataFrame(self.data)
            full_excel_path = os.path.join(self.output_dir, "applications_full_descriptions.xlsx")
            full_excel_path_abs = os.path.abspath(full_excel_path)
            df_full.to_excel(full_excel_path, index=False, engine="openpyxl")
            print(f"✅ Full descriptions saved: {full_excel_path_abs}")

        except Exception as e:
            print(f"❌ Failed writing full Excel: {e}")
            raise

    def _dom_ready(self) -> bool:
        """
        Returns True if document.readyState is 'complete' or 'interactive'.
        If Chrome is hung, this can raise a timeout/WebDriverException.
        """
        state = self.driver.execute_script("return document.readyState")
        return state in ("complete", "interactive")

    def _page_seems_stuck(self) -> bool:
        """
        Heuristic for LinkedIn "infinite spinner / never finishes loading".
        We avoid LinkedIn-specific brittle selectors and check:
        - readyState not complete
        - AND presence of common loader/progress elements
        """
        try:
            # If DOM isn't ready, it might be stuck. Check for common loader/progressbars.
            ready = self._dom_ready()
            has_loader = self.driver.execute_script("""
                const sels = [
                  'div[role="progressbar"]',
                  '.artdeco-loader',
                  '.artdeco-loader__bar',
                  '.initial-load-animation',
                  '.loading',
                  '.spinner'
                ];
                return sels.some(s => document.querySelector(s));
            """)
            return (not ready) and bool(has_loader)
        except Exception:
            # If JS can't run, the session may already be unresponsive
            return True

    def safe_get(self, url: str, context: str = "", post_wait_xpath: str | None = None, timeout_sec: int = 45):
        """
        Robust navigation wrapper.
        - driver.get(url)
        - waits for DOM ready
        - optionally waits for a key element
        If it times out / driver becomes unresponsive -> raise RestartSession.
        """
        try:
            self.driver.get(url)
        except (ReadTimeoutError, TimeoutError, WebDriverException) as e:
            raise RestartSession(f"driver.get hung/unresponsive ({context}) -> {e}")

        # Wait for DOM readiness up to timeout_sec
        t0 = time.time()
        while time.time() - t0 < timeout_sec:
            try:
                if self._dom_ready():
                    break
            except (ReadTimeoutError, TimeoutError, WebDriverException) as e:
                raise RestartSession(f"DOM check failed ({context}) -> {e}")
            time.sleep(0.5)

        # If still not ready OR looks stuck, attempt ONE refresh
        if self._page_seems_stuck():
            try:
                self.driver.refresh()
            except (ReadTimeoutError, TimeoutError, WebDriverException) as e:
                raise RestartSession(f"refresh failed ({context}) -> {e}")

            # Wait again briefly after refresh
            t1 = time.time()
            while time.time() - t1 < 20:
                try:
                    if self._dom_ready() and not self._page_seems_stuck():
                        break
                except (ReadTimeoutError, TimeoutError, WebDriverException) as e:
                    raise RestartSession(f"DOM check after refresh failed ({context}) -> {e}")
                time.sleep(0.5)

        # Optional: wait for a specific element that indicates page is usable
        if post_wait_xpath:
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, post_wait_xpath))
                )
            except TimeoutException:
                # If we can't see the expected content, treat it as stuck/unusable
                if self._page_seems_stuck():
                    raise RestartSession(f"Page stuck/unusable after navigation ({context}). URL={url}")

    def restart_browser_session(self, reason: str):
        """
        Save state + excel, close current driver, and force relogin.
        """
        print(f"\n🔁 Restarting browser session: {reason}")

        # Save progress
        try:
            self.save_to_excel()
        except Exception:
            pass
        self.save_state(reason=f"restart_session: {reason}")

        # Notify user (top-most)
        self.notify_topmost(
            "LinkedIn session restarted",
            "Selenium got stuck / unresponsive (likely infinite loading).\n\n"
            "I saved your Excel + resume state.\n"
            "A new browser session will start. Please log in again when prompted."
        )

        # Kill driver
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass

        self.driver = None
        self.wait = None

    # -----------------------
    # Main
    # -----------------------
    def scrape_page(self, page_num: int):

        print(f"\n{'='*60}\nSCRAPING PAGE {page_num}\n{'='*60}")

        # Build the correct list-page URL and open it
        page_url = self.get_page_url(page_num)

        self.safe_get(
            page_url,
            context=f"scrape_page(page={page_num})",
            post_wait_xpath="//div[contains(@class,'linked-area')]//a[contains(@href, '/jobs/view/')]"
        )

        self.last_page = page_num
        self.last_url = page_url
        self.save_state(reason=f"entered page {page_num}")

        # NEW
        self.guard_not_blocked(f"scrape_page(page={page_num}) after driver.get")

        # Confirm the list is actually visible (and you are logged in)
        try:
            self._wait_for_list_page()
        except TimeoutException:
            print(f"⚠ Could not detect job list on page {page_num}. Are you logged in / is Applied list visible?")
            return

        # Light scroll to ensure all items render (safe + cheap)
        # Move down 700px, 4 times. Then back to the top.
        for _ in range(4):
            self.driver.execute_script("window.scrollBy(0, 700);")
            self.micro_pause()

        # NEW
        self.guard_not_blocked(f"scrape_page(page={page_num}) after scroll")

        self.driver.execute_script("window.scrollTo(0, 0);")
        self.micro_pause()

        # Collect job links + hints from the list page
        jobs = self.get_jobs_from_list_page()
        print(f"Found {len(jobs)} job links on page {page_num}")

        if not jobs:
            print("⚠ Found 0 jobs on this page. Dumping current URL for debugging:")
            print("   URL:", self.driver.current_url)

        # Loop through each job and scrape the detail page
        for i, job in enumerate(jobs, 1):
            print(f"\n  → [{i}/{len(jobs)}] Opening job {job.job_id}")

            # For each job: scrape details and store row data
            try:
                row = self.scrape_job_details(job)
                row["Page Number"] = page_num
                self.data.append(row)

                self.last_url = job.url
                self.save_state(reason=f"scraped job {job.job_id} on page {page_num}")

                print(f" ✓ {row['Company']} — {row['Role']}")
                print(f" Applied: {row['Application Date']} (from: {job.applied_relative})")

            except Exception as e:
                print(f"    ✗ Failed job {job.job_id}: {e}")

            # smaller delay between jobs
            self.wait_random(0.8, 1.7)

        # Page completion summary
        print(f"\n✓ Page {page_num} complete. Total collected so far: {len(self.data)}")

        # Optional safety: save after each page
        self.save_to_excel()
        print("💾 Saved after page", page_num)


    def run(self):
        MAX_RESTARTS = 6
        restarts = 0

        while True:
            try:
                print("\n" + "=" * 60)
                print("LINKEDIN JOB APPLICATION SCRAPER (ROBUST)")
                print("=" * 60)
                print(f"Pages: {self.start_page} → {self.end_page}")
                print(f"Now (SG): {self.current_date.strftime('%Y-%m-%d %H:%M')}")
                print("=" * 60)

                self.load_state()

                # Start browser
                self.setup_driver()
                self.manual_login()

                # HEALTH CHECK (Step 2: detect failure fast)
                if not self.profile_healthy():
                    # poison strike
                    self.profile_poison_strikes += 1
                    raise RestartSession(f"profile_healthy failed (strike={self.profile_poison_strikes})")

                # If healthy, reset strikes
                self.profile_poison_strikes = 0

                resume_from = self.get_resume_start_page()

                for page_num in range(resume_from, self.end_page - 1, -1):
                    self.scrape_page(page_num)

                self.save_to_excel()
                self.save_state(reason="completed")
                print("\n✓ SCRAPING COMPLETE!")
                print(f"Output: {os.path.abspath(self.output_dir)}")
                return

            except KeyboardInterrupt:
                print("\n⚠ Interrupted. Saving partial results...")
                self.save_to_excel()
                self.save_state(reason="keyboard_interrupt")
                return

            except RuntimeError as e:
                # guard_not_blocked() throws RuntimeError
                print(f"\n⚠ RuntimeError: {e}")
                self.save_state(reason=f"runtime_error: {e}")

                restarts += 1
                if restarts > MAX_RESTARTS:
                    print("🛑 Too many restarts. Exiting.")
                    return

                # For checkpoint blocks: do NOT auto-delete profile.
                # Just restart session (you will login again).
                try:
                    if self.driver:
                        self.driver.quit()
                except Exception:
                    pass

                print("🔁 Restarting a fresh browser session... (login required)")
                continue

            except RestartSession as e:
                print(f"\n⚠ Session/Profile issue detected: {e}")
                restarts += 1

                if restarts > MAX_RESTARTS:
                    print("🛑 Too many restarts. Exiting.")
                    self.save_state(reason=f"too_many_restarts: {e}")
                    return

                # Step 1: Reuse same profile once
                if self.profile_poison_strikes <= 1:
                    self.restart_browser_session(str(e))
                    # Loop continues: will setup_driver + manual_login again
                    continue

                # Step 3: Failure repeats → reset profile folder
                self.reset_profile()
                self.profile_poison_strikes = 0

                self.notify_topmost(
                    "LinkedIn profile reset",
                    "LinkedIn/Selenium got stuck twice.\n\n"
                    "I deleted the Selenium profile folder (fresh cookies).\n"
                    "A new browser session will start. Please log in again."
                )

                # Loop continues: will setup_driver + manual_login again
                continue

            except Exception as e:
                print(f"\n❌ Unexpected error: {e}")
                print(traceback.format_exc())
                try:
                    self.save_to_excel()
                except Exception:
                    pass
                self.save_state(reason=f"unexpected_error: {e}")
                raise

            finally:
                try:
                    if self.driver:
                        self.driver.quit()
                        print("✓ Browser closed")
                except Exception:
                    pass
                self.driver = None
                self.wait = None

if __name__ == "__main__":

    START_PAGE = 37
    END_PAGE = 1
    LinkedInJobScraper(start_page=START_PAGE, end_page=END_PAGE).run()