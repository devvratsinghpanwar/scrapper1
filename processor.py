import pandas as pd
import time
import re
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium_stealth import stealth
import json

# --- Configuration ---
OUTPUT_DIR = 'output'
SEARCH_ENGINE_URL = "https://search.brave.com/"

# --- Helper Functions ---
def setup_driver():
    """Sets up a stealthy, VISIBLE Selenium Chrome driver."""
    options = webdriver.ChromeOptions()
    # Keep LinkedIn login session
    options.add_argument("--user-data-dir=F:/bot")

    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    stealth(driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
            )
    return driver

def handle_captcha_and_wait(driver):
    """
    Pauses the script and waits for user to solve CAPTCHA manually.
    """
    print("\n" + "="*60)
    print("! ACTION REQUIRED: Automation paused by a security check.")
    print("! Please solve the 'I am not a robot' CAPTCHA in the browser window.")
    input("! AFTER you have solved it and the search page is loaded, press Enter here...")
    print("! Resuming automation...")
    print("="*60 + "\n")
    time.sleep(2)
    return driver

def parse_revenue(revenue_str):
    if not revenue_str:
        return None, None
    revenue_str = str(revenue_str).lower().strip()
    billions = ['b', 'billion']
    millions = ['m', 'million']
    multiplier = 1
    if any(word in revenue_str for word in billions):
        multiplier = 1000
    elif any(word in revenue_str for word in millions):
        multiplier = 1
    num_match = re.search(r'([\d,]+\.?\d*|\d+\.?\d*)', revenue_str)
    if not num_match:
        return None, None
    try:
        value_str = num_match.group(1).replace(',', '').replace('$', '')
        value = float(value_str)
        return value * multiplier, revenue_str
    except (ValueError, IndexError):
        return None, None

def assign_tier(revenue_in_millions):
    if revenue_in_millions is None:
        return "Unknown"
    if revenue_in_millions > 1000:
        return "Super Platinum"
    elif 500 <= revenue_in_millions <= 1000:
        return "Platinum"
    elif 100 <= revenue_in_millions < 500:
        return "Diamond"
    else:
        return "Gold"

def generate_email(full_name, domain):
    if not domain:
        return "Not Found (No Domain)"
    try:
        names = full_name.lower().split()
        if len(names) > 1:
            return f"{names[0]}.{names[-1]}@{domain}"
        elif len(names) == 1:
            return f"{names[0]}@{domain}"
        return "Not Found (Name Error)"
    except Exception:
        return "Error Generating"

# --- Core Selenium Functions ---

def robust_search(driver, query):
    """
    Resilient search that loops until search is successful.
    If CAPTCHA appears, waits for manual solving.
    """
    driver.get(SEARCH_ENGINE_URL)

    while True:
        try:
            search_box = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'searchbox'))
            )
            search_box.clear()
            search_box.send_keys(query)
            search_box.send_keys(Keys.RETURN)

            # Wait for at least one result snippet
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.snippet"))
            )
            print("   -> Search submitted successfully.")
            time.sleep(2)
            return

        except TimeoutException:
            handle_captcha_and_wait(driver)

def get_company_info(driver, company_name):
    """Uses Selenium and Brave Search to find company revenue and domain."""
    print(f"   -> Searching Brave for '{company_name}' info...")
    revenue_val, revenue_display, domain = None, "Not Found", None
    try:
        robust_search(driver, f'"{company_name}" annual revenue')

        body_text = driver.find_element(By.TAG_NAME, 'body').text
        revenue_pattern = r'revenue.*?([\$€£]?\s?\d[\d,]*\.?\d*\s?(?:billion|million|B|M))'
        match = re.search(revenue_pattern, body_text, re.IGNORECASE)
        if match:
            revenue_val, revenue_display = parse_revenue(match.group(1))
            if revenue_val:
                print(f"   -> Found revenue: {revenue_display}")

        search_results = driver.find_elements(By.CSS_SELECTOR, 'div.snippet')
        for result in search_results:
            try:
                link_tag = result.find_element(By.TAG_NAME, 'a')
                href = link_tag.get_attribute('href')
                if href:
                    domain_match = re.search(r'https?://(?:www\.)?([^/]+)', href)
                    if domain_match:
                        domain = domain_match.group(1)
                        print(f"   -> Found domain: {domain}")
                        break
            except:
                continue
    except Exception as e:
        print(f"   -> [ERROR] getting company info: {e}")
    return revenue_val, revenue_display, domain

def get_contact_info(driver, full_name, company_name):
    """
    Uses Selenium and Brave Search to find LinkedIn profile and designation.
    First visits Brave search to find the LinkedIn URL, then scrapes the profile headline.
    """
    print(f"   -> Searching Brave for '{full_name}' at '{company_name}'...")
    linkedin_url, designation = "Not Found", "Not Found"
    
    try:
        # Search for LinkedIn profile on Brave
        robust_search(driver, f'site:linkedin.com/in/ "{full_name}" "{company_name}"')

        first_result = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.snippet"))
        )
        link_tag = first_result.find_element(By.TAG_NAME, 'a')
        linkedin_url = link_tag.get_attribute('href')
        print(f"   -> Found LinkedIn URL: {linkedin_url}")

        # Visit LinkedIn profile to get actual designation
        driver.get(linkedin_url)
        try:
            designation_elem = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.text-body-medium.break-words"))
            )
            designation = designation_elem.text.strip()
            print(f"   -> Extracted Designation from profile: {designation}")
        except TimeoutException:
            print("   -> Designation not visible (profile restricted or layout changed). Trying snippet...")
            try:
                title_text = first_result.find_element(By.CSS_SELECTOR, 'span.title').text
                if '-' in title_text:
                    designation = title_text.split('-')[1].split('·')[0].strip()
                    print(f"   -> Fallback designation from snippet: {designation}")
            except:
                pass

    except Exception as e:
        print(f"   -> [ERROR] getting contact info: {e}")
    
    return linkedin_url, designation


# --- Main Orchestrator ---

def run_automation(input_filepath, status_callback):
    driver = None
    try:
        status_callback("Setting up secure browser...")
        print("\n[START] Setting up Selenium driver...")
        driver = setup_driver()
        print("[SUCCESS] Driver setup complete.")

        status_callback("Reading the input Excel file...")
        company_df = pd.read_excel(input_filepath, sheet_name='Company')
        contacts_df = pd.read_excel(input_filepath, sheet_name='Contacts')
        print("[INFO] Successfully read Excel sheets.")

        company_results = []
        for index, row in company_df.iterrows():
            company_name = row['Company Name']
            status_callback(f"Company ({index+1}/{len(company_df)}): Processing {company_name}...")
            print(f"\n[INFO] Processing Company: {company_name}")
            revenue_val, revenue_display, domain = get_company_info(driver, company_name)
            tier = assign_tier(revenue_val)

            new_row = row.to_dict()
            new_row['Revenue'] = revenue_display
            new_row['Tier'] = tier
            new_row['Domain'] = domain
            company_results.append(new_row)

        enriched_companies = pd.DataFrame(company_results)
        company_domain_map = pd.Series(enriched_companies.Domain.values,
                                       index=enriched_companies['Company Name']).to_dict()

        contact_results = []
        for index, row in contacts_df.iterrows():
            full_name, company = row['Full Name'], row['Current Company']
            status_callback(f"Contact ({index+1}/{len(contacts_df)}): Processing {full_name}...")
            print(f"\n[INFO] Processing Contact: {full_name} at {company}")
            linkedin_url, designation = get_contact_info(driver, full_name, company)

            domain = company_domain_map.get(company)
            email = generate_email(full_name, domain)

            new_row = row.to_dict()
            new_row['LinkedIn URL'] = linkedin_url
            new_row['Designation'] = designation
            new_row['Work Email'] = email
            contact_results.append(new_row)

        enriched_contacts = pd.DataFrame(contact_results)

        status_callback("Saving enriched data...")
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        excel_path = os.path.join(OUTPUT_DIR, 'Enriched_Results.xlsx')
        json_path = os.path.join(OUTPUT_DIR, 'enriched_data.json')

        enriched_companies_final = enriched_companies.drop(columns=['Domain'], errors='ignore')
        with pd.ExcelWriter(excel_path) as writer:
            enriched_companies_final.to_excel(writer, sheet_name='Enriched Companies', index=False)
            enriched_contacts.to_excel(writer, sheet_name='Enriched Contacts', index=False)

        companies_json_df = enriched_companies_final.rename(columns={'Country/Region': 'Country'})
        contacts_json_df = enriched_contacts.rename(columns={'Current Company': 'Company'})
        final_json_data = {
            "companies": companies_json_df[['Company Name', 'Country', 'Revenue', 'Tier']].to_dict(orient='records'),
            "contacts": contacts_json_df[['Full Name', 'Company', 'LinkedIn URL', 'Designation', 'Work Email']].to_dict(orient='records')
        }
        with open(json_path, 'w') as f:
            json.dump(final_json_data, f, indent=2)

        status_callback("Processing complete!")
        print("\n[SUCCESS] Automation process finished.")
        return excel_path, json_path, enriched_companies_final, enriched_contacts

    finally:
        if driver:
            driver.quit()
            print("[INFO] Selenium driver closed.")
