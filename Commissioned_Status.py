import sys
import subprocess
import os
import time
from datetime import datetime, timedelta

def install_dependencies():
    packages = {
        'pandas': 'pandas',
        'selenium': 'selenium',
        'openpyxl': 'openpyxl'
    }
    for module_name, package_name in packages.items():
        try:
            __import__(module_name)
        except ImportError:
            print(f"Installing missing dependency: {package_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

# Run dependency check before importing third-party libraries
install_dependencies()

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# -------- Intro of P-schedule --------------------

print(r'''
 ,ggggggggggg,                 ,gg,                                                                         
dP"""88""""""Y8,              i8""8i             ,dPYb,                      8I              ,dPYb,         
Yb,  88      `8b              `8,,8'             IP'`Yb                      8I              IP'`Yb         
 `"  88      ,8P               `88'              I8  8I                      8I              I8  8I         
     88aaaad8P"                dP"8,             I8  8'                      8I              I8  8'         
     88"""""    aaaaaaaa      dP' `8a    ,gggg,  I8 dPgg,    ,ggg,     ,gggg,8I  gg      gg  I8 dP   ,ggg,  
     88         """"""""     dP'   `Yb  dP"  "Yb I8dP" "8I  i8" "8i   dP"  "Y8I  I8      8I  I8dP   i8" "8i 
     88                  _ ,dP'     I8 i8'       I8P    I8  I8, ,8I  i8'    ,8I  I8,    ,8I  I8P    I8, ,8I 
     88                  "888,,____,dP,d8,_    _,d8     I8, `YbadP' ,d8,   ,d8b,,d8b,  ,d8b,,d8b,_  `YbadP' 
     88                  a8P"Y88888P" P""Y8888PP88P     `Y8888P"Y888P"Y8888P"`Y88P'"Y88P"`Y88P'"Y88888P"Y888
''')

# -------- CONFIG ----------
LOGIN_URL = "https://roams.cris.org.in/uaa/login"
USERNAME = "9004441529"
PASSWORD = "PL@40028pl"
HOLD_SECONDS = 10
WAIT_TIMEOUT = 30
# --------------------------

# -------- Read Coach Numbers from Excel (cleaned) --------
def read_coach_numbers(file_path):
    """Read and clean coach numbers from Excel file"""
    try:
        df = pd.read_excel(file_path)
        coach_numbers = df.iloc[:, 0].dropna().astype(str).str.strip()
        cleaned_coach_numbers = []
        for coach_str in coach_numbers:
            cleaned = ''.join(filter(str.isdigit, coach_str))
            if cleaned:
                if len(cleaned) == 5:
                    cleaned = "0" + cleaned
                cleaned_coach_numbers.append(cleaned)
        return cleaned_coach_numbers
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# -------- Safe Click with Retry --------
def safe_click(driver, wait, by, locator, max_attempts=3):
    """Click element with retry on stale reference"""
    for attempt in range(max_attempts):
        try:
            element = wait.until(EC.element_to_be_clickable((by, locator)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", element)
            return True
        except StaleElementReferenceException:
            if attempt == max_attempts - 1:
                raise
            print(f"Stale element, retrying... (attempt {attempt + 1})")
            time.sleep(1)
        except Exception as e:
            if attempt == max_attempts - 1:
                raise
            time.sleep(1)
    return False

# -------- LOGIN AND NAVIGATION --------
def login_and_navigate(driver):
    """Login and navigate to coach profile page"""
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    try:
        print("Opening login page...")
        driver.get(LOGIN_URL)

        # Wait for the username field
        wait.until(EC.presence_of_element_located((By.ID, "username")))

        # Fill in login credentials
        print("Entering credentials...")
        driver.find_element(By.ID, "username").send_keys(USERNAME)
        driver.find_element(By.ID, "password").send_keys(PASSWORD)

        # Click login button with hold
        login_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(), 'Log In') or contains(text(), 'लॉग इन')]")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", login_button)
        print(f"Holding login button for {HOLD_SECONDS} seconds...")
        actions = ActionChains(driver)
        actions.click_and_hold(login_button).perform()
        time.sleep(HOLD_SECONDS)
        actions.release().perform()
        login_button.click()

        # Wait for successful login
        print("Waiting for login success...")
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div/div/div/a")))

        # Navigate to CMMS
        print("Navigating to CMMS...")
        cmms_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div/div/div/a")))
        cmms_link.click()

        # Wait for OTP input
        print("Waiting for OTP (please enter manually)...")
        otp_input_xpath = '//*[@id="otpForm"]/div/input[4]'
        wait.until(EC.presence_of_element_located((By.XPATH, otp_input_xpath)))
        time.sleep(10)  # Manual wait for OTP input

        # Verify OTP
        print("Verifying OTP...")
        verify_button_xpath = '//*[@id="saveButton"]'
        verify_button = wait.until(EC.element_to_be_clickable((By.XPATH, verify_button_xpath)))
        verify_button.click()
        time.sleep(3)  # Allow page load after OTP

        # Navigate to coach master
        print("Navigating to coach master...")
        safe_click(driver, wait, By.XPATH, '/html/body/div/div[2]/nav/div/ul/li[4]/a')

        # Navigate to search and edit coach
        print("Opening search and edit coach...")
        safe_click(driver, wait, By.XPATH, '//*[@id="navbarTogglerCmm"]/ul/li[4]/div/a[1]')

        wait.until(EC.presence_of_element_located((By.ID, "coachNo")))
        time.sleep(2)

        print("Login and navigation successful!\n")
        return wait

    except Exception as e:
        print(f"Error during login and navigation: {e}")
        raise

# -------- Navigate to Coach Profile --------
def navigate_to_coach_profile(driver, wait):
    """Navigate back to search and edit coach page"""
    try:
        coach_master_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[2]/nav/div/ul/li[4]/a'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", coach_master_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", coach_master_button)
        time.sleep(1)

        search_edit_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="navbarTogglerCmm"]/ul/li[4]/div/a[1]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", search_edit_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", search_edit_button)
        time.sleep(1)

        wait.until(EC.presence_of_element_located((By.ID, "coachNo")))
        wait.until(EC.element_to_be_clickable((By.ID, "coachNo")))
        time.sleep(2)

    except Exception as e:
        print(f"Error navigating to coach profile: {e}")
        raise

# -------- Search Coach and Extract Data --------
def search_coach(driver, wait, coach_no, max_retries=3):
    """Search for coach and extract data with retry mechanism"""

    coach_data = {
        "Coach Number": coach_no,
        "Status": "Unprocessed",
        "Commission Status": "N/A",
        "Railway": "N/A",
        "Coach Type": "N/A",
        "Base Depot": "N/A",
        "Owning Division": "N/A",
        "Manufacture": "N/A",
        "Workshop": "N/A",
        "Built Date": "N/A",
        "Factory Turn Out Date": "N/A",
        "Fitness Type": "N/A",
        "Commission Date": "N/A",
        "Gauge": "N/A",
        "Coach Category": "N/A",
        "Power Generation Type": "N/A",
        "Tare Weight(In Tonnes)": "N/A",
        "Is Parcel Coach": "N/A",
        "Is Parcel Facility Available": "N/A",
        "Parcel Capacity Weight": "N/A",
        "Parcel Capacity Volume": "N/A"
    }

    for attempt in range(max_retries):
        try:
            coach_input_field = wait.until(EC.presence_of_element_located((By.ID, "coachNo")))
            coach_input_field = wait.until(EC.element_to_be_clickable((By.ID, "coachNo")))

            coach_input_field.clear()
            coach_input_field.send_keys(coach_no)

            coach_go_xpath = '//*[@id="tableForCoachSearch"]/tbody/tr[1]/td[6]/button'
            coach_go = wait.until(EC.element_to_be_clickable((By.XPATH, coach_go_xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", coach_go)
            driver.execute_script("arguments[0].click();", coach_go)

            try:
                # Use a specific dynamic xpath for the search result row to ensure it belongs to the current search
                result_row_xpath = f'//*[@id="tableForCoachSearch"]/tbody/tr[2]/td[contains(., "{coach_no}")]/..'
                
                # Wait for explicitly the row holding our coach number to be present (but do not click it yet)
                wait.until(EC.presence_of_element_located((By.XPATH, result_row_xpath)))
                
                # Extract the commission status FIRST
                status_xpath = f'//*[@id="tableForCoachSearch"]/tbody/tr[2]/td[contains(., "{coach_no}")]/../td[5]'
                status_element = wait.until(EC.presence_of_element_located((By.XPATH, status_xpath)))
                commission_status = status_element.text.strip()
                coach_data["Commission Status"] = commission_status

                if commission_status == "COMMISSIONED":
                    # Because it IS commissioned, click the row to load the details pane
                    result_row_clickable = wait.until(EC.element_to_be_clickable((By.XPATH, result_row_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", result_row_clickable)
                    driver.execute_script("arguments[0].click();", result_row_clickable)
                    
                    # Wait for explicitly the Coach Type field to load completely before fetching
                    coach_type_xpath = "/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[1]/td[2]"
                    wait.until(
                        lambda d: d.find_element(By.XPATH, coach_type_xpath).text.strip() != ""
                    )
                    
                    # Extract extra fields with a shorter timeout (1 second).
                    # This prevents the script from waiting 30 seconds if an optional field is missing.
                    short_wait = WebDriverWait(driver, 1)
                    extra_xpaths = [
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[1]/div/div/form/div[1]/table/tbody/tr[2]/td[1]", "Railway"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[1]/td[2]", "Coach Type"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[1]/td[3]/select", "Base Depot"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[2]/td[1]/select", "Owning Division"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[2]/td[2]/select", "Manufacture"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[2]/td[3]/select", "Workshop"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[3]/td[1]/input", "Built Date"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[3]/td[2]/input", "Factory Turn Out Date"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[3]/td[3]/select", "Fitness Type"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[4]/div[1]/div/div/form[1]/table/tbody/tr[1]/td[2]/input", "Commission Date"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[4]/div[1]/div/div/form[1]/table/tbody/tr[2]/td[2]", "Gauge"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[4]/div[1]/div/div/form[1]/table/tbody/tr[3]/td[2]", "Coach Category"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[4]/div[1]/div/div/form[1]/table/tbody/tr[6]/td[2]/select", "Power Generation Type"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[4]/div[1]/div/div/form[1]/table/tbody/tr[7]/td[2]/input", "Tare Weight(In Tonnes)"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[4]/td/select", "Is Parcel Coach"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[5]/td[1]/select", "Is Parcel Facility Available"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[5]/td[2]/input", "Parcel Capacity Weight"),
                        ("/html/body/div[1]/div[3]/form/table/tbody/tr/td[2]/div[2]/div/div/form/table/tbody/tr[5]/td[3]/input", "Parcel Capacity Volume"),
                    ]

                    for xpath, label in extra_xpaths:
                        try:
                            # Use short_wait instead of the standard 30-second wait
                            element = short_wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                            tag_name = element.tag_name.lower()
                            
                            if tag_name == "input":
                                val = element.get_attribute("value")
                            elif tag_name == "select":
                                try:
                                    val = Select(element).first_selected_option.text
                                except Exception:
                                    val = element.get_attribute("value")
                            else:
                                val = element.text
                                
                            coach_data[label] = val.strip() if val and val.strip() else "N/A"
                        except Exception:
                            coach_data[label] = "N/A"
                
            except TimeoutException:
                print(f"  ✗ Coach {coach_no} not found in the system")
                coach_data["Status"] = "Not Found"
                return coach_data

            coach_data["Status"] = "Success"
            print(f"  ✓ Successfully extracted data for coach {coach_no}")

            navigate_to_coach_profile(driver, wait)
            return coach_data

        except StaleElementReferenceException as e:
            print(f"  ! Stale element for coach {coach_no}, attempt {attempt + 1}/{max_retries}")
            if attempt < max_retries - 1:
                time.sleep(2)
                navigate_to_coach_profile(driver, wait)
                continue
            else:
                print(f"  ✗ Failed to process coach {coach_no} after {max_retries} attempts")
                coach_data["Status"] = "Error - Stale Element"
                return coach_data

        except Exception as e:
            print(f"  ✗ Error searching coach {coach_no} (attempt {attempt + 1}/{max_retries}): {str(e)[:100]}")
            if attempt < max_retries - 1:
                time.sleep(2)
                try:
                    navigate_to_coach_profile(driver, wait)
                except:
                    pass
                continue
            else:
                coach_data["Status"] = f"Error - {str(e)[:50]}"
                return coach_data

    coach_data["Status"] = "Error - Max retries exceeded"
    return coach_data

# -------- MAIN SCRIPT --------
if __name__ == "__main__":
    driver = None

    try:
        print("="*60)
        print("COACH DATA EXTRACTION SCRIPT")
        print("="*60)

        # Initialize driver
        print("\n[1/5] Initializing Chrome WebDriver...")
        driver = webdriver.Chrome()
        driver.maximize_window()
        print("✓ WebDriver initialized")

        # Read coach numbers (Excel file must be in same dir as exe or path provided)
        print("\n[2/5] Reading coach numbers from Excel...")
        coach_file_path = os.path.join(os.path.abspath("."), "coach_numbers.xlsx")
        coach_numbers = read_coach_numbers(coach_file_path)

        if not coach_numbers:
            print("✗ No valid coach numbers found in Excel file")
            exit(1)

        print(f"✓ Found {len(coach_numbers)} coach numbers to process")
        print(f"  Coach numbers: {', '.join(coach_numbers[:5])}{'...' if len(coach_numbers) > 5 else ''}")

        # Login and navigate
        print("\n[3/5] Logging in and navigating to coach profile page...")
        wait = login_and_navigate(driver)

        # Process each coach
        print("\n[4/5] Extracting coach data...")
        print("-"*60)

        all_coach_data = []
        start_time = time.time()  # Start timer

        for idx, coach_no in enumerate(coach_numbers, 1):
            coach_start_time = time.time()  # Timer for individual coach
            
            print(f"\n[{idx}/{len(coach_numbers)}] Processing coach: {coach_no}")
            coach_data = search_coach(driver, wait, coach_no)
            all_coach_data.append(coach_data)

            coach_elapsed = time.time() - coach_start_time
            print(f"  ⏱ Time taken: {coach_elapsed:.2f} seconds")

            # Calculate and display progress statistics
            elapsed_time = time.time() - start_time
            avg_time_per_coach = elapsed_time / idx
            remaining_coaches = len(coach_numbers) - idx
            estimated_remaining = avg_time_per_coach * remaining_coaches
            
            print(f"  📊 Progress: {idx}/{len(coach_numbers)} ({(idx/len(coach_numbers)*100):.1f}%)")
            print(f"  ⏰ Elapsed: {str(timedelta(seconds=int(elapsed_time)))}")
            print(f"  🕐 Estimated remaining: {str(timedelta(seconds=int(estimated_remaining)))}")

            # Save progress every 50 coaches
            if idx % 50 == 0:
                temp_df = pd.DataFrame(all_coach_data)
                temp_df.to_excel('coach_data_progress.xlsx', index=False)
                print(f"\n  💾 Progress saved: {idx} coaches processed")

        total_time = time.time() - start_time  # Calculate total time

        # Save final results
        print("\n[5/5] Saving final results...")
        df = pd.DataFrame(all_coach_data)
        df.to_excel('coach_data.xlsx', index=False)

        # Print summary
        print("\n" + "="*60)
        print("EXTRACTION COMPLETE!")
        print("="*60)
        print(f"Total coaches processed: {len(all_coach_data)}")
        print(f"Successful: {len([c for c in all_coach_data if c.get('Status') == 'Success'])}")
        print(f"Not Found: {len([c for c in all_coach_data if c.get('Status') == 'Not Found'])}")
        print(f"Errors: {len([c for c in all_coach_data if c.get('Status', '').startswith('Error')])}")
        print(f"\n⏱ TOTAL TIME TAKEN: {str(timedelta(seconds=int(total_time)))}")
        print(f"⏱ Average time per coach: {total_time/len(all_coach_data):.2f} seconds")
        print(f"\n✓ Results saved to: coach_data.xlsx")
        print("="*60)

    except KeyboardInterrupt:
        print("\n\n⚠ Script interrupted by user")

    except Exception as e:
        print(f"\n\n✗ Critical error in main script: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if driver:
            print("\nClosing browser...")
            driver.quit()
            print("✓ Browser closed")