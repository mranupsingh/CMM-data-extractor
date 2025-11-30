import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
import sys
import os
from datetime import datetime, timedelta

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

# -------- Helper to find bundled resource --------
def resource_path(relative_path):
    """ Get absolute path to resource, works for development and PyInstaller """
    try:
        base_path = sys._MEIPASS  # PyInstaller temporary folder when bundled
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# -------- CONFIG ----------
CHROMEDRIVER_PATH = resource_path("chromedriver.exe")
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

        # Navigate to coach maintenance
        print("Navigating to coach maintenance...")
        safe_click(driver, wait, By.XPATH, '//*[@id="navbarDropdownMenuDepotLink"]')

        # Navigate to coach profile
        print("Opening coach profile...")
        safe_click(driver, wait, By.XPATH, '//*[@id="navbarTogglerCmm"]/ul/li[3]/div/a[1]')

        wait.until(EC.presence_of_element_located((By.ID, "coachNo")))
        time.sleep(2)

        print("Login and navigation successful!\n")
        return wait

    except Exception as e:
        print(f"Error during login and navigation: {e}")
        raise

# -------- Navigate to Coach Profile --------
def navigate_to_coach_profile(driver, wait):
    """Navigate back to coach profile page"""
    try:
        coach_maintenance_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="navbarDropdownMenuDepotLink"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", coach_maintenance_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", coach_maintenance_button)
        time.sleep(1)

        coach_pro_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="navbarTogglerCmm"]/ul/li[3]/div/a[1]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", coach_pro_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", coach_pro_button)
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
        "Coach Type": "N/A",
        "Built Date": "N/A",
        "Gauge": "N/A",
        "Coupling Type": "N/A",
        "Tare Weight": "N/A",
        "Commission Date(DD/MM/YYYY": "N/A",
        "Base Depot": "N/A",
        "Coach Category": "N/A",
        "Fitness Type": "N/A",
        "Max Speed": "N/A",
        "Factory TurnOut Date": "N/A",
        "Maintenance Zone": "N/A",
        "Last POH/SS2/SS3 Workshop": "N/A",
        "POH/SS2/SS3 Date": "N/A",
        "IOH/SS1 Date": "N/A",
        "Expected IOH/SS1 Due Date": "N/A",
        "Maintenance Depot": "N/A",
        "Manufactured By": "N/A",
        "Return Date": "N/A",
        "IOH/SS1 Location": "N/A",
        "Extended Return Date": "N/A"
    }

    for attempt in range(max_retries):
        try:
            coach_input_field = wait.until(EC.presence_of_element_located((By.ID, "coachNo")))
            coach_input_field = wait.until(EC.element_to_be_clickable((By.ID, "coachNo")))

            coach_input_field.clear()
            coach_input_field.send_keys(coach_no)

            coach_go_xpath = '/html/body/div/div[3]/form/div/table/tbody/tr[1]/td[5]/button'
            coach_go = wait.until(EC.element_to_be_clickable((By.XPATH, coach_go_xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", coach_go)
            driver.execute_script("arguments[0].click();", coach_go)

            try:
                found_coach_xpath = '/html/body/div/div[3]/form/div/table/tbody/tr[2]/td[2]/a/span'
                found_coach = wait.until(EC.element_to_be_clickable((By.XPATH, found_coach_xpath)))
                driver.execute_script("arguments[0].click();", found_coach)
            except TimeoutException:
                print(f"  ✗ Coach {coach_no} not found in the system")
                coach_data["Status"] = "Not Found"
                return coach_data

            # Wait for page to load completely - wait for any key element
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/div[1]/div/div[1]/div[2]"))
            )
            time.sleep(1)  # Brief pause to ensure page is stable

            data_xpaths = [
                ("/html/body/div/div[3]/div[1]/div/div[1]/div[2]", "Coach Number"),
                ("/html/body/div/div[3]/div[1]/div/div[2]/div[2]", "Base Depot"),
                ("/html/body/div/div[3]/div[1]/div/div[3]/div[2]", "Coach Category"),
                ("/html/body/div/div[3]/div[1]/div/div[4]/div[2]", "Fitness Type"),
                ("/html/body/div/div[3]/div[1]/div/div[5]/div[2]", "Max Speed"),
                ("/html/body/div/div[3]/div[1]/div/div[6]/div[2]", "Factory TurnOut Date"),
                ("/html/body/div/div[3]/div[1]/div/div[2]/div[2]", "Base Depot"),
                ("/html/body/div/div[3]/div[1]/div/div[1]/div[4]", "Coach Type"),
                ("/html/body/div/div[3]/div[1]/div/div[2]/div[4]", "Built Date"),
                ("/html/body/div/div[3]/div[1]/div/div[3]/div[4]", "Gauge"),
                ("/html/body/div/div[3]/div[1]/div/div[4]/div[4]", "Coupling Type"),
                ("/html/body/div/div[3]/div[1]/div/div[5]/div[4]", "Tare Weight"),
                ("/html/body/div/div[3]/div[1]/div/div[6]/div[4]", "Commission Date(DD/MM/YYYY"),
                ("/html/body/div/div[3]/div[2]/div/div[1]/div[2]", "Maintenance Zone"),
                ("/html/body/div/div[3]/div[2]/div/div[2]/div[2]", "Last POH/SS2/SS3 Workshop"),
                ("/html/body/div/div[3]/div[2]/div/div[3]/div[2]", "POH/SS2/SS3 Date"),
                ("/html/body/div/div[3]/div[3]/div/div[4]/div[2]", "IOH/SS1 Date"),
                ("/html/body/div/div[3]/div[2]/div/div[5]/div[2]", "Expected IOH/SS1 Due Date"),
                ("/html/body/div/div[3]/div[2]/div/div[1]/div[4]", "Maintenance Depot"),
                ("/html/body/div/div[3]/div[2]/div/div[2]/div[4]", "Manufactured By"),
                ("/html/body/div/dev[3]/div[3]/div/div[3]/div[2]", "Return Date"),
                ("/html/body/div/div[3]/div[3]/div/div[4]/div[2]", "IOH/SS1 Location"),
                ("/html/body/div/div[3]/div[3]/div/div[5]/div[2]", "Extended Return Date")
            ]

            for xpath, label in data_xpaths:
                try:
                    data_element = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    coach_data[label] = data_element.text.strip() if data_element.text.strip() else "N/A"
                except Exception:
                    coach_data[label] = "N/A"

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
        service = Service(executable_path=CHROMEDRIVER_PATH)
        driver = webdriver.Chrome(service=service)
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