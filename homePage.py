import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,NoSuchWindowException,UnexpectedAlertPresentException
from  Function.testotp import getOtp
from Function.downloadFolderPath import excelDownload
from Function.readExcelIntoDb import readExcelIntoDb
from Function.data_scraping_log_function import updateDataScrapeLog
import Function.mail as mail
import traceback

# C:\Users\stagadmin\CAPITALINE\consolidateCapitaline\downloadExcel
chrome_driver_path = 'C:\\Chrome\\chrome-win64\\chrome.exe'
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = 'C:\\Chrome\\chrome-win64\\chrome.exe'
prefs = {
    "download.default_directory": "C:\\Users\\stagadmin\\CAPITALINE\\consolidateCapitaline\\downloadExcel",
    "download.default_content_setting_value": 2,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

# chrome_options.add_argument("--headless")

browser = webdriver.Chrome(options=chrome_options)
url = "https://awsone.capitaline.com/"

browser.get(url)




# Allow some time for the page to load
time.sleep(10)

# Locate the username and password fields and provide values
username_field = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-ng-model="loginData.userName"]'))
)
username_field.send_keys("listedresearch@probeinformation.com")

password_field = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-ng-model="loginData.password"]'))
)
password_field.send_keys("cap$pro@355")

# Locate the login button and click on it
login_button = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'button[data-ng-click="login()"]'))
)
login_button.click()

# Locate the OTP field and enter the value "1234"
otp_field = WebDriverWait(browser, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-ng-model="otp"]'))
)
time.sleep(150)

otp = getOtp()
print(f"otp=============",otp)
otp_field.send_keys(otp)

try:

# Locate and click the "Verify OTP" button
    verify_otp_button = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.ID, 'btnsubotp'))
    )
    verify_otp_button.click()


    screener_button = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//a[@href="Screeener.html"]'))
    )
    screener_button.click()



# Locate the dropdown for selecting Peerset
# Locate the "Common Peerset" option and click on it
    peer_set_dropdown = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//select[@ng-model="PeerSetOpt"]'))
    )
    peer_set_dropdown.click()
    common_peer_set_option = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//select[@ng-model="PeerSetOpt"]/option[text()="Common Peerset"]'))
    )
    common_peer_set_option.click()


# Locate the dropdown for selecting Common PeerSet
# Locate the "NSE Only" option and click on it using XPath
    common_peer_set_dropdown = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//select[@ng-model="COPeerSet"]'))
    )
    common_peer_set_dropdown.click()
    nse_only_option = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//select[@ng-model="COPeerSet"]/option[contains(@ng-repeat, "SelectedPeer")][contains(text(), "NSE Only")]'))
    )
    nse_only_option.click()


# Locate the "Filter" link using CSS selector and wait until it is present
    filter_link = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="#/filter"]'))
    )
    filter_link.click()

# Locate the dropdown for selecting the table
# Locate the "Results Assets & Liabilities (Revised)" option and click on it
    results_assets_liabilities_option = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//select[@id="select1"]/option[text()="Results Assets & Liabilities(Revised)"]'))
    )
    results_assets_liabilities_option.click()


# # Locate the "Half Yearly (Assets & Liab. Revised)" option and click on it
# Locate the span element under the innerscroll div with the specified id
    half_yearly_option = WebDriverWait(browser, 30).until(
    EC.presence_of_element_located((By.XPATH, '//span[@id="HalfYearlyAssetsLiab.RevisedHYRAL_RV"]'))
    )

    # Scroll the element into view (if needed)
    browser.execute_script("arguments[0].scrollIntoView();", half_yearly_option)

    # Click the element using JavaScript
    browser.execute_script("arguments[0].click();", half_yearly_option)



# Locate the checkbox
    checkbox = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.ID, 'chkComp'))
    )
    checkbox.click()

# Locate the radio button for "consolidate"
    standalone_radio_button = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//input[@value="C" and @ng-model="OPT.priority"]'))
    )
    standalone_radio_button.click()

# # Locate the "Check All" checkbox
# Wait for the checkbox to be present
    check_all_checkbox = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//input[@class="noMargin topMargin ng-valid ng-not-empty ng-dirty ng-valid-parse ng-touched" and @type="checkbox"]'))
    )

    # Check the checkbox if it's not already checked
    if not check_all_checkbox.is_selected():
        check_all_checkbox.click()

# Locate the "Output" button
    output_button = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.XPATH, '//a[@href="#/Output"]'))
    )
    output_button.click()

    
#  locate excel
    try:
        target_element = WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'i.fa.fa-file-excel-o.green'))
        )

        # Scroll the element into view
        browser.execute_script("arguments[0].scrollIntoView();", target_element)

        # Click the element using JavaScript
        browser.execute_script("arguments[0].click();", target_element)


        print("after excel button =======")

        download_excel_file_path = excelDownload(browser)

        print(f"successfully download excel in the path =====", str(download_excel_file_path))

        readExcelIntoDb(download_excel_file_path)

    except UnexpectedAlertPresentException as e:
        # Alert Text: Please select at least one column .handled this alert box here
        print(f"Exception  excel button ======: {str(e)}")
        alert = browser.switch_to.alert
        alert.accept()  # or alert.dismiss() depending on your requirement
        profile_link = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='profile']//span[text()='Welcome Vasteam']"))
        )

        # Click the "Welcome Vasteam" profile link
        profile_link.click()


        # Wait for the "Logout" element to be clickable
        logout_element = WebDriverWait(browser, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@class='profileWrap' and @style='display: block;']//a[text()='Logout']"))
        )

        # Click the "Logout" element
        logout_element.click()
        mail.send_email("Capitaline Consolidate SCRAP DATA from web site",str(e))
        status = "failure"
        length = "NA"
        scraped_length = "NA"
        reason = "Alert Text: Please select at least one column "
        tradeDate = ""
        comments = ""
        updateDataScrapeLog("Capitaline_consolidated",status,length,scraped_length,reason,comments,tradeDate)  
        browser.quit()

except Exception as e : 
    # handled in cant locate any of the button and 
    print(f"Commonexception =============",str(e))
    mail.send_email("Capitaline Consolidate SCRAP DATA from web site",str(e))
    status = "failure"
    length = "NA"
    scraped_length = "NA"
    reason = "Invalid otp "
    tradeDate = ""
    comments = ""
    updateDataScrapeLog("Capitaline_consolidated",status,length,scraped_length,reason,comments,tradeDate) 
    print(traceback.format_exc())
    # Wait for the "Welcome Vasteam" profile link to be clickable
    profile_link = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='profile']//span[text()='Welcome Vasteam']"))
    )

    # Click the "Welcome Vasteam" profile link
    profile_link.click()


    # Wait for the "Logout" element to be clickable
    logout_element = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='profileWrap' and @style='display: block;']//a[text()='Logout']"))
    )

    # Click the "Logout" element
    logout_element.click()
    mail.send_email("Capitaline Consolidate SCRAP DATA from web site",str(e))
    browser.quit()


# Wait for the "Welcome Vasteam" profile link to be clickable
profile_link = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='profile']//span[text()='Welcome Vasteam']"))
    )

# Click the "Welcome Vasteam" profile link
profile_link.click()


# Wait for the "Logout" element to be clickable
logout_element = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='profileWrap' and @style='display: block;']//a[text()='Logout']"))
    )

# Click the "Logout" element
logout_element.click()
browser.quit()



