from pathlib import Path
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime
from selenium.common.exceptions import TimeoutException
# C:\Users\stagadmin\CAPITALINE\downloadExcel\consolidate
# \CAPITALINE\consolidateCapitaline\downloadExcel
def excelDownload(browser):
    try: 
        download_path = Path(r"C:\Users\stagadmin\CAPITALINE\consolidateCapitaline\downloadExcel")
        current_date = datetime.now().strftime("%Y%m%d")  # Format: YYYYMMDD
        
        #  Wait for the download to complete (adjust the timeout as needed)
        downloaded_file_path = WebDriverWait(browser, 100).until(
        lambda x: download_path / f"Download.xlsx" in [download_path / f for f in download_path.rglob("*.xlsx")]
        )
        print("step1 =========")

        print(downloaded_file_path)

        # Now, check if the downloaded file is fully available
        if downloaded_file_path :
            # Wait for the download to complete (you can adjust the timeout as needed)
            downloaded_file_path = WebDriverWait(browser, 10).until(
                    lambda x: max(
                        [download_path / f for f in download_path.rglob("*.xlsx")],
                        key=lambda f: f.stat().st_mtime,
                        default=None,
                    )
            )
            print("stem========")
            print(downloaded_file_path.stem)
            # Append the current date to the file name
            new_file_path = downloaded_file_path.with_name(f"{downloaded_file_path.stem}_consolidate_{current_date}{downloaded_file_path.suffix}")

                    
                # Rename the file with the new name
            downloaded_file_path.rename(new_file_path)
            print(new_file_path)
            return new_file_path
            # (Rest of your code...)
        else:
            # Handle the case where the download did not complete successfully
            print("Download not complete or file not found.")
            downloaded_file_path = WebDriverWait(browser, 10).until(
                    lambda x: max(
                        [download_path / f for f in download_path.rglob("*.xlsx")],
                        key=lambda f: f.stat().st_mtime,
                        default=None,
                    )
            )
            print(downloaded_file_path)
            return None
    except Exception as e :
            print (f"exception in folderpath ======",str(e))
            # Append the current date to the file name
            new_file_path = downloaded_file_path.with_name(f"{downloaded_file_path.stem}_consolidate_{current_date}_1{downloaded_file_path.suffix}")

                    
                # Rename the file with the new name
            downloaded_file_path.rename(new_file_path)
            
            return new_file_path #C:\Users\stagadmin\probecapitaline\downloadExcel\Download_202312_1.xlsx
    
            pass
