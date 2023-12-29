import pandas as pd
import re
import Function.mail as mail
from datetime import datetime
from config import dbconfig
from Function.data_scraping_log_function import updateDataScrapeLog
import traceback
from datetime import datetime, timedelta
current_date = datetime.now().strftime("%Y-%m-%d")

def readExcelIntoDb(excel_file_path) :
    # initializing cursor for db
    mydb = dbconfig.db_connection()
    db_cursor = mydb.cursor()

    try : 
        # Specify the substring you are looking for
        substring_to_find_capitaline_screener = 'Capitaline Screener >> Consolidated >> \(Rs In Crores\s*\)'
        substring_to_find_disclaimer = 'Disclaimer: Capitaline Database has taken due care and caution'


        # Load the Excel sheet
        # Downloaded_File_Path = r"C:\Users\stagadmin\probecapitaline\downloadExcel\downloadExcel/Download_20231220.xlsx"
        Downloaded_File_Path = rf"{excel_file_path}"
        # print(Downloaded_File_Path)
        df_copy = pd.read_excel(Downloaded_File_Path, header=None)
        df = pd.read_excel(Downloaded_File_Path, header=None)
        # [182 rows x 46 columns]


        # Find the  row containing the specified substring (case-insensitive)
        filtered_rows_capitaline_screener = df_copy[df_copy[0].str.contains(substring_to_find_capitaline_screener, na=False, case=False, regex=True)]
        # print (filtered_rows_capitaline_screener)

        filtered_rows_disclaimer = df_copy[df_copy[0].str.contains(substring_to_find_disclaimer, na=False, case=False, regex=True, flags=re.DOTALL)]
        # print (filtered_rows_disclaimer)

        # Find the  index containing the specified substring (case-insensitive)
        if not filtered_rows_capitaline_screener.empty:
            subtitle_index = filtered_rows_capitaline_screener.index[0]
            # print("Subtitle index:", subtitle_index)
        else:
            print("Substring not found in the DataFrame.")

        if not filtered_rows_disclaimer.empty:
            disclaimer_index = filtered_rows_disclaimer.index[0]
            # print("Disclaimer index:", disclaimer_index)
        else:
            print("Disclaimer not found in the DataFrame.")    

        # subtitle_index = 0 , disclaimer_index=181
        # 0  Capitaline Screener >> Standalone >> (Rs In Cr...  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN  ...  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN
        # 181  Disclaimer: Capitaline Database has taken due ...  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN  ...  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN  NaN      # 
        # range == 1:181 we will get 1 to 181-1 records 
        # [180 rows x 46 columns]
        # Create a new DataFrame from the row next to the title
        df_new = df.iloc[subtitle_index + 1:disclaimer_index]
        # Set the column names using the values from the row containing column names
        # # print("new dataframe =======")
        # print(df_new)
        # print(len(df_new.columns))
        # print(len(df_new))

        # Assuming your_column_names is a list of column names
        column_names = [
            'NatureReport',
            'CapitalineCode',
        'CompanyName',
        'YearEnd',
        'ResultType',
        'NoOfMonths',
        'ShareCapital',
        'ReservesAndSurplus',
        'MoneyReceivedAgainstShareWarrants',
        'ShareApplicationMoneyPendingAllotment',
        'Total_NCL',
        'NCL_Long_Term_Borrowings',
        'NCL_Deferred_Tax_Liabilities_Net',
        'NCL_Other_Long_Term_Liabilities',
        'NCL_Long_Term_Provisions',
        'Total_CL',
        'CL_ShortTermBorrowings',
        'CL_TradePayables',
        'CL_OtherCurrentLiabilities',
        'CL_ShortTermProvisions',
        'OtherEquityAndLiabilities',
        'TotalEquityAndLiabilities',
        'Total_NCA',
        'NCA_FixedAssetsInclCWP',
        'NCA_TangibleAssets',
        'NCA_IntangibleAssets',
        'NCA_IntangibleAssetsUnderDevelopmentRandD',
        'NCA_CWP',
        'NCA_NonCurrentInvestments',
        'NCA_DeferredTaxAssetNet',
        'NCA_LongTermLoansAndAdvances',
        'NCA_OtherNonCurrentAssets',
        'TotalCurrentAssets',
        'TCA_CurrentInvestments',
        'TCA_Inventories',
        'TCA_TradeReceivables',
        'TCA_CashAndCashEquivalents',
        'TCA_ShortTermLoansAndAdvances',
        'TCA_OtherCurrentAssets',
        'OtherAssets',
        'TotalAssets',
        'TA_DebtEquityRatio',
        'TA_CurrentRatio',
        'TA_DebtorsVelocity',
        'TA_CreditorsVelocity',
        'TA_RONW',
        'TA_ROCE',
        ]

        update_columns = [
             'NatureReport',
            'CapitalineCode',
        'CompanyName',
        'YearEnd',
        'ResultType',
        'NoOfMonths',
        'ShareCapital',
        'ReservesAndSurplus',
        'MoneyReceivedAgainstShareWarrants',
        'ShareApplicationMoneyPendingAllotment',
        'Total_NCL',
        'NCL_Long_Term_Borrowings',
        'NCL_Deferred_Tax_Liabilities_Net',
        'NCL_Other_Long_Term_Liabilities',
        'NCL_Long_Term_Provisions',
        'Total_CL',
        'CL_ShortTermBorrowings',
        'CL_TradePayables',
        'CL_OtherCurrentLiabilities',
        'CL_ShortTermProvisions',
        'OtherEquityAndLiabilities',
        'TotalEquityAndLiabilities',
        'Total_NCA',
        'NCA_FixedAssetsInclCWP',
        'NCA_TangibleAssets',
        'NCA_IntangibleAssets',
        'NCA_IntangibleAssetsUnderDevelopmentRandD',
        'NCA_CWP',
        'NCA_NonCurrentInvestments',
        'NCA_DeferredTaxAssetNet',
        'NCA_LongTermLoansAndAdvances',
        'NCA_OtherNonCurrentAssets',
        'TotalCurrentAssets',
        'TCA_CurrentInvestments',
        'TCA_Inventories',
        'TCA_TradeReceivables',
        'TCA_CashAndCashEquivalents',
        'TCA_ShortTermLoansAndAdvances',
        'TCA_OtherCurrentAssets',
        'OtherAssets',
        'TotalAssets',
        'TA_DebtEquityRatio',
        'TA_CurrentRatio',
        'TA_DebtorsVelocity',
        'TA_CreditorsVelocity',
        'TA_RONW',
        'TA_ROCE',
        ]

        # Create a string with column names separated by commas
        columns_str = ', '.join(column_names)

        # Your SQL INSERT query template with placeholders for values
        insert_query = f"INSERT INTO capitaline_new_download ({columns_str}, scraping_date) VALUES ({', '.join(['%s'] * len(column_names))}, %s)"


        # Your SQL SELECT query template
        select_query = "SELECT * FROM capitaline_new_download WHERE NatureReport = %s and  CapitalineCode = %s  and YearEnd = %s and ResultType = %s LIMIT 1"

        # Your SQL UPDATE query template
        update_query = """
            UPDATE capitaline_new_download
            SET
            NatureReport = %s,
            CapitalineCode = %s,
            CompanyName = %s,
            YearEnd = %s,
            ResultType = %s,
            NoOfMonths = %s,
            ShareCapital = %s,
            ReservesAndSurplus = %s,
            MoneyReceivedAgainstShareWarrants = %s,
            ShareApplicationMoneyPendingAllotment = %s,
            Total_NCL = %s,
            NCL_Long_Term_Borrowings = %s,
            NCL_Deferred_Tax_Liabilities_Net = %s,
            NCL_Other_Long_Term_Liabilities = %s,
            NCL_Long_Term_Provisions = %s,
            Total_CL= %s,
            CL_ShortTermBorrowings = %s,
            CL_TradePayables = %s,
            CL_OtherCurrentLiabilities = %s,
            CL_ShortTermProvisions = %s,
            OtherEquityAndLiabilities = %s,
            TotalEquityAndLiabilities = %s,
            Total_NCA = %s,
            NCA_FixedAssetsInclCWP = %s,
            NCA_TangibleAssets = %s,
            NCA_IntangibleAssets = %s,
            NCA_IntangibleAssetsUnderDevelopmentRandD = %s,
            NCA_CWP = %s,
            NCA_NonCurrentInvestments = %s,
            NCA_DeferredTaxAssetNet = %s,
            NCA_LongTermLoansAndAdvances = %s,
            NCA_OtherNonCurrentAssets = %s,
            TotalCurrentAssets = %s,
            TCA_CurrentInvestments = %s,
            TCA_Inventories = %s,
            TCA_TradeReceivables = %s,
            TCA_CashAndCashEquivalents = %s,
            TCA_ShortTermLoansAndAdvances = %s,
            TCA_OtherCurrentAssets = %s,
            OtherAssets = %s,
            TotalAssets = %s,
            TA_DebtEquityRatio = %s,
            TA_CurrentRatio = %s,
            TA_DebtorsVelocity = %s,
            TA_CreditorsVelocity = %s,
            TA_RONW = %s,
            TA_ROCE = %s,
            scraping_date = %s
            WHERE NatureReport = %s and CapitalineCode = %s and YearEnd = %s and ResultType = %s
        """


        scraped_length = 0
        update_length = 0
        for iterate in range(subtitle_index + 2, disclaimer_index):
            scraped_length = scraped_length + 1
            # Skip the header/title row and extract the data
            
            row_data = df_new.iloc[iterate -  1].tolist()
            row_data.insert(0,"C")
            # Append scraping_date to the row_data
            row_data.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

            # Check if CapitalineCode already exists
            with mydb.cursor() as db_cursor:
                db_cursor.execute(select_query, (row_data[0],row_data[1],row_data[3],row_data[4]))
                existing_row = db_cursor.fetchone()

            if existing_row:
                update_data = [row_data[column_names.index(column)] for column in update_columns]
                rearrange_existing_row = list(existing_row[:-1])
                # If row exists, perform UPDATE
                if rearrange_existing_row != update_data :
                    with mydb.cursor() as db_cursor:
                        db_cursor.execute(update_query, row_data[0:] + [row_data[0],row_data[1],row_data[3],row_data[4]])
                        mydb.commit()
                    print(f"Successfully updated row for {row_data[1]}")
                    update_length = update_length + 1
                else :
                     print(f"already updated for {row_data[1]}")     
            else:
                # If row does not exist, perform INSERT
                with mydb.cursor() as db_cursor:
                    db_cursor.execute(insert_query, row_data)
                    mydb.commit()
                print(f"Successfully inserted row for {row_data[1]}")
                update_length = update_length + 1
        status = "success"
        length = str(len(df_new)-1)
        tradeDate =str(current_date)
        if update_length > 0 :
             comments = "we have found "+ str(update_length) + " new records"
        else :              
            comments = ""
        reason = ""
        updateDataScrapeLog("Capitaline_consolidated",status,length,scraped_length,reason,comments,tradeDate) 
        print(f"Job done successfully data written on db for consolidate  ========== ,{scraped_length},{update_length}")
    except FileNotFoundError as e :
        print(traceback.format_exc())
        print(f"exception ======",str(e))
        status = "failure"
        length = "NA"
        scraped_length = "NA"
        reason = "No such file or directory "
        tradeDate = ""
        comments = ""
        updateDataScrapeLog("Capitaline_consolidated",status,length,scraped_length,reason,comments,tradeDate)  
        mail.send_email("Capitaline consolidate SCRAP DATA  from excel sheet ",reason ) 
    except Exception as e :
        print(f"exception ======",str(e))
        # Duplicate entry '71769-202309' 
        status = "failure"
        length = "NA"
        scraped_length = "NA"
        reason = "Error occurred"
        tradeDate = ""
        comments = ""
        updateDataScrapeLog("Capitaline_consolidated",status,length,scraped_length,reason,comments,tradeDate)  
        mail.send_email("Capitaline consolidate SCRAP DATA  from excel sheet ",str(e)) 