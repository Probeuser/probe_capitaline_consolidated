from config import dbconfig

mydb = dbconfig.db_connection()
cursor = mydb.cursor()

def updateDataScrapeLog(filename,status,length,scraped_length,reason,comments,tradeDate):
    insert_qry = "INSERT INTO scraping_log_bse_nse (table_name,status,no_of_data_available,no_of_data_scraped,reason,comments,trade_date) VALUES (%s,%s,%s,%s,%s,%s,%s)"
    try :
       cursor.execute(insert_qry,[filename,status,length,scraped_length,reason,comments,tradeDate])
       mydb.commit()
    except Exception as e :
        print("exception inside fn",str(e))        