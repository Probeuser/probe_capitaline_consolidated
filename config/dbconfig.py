import mysql.connector

def db_connection():
    mydb = mysql.connector.connect(
        host="localhost",
        user="root1",
        password="Mysql1234$",
        database="bse"
    )

    return mydb