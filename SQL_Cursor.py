import pyodbc 

def sql_connection_DB_QA():

    try: 
        SQL_ATTR_CONNECTION_TIMEOUT = 113
        login_timeout = 1
        connection_timeout = 3
        conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                            'Server=SERV27\JTEKT;'
                            'Database=Plc_prod;'
                            'UID=quagate;'
                            'PWD=qua+gate;'
                            'Trusted_Connection=No;',
                            timeout = login_timeout, attrs_before = {SQL_ATTR_CONNECTION_TIMEOUT : connection_timeout}
                            )
        conn.timeout = 3
        return conn
        print("QA SQL Connection : OK")
    except:
        print("QA SQL Connection impossible")



def sql_deconnection_DB_QA(cursor):
    try:
        
        cursor.close()

        print("QA SQL Deconnection : OK")
    except:
        print("QA SQL Deconnection impossible")

def sql_connection_DB_Measurlink():

    try: 
        SQL_ATTR_CONNECTION_TIMEOUT = 113
        login_timeout = 1
        connection_timeout = 3
        conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                              'Server=SERV27\JTEKT;'
                              'Database=measurlink8;'
                              'UID=quagate;'
                              'PWD=qua+gate;'
                              'Trusted_Connection=No;',
                              timeout = login_timeout, attrs_before = {SQL_ATTR_CONNECTION_TIMEOUT : connection_timeout}
                            )
        conn.timeout = 3
        return conn
        print("Measurlink SQL Connection : OK")
    except:
        print("Measurlink SQL Connection impossible")


def sql_deconnection_DB_Measurlink(cursor):
    try:
        
        cursor.close()

        print("Measurlink SQL Deconnection : OK")
    except:
        print("Measurlink SQL Deconnection impossible")
    

