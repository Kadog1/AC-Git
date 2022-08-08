# -*- coding: utf-8 -*-
"""
Created on Mon Feb 21 12:04:37 2022

@author: Paul Ferdinand Popp, Assistant, CAD
"""


import getRegisterScreenshot_Helpers as helpers
import getRegisterScreenshot_Registers as registers
import logging

import pyodbc
import pandas as pd
from datetime import datetime


isProductRun = False
if isProductRun == True:
    tTEST = ""
    Testumgebung = ""
else:
    tTEST = "_TEST"
    Testumgebung = " Testumgebung"

 
timePrefix = str(datetime.strftime(datetime.now().replace(second=0, microsecond=0), '%Y%m%d'))
logging.basicConfig(filename=r'\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich' + Testumgebung + '\\F Screenshots\\00_LogFiles\\' + timePrefix +'_getRegisterScreenshot.log',
                        format='%(asctime)s - %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S', level=logging.INFO)
summary_logger = helpers.setup_logger('summary_logger', r'\\Defrnappfl101.ey.net\101fra00010\T\TCC_SB\Z_Archive\Adressabgleich' + Testumgebung + '\\F Screenshots\\00_LogFiles\\' + timePrefix + '_getRegisterScreenshot_Summary.log')


try:
    connSQLServer = pyodbc.connect('Driver={ODBC Driver 11 for SQL Server};'
                          'Server=DERUSCMPDWASQ01.ey.net\INST02;'
                          'Database=CAD;'
                          'Trusted_Connection=yes;')
    cursor = connSQLServer.cursor()
    
    # test units
    if isProductRun == True:
        cursor.execute("SELECT OrderNo FROM [CAD].[dbo].[tCON_Orderbook{0}] WHERE AC_Status = 'InputDataAvailable'".format(tTEST))
    else:
        cursor.execute("SELECT OrderNo FROM [CAD].[dbo].[tCON_Orderbook{0}] WHERE OrderNo in ('CON0000029549','CON0000030977','CON0000030952','CON0000030789','CON0000030933','CON0000030935','CON0000030964','CON0000030883')".format(tTEST))
        # cursor.execute("SELECT OrderNo FROM [CAD].[dbo].[tCON_Orderbook{0}] WHERE OrderNo in ('CON0000030952','CON0000030892', 'CON0000030854')".format(tTEST))
        # cursor.execute("SELECT OrderNo FROM [CAD].[dbo].[tCON_Orderbook{0}] WHERE OrderNo in ('CON0000030898', 'CON0000030735', 'CON0000030892', 'CON0000030854', 'CON0000030716', 'CON0000030798', 'CON0000030799')".format(tTEST))
    
    
    columns = [column[0] for column in cursor.description]
    dataRecordSet = []
    for row in cursor.fetchall():
        row_to_list = [elem for elem in row]
        dataRecordSet.append(row_to_list)
    dfRecordSet = pd.DataFrame(dataRecordSet)
    
    if len(dfRecordSet) > 0:
        dfRecordSet.columns = columns
        print(str(len(dfRecordSet)) + " OrderNos found.") 
        for i in range(len(dfRecordSet)):
            orderNo = dfRecordSet['OrderNo'].iloc[i]
            print("   " + str(i + 1) + ": " + orderNo)
            logging.info("----- START getRegisterScreenshot.py - " + orderNo + " - isProduction " + str(isProductRun) + " -----")
            df = helpers.getDFWorkbook(orderNo)       
            if df is not None:
                summary_logger.info("{0}: Addresses {1}".format(orderNo, len(df)))
                for index, row in df.iterrows():
                    message = "Looking for {2}: {0} in {1}...".format(row.loc['Company'], row.loc['Location'], row.loc['Country'].upper())
                    print(message)
                    logging.info(message)
                    foundHandelsregister = False
                    if row.loc['Country'].upper() == 'DE':
                            foundHandelsregister = registers.findHandelsregister(row.loc['Company'], row.loc['Location'], row.loc['Country'].upper(), row.loc['OrderNo'], row.loc['index'] + 1, row.loc['timeStamp'])
                    if foundHandelsregister == False:
                        foundDBHoovers = registers.findDBHoovers(row.loc['Company'], row.loc['Location'], row.loc['Country'].upper(), row.loc['OrderNo'], row.loc['index'] + 1, row.loc['timeStamp'])
            now = datetime.strftime(datetime.now().replace(second=0, microsecond=0), '%Y-%m-%d %H:%M')
            if isProductRun:
                strSQLStatus = "UPDATE [CAD].[dbo].[tCON_Orderbook{0}] SET AC_Status = 'InputDataReceived' WHERE OrderNo = '".format(tTEST) + orderNo + "'"
                strSQLts = "UPDATE [CAD].[dbo].[tCON_Orderbook{0}] SET tsInputDataReceived = '".format(tTEST) + now + "' WHERE OrderNo = '" + orderNo + "' AND tsInputDataReceived IS NULL"
                cursor.execute(strSQLStatus)
                cursor.execute(strSQLts)
                cursor.commit()
    cursor.close()
    connSQLServer.cursor()
    logging.shutdown()
    
except Exception as e:
    logging.error('Error in Main: '+ str(e))
    logging.shutdown()
    raise



