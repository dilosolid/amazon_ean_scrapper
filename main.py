import os
import pandas as pd
import xlrd
import time
import amazonproduct
from amazonproduct.errors import InvalidParameterValue
from amazonproduct.errors import TooManyRequests
import settings
import helper
import sqlite3


conn = sqlite3.connect('data.db')
c = conn.cursor()

euro_locale_list = ['uk', 'de', 'fr', 'it', 'es']
america_locale_list = ['us', 'ca', 'mx']
amazonproduct.HOSTS['mx'] = 'ecs.amazonaws.com.mx'
index = 0


def main(): 
    if os.path.isfile(settings.fileInput) == False:
        print 'File {} does not exist in the current folder'.format(settings.fileInput)
        exit()

    createTable()  
    
    try:            
        df = pd.read_excel(settings.fileInput, sheet_name=settings.sheetName, usecols=[0,1,2,3])
        EAN_list = list(df.itertuples(index=False, name=None))

        for EANObj in EAN_list:
            if settings.do_euro_search:
                dolocale(EANObj, 'uk', euro_locale_list)

            if settings.do_america_search:
                dolocale(EANObj, 'us', america_locale_list)
                        
        if settings.do_euro_search:
            saveListToExcel(settings.fileOutput_euro , 'uk')                          

        if settings.do_america_search:  
            saveListToExcel(settings.fileOutput_america, 'us')                          

    except Exception as err:
        helper.writeToLog('Error in Main: ', '01', err)  

def dolocale(EANObj, mainlocale, locale_list):
    global index    
    EAN = ''
    Description = ''
    Volume = ''
    Brand = ''

    try:    
        if EANObj[0]:
            EAN = str(EANObj[0])            
        if EANObj[1]:
            Description = EANObj[1]
        if EANObj[2]:
            Volume = EANObj[2]
        if EANObj[3]:
            Brand = EANObj[3]

        if len(EAN) > 0:
            index = index + 1
            print '{} - Searching EAN: {}'.format(index, EAN)                                                             
            doRequest = True
            for locale in locale_list:        
                if itemExitInDB(EAN,locale):
                    print 'EAN: {} And Locale: {} already exist in DB skipping Amazon APi Request'.format(EAN,locale)
                    break

                result = doAmazonApiRequest(EAN, Description, Volume, Brand, locale, doRequest, mainlocale)
                if len(result) > 0 and len(result[0]['ASIN']) > 0 and locale == mainlocale:
                    doRequest = True
                elif locale == mainlocale:
                    doRequest = False                
    except Exception as err:
        helper.writeToLog('Error in dolocale: ', '02', err)          

def doAmazonApiRequest(EAN, Description, Volume, Brand, locale, doRequest, mainlocale):
    result_list = []
    time.sleep(settings.wait_between_api_request)        
    try:  
        if locale != mainlocale:
            EAN_temp = ''
            Description_temp = ''
            Volume_temp = ''
            Brand_temp = ''             
            Qty_temp = ''
        else:
            EAN_temp = EAN
            Description_temp = Description
            Volume_temp = Volume
            Brand_temp = Brand 
            Qty_temp = '0'
            
        result = {'EAN':EAN_temp, 'Description':Description_temp, 'Volume':Volume_temp, 'Brand':Brand_temp, 'Qty':Qty_temp , 'ASIN':'', 'Locale':locale, u'Title':''}                                    

        if doRequest == False:             
            InsertRowToDB(result, mainlocale)                                                    
            result_list.append(result)
            return result_list
        
        api = amazonproduct.API(cfg=getConfig(locale))    
        lookup_result = api.item_lookup(EAN, IdType='EAN', SearchIndex='All' )

        QtyCount = 0
        firstTitle = ''
        firstASIN = ''
        for item in lookup_result.Items.Item:
            QtyCount = QtyCount + 1
            if firstTitle == '':
                firstTitle = item.ItemAttributes.Title.text        
              
            if locale != mainlocale:                
                firstASIN = ''
            else:                    
                if firstASIN == '':
                    firstASIN = item.ASIN                

        Qty_temp = QtyCount

        if locale == mainlocale:                
            Qty_temp = str(QtyCount)
        else:
            Qty_temp = ''
        
        print 'Result From API:  Title: (%s) - ASIN: (%s) - Locale: (%s)' % (item.ItemAttributes.Title.text.encode("utf-8"), item.ASIN, locale)                                         
        result = {'EAN':EAN_temp, 'Description':Description_temp, 'Volume':Volume_temp, 'Brand':Brand_temp, 'Qty':Qty_temp , 'ASIN':firstASIN, 'Locale':locale, u'Title':firstTitle}   
        InsertRowToDB(result, mainlocale)                             
        result_list.append(result)

    except InvalidParameterValue, err:
        message = 'EAN: {} not found in locale {}'.format(EAN,locale)
        print message
        result['Title'] = ''
        InsertRowToDB(result, mainlocale)                             
        result_list.append(result)
        return result_list
    except TooManyRequests:
        print 'Amazon is throttling api request sleeping for 3 seconds'
        time.sleep(3)
        result_list = doAmazonApiRequest(EAN, Description, Volume, Brand, locale, True, mainlocale)
    except Exception as err:            
        helper.writeToLog('Error in doAmazonApiRequest', '03', err)                

    return result_list

def saveListToExcel(fileOutput, mainLocale):    
    try:
        #df = pd.DataFrame(result_list, columns=['EAN','Description','Volume','Brand','ASIN','Locale','Title'])
        df = pd.read_sql_query(sql='SELECT EAN, Description, Volume, Brand, Qty, ASIN, Locale, Title FROM items WHERE MainLocale = ? ORDER BY ID ASC', con=conn, params=[mainLocale])

        writer = pd.ExcelWriter(fileOutput)
        df.to_excel(writer,'Sheet1', index=False)
        writer.save()        
    except Exception as err:
        helper.writeToLog('Error in saveListToExcel', '04', err)        

def getConfig(locale):
    config = {
        'access_key':    settings.access_key,
        'secret_key':    settings.secret_key,
        'associate_tag': settings.associate_tag,
        'locale': locale
    }
    return config

def InsertRowToDB(data,mainlocale):
    try:        
        c.execute("INSERT INTO items (EAN, Description, Volume, Brand, Qty, ASIN, Locale, Title, mainlocale) VALUES (?,?,?,?,?,?,?,?,?)", (data['EAN'], data['Description'], data['Volume'], data['Brand'], data['Qty'], str(data['ASIN']), data['Locale'], data['Title'], mainlocale))
        conn.commit()
    except Exception as err:        
        helper.writeToLog('Error in InsertRowToDB: ', '01', err) 

def createTable():
    try:        
        c.execute('CREATE TABLE IF NOT EXISTS items (ID INTEGER PRIMARY KEY AUTOINCREMENT, EAN TEXT, Description TEXT, Volume TEXT, Brand TEXT, Qty TEXT, ASIN TEXT, Locale TEXT, Title TEXT, MainLocale TEXT)')
    except Exception as err:        
        helper.writeToLog('Error in createTable: ', '01', err) 

def itemExitInDB(EAN, Locale):
    exist = False
    try:        
        c.execute("SELECT id FROM items WHERE EAN = ? and Locale = ? LIMIT 1", (EAN, Locale))
        rows = c.fetchall()

        for row in rows:
            exist = True

    except Exception as err:        
        helper.writeToLog('Error in itemExitInDB: ', '01', err) 

    return exist

if __name__ == "__main__":
	main()
