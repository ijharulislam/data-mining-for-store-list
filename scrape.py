from lxml import html
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re
from datetime import datetime
import json



def download(url):
    print('Downloading: ' + url)
    r = requests.get(url)

    print('Status code: ' + str(r.status_code))

    if r.status_code == requests.codes.ok: 
        return r.content
    else:
        return None



def write_to_excel(workbook,worksheet,storeList):
        
        # w = tzwhere.tzwhere()
        bold = workbook.add_format({'bold': True})
        bold_italic = workbook.add_format({'bold': True, 'italic':True})
        border_bold = workbook.add_format({'border':True,'bold':True})
        border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
        border = workbook.add_format({'border':True,'bold':True})
        
        #worksheet = workbook.add_worksheet('%s_%s'%(a,j))
        worksheet.set_column('B:D', 22)
        worksheet.set_column('E:F', 33)
        row = 0
        col = 0


        worksheet.write(row,col,'Store List',bold)
        row = row + 1

        row = row + 2

        worksheet.write(row,col,'Sl No',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Name',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Formatted Store Name',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Legal Store No.',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Store No.',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Market',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Phone',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Country',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'State',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Zip Code',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Street',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Intersection Description',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Formatted Address',border_bold_grey)


        row = row + 1
        i = 0


        """{u'city': u'Westminster', u'intersectionDescription': u'SEC Beach Blvd & Edinger', 
        u'storeName': u'Westminster', u'formattedStoreName': u'Westminster-Store', u'country': u'United States', 
        u'zipCode': u'92683-7858', u'legalStoreNumber': u'T0249', u'storeNumber': u'249', u'county': u'Orange', 
        u'state': u'CA', u'street': u'16400 Beach Blvd', u'phoneNumber': u'(714) 274-6266', 
        u'formattedAddress': u'16400 Beach Blvd, Westminster, CA 92683-7858', u'typeDescription': u'Store', u'market': u'RG2'}"""

        for output in store_list:
                
            i = i + 1
            col = 0
            worksheet.write(row, col, i, border)
            col = col + 1
            worksheet.write(row, col, output["storeName"] if output.has_key('storeName') else '',border)
            col = col + 1
            worksheet.write(row, col, output["formattedStoreName"] if output.has_key('formattedStoreName') else '',border)
            col = col + 1
            worksheet.write(row, col, output["legalStoreNumber"] if output.has_key('legalStoreNumber') else '',border)
            col = col + 1
            worksheet.write(row, col, output["storeNumber"] if output.has_key('storeNumber') else '',border)
            col = col + 1
            worksheet.write(row, col, output["market"] if output.has_key('market') else '',border)
            col = col + 1
            worksheet.write(row, col, output["phoneNumber"] if output.has_key('phoneNumber') else '',border)
            col = col + 1
            worksheet.write(row, col, output["country"] if output.has_key('country') else '',border)
            col = col + 1
            worksheet.write(row, col, output["state"] if output.has_key('state') else '',border)
            col = col + 1
            worksheet.write(row, col, output["zipCode"] if output.has_key('zipCode') else '',border)
            col = col + 1
            worksheet.write(row, col, output["street"] if output.has_key('street') else '',border)

            col = col + 1
            worksheet.write(row, col, output["intersectionDescription"] if output.has_key('intersectionDescription') else '',border)
            col = col + 1
            worksheet.write(row, col, output["formattedAddress"] if output.has_key('formattedAddress') else '',border)

            col = col + 1
            row = row + 1



def scrap(url):
    def get_url(href):
        return href and re.compile("searchNav=F").search(href)

    domain = 'http://gam.target.com'
    content = download(domain + url)
    if content:
        soup = BeautifulSoup(content, "lxml")
        # print soup
        table = soup.find('div', {"id":"primaryJsonResponse"})
        json_data = table.get_text()
        data = json.loads(str(json_data))
        store_list = data["storeList"] 
        return store_list
    
                

url = '/store-locator/state-result?lnk=statelisting_stateresult&stateCode=CA&stateName=California'
store_list = scrap(url)





# #####             Writing To Excel Sheet    #####

workbook = xlsxwriter.Workbook('example.xlsx')
worksheet = workbook.add_worksheet('Store List')
write_to_excel(workbook,worksheet,store_list)
workbook.close()









