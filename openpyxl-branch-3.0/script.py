from openpyxl import load_workbook
import requests
from bs4 import BeautifulSoup
import sys


fname = '../OHSA Fatality Data.xlsx'

wb = load_workbook(filename=fname)
sheet_ranges = wb['Sheet1']

col_pos = 2
scan_count = 0

floop = True
while floop==True:
    cur_value = sheet_ranges['e'+str(col_pos)].value
    
    sys.stdout.write("Scan count: %d \r" % (col_pos))
    sys.stdout.flush()

    if sheet_ranges['H'+str(col_pos)].value != None and sheet_ranges['J'+str(col_pos)].value != None:
        col_pos = col_pos + 1
        continue

    if cur_value == None:
        floop = False
    else:
        cur_link = sheet_ranges['e'+str(col_pos)].hyperlink.display
        page = requests.get(cur_link)

        soup = BeautifulSoup(page.content, 'html.parser')
        results = soup.find(id="maincontain")

        elem = results.find('table', class_='tablei_100 table-borderedi_100')
        try:
            elems = elem.find_all('tr')
        except Exception as inst:
            sys.stdout.write("\n")
            print(inst)
            col_pos = col_pos + 1
            continue
        
        item = elems[2].find('td')
        
        try:
            sheet_ranges['h'+str(col_pos)].value = item.text
        except Exception as inst:
            print("\n IllegalCharacterError")

        try:
            item = elems[3].find('td')
            keywords = item.text.replace(item.find('strong').text, "")
            sheet_ranges['j'+str(col_pos)].value = keywords
        except Exception as inst:
            print("")
            
        try:
            items = elems[5].find_all('td')
            sheet_ranges['i'+str(col_pos)].value = items[6].text
        except Exception as inst:
            print("")

        scan_count = scan_count + 1
        if scan_count == 20:
            wb.save(filename=fname)
            scan_count = 0
        
    

    col_pos = col_pos + 1

sys.stdout.write("\n")
    
wb.save(filename=fname)

print("Successfully completed!")
exit()





# table-responsive
#tablei_100 table-borderedi_100
#tbody h 3rd tr  thumbnail
#      j 4th tr  <strong>Keywords:</strong>...
#      i 6th tr  Occupation 7th td