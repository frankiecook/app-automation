# Job Application Automation
#
# @author Frankie Cook - github.com/frankiecook
# @version python3.12
# @since 0.0
# @date 4/24/2024
#
# To-Do List
# - undo temporary solution that doesn't certify ssl
# - add other websites (only works with LinkedIn)
# - create backup of excel file
# - create backup of save data
# - GUI addition
# - check if job already exists in spreadsheet
# - deeper website interaction (highlight selections)
# - blacklist
# - safety fails for so many checks (missing sheet, missing data, url fail)
# - output successful save info (company, place, job, line)
# - limit retrys for url 
# - ampersand errors
#
# imports
from urllib import request
from datetime import date
import ssl
import openpyxl
import time

###########
# VARIABLES
###########
file_name = 'ApplicationTracker.xlsx'
sheet = 'ATLA'
date_applied = date.today().strftime("%m/%d/%y")
search_location = 'Atlanta, GA'
search_terms = 'software engineer'
website = 'LinkedIn'

###########
# FUNCTIONS
###########
# checks if worksheet has data at given cell
# ws : worksheet
# cell : cell location
def hasData(ws, cell):
    # check for data
    check_data = ws[cell].value

    if check_data != None:
        return True
    return False

# open and print request as html
#with request.urlopen(req) as f:
#    html = f.read().decode('utf-8')
def openHTML(req):
    try:
        f=request.urlopen(req)
        html_out = f.read().decode('utf-8')
        return html_out
    except:
        print('URL Failed. Waiting to Try Again.')
        time.sleep(1)   # seconds
        return openHTML(req)

def grabURL():
    global url 

    # define url and request abstraction
    url = str(input("Enter URL: "))

    # header tricks server
    try:
        req_out = request.Request(url, headers={'User-Agent': 'Mozilla'})
        return req_out
    except:
        print("Bad URL. Try Again.")
        return grabURL()

################
# READ SAVE FILE
################
# open file
file = open("data.txt",'r')
data = file.read()
data_lines = file.readlines()
file.close()
# split data into list
data = data.replace('\n',',')
data_list = data.split(",")

# find starting row
cur_row_i = data_list.index(sheet) + 1
cur_row = int(data_list[cur_row_i])

##################
# SCRAPE GIVEN URL
##################
# temporary solution to failed certs
ssl._create_default_https_context = ssl._create_unverified_context

# loop
usr_iput = ""
while(1):
    req = grabURL()
    html = openHTML(req)
        
    # PARSE HTML
    ############
    # search html for values (company, position, actual location, search terms, link)
    offset = len('<title>')
    index_start = html.find('<title>') + offset
    index_end = html.find('</title>')

    # example of scrap
    # <company> hiring <position> in <actual location> | LinkedIn
    scrap = html[index_start:index_end]

    # company
    ci_start = 0
    ci_end = scrap.find(' hiring ')
    company = scrap[ci_start:ci_end]
    # position
    offset_pis = len(' hiring ')
    pi_start = scrap.find(' hiring ') + offset_pis
    pi_end = scrap.find(' in ')
    position = scrap[pi_start:pi_end]
    # actual location
    offset_alis = len(' in ')
    ali_start = scrap.find(' in') + offset_alis
    ali_end = scrap.find(' | ')
    actual_location = scrap[ali_start:ali_end]

    # WRITE RESULT TO SPREADSHEET
    #############################
    # variables
    row = str(cur_row)

    # open workbook and sheet
    # data_only : determines either the formula or last value stored
    wb=openpyxl.load_workbook(file_name, data_only=True)
    ws= wb[sheet]
    
    # spreadsheet cell : data
    application = {
        'A'+row:str(company),
        'B'+row:str(position),
        'C'+row:str(search_location),
        'D'+row:str(actual_location),
        'E'+row:str(date_applied),
        #'F'+row:str(),
        'G'+row:str(search_terms),
        'H'+row:str(url),
        'I'+row:str(website)
    }

    # check if row has any data for each cell
    for cell in application:
        # if row contains data, then exit program
        if hasData(ws, cell):
            print("ERROR: "+cell+" has data. Check saved row in data.txt")

            wb.close()
            exit()

    # save application to row
    for cell in application:
        ws[cell]=application[cell]
    cur_row+=1

    # save file
    wb.save('ApplicationTracker-New.xlsx')
    wb.close()

    # save current row
    file = open("data.txt",'r')
    data_lines = file.readlines()
    file.close()

    for i in range(len(data_lines)):
        line = data_lines[i]

        # -1 if not found
        if line.find(sheet) != -1:
            # modify line, store current row
            new_line = sheet+","+str(cur_row)+"\n"
            data_lines[i] = new_line
    
    with open("data.txt",'w') as file:
        file.writelines(data_lines)
        file.close()
