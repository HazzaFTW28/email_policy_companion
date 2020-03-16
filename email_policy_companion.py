from tkinter import *
import time
import pandas as pd
from pandas import ExcelWriter
import xlwings as xw
from win32com.client import Dispatch
import win32com.client
from win32com.client import DispatchEx
import re
from datetime import datetime
import sys, time
import datetime
from datetime import datetime
from datetime import timedelta
import psutil
import os
import subprocess
import sys
import openpyxl
import pandas as pd
import xlwings as xw
from pandas import ExcelWriter
import shutil
import subprocess
from win32com.client import Dispatch
import win32com.client
from win32com.client import DispatchEx
import re
from datetime import datetime
import sys, time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By 
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
from selenium.webdriver.chrome.options import Options
import keyboard
from pynput.keyboard import Key, Controller
from selenium.webdriver.common.action_chains import ActionChains
from tkinter import *

class policy_infomation():
    def __init__(self,lock, old_message, white_list, number_claims, protected_before, ncb_before, canceled, end_date, count, valid, policy_number, postcode, ncb, protected, start_date, renewal_date, name, number_plate, address_1, address_2, address_3, address_4, claim_number, at_fault, bonus_affected, status, circumstances):
        self.policy_number = policy_number
        self.postcode = postcode
        self.ncb = ncb
        self.protected = protected
        self.start_date = start_date
        self.renewal_date = renewal_date
        self.name = name
        self.number_plate = number_plate
        self.address_1 = address_1
        self.address_2 = address_2
        self.address_3 = address_3
        self.address_4 = address_4
        self.claim_number = claim_number
        self.at_fault = at_fault
        self.bonus_affected = bonus_affected
        self.status = status
        self.valid = valid
        self.count = count
        self.end_date = end_date
        self.canceled = canceled
        self.ncb_before = ncb_before
        self.number_claims = number_claims
        self.protected_before = protected_before
        self.white_list = white_list
        self.old_message = old_message
        self.lock = lock
p = policy_infomation
p.white_list = ''
def update_notes(note):
    variable3.set(str(note))
    root.update()



# if the user wants to, this function will login to prestige for you and add to the underwriting notes for you.
def underwriting_notes():
    options = Options()
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--headless')
    options.add_argument('--disable-gpu') 
    options.add_argument("--window-size=1920,1080")
    options.add_argument('--safebrowsing-disable-download-protection')
    options.add_experimental_option('prefs', { "download.default_directory": r'U:\Harrys Folder\webscraping', "download.prompt_for_download": False, "plugins.always_open_pdf_externally": True } )
    driver = webdriver.Chrome(r'U:\Harrys Folder\python\chromedriver.exe', options=options)
    driver.get('https://www.prestigeunderwriting.com/Admin/System/Recall/GlobalRecallPolicy.ASP')
    Company_Username = '//*[@id="content"]/form/table/tbody/tr[1]/td[2]/input'
    Company_Username_Box = driver.find_element_by_xpath(Company_Username)
    Company_Username_Box.send_keys('prestige')
    Staff_Username = '//*[@id="content"]/form/table/tbody/tr[2]/td[2]/input'
    Staff_Username_Box = driver.find_element_by_xpath(Staff_Username)
    Staff_Username_Box.send_keys('your username')
    Password = '//*[@id="content"]/form/table/tbody/tr[3]/td[2]/input'
    Password_Box = driver.find_element_by_xpath(Password)
    Password_Box.send_keys('your password')
    Password_Box.submit()
    policy = '/html/body/div/div[4]/form/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input'
    policy_box = driver.find_element_by_xpath(policy)
    policy_box.send_keys(p.policy_number)
    search = '/html/body/div/div[4]/form/table/tbody/tr[4]/td/input'
    search_btn = driver.find_element_by_xpath(search)
    search_btn.click()
    uw_notes = '/html/body/div/div[4]/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[6]/input'
    uw_notes_btn = driver.find_element_by_xpath(uw_notes)
    uw_notes_btn.click()
    add_notes = '/html/body/div/div[4]/form/table/tbody/tr[4]/td/input[2]'
    add_notes_btn = driver.find_element_by_xpath(add_notes)
    add_notes_btn.click()
    details = '/html/body/div/div[4]/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/textarea'
    details_Box = driver.find_element_by_xpath(details)
    if p.protected == True:
        p.protected = "Protected"
    if p.protected == False:
        p.protected = "NOT Protected"
    note = ('Issued '+str(p.ncb)+' years no claims bonus which was '+str(p.protected))
    details_Box.send_keys(note)
    proceed = '/html/body/div/div[4]/form/table/tbody/tr[4]/td/input[2]'
    proceed_btn = driver.find_element_by_xpath(proceed)
    proceed_btn.click()
    driver.quit()



p.valid = int(0)
        




print('Setting up enviroment')
#as there are like over 300,000 peices of data to go through the code below will load all the data onto the computers RAM so the programe can get all the data 
#it needs without having to keep goings into the database everytime.

policy_table = pd.ExcelFile(r'U:\Harrys Folder\python\autoPolicy3.0\data\policyinfo.xlsx')
print('Loading policy table')
claim_table = pd.ExcelFile(r'U:\Harrys Folder\python\autoPolicy3.0\data\claiminfo.xlsx')
print('loading claims table')
policy_table = pd.read_excel(policy_table, 'Sheet1')
claim_table = pd.read_excel(claim_table, 'Sheet1')




# the next two functions will make the strings presentable to be pushed to the GUI 
def change(big_string):
    variable.set(str(big_string))
    root.update()

claim_dic = {}



#this is where all the formating is done.
def print_table():
    
    format_policy = '{' ':' '<22}'.format(p.policy_number)
    format_name = '{' ':' '<22}'.format(p.name)
    format_number_plate = '{' ':' '<22}'.format(p.number_plate)
    format_ncb_before = '{' ':' '<22}'.format(p.ncb_before)
    format_protected = '{' ':' '<22}'.format(p.protected)
    format_number_claims = '{' ':' '<22}'.format(p.number_claims)




    
    one = ('Policy number '+str(format_policy))
    two = ('Name                '+str(format_name))
    three = ('Number plate   '+str(format_number_plate))
    four = ('protected          '+str(format_protected))
    five = ('NCB before       '+str(format_ncb_before))
    six = ('Claims               '+str(format_number_claims))
    seven = ('Start date           '+str(p.start_date))
    eight =('End date            '+str(p.end_date))
    nine =('NCB now           '+str(p.ncb))
    ten = ('Valid Claims      '+str(p.valid))

    attach_claim = ''
    eight_claim = ''
    for x in range(1, p.number_claims + int(1)):
        one_claim = ('\n')
        two_claim = ('Claim '+str(x))
        three_claim = ('Number            ' +str(claim_dic['claim_number'+str(x)]))
        four_claim = ('Fault                  ' +str(claim_dic['claim_fault'+str(x)]))
        five_claim = ('NCB affected   ' +str(claim_dic['claim_affected'+str(x)]))
        six_claim = ('Status                ' +str(claim_dic['claim_status'+str(x)]))
        seven_claim = ('Date                   ' +str(claim_dic['Date_Of_Loss'+str(x)]))
        
        for key, value in claim_dic.items():
            if ('claim_status'+str(x)) in key and "SETTLED" in value:
                eight_claim = ('Date closed      ' +str(claim_dic['Date_Closed'+str(x)]))
        circ = str(claim_dic['circumstances'+str(x)])

        attach_claim = (attach_claim+str(one_claim)+'\n'+str(two_claim)+'\n'+str(three_claim)+'\n'+(four_claim)+'\n'+str(five_claim)+'\n'+str(six_claim)+'\n'+str(seven_claim)+'\n'+str(eight_claim)+'\n')
        eight_claim = ''
    
    if p.protected == True:
        p.protected = "Protected"
    if p.protected == False:
        p.protected = "NOT Protected"
    note = ('Issued '+str(p.ncb)+' years no claims bonus which was '+str(p.protected))
    
    
     
     
    big_string= (one+'\n'+str(two)+'\n'+str(three)+'\n'+str(four)+'\n'+str(five)+'\n'+str(six)+'\n'+str(seven)+'\n'+str(eight)+'\n'+str(nine)+'\n'+str(ten)+'\n'+str(attach_claim))
    change(big_string)
    update_notes(note)







def loop(x):
    p.policy_number = x
    p.valid = int(0)
    p.number_claims = int(0)

    # this small peice of code tests if the policy number is in the dataframe
    while True:
        if p.policy_number in policy_table.values:
            break
        else:
            return

    # grabs all the data for selected policy 
    policy_number_column = pd.DataFrame(policy_table, columns= ['PolicyNumber'])
    row = policy_number_column.loc[policy_number_column['PolicyNumber']== p.policy_number].index[0]#this finds the row number of the found policy number
    p.ncb = (policy_table.iloc[row]['YearsNCB'])
    p.name = (policy_table.iloc[row]['ProposerName'])
    p.postcode = (policy_table.iloc[row]['Postcode'])
    p.protected = (policy_table.iloc[row]['IsProtectedBonus'])
    p.start_date = (policy_table.iloc[row]['InceptionDate'])
    p.renewal_date = (policy_table.iloc[row]['Inception/RenewalDate'])
    p.number_plate = (policy_table.iloc[row]['VehicleRegistration'])
    p.address_1 = (policy_table.iloc[row]['Address1'])
    p.address_2 = (policy_table.iloc[row]['Address2'])
    p.address_3 = (policy_table.iloc[row]['Address3'])
    p.address_4 = (policy_table.iloc[row]['Address4'])
    p.ncb = int(p.ncb)
    p.renewal_date = str(p.renewal_date)
    p.renewal_date = (p.renewal_date[0:10])
    p.renewal_date = datetime.strptime(p.renewal_date, '%d/%m/%Y')
    p.ncb_before = p.ncb
    

    #gets all claims from each column for selected policy
    number_of_claims = int(0)
    policy_number_column = pd.DataFrame(claim_table, columns= ['Policy Number'])
    row = 1
    key = 1
    index = 0
    while True:
        try:
            search = (policy_number_column.iloc[row]['Policy Number'])
            if p.policy_number in search:
                row_n = policy_number_column.loc[policy_number_column['Policy Number']== p.policy_number].index[index]                
                number = (claim_table.iloc[row_n]['Previous Claim No'])
                claim_dic['claim_number'+str(key)] = number
                fault = (claim_table.iloc[row_n]['At Fault'])
                claim_dic['claim_fault'+str(key)] = fault
                affected = (claim_table.iloc[row_n]['Bonus Affected'])
                claim_dic['claim_affected'+str(key)] = affected
                status = (claim_table.iloc[row_n]['Status'])
                claim_dic['claim_status'+str(key)] = status
                circumstances = (claim_table.iloc[row_n]['Circumstances'])
                claim_dic['circumstances'+str(key)] = circumstances
                loss = (claim_table.iloc[row_n]['Date Of Loss'])
                claim_dic['Date_Of_Loss'+str(key)] = loss
                Date_Closed = (claim_table.iloc[row_n]['Date Closed'])
                claim_dic['Date_Closed'+str(key)] = Date_Closed
                number_of_claims += int(1)
                p.number_claims += int(1)
                row += 1
                key += 1
                index += 1
                
                
            else:
                row += 1
        except IndexError:
            break


    






    # code below converts dates from spreadsheet into dates that python can use.
    

    #the code below goes through the whole dict and trys to format all the dates
    from datetime import date
    today = date.today()
    today = str(today)
    today_day = (today[8:10])
    today_month = (today[5:7])
    today_year = (today[0:4])
    correct_today = (str(today_day) +'/' + str(today_month) +'/' +str(today_year))
    today = datetime.strptime(correct_today, '%d/%m/%Y')

    for x in claim_dic:
        try:
            x = datetime.strptime(claim_dic[x], '%d/%m/%Y')
        except:
            y = x
    p.renewal_date = str(p.renewal_date)
    p.start_date = str(p.start_date)
    p.renewal_date = (p.renewal_date[0:10])
    p.start_date = (p.start_date[0:10])
    try:
        p.start_date = datetime.strptime(p.start_date, '%d/%m/%Y')
    except:
        return



    # the end date is not in the database so i add a year onto the last renewl date
    renewal_day = (p.renewal_date[8:10])
    renewal_month = (p.renewal_date[5:7])
    renewal_year = (p.renewal_date[0:4])
    correct_renewal = (str(renewal_day) +'/' + str(renewal_month) +'/' +str(renewal_year))
    p.renewal_date = datetime.strptime(correct_renewal, '%d/%m/%Y')
    p.end_date = p.renewal_date + timedelta(days=365)


    can = policy_table.loc[policy_table['PolicyNumber'] == p.policy_number]
    can = str(can)
    if "CAN" in can or today > p.end_date:
        p.canceled = True
    else: 
        p.canceled = False



    if p.renewal_date == p.start_date and p.canceled == False:
        p.renewal_date = p.renewal_date + timedelta(days=364)

    #before i can calculate what the NCB will be I need to know if the claim will actually affect the bonuse such as:
    # if a claim is protected and more than 4 years old and settled I dont count that claim

    counts = int(1)
    def validate_step_1(counts):
        for key, value in claim_dic.items():
            if ('claim_fault'+str(counts)) in key and "YES" in value:
                for key, value in claim_dic.items():
                    if ('claim_status'+str(counts)) in key and "ACTIVE" in value:
                        p.valid += int(1)
                    elif ('claim_status'+str(counts)) in key and "SETTLED" in value:
                        for key, value in claim_dic.items():
                            if ('Date_Closed'+str(key)) in key and p.renewal_date < value:
                                p.valid += int(1)
                                for key, value in claim_dic.items():
                                    if ('Date_Of_Loss'+str(key)) in key and value < value - timedelta(days=1460) and p.protected == True:
                                        p.valid -= int(1)
                            
    try:
        for x in range(number_of_claims):
            validate_step_1(counts)
            counts += int(1)
    except:
        y = x



    # the NCB calculator 

    #non protected calulator
    while p.protected == False:
        if p.ncb <= int(2) and p.valid == int(1):
            p.ncb = int(0)
            break
        elif p.ncb == int(3) and p.valid == int(1):
            p.ncb = int(1)
            break
        elif p.ncb >= int(4) and p.valid == int(1):
            p.ncb = int(2)
            break
        elif p.valid >= int(2):
            p.ncb = int(0)
            break
        elif p.canceled == False and p.ncb != int(9) and p.valid == int(0):
            p.ncb += int(1)
            break
        elif p.canceled == False and p.ncb == int(9) and p.valid == int(0):
            break
        else:
            p.ncb += int(1)
            break
    #protected calulator
    while p.protected == True:
        if p.ncb == int(3) and p.valid == int(3):
            p.ncb = int(1)
            p.protected = False
            break
        elif p.ncb >= int(4) and p.valid == int(3):
            p.ncb = int(2)
            break
        elif p.valid >= int(4):
            p.ncb = int(0)
            p.protected = False
            break
        elif p.canceled == False and p.ncb != int(9):
            p.ncb += int(1)
            break
        elif p.canceled == False and p.ncb == int(9):
            break
        else:
            p.ncb += int(1)
            break

    while True:
        if p.protected == False:
            p.protected = 'False'
            break
        else:
            p.protected = 'True'
            break
    
    print_table()


# function below makes up a NCB letter 
def doc(vis):

    #takes a copy of the template doc ready for editing
    doc_name = (p.policy_number.replace("/", "-"))
    src_dir=r'U:\Harrys Folder\python\autoPolicy3.0\templates\template.xlsx'
    dst_dir=r'U:\Harrys Folder\python\autoPolicy3.0\templates\done\\'+str(doc_name)+'.xlsx'
    shutil.copy(src_dir,dst_dir)
    wb = xw.apps.add()
    wb.visible = vis
    wb = xw.Book(r'U:\Harrys Folder\python\autoPolicy3.0\templates\done\\'+str(doc_name)+'.xlsx')
    sht = wb.sheets['Sheet1']


    from datetime import date
    today = date.today()
    today = str(today)
    today_day = (today[8:10])
    today_month = (today[5:7])
    today_year = (today[0:4])
    correct_today = (str(today_day) +'/' + str(today_month) +'/' +str(today_year))
    today = datetime.strptime(correct_today, '%d/%m/%Y')



    # adding all the address together 
    p.address_1 = str(p.address_1)
    p.address_2 = str(p.address_2)
    p.address_3 = str(p.address_3)
    p.address_4 = str(p.address_4)
    if p.address_1 == 'nan':
        p.address_1 = ''
    if p.address_2 == 'nan':
        p.address_2 = '' 
    if p.address_3 == 'nan':
        p.address_3 = '' 
    if p.address_4 == 'nan':
        p.address_4 = '' 


    
    address = (p.address_1+'\n'+str(p.address_2)+'\n'+str(p.address_3)+'\n'+str(p.address_4)+'\n')
    address.title()




    # because the dates in the database are formated realy weirdly i needed to format them so the computer can read it.
    start_date_format = str(p.start_date)
    start_date_format = str(start_date_format[0:10])
    start_day = (start_date_format[8:10])
    start_month = (start_date_format[5:7])
    start_year = (start_date_format[0:4])
    start_date_format = (str(start_day) +'/' + str(start_month) +'/' +str(start_year))

    end_date_format = str(p.end_date)
    end_date_format = str(end_date_format[0:10])
    end_day = (end_date_format[8:10])
    end_month = (end_date_format[5:7])
    end_year = (end_date_format[0:4])
    end_date_format = (str(end_day) +'/' + str(end_month) +'/' +str(end_year))
    
    if p.ncb == int(9) or p.ncb == int(10):
        p.ncb = "9+"

    #below is just adding all the strings to the excel doc
    sht.range('A5').value = correct_today
    sht.range('A6').value = (address+str(p.postcode))
    sht.range('A3').value = p.name
    sht.range('A7').value = 'Policy Number:  ' +str(p.policy_number)
    sht.range('A8').value = 'Inception Date: ' +str(start_date_format)
    sht.range('A9').value = 'Expiry Date:    ' +str(end_date_format)
    sht.range('A10').value = 'Registration:  ' +str(p.number_plate)
    sht.range('A13').value = ('Dear ' +str(p.name))

    # because protected status is presented as True or False This needs to be changed to the following
    if p.protected == True:
        p.protected =str('protected')
    if p.protected == False:
        p.protected =str('NOT protected')

    wording = ('We can confirm that at the time of expiry of the above noted policy you are entitled to '+str(p.ncb)+' Years No Claims Bonus, which is '+str(p.protected))


    #if the policy has not finnished yet i add this to the end of the string
    if today < p.renewal_date:
        wording = ('We can confirm that at the time of expiry of the above noted policy you are entitled to '+str(p.ncb)+' Years No Claims Bonus, which is '+str(p.protected)+' prior to no incidents up untill ('+str(p.renewal_date)+')')

    sht.range('A15').value = wording

    # the below code is needed for grammer in the strings.
    if p.number_claims == int(0):
        claim_wording = 'We can confirm that during the period of cover there have been no claims reported.'
    if p.number_claims == int(1):
        claim_wording = 'We can confirm that during the period of cover there have been 1 claim reported.'
    if p.number_claims >= int(2):
        claim_wording = 'We can confirm that during the period of cover there have been '+str(p.number_claims)+' claims reported.'


    assemble = ''
    counts = int(0) # this long for loop adds all the claims info to a long string
    for x in range(p.number_claims):
        counts += int(1)
        for key, value in claim_dic.items():
            if ('Date_Of_Loss'+str(counts)) in key:
                assemble = (assemble+'('+str(value)+') - ')
                for key, value in claim_dic.items():
                    if ('claim_status'+str(counts)) in key:
                        assemble =(assemble +str(value))
                        for key, value in claim_dic.items():
                            if ('claim_number'+str(counts)) in key:
                                assemble =(assemble+' - '+str(value))
                                for key, value in claim_dic.items():
                                    if ('claim_fault'+str(counts)) in key:
                                        if "YES" in value:
                                            assemble =(assemble +'\nThis will affect your NCB\n')
                                            break
                                        else:
                                            assemble =(assemble +'\nThis will NOT affect your NCB\n')
                                            break




    sht.range('A17').value = claim_wording +'\n'+str(assemble)
    if vis == False:  #when make document is pressed the "vis" varible is set to false so it will save the doc in the background
        wb.save()      # else if the edit document button is pressed it misses this save code so the user can edit the doc
        wb.close()
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        excel = win32com.client.Dispatch("Excel.Application")
        wb_path = r'U:\Harrys Folder\python\autoPolicy3.0\templates\done\\'+str(doc_name)+'.xlsx'
        wb = excel.Workbooks.Open(wb_path)
        ws_index_list = [1] 
        if "-P-" in doc_name:
            path_to_pdf = r'S:\ESIedi\Production\a Maintenance\Portal Bonus Store\\'+str(doc_name)+'.pdf'
        else:
            path_to_pdf = r'S:\ESIedi\Production\a Maintenance\NCB Bonus Stores\\'+str(doc_name)+'.pdf'
        wb.WorkSheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
        excel.Quit()
        time.sleep(2)# need 2 secs for the file to acctualy save otherwise i will get a file error
        os.remove(r'U:\Harrys Folder\python\autoPolicy3.0\templates\done\\'+str(doc_name)+'.xlsx')
        



# this function finds patterns in the emails for policy numbers.
def find_policy_number(message, body):
    found_policys = []
    all_mail = (message+'\n'+str(body))
    regex1 = r"[T][P][C][/]*[P][/]*[0-9]+"
    regex2 = r"[T][C][P][/]*[P][/]*[0-9]+"     
    regex3 = r"[T][P][C][/]*[0-9]+[/]*[0-9]+"
    regex4 = r"[T][C][P][/]*[P][/]*[0-9]+"
    regex5 = r"[T][C][P][0-9]+"
    regex6 = r"[t][p][c][0-9]+"
    regex7 = r"[T][C][V][/]*[0-9]+[/]*[0-9]+"
    regex8 = r"[T][M][B][/]*[P][/]*[0-9]+"
    regex9 = r"[T][C][V][/]*[0-9]+"
    regex10 = r"[T][P][C][/]*[0-9]+"
    regexList = [regex1, regex1, regex3, regex4, regex5, regex6, regex7, regex8, regex9, regex10]

    #this peice of code finds all regex that match and if it has not been found before it will add them
    #to list so they wont be found again.


    # this for loop trys all the combonations of regex on the string
    for x in regexList:
        if re.findall(x, all_mail):
            policy_numbers = re.findall(x, all_mail)     
            for y in policy_numbers:
                if len(y) > 10: #somthimes a regex will grab half a pattern as some policy numbers have tricky patters
                    found_policys.append(y)
    found_policys= list(dict.fromkeys(found_policys)) #because two different regex could find the same policy number i need to get rid of the duplicats
    for x in found_policys:
        loop(x)






#when the user clicks show PDF this will open the NCB from the folder


def open_pdf():
    doc_name = (p.policy_number.replace("/", "-")) #As the file is saved "/" needs to be replaced as you can't save files with this "/" 

    try:
        if "-P-" in doc_name:
            os.startfile(r'S:\ESIedi\Production\a Maintenance\Portal Bonus Store\\'+str(doc_name)+'.pdf')  
        else:    #as there are two different folders for files. files with -p- in them will be save in a different folder.
            os.startfile(r'S:\ESIedi\Production\a Maintenance\NCB Bonus Stores\\'+str(doc_name)+'.pdf')
    except FileNotFoundError:
        print('Error NCB not found.')

def open_folder():
    doc_name = (p.policy_number.replace("/", "-"))
    if "-P-" in doc_name:
        subprocess.Popen(r'explorer S:\ESIedi\Production\a Maintenance\Portal Bonus Store')
    else:
        subprocess.Popen(r'explorer S:\ESIedi\Production\a Maintenance\NCB Bonus Stores')


def visable():
    vis = False
    doc(vis)

def not_visable():
    vis = True
    doc(vis)


p.lock = False
def lock_func():
    while True:
        if p.lock == True:
            p.lock = False
            break
        elif p.lock == False:
            p.lock = True
            break
        else:
            break
    

    


#everything below this code controls the GUI

root=Tk()
root.geometry("400x700")
root.title('Email Policy Companion')
root.wm_iconbitmap(r'U:\Harrys Folder\python\favicon.ico')


variable=StringVar()
your_label=Label(root,textvariable=variable, justify='left')
your_label.pack()




b = Button(root, text="Make NCB Document", command=visable, justify='left')
b.pack()

your_label2=Label(root,text="\n", justify='left')
your_label2.pack()

b8 = Button(root, text="Edit NCB Document", command= not_visable, justify='right')
b8.pack()


your_label2=Label(root,text="\n", justify='left')
your_label2.pack()



b2 = Button(root, text="Show NCB", command=open_pdf, justify='right')
b2.pack()

your_label5=Label(root,text="\n", justify='left')
your_label5.pack()

b3 = Button(root, text="Open NCB folder", command=open_folder, justify='right')
b3.pack()




variable3=StringVar()
your_label3=Label(root,textvariable=variable3, justify='left')
your_label3.pack()


b = Button(root, text="Add to underwriting notes", command=underwriting_notes)
b.pack()



lockb = Button(root, text="Lock Infomation", command=lock_func, justify='right')
lockb.pack()





p.old_message = ''

def gui():
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.ActiveExplorer().Selection.Item(1)
        body = message.body
        message = str(message)
        if p.old_message != message and p.lock == False:                    #evey 0.5 seconds this code tests if the user has selected a different email
            p.old_message = message
            find_policy_number(message, body)
        else:
            time.sleep(0.5)
    except: 
        time.sleep(0.5)
    
    root.after(1000, gui)




root.after(0, gui)

root.mainloop()












