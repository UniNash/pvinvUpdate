#--------------------------Time to check runtime------------------

import time
start = time.time()

#---------------------------Importing libraries--------------

import sys
import urllib.request as urllib
import bs4
import dateutil
import os
import pandas as pd
from datetime import datetime as dt
import warnings
import win32com.client as win32
import smtplib


#-----------------------Handle excel validation warning-----------

warnings.filterwarnings("ignore") # warnings for excel validation


#-----------------------------Defining Functions---------------

#Function to extract the first link from the goSolar URL
#------------------------------------------------------

# url - is the main GoSolar web page
# searchlink_name is the name of the tag text to search for from where the subpage url (where the download link is) is extracted"

def extract_download_link_date(url, searchlink_name):
    request_url = urllib.urlopen(url)
    base_site = bs4.BeautifulSoup(request_url)
    for item in base_site.find_all(lambda tag:tag.name=='a' and searchlink_name in tag.text):
        extractlink=item.get('href')
        scrapedate=str(item.next_sibling)   # extracting the date text; which will be used to check whether we are downloading the latest file
        dateconvert=dateutil.parser.parse(scrapedate,fuzzy=True) # converting the extracted date to datetime 
        extractedset = [extractlink,dateconvert] # creating a list of the subpage url and date
    return extractedset

# Function to extract the PV/INV file download link
#---------------------------------------------------

# "linkset" - is the output from function "extract_download_link_date" which is a list.This list is extended with filename and a boolean check for log file.

def extract_main_download_link_date(linkset): 
    url=linkset[0]
    searchlink_name = linkset[2]
    request_url = urllib.urlopen(url)
    base_site = bs4.BeautifulSoup(request_url)
    for item in base_site.find_all(lambda tag:tag.name=='a' and searchlink_name in tag.text):
        extractlink=item.get('href')
    return extractlink


#Function to send notification emails
#-----------------------------------------
def email_notification(to,cc,subj,msg1,msg2,msg3,msg4,attach1,attach2,attach3):

    outlook = win32.Dispatch('outlook.application') # starts the local outlook application using python;
                                                    # outlook should be logged in with offical email credentials
    mail = outlook.CreateItem(0)                    # creates an email object
    mail.To = to
    mail.CC = cc
    mail.Subject = subj
    mail.HTMLBody = """\
    <html>
      <head></head>
       <body>
        <br>{0}
        <br>
        <br>{1}
        <br>{2}
        <br>{3}
      </body>
    </html>
    """.format(msg1,msg2,msg3,msg4)

    # To attach a file to the email (optional):

    # attachment  = "Path to the attachment"
    
    if(attach3!=''):
        mail.Attachments.Add(attach3)               # attachment paths - to be mentioned in the arguments.
    if(attach2!=''):
        mail.Attachments.Add(attach2)
    if(attach1!=''):
        mail.Attachments.Add(attach1)

    mail.Send()


#Function to create and update log files
#----------------------------------------

# "linkset" - is the output from function "extract_download_link_date" which is a list.This list is extended with filename and a boolean check for log file.
# "path" - path of the log file
# "log_header" - headers required for the log file


# using this function to check for log file in the given path (eg."PV_log_file"), if not then creates a log file and return the log dataframe.

def log (linkset,path,log_header):   
    if(linkset[3]==True):                       # log file exists check (eg."if_exists_PV_log")
        goSolar_updates_log=pd.read_csv(path,index_col='Index_ID') 
        return goSolar_updates_log
    else:
        print("Log files not available")
        print("Creating log files......")
        goSolar_updates_log = pd.DataFrame(columns=log_header)
        goSolar_updates_log.to_csv(path,index=True,index_label='Index_ID')
        goSolar_updates_log=pd.read_csv(path,index_col='Index_ID')
        print("Created log files")
        return goSolar_updates_log

#Function to check if we are downloading the latest file with the log and then extract the main download link
#-----------------------------------------------------------------------------------------------

# "linkset" - is the output from function "extract_download_link_date" which is a list.This list is extended with filename and a boolean check for log file.
# log_file - is the log dataframe

def check_dates_download_main (linkset,log_file):
    today=dt.date(dt.now())
    day = int((today-dt.date(linkset[1])).days) # no. of days from today and the date from the subpage url ( to check if the latest update is within 2 weeks )
    log_file['date_sort']=pd.to_datetime(log_file['Update_Date'])
    log_file.sort_values('date_sort', inplace=True, ascending=True)
    log_file.drop('date_sort', inplace=True, axis=1)
    check_date = log_file.tail(1)['Update_Date'].values[0]
    check_date = dt.date(dt.strptime(check_date,'%m/%d/%Y')) #getting the latest updated data from logfile.
    if(day<13):
        print("Latest File. {} days old".format(day))
    else:
        print("Old File. {} days old".format(day)) # if it has been more than 2 weeks then the latest data was not update
    if (len(log_file)==0):
        print("Extracting link......")
        download_link = extract_main_download_link_date(linkset) "using the function to extract file download link)
        print("Extracted")
        return download_link
    elif (len(log_file)>0):
        if (check_date==dt.date(linkset[1])):  # if the last updated log data and the linkset date are same then the data has already been updated so exiting from the program
            tell_1="Abort"
            print("Abort")
            return tell_1
        else:
            print("Extracting link......")
            download_link = extract_main_download_link_date(linkset)
            print("Extracted")
            return download_link



#Fucntion to check and download the file from main download link
-------------------------------------------------------------------------

def download_file (linkset,downloadlink,path):  
    dir_path = os.path.join(path, dt.strftime(dt.date(linkset[1]),'%m.%d.%Y')) 
    is_dir_available = os.path.isdir(dir_path) # download folder path
    today=dt.date(dt.now())
    if(downloadlink =="Abort"):  # if the data is upto data
        if(is_dir_available==True):
            sub_1 = "goSolar"+"PV/INV Update - " + dt.strftime(today,'%m/%d/%Y') 
            msg_1 = "<i>**This is an automated email**</i>"+"<br><br>Hello,"+"<br><br>Logs suggest that "+'<b><p style="color:red">' + linkset[2] + "</p></b>"+" data dated " +dt.strftime(dt.date(linkset[1]),'%m/%d/%Y') + " were already updated. Please check the logs. <br><i><b>Respective data folder is also available in local.</b></i>"+"<br><br><br>"+"Regards,<br>Naresh.R"
            to_cont_1='r.naresh@anbsystems.com'
            cc_cont_1= 'shenbagaveni@anbsystems.com ; Venkatesh.R@anbsystems.com'
            #cc_cont_1= 'r.naresh@anbsystems.com ; r.naresh@anbsystems.com'
            email_notification(to_cont_1,cc_cont_1,sub_1,msg_1,"","","","","","")
            tell_2="Abort_1"
            print("Abort_1. Latest file updated as per log and folder is also available")
            return tell_2
        else:  # if download folder is not available
            sub_1 = "goSolar"+"PV/INV Update - " + dt.strftime(today,'%m/%d/%Y') 
            msg_1 = "<i>**This is an automated email**</i>"+"<br><br>Hello,"+"<br><br>Logs suggest that "+'<b><p style="color:red">' + linkset[2] + "</p></b>"+" data dated " +dt.strftime(dt.date(linkset[1]),'%m/%d/%Y') + " were already updated. Please check the logs. <br><i><b>Looks like the reespective data folder is not available in local.</b></i>"+"<br><br><br>"+"Regards,<br>Naresh.R"
            to_cont_1='r.naresh@anbsystems.com'
            cc_cont_1= 'shenbagaveni@anbsystems.com ; Venkatesh.R@anbsystems.com'
            #cc_cont_1= 'r.naresh@anbsystems.com ; r.naresh@anbsystems.com'
            email_notification(to_cont_1,cc_cont_1,sub_1,msg_1,"","","","","","")
            tell_2="Abort_2"
            print("Abort_2. Latest file updated as per log, but folder not found")
            return tell_2
    else:
        if(is_dir_available==True):
            sub_1 = "goSolar"+"PV/INV Update - " + dt.strftime(today,'%m/%d/%Y') 
            msg_1 = "<i>**This is an automated email**</i>"+"<br><br>Hello,"+"<br><br>"+'<b><p style="color:red">' + linkset[2] + "</p></b>"+" data folder dated " +dt.strftime(dt.date(linkset[1]),'%m/%d/%Y') + " is already available in local. Please check."+"<br><br><br>"+"Regards,<br>Naresh.R"
            to_cont_1='r.naresh@anbsystems.com'
            cc_cont_1= 'shenbagaveni@anbsystems.com ; Venkatesh.R@anbsystems.com'
            #cc_cont_1= 'r.naresh@anbsystems.com ; r.naresh@anbsystems.com'
            email_notification(to_cont_1,cc_cont_1,sub_1,msg_1,"","","","","","")
            tell_2="Abort_3"
            print("Abort_3.Folder already available for {0}".format(linkset[2])) # aborting file download if the folder already available for the given date
            return tell_2
        else:  # downloading the actual file into the folder.
            print("creating folder.........")
            os.mkdir(dir_path)
            print("Folder created {0}".format(dir_path))
            file_path=os.path.join(dir_path,linkset[2])
            print("Downloading file.........")
            urllib.urlretrieve(downloadlink,file_path)
            print("File downloaded {0}".format(file_path))
            pathreturn=[dir_path,file_path]
            return pathreturn



#---------------Initializing Constants---------------



url="https://www.energy.ca.gov/programs-and-topics/topics/renewable-energy/solar-equipment-lists"
today=dt.date(dt.now())

#Initialize constants for PV
PV = "PV Module List - Full Data"
PV_Filename = 'PV_Module_List_Full_Data_ADA.xlsx'
PV_log_file = r'C:\Users\r.naresh\OneDrive - ANB Systems Private Limited\General Tasks\goSolar Update\log\goSolar_PV_update_log.csv'
goSolar_PV_log_headers=['Serial_No'
                        ,'Update_Date'
                        ,'Date_Downloaded'
                        ,'PV_Records'
                        ,'New_PV_Records'
                        ,'Created_Date']
PV_main_path = r'C:\Users\r.naresh\OneDrive - ANB Systems Private Limited\General Tasks\goSolar Update\Equipment_Updates\PV'



#Initialize constants for INV
INV = "Grid Support Inverter List - Full Data"
INV_Filename = 'Grid_Support_Inverter_List_Full_Data_ADA.xlsm'
INV_log_file = r'C:\Users\r.naresh\OneDrive - ANB Systems Private Limited\General Tasks\goSolar Update\log\goSolar_INV_update_log.csv'
goSolar_INV_log_headers=['Serial_No'
                         ,'Update_Date'
                         ,'Date_Downloaded'
                         ,'Utility_INV_Records'
                         ,'Solor_INV_Records'
                         ,'Battery_INV_Records'
                         ,'Grid_INV_Records'
                         ,'INV_Records'
                         ,'New_INV_Records'
                         ,'Created_Date']
INV_main_path = r'C:\Users\r.naresh\OneDrive - ANB Systems Private Limited\General Tasks\goSolar Update\Equipment_Updates\INV'


#----------------Validating and running the functions-----------------


#PV
if_exists_PV_log = os.path.isfile(PV_log_file)

PV_Set = extract_download_link_date(url,PV)

PV_Set.extend([PV_Filename,if_exists_PV_log])

goSolar_PV_update_log=log(PV_Set,PV_log_file,goSolar_PV_log_headers)

PV_download_link=check_dates_download_main(PV_Set,goSolar_PV_update_log)


#INV
if_exists_INV_log = os.path.isfile(INV_log_file)

INV_Set = extract_download_link_date(url,INV)

INV_Set.extend([INV_Filename,if_exists_INV_log])

goSolar_INV_update_log=log(INV_Set,INV_log_file,goSolar_INV_log_headers)

INV_download_link=check_dates_download_main(INV_Set,goSolar_INV_update_log)


#PV main link ; downloading the actual file
PV_path_set = download_file(PV_Set,PV_download_link,PV_main_path)

#INV main link ; downloading the actual file
INV_path_set = download_file(INV_Set,INV_download_link,INV_main_path)


#PV and INV validate ; exiting the system if the downloading file functions return abort message.
if(PV_path_set == "Abort_1" or PV_path_set == "Abort_2" or  PV_path_set == "Abort_3"):
    if(INV_path_set == "Abort_1" or INV_path_set == "Abort_2" or  INV_path_set == "Abort_3"):
        end = time.time()
        total_execution_time = (end - start)/60
        print("Total time taken {} mins.".format(total_execution_time))
        sys.exit(0)


#----------------------Process PV files------------------------


if(PV_path_set != "Abort_1" and PV_path_set != "Abort_2" and  PV_path_set != "Abort_3"): # making sure if no abort messages are produced

    #processing PV files
    goSolarPV = pd.read_excel (PV_path_set[1], engine ='openpyxl') # read the downloaded excel file using openpyxl engine
    idxpv = int(goSolarPV[goSolarPV['PV Module List (Full Data)']=='Manufacturer'].index.values) # finding the header column index as there are few columns above with some text.

    newcol = goSolarPV.iloc[idxpv,:] 

    goSolarPV = goSolarPV.set_axis(newcol, axis=1, inplace=False) #setting the first row as headers
    
    del_rows_pv = []   #getting the list of rows to delete
    for i in range(idxpv+2):
        del_rows_pv.append(i)
    del_rows_pv
    
    goSolarPV = goSolarPV.drop(del_rows_pv) # deleting all unrequired rows
    goSolarPV = pd.DataFrame(goSolarPV,columns=['Manufacturer'
                                                ,'Model Number'
                                                ,'Description'
                                                ,'BIPV'
                                                ,'Nameplate Pmax'
                                                ,'PTC'])
    trim_strings = lambda tr1: tr1.strip() if isinstance(tr1, str) else tr1 #custom function stripping white spaces from all the columns ; trim if string else return the value
    
    goSolarPV = goSolarPV.applymap(trim_strings) #applying the custom function
    
    newcol_PV = ['manufacturer','modelNumber','description','bipv','watts','ptc']
    
    goSolarPV = goSolarPV.set_axis(newcol_PV, axis=1, inplace=False)
    
    goSolarPV.to_csv(os.path.join(PV_path_set[0],'goSolarPV.csv'), sep=",",index=False)
    
    goSolarPV = goSolarPV[['manufacturer','modelNumber','description','bipv','watts']]
    
    goSolarPV.to_excel(os.path.join(PV_path_set[0],'goSolarPV.xlsx'), sheet_name='goSolarPV',index=False) # for MASSCEC project
    goSolarPV.to_csv(os.path.join(PV_path_set[0],'solar.csv'), sep="|",index=False) # for RMLD project

    #Save goSolarPV as .xls form MASSCEC project
    --------------------------------------------

    #excel = win32.gencache.EnsureDispatch('Excel.Application')   #using the system's excel application to process the file and save it as .xls format
    excel = win32.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(os.path.join(PV_path_set[0],'goSolarPV.xlsx'))
    wb.SaveAs(os.path.join(PV_path_set[0],'goSolarPV.xls'), FileFormat=56)
    wb.Close(True)
    #excel.Application.Quit()
    
    
    #------log update and Checking file-----
    
    #Checking file
    
    goSolar_PV_update_log['pv_date_sort']= pd.to_datetime(goSolar_PV_update_log['Update_Date'])
    goSolar_PV_update_log = goSolar_PV_update_log.sort_values('pv_date_sort',ascending=True)
    goSolar_PV_update_log.drop('pv_date_sort', inplace=True, axis=1)
    
    if (len(goSolar_PV_update_log)>0): # if the update file has records then 
        prev_date_pv = goSolar_PV_update_log.tail(1)['Update_Date'].values[0]
        prev_date_pv = dt.strptime (prev_date_pv,'%m/%d/%Y')
        prev_dir_path_pv = os.path.join(PV_main_path, dt.strftime(dt.date(prev_date_pv),'%m.%d.%Y')) 
        prev_file_path_pv = os.path.join(prev_dir_path_pv,'goSolarPV.xlsx') 
        goSolarPV_prev = pd.read_excel (prev_file_path_pv, engine ='openpyxl')

        # new records into excel file
        goSolarPV_check = goSolarPV[['manufacturer','modelNumber']].append(goSolarPV_prev[['manufacturer','modelNumber']],ignore_index=True)
        goSolarPV_check = goSolarPV_check.drop_duplicates(keep=False)
    
        goSolarPV_check.to_excel(os.path.join(PV_path_set[0],'goSolarPV_check.xlsx'), sheet_name='goSolarPV_check',index=False)
        
        new_pv_rec = len(goSolarPV_check) # no. of new records

    else:
        new_pv_rec = len(goSolarPV)- (0 if goSolar_PV_update_log.tail(1)['PV_Records'].empty == True else goSolar_PV_update_log.tail(1)['PV_Records'].values[0])

        
    #Log update
      
    PV_new_updates = [[len(goSolar_PV_update_log)+1
                      ,dt.strftime(dt.date(PV_Set[1]),'%m/%d/%Y')
                      ,dt.strftime(today,'%m/%d/%Y')
                      ,len(goSolarPV)       # total records
                      ,new_pv_rec           # new records
                      ,dt.strftime(today,'%m/%d/%Y')]]
    goSolar_PV_new_updates = pd.DataFrame(PV_new_updates,columns=goSolar_PV_log_headers)
    goSolar_PV_update_log = goSolar_PV_update_log.append(goSolar_PV_new_updates,ignore_index=True)
    goSolar_PV_update_log.to_csv(PV_log_file,index=True,index_label='Index_ID',mode='w')
    finalPV = "PV processed"
    print("PV processed")

else:
    finalPV = "Did not process PV"
    print("Did not process PV")


#---------------------Process INV files----------------------------


if(INV_path_set != "Abort_1" and INV_path_set != "Abort_2" and  INV_path_set != "Abort_3"):
    goSolarINV_Solar = pd.read_excel (INV_path_set[1], sheet_name = 'Solar_Inverters', engine ='openpyxl') #solar sheet
    goSolarINV_Bat = pd.read_excel (INV_path_set[1], sheet_name = 'Battery_Inverters', engine ='openpyxl') #battery sheet

    idxsinv=int(goSolarINV_Solar[goSolarINV_Solar['Grid Support Solar Inverter List (Full Data)']=='Manufacturer Name'].index.values)
    idxbinv=int(goSolarINV_Bat[goSolarINV_Bat['Grid Support Battery Inverter List (Full Data)']=='Manufacturer Name'].index.values)

    newcol1 = goSolarINV_Solar.iloc[idxsinv,:]
    newcol2 = goSolarINV_Bat.iloc[idxbinv,:]

    goSolarINV_Solar = goSolarINV_Solar.set_axis(newcol1, axis=1, inplace=False)
    goSolarINV_Bat = goSolarINV_Bat.set_axis(newcol2, axis=1, inplace=False)


    del_rows_sINV = [] # rows to delete
    for i in range(idxsinv+2):
        del_rows_sINV.append(i)

    del_rows_bINV = []
    for i in range(idxbinv+2):
        del_rows_bINV.append(i)

    goSolarINV_Solar=goSolarINV_Solar.drop(del_rows_sINV)
    goSolarINV_Bat=goSolarINV_Bat.drop(del_rows_bINV)

    goSolarINV_Solar = goSolarINV_Solar.loc[:,~goSolarINV_Solar.columns.duplicated()]
    goSolarINV_Bat = goSolarINV_Bat.loc[:,~goSolarINV_Bat.columns.duplicated()]
    
    goSolarINV_Solar = pd.DataFrame(goSolarINV_Solar,columns=['Manufacturer Name','Model Number1','Description','Maximum Continuous Output Power at Unity Power Factor','Nominal Voltage','Weighted Efficiency','Built-In Meter'])
    goSolarINV_Bat = pd.DataFrame(goSolarINV_Bat,columns=['Manufacturer Name','Model Number1','Description','Maximum Continuous Output Power at Unity Power Factor','Nominal Voltage','Weighted Efficiency','Built-In Meter'])

    trim_strings = lambda tr2: tr2.strip() if isinstance(tr2, str) else tr2
    goSolarINV_Solar = goSolarINV_Solar.applymap(trim_strings)

    trim_strings = lambda tr3: tr3.strip() if isinstance(tr3, str) else tr3
    goSolarINV_Bat = goSolarINV_Bat.applymap(trim_strings)
    
    newcol_INV=['manufacturer','modelNumber','description','powerRating','nominalVoltage','weightedEfficiency','builtInMeter']

    goSolarINV_Solar = goSolarINV_Solar.set_axis(newcol_INV, axis=1, inplace=False)
    goSolarINV_Bat = goSolarINV_Bat.set_axis(newcol_INV, axis=1, inplace=False)

    goSolarINV_Grid = goSolarINV_Solar.append(goSolarINV_Bat)

    goSolarINV_Grid['description']=goSolarINV_Grid['description'].str[0:240]
    goSolarINV_Grid['watts']=goSolarINV_Grid['powerRating']*1000
    
    if_exists_INV_Utility_file = os.path.isfile(os.path.join(INV_main_path, 'goSolarINV_Utility.xlsx'))
    
    if (if_exists_INV_Utility_file == True):
        goSolarINV_Utility = pd.read_excel (os.path.join(INV_main_path, 'goSolarINV_Utility.xlsx'), engine ='openpyxl')
    else:
        to_cont_2='r.naresh@anbsystems.com'
        cc_cont_2= 'shenbagaveni@anbsystems.com ; Venkatesh.R@anbsystems.com'
        #cc_cont_2= 'r.naresh@anbsystems.com ; r.naresh@anbsystems.com'
        sub_2 = "goSolar"+"PV/INV Update - " + dt.strftime(today,'%m/%d/%Y')
        msg_2 = "<i>**This is an automated email**</i>"+"<br><br>Hello,"+"<br><br> Could not find goSolarINV_Utility.xlsx. Please place the file and re-run."
        email_notification(to_cont_2,cc_cont_2,sub_2,msg_2,"","","","","","")
        end = time.time()
        total_execution_time = (end - start)/60
        print("Total time taken {} mins.".format(total_execution_time))
        sys.exit(0)
        
    goSolarINV=goSolarINV_Utility.append(goSolarINV_Grid)

    #Save as .xlsx

    goSolarINV.to_excel(os.path.join(INV_path_set[0],'goSolarINV.xlsx'), sheet_name='solar inverter',index=False) #for MASSCEC project
    
    #Save as .csv
    goSolarINV.to_csv(os.path.join(INV_path_set[0],'goSolarINV.csv'),index=False) #for RMLD project
    
    
    #Save goSolarINV as .xls for MASSCEC project

    #excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel = win32.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(os.path.join(INV_path_set[0],'goSolarINV.xlsx'))
    wb.SaveAs(os.path.join(INV_path_set[0],'goSolarINV.xls'), FileFormat=56)
    wb.Close(True)
    #excel.Application.Quit()
    
    #---log update and check file----
    
    goSolar_INV_update_log['inv_date_sort']= pd.to_datetime(goSolar_INV_update_log['Update_Date'])
    goSolar_INV_update_log=goSolar_INV_update_log.sort_values('inv_date_sort',ascending=True)
    goSolar_INV_update_log.drop('inv_date_sort', inplace=True, axis=1)
    
    if (len(goSolar_INV_update_log)>0):
        prev_date_inv = goSolar_INV_update_log.tail(1)['Update_Date'].values[0]
        prev_date_inv = dt.strptime (prev_date_inv,'%m/%d/%Y')
        prev_dir_path_inv = os.path.join(INV_main_path, dt.strftime(dt.date(prev_date_inv),'%m.%d.%Y')) 
        prev_file_path_inv = os.path.join(prev_dir_path_inv,'goSolarINV.xlsx') 
        goSolarINV_prev = pd.read_excel (prev_file_path_inv, engine ='openpyxl')
    
        goSolarINV_check = goSolarINV[['manufacturer','modelNumber']].append(goSolarINV_prev[['manufacturer','modelNumber']],ignore_index=True)
        goSolarINV_check = goSolarINV_check.drop_duplicates(keep=False)
        goSolarINV_check.to_excel(os.path.join(INV_path_set[0],'goSolarINV_check.xlsx'), sheet_name='goSolarINV_check',index=False)
        
      
        new_inv_rec = len(goSolarINV_check)

    else:
        new_inv_rec = len(goSolarINV) - (0 if goSolar_INV_update_log.tail(1)['INV_Records'].empty == True else goSolar_INV_update_log.tail(1)['INV_Records'].values[0])        
    
    
    INV_new_updates = [[len(goSolar_INV_update_log)+1
                      ,dt.strftime(dt.date(INV_Set[1]),'%m/%d/%Y')
                      ,dt.strftime(today,'%m/%d/%Y')
                      ,len(goSolarINV_Utility)
                      ,len(goSolarINV_Solar)
                      ,len(goSolarINV_Bat)
                      ,len(goSolarINV_Grid)
                      ,len(goSolarINV_Utility)+len(goSolarINV_Grid)
                      ,new_inv_rec
                      ,dt.strftime(today,'%m/%d/%Y')]]
    goSolar_INV_new_updates = pd.DataFrame(INV_new_updates,columns=goSolar_INV_log_headers)
    goSolar_INV_update_log = goSolar_INV_update_log.append(goSolar_INV_new_updates,ignore_index=True)
    goSolar_INV_update_log.to_csv(INV_log_file,index=True,index_label='Index_ID',mode='w')
    finalINV = "INV processed"
    print("INV processed")
else:
    finalINV = "Did not process INV"
    print("Did not process INV")


#------------send email to eTP BAs-----------------------

if(finalPV == "PV processed"):
    msg_etp_a= "<b><u>PV-updates dated - </u></b>"+"<b><u>"+dt.strftime(dt.date(PV_Set[1]),'%m/%d/%Y')+"</u></b>"+"<br>"+goSolar_PV_new_updates.to_html()+"<br><br>"
    attachmentPV = os.path.join(PV_path_set[0],'goSolarPV.csv')
    
else:
    msg_etp_a = "No latest PV updates available"+"<br><br>"
    attachmentPV =""
        
if(finalINV == "INV processed"):
    msg_etp_b = "<b><u>INV-updates dated - </u></b>"+"<b><u>"+dt.strftime(dt.date(INV_Set[1]),'%m/%d/%Y')+"</u></b>"+"<br>"+goSolar_INV_new_updates.to_html()+"<br><br>"
    attachmentINV = os.path.join(INV_path_set[0],'goSolarINV.csv')
else:
    msg_etp_b = "No latest INV updates available"+"<br><br>"
    attachmentINV = ""
    
sub_main_etp = "goSolar"+"PV/INV Update - " + dt.strftime(today,'%m/%d/%Y')
msg_main_etp= "<i>**This is an automated email**</i>"+"<br><br>Hello,"+"<br><br>Please find attached the Go Solar data as per the latest update received."+"<br><br>"+msg_etp_a+msg_etp_b+"<br><br><br>"+"Regards,<br>Naresh.R"
#to_cont_etp = 'myself@anbsystems.com'
to_cont_etp = 'BA1@anbsystems.com ; BA2@anbsystems.com'
cc_cont_etp = 'myself@anbsystems.com'
email_notification(to_cont_etp,cc_cont_etp,sub_main_etp,msg_main_etp,"","","",attachmentPV,attachmentINV,"")

#--------------------Transfer files to SFTP------------------------

import paramiko # library to connect to sftp

#create a function to connect to SFTP

def sftpconnect (address,port,username,password):
    try:
        transport = paramiko.Transport((address, port))
        transport.connect(username=username,password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        print("Connected to SFTP")
        connections= [sftp, transport]
        return connections
    except:
        print("Connection failed. Check SFTP login credentials")
        to_cont_3='myself@anbsystems.com'
        cc_cont_3= 'BA1@anbsystems.com ; BA2@anbsystems.com'
        #cc_cont_3= 'myself@anbsystems.com; myself@anbsystems.com'
        subj_3 = "goSolar Update"
        msg_3 = "Some exception error has occured or the connection to SFTP failed, check login credentials"
        email_notification(to_cont_3,cc_cont_3,subj_3,msg_3,"","","","","","")
        sys.exit


connect = sftpconnect("FTP ADDRESS", 22,"USERNAME","PASSWORD")

sftp = connect[0]
transport = connect[1]

sftp.chdir('/MASSCEC/goSolar/')

sftp.mkdir(dt.strftime(today,'%m.%d.%Y'))


if(finalPV == "PV processed"):
    localpath1 = os.path.join(PV_path_set[0],'goSolarPV.xls')
    localpath2 = os.path.join(PV_path_set[0],'solar.csv')           
    sfpath1 = '/MASSCEC/goSolar/'+dt.strftime(today,'%m.%d.%Y')+'/goSolarPV.xls'
    sfpath2 = '/MASSCEC/goSolar/'+dt.strftime(today,'%m.%d.%Y')+'/solar.csv'
    sftp.put(localpath1,sfpath1)
    print ("PV file 1 uploaded")
    sftp.put(localpath2,sfpath2)
    print ("PV file 2 uploaded")
else:
    print("Did not process PV")
    
if(finalINV == "INV processed"):
    localpath3 = os.path.join(INV_path_set[0],'goSolarINV.xls')
    sfpath3 = '/MASSCEC/goSolar/'+dt.strftime(today,'%m.%d.%Y')+'/goSolarINV.xls'
    sftp.put(localpath3,sfpath3)
    print ("INV file uploaded")
else:
    print("Did not process INV")


sftp.close()
transport.close()

print("Export completed")


#----------------Send notification and exit---------------------

if(finalINV == "INV processed"):
    msg_a = "<b><u>INV-updates dated - </u></b>"+"<b><u>"+dt.strftime(dt.date(INV_Set[1]),'%m/%d/%Y')+"</u></b>"+"<ol><li>MassCEC – Inverter</li></ol>"+"<br>"+goSolar_INV_new_updates.to_html()+"<br><br>"
else:
    msg_a = "No latest INV updates available"+"<br><br>"

if(finalPV == "PV processed"):
    msg_b= "<b><u>PV-updates dated - </u></b>"+"<b><u>"+dt.strftime(dt.date(PV_Set[1]),'%m/%d/%Y')+"</u></b>"+"<ol> <li>MassCEC – PV</li> <li>RMLD – PV</li></ol>"+"<br>"+goSolar_PV_new_updates.to_html()+"<br><br>"
else:
    msg_b = "No latest PV updates available"+"<br><br>"


sub_main = "goSolar"+"PV/INV Update - " + dt.strftime(today,'%m/%d/%Y')
msg_main= "<i>**This is an automated email**</i>"+"<br><br>Hello,"+"<br><br>Please update the Go Solar data as per the latest update received."+"<br>The data files are available in SFTP in the folder <b><u>" + dt.strftime(today,'%m.%d.%Y')+"</u></b>."+"<br><br>"+msg_b+msg_a+"<br><br><br>"+"Regards,<br>Naresh.R"
to_cont = 'SUPPORTGROUPRELEASE@anbsystems.com'
#to_cont = 'myself@anbsystems.com'
cc_cont = 'MANAGER1@anbsystems.com ; MANAGER2@anbsystems.com ; DEV1@anbsystems.com ; BA1@anbsystems.com ; BA2@anbsystems.com'
#cc_cont = 'myself@anbsystems.com ; myself@anbsystems.com; myself@anbsystems.com ; myself@anbsystems.com ; myself@anbsystems.com'

email_notification(to_cont,cc_cont,sub_main,msg_main,"","","","","","")

end = time.time()
total_execution_time = (end - start)/60
print("Process Completed. Total time taken {} mins.".format(total_execution_time))
sys.exit
