from selenium import webdriver #Web driver activities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC #Error handling
from selenium.webdriver.support.ui import WebDriverWait #Web driver wait
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from configparser import ConfigParser #Configuration read
from pathlib import Path #Path conversions
from os import listdir
from os.path import isfile, join
import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
# import undetected_chromedriver as uc
import os #Path
import sys #System paths
import re
import csv
import ctypes #message box
import pyodbc
import logging #Activity logging
import time #Sleep
import pyautogui
import time
import pygetwindow
import shutil
import keyboard   #from webdriver_manager.chrome import ChromeDriverManager
 
 
download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'
firefox_location='./'
 
 
global wait, driver, webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection,min
 
 
def read_config_file():
    global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
    global sql_server_name, sql_user_name, sql_password, sql_db
    try:
        # configuration entries
        config = ConfigParser()
        config.read(ConfigPath)
        webpage_url = config.get ("Data", "webpage_URL")
        user_name = config.get ("Data", "user_name")
        password = config.get ("Data", "password")
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password")
        sql_db = config.get ("SQL", "database")
        return(0)
    except Exception as e:
        print('config file read error')
        return(-1)
 
#--------------------------------------------------------------------------------------------------------
 
 
def setup():
    global firefox_location, download_dir, webpage_url, driver, wait
    binary = FirefoxBinary(firefox_location)
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/zip")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "image/jpeg")
    profile.set_preference("browser.helperApps.saveToDisk.image/jpeg", "application/octet-stream")
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference("browser.download.viewableInternally.enabledTypes", "")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("browser.download.manager.useWindow", False)
    profile.set_preference("browser.download.manager.closeWhenDone", True)
    profile.set_preference("print_printer", "Microsoft Print to PDF")
    profile.set_preference("print.always_print_silent", True)
    profile.set_preference("print.show_print_progress", False)
    profile.set_preference("print.save_as_pdf.links.enabled", True)
    profile.set_preference("browser.helperApps.alwaysAsk.force", False)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp,, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    # driver = webdriver.Firefox(service_log_path='NUL', firefox_profile=profile)
    driver = webdriver.Firefox()
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    webpage_url = "https://hrginc.deltekfirst.com/hrginc/app/"
    driver.get(webpage_url)
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    return 0
 
#--------------------------------------------------------------------------------------------------------
 
# def addToDatabase(emp_id,emp_name,category,document_title,each_file):
#     global sql_db
#     try:
#         cursor = SQLconnection.cursor()
#         insert_statement = """INSERT INTO [Studio_C].[dbo].[Emp_Files1] (
#                     [Src_Id],[Employee_ Name],[Category],[File_Name],[Document_Title],[IsDownloaded],[DownloadedDateTime])
#                     VALUES (?, ?, ?, ?, ?, ?, ?);"""
#         cursor.execute(insert_statement,(emp_id, emp_name,category,each_file,document_title,1,time.strftime('%Y-%m-%d %H:%M:%S')))
#         SQLconnection.commit()
#         print(document_title)
#         cursor.execute("UPDATE [Studio_C].[dbo].[Emp_Files] SET [File_Name] = ?, [IsDownloaded] = ?, [DownloadedDateTime] = ? WHERE [Src_Id] = ? AND [Document_Title] = ? ;",
#         (each_file,1, time.strftime('%Y-%m-%d %H:%M:%S'), emp_id,document_title))
#         SQLconnection.commit()
#     except:
#         print('insert to database error')

def addToDatabase(src_id,emp_name,document_title,category):
    try:
            cursor = SQLconnection.cursor()
            insert_statement = """
                INSERT INTO [Orchestra].[dbo].[Emp_Files_emp_docs] ([Employee_Number],[Emp_Name],[Category],[File_Name],[IsDownloaded]
                ) VALUES (?, ?, ?, ?, ?)
            """
            values = (src_id,emp_name,category,document_title, 1)
            cursor.execute(insert_statement, values)
            SQLconnection.commit()
            print(src_id," :Added to Emp_Files")

    except Exception as d:
        print('Database error which is ', d)
 
#--------------------------------------------------------------------------------------------------------
 
def remove_files():
    files_to_delete = os.listdir(download_dir)
    for filename in files_to_delete:
        file_path = os.path.join(download_dir, filename)
        os.remove(file_path)  # Delete the file
 
#--------------------------------------------------------------------------------------------------------
 
def createFolder(sFolder):
    isExist = os.path.exists(sFolder)
    if not isExist:
        os.makedirs(sFolder)
    return(0)
 
#--------------------------------------------------------------------------------------------------------
 
def login():
    global min
    #username input xpath
    # Find and interact with username field
   
    xpath_username = '/html/body/div[3]/div[2]/div/div[2]/form/div[2]/input[1]'
    xpath_username = "//input[@id='userID']"
    element_username = driver.find_element(By.XPATH, xpath_username)
    element_username.click()
    element_username.clear()
    element_username.send_keys("PAYCOR")

    time.sleep(2)
 
    # Find and interact with password field
    xpath_password = '/html/body/div[3]/div[2]/div/div[2]/form/div[2]/input[2]'
    xpath_password ="//input[@id='password']"
    element_password = driver.find_element(By.XPATH, xpath_password)
    element_password.click()
    element_password.clear()
    element_password.send_keys("P@yC0R385!")
    time.sleep(2)

    # Find and click login button
    xpath_login_button = '/html/body/div[3]/div[2]/div/div[2]/form/div[6]/div[1]/button[1]'
    xpath_login_button = "//button[@id='loginBtn']"
    element_login_button = driver.find_element(By.XPATH, xpath_login_button)
    element_login_button.click()
    time.sleep(5)
    user_input = input("Please enter something: ")
#--------------------------------------------------------------------------------------------------------
 
def createFolder(sFolder):
     isExist = os.path.exists(sFolder)
     if not isExist:
        os.makedirs(sFolder)
        return(0)
 
def connectSQL():
    global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
    try:
        SQLconnection = pyodbc.connect('Driver={SQL Server};'
                            'Server=' + sql_server_name + ';'
                            'Database=' + sql_db + ';'
                            'UID=' + sql_user_name + ';'
                            'PWD=' + sql_password + ';'
                            'Trusted_Connection=no;')
        return(0)
    except Exception as e:
        print('SQL connection error')
        return(-1)
   
 

def rename_and_move_file(file_name, document_title, category, file_extension, target_dir, fld_name, download_dir, src_id, emp_name):
    try:
        # Define new folder path in the target directory based on category
        new_folder = os.path.join(target_dir, category)
        if not os.path.exists(new_folder):
            os.makedirs(new_folder)

        # Define the base new file path
        base_new_file_path = os.path.join(new_folder, document_title + file_extension)
        
        # Check if the file already exists and generate a new name if necessary
        new_file_path = base_new_file_path
        counter = 1
        while os.path.exists(new_file_path):
            new_file_path = os.path.join(new_folder, f"{document_title}_{counter}{file_extension}")
            counter += 1

        # Rename and move the file
        original_file = os.path.join(download_dir, file_name)
        shutil.move(original_file, new_file_path)
        print(f"Moved file '{file_name}' to '{new_file_path}'")

    # except Exception as e:
    #     print(f"Error: {e}")
        value = '"'+fld_name+'"'+","+'"'+document_title+'"'
        f= open('Berlin1_Downloaded_files.csv',"a")
        f.write(value)
        f.write('\n')
        f.close()
        addToDatabase(src_id,emp_name,document_title,category)
    except Exception as e:
        print(f"Error while renaming and moving the file: {e}")
 
def wait_for_download(download_dir, timeout=30):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = any([filename.endswith('.part') for filename in os.listdir(download_dir)])
        seconds += 1
    return not dl_wait
   
#--------------------------------------------------------------------------------------------------------
 
def searchanddownload(fld_name, src_id, emp_name):
    global SQLconnection,min
    try:
        #time.sleep(50)
        # createFolder("Z:/HRG/Documents/Vantage Point/"+fld_name)
        src_id = src_id.zfill(4)
        print("employee id",src_id)
        driver.get("https://hrginc.deltekfirst.com/hrginc/app/#!Employees/view/0/filesAndLinksTab/"+str(src_id)+"/hybrid")
        #  https://hrginc.deltekfirst.com/hrginc/app/#!Employees/view/0/filesAndLinksTab/01176/hybrid
        time.sleep(10)
        xpath="/html/body/div[2]/div[2]/div[3]/div[2]/div/div[2]/div/div[1]/div[4]/div[1]/div[3]/div/div[3]/div/div[2]/div/div[1]"
        element=driver.find_element(By.XPATH, xpath).text
        print(element)
        if src_id==element:
            print("ok")
            try:
                xpath='//*[@id="c5ad977be744429980ddb80415700a32Body"]/table/tfoot/tr/td'
                element=driver.find_element(By.XPATH,xpath)
                elementt=element.text

                if elementt=="There are no records to show in this grid.":
                    f= open('No Records.csv',"a")
                    f.write(f'{{"Files & Links": "{fld_name}": {elementt}}}\n')    
                
            except:
                xpath="/html/body/div[2]/div[2]/div[3]/div[2]/div/div[2]/div/div[1]/div[4]/div[2]/div/div/div[14]/div/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr/td[2]/span"
                elements=driver.find_elements(By.XPATH,xpath)
                elements=len(elements)

                f= open('count.csv',"a")
                f.write(f'{{{elements}: "{fld_name}"}}\n')

                for i in range(1,elements+1):
                    xpath='/html/body/div[2]/div[2]/div[3]/div[2]/div/div[2]/div/div[1]/div[4]/div[2]/div/div/div[14]/div/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr['+str(i)+']/td[1]'
                    element=driver.find_element(By.XPATH,xpath)
                    elementt=element.text
                    time.sleep(5)
                    createFolder("Z:/HRG/Documents/Vantage Point/"+fld_name+"/"+elementt)
                    xpath='/html/body/div[2]/div[2]/div[3]/div[2]/div/div[2]/div/div[1]/div[4]/div[2]/div/div/div[14]/div/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr['+str(i)+']/td[2]/span'
                    textelement=driver.find_element(By.XPATH,xpath).text                    

                    f= open('Records.csv',"a")
                    f.write(f'{{{elementt}: "{textelement}": "{fld_name}"}}\n')

                    xpath='/html/body/div[2]/div[2]/div[3]/div[2]/div/div[2]/div/div[1]/div[4]/div[2]/div/div/div[14]/div/div[1]/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr['+str(i)+']/td[2]/span'
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(10)

                    download_dir = r'C:/Users/RPATEAMADMIN/Downloads/'
                    files = os.listdir(download_dir)   
                    for file in files:
                        old_path = os.path.join(download_dir, file)
                        new_filename = file
                        #src=download_dir+file
                        dst="Z:/HRG/Documents/Vantage Point/"+fld_name+'/'+elementt+"/"
                        new_path = os.path.join(dst, new_filename)
                        shutil.move(old_path,new_path)  
                        value = '"'+fld_name +'"'+", downloaded"
                        print(value)
    except Exception as e:
        print("Employee documents Not downloaded\n")
        f= open('employee error.csv',"a")
        f.write(fld_name)
        f.write('\n')
        f.close()
        print("error ",e)
        return -1
    return 0
#--------------------------------------------------------------------------------------------------------
 
 
def main():
    global SQLconnection
    setup()
    read_config_file()
    connectSQL() 
    login()
 
    import csv
    rows = []
    # with open("a2.csv", 'r') as file:
    #     csvreader = csv.reader(file)
    #     header = next(csvreader)
    #     for row in csvreader:
    #         rows.append(row)
    # print(len(rows))
    rows = []
    select_statement = "SELECT [Employee_ID],[Last_Name],[First_Name],[IsDownloaded]"
    select_statement = select_statement + "FROM [HRG].[dbo].[Emp_Roster]"
    # select_statement += "WHERE [Company] = 'BerlinRosen';"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    rows = cursor.fetchall()
    cursor.close()
    #266,267             #Term 448(Vm24)
    for i in range(0,len(rows)):
        each = rows[i]
        src_id_1 = str(each[0]).strip()  
        src_id = src_id_1.zfill(5)
        print("initial print:",src_id)
        #src_id = src_id.zfill(6)
        last_name = each[1].strip()
        first_name = each[2].strip()
        #emp_name = each[1].strip()
        emp_name = last_name+", "+first_name
        print(emp_name)
        isdownloaded = each[3]
        # path_1 = r'C:\Users\RPATEAMADMIN\Downloads'
        path_1 = r'C:/Users/RPATEAMADMIN/Downloads/'
        files = os.listdir(path_1)
        # Iterate over the files and delete them
        for file in files:
            file_path = os.path.join(path_1, file)
            if os.path.isfile(file_path):
                break
                #os.remove(file_path)
                #print(f"Deleted {file_path}")
        
        path = "Z:/HRG/Documents/Vantage Point/"
        directory_contents = os.listdir(path)
        fld_name = emp_name+" ("+src_id+")"
        # if fld_name not in directory_contents:
        if not isdownloaded:
            print("entered_searchanddownload")
            # err_f = 1
            # count_flag = 1
            # while(err_f):
            res = searchanddownload(fld_name, src_id, emp_name)
            print("res=",res)
            if res == 0:
                cursor = SQLconnection.cursor()
                cursor.execute("UPDATE [HRG].[dbo].[Emp_Roster] SET [IsDownloaded] = ? WHERE [Employee_ID] = ?;",
                (1, src_id_1))
                print(src_id," :Updated in Emp_List")
                SQLconnection.commit()


            # elif res<0:
            #     print('Employee error '+src_id+" "+emp_name)
            #     continue

 
#--------------------------------------------------------------------------------------------------------
 
main()
