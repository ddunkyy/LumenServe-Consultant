# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from sqlalchemy import create_engine
import time 
import pandas as pd
import getpass
import base64
import math
import os.path
import os
import zipfile
from datetime import date


from tkinter import *
from tkinter import ttk
from tkinter import filedialog

gui = Tk()
gui.geometry("450x300")
gui.title("Input Values")

def getFolderPath():
    folder_selected = filedialog.askdirectory()
    folderPath.set(folder_selected)
    
folderPath = StringVar()
a = Label(gui ,text="Download Folder")
a.grid(row=0,column = 0)

E = Entry(gui,textvariable=folderPath)
E.grid(row=0,column=1)

btnFind = ttk.Button(gui, text="Browse Folder",command=getFolderPath)
btnFind.grid(row=0,column=2)

Label(gui, text = "CRM Email").grid(row = 1, sticky = W)
Label(gui, text = "CRM Password").grid(row = 2, sticky = W)

email = Entry(gui)
password = Entry(gui, show="*")

email.grid(row = 1, column = 1)
password.grid(row = 2, column = 1)

def getInput():
    a = email.get()
    b = password.get()
    c = folderPath.get()
    global params
    params = [c,a,b]
    
Button(gui, text = "submit",command = getInput).grid(row = 4, sticky = W)
gui.mainloop()

def wait_for_downloads():
    print("Waiting for downloads", end="")
    while any([filename.endswith(".crdownload") for filename in 
               os.listdir(params[0])]):
        time.sleep(2)
        print(".", end="")
    print("done!")

#when you view a dataset within the code, make it so that all columns appear, even if there are more than the default
pd.set_option('display.max_columns', None)

# %%
#SCRAPE THE CRM DATA

#change this to match where your chromedriver is located on your computer
driver = webdriver.Chrome(params[0] + '/chromedriver') 

#opens website, will open the window on your computer
driver.set_window_size(1400, 900)
driver.get("https://lumenserve.bitrix24.com/stream/?current_fieldset=SOCSERV")

#enter your login email into the Bitrix site
inputElement = driver.find_element_by_xpath('//*[@id="login"]')
inputElement.send_keys(params[1], Keys.ENTER) #edit this line to match the email of the user
time.sleep(2)
driver.find_element_by_xpath('//*[@id="authorize-layout"]/div/div[3]/div/form/div/div[5]/button[1]').click()

time.sleep(5)

#asks the user (you) to enter your password, where it is immediately encoded and remains hidden throughout the code
password = params[2]#getpass.getpass("Enter your password: ")
password = base64.b64encode(password.encode("utf-8"))

inputElement = driver.find_element_by_xpath('//*[@id="password"]')
inputElement.send_keys(base64.b64decode(password).decode("utf-8"), Keys.ENTER)
driver.find_element_by_xpath('//*[@id="authorize-layout"]/div/div[3]/div/form/div/div[3]/button[1]').click()
time.sleep(5)

driver.find_element_by_xpath('//*[@id="bx_left_menu_menu_crm_favorite"]/a').click()

time.sleep(2)

#navigates to the CRM tab of the site
driver.find_element_by_xpath('//*[@id="crm_control_panel_menu_menu_crm_company"]/a[1]/span[2]/span[2]').click()

time.sleep(5)

#navigates to the Companies tab of the CRM area of the site
driver.find_element_by_xpath('//*[@id="toolbar_company_list_button"]').click()

time.sleep(5)

#selects to download CRM into a CSV file 
driver.find_element_by_xpath('//*[@id="popup-window-content-menu-popup-toolbar_company_list_menu"]/div/div/span[4]/span[2]').click()

#checks the necessary boxes to include all details in the Companies file
driver.find_element_by_xpath('//*[@id="STEXPORT_COMPANY_MANAGER_MEg2VEg38A_LrpDlg_opt_REQUISITE_MULTILINE"]').click()
driver.find_element_by_xpath('//*[@id="STEXPORT_COMPANY_MANAGER_MEg2VEg38A_LrpDlg_opt_EXPORT_ALL_FIELDS"]').click()
driver.find_element_by_xpath('//*[@id="stexport_company_manager_meg2veg38a_lrpdlg"]/div[3]/span[1]').click()

time.sleep(150)
    
#clicks button to download the file to your downloads file
driver.find_element_by_xpath('//*[@id="popup-window-content-stexport_company_manager_meg2veg38a_lrpdlg"]/div/div[1]/a').click()

time.sleep(30)

#close the browser
driver.close()


# %%
#SCRAPE THE NOTAM DATA

#adjust this to be the path that leads to where your chromedriver is stored, most likely in Downloads folder
driver = webdriver.Chrome(params[0] + '/chromedriver') 


#opens disclaimer website, will open the window on your computer
driver.set_window_size(1920, 900)
driver.get("https://notams.aim.faa.gov/notamSearch/disclaimer.html")

#click 'I've read and understood terms' button, moves onto the actual website
button_path = "//html/body/table/tbody/tr[4]/td/button"
driver.find_element_by_xpath(button_path).click()


time.sleep(5)

#locates the dropdownm menu in the HTML path
element= driver.find_element_by_xpath('/html/body/div[2]/div/div/div[2]/div/div[2]/div[2]/div[1]/button')
 

#selects 'Free Text' from the drop down menu, now you are able to input text into the search field!
element.send_keys("Free Text")
element.send_keys("Free Text", Keys.ENTER)

time.sleep(5)

#inputs 'OBST TOWER LGT' into the text box to search
inputElement = driver.find_element_by_xpath('//*[@id="searchCriteria"]/div/input')
inputElement.send_keys('OBST TOWER LGT', Keys.ENTER)

time.sleep(7)

#downloads the searched data as an Excel file into your computer's downloads file
driver.find_element_by_xpath('/html/body/div[2]/div/div[3]/div[1]/div/div[3]/div[2]/div').click()

#if the browser is closing before the Excel has finished downloaded, increase this number from 60 to 120
time.sleep(60)

#close browser
driver.close()


# %%
#SCRAPE THE FCC DATA 

#change this to match where your chromedriver is located on your computer
driver = webdriver.Chrome(params[0] + '/chromedriver') 

#opens website, will open the window on your computer
driver.set_window_size(1400, 900)
driver.get("https://www.fcc.gov/uls/transactions/daily-weekly")

#downloads three necessary zip files to your downloads folder on your computer
driver.find_element_by_xpath('//*[@id="fcc-uls-transaction-files-weekly"]/div/div/table/tbody[5]/tr[2]/td[1]/a').click()
driver.find_element_by_xpath('//*[@id="fcc-uls-transaction-files-weekly"]/div/div/table/tbody[5]/tr[3]/td[1]/a').click()
driver.find_element_by_xpath('//*[@id="fcc-uls-transaction-files-weekly"]/div/div/table/tbody[5]/tr[4]/td[1]/a').click()

time.sleep(100)
#close browser
driver.close()

# %% [markdown]
# The following chunks of code will clean the files within the FCC data that are useful for the final combination. This includes variables from CO.dat, EN.dat, and RA.dat.

# %%
file = params[0] + '/' + 'r_tower.zip'
with zipfile.ZipFile(file,"r") as zip_ref:
    zip_ref.extractall(params[0])


# %%
#read in the .dat files to prepare for cleaning
tower = '/CO.dat'
path = params[0] + tower
co_messy = pd.read_table(path, error_bad_lines=False, encoding='cp1252')

tower = '/EN.dat'
path = params[0] + tower
en_messy = pd.read_table(path, error_bad_lines=False, encoding='cp1252')

tower = '/RA.dat'
path = params[0] + tower
ra_messy = pd.read_table(path, error_bad_lines=False, encoding='cp1252')


# %%
#splitting the data on the | symbol so all of the data has its own columns
def Split(column, on, df):
    '''
    Splits a column on any type of symbol
    column = The name of the column you want to split 
    on = the symbol that you want to split on 
    df = the dataframe containing the column that you want to split 
    '''
    new = df[column].str.split(on, expand = True)
    return new

#adding the rename function for later use 
def renameColumn(df, index, Name):
    '''
    Renames a column 
    df = dataframe with the column you want to rename 
    index = the index of the column you want to change of the name of the column you want to change 
    name = the new name that you want the column to be 
    '''
    New = df.rename(columns = {index: Name})
    return New 

#adding the drop column function for later use 
def dropColumns(df, column):
    '''Drops any column wanted within a pandas dataframe
           Df = Dataframe you want to drop a column from 
           Column = The name of the column that you want to drop 
    ''' 
    
    new = df.drop([column], axis=1)
    return new

#adding the add column function for later use 
def addColumn(df1, name, df2, column):
    '''Adds a column to pandas dataframe based off of exisiting column in another dataframe 
       df1 = dataframe you are adding the column to 
       name = name of the new column you are adding 
       df2 = name of the dataframe you are taking the column from
       column = name of the column you are putting into the new dataframe
    '''
    df1[name] = df2[column]
    return df1


# %%
co_messy.columns = ["New"]
en_messy.columns = ["New"]
ra_messy.columns = ["New"]


# %%
#working with the CO data using the functions above 
CO = Split('New','|',co_messy)
CO = dropColumns(CO, 16) 
CO = dropColumns(CO, 17)
CO = renameColumn(CO, 0, 'Record Type') 
CO = renameColumn(CO, 1, 'Content Indicator') 
CO = renameColumn(CO, 2, 'File Number') 
CO = renameColumn(CO, 3, 'Registration Number') 
CO = renameColumn(CO, 4, 'Unique System Identifier') 
CO = renameColumn(CO, 5, 'Coordinate Type') 
CO = renameColumn(CO, 6, 'Latitude Degrees') 
CO = renameColumn(CO, 7, 'Latitude Minutes') 
CO = renameColumn(CO, 8, 'Latitude Seconds') 
CO = renameColumn(CO, 9, 'Latitude Direction') 
CO = renameColumn(CO, 10, 'Latitude_Total_Seconds') 
CO = renameColumn(CO, 11, 'Longitude Degrees') 
CO = renameColumn(CO, 12, 'Longitude Minutes') 
CO = renameColumn(CO, 13, 'Longitude Seconds') 
CO = renameColumn(CO, 14, 'Longitude Direction') 
CO = renameColumn(CO, 15, 'Longitude Total Seconds') 


# %%
CO_to_use = CO[['File Number', 'Registration Number', 'Unique System Identifier', 'Latitude Degrees', 
               'Latitude Minutes', 'Latitude Seconds', 'Latitude Direction', 'Longitude Degrees', 
                'Longitude Seconds', 'Longitude Direction']]


# %%
#working with the EN data using the functions above 
EN = Split('New','|',en_messy)
EN = renameColumn(EN, 0, 'Record Type')
EN = renameColumn(EN, 1, 'Content Inidcator')
EN = renameColumn(EN, 2, 'File Number')
EN = renameColumn(EN, 3, 'Registration Number')
EN = renameColumn(EN, 4, 'Unique System Identifier')
EN = renameColumn(EN, 5, 'Contact Type')
EN = renameColumn(EN, 6, 'Entity Type')
EN = renameColumn(EN, 7, 'Entity Type - Other')
EN = renameColumn(EN, 8, 'Licensee ID')
EN = renameColumn(EN, 9, 'Entity Name')
EN = renameColumn(EN, 10, 'First Name')
EN = renameColumn(EN, 11, 'MI')
EN = renameColumn(EN, 12, 'Last Name')
EN = renameColumn(EN, 13, 'Suffix')
EN = renameColumn(EN, 14, 'Phone')
EN = renameColumn(EN, 15, 'Fax Number')
EN = renameColumn(EN, 16, 'Internet Address')
EN = renameColumn(EN, 17, 'Street Address')
EN = renameColumn(EN, 18, 'Street Address 2')
EN = renameColumn(EN, 19, 'PO Box')
EN = renameColumn(EN, 20, 'City')
EN = renameColumn(EN, 21, 'State')
EN = renameColumn(EN, 22, 'Zip Code')
EN = renameColumn(EN, 23, 'Attention')
EN = renameColumn(EN, 24, 'FRN')


# %%
EN_to_use = EN[['File Number', 'Registration Number', 'Unique System Identifier', 'Entity Name', 'Phone', 
               'Fax Number', 'Internet Address', 'Street Address', 'Street Address 2', 'PO Box', 'City',
               'State', 'Zip Code', 'FRN']]


# %%
r_tower = pd.merge(CO_to_use, EN_to_use, on='Unique System Identifier', how='outer')

r_tower_cols = []
for col in r_tower.columns:
    r_tower_cols.append(col)
    
r_tower = r_tower.astype({"Registration Number_x": float, 'Registration Number_y': float})
r_tower_list = r_tower.values.tolist()

for entry in r_tower_list:
    if entry[1] != entry[11] and math.isnan(entry[1]):
        entry[1] = entry[11]
    if entry[1] != entry[11] and math.isnan(entry[11]):
        entry[11] = entry[1]
        
r_tower = pd.DataFrame(r_tower_list, columns = r_tower_cols)
r_tower = r_tower.drop(['Registration Number_y'], axis = 1)
r_tower = r_tower.rename(columns={"Registration Number_x": "Registration Number"})


# %%
#working with the RA data using the functions above 
RA = Split('New','|',ra_messy)
split_RA9 = Split(9, '/', RA)
split_RA9 = renameColumn(split_RA9, 0, 'Month')
split_RA9 = renameColumn(split_RA9, 1, 'Day')
split_RA9 = renameColumn(split_RA9, 2, 'Year')
RA = addColumn(RA, 'Date Entered Month', split_RA9, 'Month' )
RA = addColumn(RA, 'Date Entered Day', split_RA9, 'Day' )
RA = addColumn(RA, 'Date Entered Year', split_RA9, 'Year' )
RA = dropColumns(RA, 9)
split_RA10 = Split(10, '/', RA)
split_RA10 = renameColumn(split_RA10, 0, 'Month')
split_RA10 = renameColumn(split_RA10, 1, 'Day')
split_RA10 = renameColumn(split_RA10, 2, 'Year')
RA = addColumn(RA, 'Date Received Month', split_RA10, 'Month' )
RA = addColumn(RA, 'Date Received Day', split_RA10, 'Day' )
RA = addColumn(RA, 'Date Received Year', split_RA10, 'Year' )
RA = dropColumns(RA, 10)
split_RA11 = Split(11, '/', RA)
split_RA11 = renameColumn(split_RA11, 0, 'Month')
split_RA11 = renameColumn(split_RA11, 1, 'Day')
split_RA11 = renameColumn(split_RA11, 2, 'Year')
RA = addColumn(RA, 'Date Issued Month', split_RA11, 'Month' )
RA = addColumn(RA, 'Date Issued Day', split_RA11, 'Day' )
RA = addColumn(RA, 'Date Issued Year', split_RA11, 'Year' )
RA = dropColumns(RA, 11)
split_RA12 = Split(12, '/', RA)
split_RA12 = renameColumn(split_RA12, 0, 'Month')
split_RA12 = renameColumn(split_RA12, 1, 'Day')
split_RA12 = renameColumn(split_RA12, 2, 'Year')
RA = addColumn(RA, 'Date Constructed Month', split_RA12, 'Month' )
RA = addColumn(RA, 'Date Constructed Day', split_RA12, 'Day' )
RA = addColumn(RA, 'Date Constructed Year', split_RA12, 'Year' )
RA = dropColumns(RA, 12)
split_RA13 = Split(13, '/', RA)
split_RA13 = renameColumn(split_RA13, 0, 'Month')
split_RA13 = renameColumn(split_RA13, 1, 'Day')
split_RA13 = renameColumn(split_RA13, 2, 'Year')
RA = addColumn(RA, 'Date Dismantled Month', split_RA13, 'Month' )
RA = addColumn(RA, 'Date Dismantled Day', split_RA13, 'Day' )
RA = addColumn(RA, 'Date Dismantled Year', split_RA13, 'Year' )
RA = dropColumns(RA, 13)
split_RA14 = Split(14, '/', RA)
split_RA14 = renameColumn(split_RA14, 0, 'Month')
split_RA14 = renameColumn(split_RA14, 1, 'Day')
split_RA14 = renameColumn(split_RA14, 2, 'Year')
RA = addColumn(RA, 'Date Action Month', split_RA14, 'Month' )
RA = addColumn(RA, 'Date Action Day', split_RA14, 'Day' )
RA = addColumn(RA, 'Date Action Year', split_RA14, 'Year' )
RA = dropColumns(RA, 14)
split_RA33 = Split(33, '/', RA)
split_RA33 = renameColumn(split_RA33, 0, 'Month')
split_RA33 = renameColumn(split_RA33, 1, 'Day')
split_RA33 = renameColumn(split_RA33, 2, 'Year')
RA = addColumn(RA, 'Date FAA Determination Issued Month', split_RA33, 'Month' )
RA = addColumn(RA, 'Date FAA Determination Issued Day', split_RA33, 'Day' )
RA = addColumn(RA, 'Date FAA Determination Issued Year', split_RA33, 'Year' )
RA = dropColumns(RA, 33)
split_RA42 = Split(42, '/', RA)
split_RA42 = renameColumn(split_RA42, 0, 'Month')
split_RA42 = renameColumn(split_RA42, 1, 'Day')
split_RA42 = renameColumn(split_RA42, 2, 'Year')
RA = addColumn(RA, 'Date Signed Month', split_RA42, 'Month' )
RA = addColumn(RA, 'Date Signed Day', split_RA42, 'Day' )
RA = addColumn(RA, 'Date Signed Year', split_RA42, 'Year' )
RA = dropColumns(RA, 42)


# %%
RA = renameColumn(RA, 0, 'Record Type')
RA = renameColumn(RA, 1, 'Content Indicator')
RA = renameColumn(RA, 2, 'File Number')
RA = renameColumn(RA, 3, 'Registration Number')
RA = renameColumn(RA, 4, 'Unique System Identifier')
RA = renameColumn(RA, 5, 'Application Purpose')
RA = renameColumn(RA, 6, 'Previous Purpose')
RA = renameColumn(RA, 7, 'Input Source Code')
RA = renameColumn(RA, 8, 'Status Code')
RA = renameColumn(RA, 15, 'Archive Flag Code')
RA = renameColumn(RA, 16, 'Version')
RA = renameColumn(RA, 17, 'Signature First Name')
RA = renameColumn(RA, 18, 'Signature Middle Initial')
RA = renameColumn(RA, 19, 'Signature Last Name')
RA = renameColumn(RA, 20, 'Signature Suffix')
RA = renameColumn(RA, 21, 'Signature Title')
RA = renameColumn(RA, 22, 'Invalid Signature')
RA = renameColumn(RA, 23, 'Structure_Street Address')
RA = renameColumn(RA, 24, 'Structure_City')
RA = renameColumn(RA, 25, 'Structure_State Code')
RA = renameColumn(RA, 26, 'County Code')
RA = renameColumn(RA, 27, 'Zip Code')
RA = renameColumn(RA, 28, 'Height of Structure')
RA = renameColumn(RA, 29, 'Ground Elevation')
RA = renameColumn(RA, 30, 'Overall Height Above Ground')
RA = renameColumn(RA, 31, 'Overall Height AMSL')
RA = renameColumn(RA, 32, 'Structure Type')
RA = renameColumn(RA, 34, 'FAA Study Number')
RA = renameColumn(RA, 35, 'FAA Circular Number')
RA = renameColumn(RA, 36, 'Specification Option')
RA = renameColumn(RA, 37, 'Painting and Lighting')
RA = renameColumn(RA, 38, 'Proposed Marking and Lighting')
RA = renameColumn(RA, 39, 'Marking and Lighting Other')
RA = renameColumn(RA, 40, 'FAA EMI Flag')
RA = renameColumn(RA, 41, 'NEPA Flag')
RA = renameColumn(RA, 43, 'Assignor Signature Last Name')
RA = renameColumn(RA, 44, 'Assignor Signature First Name')
RA = renameColumn(RA, 45, 'Assignor Signature Middle Initial')
RA = renameColumn(RA, 46, 'Assignor Signature Suffix')
RA = renameColumn(RA, 47, 'Assignor Signature Title')
RA = renameColumn(RA, 48, 'Assignor Date Signed')


# %%
RA_to_use = RA[['File Number', 'Registration Number', 'Unique System Identifier', 'Structure_Street Address', 
               'Structure_City', 'Structure_State Code', 'County Code', 'Zip Code', 'Overall Height Above Ground', 
                'Structure Type', 'FAA Study Number', 'FAA Circular Number', 'Date Constructed Month', 
                'Date Constructed Day', 'Date Constructed Year']]


# %%
r_tower = pd.merge(r_tower, RA_to_use, on='Unique System Identifier', how='outer')

r_tower_list = r_tower['Registration Number_x'].tolist()
r_tower_list2 = r_tower['Registration Number_y'].tolist()

r_tower = r_tower.astype({'Registration Number_x': float})
r_tower = r_tower.astype({'Registration Number_y': float})
r_tower_list = r_tower.values.tolist()

r_tower_cols = []
for col in r_tower.columns:
    r_tower_cols.append(col)
    
for entry in r_tower_list:
    if entry[1] != entry[23] and math.isnan(entry[1]):
        entry[1] = entry[23]
    if entry[1] != entry[23] and math.isnan(entry[23]):
        entry[23] = entry[1]
        
r_tower = pd.DataFrame(r_tower_list, columns = r_tower_cols)
r_tower = r_tower.drop(['Registration Number_y'], axis = 1)
r_tower = r_tower.rename(columns={"Registration Number_x": "Registration Number"})

# %% [markdown]
# The following chunks of code will clean the files within the NOTAM data.

# %%
partialFileName = "fnsNotams"
a = [f for f in os.listdir(params[0]) if partialFileName == f[:len(partialFileName)]]

notam = a[0]
for entry in a:
    if entry[16:20] > notam[16:20] and entry:
        notam = entry
    
    if entry[10:12] > notam[10:12]:
        notam = entry
    
    if entry[10:12] == notam[10:12]:
        if entry[13:15] > notam[13:15]:
            notam = entry
    
    if entry[10:12] == notam[10:12] and entry[13:15] == notam[13:15]:
        if entry[21:27] > notam[21:27]:
            notam = entry


# %%
file = params[0] + '/' + notam
Notam = pd.read_excel(file)


#Rename existing columns
Notam.columns = ["Location", "NOTAM #/LTA#", "Class", "Issue Date (UTC)", "Effective Date (UTC)", 
                 "Expiration Date (UTC)", "Condition"]

#drop unnecessary rows
def dropRows(df, index):
    ''' 
    Drop un-needed rows that appear
    df = database that you want to drop rows from
    index = index of the row that you want to drop
    '''
    new = df.drop([index])
    
    return new 
Notam = dropRows(Notam, 0) 
Notam = dropRows(Notam, 1) 
Notam = dropRows(Notam, 2) 
Notam = dropRows(Notam, 3) 

#Initializing SQL  and only selecting towers where Class is Obstruction
engine = create_engine('sqlite://', echo=False) 
Notam.to_sql('Notam2', con=engine) 
Notam = engine.execute("SELECT * FROM Notam2 WHERE Class LIKE 'O%'").fetchall() 
Notam = pd.DataFrame(Notam) 

#drop unnecessary columns
def dropColumns(df, column):
    '''Drops any column wanted within a pandas dataframe
           Df = Dataframe you want to drop a column from 
           Column = The name of the column that you want to drop 
    ''' 
    
    new = df.drop([column], axis=1)
    return new
Notam = dropColumns(Notam, 0)

#rename more columns
def renameColumn(df, index, Name):
    '''
    Renames a column 
    df = dataframe with the column you want to rename 
    index = the index of the column you want to change of the name of the column you want to change 
    name = the new name that you want the column to be 
    '''
    New = df.rename(columns = {index: Name})
    return New 
Notam = renameColumn(Notam, 1, 'Location')
Notam = renameColumn(Notam, 2, 'NOTAMLTA')
Notam = renameColumn(Notam, 3, 'Class')
Notam = renameColumn(Notam, 4, 'Issue Date (UTC)')
Notam = renameColumn(Notam, 5, 'EffectiveDate')
Notam = renameColumn(Notam, 6, 'Expiration Date (UTC)')
Notam = renameColumn(Notam, 7, 'Condition')

#Split columns that contain more than one piece of data
def Split(column, on, df):
    '''
    Splits a column on any type of symbol
    column = The name of the column you want to split 
    on = the symbol that you want to split on 
    df = the dataframe containing the column that you want to split 
    '''
    new = df[column].str.split(on, expand = True)
    return new
Split_NOTAM = Split('NOTAMLTA','/',Notam) 
Split_NOTAM = renameColumn(Split_NOTAM, 0, 'NOTAM') 
Split_NOTAM = renameColumn(Split_NOTAM, 1, 'LTA') 

#adding columns to account for splitting above columns
def addColumn(df1, name, df2, column):
    '''Adds a column to pandas dataframe based off of exisiting column in another dataframe 
       df1 = dataframe you are adding the column to 
       name = name of the new column you are adding 
       df2 = name of the dataframe you are taking the column from
       column = name of the column you are putting into the new dataframe
    '''
    df1[name] = df2[column]
    return df1
Notam = addColumn(Notam, 'NOTAM', Split_NOTAM, 'NOTAM') 
Notam = addColumn(Notam, 'LTA', Split_NOTAM, 'LTA') 
Notam = dropColumns(Notam, 'NOTAMLTA') 

#splits and adds more columns for further cleaning
Split_Effective = Split('EffectiveDate','/', Notam) 
Split_Effective = renameColumn(Split_Effective, 0, 'Effective Month (UTC)') 
Split_Effective = renameColumn(Split_Effective, 1, 'Effective Day (UTC)') 
Split_Effective = renameColumn(Split_Effective, 2, 'EffectiveYear') 
Notam = addColumn(Notam, 'Effective Month (UTC)', Split_Effective, 'Effective Month (UTC)') 
Notam = addColumn(Notam, 'Effective Day (UTC)', Split_Effective, 'Effective Day (UTC)') 
Notam = addColumn(Notam, 'EffectiveYear', Split_Effective, 'EffectiveYear') 
Notam = dropColumns(Notam, 'EffectiveDate') 

#More splitting and adding
def splitSpace(column, df):
    '''Splits rows in a column on the basis of a single space
       Column = name of the column that you want to split 
       df = Dataframe that contains the column in which you are splitting 
    '''
    new = df[column].str.split(expand = True)
    return new     

Split_Year = splitSpace('EffectiveYear',Notam) 
Split_Year = renameColumn(Split_Year, 0, "Effective Year (UTC)") 
Split_Year = renameColumn(Split_Year, 1, "Effective Military Time (UTC)") 
Notam = addColumn(Notam, 'Effective Year (UTC)', Split_Year, 'Effective Year (UTC)') 
Notam = addColumn(Notam, 'Effective Military Time (UTC)', Split_Year, 'Effective Military Time (UTC)') 
Notam = dropColumns(Notam, 'EffectiveYear') 
Split_Condition = splitSpace('Condition', Notam) 
Split_Condition = dropColumns(Split_Condition, 0) 
Split_Condition = dropColumns(Split_Condition, 1) 
Split_Condition = dropColumns(Split_Condition, 2) 
Split_Condition = dropColumns(Split_Condition, 3) 
Split_Condition = dropColumns(Split_Condition, 4)  
Split_Condition = dropColumns(Split_Condition, 5) 
Split_Condition = dropColumns(Split_Condition, 19) 
Split_Condition = dropColumns(Split_Condition, 20) 
Split_Condition = dropColumns(Split_Condition, 21) 
Split_Condition = renameColumn(Split_Condition, 7, 'ASR') 
Split_Condition = renameColumn(Split_Condition, 13, 'AGL FT') 
Split_Condition = renameColumn(Split_Condition, 15, 'Light Status') 
Notam = addColumn(Notam, 'ASR', Split_Condition, 'ASR') 
Notam = addColumn(Notam, 'AGL ft', Split_Condition, 'AGL FT') 
Notam = addColumn(Notam, 'Light Status', Split_Condition, 'Light Status') 

#remove unnecessary symbols from columns
def removeSigns(df, column1, df2, column2, symbol):
    '''Removes any symbol, such as a comma, from a column in a pandas dataframe
       df = dataframe you are removing the symbol from
       column1 = Name of the column you are removing the symbol from
       df2 = dataframe you are removing the symbol from
       column2 = Name of the column you are removing the symbol from
       symbol = The symbol you want to remove
    '''
    df[column1] = df2[column2].str.replace(symbol,"")
    return df
Notam = removeSigns(Notam, "ASR", Notam, "ASR",")")
Notam = removeSigns(Notam, "AGL ft", Notam, "AGL ft", "(")
Notam = removeSigns(Notam, "AGL ft", Notam, "AGL ft", "FT")
Notam = dropColumns(Notam, 'Condition')
Notam = dropColumns(Notam, 'Issue Date (UTC)')
Notam = renameColumn(Notam, 'Expiration Date (UTC)', "Expiration")

#replaces text of column with appropriate text
def replaceString(df1, column1, df2, column2, text1, text2):
    df1[column1] = Notam[column2].str.replace(text1, text2)
    return df1
Notam = replaceString(Notam, 'Light Status', Notam, 'Light Status', 'U/S', 'ok')

#edits to correct NOTAM expiration dates
def changeExpiration(df, column, condition, column2, sign, df2):
    for i in df[column]:
        if i != condition:
            new = Split(column2,sign, df2)
    return new
Split_Expiration = changeExpiration(Notam, 'Expiration', 'PERM', 'Expiration', '/', Notam)
Split_Expiration = renameColumn(Split_Expiration, 0, 'Expiration Month') 
Split_Expiration = renameColumn(Split_Expiration, 1, 'Expiration Day') 
Split_Expiration = renameColumn(Split_Expiration, 2, 'ExpirationYear')
Split_Expiration = Split_Expiration.fillna('PERM') 
Notam = addColumn(Notam, 'Expiration Month', Split_Expiration, 'Expiration Month') 
Notam = addColumn(Notam, 'Expiration Day', Split_Expiration, 'Expiration Day') 
Notam = addColumn(Notam, 'ExpirationYear', Split_Expiration, 'ExpirationYear') 
Split_ExpirationYear = splitSpace('ExpirationYear',Notam)
Split_ExpirationYear = Split_ExpirationYear.fillna('PERM')
Split_ExpirationYear = renameColumn(Split_ExpirationYear, 0, 'Expiration Year') 
Split_ExpirationYear = renameColumn(Split_ExpirationYear, 1, 'Expiration Military Time')
Notam = addColumn(Notam, 'Expiration Year', Split_ExpirationYear, 'Expiration Year') 
Notam = addColumn(Notam, 'Expiration Military Time', Split_ExpirationYear, 'Expiration Military Time')
Notam = dropColumns(Notam, 'ExpirationYear')
Notam = dropColumns(Notam, 'Expiration')

#selects appropriate columns
Notam = Notam[['Location','ASR','Class','Effective Month (UTC)','Effective Day (UTC)','Effective Year (UTC)', 
               'Effective Military Time (UTC)','Expiration Month', 'Expiration Day', 'Expiration Year', 
               'Expiration Military Time', 'NOTAM', 'LTA', 'AGL ft', 'Light Status']]

# %% [markdown]
# The following code chunks will clean the CRM data file to prepare for future combination

# %%
partialFileName = "COMPANY"
a = [f for f in os.listdir(params[0]) if partialFileName == f[:len(partialFileName)]]
for i in range(len(a)):
    if i == 0:
        crm = a[i]
    else:
        if a[i][8:16] > crm[8:16]:
            crm = a[i]

file = params[0] + '/' + crm
CRM = pd.read_csv(file, delimiter = ";")


# %%
HG = pd.read_excel("Holy Grail Target Twrs 2020 08 25.xlsx")
HG = HG.astype({"ASR": float})

CRM_columns = []
for col in CRM.columns:
    CRM_columns.append(col)

#cleans the company name columns in both datasets to remove bad characters and change it to lowercase
bad_chars = [' ', '.', ',', '=']

name = CRM['Company Name']

for i in name:
    i = str(i).casefold()
    i = ''.join((filter(lambda j: j not in bad_chars, i)))

CRM['Company Name'] = CRM['Company Name'].str.lower()

bad_chars = [' ', '.', ',', '=']

name = HG['Owner']

for i in name:
    i = str(i).casefold()
    i = ''.join((filter(lambda j: j not in bad_chars, i)))

HG['Owner'] = HG['Owner'].str.lower()

CRM_names = CRM['Company Name'].tolist()
HG_names = HG['Owner'].tolist()


# %%
#find all instances where the owner name column matches in the CRM and Holy Grail
names = []
for entry in CRM_names:
    if entry in HG_names:
        names.append(entry)

CRM_match = CRM[CRM['Company Name'].isin(names)]
HG_match = HG[HG['Owner'].isin(names)]

CRM_match_small = CRM_match[["ID", "Company Name", 'Company Type', 'ASR Number']]

#pull out the unique, 1-1 matches of Company Name and Owner Name
HG_match['group_count'] = HG_match.groupby(by='Owner')['Owner'].transform('count')
HG_match = HG_match.sort_values(by=['Owner'])
HG_match = HG_match[HG_match['group_count']==1]

HG_match_list = HG_match.values.tolist()

CRM_match_list = CRM_match_small.values.tolist()


# %%
#adjust all CRM ASR values to match those in Holy Grail, where the names match
for entry in CRM_match_list:
    for entry2 in HG_match_list:
        if entry[1] == entry2[11] and (entry[3] == 0.0 or math.isnan(entry[3])) and pd.notna(entry[1]):
            entry[3] = entry2[0]
            
CRM_list = CRM.values.tolist()

#scaling up this smaller CRM matching to implement the changes in the larger CRM dataset
for entry in CRM_list:
    for entry2 in CRM_match_list:
        if entry[2] == entry2[1] and entry[66] != entry2[3]:
            entry[66] = entry2[3]


# %%
CRM = pd.DataFrame(CRM_list, columns = CRM_columns)

CRM = CRM.rename(columns={"ASR Number": "ASR"})

CRM_ASR = CRM[(CRM['ASR'] != 0.0) & (CRM['ASR'].notna())]
CRM_no_ASR = CRM[(CRM['ASR'] == 0.0) | (CRM['ASR'].isna())]


# %%
#create a column that represents the full structure addresses in the CRM and Holy Grail
CRM['Full Address']=CRM['Address'].astype(str)+'_'+CRM['City']+'_'+CRM['State']
CRM['Full Address'] = CRM['Full Address'].str.lower()

HG['Full Address']=HG['Structure Address'].astype(str)+'_'+HG['Structure City']+'_'+HG['Structure State']
HG['Full Address'] = HG['Full Address'].str.lower()

CRM_address = CRM['Full Address'].tolist()
HG_address = HG['Full Address'].tolist()

#find all instances where the owner name column matches in the CRM and Holy Grail
addresses = []
for entry in CRM_address:
    if entry in HG_address and pd.notna(entry):
        addresses.append(entry)

CRM_match2 = CRM[CRM['Full Address'].isin(addresses)]
HG_match2 = HG[HG['Full Address'].isin(addresses)]

#pull out the unique, 1-1 matches of full structure address
HG_match2['group_count'] = HG_match2.groupby(by='Full Address')['Full Address'].transform('count')
HG_match2 = HG_match2.sort_values(by=['Full Address'])
HG_match2 = HG_match2[HG_match2['group_count']==1]


# %%
HG_match2_list = HG_match2.values.tolist()
CRM_match2_list = CRM_match2.values.tolist()

#adjust all CRM ASR values to match those in Holy Grail, where the addresses match
for entry in CRM_match2_list:
    for entry2 in HG_match2_list:
        if entry[116] == entry2[25] and (entry[66] == 0.0 or math.isnan(entry[66])):
            entry[66] = entry2[0]

CRM_cols = []
for col in CRM.columns:
    CRM_cols.append(col)
    
CRM_list = CRM.values.tolist()

#scaling up this smaller CRM matching to implement the changes in the larger CRM dataset
for entry in CRM_list:
    for entry2 in CRM_match2_list:
        if entry[116] == entry2[116] and entry[66] != entry2[66]:
            entry[66] = entry2[66]

CRM = pd.DataFrame(CRM_list, columns = CRM_cols)


# %%
CRM_ASR = CRM[(CRM['ASR'] != 0.0) & (CRM['ASR'].notna())]
CRM_no_ASR = CRM[(CRM['ASR'] == 0.0) | (CRM['ASR'].isna())]

# %% [markdown]
# The following chunks of code will combine the FCC, NOTAM, and CRM datasets

# %%
#remove rows in the data where the Light Status is something other than "OUT" or "ok"
notam = Notam[(Notam['Light Status'] =='OUT') | (Notam['Light Status'] =='ok')]

notam_cols = []
for col in notam.columns:
    notam_cols.append(col)

notam_list = notam.values.tolist()

#remove rows in the NOTAM data where the ASR value does not follow the format that the value should
notam_usable = []
for entry in notam_list:
    if entry[1] != 'UNKNOWN' and '-' not in entry[1]:
        notam_usable.append(entry)

notam_usable = pd.DataFrame(notam_usable, columns = notam_cols)


# %%
r_tower = r_tower.rename(columns={"Registration Number": "ASR"})

notam_usable = notam_usable.astype({'ASR':float})
r_tower = r_tower.astype({'ASR':float})


#merge the FCC data with the NOTAM data on the ASR column (note: this keeps all columns and rows from both)
r_tower_merged = pd.merge(r_tower, notam_usable, on='ASR', how='outer')
r_tower_merged = r_tower_merged.drop(['File Number_x', 'File Number_y', 'File Number'], axis=1)

#creates an empty column that will indicate whether or not the NOTAM data matched (to be later filtered on)
r_tower_merged['NOTAM Included'] = ""

r_tower_merged_cols = []
for col in r_tower_merged.columns:
    r_tower_merged_cols.append(col)

r_tower_merged_list = r_tower_merged.values.tolist()

#adds values to the NOTAM Included column, inputting a 1 if the NOTAM matched and a 0 if it did not
for entry in r_tower_merged_list:
    if entry[45] == 'ok' or entry[45] == 'OUT':
        entry[46] = 1
    else:
        entry[46] = 0
        
r_tower_merged = pd.DataFrame(r_tower_merged_list, columns = r_tower_merged_cols )


# %%
CRM = CRM.astype({"ASR": float})

#pulls out only the specific columns that we want to include in the combined CRM
#NOTE: THESE CAN EASILY BE CHANGED TO MATCH THE DESIRES OF YOUR COMPANY, SIMPLY BY REMOVING OR ADDING VARIABLES FROM
#THIS LIST

CRM = CRM[['Company Name', 'Company Type', 'Employees', 'Work Phone', 'Mobile', 'Fax', 'Home Phone', 'Corporate Website',
'Work E-mail', 'Home E-mail', 'Responsible', 'Industry', 'First Name', 'Last Name', 'Contact Verified?',
'Proposal Requested', 'Propsal Walk-Thru Scheduled', 'Proposal Walk-Thru Complete', 'ASR', 'Already Upgraded to LED',
'Do Not Call', 'Ready to Upgrade to LED', 'Interested in Monitoring and Compliance', 'City', 'State', 'Zip Code',
'Address']]

CRM_ASR = CRM[(CRM['ASR'] != 0.0) & (CRM['ASR'].notna())]
CRM_no_ASR = CRM[(CRM['ASR'] == 0.0) | (CRM['ASR'].isna())]


# %%
#merge the combined FCC and NOTAM data with the CRM data that contains a valid ASR, on the ASR column
#note: all columns and rows are kept from both datasets

r_tower_merged2 = pd.merge(r_tower_merged, CRM_ASR, on='ASR', how='outer')

r_tower_merged2['CRM Included'] = ""

r_tower_merged2_cols = []
for col in r_tower_merged2.columns:
    r_tower_merged2_cols.append(col)

r_tower_merged2_list = r_tower_merged2.values.tolist()

#adds values to the CRM Included column, inputting a 1 if the CRM matched and a 0 if it did not
for entry in r_tower_merged2_list:
    if pd.isna(entry[47]) & pd.isna(entry[48]) & pd.isna(entry[49]) & pd.isna(entry[50]) & pd.isna(entry[51]) & pd.isna(entry[52]) & pd.isna(entry[53]) & pd.isna(entry[54]) & pd.isna(entry[55]) & pd.isna(entry[56]) & pd.isna(entry[57]) & pd.isna(entry[58]) & pd.isna(entry[59]) & pd.isna(entry[60]) & pd.isna(entry[61]) & pd.isna(entry[62]) & pd.isna(entry[63]) & pd.isna(entry[64]) & pd.isna(entry[65]) & pd.isna(entry[66]) & pd.isna(entry[67]) & pd.isna(entry[68]) & pd.isna(entry[69]) & pd.isna(entry[70]) & pd.isna(entry[71]) & pd.isna(entry[72]):
        entry[73] = 0
    else:
        entry[73] = 1
r_tower_merged2 = pd.DataFrame(r_tower_merged2_list, columns = r_tower_merged2_cols )


# %%
#creates 3 subsets of the data based on the columns NOTAM Included and CRM Inluded

#subset where the NOTAM matched by ASR
FCC_with_notam = r_tower_merged[(r_tower_merged['NOTAM Included'] ==1)]

#subset where the CRM matched by ASR
FCC_with_CRM = r_tower_merged2[(r_tower_merged2['CRM Included'] ==1)]

#subset where both the NOTAM and CRM matched by ASR
FCC_with_NOTAM_CRM = r_tower_merged2[(r_tower_merged2['CRM Included'] == 1) & r_tower_merged2['NOTAM Included'] == 1]


# %%
#obtain holier grail as an Excel file on your computer

excel_name = 'holier_grail.{}.xlsx'.format(date.today())
r_tower_merged2.to_excel(excel_name)


# %%
#obtain holier grail as a CSV file on your computer

csv_name = 'holier_grail.{}.csv'.format(date.today())
r_tower_merged2.to_csv(csv_name)


# %%
#obtain holier grail only where the CRM data matched as an Excel file on your computer


excel_name = 'FCC_with_CRM.{}.xlsx'.format(date.today())
FCC_with_CRM.to_excel(excel_name)


# %%
#obtain holier grail only where the CRM data matched as a CSV file on your computer


csv_name = 'FCC_with_CRM.{}.csv'.format(date.today())
FCC_with_CRM.to_csv(csv_name)


# %%
#obtain holier grail only where the NOTAM data matched as an Excel file on your computer

excel_name = 'FCC_with_NOTAM.{}.xlsx'.format(date.today())
FCC_with_notam.to_excel(excel_name)


# %%
#obtain holier grail only where the NOTAM data matched as a CSV file on your computer

csv_name = 'FCC_with_NOTAM.{}.csv'.format(date.today())
FCC_with_notam.to_csv(csv_name)


# %%
#obtain holier grail only where both the NOTAM data and the CRM data matched as an Excel file on your computer


excel_name = 'FCC_with_NOTAM_CRM.{}.xlsx'.format(date.today())
FCC_with_NOTAM_CRM.to_excel(excel_name)


# %%
#obtain holier grail only where both the NOTAM data and the CRM data matched as a CSV file on your computer


csv_name = 'FCC_with_NOTAM_CRM.{}.csv'.format(date.today())
FCC_with_NOTAM_CRM.to_csv(csv_name)

#%%

done = Tk()

done.title("LS Full finished")

lbl = Label(done, text="The program has completed. You can now close every windows")

lbl.grid(column=0, row=0)

done.mainloop()

# %%
