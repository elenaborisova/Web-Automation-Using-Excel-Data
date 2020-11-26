import xlrd
from selenium import webdriver
import tkinter
import tkinter.filedialog
import os


# Navigate to Excel file
root = tkinter.Tk()
currdir = os.getcwd()
tempdir = tkinter.filedialog.askopenfilename(parent=root, initialdir=currdir, title='Please select a directory')
if len(tempdir) > 0:
    print(f"You chose {tempdir}")
book = xlrd.open_workbook(tempdir)


stream_names = {}
column_position_name = 0
column_position_start = 0
column_position_end = 0
stream_name = ""
start = ""
end = ""

# Extract data from Excel sheets
for sheet in book.sheets():
    headers = sheet.row_values(0)
    for index, header in enumerate(headers):
        if header == "name":
            column_position_name = index
        elif header == "start":
            column_position_start = index
        elif header == "end":
            column_position_end = index

    for i in range(1, sheet.nrows):
        stream_name = sheet.row_values(i)[column_position_name]
        start = sheet.row_values(i)[column_position_start]
        end = sheet.row_values(i)[column_position_end]
        stream_names[stream_name] = [start, end]


# Automate browser navigation
eid = input("Event eid: ")
baseurl = "insert website name here" + eid
email = input("Enter your email: ")
password = input("Enter your password: ")

xpaths = { 'emailTxtBox': "insert x-path here",
           'passwordTxtBox': "insert x-path here",
           'submitButton': "insert x-path here",
           'liveButton': "insert x-path here",
           'liveStream': "insert x-path here",
           'searchBox': "insert x-path here",
           'streamResult': "insert x-path here",
           'analytics': "insert x-path here",
           'advancedSearch': "insert x-path here",
           'startBox': "insert x-path here",
           'endBox': "insert x-path here",
           'downloadButton': "insert x-path here"
         }


mydriver = webdriver.Chrome()
mydriver.get(baseurl)
mydriver.maximize_window()

# Automate log in
mydriver.find_element_by_xpath(xpaths['emailTxtBox']).clear()
mydriver.find_element_by_xpath(xpaths['emailTxtBox']).send_keys(email)
mydriver.find_element_by_xpath(xpaths['passwordTxtBox']).clear()
mydriver.find_element_by_xpath(xpaths['passwordTxtBox']).send_keys(password)
mydriver.find_element_by_xpath(xpaths['submitButton']).click()

# Navigate and click on liveStream buttons
mydriver.find_element_by_xpath(xpaths['liveButton']).click()
mydriver.find_element_by_xpath(xpaths['liveStream']).click()


# Insert excel data into the Search box
for stream_name, (start, end) in stream_names.items():
    mydriver.find_element_by_xpath(xpaths['searchBox']).click()
    mydriver.find_element_by_xpath(xpaths['searchBox']).clear()
    mydriver.find_element_by_xpath(xpaths['searchBox']).send_keys(stream_name)
    mydriver.find_element_by_xpath(xpaths['streamResult']).click()
    mydriver.find_element_by_xpath(xpaths['analytics']).click()
    mydriver.find_element_by_xpath(xpaths['advancedSearch']).click()

    start = (start, end)[0]
    end = (start, end)[1]
    y, m, d, h, _, s = xlrd.xldate_as_tuple(start, book.datemode)
    start_paste = f"{y}-{m}-{d} {h}:{s}"
    y, m, d, h, _, s = xlrd.xldate_as_tuple(end, book.datemode)
    end_paste = f"{y}-{m}-{d} {h}:{s}"

    mydriver.find_element_by_xpath(xpaths['startBox']).send_keys(start_paste)
    mydriver.find_element_by_xpath(xpaths['endBox']).send_keys(end_paste)

    # Finally download the document needed
    mydriver.find_element_by_xpath(xpaths['downloadButton']).click()
