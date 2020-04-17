import unittest
import testrunner
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import chromedriver_binary
import xlsxwriter 


def write_file():
    """this function is used to wrtie values to excel file"""
    workbook = xlsxwriter.Workbook('emailpass.xlsx') 
    
    # By default worksheet names in the spreadsheet will be  
    # Sheet1, Sheet2 etc., but we can also specify a name. 
    worksheet = workbook.add_worksheet("My sheet") 
    # Iterate over the data and write it out row by row. 
    worksheet.write(0,0,'treat@xyz.com') 
    worksheet.write(1,0,'12345')
    
    workbook.close() 



class logintest(unittest.TestCase):
    """ this class is used to Generate result of the login"""
   
    def setUp(self):
        """ to open browser and insert the webpage"""
        self.driver = webdriver.Chrome()
        self.driver.get("http://demo.guru99.com/test/login.html")


    
    def test_login(self):

        driver=self.driver

        import xlrd 
  
        # To open Workbook 
        wb = xlrd.open_workbook('emailpass.xlsx') 
        sheet = wb.sheet_by_index(0) 
        
        # to read values 
        user_name = sheet.cell_value(0, 0)
        password = sheet.cell_value(1, 0)

        #find and insert email and password in relevant fields
        element = driver.find_element_by_id("email")
        element.send_keys(user_name)
        element = driver.find_element_by_id("passwd")
        element.send_keys(password)
        #find and click on submit button
        login = driver.find_element_by_id("SubmitLogin")     
        login.click()

        driver.implicitly_wait(5)

    def tearDown(self):
        self.driver.quit()

if __name__== '__main__':
    write_file() # to Generate Excel File
    testrunner.main() # To generate Report
    