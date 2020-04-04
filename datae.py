
from selenium import webdriver
import re
import csv
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import datetime
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
from PIL import Image
import sys
import time
from webdriver_manager.chrome import ChromeDriverManager

def loadpage(driver):
   
    action = ActionChains(driver)
     # adding dates
    d1 = driver.find_element_by_id('openedFrom')
    driver.execute_script("arguments[0].value = arguments[1]", d1, '03/16/2020')
  
    d2 =driver.find_element_by_id('openedTo')
    driver.execute_script("arguments[0].value = arguments[1]", d2, '03/22/2020')
    # select court type
    submitbtn = driver.find_element_by_id('courTypesButton')
    submitbtn.click()
    
    #click on the select all first then choose court type
    driver.find_element_by_xpath("(//label[contains(text(),'Select all')])[1]").click()

    action.move_by_offset(20, 10)    
    action.click().perform()
    driver.find_element_by_id('courTypesButton').click()
    # select required court type
    y = driver.find_element_by_xpath("//label[contains(text(),' Circuit Civil')]")
    y.click()
    
    y.send_keys(Keys.TAB)
    time.sleep(2)
    #check if the  button is available if not quit the program
    secondbutton = driver.find_element_by_id('caseTypesButton').text
    if(secondbutton == ""):
        print("invalid selection please run the driver again  ")
        driver.quit()
    driver.find_element_by_id('caseTypesButton').click()
    #click on the select all first then choose case typehaving text REAL PROPERTY/
    ele = driver.find_element_by_xpath("(//label[contains(text(),'Select all')])[2]")
    ele.click()
    ele.send_keys(Keys.TAB)
   
    ss =driver.find_element_by_id('caseTypesButton')
    ss.click()
    s1 =driver.find_elements_by_xpath("//label[starts-with(text(),' REAL PROPERTY/')]")
    for s in s1:
        s.click()
    
    ss.send_keys(Keys.TAB)
    element = driver.find_element_by_xpath("/html/body/div[3]/table/tbody/tr/td[2]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/div/div[3]/form/img")
   
    # taking ss of captcha image and apply ocr,ocr gives incurrate results henece quitting the driver til i get correct text from captcha
    element.screenshot('myfile.png')
    image_file = Image.open("myfile.png") # open colour image
    text = pytesseract.image_to_string(image_file)
    return text
def start():
    chromedriver = "C:/webdrivers/chromedriver.exe"
    
    download_dir = "C:/Users/Tyson/Desktop/CHROME"
   
    options = webdriver.ChromeOptions()

     

    profile = {"plugins.plugins_list": [{"enabled": False,
                                         "name": "Chrome PDF Viewer"}],
               "download.default_directory": download_dir,
               "download.extensions_to_open": ""}

    options.add_experimental_option("prefs", profile)

    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=options)

    
    driver.delete_all_cookies()

    
    driver.get("https://court.martinclerk.com/Home.aspx/Search")
    driver.refresh()
    driver.maximize_window()

    prefs = {"plugins.always_open_pdf_externally": True}
    time.sleep(2)
    text = loadpage(driver)
    # ocr text replace blank spaces
    text = text.replace(" ", "")
   
    print(text)
    bol  = "no"
    if(re.match(r"^[0-9]{2}[+#@&*-¥4][0-9]{1}[=][?]", text)):#regex to validate ocr sometimes + symbol is also read as (-#@) so validating that too
        bol = "yes"
      
    if(bol== "no"):
        print("please quit")
        driver.quit()
        start()
    else:
        num = int(text[:2]) + int(text[3:4])
        print(num)
        # add ocr result in input box and click search
        driver.find_element_by_xpath("/html/body/div[3]/table/tbody/tr/td[2]/div/table/tbody/tr/td/table/tbody/tr[2]/td[2]/div/div[3]/form/input[2]").send_keys(num)
        driver.find_element_by_id('searchButton').click()
        time.sleep(2)
        row_count = len(driver.find_elements_by_xpath("//table[@id='gridSearchResults']/tbody/tr"))
        print(row_count) # getting the row count and iterate throuth all the rows
        
        invalidcaptcha = driver.find_element_by_xpath("//*[@id='title']").text
        #validation of captcha calculation
        print(invalidcaptcha)
        if(invalidcaptcha != "Case Search Results"):
            print("wrong captcha calculation")
        else:
            #creating list of all rows
            rowout = []
            for h in range(1,row_count+1):#iterate throuth all the rows to get required data
                st = str(h)
                #wait till element loads
                WebDriverWait(driver, 10).until(lambda x: x.find_elements_by_class_name("searchSection"))
                WebDriverWait(driver, 10).until(lambda x: x.find_elements_by_xpath("//*[@id='gridSearchResults']/tbody/tr["+st+"]/td[3]"))
                driver.find_element_by_xpath("//*[@id='gridSearchResults']/tbody/tr["+st+"]/td[3]").click()#wait till the element appears 
                time.sleep(2)
               
                print(str(h)+"done")
                time.sleep(2)
                WebDriverWait(driver, 10).until(
                lambda x: x.find_element_by_class_name("detailsSummary"))
                #wait till element loads
                banknameandcaseid= driver.find_element_by_xpath("//*[@id='page_header']/div/div[1]/div").text
                case_date = driver.find_element_by_xpath("//*[@id='summaryAccordionCollapse']/table/tbody/tr/td[1]/dl/dd[3]").text
                case_type = driver.find_element_by_xpath("//*[@id='summaryAccordionCollapse']/table/tbody/tr/td[3]/dl/dd[1]").text
                defendent = driver.find_element_by_xpath("//td[contains(text(),'DEFENDANT')]/following::td").text
                WebDriverWait(driver, 10).until(lambda x: x.find_element_by_xpath("//td[contains(text(),'VALUE OF REAL PROPERTY OR MORTGAGE FORECLOSURE CLAIM:')]/preceding-sibling::td[2]"))
                driver.find_element_by_xpath("//td[contains(text(),'VALUE OF REAL PROPERTY OR MORTGAGE FORECLOSURE CLAIM:')]/preceding-sibling::td[2]").click()
                print(case_date,banknameandcaseid[:13],case_type,defendent,banknameandcaseid[16:])
                # adding one row to list
                row = [case_date,banknameandcaseid[:13],case_type,defendent,banknameandcaseid[16:]]
                # writer.writerow(row)
                rowout.append(row)
                driver.back()
                time.sleep(2)
            print(rowout)    
            filename = "sample.csv"
            # empty the sample.csv file and the write in file
            f = open("sample.csv", "w")
            f.truncate()
            f.close()
            fields = ['CaseDate','CaseID','CaseType ','DefendantName','BankName'] 
# writing to csv file 
            with open(filename, 'w') as csvfile: 
    # creating a csv writer object 
                csvwriter = csv.writer(csvfile) 
                
                # writing the fields 
                csvwriter.writerow(fields) 
                
                # writing the data rows 
                csvwriter.writerows(rowout)
        
     


start() 
#the program quits itself until it finds correct captch with ocr
# ocr is not able to read the pdfs as every file every different case has p different format and ocr results are very inaccurate

    
