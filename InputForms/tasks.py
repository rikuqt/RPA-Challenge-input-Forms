import time

from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables

page = browser.page()
http = HTTP()
excel = Files()
tables = Tables()


@task
def my_task():
    browser.configure(
        slowmo=3000,
    )
    open_browser()
    download_excel()
    press_start_button()
    read_csv_file()
    time.sleep(5)
    

def open_browser():
    browser.goto("https://rpachallenge.com/")

def download_excel():
    http.download("https://rpachallenge.com/assets/downloadFiles/challenge.xlsx", overwrite=True)

def read_csv_file():
    excel.open_workbook("challenge.xlsx")
    worksheet = excel.read_worksheet_as_table("Sheet1" ,header=True)
    excel.close_workbook()
    for row in worksheet:
        fill_form_with_excel_data(row)
    
def press_start_button():
    page.click("button.waves-effect.col.s12.m12.l12.btn-large.uiColorButton:has-text('Start')")
    

def fill_form_with_excel_data(person_infromation):
   
    page.fill("input[ng-reflect-name=labelFirstName]", person_infromation["First Name"])
    page.fill("input[ng-reflect-name=labelLastName]", person_infromation["Last Name"])
    page.fill("input[ng-reflect-name=labelRole]", person_infromation["Role in Company"])
    page.fill("input[ng-reflect-name=labelAddress]", person_infromation["Address"])
    page.fill("input[ng-reflect-name=labelEmail]", person_infromation["Email"])
    page.fill("input[ng-reflect-name=labelPhone]", str(person_infromation["Phone Number"]))
    page.fill("input[ng-reflect-name=labelCompanyName]", person_infromation["Company Name"])
    


    page.click("input[value=Submit]")




