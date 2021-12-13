from os import name
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import time, os


class Automation:
    def __init__(self, url, name_agencie): 
        os.mkdir('output')
        self.browser_lib = Selenium()
        self.file = Files()
        self.browser_lib.set_download_directory(os.path.join(os.getcwd(), f"output/"))
        self.url = url
        self.name_agencie = name_agencie
        self.agencies = []
        self.values = []
        self.links = []
        self.text = ""
        

        
        

    def open_the_website(self):
        try:
            self.browser_lib.open_available_browser(self.url)
            self.browser_lib.click_element_when_visible('xpath://*[@id="node-23"]/div/div/div/div/div/div/div/a')
            time.sleep(2)
            self.text = self.browser_lib.find_element('xpath://*[@id="agency-tiles-widget"]/div  ')
            self.browser_lib.wait_until_element_is_visible('xpath://*[@id="agency-tiles-widget"]/div  ')
            self.text = self.text.text
        except(AssertionError):
            print("Invalid link")
    
    def write_excel(self):
        self.text = self.text.replace("Total FY2021 Spending:","")
        self.text = self.text.replace("view","")
        self.text = self.text.split("\n")
        
        self.file.create_workbook("output/agencies.xlsx")
        self.file.create_worksheet("agencies")   
        self.file.set_active_worksheet("agencies")
        for n in self.text:
            if len(n) == 5 or len(n) == 6 or len(n) == 4:
                self.values.append(n)
            elif len(n) == 0:
                continue
            else:
                self.agencies.append(n)
        
        row = 1
        try:
            for n in self.agencies:
                self.file.set_cell_value(row, 1, n)
                row+=1
            row = 1
            for n in self.values:
                self.file.set_cell_value(row, 2, n)
                row+=1
        except(PermissionError):
            print("Close the excel file to continue")
        self.file.save_workbook()
        self.file.close_workbook()


    def click_agencie(self):
        self.agencies_id = {}
        a=0
        while a < len(self.agencies):
            for n1 in range(1,len(self.agencies)):
                for n2 in range(1,4):
                    if a ==len(self.agencies):
                        break
                    else:
                        self.agencies_id[self.agencies[a]] = (n1,n2)
                        a=a+1
        num = self.agencies_id[self.name_agencie]
        self.browser_lib.wait_until_element_is_visible('xpath://*[@id="agency-tiles-widget"]/div/div[{}]/div[{}]'.format(num[0],num[1]))
        self.browser_lib.click_element_when_visible('xpath://*[@id="agency-tiles-widget"]/div/div[{}]/div[{}]'.format(num[0],num[1]))
    
    def table(self):
        self.browser_lib.set_browser_implicit_wait(10)
        self.browser_lib.click_element_when_visible('xpath://*[@id="investments-table-object_length"]/label/select')
        time.sleep(2)
        self.browser_lib.set_browser_implicit_wait(10)
        self.browser_lib.click_element_when_visible('xpath://*[@id="investments-table-object_length"]/label/select/option[4]')
        self.browser_lib.set_browser_implicit_wait(15)
        time.sleep(2)
        self.browser_lib.set_browser_implicit_wait(10)
        self.file.open_workbook("output/agencies.xlsx")
        try:
            self.file.create_worksheet("Individual Investments")
        except(ValueError):
            self.file.set_active_worksheet("Individual Investments")
        row = 1
        try:
            for n in range(1,8):
                celd = self.browser_lib.find_element('xpath://*[@id="investments-table-object_wrapper"]/div[3]/div[1]/div/table/thead/tr[2]/th[{}]'.format(n))
                celd = celd.text
                self.file.set_cell_value(1, row, celd)
                row+=1
        except(PermissionError):
            print("Close the excel file to continue")
        entries = self.browser_lib.find_element('xpath://*[@id="investments-table-object_info"]')
        entries= entries.text
        entries = entries.split(" ")
        entries = entries[5]
        row = 2
        try:
            for col in range(1,8):
                row=2
                for n in range(1,int(entries)+1):
                    val = self.browser_lib.find_element('xpath://*[@id="investments-table-object"]/tbody/tr[{}]/td[{}]'.format(n,col))
                    val = val.text
                    self.file.set_cell_value(row, col, val)
                    row+=1
        except(PermissionError):
            print("Close the excel file to continue")
        try:
            self.file.save_workbook()
            self.file.close_workbook()
        except(PermissionError):
            print("Close the excel file to continue")

    def donwload_pdf_if_exists(self):
        
        self.browser_lib.set_browser_implicit_wait(10)
        self.browser_lib.click_element_when_visible('xpath://*[@id="investments-table-object_length"]/label/select')
        time.sleep(2)
        self.browser_lib.set_browser_implicit_wait(10)
        self.browser_lib.click_element_when_visible('xpath://*[@id="investments-table-object_length"]/label/select/option[4]')
        self.browser_lib.set_browser_implicit_wait(15)
        time.sleep(2)

        data = self.browser_lib.find_elements('//*[@id="investments-table-object"]/tbody/tr/td[1]/a')
        for n in data:
            self.links.append(n.get_attribute("href"))

        for link in self.links:
            self.browser_lib.go_to(link)
            self.browser_lib.click_element_when_visible('//*[@id="business-case-pdf"]/a')

            while True:
                time.sleep(2)
                try:
                    if self.browser_lib.find_element('//div[@id="business-case-pdf"]').find_element_by_tag_name("span"):
                        time.sleep(2)
                    else:
                        break
                except:
                    if self.browser_lib.find_element('//*[contains(@id,"business-case-pdf")]//a[@aria-busy="false"]'):
                        time.sleep(2)
                        break


if __name__ == "__main__":
    link = Automation("https://itdashboard.gov", "National Archives and Records Administration")
    link.open_the_website()
    link.write_excel()
    link.click_agencie()
    link.table()
    link.donwload_pdf_if_exists()

