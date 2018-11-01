from unittest import defaultTestLoader as loader, TextTestRunner
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.utils import keys_to_typing
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from pkg.BasePython import BasePython
from Configs import Configs
import unittest
import datetime
import argparse
import requests
import logging
import json
import time
import sys
import os


v_low = 1
low = 2
mid = 4
high = 8
v_high = 10
data ={}


class CustomResult(unittest.TestResult):

    def addFailure(self, test, err):
        critical = ['test_createAccount', 'test_loginAndUpdate', 'test_loginTransIpMail', 'test_preferencesUpdation', 'test_serverConfiguration', 'test_xEmailsCleanup']
        if test._testMethodName in critical:
            print("Critical Failure!")
            self.stop()
        unittest.TestResult.addFailure(self, test, err)


class TestSuite1(BasePython):
    
    def updateStatus(self, step_number, step_name, start_timestamp, end_timestamp, step_status, message, exception_status, build_number, consoleUrl):
        
        self.consoleUrl = consoleUrl
        self.build_number = build_number
        rollbase_xml_payload=('C:/robotic_process_automation/work/rollbase_xml_payload_{}.xml'.format(self.build_number))
        print()
        self.step_number = step_number
        self.step_name = step_name
        self.start_timestamp = start_timestamp
        self.end_timestamp = end_timestamp
        self.step_status = step_status
        self.message = message
        self.exception_status = exception_status
        print("Step Name : " + self.step_name)
        print("Start timestamp : " + self.start_timestamp)
        print("End timestamp : " + self.end_timestamp)
        print("Status : " + self.step_status)
        if (self.step_status=="Failed"):
            print("Message : " + self.message)
            self.driver.quit()
        
        if (self.step_number == 1):
            if os.path.exists(rollbase_xml_payload):
                os.remove(rollbase_xml_payload)
            
            time.sleep(v_high)
            xml_req= open(rollbase_xml_payload,"a+")
            xml_req.write('<?xml version="1.0" encoding="utf-8" ?>')
            xml_req.write('<data objName="Robot_Job" useIds="false">')
            xml_req.write('<Field name="R32990512">32990533</Field>')
            xml_req.write('<Field name="R32990522">32990091</Field>')
            xml_req.write('<Field name="Startdate_and_time">{}</Field>'.format(self.start_timestamp))
            xml_req.write('<Field name="Step_1">{}</Field>'.format(self.step_status))

            if (self.step_status == 'False'):
                xml_req.write('<Field name="Step_7">{}</Field>'.format(self.consoleUrl))
                xml_req.write('<Field name="enddate_and_time">{}</Field>'.format(self.end_timestamp))
                xml_req.write('<Field name="status">status4</Field>')
                xml_req.write('<Field name="end_result">Error: Step 1, Element could not found.</Field></data>')
                xml_req.close()
                
            
        if (self.step_number >= 2 and self.step_number <= 5):
            xml_req= open(rollbase_xml_payload,"a+")
            xml_req.write('<Field name="Step_{}">{}</Field>'.format(self.step_number,self.step_status))

            if (self.step_status == 'False'):
                xml_req.write('<Field name="Step_7">{}</Field>'.format(self.consoleUrl))
                xml_req.write('<Field name="enddate_and_time">{}</Field>'.format(self.end_timestamp))
                xml_req.write('<Field name="status">status4</Field>')
                xml_req.write('<Field name="end_result">Error: Step_{}, Element could not found.</Field></data>'.format(self.step_number))
                xml_req.close()
            

        if (self.step_number == 6):
            xml_req= open(rollbase_xml_payload,"a+")
            xml_req.write('<Field name="Step_6">{}</Field>'.format(self.step_status))
            
            if (self.step_status == 'False'):
                xml_req.write('<Field name="Step_7">{}</Field>'.format(self.consoleUrl))
                xml_req.write('<Field name="Enddate_and_time">{}</Field>'.format(self.end_timestamp))
                xml_req.write('<Field name="status">status4</Field>')
                xml_req.write('<Field name="End_result">Error: Step_6, Element could not found.</Field></data>')
                xml_req.close()
                
            else:
                xml_req.write('<Field name="Step_7">{}</Field>'.format(self.consoleUrl))
                xml_req.write('<Field name="Enddate_and_time">{}</Field>'.format(self.end_timestamp))
                xml_req.write('<Field name="status">status3</Field>')
                xml_req.write('<Field name="End_result">All steps are completed successfully!</Field></data>')
                xml_req.close()


    def test_createAccount(self):

        step_num = 1
        step = "User Account Creation"
        start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            for arg in sys.argv[1:]:
                text = arg.split("=")
                data[text[0]] = text[1]
            driver = self.driver
            driver.implicitly_wait(6)
            driver.get(Configs.baseUrl)
            driver.implicitly_wait(6)
            build_num = str(data["build_num"])
            print("Jenkins Build number is: " + build_num)
            consoleUrl=('https://rpa.cloudapps.services/job/RPA_Cot_Bot/{}/console'.format(build_num))
            companyName = driver.find_element(By.ID, "companyName")
            companyName.send_keys(data["companyname"])                              #inserting company name
            time.sleep(4)

            firstName = driver.find_element(By.ID, "firstName")
            firstName.send_keys(data["firstname"])                                  #inserting first name of user     
            firstName.send_keys(Keys.RETURN)

            lastName = driver.find_element(By.ID, "lastName")
            lastName.send_keys(data["lastname"])                                    #inserting last name     

            email = driver.find_element(By.ID, "email")
            email.send_keys("rpa@brixxs.nl")                                        #inserting email address    

            phone = driver.find_element(By.ID, "phone")
            phone.send_keys(data["mobile"])                                         #inserting mobile number     

            loginName = driver.find_element(By.ID, "loginName")
            loginName.clear()
            loginName.send_keys(data["loginname"])

            submit = driver.find_element(By.XPATH, "//form[@name='theForm']//div[@name='Submission Form End']//td[@class='center']/input[@value=' Verzenden ']")
            submit.click()
            time.sleep(5)
        except:
            msg = "Timeout Error!! Element not found."
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "false"
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=True, build_number=build_num, consoleUrl=consoleUrl)
            print("Error! Element could not found.")
            self.fail()
        else:
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "true"
            msg = ""
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=False, build_number=build_num, consoleUrl=consoleUrl)

    
    def test_loginAndUpdate(self):

        step_num = 2
        step = "Login and Update"
        start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            for arg in sys.argv[1:]:
                text = arg.split("=")
                data[text[0]] = text[1]
            driver = self.driver
            driver.implicitly_wait(6)
            driver.get(Configs.secondUrl)
            driver.implicitly_wait(6)
            build_num = str(data["build_num"])
            consoleUrl=('https://rpa.cloudapps.services/job/RPA_Cot_Bot/{}/console'.format(build_num))
            loginName2 = driver.find_element(By.ID, "loginName")
            loginName2.send_keys("Officeportaal")                                                                                   #insert username
            password = driver.find_element(By.ID, "password")
            password.send_keys("Officeportaal!")                                                                                   #insert password
            submit = driver.find_element(By.XPATH, "//button[@id='btnLogin']")
            submit.click()                                                                                                          #submit creadentials
            time.sleep(v_high)
            driver.find_element_by_link_text(data["companyname"]).click()                                                           #company name found                  
            driver.implicitly_wait(v_high)
            driver.find_element_by_link_text("Bewerken").click()                                                                    #edit company information
            driver.implicitly_wait(high)
            expiredAt = driver.find_element_by_name("expiredAt")                                                                    #remove the date
            expiredAt.clear()
            driver.implicitly_wait(high)
            driver.find_element(By.XPATH, "/html[1]/body[1]/div[3]/div[2]/div[1]/form[1]/section[1]/ul[1]/li[1]/div[1]/div[1]/div[1]/div[1]/div[11]/div[2]/div[1]/div[3]/div[1]/span[1]/span[1]/span[1]").click()
            driver.implicitly_wait(high)    
            office = driver.find_element(By.XPATH, "//ul[@id='R5683293_listbox']/li[5]")
            time.sleep(v_high)
            office.click()                                                                                                          #select officeportaal
            time.sleep(high)
            driver.find_element(By.XPATH, "/html[1]/body[1]/div[3]/div[2]/div[1]/form[1]/section[1]/ul[1]/li[1]/div[1]/div[1]/div[1]/div[1]/div[12]/div[2]/div[1]/div[3]/div[1]/button[1]/span[1]").click()
            driver.implicitly_wait(high)
            driver.switch_to.frame(0)                                                                                               #switching to iframe
            driver.implicitly_wait(high)
            box = driver.find_element(By.XPATH, "//div[@id='rb-content-box']/section[@name='Selecteer Gebruikers']/ul[@role='menu']/li[@role='menuitem']/div[@role='region']//div[@class='row']/div/div[2]/div[2]/table[@role='grid']//input[@name='selFromDlg_28958816']")
            time.sleep(5)
            box.click()
            saver = driver.find_element(By.XPATH, "//div[@id='rb-content-box']/section[@name='Selecteer Gebruikers']/ul[@role='menu']/li[@role='menuitem']/div[@role='region']//div[@class='gridToolbar k-toolbar k-widget']/a[1]")
            saver.click()                                                                                                           #save changes
            driver.implicitly_wait(high)
            driver.switch_to.default_content()                                                                                      #switch to default frame
            driver.implicitly_wait(high)
            driver.find_element(By.XPATH, "/html//div[@id='rbi_F_maxUsers']//span[@class='k-numeric-wrap k-state-default']/input[1]").send_keys("00")  #update maximum users
            driver.implicitly_wait(high)
            mailSender = driver.find_element(By.ID, "mailSender")
            mailSender.clear()                                                                                                      #clear mail sender
            driver.implicitly_wait(high)
            mailSender.send_keys("noreply@apptomorrow.nl")                                                                          #enter new mail id
            driver.implicitly_wait(high)
            driver.find_element(By.XPATH, "//div[@id='rb-header-box']//a[@title='Opslaan']/span").click()                           #save changes
            driver.implicitly_wait(high)
        except:
            msg = "Timeout Error!! Element not found."
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "false"
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=True, build_number=build_num, consoleUrl=consoleUrl)
            print("Error! Element could not found.")
            self.fail()
        else:
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "true"
            msg = ""
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=False, build_number=build_num, consoleUrl=consoleUrl)


    def test_loginTransIpMail(self):

        step_num = 3
        step = "Login Transip Mail"
        start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            for arg in sys.argv[1:]:
                text = arg.split("=")
                data[text[0]] = text[1]
            driver = self.driver
            time.sleep(mid)
            driver.get(Configs.mailUrl)                                                                                                     #getting transip mail url
            driver.implicitly_wait(high)
            build_num = str(data["build_num"])
            consoleUrl = ('https://rpa.cloudapps.services/job/RPA_Cot_Bot/{}/console'.format(build_num))
            mailuser = driver.find_element(By.ID, "rcmloginuser_tip")
            mailuser.send_keys("rpa@brixxs.nl")                                                                                     #insert username
            mailpass = driver.find_element(By.ID, "rcmloginpwd_tip")
            mailpass.send_keys("Qwerty123!")                                                                                        #insert password
            driver.find_element(By.XPATH, "/html//button[@id='btnId']").click()                                                     #submit credentials
            time.sleep(v_high)
            mail = driver.find_element_by_link_text("Welkom bij " + data["companyname"])
            mail.click()                                                                                                            #select mail to set password
            meer = driver.find_element(By.XPATH, "/html//a[@id='messagemenulink']").click()                                         #select action menu to open mail
            time.sleep(mid)
            newtab = driver.find_element(By.XPATH, "//ul[@id='messagemenu-menu']/li[7]/a[@role='button']/span[.='Openen in een nieuw venster']")
            newtab.click()                                                                                                          #mail opened in new tab
            time.sleep(mid)
            driver.switch_to.window(driver.window_handles[1])                                                                       #switch to second tab
            time.sleep(low)
            driver.find_element_by_link_text("Instellen wachtwoord").click()                                                        #set new password
            time.sleep(low)
            driver.switch_to.window(driver.window_handles[2])                                                                       #switch to third tab
            time.sleep(mid)

            try:
                driver.find_element_by_link_text(", klik dan hier")
                time.sleep(mid)
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(mid)
                dropdown = driver.find_element(By.XPATH, "/html//select[@id='messagessearchfilter']")
                dropdown.click()
                time.sleep(mid)
                unseen = driver.find_element(By.XPATH, "//select[@id='messagessearchfilter']/option[@value='UNSEEN']")
                unseen.click()
                time.sleep(mid)
                driver.find_element_by_link_text("Welkom bij " + data["companyname"]).click()
                meer = driver.find_element(By.XPATH, "/html//a[@id='messagemenulink']").click()                                         #select action menu to open mail
                time.sleep(mid)
                newtab = driver.find_element(By.XPATH, "//ul[@id='messagemenu-menu']/li[7]/a[@role='button']/span[.='Openen in een nieuw venster']")
                newtab.click()                                                                                                          #mail opened in new tab
                driver.switch_to.window(driver.window_handles[3])                                                                       #switch to second tab
                time.sleep(high)
                driver.find_element_by_link_text("Instellen wachtwoord").click()                                                        #set new password
                time.sleep(high)
                driver.switch_to.window(driver.window_handles[4])
                time.sleep(mid)
                loginname2 = driver.find_element(By.ID, "loginName")
                loginname2.send_keys(data["loginname"])                                                                             #enter the username
                time.sleep(low)
                submit2 = driver.find_element(By.XPATH, "//button[@id='btnLogin']")
                submit2.click()                                                                                                     #submit username
                time.sleep(mid)

            except:
                time.sleep(mid)
                loginname2 = driver.find_element(By.ID, "loginName")
                loginname2.send_keys(data["loginname"])                                                                             #enter the username
                time.sleep(low)
                submit2 = driver.find_element(By.XPATH, "//button[@id='btnLogin']")
                submit2.click()                                                                                                     #submit username
                time.sleep(mid)

            time.sleep(mid)
            newPassword1 = driver.find_element(By.ID, "newPassword1_field")
            newPassword1.send_keys("123456")                                                                                        #enter new password
            conformPassword = driver.find_element(By.ID, "newPassword2_field")
            conformPassword.send_keys("123456")                                                                                     #confirm new password
            driver.find_element(By.XPATH, "//button[@id='submitBtn']").click()                                                      #password created successfully
            time.sleep(low)
            
        except:
            msg = "Timeout Error!! Element not found."
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "false"
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=True, build_number=build_num, consoleUrl=consoleUrl)
            print("Error! Element could not found.")
            self.fail()
        else:
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "true"
            msg = ""
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=False, build_number=build_num, consoleUrl=consoleUrl)



    def test_preferencesUpdation(self):                                                                                         #this method login with new user

        step_num = 4
        step = "Preferences Updation"
        start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            for arg in sys.argv[1:]:
                text = arg.split("=")
                data[text[0]] = text[1]
            
            driver = self.driver
            driver.get(Configs.brixxsURL)                                                                                            #getting test.brixxs url
            build_num = str(data["build_num"])
            consoleUrl=('https://rpa.cloudapps.services/job/RPA_Cot_Bot/{}/console'.format(build_num))
            loginName3 = driver.find_element(By.ID, "loginName")
            loginName3.send_keys(data["loginname"])                                                                          #enter new username
            password3 = driver.find_element(By.ID, "password")
            password3.send_keys("123456")                                                                                    #enter password
            submit3 = driver.find_element(By.XPATH, "//button[@id='btnLogin']")
            submit3.click()                                                                                                  #logged in successfully
            time.sleep(mid)

            dropdown = driver.find_element(By.XPATH, "/html[1]/body[1]/div[1]/div[3]/div[2]/ul[2]/li[1]/a[1]/span[2]")
            dropdown.click()                                                                                                #clicked on dropdown button
            time.sleep(low)
            driver.find_element_by_link_text("Inschakelen support toegang").click()                                         #clicked on enable support access
            time.sleep(low)
            driver.find_element(By.XPATH, "//body[@class='pacific-bootstrap']/table[@class='rbs_mainContentTable']/tbody/tr[2]/td[3]/table//table[@class='rbs_mainComponentTable']/tbody/tr[1]/td/div[2]/input[@value=' Bewerken ']").click()
            time.sleep(low)
            access = driver.find_element(By.ID, "supportLoginEnabledTime")
            access.clear()                                                                                                  #access users cleared
            time.sleep(low)
            access.send_keys("30")                                                                                          #set access user=30
            time.sleep(low)
            unit = driver.find_element(By.ID, "supportLoginEnabledTimeUnit")                                                #found the time unit tab
            unit.click()
            unit.send_keys("dagen")                                                                                         #set support login enable time unit = dagen
            driver.find_element(By.XPATH, "//body[@class='pacific-bootstrap']/table[@class='rbs_mainContentTable']/tbody/tr[2]/td[3]/table//form[@name='theForm']/table[@class='rbs_mainComponentTable']/tbody/tr[1]/td/div[2]/input[@value=' Opslaan ']").click()
            time.sleep(high)

            back = driver.find_element_by_link_text("<< terug naar OfficePortaal Algemeen")
            back.click()                                                                                                    #go back to home page
            time.sleep(mid)
            
            dropdown2 = driver.find_element(By.XPATH, "/html[1]/body[1]/div[1]/div[3]/div[2]/ul[2]/li[1]/a[1]/span[2]")
            dropdown2.click()                                                                                               #clicked dropdown button
            time.sleep(low)
            medewerkers = driver.find_element_by_link_text("Medewerkers")
            medewerkers.click()                                                                                             #edit employees permissions
            time.sleep(low)

            klant = driver.find_element_by_link_text(data["firstname"] + " " + data["lastname"])
            klant.click()                                                                                                   #client name found
            time.sleep(low)
            Bewerken2 = driver.find_element_by_link_text("Bewerken")
            Bewerken2.click()                                                                                               #edit client
            time.sleep(mid)
            search2 = driver.find_element(By.XPATH, "//div[@id='rbi_F_locationId']/button[@role='button']")
            search2.click()
            time.sleep(mid)
            driver.switch_to.frame(0)                                                                                       #switch to iframe0
            time.sleep(mid)
            alltext1 = driver.find_element_by_link_text("All")
            alltext1.click()                                                                                                #selected all
            time.sleep(mid)
            driver.switch_to.default_content()                                                                              #switched to default frame
            time.sleep(low)
            search3 = driver.find_element(By.XPATH, "//div[@id='rbi_F_functionId']/button[@role='button']")
            search3.click()
            time.sleep(mid)
            driver.switch_to.frame(1)                                                                                       #switch to iframe1
            time.sleep(mid)
            alltext2 = driver.find_element_by_link_text("All").click()                                                      #selected all
            time.sleep(low)
            driver.switch_to.default_content()                                                                              #switched to default frame
            time.sleep(low)
            search4 = driver.find_element(By.XPATH, "//div[@id='rbi_F_departmentId']/button[@role='button']").click()
            time.sleep(mid)
            driver.switch_to.frame(2)                                                                                       #switch to iframe2
            time.sleep(low)
            alltext3 = driver.find_element_by_link_text("All").click()                                                      #selected all
            time.sleep(low)
            driver.switch_to.default_content()                                                                              #switched to default frame
            time.sleep(low)
            
            dropdown3 = driver.find_element(By.XPATH, "/html[1]/body[1]/div[3]/div[2]/div[1]/form[1]/section[7]/ul[1]/li[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/span[1]/span[1]/span[1]")
            dropdown3.click()
            time.sleep(low)
            utf8 = driver.find_element(By.XPATH, "//ul[@id='emailEncoding_listbox']/li[1]")
            utf8.click()                                                                                                    #select utf8 incoding method
            time.sleep(low)
            save4 = driver.find_element(By.XPATH, "//div[@id='rb-header-box']//a[@title='Opslaan']/span")
            save4.click()                                                                                                   #saved changes
            time.sleep(low)
            driver.quit()
        except:
            msg = "Timeout Error!! Element not found."
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "false"
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=True, build_number=build_num, consoleUrl=consoleUrl)
            print("Error! Element could not found.")
            self.fail()
        else:
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "true"
            msg = ""
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=False, build_number=build_num, consoleUrl=consoleUrl)



    def test_serverConfiguration(self):

        step_num = 5
        step = "Server and Email Configuration"
        start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            for arg in sys.argv[1:]:
                text = arg.split("=")
                data[text[0]] = text[1]
            self.driver = webdriver.Chrome(executable_path=r'C:/robotic_process_automation/lib/chromedriver.exe')
            self.driver.maximize_window()
            self.driver.implicitly_wait(2)
            driver = self.driver
            driver.get(Configs.brixxsURL)                                                                                             #getting test.brixxs url
            time.sleep(v_low)
            build_num = str(data["build_num"])
            consoleUrl=('https://rpa.cloudapps.services/job/RPA_Cot_Bot/{}/console'.format(build_num))
            loginName3 = driver.find_element(By.ID, "loginName")
            loginName3.send_keys(data["loginname"])                                                                           #enter new username
            password3 = driver.find_element(By.ID, "password")
            password3.send_keys("123456")
            submit3 = driver.find_element(By.XPATH, "//button[@id='btnLogin']")
            submit3.click()                                                                                                   #logged in successfully
            time.sleep(high)
            grid = driver.find_element(By.XPATH, "//div[@id='rb-header-box']/div[2]/ul[2]/li/span/span").click()              #click on grid
            time.sleep(high)
            driver.find_element_by_link_text("Startpagina configuratie").click()                                              #configuration
            time.sleep(high)
            dropdown4 = driver.find_element(By.XPATH, "//div[@id='nonPacificHeaderDivId']/table[2]/tbody/tr/td[3]/table[@class='wide']/tbody/tr/td[3]//button[@class='btn btn-small dropdown-toggle']/span[@class='caret']").click()
            time.sleep(high)
            driver.find_element_by_link_text(" Beheerconfiguratie").click()                                                   #edit configuration
            time.sleep(high)
            preferences = driver.find_element_by_css_selector("[href='preferences\.jsp\?domain\=custPref'] strong").click()   #edit prefrences
            time.sleep(high)
            checkbox2 = driver.find_element(By.XPATH, "//input[@name='turnOffWelcomeMail']").click()                        #checkbox enable, It should be unchecked before running the code
            time.sleep(high)
            opslaan2 = driver.find_element(By.XPATH, "//body[@class='pacific-bootstrap']//table[@class='rbs_mainContentTable']//tbody//tr[@height='100%']//td[@class='wide top']//table[@class='wide top']//tbody//tr//td//form[@name='theForm']//table[@class='rbs_mainComponentTable']//tbody//tr//td//div[2]//input[1]").click()  #save changes
            time.sleep(high)
            dropdown5 = driver.find_element(By.XPATH, "//div[@id='nonPacificHeaderDivId']/table[2]/tbody/tr/td[3]/table[@class='wide']/tbody/tr/td[3]//button[@class='btn btn-small dropdown-toggle']/span[@class='caret']").click()
            time.sleep(high)
            driver.find_element_by_link_text(" Beheerconfiguratie").click()                                                 #again click on configuration
            time.sleep(high)
            emailserversettings = driver.find_element_by_css_selector("[href='custEmailSettings\.jsp'] strong").click()     #edit email server settings
            time.sleep(high)
            radiobtn = driver.find_element(By.XPATH, "/html//input[@id='emailServerTypeSMTP']").click()                     #click on smtp radio button
            driver.implicitly_wait(high)
            mailhost = driver.find_element(By.XPATH, "//td[@id='MailHost_field']/input[@name='MailHost']")
            mailhost.send_keys("smtp.officeportaal.nl")                                                                     #enter mail host

            mailuser = driver.find_element(By.XPATH, "//td[@id='MailUser_field']/input[@name='MailUser']") 
            mailuser.send_keys("noreply@officeportaal.nl")                                                                  #enter mail user
            
            paswrd = driver.find_element(By.XPATH, "//div[@id='EmailSettings_Table']/table[1]/tbody/tr[5]/td[@class='rbs_leftDataColWide']/input[@name='MailPassword']")
            paswrd.send_keys("Tomorrow1000!")                                                                               #enter password
            
            email2 = driver.find_element(By.XPATH, "/html//td[@id='email_field']/input[@name='email']")
            email2.send_keys("rpa@brixxs.nl")                                                                               #enter receiver email
            
            automail = driver.find_element(By.XPATH, "//td[@id='AutoReplyAddress_field']/input[@name='AutoReplyAddress']")
            automail.send_keys("noreply@officeportaal.nl")                                                                  #enter automail
            time.sleep(high)
            parentHandle = driver.current_window_handle                                                                     #points to parent handle
            driver.find_element(By.XPATH, "//div[@id='EmailSettings_Table']/table[2]/tbody/tr[3]/td[@class='rbs_leftDataColWide']/input[@name='testButton']").click()
            time.sleep(high)                                                                                                #test connection
            handles = driver.window_handles                                                                                 # it gives all the handles which are currently open
            for handle in handles:
                if handle not in parentHandle:
                    driver.switch_to.window(handle)                                                                         #switch to new popup window
                    time.sleep(mid)
                    con = driver.find_element(By.XPATH, "//div[@id='rb-styleable-content-box']/input[@value='Bericht is ontvangen in mailbox']")
                    time.sleep(mid)
                    con.click()
                    time.sleep(low)
                    driver.close                                                                                            #close only current window
                    break

            driver.switch_to.window(parentHandle)                                                                           #switched back to the parent handle
            time.sleep(low)
            opslaan5 = driver.find_element(By.XPATH, "//body[@class='pacific-bootstrap']/table[@class='rbs_mainContentTable']/tbody/tr[2]/td[3]/table//form[@name='theForm']/table[@class='rbs_mainComponentTable']/tbody/tr[1]/td/div[2]/input[@value=' Opslaan ']")
            opslaan5.click()
            driver.implicitly_wait(v_high)
            driver.quit()
        except:
            msg = "Timeout Error!! Element not found."
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "false"
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=True, build_number=build_num, consoleUrl=consoleUrl)
            print("Error! Element could not found.")
            self.fail()
        else:
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "true"
            msg = ""
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=False, build_number=build_num, consoleUrl=consoleUrl)



    def test_xEmailsCleanup(self):
        
        step_num = 6
        step = "Emails CleanUp"
        start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            for arg in sys.argv[1:]:
                text = arg.split("=")
                data[text[0]] = text[1]
            self.driver = webdriver.Chrome(executable_path=r'C:/robotic_process_automation/lib/chromedriver.exe')
            self.driver.maximize_window()
            self.driver.implicitly_wait(2)
            driver = self.driver
            driver.get(Configs.mailUrl)                                                                                                     #getting transip mail url
            driver.implicitly_wait(high)
            build_num = str(data["build_num"])
            consoleUrl=('https://rpa.cloudapps.services/job/RPA_Cot_Bot/{}/console'.format(build_num))
            mailuser = driver.find_element(By.ID, "rcmloginuser_tip")
            mailuser.send_keys("rpa@brixxs.nl")                                                                                     #insert username
            mailpass = driver.find_element(By.ID, "rcmloginpwd_tip")
            mailpass.send_keys("Qwerty123!")                                                                                        #insert password
            driver.find_element(By.XPATH, "/html//button[@id='btnId']").click()                                                     #submit credentials
            time.sleep(mid)
            mail = driver.find_element_by_link_text("Welkom bij " + data["companyname"])
            mail.click()                                                                                                            #select mail to set password
            time.sleep(mid)
            delete = driver.find_element(By.XPATH, "//div[@id='messagetoolbar']/a[4]")                                              #but not required because new mails displays always on top
            delete.click()
            time.sleep(mid)
            mail = driver.find_element_by_link_text("Welkom bij " + data["companyname"])
            mail.click()                                                                                                            #select mail to set password
            time.sleep(mid)
            delete = driver.find_element(By.XPATH, "//div[@id='messagetoolbar']/a[4]")                                              #but not required because new mails displays always on top
            delete.click()
            time.sleep(mid)
            
            mail = driver.find_element_by_link_text("Wachtwoord succesvol gewijzigd")
            mail.click()                                                                                                            #select mail to set password
            time.sleep(mid)
            delete = driver.find_element(By.XPATH, "//div[@id='messagetoolbar']/a[4]")                                              #but not required because new mails displays always on top
            delete.click()
            time.sleep(mid)
            driver.find_element_by_link_text("Connection Test").click()
            time.sleep(low)
            delete = driver.find_element(By.XPATH, "//div[@id='messagetoolbar']/a[4]")                                              #but not required because new mails displays always on top
            delete.click()
            time.sleep(mid)
            driver.quit()
        except:
            msg = "Timeout Error!! Element not found."
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "false"
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=True, build_number=build_num, consoleUrl=consoleUrl)
            print("Error! Element could not found.")
            self.fail()
        else:
            end = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "true"
            msg = ""
            self.updateStatus(step_number=step_num, step_name=step, start_timestamp=start, end_timestamp=end, step_status=status, message=msg, exception_status=False, build_number=build_num, consoleUrl=consoleUrl)


    def test_zRollbase_Update(self):                                                                                        # method used to update rollbase

        build_num = str(data["build_num"])                                                                          # stored jenkins build number in a variable
        rollbase_xml_payload=('C:/robotic_process_automation/work/rollbase_xml_payload_{}.xml'.format(build_num))           # creates a file with the build nuber to this path
        print("\nXML File Created Successfully!\n")
        url = "https://cloudapps.services/rest/api/login?output=json"

        headers = {
            'Content-Type': "application/json",
            'Acc': "application/json",
            'loginName': "restuser01",
            'password': "Qwerty123",
            'custId': "16384948",
            'cache-control': "no-cache",
            }

        response = requests.request("POST", url, headers=headers)                                                          # login api called

        print(response.json())
        json_response=response.json()
        ATOKEN=json_response["sessionId"]

        #Create API call

        url = "https://cloudapps.services/rest/api/create"

        querystring = {"output":"json","useIds":"false","sessionId":ATOKEN}

        with open(rollbase_xml_payload, "r+") as myfile:
            payload = myfile.read()
        print()
        print(payload)

        headers = {
            'Content-Type': "text/xml",
            'Accept': "application/xml",
            'Cache-Control': "no-cache",
            'cache-control': "no-cache",
            'Postman-Token': "32e4f32d-5dfd-41ee-b48e-7f34210999dd"
            }

        response = requests.request("POST", url, data=payload, headers=headers, params=querystring)                       # create api called
        print()
        print(response.text)
        print()



if __name__ == '__main__':

    command_line_param = sys.argv[1]                                                                                       # takes input from command line
    suite = unittest.TestLoader().loadTestsFromTestCase(TestSuite1)                                                        # it loads all the test cases
    result = not unittest.TextTestRunner(stream=sys.stderr, failfast=True, verbosity=2).run(suite).wasSuccessful()         # it runs all test cases and stores
    sys.exit(result)
