import PySimpleGUI as sg
import datetime
from datetime import timedelta
import hashlib
import os
import time

import chromedriver_autoinstaller
import pandas as pd
import selenium
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


""" Create the window for the program """
sg.theme('DarkBlue3')

""" Change python's print to sg.print to print to PySimpleGUI debug window """
print = sg.Print


layout = [
    [sg.Text('Username:', size=(10,1)), sg.InputText(size=(15,1), password_char='*'), sg.Text('Password:', size=(10,1)), sg.InputText(size=(15,1), password_char='*')],
    [sg.Text(key='login_message', size=(55,1))],
    [sg.Text('Customer Invoice Sheet'), sg.Input(''), sg.FileBrowse('Browse')],
    [sg.Text(key='time_taken', size=(55,1))],
    [sg.Button('Start', size=(10,1)), sg.Button('Cancel', size=(10,1))],
]

window = sg.Window('billing stages - temp', layout)
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break
    if event == 'Start':
        username = values[0]
        password = values[1]
        complete = False
        if complete != True:

            ukey = 'USERNAME KEY'
            usalt = 'USERNAME SALT'

            pkey = 'PASSWORD KEY'
            psalt = 'PASSWORD SALT'

            provided_username = hashlib.pbkdf2_hmac(
                'sha256', # Hash digest algorithm for HMAC
                username.encode('utf-8'),
                usalt,
                100000   
                )

            provided_password = hashlib.pbkdf2_hmac(
                'sha256', # Hash digest algorithm for HMAC
                password.encode('utf-8'),
                psalt,
                100000   
                )

            if provided_username == ukey and provided_password == pkey:
                window['login_message'].update('Username and Password are Correct. Billing Automation Script is Starting.')
                print('Username and Password are Correct') 
                complete = True
            else:
                window['login_message'].update('Username or Password are Incorrect. Try again.')
                print('Username or Password is Incorrect')
                window[0].update(''), window[1].update('')
                continue
            complete = True
        
        print('This Billing Automation Script is Starting')
        chromedriver_autoinstaller.install()
        '''
        This is the Class object BillingAutomation which uses Pandas (located at the end of the program) to unpack an Excel
        Spreadsheet of variables per school to pass along to the functions within the Class. Other functions within this Class
        will have their own documentation about what they do.
        '''

        ''' This is the beginning of the Automated Billing Program. It uses Selenium Chrome WebDriver to run the automation. '''

        # current date variable
        current_date = datetime.datetime.now()

        class BillingAutomation:
            def __init__(self, school_name, current_invoice, period, course_changed_credit, no_start_credit, withdraw_credit, address):
                school_name = self.school_name = school_name
                current_invoice = self.current_invoice = str(current_invoice)
                period = self.period = period.upper()
                course_changed_credit = self.course_changed_credit = course_changed_credit.upper()
                no_start_credit = self.no_start_credit = no_start_credit.upper()
                withdraw_credit = self.withdraw_credit = withdraw_credit.upper()
                address = self.address = address
                # browser = self.browser = browser
                driver = self.driver = webdriver.Chrome()
                driver.maximize_window()

                # This is where the list of dictionaries that holds refunds is initialized
                # This is inside the Class to have a fresh excel sheet generated for each school
                # Reference: https://stackoverflow.com/questions/20638006/convert-list-of-dictionaries-to-a-pandas-dataframe
                auto_billing_report = []

                print('Completing billing for: ' + school_name + '\nInvoice: ' + str(current_invoice))

                # Open Browser
                print('Driver initializing')
                driver.get(address)
                print('Finished driver init and completed navigation')

                # Log in to billing system
                print('Finding username and password and sending both')
                driver.find_element_by_name('uname').send_keys(username)
                driver.find_element_by_name('password').send_keys(password)
                print('Logging in')
                try:
                    driver.find_element_by_class_name("bttn").click()
                except NoSuchElementException:
                    driver.find_element_by_xpath('//*[@id="butlogin"]').click()

                # If the browser is redirected to the school's Home page instead of Billing, this will reopen browser to change it to Billing
                try:
                    if WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'location1'))).text == 'Location':
                        try:
                            driver.switch_to.alert.accept()
                            main_window = driver.window_handles[0]
                            driver.switch_to.window(main_window)
                            driver.get(address)
                            print('\nMessages handled')
                        except:
                            pass
                            print('\nNo messages to handle')
                except TimeoutException:
                    driver.get(address)
                                
                # If a New Message Alert is found, this will click OK to open New Message pop-up, then return to billing page.
                # When the browser is closed at the end, the pop up should also close.
                try:
                    Alert(driver).accept()
                    main_window = driver.window_handles[0]
                    driver.switch_to.window(main_window)
                    print('\nMessages handled')
                except:
                    pass
                    print('\nNo messages to handle')

                # This creates variables depending on if using Xbrowser or not
                    
                def text_present():
                    WebDriverWait(driver, 20).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="tblList_info"]/div[1]'), 'Total'))

                def press_filter():
                    text_present()
                    driver.find_element_by_xpath('//*[@id="filterButton"]').click()
                    time.sleep(1)
                    text_present()

                def select_all():
                    text_present()
                    driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[3]/td[4]/input[1]').click()
                    time.sleep(1)
                    text_present()

                def press_update():
                    text_present()
                    select_all()
                    driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[3]/td[4]/input[2]').click()
                    driver.switch_to.alert.accept()
                    time.sleep(1)
                    text_present()
                    

                # This function called when in ACTIVE/COMPLETED - OPEN. This is the IF/THEN logic of billing
                def active_completed_logic():
                    try:
                        row_count = len(driver.find_elements_by_xpath('//*[@id="tblList"]/tbody/tr'))
                        r = 1
                        if driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[8]')).text == '':
                            print('NEW COURSE ALERT! There is a NEW COURSE: {}'.format(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text))
                            auto_billing_report.append({
                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                'Course': 'NEW COURSE',
                                'New Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text,
                            })
                        else:
                            while row_count >= r:
                                course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                if invoice_num == current_invoice:
                                    pass
                                elif invoice_num == '' and course_charge != '.00':
                                    print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Open')
                                    Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Open') # Sets Charge Status to OPEN
                                    driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                elif invoice_num != '' and course_charge != '.00':
                                    print('Has invoice, has charge. Invoice: Same Charge Status: Paid')
                                    Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                elif invoice_num == '' and course_charge == '.00':
                                    print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                    Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                elif invoice_num != '' and course_charge == '.00':
                                    print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                    Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                r+=1
                            press_update()
                    except NoSuchElementException:
                        print('No Data in ' + Select(driver.find_element_by_name('location1')).first_selected_option.text + ': ' + Select(driver.find_element_by_name('course')).first_selected_option.text)
                        #time.sleep(.5)

                # This function called when in COURSE CHANGED - OPEN. This is the IF/THEN logic of billing
                # It also takes input from billing spreadsheet to determine how it should handle refunds.
                def course_changed_logic():
                    try:
                        row_count = len(driver.find_elements_by_xpath('//*[@id="tblList"]/tbody/tr'))
                        r = 1
                        if driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[8]')).text == '':
                            print('NEW COURSE ALERT! There is a NEW COURSE: ' + driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text)
                            auto_billing_report.append({
                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                'Course': 'NEW COURSE',
                                'New Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text,
                            })
                        else:
                            if course_changed_credit == 'FULL':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Course Changed',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Course Changed',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                        if student_charge == '':
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'Course Changed',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': float(course_charge),
                                            })
                                        elif student_charge != '':
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'Course Changed',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': float(student_charge),
                                            })
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif course_changed_credit == 'HALF':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    course_completion_stats = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).get_attribute('title').split()
                                    completed = float(course_completion_stats[1])
                                    omitted = float(course_completion_stats[3])
                                    remaining = float(course_completion_stats[5])
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Course Changed',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Course Changed',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00' and completed <= ((completed + omitted + remaining) / 2):
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                        if student_charge == '':
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'Course Changed',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': float(course_charge),
                                            })
                                        elif student_charge != '':
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'Course Changed',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': float(student_charge),
                                            })
                                    elif invoice_num != "" and course_charge != '.00' and completed > ((completed + omitted + remaining) / 2):
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Paid')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif course_changed_credit == 'PRO':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    course_completion_stats = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).get_attribute('title').split()
                                    completed = float(course_completion_stats[1])
                                    omitted = float(course_completion_stats[3])
                                    remaining = float(course_completion_stats[5])
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if student_charge == '':
                                                cc_prorate = float(course_charge) * ((remaining) / (completed + omitted + remaining))
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Course Changed',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': round(cc_prorate, 2),
                                                })
                                            elif student_charge != '':
                                                sc_prorate = float(student_charge) * ((remaining) / (completed + omitted + remaining))
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Course Changed',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': round(sc_prorate, 2),
                                                })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Refunded Prorating Fee')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                        if student_charge == '':
                                            cc_prorate = float(course_charge) * ((remaining) / (completed + omitted + remaining))
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'Course Changed',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': round(cc_prorate, 2),
                                            })
                                        elif student_charge != '':
                                            sc_prorate = float(student_charge) * ((remaining) / (completed + omitted + remaining))
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'Course Changed',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': round(sc_prorate, 2),
                                            })
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif course_changed_credit == 'NONE':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    if invoice_num == current_invoice:
                                        pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Open')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Open') # Sets Charge Status to OPEN
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Paid')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            else:
                                print('UNKNOWN VARIABLE! ' + course_changed_credit + ' NOT A VALID VARIABLE')
                                auto_billing_report.append({
                                    'Course': 'UNKNOWN VARIABLE', 
                                    'Unknown Variable': withdraw_credit
                                })
                    except NoSuchElementException:
                        print('No Data in ' + Select(driver.find_element_by_name('location1')).first_selected_option.text + ': ' + Select(driver.find_element_by_name('course')).first_selected_option.text)
                        #time.sleep(.5)

                # This function called when in NO START - OPEN. This is the IF/THEN logic of billing
                # It also takes input from billing spreadsheet to determine how it should handle refunds.        
                def no_start_logic():
                    try:
                        row_count = len(driver.find_elements_by_xpath('//*[@id="tblList"]/tbody/tr'))
                        r = 1
                        if driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[8]')).text == '':
                            print('NEW COURSE ALERT! There is a NEW COURSE: ' + driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text)
                            auto_billing_report.append({
                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                'Course': 'NEW COURSE',
                                'New Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text
                            })
                        else:
                            if no_start_credit == 'FULL':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'No Start',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'No Start',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                        if student_charge == '':
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'No Start',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': float(course_charge),
                                            })
                                        elif student_charge != '':
                                            auto_billing_report.append({
                                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                'Course Status': 'No Start',
                                                'Billing Status': 'Credit',
                                                'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                'Charge': float(student_charge),
                                            })
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif no_start_credit == 'NONE':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    if invoice_num == current_invoice:
                                        pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Open')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Open') # Sets Charge Status to OPEN
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Paid')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            else:
                                print('UNKNOWN VARIABLE! ' + no_start_credit + ' NOT A VALID VARIABLE')
                                auto_billing_report.append({
                                    'Course': 'UNKNOWN VARIABLE', 
                                    'Unknown Variable': withdraw_credit
                                })
                    except NoSuchElementException:
                        print('No Data in ' + Select(driver.find_element_by_name('location1')).first_selected_option.text + ': ' + Select(driver.find_element_by_name('course')).first_selected_option.text)
                        #time.sleep(.5)

                # This function called when in WITHDRAW - OPEN. This is the IF/THEN logic of billing
                # It also takes input from billing spreadsheet to determine how it should handle refunds. 
                def withdraw_logic():
                    try:
                        row_count = len(driver.find_elements_by_xpath('//*[@id="tblList"]/tbody/tr'))
                        r = 1
                        if driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[8]')).text == '':
                            print('NEW COURSE ALERT! There is a NEW COURSE: ' + driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text)
                            auto_billing_report.append({
                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                'Course': 'NEW COURSE',
                                'New Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text
                            })
                        else:
                            if withdraw_credit == 'FULL':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    course_completion_stats = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).get_attribute('title').split()
                                    completed = float(course_completion_stats[1])
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if no_start_credit == 'FULL' and completed == 0.0:
                                                print('Actually No Start. Adding to Refund list as No Start')
                                                if student_charge == '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(course_charge),
                                                        })
                                                elif student_charge != '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(student_charge),
                                                    })
                                            else:                                                              
                                                print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                                if student_charge == '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(course_charge),
                                                    })
                                                elif student_charge != '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(student_charge),
                                                    })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        if no_start_credit == 'FULL' and completed == 0.0:
                                            print('Actually No Start. Adding to Refund list as No Start')
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                    })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                        else:                                                              
                                            print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif withdraw_credit == 'HALF':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    course_completion_stats = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).get_attribute('title').split()
                                    completed = float(course_completion_stats[1])
                                    omitted = float(course_completion_stats[3])
                                    remaining = float(course_completion_stats[5])
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if no_start_credit == 'FULL' and completed == 0.0:
                                                print('Actually No Start. Adding to Refund list as No Start')
                                                if student_charge == '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(course_charge),
                                                    })
                                                elif student_charge != '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(student_charge),
                                                    })
                                            else:                                                              
                                                print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                                if student_charge == '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(course_charge),
                                                    })
                                                elif student_charge != '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(student_charge),
                                                    })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00' and completed <= ((completed + omitted + remaining) / 2):
                                        if no_start_credit == 'FULL' and completed == 0.0:
                                            print('Actually No Start. Adding to Refund list as No Start')
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                        else:                                                              
                                            print('Has invoice, has charge. Invoice: Same Charge Status: Refunded')
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                    elif invoice_num != "" and course_charge != '.00' and completed > ((completed + omitted + remaining) / 2):
                                        print('Has invoice, has charge. Invoice: Same Charge Status: Paid')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif withdraw_credit == 'PRO':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    course_completion_stats = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).get_attribute('title').split()
                                    completed = float(course_completion_stats[1])
                                    omitted = float(course_completion_stats[3])
                                    remaining = float(course_completion_stats[5])
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if no_start_credit == 'FULL' and completed == 0.0:
                                                print('Actually No Start. Adding to Refund list as No Start')
                                                if student_charge == '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(course_charge),
                                                    })
                                                elif student_charge != '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(student_charge),
                                                    })
                                            else:                                                              
                                                print('Has invoice, has charge. Invoice: Same Charge Status: Refunded Prorating Fee')
                                                if completed > ((completed + omitted + remaining) / 2):
                                                    Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                                elif completed <= ((completed + omitted + remaining) / 2):
                                                    Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets C   harge Status to REFUNDED, leaving Invoice alone
                                                    if student_charge == '':
                                                        cc_prorate = float(course_charge) * ((remaining) / (completed + omitted + remaining))
                                                        auto_billing_report.append({
                                                            'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                            'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                            'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                            'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                            'Course Status': 'Withdraw',
                                                            'Billing Status': 'Credit',
                                                            'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                            'Charge': round(cc_prorate, 2),
                                                        })
                                                    elif student_charge != '':
                                                        sc_prorate = float(student_charge) * ((remaining) / (completed + omitted + remaining))
                                                        auto_billing_report.append({
                                                            'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                            'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                            'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                            'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                            'Course Status': 'Withdraw',
                                                            'Billing Status': 'Credit',
                                                            'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                            'Charge': round(sc_prorate, 2),
                                                        })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        if no_start_credit == 'FULL' and completed == 0.0:
                                            print('Actually No Start. Adding to Refund list as No Start')
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                        else:                                                              
                                            print('Has invoice, has charge. Invoice: Same Charge Status: Refunded Prorating Fee')
                                            if completed > ((completed + omitted + remaining) / 2):
                                                Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                            elif completed <= ((completed + omitted + remaining) / 2):
                                                Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets C   harge Status to REFUNDED, leaving Invoice alone
                                                if student_charge == '':
                                                    cc_prorate = float(course_charge) * ((remaining) / (completed + omitted + remaining))
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': round(cc_prorate, 2),
                                                    })
                                                elif student_charge != '':
                                                    sc_prorate = float(student_charge) * ((remaining) / (completed + omitted + remaining))
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': round(sc_prorate, 2),
                                                    })
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            elif withdraw_credit == 'NONE':
                                while row_count >= r:
                                    course_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[8]').format(r)).text # NO PARENTHESIS ON .text
                                    student_charge = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[9]/input').format(r)).get_attribute('value')
                                    invoice_num = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value")
                                    course_completion_stats = driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).get_attribute('title').split()
                                    completed = float(course_completion_stats[1])
                                    if invoice_num == current_invoice:
                                        if Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select'.format(r)))).first_selected_option.text == 'Refunded':
                                            if no_start_credit == 'FULL' and completed == 0.0:
                                                print('Actually No Start. Adding to Refund list as No Start')
                                                if student_charge == '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(course_charge),
                                                    })
                                                elif student_charge != '':
                                                    auto_billing_report.append({
                                                        'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                        'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                        'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                        'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                        'Course Status': 'Withdraw (Actually No Start)',
                                                        'Billing Status': 'Credit',
                                                        'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                        'Charge': float(student_charge),
                                                    })
                                        else:
                                            pass
                                    elif invoice_num == "" and course_charge != '.00':
                                        print('No invoice, has charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Open')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Open') # Sets Charge Status to OPEN
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge != '.00':
                                        if no_start_credit == 'FULL' and completed == 0.0:
                                            print('Actually No Start. Adding to Refund list as No Start')
                                            if student_charge == '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(course_charge),
                                                })
                                            elif student_charge != '':
                                                auto_billing_report.append({
                                                    'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                                    'Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[3]').format(r)).text,
                                                    'Student': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[2]').format(r)).text,
                                                    'ETA ID': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[6]').format(r)).text,
                                                    'Course Status': 'Withdraw (Actually No Start)',
                                                    'Billing Status': 'Credit',
                                                    'Invoice': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).get_attribute("value"),
                                                    'Charge': float(student_charge),
                                                })
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Refunded') # Sets Charge Status to REFUNDED, leaving Invoice alone
                                        else:
                                            print('Has invoice, has charge. Invoice: Same Charge Status: Paid')
                                            Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Paid') # Sets Charge Status to PAID, leaving Invoice alone
                                    elif invoice_num == "" and course_charge == '.00':
                                        print('No invoice, no charge. Setting Invoice: ' + str(current_invoice) + ' Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                        driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[11]/input[5]').format(r)).send_keys(current_invoice) # Sets Invoice to CURRENT INVOICE
                                    elif invoice_num != "" and course_charge == '.00':
                                        print('Has invoice, no charge. Setting Invoice: Same Charge Status: Waived')
                                        Select(driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[{}]/td[10]/select').format(r))).select_by_visible_text('Waived') # Sets Charge Status to WAIVED
                                    r+=1
                                press_update()
                            else:
                                print('UNKNOWN VARIABLE! ' + withdraw_credit + ' NOT A VALID VARIABLE')
                                auto_billing_report.append({
                                    'Course': 'UNKNOWN VARIABLE', 
                                    'Unknown Variable': withdraw_credit
                                })
                    except NoSuchElementException:
                        print('No Data in ' + Select(driver.find_element_by_name('location1')).first_selected_option.text + ': ' + Select(driver.find_element_by_name('course')).first_selected_option.text)
                        #time.sleep(.5)

                def course_status_changer():
                    Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[2]/select')).select_by_visible_text('All')
                    Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[3]/select')).select_by_visible_text('Open')
                    press_filter()
                    try:
                        if driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[8]')).text == '':
                            print('NEW COURSE ALERT! There is a NEW COURSE: ' + driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text)
                            auto_billing_report.append({
                                'Location': Select(driver.find_element_by_name('location1')).first_selected_option.text,
                                'Course': 'NEW COURSE',
                                'New Course': driver.find_element_by_xpath(('//*[@id="tblList"]/tbody/tr[1]/td[3]')).text,
                            })
                        else:
                            # Changing Charge and Course Status Definitions
                            def charge_status_open():
                                Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[3]/select')).select_by_visible_text('Open')

                            # Course Status Active
                            text_present()
                            Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[2]/select')).select_by_visible_text('Active')
                            charge_status_open()
                            press_filter()
                            active_completed_logic()

                            # Course Status Completed
                            text_present()
                            Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[2]/select')).select_by_visible_text('Completed')
                            charge_status_open()
                            press_filter()
                            active_completed_logic()

                            # Course Status Course Changed
                            text_present()
                            Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[2]/select')).select_by_visible_text('Course Changed')
                            charge_status_open()
                            press_filter()
                            course_changed_logic()

                            # Course Status No Start
                            text_present()
                            Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[2]/select')).select_by_visible_text('No Start')
                            charge_status_open()
                            press_filter()
                            no_start_logic()

                            # Course Status Withdraw
                            text_present()
                            Select(driver.find_element_by_xpath('//*[@id="form1"]/table/tbody/tr[2]/td[2]/select')).select_by_visible_text('Withdraw')
                            charge_status_open()
                            press_filter()
                            withdraw_logic()

                    except NoSuchElementException:
                        pass
                        print('No Data in ' + Select(driver.find_element_by_name('location1')).first_selected_option.text + ': ' + Select(driver.find_element_by_name('course')).first_selected_option.text)
                        #time.sleep(.5)

                # Determines the number of Locations at school to select either Multi or Single location script, then builds list of courses to iterate through
                def location_selection():
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'location1')))
                    print('\nFinding and listing locations at ' + school_name)
                    location = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'location1'))))
                    location_list = ([l.text for l in location.options])
                    location_list.remove('')
                    print(location_list)
                    while location_list == []:
                        location = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'location1'))))
                        location_list = ([l.text for l in location.options])
                        location_list.remove('')
                        print(location_list)
                    l = 1
                    if len(location_list) != 1:
                        print('Running Script for multiple locations')
                        while l <= len(location_list):
                            location = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'location1'))))
                            location.select_by_index(l)
                            print('\nFinding and listing courses at ' + school_name + '\'s ' + str(location_list[l-1]) + ' location\n')
                            course = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'course'))))
                            course_list = ([c.text for c in course.options])
                            del course_list[0:2]
                            print(course_list)
                            while course_list == []:
                                course = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'course'))))
                                course_list = ([c.text for c in course.options])
                                del course_list[0:2]
                                print(course_list)
                            l += 1
                            c = 2
                            while c <= len(course_list) + 1:
                                Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'course')))).select_by_index(c)
                                course_status_changer()
                                c +=1
                    elif len(location_list) == 1:
                        print('Running script for single location')
                        print('\nFinding and listing courses at ' + school_name + '\'s ' + str(location_list[l-1]) + ' location\n')
                        location.select_by_index(1)
                        course = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'course'))))
                        course_list = ([c.text for c in course.options])
                        del course_list[0:2]
                        print(course_list)
                        while course_list == []:
                            course = Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'course'))))
                            course_list = ([c.text for c in course.options])
                            del course_list[0:2]
                            print(course_list)
                        c = 2
                        while c <= len(course_list) + 1:
                            Select(WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'course')))).select_by_index(c)
                            course_status_changer()
                            c += 1

                location_selection()


                # create a new folder to put credits in if it doesn't exist
                if not os.path.exists('../3 - Credits - 3/{}_{}'.format(current_date.year, current_date.month)):
                    os.makedirs('../3 - Credits - 3/{}_{}'.format(current_date.year, current_date.month))

                # Uses pandas to convert to dataframe to then create Excel spreadsheet from list-dict 'auto_billing_report
                if auto_billing_report != []:
                    auto_billing_report_df = pd.DataFrame.from_dict(auto_billing_report)
                    writer = pd.ExcelWriter('../3 - Credits - 3/{}_{}/{} {} Credits Report.xlsx'.format(current_date.year, current_date.month, current_invoice, school_name), engine='xlsxwriter')
                    print(auto_billing_report_df)
                    auto_billing_report_df.to_excel(writer, index=False, header=True)
                    writer.save()
                elif auto_billing_report == []:
                    #pass
                    print('No Credits for ' + school_name)

                # create log for faster restarts after program hangs up or stops
                if not os.path.exists('./logs/'):
                    os.makedirs('./logs/')
                with open('./logs/{}_{}_complete.log'.format(current_date.year, current_date.month), 'a') as log:
                    log.write('{}\n'.format(school_name))

                # Close browser, exiting the billing system
                print('Closing Browser')
                driver.quit()


        ''' 
        -----------------------------------
        Here is the beginning of the program
        -----------------------------------
        '''

        # timer to show how long the billing process has taken for benchmarking
        start_time = time.time()

        ''' This function takes an Excel spreadsheet with the schools and 
        their variables, reads it, then makes it understandable by the above 
        Class BillingAutomation to itterate through the schools'''

        try: # if there is a current log for this billing cycle, the program will read it to determine the next school to start on
            with open('./logs/{}_{}_complete.log'.format(current_date.year, current_date.month), 'r') as log:
                school_search = ''.join(log.readlines()[-1:]).strip('\n')
                school_df = pd.read_excel(values[2], sheet_name='Sheet1')
                print(school_search)
                for row in range(school_df.shape[0]):
                    for col in range(school_df.shape[1]):
                        if school_df.iat[row,col] == school_search:
                            row_start = row
                            restart_school_df = school_df.loc[row_start+1:]
                            schools = restart_school_df.values.tolist()
                            print(restart_school_df)

                            billingautomation_instances = []
                            for s in schools:
                                billingautomation_instances.append(BillingAutomation(*s))
        except FileNotFoundError: # if there is no current log available, the program starts at the beginning of the school excel sheet
            school_df = pd.read_excel(values[2], sheet_name='Sheet1')

            schools = school_df.values.tolist()
            print(school_df)

            billingautomation_instances = []
            for s in schools:
                billingautomation_instances.append(BillingAutomation(*s))

        elapsed_time_secs = time.time() - start_time
        window['time_taken'].update('Billing Duration: %s (hh:mm:ss)' % timedelta(seconds=round(elapsed_time_secs)))
        print('Billing Complete')


window.close()