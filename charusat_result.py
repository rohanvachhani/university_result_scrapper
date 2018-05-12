# -*- coding: utf-8 -*-
"""
Created on Fri May 11 20:46:53 2018

@author: "Rohan Vachhani"
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import csv


#========================================================================
def select_from_menu(ele_id, option_name):
    el = driver.find_element_by_id(ele_id)
    for option in el.find_elements_by_tag_name('option'):
        if option.text == option_name:
            option.click() # select() in earlier versions of webdriver  _rv
            break


driver = webdriver.Chrome("C:\cds\py\chromedriver.exe")
driver.maximize_window()
url = "http://117.239.83.200:2020/" 
driver.get(url)

BRANCH = 'CSPIT'
FIELD = 'BTECH(IT)'
SEM = '5'

select_from_menu('ddlInst', BRANCH)
select_from_menu('ddlDegree',FIELD)
select_from_menu('ddlSem', SEM)         #insert sem here on place of 5  _rv
select_from_menu('ddlScheduleExam', 'NOVEMBER 2017')    #insert year of exam here non place of november 2017   _rv

#el2 = driver.find_element_by_name('txtEnrNo')
#el2.send_keys('15ce146')
#el2.send_keys(Keys.RETURN)

#=====================================================================
#start scrapping result

file = open("2_charusat_result_data_it.csv", "w", newline='')
field_names = ['s_number', 'Name', 'semester', 'exame_year', 'SGPA', 'CGPA']
writer = csv.DictWriter(file, fieldnames=field_names)
writer.writeheader()

for i in range(1,151):        #enter ID number range in place od 1 to 151
    try:
        el2 = driver.find_element_by_name('txtEnrNo')
        iid = '15'+ FIELD[6:8]+str('{:03}'.format(i))   #FIELD[6:8]=CE
        #print(iid)
        
        el2.send_keys(Keys.CONTROL + "a")
        el2.send_keys(Keys.DELETE)
        
        el2.send_keys(iid)
        el2.send_keys(Keys.RETURN)
        
        s_num = driver.find_element_by_id('uclGrd1_lblExamNo')
        name = driver.find_element_by_id('uclGrd1_lblStudentName')
        sem = driver.find_element_by_id('uclGrd1_lblSemester')
        year = driver.find_element_by_id('uclGrd1_lblMnthYr')
        sgpa = driver.find_element_by_id('uclGrd1_lblSGPA')
        cgpa = driver.find_element_by_id('uclGrd1_lblCGPA')
        
        print (s_num.text," ",name.text, " ", year.text, " ", sem.text," ", sgpa.text, " ", cgpa.text)
        writer.writerow({'s_number':s_num.text, 'Name': name.text, 'semester': sem.text, 'exame_year': year.text, 'SGPA': sgpa.text, 'CGPA': cgpa.text})
        
        driver.execute_script("window.history.go(-1)")
    except Exception as e:
        continue            #skip the invalid ID due to kami students.._rv
        
    
        
file.close()   

