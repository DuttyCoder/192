#ele -*- coding: utf-8 -*-
"""
Created on Fri Nov  2 00:32:54 2018

@author: Profess0r
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from pandas import *
from _collections import defaultdict
from selenium.webdriver.common.proxy import Proxy, ProxyType
import re
save_file = input('Save file as')

profile = webdriver.FirefoxProfile()
profile.set_preference("network.proxy.type", 1)
profile.set_preference("network.proxy.http", '37.48.118.4')
profile.set_preference("network.proxy.http_port", 13041)
profile.set_preference("network.proxy.ssl", '37.48.118.4')
profile.set_preference("network.proxy.ssl_port", 13041)
driver = webdriver.Firefox(firefox_profile=profile)

df = read_excel('C:/Users/Profess0r/Downloads/2.xls', sheet_name='Sheet1')
result_dict = defaultdict(list)
for i in df.index:
    name = df['name'][i]
    postcode = df['postcode'][i]
    details = []
    driver = webdriver.Firefox()
    driver.implicitly_wait(5)
    driver.get("https://www.192.com/people/")
    assert "People Search UK - People Finder - 192.com" in driver.title
    name_field = driver.find_element_by_id("peopleBusinesses_name")
    name_field.send_keys(name)
    postcode_field = driver.find_element_by_id("where_location")
    postcode_field.send_keys(postcode)
    postcode_field.submit()

    try:
        name_element = driver.find_element_by_xpath("//*[@id='resultsPremium']/div[3]/table/tbody/tr[1]/td[2]/a")
        name = name_element.text
        if len(name) < 1:
            driver.close()
            continue
    except:
        driver.close()
        continue


    try:
        postcode_element = driver.find_element_by_xpath("//*[@id='resultsPremium']/div[3]/table/tbody/tr[1]/td[3]/span[2]")
        postcode = postcode_element.text
        postcode = re.sub('\nFull Address', postcode)
    except:
        pass

    try:
        age_element = driver.find_element_by_xpath("//*[@id='resultsPremium']/div[3]/table/tbody/tr[1]/td[2]/div/span[2]")
        age = age_element.text
    except:
        print("no age found")
        age = 'NA'
    
    try:
        elec_roll_element = driver.find_element_by_xpath("//*[@id='resultsPremium']/div[3]/table/tbody/tr[1]/td[5]")
        elec_roll = elec_roll_element.text
    except:
        print("no electoral roll data")
        elec_roll = 'NA'

    try:
        other_occ_element = driver.find_element_by_xpath("//*[@id='resultsPremium']/div[3]/table/tbody/tr[1]/td[4]")
        other_occ = other_occ_element.text
    except:
        print("no other occupants found")
        other_occ = 'NA'
    driver.close()
    
    details.append(name)
    details.append(age)
    details.append(postcode)
    details.append(elec_roll)
    details.append(other_occ)
    print(details)
    for x in details:
        result_dict[i].append(x)
    
    
    
df = DataFrame(result_dict)
df.to_excel(save_file +'.xls')


