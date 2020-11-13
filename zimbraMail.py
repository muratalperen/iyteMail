#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 16 12:51:47 2020

@author: murat
"""

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import datetime
from os import path
import re
import csv


class zimbraMail:
    def __init__(self, username, password):
        
        self.DTF = "%d/%m/%Y %H:%M" # datetime format
        EXECPATH = "/home/murat/Documents/Bash/Sistem/geckodriver"
        MAILSRVR = "https://std.iyte.edu.tr/"
        
        # ----------- Login -----------
        self.driver = webdriver.Firefox(executable_path=EXECPATH)
        self.driver.get(MAILSRVR)
        self.driver.find_element_by_id("username").send_keys(username)
        self.driver.find_element_by_id("password").send_keys(password)
        self.driver.find_element_by_xpath("//input[@type='submit']").click()
        
        # ----------- Check login -----------
        try:
            self.driver.find_element_by_id("z_banner")
        except NoSuchElementException:
            self.driver.close()
            raise Exception("Wrong user name or password. You can change it from config.json file.") # todo: do it auto
            
        # ----------- Wait for getting mail list -----------
        self.driver.implicitly_wait(2)
        while len(self.driver.find_elements_by_css_selector("#zl__TV-main__rows")) == 0:
            self.driver.implicitly_wait(2)
            
        # ----------- Get mail list -----------
        mailUl = self.driver.find_element_by_id("zl__TV-main__rows")
        self.mails = []
        for mailLi in mailUl.find_elements_by_css_selector("li"):
            elementId = mailLi.get_attribute("id")
            labelText = mailLi.get_attribute("aria-label")
            if len(re.findall(", \d+/\d+/\d+,", labelText)) == 0:
                dateStr = datetime.date.today().strftime("%d/%m/%Y")
            else:
                dateStr = re.sub("/(\d\d)$", r"/20\1", re.findall(", \d+/\d+/\d+,", labelText)[0][2:-1])
                dateStr = re.sub("(\d+)/(\d+)/20(\d+)", r"\2/\1/20\3", dateStr) # m/d/Y to d/m/Y
            zlifId = elementId.replace("zli", "zlif")
            hour = datetime.datetime.strptime(re.findall(", \d+:\d+ [A|P]" + "M", labelText)[0][2:], "%I:%M %p").strftime("%H:%M") # convert AM/PM to 24
            self.mails.append({
                "id"        : elementId,
                "read"      : "Read, " in labelText,
                "flagged"   : "Flagged, " in labelText,
                "attachment": ", has attachment," in labelText,
                "date"      : dateStr,
                "hour"      : hour,
                "size"      : re.findall(", [0-9]+ .B,", labelText)[0][2:-1],
                "sender"    : mailLi.find_element_by_id(zlifId + "__fr").get_attribute('textContent'),
                "title"     : mailLi.find_element_by_id(zlifId + "__su").get_attribute('textContent'),
                "text"      : mailLi.find_element_by_id(zlifId + "__fm").get_attribute('textContent')[3:]
            })
        
    def clickDelete(self):
        self.driver.execute_script("document.getElementById(\"zb__TV-main__DELETE\").click()")
    
    def selectMail(self, mailId):
        self.driver.execute_script("document.getElementById('" + mailId.replace("zli", "zlif") + "__se').click()")
    
    def getMailList(self):
        return self.mails
    
    def export(self, logFile, lastSavingTime):
        if not path.exists(logFile):
            open(logFile, "w").close()
        with open(logFile, "a") as f:
            writer = csv.DictWriter(f, fieldnames=list(self.mails[0].keys())[1:], extrasaction="ignore")
            writer.writeheader()
            savedTime = datetime.datetime.strptime(lastSavingTime, self.DTF)
            for amail in self.mails:
                mailDate = datetime.datetime.strptime(amail["date"] + " " + amail["hour"], self.DTF)
                if savedTime < mailDate:
                    writer.writerow(amail)
                    #print("Eklendi:", amail["title"])
    
    def strTimeNow(self):
        return datetime.datetime.now().strftime(self.DTF)
            
    def close(self):
        self.driver.quit()
    