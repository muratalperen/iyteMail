# -*- coding: utf-8 -*-

"""
zimbraMail
~~~~~~~~~~~~~

This module contains the Zimbra Mail operations.
"""

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import datetime
from os import path
from time import sleep
import re
import csv


class zimbraMail:
    """Zimbra mail operations"""
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
        
        # ----------- Go bottom so load all mails -----------
        scrollScript = "var objDiv = document.getElementById('zl__TV-main__rows');objDiv.scrollTop = objDiv.scrollHeight;return objDiv.scrollHeight;"
        lenOfPage = self.driver.execute_script(scrollScript)
        match=False
        while(match==False):
            print("ileri")
            lastCount = lenOfPage
            sleep(2) # todo: implicitly_wait gibi daha mantıklı birşey
            lenOfPage = self.driver.execute_script(scrollScript)
            if lastCount==lenOfPage:
                match=True

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

    def selectWithSender(self, sender):
        """
        Finds mails by sender name and returns mail id list.

        :param sender: Sender name
        """
        for row in self.mails:
            if row["sender"] == sender:
                yield row["id"]

    def selectWithWords(self, words):
        """
        Finds mails by word list.

        :param sender: Sender name
        """
        for row in self.mails:
            for word in words:
                if re.search(word, row["title"], re.IGNORECASE):
                    yield row["id"]
                    break

    def clickDelete(self):
        """Clicks delete button"""
        self.driver.execute_script("document.getElementById(\"zb__TV-main__DELETE\").click()")
    
    def selectMail(self, mailId):
        """
        Makes mails selected by Id

        :param mailId: id of mail
        """
        if not self.isMailSelected(mailId):
            while not self.isMailSelected(mailId): # DEBUG: Sometimes first click not working so I added to while. And selenium click is slow
                self.driver.execute_script("document.getElementById('" + mailId.replace("zli", "zlif") + "__se').click()")

    def getSelectedMailList(self):
        """Returns selected mails"""
        for m in self.mails:
            if self.isMailSelected(m["id"]):
                yield m
    
    def getMailList(self):
        """Returns all mails"""
        return self.mails
    
    def isMailSelected(self, mailId):
        """
        Detects is mail given Id selected

        :param mailId: Mail Id
        """
        return self.driver.find_element_by_css_selector("#" + mailId.replace("zli", "zlif") + "__se > div").get_attribute("class") == "ImgCheckboxChecked"
    
    def export(self, logFile, lastSavingTime):
        """
        Exports mail infos to csv file

        :param logFile: Csv folder name
        :param lastSavingTime: last time mails are saved to csv file for adding only new mails
        """
        newCsv = False
        if not path.exists(logFile):
            open(logFile, "w").close()
            newCsv = True
        with open(logFile, "a") as f:
            writer = csv.DictWriter(f, fieldnames=list(self.mails[0].keys())[1:], extrasaction="ignore", )
            writer.writeheader()
            savedTime = datetime.datetime.strptime(lastSavingTime, self.DTF)
            for amail in self.mails[1:]:
                mailDate = datetime.datetime.strptime(amail["date"] + " " + amail["hour"], self.DTF)
                if savedTime < mailDate:
                    writer.writerow(amail)
                    # DEBUG: bazı mailleri tekrar ekliyor print("Added Mail:", amail["title"])

        # TODO: if csv file exists, Dictwriter adding col names again. find better solution
        if not newCsv:
            with open(logFile, "r") as r:
                csvContent = r.read()
                csvContent = csvContent.replace("read,flagged,attachment,date,hour,size,sender,title,text\n", "")
                csvContent = "read,flagged,attachment,date,hour,size,sender,title,text\n" + csvContent
            with open(logFile, "w") as w:
                w.write(csvContent)
                
    def mailById(self, mId):
        """
        Finds mail by Id

        :param mId: Mail Id
        """
        for m in self.mails:
            if m["id"] == mId:
                return m
    
    def strTimeNow(self):
        """Returns now date as standart format"""
        return datetime.datetime.now().strftime(self.DTF)
            
    def close(self):
        """Closes browser window"""
        self.driver.quit()
    