# -*- coding: utf-8 -*-

"""
"""

import json
from sys import exit
from os import path
from time import sleep

from zimbraMail import zimbraMail

# Open Config
if path.exists("config.json"):
    with open("config.json", "r") as f:
        config = json.load(f)
else:
    print("Missing config file.")
    exit()
    
# Check if first time
if config["userName"] == "":
    config["userName"] = input("User Name:")
    config["password"] = input("Password:")
    with open("config.json", "w") as f:
        json.dump(config, f)

# Connect mail service
try:
    mail = zimbraMail(config["userName"], config["password"])
except Exception as e:
    print("Error:", e)
    exit()

mailList = mail.getMailList()
# write operation you want here

# Configde belirttiğin, mail almak istemediğin kelimeleri seç
for selectMailId in mail.selectWithWords(config["wordBlackList"]):
    mail.selectMail(selectMailId)

# Configde belirttiğin, mail almak istemediğin kullanıcıları seç
for blackSender in config["senderBlackList"]:
    for blackMailId in mail.selectWithSender(blackSender):
        mail.selectMail(blackMailId)


# Seçili mailleri sil
print(*[tit["title"] for tit in mail.getSelectedMailList()], sep="\n")
if input("Yukarıdaki mailler silinecektir. Onaylıyor musunuz? (y/n): ") == "y":
    mail.clickDelete()
    sleep(1)
    print("Silindi")

# Export csv data if chosen

if config["exportCsvFile"] != "":
    mail.export(config["exportCsvFile"], config["lastExportTime"])
    config["lastExportTime"] = mail.strTimeNow()
    with open("config.json", "w") as f:
        json.dump(config, f)

if config["closeAfterProcess"]:
    mail.close()
