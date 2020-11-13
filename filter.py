#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 16 12:53:50 2020

@author: murat
"""

from sys import exit
from zimbraMail import zimbraMail
import json
from os import path

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
finally:
    mail.close()


# todo: auto delete messages

# write operation you want here


# Export csv data if chosen
if config["exportCsvFile"] != "":
    mail.export(config["exportCsvFile"], config["lastExportTime"])
    config["lastExportTime"] = mail.strTimeNow()
    with open("config.json", "w") as f:
        json.dump(config, f)

if config["closeAfterProcess"]:
    mail.close()
