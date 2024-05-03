# title: "AutomatedBiddingNordpoolAPI"
# description: "This script executes the NordpoolAPI console application for placing latest bids at the exchange (e.g. 'nordpool_nl')."
# output: ""
# parameters: {}
# owner: "MECO, Lukas Dicke"

"""

Usage:
    AutomatedBiddingNordpoolAPI.py <job_path> --exchangeName=<string> --daysAhead=<int> --environment=<string>

Options:
    --exchangeName=<string> Matchname of exchange name, see file 'ConfigDataNordpool.xml' in known location.
    --daysAhead=<int> rel. days to today => Delivery-Day.
    --environment=<string> Test ("test") /Prod ("prod") environment (default: "test")

"""

import subprocess
from subprocess import run
import sys
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import json
import time

def SendMailPythonServer(send_to, send_cc, send_bcc, subject, body, files=[]):
    msgBody = """<html><head></head>
        <style type = "text/css">
            table, td {height: 3px; font-size: 14px; padding: 5px; border: 1px solid black;}
            td {text-align: left;}
            body {font-size: 12px;font-family:Calibri}
            h2,h3 {font-family:Calibri}
            p {font-size: 14px;font-family:Calibri}
         </style>"""

    msgBody += "<h2>" + subject + "</h2>"
    # msgBody += "<h3>" + message + "</h3>"
    msgBody += body

    strFrom = "no-reply-duswvpyt002p@statkraft.de"
    #strFrom = "nominations@statkraft.de"
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = strFrom
    if len(send_to) == 1:
        msgRoot['To'] = send_to[0]
    else:
        msgRoot['To'] = ",".join(send_to)

    if len(send_cc) == 1:
        msgRoot['Cc'] = send_cc[0]
    else:
        msgRoot['Cc'] = ",".join(send_cc)

    if len(send_cc) == 1:
        msgRoot['Bcc'] = send_bcc[0]
    else:
        msgRoot['Bcc'] = ",".join(send_bcc)
    msgRoot.preamble = 'This is a multi-part message in MIME format.'

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msgRoot.attach(part)

    msgText = MIMEText('Sorry this mail requires your mail client to allow for HTML e-mails.')
    msgAlternative.attach(msgText)

    msgText = MIMEText(msgBody, 'html')
    msgAlternative.attach(msgText)

    smtp = smtplib.SMTP('smtpdus.energycorp.com')
    smtp.sendmail(strFrom, send_to, msgRoot.as_string())
    smtp.quit()

    print("Mail sent successfully from " + strFrom)

def GetLatestBidUpdateInfo(delDate, mergerName, environment):
    with open(GetLatestBidUpdateInfoFile(delDate, mergerName, environment)) as f:
        d = json.load(f)
    return d

def GetLatestBidUpdateInfoFile(delDate, mergerName, environment):
    fileUpdateInfo = delDate + "_" + mergerName + "_UpdateInfo_" + environment + ".json"
    uncPath = os.path.join(folderUpdateInfo, fileUpdateInfo)
    return uncPath

def GetContentConfigFile():
    uncPath = r"\\energycorp.com\common\Divsede\Operations\IT\APIs\NordpoolAPI_CWE\ConfigDataNordpool.json"
    with open(uncPath) as f:
        d = json.load(f)
    return d

def GetClearname(exchangeName):
    availables = GetContentConfigFile()["AvailableExchanges"]
    clearname=""
    for i in enumerate(availables):
        if i[1]["Matchname"] == exchangeName:
            clearname = i[1]["Clearname"]
            break
    return clearname

exchangeName = str(exchangeName)
daysAhead = int(daysAhead)
environment = str(environment)

# exchangeName = "nordpool_de_IDA1"
# daysAhead = 1
# environment="test"

recipientsTo = ["lukas.dicke@statkraft.de"]

processFilename = "AutomatedBidding_NordpoolAPI_CWE.exe"

folderUpdateInfo = r"\\energycorp.com\common\Divsede\Operations\IT\APIs\NordpoolAPI_CWE\BidUpdateInfo"

#path = r"\\energycorp.com\common\Divsede\Operations\IT\APIs\NordpoolAPI_CWE\64bit"
path=r"\\energycorp.com\common\Divsede\Operations\Personal_OPS\Lukas\DevelopedApplications\NordpoolAPI_CWE\AutomatedBidding_NordpoolAPI_CWE\bin\Debug"

process = subprocess.run(os.path.join(path, processFilename) + " " + exchangeName + " " + str(daysAhead) + " " + str(environment), capture_output=True)

delDate = datetime.datetime.strftime(datetime.date.today() + datetime.timedelta(days=daysAhead),"%d.%m.%Y")

timestampBidPlacing = datetime.datetime.now().strftime('%H:%M:%S')

#time.sleep(10)

exchangeClearName = GetClearname(exchangeName)

if exchangeClearName=="":
    exchangeClearName=exchangeName

if process.returncode != 0 :

    error = str(process.stdout) + "<br>Return-code: " + str( process.returncode)

    error = str(process.stdout) + "<br>Return-code: " + str(process.returncode)

    print(error)

    emailSubject = exchangeClearName + ": Error occurred when placing exchange-bid"

    message = "Exchange-bid for '" + exchangeClearName + "' " + "(Environment: " + environment + "; Delivery-Date: " + str(delDate) + ")" + " could not be placed, because the following issue occurred:<br><br>" + error
    print(message)

    messageHeader = "Hi colleagues on the intraday-desk,<br><br>please keep track on the issues below. Thanks.<br><br><br>"

    sendEmailYesOrNo = True

else:

    latestBidUpdate = GetLatestBidUpdateInfo(datetime.datetime.strftime(datetime.date.today() + datetime.timedelta(days=daysAhead), "%Y%m%d"), exchangeName,environment)

    sendEmailYesOrNo = latestBidUpdate['IsBidUpdatedComparedToFormerVersion']

    createdTimestamp = latestBidUpdate['CreationDate']

    emailSubject = exchangeClearName + ": Exchange-bid successfully placed (Del.: " + str(delDate) + "; Env.: " + environment + ")"
    messageHeader = "Hi colleagues on the intraday-desk,<br><br>"

    message = "Exchange-bid for '" + exchangeClearName + "' " + "(Environment: " + environment + "; Delivery-Date: " + str(delDate) + "; created: " + latestBidUpdate['CreationDate'] +")" +  " has been successfully placed at " + timestampBidPlacing + "<br><br>" + latestBidUpdate['UncPathLastBidPlaced'] + ".<br><br>"

    print("Exchange bidding successful!")

messageEnd = "<br><br>BR<br><br>Statkraft Operations"

if sendEmailYesOrNo == True:
    SendMailPythonServer(send_to=recipientsTo, send_cc=[], send_bcc=[], subject=emailSubject,body=messageHeader + message + messageEnd, files=[])