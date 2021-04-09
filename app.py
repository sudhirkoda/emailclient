import atexit
import configparser
import json
import logging
import os
import re
import smtplib
import socket
import time
from datetime import datetime
from email.message import EmailMessage
from logging.handlers import RotatingFileHandler

import matplotlib.pyplot as plt
from apscheduler.schedulers.background import BackgroundScheduler
from flask import Flask, request, jsonify

import csv
from db import Database

rootpath = os.path.abspath(os.path.dirname(__file__))
app = Flask(__name__)

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
# Adding logging function
logfile = os.path.join(rootpath, "emailclient.log")
logger = logging.getLogger("emailClient")
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
handler = RotatingFileHandler(logfile, mode="a", maxBytes=10 * 1024 * 1024, backupCount=5)
handler.setLevel(logging.DEBUG)
handler.setFormatter(formatter)
logger.addHandler(handler)


# Adding config reader
def getconfig(section, key):
    try:
        configfile = os.path.join(rootpath, "emailclient.ini")
        if not os.path.isfile(configfile):
            logger.error(f" getconfig :: No config file or cant read config file - {configfile}")
            return ""
        configmanager = configparser.RawConfigParser()
        configmanager.read(configfile)
        return configmanager.get(section, key).strip()
    except Exception as error:
        logger.error(f" getconfig :: Error in getting configuration information - {error}")
        return ""


# Add scheduler timing
schedulertime = getconfig("MAIN", "statistics email time")
hour, mins = 23, 57
if schedulertime is not None or str(schedulertime).find(":") != -1:
    # noinspection PyBroadException
    try:
        hour = int(str(schedulertime).split(":")[0])
        mins = int(str(schedulertime).split(":")[1])
    except Exception:
        pass

# Adding scheduler
scheduler = BackgroundScheduler()


def isvalidemail(email):
    regex = '^(\w|\.|\_|\-)+[@](\w|\_|\-|\.)+[.]\w{2,3}$'
    if re.search(regex, email):
        return True
    return False


def flatenlist(inpuarg):
    if isinstance(inpuarg, list):
        return ",".join(inpuarg)
    return inpuarg


def updatetable(context: dict) -> (bool, str):
    insertsql = """
    INSERT INTO tbl_emailclient("from", "to", cc, bcc, subject, status, response, emailsenttime, "type")
    VALUES (TRIM(?), TRIM(?), TRIM(?), TRIM(?), TRIM(?), TRIM(?), TRIM(?), ?, TRIM(?))"""
    db = Database()

    params = [context.get("from"), flatenlist(context.get("to")), flatenlist(context.get("cc")),
              flatenlist(context.get("bcc")), context.get("subject"),
              context.get("status"), context.get("response"), context.get("emailsenttime"), context.get("type")]
    status, message = db.execute(sql=insertsql, param=params)
    return status, message


def sendemail(context: dict, msg=None) -> dict:
    module = "sendemail"
    logger.info(f" {module} :: sendemail func started")
    if not isinstance(context, dict):
        context["status"] = "NOK"
        context["response"] = "Error in retrieving JSON data"
        logger.error(f" {module} :: Error in retrieving JSON data")
        return context

    for value in ["from", "password"]:
        if context.get(value) is None or len(str(context[value]).strip()) == 0:
            context["status"] = "NOK"
            context["response"] = "Data is not complete to send email"
            logger.error(f" {module} :: Data is not complete to send email")
            return context

    toaddr = context.get("to")
    if isinstance(toaddr, str):
        context["to"] = [toaddr]

    try:
        if msg is None:
            message = EmailMessage()
            message['From'] = context.get("from")
            message['To'] = context.get("to")
            if context.get("cc") is not None:
                message['Cc'] = context.get("cc")
            if context.get("bcc") is not None:
                message['Bcc'] = context.get("bcc")
            message['MIME-Version'] = '1.0'
            message['Content-type'] = 'text/html'
            message['Subject'] = context.get("subject")
            message.set_content(context.get("body"))
            logger.debug(" {module} :: from-{fromaddr}".format(module=module, fromaddr=context.get("from")))
            logger.debug(" {module} :: to-{toaddr}".format(module=module, toaddr=context.get("to")))
            logger.debug(" {module} :: subject-{subject}".format(module=module, subject=context.get("subject")))
        else:
            message = msg

        logger.info(" {module} :: Email Type- {msgtype}".format(module=module, msgtype=context.get("type")))
        fromaddr = context.get("from")
        password = context.get("password")
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(fromaddr, password)
        server.send_message(message)
        server.close()
        context["emailsenttime"] = datetime.now()
        context["status"] = "OK"
        context["response"] = "Email sent"
    except Exception as error:
        context["emailsenttime"] = datetime.now()
        context["status"] = "NOK"
        context["response"] = str(error)
    finally:
        logger.info(" {0} :: Email sent status-{1} and response-{2}".format(module, context["status"],
                                                                            context["response"]))
        status, msg = updatetable(context)
        logger.info(" {0} :: Database update status-{1} and response-{2}".format(module, status,
                                                                                 msg))
        logger.info(f" {module} :: sendemail func end")
        return context


@app.route("/emailclient", methods=["POST", "GET"])
def emailclient():
    module = "emailclient"
    if request.method == "GET":
        ret = {
            "from": "",
            "password": "",
            "to": "",
            "cc": "",
            "bcc": "",
            "subject": "",
            "body": ""
        }
        return json.dumps(ret, indent=4)
    else:
        logger.info(f" {module} :: Received emailclient request")
        emailcontext = request.get_json()
        if emailcontext is None:
            logger.error(f" {module} :: Data not found in request")
            return jsonify({
                "Status": "NOK",
                "Response": "Data not found in request"
            })
        if emailcontext.get("from") is None:
            fromaddr = getconfig("MAIN", "from")
            password = getconfig("MAIN", "password")
            emailcontext["from"] = fromaddr
            emailcontext["password"] = password
            logger.info(f" {module} :: from address not received in request, using config file values")

        emailcontext["type"] = "normal"
        retdict = sendemail(emailcontext)
        logger.info(f" {module} :: End of emailclient request")
        return jsonify({
            "Status": retdict.get("status"),
            "Response": retdict.get("response")
        })


@app.route("/bulkemail", methods=["POST", "GET"])
def bulkemail():
    module = "bulkemail"
    if request.method == "GET":
        return ""
    else:
        logger.info(f" {module} :: Received of bulkemail request")
        bulkcsvfileobj = request.files.get('filename')
        if bulkcsvfileobj is None:
            logger.error(f" {module} :: No csv file available with the request")
            return jsonify({
                "Status": "NOK",
                "Response": "No csv file available with the request"
            })
        filename = bulkcsvfileobj.filename
        bulkcsvfileobj.save(filename)
        resultlist = []
        with open(filename) as bulkemailcsv:
            csvreader = csv.reader(bulkemailcsv)
            next(csvreader)
            for emaildetails in csvreader:
                emailcontext = {
                    "from": emaildetails[0],
                    "password": emaildetails[1],
                    "to": emaildetails[2],
                    "cc": emaildetails[3],
                    "bcc": emaildetails[4],
                    "subject": emaildetails[5],
                    "body": emaildetails[6],
                    "type": "bulk"
                }
                retdict = sendemail(emailcontext)
                retdict.pop("password", None)
                retdict.pop("body", None)
                resultlist.append(retdict)
        os.remove(filename)
        logger.info(f" {module} :: End of bulkemail request")
        return jsonify(resultlist)


@scheduler.scheduled_job('cron', id='sendstatics', hour=hour, minute=mins)
def sendstatics():
    module = "sendstatics"
    logger.info(f" {module} :: Scheduled email client notification started.")
    fromaddr = getconfig("MAIN", "from")
    if not isvalidemail(fromaddr):
        logger.error(f" {module} :: from email address configured \"{fromaddr}\" is not valid, Exit sendstatics.")
        return

    password = getconfig("MAIN", "password")
    if password is None:
        logger.error(f" {module} :: password configured \"{password}\" is not valid, Exit sendstatics.")
        return

    toaddr = getconfig("MAIN", "admin email")
    if not isvalidemail(toaddr):
        logger.error(f" {module} :: admin email address configured \"{toaddr}\" is not valid, Exit sendstatics.")
        return

    staticsdate = datetime.now().strftime("%Y-%m-%d")
    hourdatasql = f"""SELECT COUNT(Id), strftime ('%H',emailsenttime) hour FROM tbl_emailclient 
    WHERE emailsenttime like "{staticsdate}%"
    GROUP BY strftime ('%H',emailsenttime)"""

    db = Database()
    dbreturnvalue = db.execute(sql=hourdatasql)
    hourcountmapdict = {}
    if dbreturnvalue[0]:
        for hourcount in dbreturnvalue[1]:
            hourcountmapdict[int(hourcount[1])] = hourcount[0]

    countvalues = [0 for _ in range(24)]
    for keys, values in hourcountmapdict.items():
        countvalues[keys] = values

    plt.style.use('ggplot')
    fig = plt.figure()
    ax = fig.add_axes([0.1, 0.1, 0.75, 0.75])
    ax.set_title("Day Email Count Stats")
    ax.set_xlabel("Time (Hours)")
    ax.set_ylabel("Email Count")
    ax.bar([_ for _ in range(24)], countvalues, 1, color="blue")

    rects = ax.patches
    for rect, label in zip(rects, countvalues):
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width() / 2, height, label,
                ha='center', va='bottom')

    fig.savefig("fig.png")
    time.sleep(1)

    from email.utils import make_msgid
    import mimetypes
    image_cid = make_msgid()

    html_template = """
    <html lang="en">
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>Email Client Statistics Report</title>
    <style type="text/css">
    a {color: #d80a3e;}
    body, #header h1, #header h2, p {margin: 0; padding: 0;}
    #main {border: 1px solid #cfcece;}
    #top-message p, #bottom p {color: #3f4042; font-size: 12px; font-family: Arial, Helvetica, sans-serif; }
    #header h1 {color: #ffffff !important; font-family: "Lucida Grande", sans-serif; font-size: 24px; margin-bottom: 0!important; padding-bottom: 0; }
    #header p {color: #ffffff !important; font-family: "Lucida Grande", "Lucida Sans", "Lucida Sans Unicode", sans-serif; font-size: 12px;  }
    h5 {margin: 0 0 0.8em 0;}
    h5 {font-size: 18px; color: #444444 !important; font-family: Arial, Helvetica, sans-serif; }
    p {font-size: 12px; color: #444444 !important; font-family: "Lucida Grande", "Lucida Sans", "Lucida Sans Unicode", sans-serif; line-height: 1.5;}
    </style>
    </head>
    <body>
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="e4e4e4"><tr><td>
    <table id="main" width="600" align="center" cellpadding="0" cellspacing="15" bgcolor="ffffff">
    <tr>
    <td>
    <table id="header" cellpadding="10" cellspacing="0" align="center" bgcolor="8fb3e9">
        <tr>
            <td width="570" align="center"  bgcolor="#d80a3e"><h1>Email Client Statistics Report</h1></td>
        </tr>
    </table>
    </td>
    </tr>
    <tr>
    <td>
    <table id="content-4" cellpadding="10" cellspacing="0" align="left">
        <tr>
            <td width="600" valign="top">
            <h5>Hello Admin</h5>
            <p>Good Day, Email Client Statistics for date """ + str(staticsdate) + """.</p>
             </br>
                <img src="cid:""" + str(image_cid[1:-1]) + """">
                </br>
                <p>Thanks!,</p>
                <p>Team Email Client</p>
            </td>
        </tr>
    </table>
    </td>
    </tr>
    </td></tr></table><!-- wrapper -->
    </table>
    </body>
    </html>"""

    emailcontext = {
        "from": fromaddr,
        "password": password,
        "to": toaddr,
        "cc": None,
        "bcc": None,
        "subject": f"Email Client Statistic Report {staticsdate}",
        "body": html_template,
        "type": "admin"
    }

    msg = EmailMessage()
    msg['From'] = emailcontext.get("from")
    msg['To'] = emailcontext.get("to")
    msg['Subject'] = f"Email Client Statistic Report {staticsdate}"

    msg.set_content('Email Client Statistic Report.')

    msg.add_alternative(html_template, subtype='html')
    with open('fig.png', 'rb') as img:
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
        msg.get_payload()[1].add_related(img.read(),
                                         maintype=maintype,
                                         subtype=subtype,
                                         cid=image_cid)

    sendemail(emailcontext, msg=msg)
    logger.info(f" {module} :: Scheduled email client notification end.")


# noinspection PyBroadException
@scheduler.scheduled_job('cron', id='cleandbdata', hour=14, minute=14)
def cleandbdata():
    module = "cleandbdata"
    logger.info(f" {module} :: Database cleanup started")
    try:
        days = int(getconfig("MAIN", "backup days"))
    except Exception:
        days = 180
        logger.info(f" {module} :: Database backup days not configured, using default {days} days")

    deletesql = f"DELETE FROM tbl_emailclient WHERE emailsenttime <= date('now','-{days} day')"
    db = Database()
    retvalues = db.execute(deletesql)
    if retvalues[0]:
        logger.info(f" {module} :: Database cleanup completed")
        return
    logger.info(f" {module} :: Database cleanup failed, response-{str(retvalues[1])}")


# noinspection PyBroadException
def gethostip():
    sok = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sok.connect(("10.255.255.255", 1))
        ip = sok.getsockname()[0]
    except Exception:
        sok.connect(("8.8.8.8", 80))
        ip = sok.getsockname()[0]
    finally:
        sok.close()
    return ip


try:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(("127.0.0.1", 47200))
except socket.error:
    pass
else:
    scheduler.start()

atexit.register(lambda: scheduler.shutdown())
logger.info(f" MAIN :: {scheduler.get_jobs()}")

if __name__ == "__main__":
    app.run(host=gethostip(), port="5001")
