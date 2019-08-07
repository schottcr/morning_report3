from dateutil.relativedelta import relativedelta
import datetime
import pyodbc


cnxn = pyodbc.connect("DSN=L5QNS_PROD;UID=;DBQ=TNS:L5QNS_PROD.WORLD;ASY=OFF")
cursor = cnxn.cursor()


today = datetime.date.today()

start_month_temp = today - relativedelta(days = (today.day - 1))
end_month_temp = (today + relativedelta(months = 1))
end_month_temp = datetime.date(end_month_temp.year, end_month_temp.month, 1)


# end_month_temp = (today + relativedelta(months = 1)) - relativedelta(days = today.day - 1)
#idx = (today.weekday() + 1) % 7

start_month = datetime.datetime.combine(start_month_temp, datetime.datetime.min.time())
end_month = datetime.datetime.combine(end_month_temp, datetime.datetime.min.time())

engineers = []
emp = {}
emp['Name'] = 'Chris Schottmiller'
emp['SC_Num'] = 55268
emp['Email'] = 'Chris.Schottmiller@aksteel.com'
# emp['ToEmail'] = ['schottcr@gmail.com']
emp['ToEmail'] = ['Chris.Schottmiller@aksteel.com']
# emp['ToEmail'] = ['Chris.Schottmiller@aksteel.com', 'Sal.Ciriello@aksteel.com']
engineers.append(emp)

emp = {}
emp['Name'] = 'Sal Ciriello'
emp['SC_Num'] = 8212
emp['Email'] = 'Sal.Ciriello@aksteel.com'
# emp['ToEmail'] = ['Chris.Schottmiller@aksteel.com']
emp['ToEmail'] = ['Sal.Ciriello@aksteel.com']
engineers.append(emp)


emp = {}
emp['Name'] = 'Aaron Boyer'
emp['SC_Num'] = 58068
emp['Email'] = 'Aaron.Boyer@aksteel.com'
# emp['ToEmail'] = ['Chris.Schottmiller@aksteel.com']
emp['ToEmail'] = ['Aaron.Boyer@aksteel.com']
engineers.append(emp)


emp = {}
emp['Name'] = 'Josh Hartman'
emp['SC_Num'] = 69068
emp['Email'] = 'Josh.Hartman@aksteel.com'
# emp['ToEmail'] = ['Chris.Schottmiller@aksteel.com']
emp['ToEmail'] = ['Josh.Hartman@aksteel.com']
engineers.append(emp)



for emp in engineers:
    # print emp
    sql = """SELECT SC_OBSERVATION_DETAIL.OBSERVED_DATE, SC_OBSERVATION_DETAIL.DAILY_SEQ_NUMBER, SC_OBSERVATION_DETAIL.DEPT_ID, SC_OBSERVATION_DETAIL.GROUP_ID, SC_OBSERVATION_DETAIL.OBS_STATUS_CODE, SC_OBSERVATION_DETAIL.OBS_TOPIC, SC_OBSERVATION_DETAIL.REVIEW_DOC_TYPE_DESC, SC_OBSERVATION_DETAIL.DOC_TITLE, SC_OBSERVATION_DETAIL.STEPS_COMPLETED, SC_OBSERVATION_DETAIL.OBS_CATEGORY_ID, SC_OBSERVATION_DETAIL.CORRECTIVE_ACTION_ID, SC_OBSERVATION_DETAIL.OBS_CATEGORY_TYPE_ID, SC_OBSERVATION_DETAIL.OBS_BY_ID, SC_OBSERVATION_DETAIL.WEEK_NUMBER, SC_OBSERVATION_DETAIL.INITIATED_BY_INCIDENT_IND, SC_OBSERVATION_DETAIL.UPDATE_LOG_DATE, SC_OBSERVATION_DETAIL.UPDATE_LOG_USER_ID, SC_OBSERVATION_DETAIL.UPDATE_LOG_PROGRAM_ID, SC_OBSERVATION_DETAIL.DOC_TYPE_NUMBER, SC_OBSERVATION_DETAIL.OBS_NATURE_ID, SC_OBSERVATION_DETAIL.AUDITOR_OBS_BY_ID, SC_OBSERVATION_DETAIL.TARGETED_OBSERVATION_IND, SC_OBSERVATION_DETAIL.OBS_GRADE_ID, SC_OBSERVATION_DETAIL.LOTO_OBSERVATION_IND, SC_OBSERVATION_DETAIL.EMPLOYEE_VIOLATION_EXPLANATION, SC_OBSERVATION_DETAIL.OBS_EMPLOYEE_CONTACTED_IND, SC_OBSERVATION_DETAIL.OBS_EMPL_NO_CONTACT_REASON
    FROM SCDB.SAFETY_CARDS SAFETY_CARDS, SCDB.SC_OBSERVATION_DETAIL SC_OBSERVATION_DETAIL, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
    WHERE SAFETY_CARDS.SC_NUMBER = SC_OBSERVED_BY.SC_NUMBER AND SC_OBSERVED_BY.OBS_BY_ID = SC_OBSERVATION_DETAIL.OBS_BY_ID AND ((SAFETY_CARDS.SC_NUMBER=""" + str(emp['SC_Num']) + """) AND (SC_OBSERVATION_DETAIL.OBSERVED_DATE Between {ts '""" + str(start_month) + """'} And {ts '""" + str(end_month) + """'}) AND (SC_OBSERVATION_DETAIL.OBS_NATURE_ID=1))"""
    obs_return = cursor.execute(sql).fetchall()
    emp['Observations'] = len(obs_return)

# for emp in engineers:
#     print emp['Name'], emp['Observations']


for emp in engineers:
    sql = """SELECT SC_MEETING_DETAIL.MEETING_DATE, SC_MEETING_DETAIL.DAILY_SEQ_NUMBER, SC_MEETING_DETAIL.DEPT_ID, SC_MEETING_DETAIL.MEETING_TYPE_ID, SC_MEETING_DETAIL.MEETING_DESC, SC_MEETING_DETAIL.HELD_BY_INITIALS, SC_MEETING_DETAIL.OBS_BY_ID, SC_MEETING_DETAIL.MEETING_LEADER_FIRST_NAME, SC_MEETING_DETAIL.MEETING_LEADER_LAST_NAME, SC_MEETING_DETAIL.UPDATE_LOG_DATE, SC_MEETING_DETAIL.UPDATE_LOG_USER_ID, SC_MEETING_DETAIL.UPDATE_LOG_PROGRAM_ID, SC_MEETING_DETAIL.AUDITOR_OBS_BY_ID
    FROM SCDB.SAFETY_CARDS SAFETY_CARDS, SCDB.SC_MEETING_DETAIL SC_MEETING_DETAIL, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
    WHERE SAFETY_CARDS.SC_NUMBER = SC_OBSERVED_BY.SC_NUMBER AND SC_OBSERVED_BY.OBS_BY_ID = SC_MEETING_DETAIL.AUDITOR_OBS_BY_ID AND ((SAFETY_CARDS.SC_NUMBER=""" + str(emp['SC_Num']) + """) AND (SC_MEETING_DETAIL.MEETING_DATE Between {ts '""" + str(start_month) + """'} And {ts '""" + str(end_month) + """'}) AND (SC_MEETING_DETAIL.MEETING_TYPE_ID=64))"""
    toolbox_return = cursor.execute(sql).fetchall()

    sql2 = """SELECT SC_MEETING_DETAIL_AUDITORS.MEETING_DATE, SC_MEETING_DETAIL_AUDITORS.DAILY_SEQ_NUMBER, SC_MEETING_DETAIL_AUDITORS.DEPT_ID, SC_MEETING_DETAIL_AUDITORS.AUDITOR_OBS_BY_ID, SC_MEETING_DETAIL_AUDITORS.UPDATE_LOG_DATE, SC_MEETING_DETAIL_AUDITORS.UPDATE_LOG_USER_ID, SC_MEETING_DETAIL_AUDITORS.UPDATE_LOG_PROGRAM_ID
    FROM SCDB.SAFETY_CARDS SAFETY_CARDS, SCDB.SC_MEETING_DETAIL SC_MEETING_DETAIL, SCDB.SC_MEETING_DETAIL_AUDITORS SC_MEETING_DETAIL_AUDITORS, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
    WHERE SAFETY_CARDS.SC_NUMBER = SC_OBSERVED_BY.SC_NUMBER AND SC_OBSERVED_BY.OBS_BY_ID = SC_MEETING_DETAIL_AUDITORS.AUDITOR_OBS_BY_ID AND SC_MEETING_DETAIL_AUDITORS.MEETING_DATE = SC_MEETING_DETAIL.MEETING_DATE AND SC_MEETING_DETAIL_AUDITORS.DEPT_ID = SC_MEETING_DETAIL.DEPT_ID AND SC_MEETING_DETAIL_AUDITORS.DAILY_SEQ_NUMBER = SC_MEETING_DETAIL.DAILY_SEQ_NUMBER AND ((SAFETY_CARDS.SC_NUMBER=""" + str(emp['SC_Num']) + """) AND (SC_MEETING_DETAIL.MEETING_TYPE_ID=64) AND (SC_MEETING_DETAIL.MEETING_DATE Between {ts '""" + str(start_month) + """'} And {ts '""" + str(end_month) + """'}))
    ORDER BY SC_MEETING_DETAIL_AUDITORS.MEETING_DATE DESC"""
    toolbox_return_multi = cursor.execute(sql2).fetchall()

    emp['Toolbox'] = len(toolbox_return) + len(toolbox_return_multi)




# New Safety Meeting Method
for emp in engineers:
    sql = """SELECT SC_MEETING_DETAIL.MEETING_DATE, SC_MEETING_DETAIL.DAILY_SEQ_NUMBER, SC_MEETING_DETAIL.DEPT_ID, SC_MEETING_DETAIL.MEETING_TYPE_ID, SC_MEETING_DETAIL.MEETING_DESC, SC_MEETING_DETAIL.HELD_BY_INITIALS, SC_MEETING_DETAIL.OBS_BY_ID, SC_MEETING_DETAIL.MEETING_LEADER_FIRST_NAME, SC_MEETING_DETAIL.MEETING_LEADER_LAST_NAME, SC_MEETING_DETAIL.UPDATE_LOG_DATE, SC_MEETING_DETAIL.UPDATE_LOG_USER_ID, SC_MEETING_DETAIL.UPDATE_LOG_PROGRAM_ID, SC_MEETING_DETAIL.AUDITOR_OBS_BY_ID
    FROM SCDB.SAFETY_CARDS SAFETY_CARDS, SCDB.SC_MEETING_DETAIL SC_MEETING_DETAIL, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
    WHERE SAFETY_CARDS.SC_NUMBER = SC_OBSERVED_BY.SC_NUMBER AND SC_OBSERVED_BY.OBS_BY_ID = SC_MEETING_DETAIL.AUDITOR_OBS_BY_ID AND ((SAFETY_CARDS.SC_NUMBER=""" + str(emp['SC_Num']) + """) AND (SC_MEETING_DETAIL.MEETING_DATE Between {ts '""" + str(start_month) + """'} And {ts '""" + str(end_month) + """'}) AND (SC_MEETING_DETAIL.MEETING_TYPE_ID=67))"""
    safety_meeting_return = cursor.execute(sql).fetchall()

    sql2 = """SELECT SC_MEETING_DETAIL_AUDITORS.MEETING_DATE, SC_MEETING_DETAIL_AUDITORS.DAILY_SEQ_NUMBER, SC_MEETING_DETAIL_AUDITORS.DEPT_ID, SC_MEETING_DETAIL_AUDITORS.AUDITOR_OBS_BY_ID, SC_MEETING_DETAIL_AUDITORS.UPDATE_LOG_DATE, SC_MEETING_DETAIL_AUDITORS.UPDATE_LOG_USER_ID, SC_MEETING_DETAIL_AUDITORS.UPDATE_LOG_PROGRAM_ID
    FROM SCDB.SAFETY_CARDS SAFETY_CARDS, SCDB.SC_MEETING_DETAIL SC_MEETING_DETAIL, SCDB.SC_MEETING_DETAIL_AUDITORS SC_MEETING_DETAIL_AUDITORS, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
    WHERE SAFETY_CARDS.SC_NUMBER = SC_OBSERVED_BY.SC_NUMBER AND SC_OBSERVED_BY.OBS_BY_ID = SC_MEETING_DETAIL_AUDITORS.AUDITOR_OBS_BY_ID AND SC_MEETING_DETAIL_AUDITORS.MEETING_DATE = SC_MEETING_DETAIL.MEETING_DATE AND SC_MEETING_DETAIL_AUDITORS.DEPT_ID = SC_MEETING_DETAIL.DEPT_ID AND SC_MEETING_DETAIL_AUDITORS.DAILY_SEQ_NUMBER = SC_MEETING_DETAIL.DAILY_SEQ_NUMBER AND ((SAFETY_CARDS.SC_NUMBER=""" + str(emp['SC_Num']) + """) AND (SC_MEETING_DETAIL.MEETING_TYPE_ID=67) AND (SC_MEETING_DETAIL.MEETING_DATE Between {ts '""" + str(start_month) + """'} And {ts '""" + str(end_month) + """'}))
    ORDER BY SC_MEETING_DETAIL_AUDITORS.MEETING_DATE DESC"""
    safety_meeting_return_multi = cursor.execute(sql2).fetchall()

    emp['Safety_Meeting'] = len(safety_meeting_return) + len(safety_meeting_return_multi)


for emp in engineers:
    print emp['Name'], emp['Observations'], emp['Toolbox'], emp['Safety_Meeting']

# sun = today - datetime.timedelta(7 + idx)
# sat = today - datetime.timedelta(7 + idx - 7)


#### SEND THE EMAIL
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders

cc_address = ['sal.ciriello@aksteel.com']
if today.day > 15 or today.weekday() == 0:
    for emp in engineers:
        todo = []
        # Commented out 2/20/19 due to change in monthly requirements from 2 planned observations to 0 for engineers.
        # if emp['Observations'] < 2:
        #     todo.append('You have {} planned observation in for the month. The requirement is 2.'.format(emp['Observations']))
        if emp['Toolbox'] < 2:
            todo.append('You have {} audited toolbox meetings in for the month. The requirement is 2.'.format(emp['Toolbox']))
        if emp['Safety_Meeting'] < 1:
            todo.append('You have {} audited safety meetings in for the month. The requirement is 1.'.format(emp['Safety_Meeting']))
        if len(todo) > 0:
            # me == my email address
            # you == recipient's email address
            me = "chris.schottmiller@aksteel.com"
            you = emp['ToEmail']
            # Create message container - the correct MIME type is multipart/alternative.
            msg = MIMEMultipart('alternative')
            msg['Subject'] = "{} Monthly Numbers".format(emp['Name'])
            msg['From'] = me
            msg['To'] = ", ".join(you)
            msg['CC'] = ", ".join(cc_address)
            # Create the body of the message (a plain-text and an HTML version).
            text = "Engineer Numbers"
            html1 = """\
            <html>
              <head></head>
              <body>"""
            html2 = ""
            for val in todo:
                html2 = html2 + """<br>""" + str(val)
            html2 = html2
            html3 = """
            </body>
            </html>
            """
            html = html1 + html2 + html3
            # Record the MIME types of both parts - text/plain and text/html.
            #part1 = MIMEText(text, 'plainA')
            part2 = MIMEText(html, 'html')
            # Attach parts into message container.
            # According to RFC 2046, the last part of a multipart message, in this case
            # the HTML message, is best and preferred.
            #msg.attach(part1)
            msg.attach(part2)
            # Send the message via local SMTP server.
            s = smtplib.SMTP('smtp.akst.com')
            # sendmail function takes 3 arguments: sender's address, recipient's address
            # and message to send - here it is sent as one string.
            s.sendmail(me, you+cc_address, msg.as_string())
            s.quit()
else:
    print('not in range')