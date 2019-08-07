import pandas as pd
import pyodbc
import datetime
from openpyxl import Workbook

cnxn = pyodbc.connect("DSN=L5QNS_PROD;UID=;DBQ=TNS:L5QNS_PROD.WORLD;ASY=OFF")
cursor = cnxn.cursor()


today = datetime.date.today()
idx = (today.weekday() + 1) % 7
# sun = today - datetime.timedelta(7 + idx)
# sat = today - datetime.timedelta(7 + idx - 7)

#sat_date = today - datetime.timedelta(7 + idx - 6)
current_sun = today - datetime.timedelta(idx)
current_sat = current_sun + datetime.timedelta(6)


department_shiftmanagers = """SELECT SAFETY_CARDS.SC_NUMBER, SAFETY_CARDS.LAST_NAME, SAFETY_CARDS.FIRST_NAME, SAFETY_CARDS.DEPT_ID
FROM SCDB.SAFETY_CARDS SAFETY_CARDS
WHERE (SAFETY_CARDS.DEPT_ID In ('MWOP')) AND (SAFETY_CARDS.EMPLOYEE_STATUS_CODE='01') AND (SAFETY_CARDS.EMPLOYEE_TYPE_CODE In ('B', 'C'))  AND (SAFETY_CARDS.MANAGER_TYPE_ID In (5, 8))
ORDER BY SAFETY_CARDS.DEPT_ID, SAFETY_CARDS.LAST_NAME, SAFETY_CARDS.FIRST_NAME"""

shiftmanagers_return = pd.read_sql(department_shiftmanagers, cnxn)
shift_managers_all = shiftmanagers_return.T.to_dict().values()

employee_schedule_sql = """SELECT SAFETY_CARDS.SC_NUMBER, SAFETY_CARDS.LAST_NAME, SAFETY_CARDS.FIRST_NAME, SAFETY_CARDS.DEPT_ID, EMPLOYEE_SCHEDULE.SHIFT_BEGIN_DATE
FROM SCDB.EMPLOYEE_SCHEDULE EMPLOYEE_SCHEDULE, SCDB.SAFETY_CARDS SAFETY_CARDS
WHERE SAFETY_CARDS.SC_NUMBER = EMPLOYEE_SCHEDULE.SC_NUMBER 
AND ((SAFETY_CARDS.DEPT_ID In ('MWOP')) 
AND (SAFETY_CARDS.EMPLOYEE_STATUS_CODE='01') 
AND (SAFETY_CARDS.EMPLOYEE_TYPE_CODE In ('B', 'C'))  
AND (SAFETY_CARDS.MANAGER_TYPE_ID In (5, 8)) 
AND (EMPLOYEE_SCHEDULE.SHIFT_BEGIN_DATE 
Between {ts '""" + str(current_sun) + """ 00:00:00'} And {ts '""" + str(current_sat) + """ 00:00:00'}))
ORDER BY SAFETY_CARDS.LAST_NAME, SAFETY_CARDS.FIRST_NAME"""


shiftmanager_schedule = pd.read_sql(employee_schedule_sql, cnxn)
schedule = shiftmanager_schedule.T.to_dict().values()


weekly_lists = []
daily_alerts_list = []
next_weekly_lists = []
groupings = [['MWOP']]
for grouping in groupings:
    shift_managers = [x for x in shift_managers_all if x['DEPT_ID'] in grouping]
    sm_list = [[x['LAST_NAME'], x['SC_NUMBER']] for x in schedule if x['DEPT_ID'] in grouping]

    current_week = [current_sun + datetime.timedelta(days = x) for x in range((current_sat - current_sun).days + 1)]
    current_week = [b for b in current_week if b <= today]


    header = ['SC#', 'Shift Manager', 'Department'] + current_week
    weekly_list = []
    weekly_list.append(header)

    for sm in shift_managers:
        sm['SCHEDULED'] = {}
        sm['OBSERVATIONS'] = {}
        sm['TOOLBOX'] = {}

    for day in current_week:
        for idx, sm in enumerate(shift_managers):
            if any([x for x in schedule if (x['SHIFT_BEGIN_DATE'].date() == day) and x['SC_NUMBER'] == sm['SC_NUMBER']]):
                shift_managers[idx]['SCHEDULED'][str(day)] = 2

    for idx, sm in enumerate(shift_managers):
        sm_obs_sql = """SELECT SAFETY_OBSERVATIONS.OBSERVED_DATE
        FROM SCDB.SAFETY_OBSERVATIONS SAFETY_OBSERVATIONS, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
        WHERE SC_OBSERVED_BY.OBS_BY_ID = SAFETY_OBSERVATIONS.OBS_BY_ID 
        AND ((SC_OBSERVED_BY.SC_NUMBER=""" + str(int(sm['SC_NUMBER'])) + """) 
        AND (SAFETY_OBSERVATIONS.OBSERVED_DATE 
        Between {ts '""" + str(current_sun) + """ 00:00:00'} 
        And {ts '""" + str(current_sat) + """ 00:00:00'}))
        ORDER BY SAFETY_OBSERVATIONS.OBSERVED_DATE"""
        sm_obs = cnxn.execute(sm_obs_sql).fetchall()
        obs_date_list = []
        for b in sm_obs:
            obs_date_list.append((b[0].date()))
        for day in current_week:
            total_obs = len([x for x in obs_date_list if x == day])
            shift_managers[idx]['OBSERVATIONS'][str(day)] = total_obs

        sm_tb_sql = """SELECT SC_MEETING_DETAIL.MEETING_DATE
        FROM SCDB.SC_MEETING_DETAIL SC_MEETING_DETAIL, SCDB.SC_OBSERVED_BY SC_OBSERVED_BY
        WHERE SC_OBSERVED_BY.OBS_BY_ID = SC_MEETING_DETAIL.OBS_BY_ID 
        AND ((SC_OBSERVED_BY.SC_NUMBER=""" + str(int(sm['SC_NUMBER'])) + """) 
        AND (SC_MEETING_DETAIL.MEETING_DATE Between {ts '""" + str(current_sun) + """ 00:00:00'} 
        And {ts '""" + str(current_sat) + """ 00:00:00'}) 
        AND (SC_MEETING_DETAIL.MEETING_TYPE_ID=64))
        ORDER BY SC_MEETING_DETAIL.MEETING_DATE DESC"""

        sm_tb = cnxn.execute(sm_tb_sql).fetchall()
        obs_date_list = []
        for b in sm_tb:
            obs_date_list.append((b[0].date()))
        for day in current_week:
            total_obs = len([x for x in obs_date_list if x == day])
            shift_managers[idx]['TOOLBOX'][str(day)] = total_obs


    for sm in shift_managers:
        sc_number = int(sm['SC_NUMBER'])
        name = "{}".format(sm['LAST_NAME'])
        department = sm['DEPT_ID']
        sm_line = []
        sm_line.append(sc_number)
        sm_line.append(name)
        sm_line.append(department)
        for day in current_week:
            try:
                if sm['SCHEDULED'][str(day)] == 2:
                    if sm['OBSERVATIONS'][str(day)] < 2:
                        # sm_line.append('X')
                        if sm['TOOLBOX'][str(day)] > 0:
                            sm_line.append('PX')
                        else:
                            sm_line.append('XX')
                    else:
                        if sm['TOOLBOX'][str(day)] > 0:
                            sm_line.append('P')
                        else:
                            sm_line.append('XP')
                        # sm_line.append('P')
                else:
                    sm_line.append('_')
            except:
                sm_line.append('_')
        weekly_list.append(sm_line)
    for b in weekly_list:
        print(b)
    next_sat = current_sat + datetime.timedelta(days=7)
    next_sun = current_sun + datetime.timedelta(days=7)
    next_week = [next_sun + datetime.timedelta(days = x) for x in range((next_sat - next_sun).days + 1)]
    header = ['SC#', 'Shift Manager', 'Department'] + next_week
    next_weekly_list = []
    next_weekly_list.append(header)


    next_week_employee_schedule_sql = """SELECT SAFETY_CARDS.SC_NUMBER, SAFETY_CARDS.LAST_NAME, SAFETY_CARDS.FIRST_NAME, SAFETY_CARDS.DEPT_ID, EMPLOYEE_SCHEDULE.SHIFT_BEGIN_DATE
    FROM SCDB.EMPLOYEE_SCHEDULE EMPLOYEE_SCHEDULE, SCDB.SAFETY_CARDS SAFETY_CARDS
    WHERE SAFETY_CARDS.SC_NUMBER = EMPLOYEE_SCHEDULE.SC_NUMBER 
    AND ((SAFETY_CARDS.DEPT_ID In ('MWOP')) 
    AND (SAFETY_CARDS.EMPLOYEE_STATUS_CODE='01') 
    AND (SAFETY_CARDS.EMPLOYEE_TYPE_CODE In ('B', 'C')) 
    AND (SAFETY_CARDS.MANAGER_TYPE_ID In (5, 8)) 
    AND (EMPLOYEE_SCHEDULE.SHIFT_BEGIN_DATE 
    Between {ts '""" + str(next_sun) + """ 00:00:00'} And {ts '""" + str(next_sat) + """ 00:00:00'}))
    ORDER BY SAFETY_CARDS.LAST_NAME, SAFETY_CARDS.FIRST_NAME"""


    next_week_shiftmanager_schedule = pd.read_sql(next_week_employee_schedule_sql, cnxn)
    next_week_schedule = next_week_shiftmanager_schedule.T.to_dict().values()


    for sm in shift_managers:
        sm['SCHEDULED'] = {}

    for day in next_week:
        for idx, sm in enumerate(shift_managers):
            if any([x for x in next_week_schedule if (x['SHIFT_BEGIN_DATE'].date() == day) and x['SC_NUMBER'] == sm['SC_NUMBER']]):
                shift_managers[idx]['SCHEDULED'][str(day)] = 2

    for sm in shift_managers:
        sc_number = int(sm['SC_NUMBER'])
        name = "{}".format(sm['LAST_NAME'])
        department = sm['DEPT_ID']
        sm_line = []
        sm_line.append(sc_number)
        sm_line.append(name)
        sm_line.append(department)
        for day in next_week:
            try:
                if sm['SCHEDULED'][str(day)] > 1:
                    sm_line.append('P')
                else:
                    sm_line.append('_')
            except:
                sm_line.append('_')
        next_weekly_list.append(sm_line)
    for b in next_weekly_list:
        print(b)

    daily_alerts = []
    next_week_sms = []
    for day in next_week:
        daily_sms = sum([1 for x in next_week_schedule if x['SHIFT_BEGIN_DATE'].date() == day])
        if daily_sms == 0:
            daily_message = """No shift managers for {}""".format(str(day))
            daily_alerts.append(daily_message)
        next_week_sms.append([day, daily_sms])
    weekly_lists.append(weekly_list)
    next_weekly_lists.append(next_weekly_list)
    daily_alerts_list.append(daily_alerts)

#### SEND THE EMAIL
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders

# me == my email address
# you == recipient's email address

me = "chris.schottmiller@aksteel.com"
you = "chris.schottmiller@aksteel.com"
you_df = pd.read_excel(r'S:\QualityCI\Continuous Improvement\South Annealing\safety_email_lists\bof.xlsx', sheet_name='bof_mtc')
you = you_df.shift_manager_list.tolist()
you = [y for y in you if pd.isnull(y) == False]# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = "SM Observations"
msg['From'] = me
msg['To'] = ", ".join(you)
# Create the body of the message (a plain-text and an HTML version).
text = "SM OBS"
html1 = """\
<html>
  <head></head>
  <body>"""
html2 = ""
for weekly_list in weekly_lists:
    for idx, line in enumerate(weekly_list):
        if idx == 0:
            html2 = html2 + """<table width="800" border="1">
      <tr>"""
            for val in line:
                html2 = html2+ """<th>""" + str(val) + """</th>"""
            html2 = html2 + """</tr>"""
        elif idx > 0:
            html2 = html2 + """<tr>"""
            for val in line:
                html2 = html2 + """<td align="center">""" + str(val) + """</td>"""
            html2 = html2 + """</tr>"""
    html2 = html2 + """</table><br><br>"""
html3 = """
Key:<br>
P : Both Objectives Complete<br>
XP : Toolbox Meeting Not Entered, Observation Target Met<br>
PX : Toolbox Meeting Entered, Observation Target Not Met <br><br>
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
s.sendmail(me, you, msg.as_string())
s.quit()

for idx, next_weekly_list in enumerate(next_weekly_lists):
    daily_alerts = daily_alerts_list[idx]
    if all([len(daily_alerts) > 0, datetime.datetime.today().isoweekday() >= 1]):
        #### SEND THE EMAIL
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.MIMEMultipart import MIMEMultipart
        from email.MIMEBase import MIMEBase
        from email import Encoders

        # me == my email address
        # you == recipient's email address
        me = "chris.schottmiller@aksteel.com"
        you = "chris.schottmiller@aksteel.com"
        you_df = pd.read_excel(r'S:\QualityCI\Continuous Improvement\South Annealing\safety_email_lists\bof.xlsx', sheet_name='bof_mtc')
        you = you_df.next_week_schedule_list.tolist()
        you = [y for y in you if pd.isnull(y) == False]
        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "SM Schedule Next Week"
        msg['From'] = me
        msg['To'] = ", ".join(you)
        # Create the body of the message (a plain-text and an HTML version).
        text = "SM OBS"
        html1 = """\
        <html>
          <head></head>
          <body>"""
        html2 = ""
        for idx, line in enumerate(daily_alerts):
            html2 = html2 + """<b> {} </b><br>""".format(str(line))
        html2 = html2 + """<br><br>Next Week's Schedule In Safety Card"""
        for idx, line in enumerate(next_weekly_list):
            if idx == 0:
                html2 = html2 + """<table width="1200" border="1">
          <tr>"""
                for val in line:
                    html2 = html2+ """<th>""" + str(val) + """</th>"""
                html2 = html2 + """</tr>"""
            elif idx > 0:
                html2 = html2 + """<tr>"""
                for val in line:
                    html2 = html2 + """<td align="center">""" + str(val) + """</td>"""
                html2 = html2 + """</tr>"""
        html2 = html2 + """</table>"""
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
        s.sendmail(me, you, msg.as_string())
        s.quit()