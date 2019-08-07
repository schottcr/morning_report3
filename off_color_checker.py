from dateutil.relativedelta import relativedelta
import os
import datetime
import xlwings as xw

todays_date = datetime.datetime.now().date()
yesterdays_date = todays_date - relativedelta(days=1)

yesterdays_month = yesterdays_date.strftime("%B")

yesterdays_year = yesterdays_date.strftime("%Y")

yesterday_year_folder = "\Annealing Turn Reports {}".format(yesterdays_year)
yesterday_month_folder = "\{} Turn Report {}".format(yesterdays_month, yesterdays_year)

yesterdays_dir = "S:\Annealing\94 Annealing Turn Reports" + yesterday_year_folder + yesterday_month_folder

todays_month = todays_date.strftime("%B")
todays_year = todays_date.strftime("%Y")

todays_year_folder = "\Annealing Turn Reports {}".format(todays_year)
todays_month_folder = "\{} Turn Report {}".format(todays_month, todays_year)

todays_dir = "S:\Annealing\94 Annealing Turn Reports" + yesterday_year_folder + yesterday_month_folder

files = []

t2 = yesterdays_dir + "\{} {} {} T2.xlsb".format(yesterdays_date.strftime("%m"), yesterdays_date.strftime("%d"), yesterdays_date.strftime("%y"))
t3 = yesterdays_dir + "\{} {} {} T3.xlsb".format(yesterdays_date.strftime("%m"), yesterdays_date.strftime("%d"), yesterdays_date.strftime("%y"))
t1 = todays_dir + "\{} {} {} T1.xlsb".format(todays_date.strftime("%m"), todays_date.strftime("%d"), todays_date.strftime("%y"))

files.append(t2)
files.append(t3)
files.append(t1)

off_color = []
oc_headers = ['TurnReport', 'Base', 'Cover', 'Flopper', '# Coils', 'Color', 'Tons' ,'Charge', 'Comments']
off_color.append(oc_headers)
for fname in files:
    # print fname
    wb = xw.Book(fname)
    # print wb.sheets
    sht = wb.sheets('TurnReport')
    oc = sht.range("C11:J13").value
    for line in oc:
        try:
            if len(line[0]) < 1:
                pass
            elif str(line[0]) == ' ':
                pass
            else:
                # print len(line[0])
                # print(line)
                oc_line = [fname.split('\\')[-1]]
                oc_line.extend(line)
                off_color.append(oc_line)
        except:
            pass

app = xw.apps.active
app.quit()
if len(off_color) > 1:
    # oc_emails = ["chris.schottmiller@aksteel.com", "glenn.stamper@aksteel.com", "Sal.Ciriello@aksteel.com"]
    oc_emails = ["chris.schottmiller@aksteel.com"]
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
    #you = "chris.schottmiller@aksteel.com"
    you = oc_emails
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Off Color Alert {}".format(todays_date.strftime("%B %d, %Y"))
    msg['From'] = me
    msg['To'] = ", ".join(you)
    # Create the body of the message (a plain-text and an HTML version).
    text = "OC Alert"
    html1 = """\
    <html>
      <head></head>
      <body>"""
    html2 = ""
    for idx, line in enumerate(off_color):
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