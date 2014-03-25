import xlrd

from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

from creds import OT04, OT04_PASSWORD
from mail_lists import OFFICER_LOOKUP
from email_text import EMAIL_TEXT

wkbk = xlrd.open_workbook('olentangy_2004_2014.xlsx')
sht = wkbk.sheets()[0]
#TEMP
offs = []
for r in range(1,sht.nrows, 20):
    if sht.row(r)[5].value:
        btext = ''
        first_name = sht.row(r)[2].value.capitalize()
        full_name = '%s %s' % (first_name, sht.row(r)[1].value.capitalize())
        cell = sht.row(r)[3].value if sht.row(r)[3].value != '' else 'We don\'t have this info from you!'
        parent_contact = sht.row(r)[4].value if sht.row(r)[4].value != '' else 'We don\'t have this info from you!'
        email = sht.row(r)[5].value if sht.row(r)[5].value != '' else 'We don\'t have this info from you!'
        addr = '%s   %s, %s   %s' % (sht.row(r)[7].value, sht.row(r)[8].value, sht.row(r)[9].value, sht.row(r)[10].value)
        friends = '%s and %s' % (sht.row(r)[13].value, sht.row(r)[14].value)
        if friends == ' and ':
            friends = 'We don\'t have this info from you!'
        btext += EMAIL_TEXT % (first_name, full_name, cell, parent_contact, email, addr, friends)
        FROM = OFFICER_LOOKUP[sht.row(r)[0].value]
        TO = email
        if FROM not in offs:
            offs.append(FROM)
            print 'From: %s, To: %s' % (FROM, TO)
            msg = MIMEMultipart()
            msg['Cc'] = FROM
            RFROM = OT04
            msg['From'] = RFROM
            msg['To'] = FROM
            msg['Subject'] = 'TBD Subject Line'
            msg.attach(MIMEText(btext.encode('utf-8'),'plain'))

            import smtplib
            server = smtplib.SMTP('smtp.gmail.com',587)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(OT04, OT04_PASSWORD)
            text = msg.as_string()

            server.sendmail(RFROM, FROM.split(','), text)
