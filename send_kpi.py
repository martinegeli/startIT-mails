# coding=utf-8
import smtplib

import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from read_mails import read_hoopla


def send_to_all(dict):
    for key, value in dict.items():
        fromaddr = "mail@mail.no"
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['Subject'] = "Spørreundersøkelse StartIT"

        names = ""
        plurar = ""
        if len(value) == 1:
            names = value[0]
            plural = "du"
        else:
            names = ", ".join(value[0:len(value)-1]) + "og " + value[-1]
            plural = "dere"

        html = """\
        <html>
            <body>
                <h4>Hei {}!</h4>
                <p><strong>Vi håper {} hadde en fantastisk kveld under årets StartIT</strong></p>
                
                <p>I vårt videre arbeid med StartIT er din tilbakemelding utrolig viktig!</p>
                
                <p>Vi sender derfor ut en spørreundersøkelse vi ønsker at du skal svare på, som kun tar 2-3 minutter</p>
                
                <p>Blant besvarelsene vil vi trekke tre heldige vinnere som får en chromecast!</p>
                
                <a href='link'>Trykk her for link til spørreundersøkelse</a>
                
                <p>Vi håper å se deg til neste år også. Ha en fin dag videre! :) </p>
                
                <p>Med vennlig hilsen</p>
                <p>StartIT-gjengen :) </p>
                
            <body>
        <html>
        """.format(names, plural)

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()

        server.login(fromaddr, 'pass')
        toaddr = key
        msg['To'] = key
        msg.attach(MIMEText(html, 'html'))
        server.sendmail(fromaddr, toaddr, msg.as_string())
        server.quit()

def read_hoopla(file):
    # Get the cols needed in the DataFrame, [name, email]
    df = pd.read_excel(file, usecols=[22, 30])

    mail_dict = {}
    # Iterate over the DataFrame and make a dictionary containing name and email
    for idx, row in df.iterrows():
        if row[1] not in mail_dict:
            mail_dict[row[1]] = [row[0]]
        else:
            mail_dict[row[1]].append(row[0])

    # Return the mails in a dictionary containing names and emails
    return mail_dict

# mail_dict = read_hoopla('hoopla.xls')


test_dict = {

}

# send_to_all(mail_dict)