# coding=utf-8
import smtplib

import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def send_to_all(dict):
    for key, value in dict.items():
        fromaddr = "mail@mail.no"
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['Subject'] = "Starterbillett til StartIT"

        names = value[0]

        html = """\
        <html>
            <body>
                <h4>Hei {}!</h4>
                <p><strong>Vi bekrefter at din billett til StartIT 2019 er registrert!</strong></p>

                <p><strong>Generell informasjon:</strong><p>
                <ul>
                    <li>Registrering begynner 17:45</li>
                    <li>Du må ha registrert deg ved inngangen før 18:15</li>
                    <li>Det blir middag 20:00. Tapasbuffet og desserttallerken med én bong til alle</li>
                    <li>Kvelden er ferdig rundt 23:00</li>
                    <li>Etterfest for de som ønsker det!</li>
                </ul>
                
                <p>Vi har selvsagt oversolgt noen billetter til vanlige billetter, og da går det ut over oss startere om for mange møter opp</p>
                <p>Det vil si at vi kanskje må vente på at andre har forsynt seg, eventuelt at vi bestiller inn noe ekstra mat :) </p>
                
                <p>Om du ønsker å bidra er det bare å møte opp før 17:45</p>
                <p>Andreas Engebretsen er sjef for kvelden, så det er bare å finne han om du skulle ønske å bidra.</p>
                
                <br>
                <p>Med vennlig hilsen</p>
                <p>Gjengen i StartIT :) </p>

                <img src="cid:image1" style="width: 600px"/>
            <body>
        <html>
        """.format(names)

        img = open('startit_img.jpg', 'rb').read()
        msg_img = MIMEImage(img, 'jpg')
        msg_img.add_header('Content-ID', '<image1>')
        msg_img.add_header('Content-Disposition', 'inline', filename="startit_image.jpg")

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()

        server.login(fromaddr, 'pass')
        toaddr = key
        msg['To'] = key
        msg.attach(MIMEText(html, 'html'))
        # msg.attach(msg_img)
        server.sendmail(fromaddr, toaddr, msg.as_string())
        server.quit()

def read_starters(file):
    # Get the cols needed in the DataFrame, [name, email]
    df = pd.read_excel(file, usecols=[7, 19])

    mail_dict = {}
    # Iterate over the DataFrame and make a dictionary containing name and email
    for idx, row in df.iterrows():
        mail_dict[row[1]] = [row[0]]

    # Return the mails in a dictionary containing names and emails
    return mail_dict

start_dict = read_starters('hoopla_start.xls')

send_to_all(start_dict)