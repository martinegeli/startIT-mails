# coding=utf-8
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# from read_mails import read_hoopla


def send_to_all(dict):
    for key, value in dict.items():
        fromaddr = "martin.egeli@startntnu.no"
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
                
                <p>Blant alle svar velger vi ut tre stykker som får en chromecast som premie! Dette er store sjanser</p>
                
                <a href='https://goo.gl/forms/LLukrGjTvv3dfkyo2'>Trykk her for link til spørreundersøkelse</a>
                
                <p>Vi håper å se deg til neste år også. Ha en fin dag videre! :) </p>
                
                <img src="cid:image1" />
            <body>
        <html>
        """.format(names, plural)

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

# mail_dict = read_hoopla('hoopla.xls')


test_dict = {
    #'erling.olweus@startntnu.no': ['Erling'],
    #'andreas.engebretsen@startntnu.no': ['Andreas'],
    #'isabel.slorer@startntnu.no': ['Isabel'],
    #'sanne.saetre@startntnu.no': ['Sanne'],
    'martinegeli9@gmail.com': ['Martin']
}

send_to_all(test_dict)