# coding=utf-8
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

from read_mails import read_hoopla

def send_to_all(dict):
    for key, value in dict.items():
        fromaddr = "martin.egeli@startntnu.no"
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['Subject'] = "Confirmation av påmelding StartIT"

        names = ""
        plurar = ""
        if len(value) == 1:
            names = value[0]
            plural = "din billett"
        else:
            names = ", ".join(value[0:len(value)-1]) + " og " + value[-1]
            plural = "deres billetter"

        html = """\
        <html>
            <body>
                <h4>Hei {}!</h4>
                <p><strong>Vi bekrefter at {} til StartIT 2019 er registrert!</strong></p>
                
                <p>Vi kan garantere en spennende kveld på Radisson Blu Royal Garden nå til onsdag.</p>
                
                <p><strong>Generell informasjon:</strong><p>
                <ul>
                    <li>Registrering begynner 17:45</li>
                    <li>Du må ha registrert deg ved inngangen før 18:15</li>
                    <li>Det blir middag 20:00. Tapasbuffet og desserttallerken med én bong til alle</li>
                    <li>Kvelden er ferdig rundt 23:00</li>
                    <li>Etterfest for de som ønsker det!</li>
                </ul>
                
                <p>Under kvelden blir det mange konkurranser med mulighet til å vinne kule premier! Vi gleder oss til å se deg.</p>
                
                <p>Om du ikke har mulighet til å delta, så ønsker vi at du svarer på denne mailen :) </p>
                
                <br>
                <p>Med vennlig hilsen,</p>
                <p>Martin Egeli</p>
                <p>Prosjektleder StartIT</p>
                
                <br>
                
                <img src="cid:image1" style="width: 600px"/>
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

}

# send_to_all(mail_dict)


