# -*- coding: utf-8 -*-

import os
import smtplib
import getpass
#import mimetypes

from email.utils import formataddr
from email.utils import formatdate
from email.utils import COMMASPACE

from email.header import Header
from email import encoders

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
#from email.mime.application import MIMEApplication


#n = 0
#while n == 0:
#    print('Qual o email da pessoa para a qual você deseja enviar o email?')
#    ad1 = input()
#    print('Qual o nome dessa mesma pessoa?')
#    nam1 = input()
#    print('Confirmar    ' + nam1 +' | '+ ad1)
#    print('Y or N')
#    ss = input()
#    if ss == 'Y':
#        n = 1

nam1 = ['Pedro', 'José']
ad1 = ['pandrade@empresajunior.com.br', 'pefandrade@hotmail.com']

def send_email(sender_name: str, sender_addr: str, smtp: str, port: str,
               recipient_addr: list, subject: str, html: list, text: str,
               img_list: list=[], attachments: list=[],
               fn: str='last.eml', save: bool=False):

    passwd = getpass.getpass('Password: ')

    sender_name = Header(sender_name, 'utf-8').encode()

    for index, f in enumerate(recipient_addr):
        msg_root = MIMEMultipart('mixed')
        msg_root['Date'] = formatdate(localtime=1)
        msg_root['From'] = formataddr((sender_name, sender_addr))
        msg_root['To'] = f
        msg_root['Subject'] = Header(subject, 'utf-8')
        msg_root.preamble = 'This is a multi-part message in MIME format.'

        msg_related = MIMEMultipart('related')
        msg_root.attach(msg_related)

        msg_alternative = MIMEMultipart('alternative')
        msg_related.attach(msg_alternative)

        msg_text = MIMEText(text.encode('utf-8'), 'plain', 'utf-8')
        msg_alternative.attach(msg_text)

        msg_html = MIMEText(html.format(nam1[index]).encode('utf-8'), 'html', 'utf-8')
        msg_alternative.attach(msg_html)

        for i, img in enumerate(img_list):
            with open(img, 'rb') as fp:
                msg_image = MIMEImage(fp.read())
                msg_image.add_header('Content-ID', '<image{}>'.format(i))
                msg_related.attach(msg_image)

        for attachment in attachments:
            fname = os.path.basename(attachment)

            with open(attachment, 'rb') as f:
                msg_attach = MIMEBase('application', 'octet-stream')
                msg_attach.set_payload(f.read())
                encoders.encode_base64(msg_attach)
                msg_attach.add_header('Content-Disposition', 'attachment',
                                      filename=(Header(fname, 'utf-8').encode()))
                msg_root.attach(msg_attach)

        mail_server = smtplib.SMTP(smtp, port)
        mail_server.ehlo()

        try:
            mail_server.starttls()
            mail_server.ehlo()
        except smtplib.SMTPException as e:
            print(e)

        mail_server.login(sender_addr, passwd)
        mail_server.send_message(msg_root)
        mail_server.quit()

        if save:
            with open(fn, 'w') as f:
                f.write(msg_root.as_string())

if __name__ == '__main__':
    # Usage:
    sender_name = 'Pedro Enrique Andrade'
    sender_addr = 'pandrade@empresajunior.com.br'
    smtp = 'smtp.gmail.com'
    port = '587'
    recipient_addr = ad1
    subject = 'Contato | Empresa Júnior PUC-Rio'
    text = ''
    #text_s = [nam1]
    html = """
        <html>
        <head>
        <meta http-equiv="content-type" content="text/html;charset=utf-8" />
        </head>
        <body>
        <font face="georgia" size=2>
        Boa tarde {}, tudo bem?<br/>
        Aqui é o Pedro da Empresa Júnior-PUC-Rio, somos uma consultoria especializada em....
        </font>
        <img src="cid:image0" border=0 />
        <img alt=3D"" style=3D"width:0px;max-height:0px;overflow:hidden" src=3D"https://mailfoogae.appspot.com/t?sender=3DacGFuZHJhZGVAZW1wcmVzYWp1bmlvci5jb20uYnI%3D&amp;type=3Dzerocontent&amp;guid=3D5a25310f-2be3-41a7-a7c2-0eb008c0ebf3"><font color=3D"#ffffff" size=3D"1">=E1=90=A7</font></div>
        </body>
        </html>
        """

        #.format(nam1)  # ver como .format() funciona https://www.programiz.com/python-programming/methods/string/format

    #img_list = ['/Users/pedroenriqueandrade/Desktop/roxo.jpg']  # -> image0, image1, image2, ...
    #attachments = ['/Users/pedroenriqueandrade/Desktop/Institucional - Empresa Júnior PUC-Rio.pdf']

    send_email(
                sender_name,
                sender_addr,
                smtp, port,
                recipient_addr,
                subject,
                html,
                text,
                #img_list,
                #attachments,
                fn='my.eml',
                save=True
    )
