import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

#
# #
# # Dictionary with list object in values
# studentData = {
#     'adjclose' : {'m':56,'p':45,'r':59, 's': 42,'t':57},
#     'adjclose_N' : {'m':76,'p':88,'r':79, 's': 41,'t':54},
#     'adjclose_X' : {'m':7,'p':8,'r':5, 's': 41,'t':54},
#     'adjclose_Y' : {'m':4,'p':33,'r':87, 's': 41,'t':54}}
#
# df = pd.DataFrame(studentData)

def sendme_dataframe(dataframe):
    pd.set_option('display.max_colwidth', -1)
    gmail_user = 'urusfin@gmail.com'
    gmail_password = 'Misys123$'

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        msg = MIMEMultipart()
        msg['From'] = gmail_user
        msg['To'] = "urusfin@gmail.com"
        msg['Subject'] = "PDA Covid News Scrapper"
        message = dataframe.to_string

        html = """\
        <html>
          <head></head>
          <body>
            {0}
          </body>
        </html>
        """.format(dataframe.to_html())

        dataframe_html = MIMEText(html, 'html')
        msg.attach(dataframe_html)
        server.send_message(msg)
        del msg
        server.quit()

    except Exception as e:
        print(e)
        print('Something went wrong...')

