import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = pd.read_excel("Student.xlsx")
print(type(data))
emailcol = data.get("Email")
list_of_emails=list(emailcol)
print(list_of_emails)


try:

    #object of SMTP
    server=sm.SMTP("smtp.gmail.com",587)
    server.starttls()

    #Login using Email
    server.login("pradhuyamn@gmail.com","Enter_Your_Password")
    from_ ="pradhuyamn@gmail.com"
    to_=list_of_emails
    message=MIMEMultipart("alternative")
    message['Subject']="This is an Test Message"
    message["from"]="pradhuyamn@gmail.com"

    #Creating a Text in form of HTML
    html='''
    <html>
    <head>
    </head>
    <body>
    <h1>Testing</h1>
    <button style="padding:20px;background:green;color:white" ><Varify/button>
    </body>
    </html>
    '''

    part2=MIMEText(html,"html")
    message.attach(part2)
    #Send The Message
    server.sendmail(from_ ,to_ ,  message.as_string())
    print("Done")

except Exception as e:
    print(e)
