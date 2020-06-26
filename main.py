import pandas as pd
import datetime
import smtplib
from twilio.rest import Client
from credentials import account_sid, auth_token, my_twilio
# import notify2
import os

os.chdir(r"C:\Users\Ayush\Desktop\Ayush\Engineernig Data\GrabCAD\python\Automatic_Wishing_Software")
# os.mkdir("testing")

# Enter your details

GMAIL_ID = ''  # Your Mail id here
GMAIL_PSWD = ''  # mail id password here


def sendEmail(to, sub, msg):
    # print(f'Email to {to} sent with subject: {sub} and msg : {msg}')

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f'Subject: {sub}\n\n{msg}')
    s.quit()


def sendSMS(message, my_cell):
    print(my_cell[0:13])
    client = Client(account_sid, auth_token)
    my_msg = message
    message = client.messages.create(to=my_cell[0:13], from_=my_twilio, body=my_msg)
    # message_whatsapp = client.messages.create(from_='whatsapp:+14155238886',
    #                           body=my_msg,
    #                           to= f'whatsapp:{my_cell[0:13]}')
    # print(message_whatsapp.sid)
    # a = notify2.Notification('Status', 'I have wished birthdays and anniversaries today')


if __name__ == '__main__':

    # sendEmail(GMAIL_ID, 'subject', 'test message')
    # exit()

    df = pd.read_excel("data.xlsx")
    today = datetime.datetime.now().strftime('%d-%m')
    yearNow = datetime.datetime.now().strftime('%Y')

    writeInd = []
    for index, item in df.iterrows():
        try:
            bday = item['BirthdayDate'].strftime('%d-%m')
            if (today == bday) and yearNow not in str(item['YearBirthday']):
                sendEmail(item['Email'], "Happy Birthday", item['BirthdayDialogue'])
                writeInd.append(index)
                sendSMS(item['BirthdayDialogue'], f"+91{item['Phone']}")
        except:
            pass

    for i in writeInd:
        yr = df.loc[i, 'YearBirthday']
        df.loc[i, 'YearBirthday'] = str(yr) + ', ' + str(yearNow)
    df.to_excel("data.xlsx", index=False)

    # anniversary

    today_ani = datetime.datetime.now().strftime('%d-%m')

    yearNow_ani = datetime.datetime.now().strftime('%Y')

    writeInd_ani = []
    for index, item in df.iterrows():
        try:
            ani = item['AnniversaryDate'].strftime('%B-%d')
            if (today == ani) and yearNow_ani not in str(item['YearAnniversary']):
                sendEmail(item['Email'], "Happy Anniversary", item['AnniversaryDialogue'])
                writeInd_ani.append(index)
                sendSMS(item['AnniversaryDialogue'], f"+91{item['Phone']}")

        except:
            pass

    for i_1 in writeInd_ani:
        yr = df.loc[i_1, 'YearAnniversary']
        df.loc[i_1, 'YearAnniversary'] = str(yr) + ', ' + str(yearNow_ani)

    df.to_excel('data.xlsx', index=False)
