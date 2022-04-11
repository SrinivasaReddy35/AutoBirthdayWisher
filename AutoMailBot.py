import pandas as p
import datetime
import smtplib

# Enter your authentication details i.e your correct gmail and password below
GMAIL_ID = 'srinivasareddy1619@gmail.com'
GMAIL_PASSWORD = 'reddy1619'


def SendEmail(to, sub ,msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}")
    # s = smtplib.SMTP('smtp.gmail.com', 587)
    # s.starttls()
    # s.login(GMAIL_ID, GMAIL_PASSWORD)
    # s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    # s.quit()


if __name__ == "__main__":
    df = p.read_excel("Demo.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(type(today))
    writeInd = []

    for index, item in df.iterrows():
        # print(index, item['DOB'])
        Bday = item['DOB'].strftime("%d-%m")
        # print(Bday)
        if(today==Bday) and yearNow not in str(item['Year']):
            name = item['Name']
            mssg = f"Wishing you a very HAPPY BIRTHDAY!<---{name}--->From the management of SVREC...Wishing you a happy a HAPPY LEARNING"
            SendEmail(item['Email'], "SURPRISE", mssg)
            writeInd.append(index)

    # print(writeInd)
    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) +", " + str(yearNow)
        # print(df.loc[i, 'Year'])

    # print(df)
    df.to_excel('Demo.xlsx', index=False)