# -*- coding: utf-8 -*-
"""
Created on Mon Feb 14 20:52:10 2022

@author: reda3
"""




import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage
import random
from quote import quote



mylist = ["leo tolstoy", "simon sinek", "Paulo Coelho", "Stephen Covey", "Tony Robbins", "Mark Twain", "Fyodor Dostoevsky", "Lao Tzu", "Sun Tzu"]

def quote_generate():
    
    search = random.choice(mylist)
    result = quote(search, limit=1)

    df = pd.DataFrame(result)
    
    return df




quote_result = quote_generate()



#x = "Your Quote Today: '" + quote_result.iloc[0][2] + "' By " + quote_result.iloc[0][0]
#print(x)

def sendEmail(to, sub, msg):
    print(f"email to {to} \nsend with subject: {sub}\n message: {msg}")
    email = EmailMessage()
    email['from'] = 'Reda Aldahan'
    email['to'] = f"{to}"
    email['subject'] = f'{sub}'

    email.set_content(f'{msg}')

    with smtplib.SMTP(host='smtp.gmail.com', port=587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login('your_Email@gmail.com','your password')
        smtp.send_message(email)
        print("Email send")
    pass
if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    #print(type(today))
    update = []
    td =  datetime.datetime.now().strftime("%d-%m")
    #print(yearnow)
    for index, item in df.iterrows():
        #print(index,item['birthday'])
        #print(type(bday))
        if td not in str(item["Date"]): 
            sendEmail(item['Email'] , "Happy QUOTE "+item["Name"], "Your Quote Today: '" + quote_result.iloc[0][2] + "' By " + quote_result.iloc[0][0])
            update.append(index)
    for i in update:
        dt = df.loc[i, 'Date']
        #print(yr)
        df.loc[i,'Date'] = f"{td}, {dt}"
        #print((df.loc[i, 'Year'])
    #print(df)
    df.to_excel("data.xlsx", index=False)