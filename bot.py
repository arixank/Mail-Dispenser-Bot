# Mail Dispenser bot
import pandas as pd
import smtplib
from email.message import EmailMessage

MAIL_ID = "validlogbot@gmail.com"
PASSWORD = "maildispenser101"

# ? Read the file
data = pd.read_excel(
    "mail dispenser\Mail-Dispenser-Bot\data.xlsx", engine="openpyxl")
# ? Get the Mail Id
mail_list = [mail_id["Mail id"] for (index, mail_id) in data.iterrows()]
# ? Get recipient Names
name_list = [name["Name"] for(index, name) in data.iterrows()]
# # ? Framing a Dictionary
# data_list = {row["Name"]: row["Mail id"] for (index, row) in data.iterrows()}

# ? Message
msg = EmailMessage()
msg["Subject"] = 'Greetings from Message Dispenser Bot'
msg["From"] = MAIL_ID
msg.set_content("Hey , this is a test message from the bot!")


# ? Sending Mails
with smtplib.SMTP(host="smtp.gmail.com", port=587) as connection:
    connection.starttls()
    connection.login(user=MAIL_ID, password=PASSWORD)
    for each in mail_list:
        connection.send_message(to_addrs=each, msg=msg)
        print(f"Mail was sent to {each}")
