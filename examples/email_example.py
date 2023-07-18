import smtplib
from email.message import EmailMessage


# instantiate email message
msg = EmailMessage()

# set email headers
msg["From"] = "franciskohara@gmail.com"  # sender address
msg["To"] = "franciskohara@gmail.com"  # recipient address
msg["Subject"] = "Batman Monologue"

# set email body
email_text = """Thursday, October 31st. The city streets are crowded for the holiday.
Even with the rain. But hidden in the chaos is the element. Waiting to strike like snakes.
But I'm there too. Watching. Two years of nights have turned me into a nocturnal animal.
I must choose my targets carefully. It's a big city. I can't be everywhere. But they don't where I am.
"""
msg.set_content(email_text)

# add email attachment
with open("../data/attachment.docx", "rb") as attachment:
    msg.add_attachment(attachment.read(), filename="Batman Monologue.docx", maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document")

# start smtp client & send
smtp = smtplib.SMTP("smtp.gmail.com", 587)
smtp.ehlo()
smtp.starttls()
sender_email = input("Enter your email address:\t")
password = input("Enter your password:\t")
smtp.login(sender_email, password)
smtp.send_message(msg)
smtp.quit()
