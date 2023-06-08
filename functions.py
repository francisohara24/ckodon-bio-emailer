from docx import Document
from email.message import EmailMessage
from smtplib import SMTP


# function for creating word document given file name and content.
def create_doc(filename: str, content: str) -> str:
    """Creates word document in the attachments directory with the specified content and file name.
    Returns path to the created document."""
    document = Document()
    document.add_paragraph(content)
    filepath = "./data/attachments/" + filename
    document.save(filepath)
    return filepath


def create_email(recp_address: str, subject: str, content: str, attachment_path: str) -> EmailMessage:
    """Creates an email message with specified address, subject, content, and attachment.
     Returns EmailMessage object representing the created email."""
    msg = EmailMessage()
    msg["To"] = recp_address
    msg["Subject"] = subject
    msg.set_content(content)
    with open(attachment_path, "rb") as document:
        attachment_name = attachment_path.split("/")[-1]
        msg.add_attachment(document.read(),
                           filename=attachment_name,
                           maintype="application",
                           subtype="vnd.openxmlformats-officedocument.wordprocessingml.document")
    return msg


def create_smtp(address: str, port: int, username: str, password: str) -> SMTP:
    """Instantiates an SMTP client that connects to a server with the specified credentials"""
    smtp = SMTP(address, port)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username, password)
    return smtp


def send_email(msg: EmailMessage, smtp_client: SMTP):
    """Sends email message using the specified smtp client."""
    smtp_client.send_message(msg)
    smtp_client.quit()
    print("EMAIL SENT")
