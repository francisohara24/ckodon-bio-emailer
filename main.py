# import required modules
import pandas as pd
from functions import create_doc, create_email, send_email, create_smtp

data = pd.read_excel("data/Ckodon Bio Submission Form (Responses).xlsx")
students = data.head(50)  # select first 50 students

smtp = create_smtp("smtp.gmail.com", 587, "ckodontech@gmail.com", "fzdbwumpxpyolpny")

for row in students.index:
    # extract name, email, bio
    student = students.loc[row]
    student_name = student["Full Name"]
    student_email = "franciskohara@gmail.com"#student["Email"]

    # create word document containing bio
    doc_content = student["Bio"]
    doc_name = f"Ckodon Bio [{student_name}].docx"
    doc_path = create_doc(doc_name, doc_content)

    # create email message
    msg_subject = "Ckodon Bio Update"
    msg_recipient = student_email
    msg_attachment = doc_path
    msg_content = f"""Dear {student_name},

    Your Ckodon Bio Submission has been reviewed.
    Find attached a Word document containing your updated Bio.
    You are not required to take any further action if your Bio has already been reviewed.
    You may reply to this email with any concerns or questions.

    Best,
    The Ckodon Foundation."""
    msg = create_email(msg_recipient, msg_subject, msg_content, msg_attachment)

    # send email message
    try:
        send_email(msg, smtp)
    except:
        print("Error sending mail at-")
        print("Row No:", row)
        print("Student Name:", student_name)
        print("Student Email:", student_email)
    break
