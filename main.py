# import required modules
import pandas as pd
from functions import create_doc, send_mail

data = pd.read_excel("data/Ckodon Bio Submission Form (Responses).xlsx")
students = data.head(50)  # select first 50 students

for row in students.index:
    # extract name, email, bio
    student = students.loc[row]
    student_name = student["Full Name"]
    student_email = student["Email"]

    # create word document containing bio
    doc_content = student["Bio"]
    doc_name = f"Ckodon Bio [{student_name}].docx"
    doc_path = create_doc(doc_name, doc_content)

    # send email
    eml_subject = "Ckodon Bio Update"
    eml_recipient = student_email
    eml_attachment = doc_path
    eml_content = f"""Dear {student_name},

    Your Ckodon Bio Submission has been reviewed.
    Find attached a Word document containing your updated Bio.
    You are not required to take any further action if your Bio has already been reviewed.
    You may reply to this email with any concerns or questions.

    Best,
    The Ckodon Foundation."""

    try:
        send_mail(eml_recipient, eml_subject, eml_content, eml_attachment)
    except:
        print("Error sending mail at-")
        print("Row No:", row)
        print("Student Name:", student_name)
        print("Student Email:", student_email)

