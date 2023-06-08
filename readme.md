# doc-emailer
This is a Python program I wrote to submit word documents to students via email.

Documents and student email addresses are extracted from a Google form responses Excel sheet using ```pandas```.

```python-docx``` is used to save the extracted text in Word documents, and ```email``` and ```smtplib``` are used to send the documents to students as email attachments.