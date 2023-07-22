# ckodon-essay-emailer
This is a Python program I wrote to automate the process of submitting reviewed college essay bios via email to students at [the Ckodon Foundation](https://www.ckodon.com/ckodon-foundation).

Reviewed essays and student email addresses are extracted from an Excel Sheet of Google Form responses using `pandas`.

`python-docx` is used to save the extracted text in Word documents, and ```email``` and ```smtplib``` are used to send the documents to students as email attachments.

The Excel sheet and the `data/attachments` directory are not included for privacy reasons.
