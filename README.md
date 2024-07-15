- 👋 Hi, I’m @Shawn730-coder
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...

<!---
Shawn730-coder/Shawn730-coder is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
import openpyxl
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class Student:
    def __init__(self, name, grade, email, homework_file):
        """
        Initialize Student object
        :param name: Student's name
        :param grade: Student's grade
        :param email: Student's email
        :param homework_file: Path to student's homework PDF file
        """
        self.name = name
        self.grade = grade
        self.email = email
        self.homework_file = homework_file

def read_students_from_excel(filename):
    """
    Read student information from an Excel file and return a list of Student objects
    :param filename: Excel filename
    :return: List of Student objects
    """
    students = []
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, grade, email, homework_file = row
        students.append(Student(name, grade, email, homework_file))

    return students

def send_email(from_email, from_password, to_email, subject, message, attachments):
    """
    Send an email
    :param from_email: Sender's email
    :param from_password: Sender's email password
    :param to_email: Recipient's email
    :param subject: Email subject
    :param message: Email body
    :param attachments: List of paths to attachment files
    """
    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Add email body
    msg.attach(MIMEText(message, 'plain'))

    # Add attachments
    for attachment_path in attachments:
        attachment_filename = os.path.basename(attachment_path)
        student_name = attachment_filename.split(' - ')[0]  # Extract student name from attachment filename
        new_attachment_filename = f"{student_name} - Detailed Marks.pdf"

        with open(attachment_path, "rb") as attachment_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.read())

        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{new_attachment_filename}"')
        msg.attach(part)

    # Connect to the SMTP server and send the email
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(from_email, from_password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

def main():
    excel_filename = 'D:/Canada/Teaching Assistant/Python Test/Name Path.xlsx'
    students = read_students_from_excel(excel_filename)

    outlook_email = '@outlook.com'
    outlook_password = 'ehsufvvdtgirbjkt'
    common_attachment_path = 'D:/Canada/Teaching Assistant/TA for 730/Assignment 2/Solutions of Assignment 2 for ENGY 730.pdf'  # Unified PDF file path

    for student in students:
        if not student.email:  # Check if email is empty
            print(f"Skipping {student.name} because email is empty.")
            continue  # If email is empty, skip this student and move to the next one

        subject = '<From 730 TA> Grade of Assignment 2 for ENGY 730'
        message = f'Dear {student.name},\n\nI hope you are enjoying a sunny day.'
        attachments = [student.homework_file, common_attachment_path]  # List of attachment file paths
        send_email(outlook_email, outlook_password, student.email, subject, message, attachments)
        print(f"Email sent to {student.name} at {student.email}.")

if __name__ == "__main__":
    main()
