#! python3
# readCensusExcel.py - Tabulates population and number of census tracts for
# each county.
import openpyxl, os, docx, smtplib
from docx2pdf import convert
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def writName(Name):
    doc = docx.Document("C:\\Users\\fadia\\Desktop\\ourproject\\Certificate.docx")
    p = doc.paragraphs[9]
    p.add_run(Name)
    p.runs[0].bold = True
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = docx.shared.Pt(18)
    doc.save("C:\\Users\\fadia\\Desktop\\ourproject\\Certificate1.docx")
    convert("Certificate1.docx", "C:\\Users\\fadia\\Desktop\\ourproject\\Certificate\\" + Name + ".pdf")


def sendEmail(Email, Name):
    body = 'Dear ' +Name + '''

Thank you for your interest in improving your skills, We appreciate your time that you spent it with us. We will achieve our goals together soon.

Because your experience is valuable to us, in the attachment you will find your own certificate of your successful completion of Automate the Boring Stuff with Python Programming course


Good luck,
Best regards,
ROB and FADIA'''

    sender = 'angelrobbor@gmail.com'

    password = 'AngelrobborAngel'

    receiver = Email

    mail = MIMEMultipart()
    mail['From'] = sender
    mail['To'] = receiver
    mail['Subject'] = 'your certificate'

    mail.attach(MIMEText(body, 'plain'))

    pdfname = "C:\\Users\\fadia\\Desktop\\ourproject\\Certificate\\" + Name + ".pdf"

    binary_pdf = open(pdfname, 'rb')

    payload = MIMEBase('application', 'octate-stream', Name=pdfname)
    payload = MIMEBase('application', 'pdf', Name=Name)
    payload.set_payload((binary_pdf).read())

    encoders.encode_base64(payload)

    payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
    mail.attach(payload)

    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.ehlo()

    session.starttls()

    session.login(sender, password)

    text = mail.as_string()
    session.sendmail(sender, receiver, text)
    session.quit()
    print('Mail Sent')


os.chdir("C:\\Users\\fadia\\Desktop\\ourproject")
vwb = openpyxl.load_workbook('Email.xlsx')
# sheet = vwb.get_sheet_by_name('Sheet1')
sheet = vwb["Sheet1"]

# TODO: Fill in countyData with each county's population and tracts.
print('Reading rows...')
for row in range(1, sheet.max_row + 1):
    # Each row in the spreadsheet has data for one census tract.
    Name = sheet['A' + str(row)].value
    writName(Name)
    Email = sheet['B' + str(row)].value
    sendEmail(Email, Name)



