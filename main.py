import jinja2
import pdfkit
from openpyxl import load_workbook
import pandas as pd

file = "ACES_MembershipData.xlsx"  # Name of your Excel File
data = pd.ExcelFile(file)
wb = load_workbook(file)
ws = wb['Sheet1']  # Name of Sheet
all_rows = list(ws.rows)

SEACounter = 0
SEBCounter = 0
TEACounter = 0
TEBCounter = 0
InvoiceCounter = 0
MembershipID = "NULL"

for row in all_rows[1:10]:  # Change 10 to total number of rows in your sheet
    StudentName = row[0].value
    StudentClass = row[1].value
    PhoneNumber = row[2].value
    EmailID = row[3].value
    Date = row[4].value
    if StudentClass in {"SE A", "SE B"}:
        ItemName = "ACES Membership Fees SE 2024-27"
        ItemDescription = "Membership fees for SE Students for 3 years"
        if StudentClass[3] == "A":
            SEACounter += 1
            MembershipID = "ACES/2023/SE/%s/%.3d" % (StudentClass[3], SEACounter)
        elif StudentClass[3] == "B":
            SEBCounter += 1
            MembershipID = "ACES/2023/SE/%s/%.3d" % (StudentClass[3], SEBCounter)
        Subtotal = 400
    elif StudentClass in {"TE A", "TE B"}:
        ItemName = "ACES Membership Fees TE 2024-26"
        ItemDescription = "Membership fees for TE Students for 2 years"
        if StudentClass[3] == "A":
            TEACounter += 1
            MembershipID = "ACES/2023/TE/%s/%.3d" % (StudentClass[3], TEACounter)
        elif StudentClass[3] == "B":
            TEBCounter += 1
            MembershipID = "ACES/2023/TE/%s/%.3d" % (StudentClass[3], TEBCounter)
        Subtotal = 300
    else:
        ItemName = "ACES Membership Fees 2024-25"
        ItemDescription = "Membership fees for 1 year"
        Subtotal = 400
    Total = Subtotal
    InvoiceCounter += 1
    InvoiceID = f"{InvoiceCounter:04d}"
    print(f"{StudentName}\t{StudentClass}\t{PhoneNumber}\t{EmailID}\t{ItemName}\t{ItemDescription}\t{Total}\t{Date}")
    context = {'StudentName': StudentName, 'Date': Date, 'Total': f'Rs. {Total:.2f}',
               'ItemName': ItemName, 'ItemDescription': ItemDescription, 'Subtotal': f'Rs. {Subtotal:.2f}',
               'PhoneNumber': PhoneNumber, 'EmailID': EmailID, 'MembershipID': MembershipID, 'InvoiceID': InvoiceID
               }
    template = jinja2.Environment(loader=jinja2.FileSystemLoader('./')).get_template('invoice.html').render(context)
    config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')
    pdfkit.from_string(template, 'GeneratedInvoices/Invoice_%s.pdf' % InvoiceID, configuration=config,
                       css='invoice.css', options={"enable-local-file-access": ""})
    print("Invoice Generated for %s" % StudentName)
