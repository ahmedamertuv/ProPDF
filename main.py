from PyPDF2 import PdfReader, PdfWriter
from openpyxl import Workbook
import re


reader = PdfReader("Egypt foods - TQM Certificates.pdf")
wb = Workbook()
ws = wb.active

issue_date = "9/7/2023"
company_name = "Egypt Foods Group"
certificate_field = "Total Quality Management Training (TQM)"

titles = [
    "certificate_holder",
    "certificate_code",
    "certificate_field",
    "issue_date",
    "company_name",
]

ws.append(titles)

for page in reader.pages:
    text = page.extract_text()

    # Cert Holder
    exp = re.compile(r"It is hereby certified that\s*(.*?)\s*Attended the training on")
    certificate_holder = exp.findall(text)
    print(certificate_holder)

    # Cert Code
    exp = re.compile(r"\b[A-Z]{2}-[A-Z]{3}-\d{3}-\d{4}-\d{2}\b")
    certificate_code = exp.findall(text)
    ws.append(
        [
            certificate_holder[0],
            certificate_code[0],
            certificate_field,
            issue_date,
            company_name,
        ]
    )
wb.save("final.xlsx")
