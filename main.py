from PyPDF2 import PdfReader
import re


reader = PdfReader("Egypt foods - TQM Certificates.pdf")
number_of_pages = len(reader.pages)

issue_date = '9/7/2023'
company_name = "Egypt Foods Group"
certificate_field = "Total Quality Management Training (TQM)"

for page in reader.pages:
    text = page.extract_text()

    # Cert Holder
    exp = re.compile(r'It is hereby certified that\s*(.*?)\s*Attended the training on')
    certificate_holder = exp.findall(text)
    print(certificate_holder)

    # Cert Code
    exp = re.compile(r'\b[A-Z]{2}-[A-Z]{3}-\d{3}-\d{4}-\d{2}\b')
    certifictae_code = exp.findall(text)
    print(certifictae_code)