from PyPDF2 import PdfReader, PdfWriter
from openpyxl import Workbook
import re
import os


def writePDF(filename, page_or_file, output="output"):
    writer = PdfWriter()
    writer.add_page(page_or_file)

    if ".pdf" in filename:
        filename = filename[:-4]

    try:
        writer.write(f"./{output}/{filename}.pdf")
    except:
        os.mkdir(output)
        writer.write(f"./{output}/{filename}.pdf")

    return os.path.join(os.getcwd(), output)


def main():
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
        "file_path",
    ]

    ws.append(titles)

    for page in reader.pages:
        text = page.extract_text()

        # Cert Holder
        exp = re.compile(
            r"It is hereby certified that\s*(.*?)\s*Attended the training on"
        )
        certificate_holder = exp.findall(text)
        print(certificate_holder)

        # Cert Code
        exp = re.compile(r"\b[A-Z]{2}-[A-Z]{3}-\d{3}-\d{4}-\d{2}\b")
        certificate_code = exp.findall(text)

        certificate_dir = writePDF(certificate_code[0], page)
        ws.append(
            [
                certificate_holder[0],
                certificate_code[0],
                certificate_field,
                issue_date,
                company_name,
                certificate_dir,
            ]
        )

    wb.save("final.xlsx")


if __name__ == "__main__":
    main()
