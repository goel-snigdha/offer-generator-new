from io import BytesIO
from docx import Document
from mailmerge import MailMerge
from datetime import datetime
import calendar

SALESFORCE_CODES = {
    "A": ("Divyam Goel", "+91 9818806094", "divyam.goel@vibrant-technik.com"),
    "B": ("Devashish Sharma", "+91 9166668377", "devashish.sharma@vibrant-technik.com"),
    "C": ("Sanjay Goel", "+91 9829055494", "sanjay.goel@vibrant-technik.com"),
    "D": ("Anmol Mathur", "+91 9772878666", ""),
    "E": ("Narasimha Santosh Batchu", "+91 7728060018", "narasimha@vibrant-technik.com"),
    "F": ("Mahendra Sanglikar", "+91 9166669595", "pune@vibrant-technik.com"),
    "G": ("Divyam Goel", "+91 9818806094", "divyam.goel@vibrant-technik.com"),
    "I": ("Sumitra Iyer", "+91 9660033816", "sumitra.iyer@vibrant-technik.com"),
    "K": ("Sumit Sharma", "+91 7756970871", "sumit.sharma@vibrant-technik.com")
}


def get_merge_fields(data):

    offer_data = data['offer_data']
    offer_code = "A"
    offer_number = offer_data["OfferNumber"]
    if len(offer_number) >= 9:
        if offer_number[8] in SALESFORCE_CODES:
            offer_code = offer_number[8]
    offer_rep_name = SALESFORCE_CODES[offer_code][0]
    offer_rep_phone = SALESFORCE_CODES[offer_code][1]
    offer_rep_email = SALESFORCE_CODES[offer_code][2]

    year = datetime.today().year
    month = datetime.today().month
    date = datetime.today().day

    expiry_year = year
    expiry_month = month + 1 if month < 12 else 1
    if expiry_month == 1:
        expiry_year += 1
    if date < 15:
        expiry_date = 15
    else:
        expiry_date = calendar.monthrange(expiry_year, expiry_month)[1]

    merge_fields = {
        "OfferNumber": offer_data['OfferNumber'],
        "Date": str(date),
        "Month": calendar.month_name[month],
        "Year": str(year),
        "ExpiryDate": str(expiry_date),
        "ExpiryMonth": calendar.month_name[expiry_month],
        "ExpiryYear": str(expiry_year),
        "OfferRep": offer_rep_name,
        "OfferRepPhone": offer_rep_phone,
        "OfferRepEmail": offer_rep_email,
        "FullName": offer_data["FullName"],
        "CompanyCity": offer_data["CompanyCity"],
        "Mobile": offer_data["Mobile"],
        "CompanyName": offer_data["CompanyName"],
        "ProjectName": offer_data["ProjectName"],
        "ProjectCity": offer_data["ProjectCity"]
    }

    return merge_fields


def convert_to_doc(document):

    temp_file = BytesIO()
    if isinstance(document, MailMerge):
        document.write(temp_file)
    else:
        document.save(temp_file)
    temp_file.seek(0)

    return Document(temp_file)
