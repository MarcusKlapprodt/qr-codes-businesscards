import pandas as pd
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import mm
from reportlab_qr_code import qr_draw

## Excel Functions to create a DataFrame out of an excel file

def fetch_excelfile(file):
    with open(file):
        cards_file = pd.read_excel(file).fillna("")
    return cards_file

def create_company_instance(pandas_row):
    pdr = pandas_row
    return Company(pdr["Organization"], pdr["Website"])

def create_address_instance(pandas_row):
    pdr = pandas_row
    return Address(pdr["Street"], pdr["Street 2"], pdr["Zip Code"], pdr["City"], pdr["State"], pdr["Country"])

def create_person_instance(pandas_row):
    pdr = pandas_row
    return Person(pdr["First Name"], pdr["Last Name"], pdr["Title"], pdr["Phone"], pdr["Mobile"], pdr["Email"], pdr["Position"])


## Collection of functions to create a PDF

def create_one_pdf(card):
    # define canvas size
    canvas = Canvas("mycard.pdf", pagesize=(87.0*mm,57.0*mm))
    
    # Write Name as Headline
    canvas.setFont('Helvetica-Bold', 12)
    canvas.drawString(4*mm, 45*mm, f'{card.person.fname} {card.person.lname}')

    # Job position
    canvas.setFont('Helvetica', 8)
    canvas.drawString(4*mm, 40*mm, card.person.jobposition)

    # Address Block
    canvas.setFont('Helvetica', 8)
    canvas.drawString(4*mm, 28*mm, card.company.cname)
    canvas.drawString(4*mm, 24*mm, card.address.address1)
    # if card.address.address2 != "" or "":
        # canvas.drawString(4*mm, *mm, card.address.address2)
    canvas.drawString(4*mm, 20*mm, f'{card.address.zipcode}, {card.address.city}, {card.address.state}')
    canvas.drawString(4*mm, 16*mm, card.address.country)
    canvas.drawString(4*mm, 12*mm, card.person.email)
    if card.person.tel != "":
        canvas.drawString(4*mm, 8*mm, f'tel.: {card.person.tel}')
    if card.person.mobile != "":
        canvas.drawString(4*mm, 4*mm, f'mobile: {card.person.mobile}')

    ## Create the QR Code 
    qr_draw(canvas, "Hello world", x="55mm", y="23mm", size="30mm")

    # Finalize Business Card PDF
    canvas.showPage()
    canvas.save()

def create_pdf_body(canvas, card):
    
    # Write Name as Headline
    canvas.setFont('Helvetica-Bold', 12)
    canvas.drawString(4*mm, 45*mm, f'{card.person.fname} {card.person.lname}')

    # Job position
    canvas.setFont('Helvetica', 8)
    canvas.drawString(4*mm, 40*mm, card.person.jobposition)

    # Address Block
    canvas.setFont('Helvetica', 8)
    canvas.drawString(4*mm, 28*mm, card.company.cname)
    canvas.drawString(4*mm, 24*mm, card.address.address1)
    # if card.address.address2 != "" or "":
        # canvas.drawString(4*mm, *mm, card.address.address2)
    canvas.drawString(4*mm, 20*mm, f'{card.address.zipcode}, {card.address.city}, {card.address.state}')
    canvas.drawString(4*mm, 16*mm, card.address.country)
    canvas.drawString(4*mm, 12*mm, card.person.email)
    if card.person.tel != "":
        canvas.drawString(4*mm, 8*mm, f'tel.: {card.person.tel}')
    if card.person.mobile != "":
        canvas.drawString(4*mm, 4*mm, f'mobile: {card.person.mobile}')

    ## Create the QR Code 
    qr_draw(canvas, card.cardqrcode, x="55mm", y="23mm", size="30mm")


def finalize_bc_pdf(excel_file):
    # Get Excel File and create a Pandas Data Frame
    e_file = fetch_excelfile(excel_file)

    # Define a Canvas for the Business-Cards-PDF
    canvas = Canvas("mycards.pdf", pagesize=(87.0*mm,57.0*mm))
    
    for key, row in e_file.iterrows():
        # Create Cards Instances
        comp1 = create_company_instance(row)
        add1 = create_address_instance(row)
        pers1 = create_person_instance(row)

        # Create Card Instance
        card = Card(pers1,comp1,add1,row["Business Card URL"])

        # Create PDF Page
        create_pdf_body(canvas, card)

        # End of page
        canvas.showPage()

    canvas.save()
        
    return None