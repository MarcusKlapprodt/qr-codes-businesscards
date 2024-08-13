# from dataclasses import dataclass
import sys
from modules.bcsclasses import Address, Company, Person, Card
from modules.bcspdf import create_pdf_body, finalize_bc_pdf
from modules.bcsexcel import fetch_excelfile, create_company_instance, create_address_instance, create_person_instance

def main():
    file = finalize_bc_pdf(sys.argv[1])
    #("Business-Cards-List.xlsx")

if __name__=="__main__":
    main()