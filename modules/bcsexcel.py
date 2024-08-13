from modules.bcsclasses import Address, Company, Person, Card

## Excel Functions to create a DataFrame out of an excel file
import pandas as pd

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