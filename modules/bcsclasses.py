from dataclasses import dataclass

@dataclass
class Address:
        address1: str = ""
        address2: str = ""
        zipcode: str = ""
        city: str = ""
        state: str = ""
        country: str = ""

@dataclass
class Company:
        cname: str
        website: str = ""

@dataclass    
class Person:
        fname: str
        lname: str
        salutation: str  = ""
        tel: str = ""
        mobile: str = ""
        email: str = ""
        jobposition: str = ""

        def greet(self):
            print("Hello " + self.fname + " " + self.lname)

@dataclass
class Card:
    person: Person
    company: Company
    address: Address
    cardqrcode: str = ""