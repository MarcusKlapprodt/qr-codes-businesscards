# Business Card Generator

This Python application generates business cards in PDF format from an Excel file. Each row in the Excel file corresponds to a separate business card (with a qr code), and each card is placed on a separate page in the resulting PDF. The PDF is then ready to be sent to a print shop for commissioning the printing of the business cards.

## Features

- Converts an Excel file into a PDF where each row is a separate business card.
- Each business card is placed on a separate page in the PDF.
- Designed for easy bulk creation of business cards.

## Requirements

- Python 3.12
- Required Python modules listed in `requirements.txt`

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/MarcusKlapprodt/qr-codes-businesscards.git
    cd qr-codes-businesscards
    ```

2. **Install the required dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

## Usage

To generate a PDF of business cards, use the following command:

```bash
python bcspdf.py "name_of_the_excel_file.xlsx"
```

### Excel File Format

The Excel file must follow a specific format to ensure that the application works correctly. The column headers must be in the exact order and spelled exactly as follows:

```
Username | Title | First Name | Last Name | Organization | Position | Phone | Mobile | Fax | Email | Website | Street | Street 2 | P.O. Box | City | Zip Code | State | Country | Business Card URL
```

- **Each column** in the Excel file represents a different piece of information that will be included on the business card.
- **Each row** in the Excel file will be converted into a separate business card in the PDF.

### Example

Here is a quick example of how the Excel file should be structured:

| Username | Title | First Name | Last Name | Organization | Position | Phone | Mobile | Fax | Email | Website | Street | Street 2 | P.O. Box | City | Zip Code | State | Country | Business Card URL |
|----------|-------|------------|-----------|--------------|----------|-------|--------|-----|-------|---------|--------|----------|---------|------|----------|-------|---------|------------------|
| jdoe     | Mr.   | John       | Doe       | Acme Corp    | Manager  | 123-456-7890 | 098-765-4321 | 555-5555 | john.doe@example.com | www.example.com | 123 Elm St | Suite 100 |  | Springfield | 62701 | IL | USA |  |

## Contributing

Feel free to submit issues, fork the repository, and make a pull request if you have any improvements or suggestions.

## License

This project is licensed under the GPL-3.0 license - see the [LICENSE](https://github.com/MarcusKlapprodt/qr-codes-businesscards?tab=GPL-3.0-1-ov-file#readme) file for details.
