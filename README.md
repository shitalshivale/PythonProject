Invoice Generator
This project is a simple invoice generator that reads sales data from an Excel file and generates PDF invoices for each customer using the pandas, openpyxl, jinja2, and pdfkit libraries.

Table of Contents
Features
Requirements
Installation
Usage
License
Features
Reads sales data from an Excel file.
Groups data by customer ID.
Calculates subtotal, tax, and grand total for each invoice.
Generates PDF invoices using an HTML template.
Requirements
To run this project, you need the following:

Python 3.x
Required Python packages:
pandas
openpyxl
jinja2
pdfkit
wkhtmltopdf installed on your system.
Installation
Clone the repository (if applicable):

bash
Copy code
git clone <repository-url>
cd <repository-directory>
Install required Python packages: You can install the required packages using pip:

bash
Copy code
pip install pandas openpyxl jinja2 pdfkit
Install wkhtmltopdf:

Download and install wkhtmltopdf from the official site: wkhtmltopdf Downloads.
Make sure to add the installation path (e.g., C:/Program Files/wkhtmltopdf/bin) to your system's PATH environment variable or specify the path directly in the code.
Usage
Prepare your Excel file (sales_data.xlsx) with the following columns:

Customer ID
Product ID
Product Name
Quantity
Unit Price
Example:

Copy code
| Customer ID | Product ID | Product Name | Quantity | Unit Price |
|-------------|------------|--------------|----------|------------|
| 1           | 101        | Widget A     | 2        | 10.00      |
| 1           | 102        | Widget B     | 1        | 15.00      |
| 2           | 103        | Widget C     | 3        | 20.00      |
Run the script:

bash
Copy code
python InvoiceGenerator.py
The generated PDF invoices will be saved in the same directory as the script, named invoice_customer_<customer_id>.pdf.

Troubleshooting
If you encounter an error related to wkhtmltopdf, ensure that it is installed correctly and that the path is set in the code or in your system's PATH.
If the Excel file is empty or not found, the script will print an error message.