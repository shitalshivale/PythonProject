import pandas as pd
import openpyxl
from jinja2 import Template
import pdfkit
from datetime import datetime

config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')

#SOLUTION-------------------------------------------------------------
#Configuration of constant
Company_Name="Tech Solutions Inc."
Company_Address="123 Main Street, Anytown, CA 91234"
Invoice_Prefix= "INV-"
Tax_Rate=0.05

#read data from excel  #Use openpyxl to read the data from sales_data.xlsx into a Pandas DataFrame.
def read_excel_data(file):
    try:
        df=pd.read_excel(file, engine='openpyxl')
        return df
    except Exception as e:
        print(f"Error in reading excel file: {e}")
        
#generate invoices as per customer id
def generate_invoice(df):
    grouped_df = df.groupby('Customer ID')
    Invoice_Number =1

    for customer_id, group in grouped_df:
        group['Total Price']= group['Quantity']* group['Unit Price']
        subtotal=group['Total Price'].sum()
        tax=subtotal*Tax_Rate
        grand_total=subtotal+tax
        
        invoice_data={
            'company_name': Company_Name,
            'company_address': Company_Address,
            'invoice_number': f"{Invoice_Prefix}{datetime.now().strftime('%Y-%m-%d')}-{Invoice_Number:03d}",
            'invoice_date': datetime.now().strftime('%Y-%m-%d'),
            'customer_id': customer_id,
            'products': group.to_dict(orient='records'),
            'subtotal': subtotal,
            'tax': tax,
            'grand_total': grand_total
        }

        generate_pdf(invoice_data)
        Invoice_Number +=1
        

#Generate pdf invoices by using jinja HTML
def generate_pdf(data):
    html_template = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Invoice</title>
        <style>
            body { font-family: Arial, sans-serif; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #000; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
        </style>
    </head>
    <body>
        <h1>{{ company_name }}</h1>
        <p>{{ company_address }}</p>
        <p>Invoice Number: {{ invoice_number }}</p>
        <p>Invoice Date: {{ invoice_date }}</p>
        <p>Customer ID: {{ customer_id }}</p>

        <h3>Product Details:</h3>
        <table>
            <thead>
                <tr>
                    <th>Product ID</th>
                    <th>Product Name</th>
                    <th>Quantity</th>
                    <th>Unit Price</th>
                    <th>Total Price</th>
                </tr>
            </thead>
            <tbody>
                {% for product in products %}
                <tr>
                    <td>{{ product['Product ID'] }}</td>
                    <td>{{ product['Product Name'] }}</td>
                    <td>{{ product['Quantity'] }}</td>
                    <td>{{ product['Unit Price'] }}</td>
                    <td>{{ product['Total Price'] }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>

        <h3>Totals:</h3>
        <p>Subtotal: {{ subtotal }}</p>
        <p>Tax (5%): {{ tax }}</p>
        <p>Grand Total: {{ grand_total }}</p>
    </body>
    </html>
    """

    #create jinja2 template from above html template
    template = Template(html_template)
    html_content = template.render(data)

    pdf_file_name = f"invoice_customer_{data['customer_id']}.pdf"
    pdfkit.from_string(html_content, pdf_file_name, configuration=config)
    print(f"Generated invoice: {pdf_file_name}")

def main():
    data_file = 'sales_data.xlsx'
    sale_data= read_excel_data(data_file)

    if sale_data is None:
        print(f"File is empty:{data_file}")
    else:
        generate_invoice(sale_data)

# Entry point of the script
if __name__ == "__main__":
    main()




    