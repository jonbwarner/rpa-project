import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
directory = os.getcwd()+'\\'
directory

while True:
    directory=single_input('Please input the file directory.')
    directory=directory.strip('"').strip('"')
    if directory[-1]!= '\\':
        directory = directory + "\\"
        
    filesindir=os.listdir(directory)

    if 'ProductData.xlsx' and 'MasterInvoice.pdf' and 'Order1.xlsx' and 'Order2.xlsx' and 'Order3.xlsx' and 'Order4.xlsx' in filesindir:
        break
    else:
        show_message('Error', "The directory you entered is not correct. Please try again.")
        
        # Insert your python code here
      # INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions= {}
product_prices={}



df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')

for index,row in df.iterrows():
    product_descriptions [row['Unit number']]= row['Description']

df1= pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')

for index,row in df1.iterrows():
    product_prices [row['Unit number']]= row['Price per unit']
    
product_descriptions

# INSERT ANSWER FOR QUESTION 3 HERE
df=get_pdf(directory+'MasterInvoice.pdf')
display(df)

get_pdf(directory + 'MasterInvoice.pdf')
def get_order_data(invoice_num, filename, directory):
    order_data = {}
    template = openpyxl.load_workbook(directory + filename, data_only = True)
    data = template['OrderForm']
    name = data['C14'].value
    email = data['C13'].value
    date = data['C9'].value.strftime('%m/%d/%Y')
    order_data['InvoiceNum'] = f'{invoice_num}'
    order_data['Date'] = date
    order_data['CustomerAddress1'] = data['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = data['C15'].value
    order_data['CustomerAddress2'] = f"{data['C17'].value}, {data['C18'].value} {data['C19'].value}"
    # product 1
    order_data['Description0'] = product_descriptions[data['B24'].value]
    price0 = product_prices[data['B24'].value]
    order_data['Price0'] = f'${price0:,.2f}'
    quantity0 = data['C24'].value
    order_data['Quantity0'] = str(quantity0)
    amount0 = price0*quantity0
    order_data['Amount0'] = f'${amount0:,.2f}'
    order_data['Total'] = f'${amount0:,.2f}'
    # product 2
    if data['B25'].value is not None:
        order_data['Description1'] = product_descriptions[data['B25'].value]
        price1 = product_prices[data['B25'].value]
        order_data['Price1'] = f'${price1:,.2f}'
        quantity1 = data['C25'].value
        order_data['Quantity1'] = str(quantity1)
        amount1 = price1*quantity1
        order_data['Amount1'] = f'${amount1:,.2f}'
        total1 = amount0 + amount1
        order_data['Total'] = f'${total1:,.2f}'
    else:
        order_data['Description1'] = ''
        order_data['Price1'] = ''
        order_data['Quantity1'] = ''
        order_data['Amount1'] = ''
    # product 3
    if data['B26'].value is not None:
        order_data['Description2'] = product_descriptions[data['B26'].value]
        price2 = product_prices[data['B26'].value]
        order_data['Price2'] = f'${price2:,.2f}'
        quantity2 = data['C26'].value
        order_data['Quantity2'] = str(quantity2)
        amount2 = price2*quantity2
        order_data['Amount2'] = f'${amount2:,.2f}'
        total2 = amount0 + amount1 + amount2
        order_data['Total'] = f'${total2:,.2f}'
    else:
        order_data['Description2'] = ''
        order_data['Price2'] = ''
        order_data['Quantity2'] = ''
        order_data['Amount2'] = ''
    return order_data, name, email, date
order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
order_data

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)
    

create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
gmail_user = "jbwarner9@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email (invoice_num, name, email, date, directory):
    
    
    
    recipients= [email]
    
    subject= f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    
    body= f"""Dear {name},
Attached is your invoice dated {date}.

Thank you for business.

Python Tech Co.
18 Python Drive
Pleasant Grove, UT 84602
Phone: (801) 555-2372
Email: sales@pythontechco.com"""


    message = EmailMessage()
    message['From']= gmail_user
    message['To']= recipients
    message['Subject']= subject
    message.set_content(body)
    
    path_to_file = os.getcwd()+ f'\\Invoice {invoice_num}.pdf'
    message= add_attachment(path_to_file, message)
    
    mail_server= smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    
    mail_server.send_message(message)
    mail_server.quit()
    
send_email(8515, 'Thomas Holt', 'jbwarner9@gmail.com', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE

filenames=os.listdir(directory)
filenames.sort()
invoice_num=8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filenames:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num= invoice_num + 1
        

 


