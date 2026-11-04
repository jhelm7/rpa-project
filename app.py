# Insert your python code here
while True:  
    directory = single_input("Please enter your directory")
    directory = directory.strip('"')
    directory = directory + "\\"
    try:
        filenames = os.listdir(directory)
    except:
        show_message("Error","That is not a valid directory, please try again")
        continue
    print(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message("Error", "The directory you entered is invalid. Please try again.")
Descriptions = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
Prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
product_descriptions = {}
product_prices = {}

for index,row in Descriptions.iterrows():
    unit_num = row['Unit number']
    descript = row['Description']
    product_descriptions[unit_num] = descript

for index,row in Prices.iterrows():
    unit_num1 = row['Unit number']
    price_unit = row['Price per unit']
    product_prices[unit_num1] = price_unit   
def get_order_data(invoice_num, filename, directory):
  
    update_dict = {}
    template = openpyxl.load_workbook(directory+filename)
    file = template['OrderForm']

    update_dict['InvoiceNum'] = invoice_num

    date1 = str(file['C9'].value).split(' ')
    date2 = date1[0].split('-')
    date = date2[1]+'/'+date2[2]+'/'+date2[0]

    update_dict['Date'] = str(date)

    if file['B24'].value != None:
        unitpr1 = product_prices[file['B24'].value]
        subtotal1 = file['C24'].value * unitpr1
    else:
        subtotal1 = 0
    if file['B25'].value != None:
        unitpr2 = product_prices[file['B25'].value]
        subtotal2 = file['C25'].value * unitpr2
    else:
        subtotal2 = 0
        unitpr2 = 0
    if file['B26'].value != None:
        unitpr3 = product_prices[file['B26'].value]
        subtotal3 = file['C26'].value * unitpr3
    else:
        subtotal3 = 0
        unitpr3 = 0
    if file['B27'].value != None:
        unitpr4 = product_prices[file['B27'].value]
        subtotal4 = file['C27'].value * unitpr4
    else:
        subtotal4 = 0
        unitpr4 = 0
    if file['B28'].value != None:
        unitpr5 = product_prices[file['B28'].value]
        subtotal5 = file['C28'].value * unitpr5
    else:
        subtotal5 = 0
        unitpr5 = 0
  
    total1 = subtotal1+subtotal2+subtotal3+subtotal4+subtotal5

    update_dict['Total'] = f'${total1:,.2f}'
    update_dict['CustomerAddress1'] = file['C16'].value
    update_dict['CustomerName'] = file['C14'].value
    update_dict['CustomerCompany'] = file['C15'].value
    update_dict['CustomerAddress2'] = str(file['C17'].value) + ', ' + str(file['C18'].value) + ', ' + str(file['C19'].value)

    if unitpr1 == 0:
        unitpr1 = ''
    else:
        unitpr1 = unitpr1
    update_dict['Price0'] = f'${unitpr1:,.2f}'
    update_dict['Quantity0'] = file['C24'].value

    if subtotal1 == 0:
        subtotal1 = ''
    else:
        subtotal1 = subtotal1
    update_dict['Amount0'] = f'${subtotal1:,.2f}'

    if unitpr2 == 0:
        unitpr2 = ''
    else:
        unitpr2 = f'${unitpr2:,.2f}'
    update_dict['Price1'] = unitpr2

    if file['C25'].value != None:
        update_dict['Quantity1'] = file['C25'].value
    else:
        update_dict['Quantity1'] = ''

    if subtotal2 == 0:
        subtotal2 = ''
    else:
        subtotal2 = f'${subtotal2:,.2f}'
    update_dict['Amount1'] = subtotal2

    if unitpr3 == 0:
        unitpr3 = ''
    else:
        unitpr3 = f'${unitpr3:,.2f}'

    update_dict['Price2'] = unitpr3

    if file['C26'].value != None:
        update_dict['Quantity2'] = file['C26'].value
    else:
        update_dict['Quantity2'] = ''

    if subtotal3 == 0:
        subtotal3 = ''
    else:
        subtotal3 = f'${subtotal3:,.2f}'
    update_dict['Amount2'] = subtotal3

    if unitpr4 == 0:
        unitpr4 = ''
    else:
        unitpr4 = f'${unitpr4:,.2f}'
    update_dict['Price3'] = unitpr4
    if file['C27'].value != None:
        update_dict['Quantity3'] = file['C27'].value
    else:
        update_dict['Quantity3'] = ''

    if subtotal4 == 0:
        subtotal4 = ''
    else:
        subtotal4 = f'${subtotal4:,.2f}'
    update_dict['Amount3'] = subtotal4

    if unitpr5 == 0:
        unitpr5 = ''
    else:
        unitpr5 = f'${unitpr5:,.2f}'
    update_dict['Price4'] = unitpr5
    if file['C28'].value != None:
        update_dict['Quantity4'] = file['C28'].value
    else:
        update_dict['Quantity4'] = ''

    if subtotal5 == 0:
        subtotal5 = ''
    else:
        subtotal5 = f'${subtotal5:,.2f}'
    update_dict['Amount4'] = subtotal5

    if file['B27'].value != None:
        update_dict['Description3'] = product_descriptions[file['B27'].value]
    else:
        update_dict['Description3'] = ''
    if file['B26'].value != None:
        update_dict['Description2'] = product_descriptions[file['B26'].value]
    else:
        update_dict['Description2'] = ''
    if file['B24'].value != None:
        update_dict['Description0'] = product_descriptions[file['B24'].value]
    else:
        update_dict['Description0'] = ''
    if file['B25'].value != None:
        update_dict['Description1'] = product_descriptions[file['B25'].value]
    else:
        update_dict['Description1'] = ''
    if file['B28'].value != None:
        update_dict['Description4'] = product_descriptions[file['B28'].value]
    else:
        update_dict['Description4'] = ''

    name = str(file['C14'].value)
    email = str(file['C13'].value)
    date3 = str(update_dict['Date'])

    return update_dict, name, email, date3
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)
def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "johnnyhelm03@gmail.com"
    gmail_password = open('gmail_password.txt').read()

# List of recipients

    recipients = email

# Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

# Body of the email

    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

# Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

# Add attachment

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
#path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
        
# Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

# Send the email

    mail_server.send_message(message)
    mail_server.quit()
filenames = os.listdir(directory)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1