import xlsxwriter
import pdb
import datetime
from bs4 import BeautifulSoup
import traceback
input_file_path = 'input.xml'
with open('input_file_path', 'r') as f:
    data = f.read()
soup = BeautifulSoup(data, "xml")
voucher_tags = soup.findAll('VOUCHER')


Date = []
Vch_Type = []
Vch_No = []
Transaction_Type = []
Debtor = []
Ref_Amount =[]
Ref_No = []
Ref_Type =[]
Ref_Date = []
Amount = []
Particulars = []
Amount_Varified = []
Cell_Value = ["NA",""]
T_Type = ["Parent", "Child", "Other"]











def extractData(data, transaction_type, date_arg="", vch_no=""):
    Transaction_Type.append(transaction_type)
    # It will be same for Parent,Child,Other
    if transaction_type == T_Type[0] or transaction_type == T_Type[2]:
        Ref_No.append('NA')
        Ref_Type.append('NA')
        Ref_Date.append('NA')
        Ref_Amount.append('NA')

    if transaction_type == T_Type[0]:
        Vch_Type.append('Receipt')
        Date.append(date_arg)
        Vch_No.append(vch_no)
        try:
            Debtor.append(data.find('PARTYLEDGERNAME').text)
        except:
            Debtor.append("")
        try: 
            Particulars.append(data.find('PARTYLEDGERNAME').text)
        except:
            Particulars.append("")
        try:
            Amount.append(data.find('AMOUNT').text)
        except:
            Amount.append("")

        Amount_Varified.append('YES')   # Need to calculate

    elif transaction_type == T_Type[1]:  # for child
        try:
            Ref_No.append(data.find('NAME').text)
        except:
            Ref_No.append("")
        try:
            Ref_Type.append(data.find('BILLTYPE').text)
        except:
            Ref_Type.append("")
        try:
            Ref_Date.append(data.find('DATE').text)
        except:
            Ref_Date.append("")
        try:
            Ref_Amount.append(data.find('AMOUNT').text)
        except:
            Ref_Amount.append("")

for voucher in voucher_tags:

    try:
        voucher_attrs = voucher.attrs
        if voucher_attrs['VCHTYPE'] == 'Receipt':
            date = str(datetime.datetime.strptime(
                voucher.find('DATE').text, "%Y%m%d").date())
            voucher_number = voucher.find('VOUCHERNUMBER').text
            
            # Getting only parent data out
            extractData(
                voucher, transaction_type=T_Type[0], date_arg=date, vch_no=voucher_number)

            for ledger in voucher.findAll('ALLLEDGERENTRIES.LIST'):
                

                def set_defaults():
                    particular = ledger.find('LEDGERNAME').text
                    debtor = particular
                    Particulars.append(particular)  # same for child and bank
                    Debtor.append(debtor)  # same for child and bank
                    Amount_Varified.append("NA")  # same for child and bank

                    Vch_Type.append('Receipt')
                    Date.append(date)
                    Vch_No.append(voucher_number)


                bill_locations = ledger.findAll('BILLALLOCATIONS.LIST')
                bank_locations = ledger.findAll('BANKALLOCATIONS.LIST')
                if len(bill_locations) > 1:  # for child
                    for bill_location in bill_locations:
                            Amount.append("NA")
                            set_defaults()
                            extractData(bill_location,
                                        transaction_type=T_Type[1])
                elif len(bank_locations) > 0:
                    for bank_location in bank_locations:
                        try:
                            Amount.append(bank_location.find('AMOUNT').text)
                        except:
                            Amount.append("")
                            
                        set_defaults()
                        extractData(bank_locations,
                                    transaction_type=T_Type[2])

    except Exception as e:
        print("Error : ",e)
        traceback.print_exc()
        pass


# Making final_result.xlsx
xlsxData = [
    ["Date", "Transaction Type", "Vch No.",
        "Ref No", "Ref Type", "Ref Date", "Debtor", "Ref Amount", "Amount", "Particulars", "Vch Type", "Amount Verified"]
]
workbook = xlsxwriter.Workbook('final_result.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0
for data in range(len(Vch_Type)):
    xlsxData.append([Date[data],
                    Transaction_Type[data],
                    Vch_No[data],
                    Ref_No[data],
                    Ref_Type[data],
                    Ref_Date[data],
                    Debtor[data],
                    Ref_Amount[data],
                    Amount[data],
                    Particulars[data],
                    Vch_Type[data],
                    Amount_Varified[data], ])

for Date, Vch_Type, Vch_No, Transaction_Type, Debtor, Ref_Amount, Ref_No, Ref_Type, Ref_Date, Amount, Particulars, Amount_Varified in (xlsxData):
    worksheet.write(row, col, Date)
    worksheet.write(row, col + 1, Transaction_Type)
    worksheet.write(row, col + 2, Vch_No)
    worksheet.write(row, col + 3, Ref_No)
    worksheet.write(row, col + 4, Ref_Type)
    worksheet.write(row, col + 5, Ref_Date)
    worksheet.write(row, col + 6, Debtor)
    worksheet.write(row, col + 7, Ref_Amount)
    worksheet.write(row, col + 8, Amount)
    worksheet.write(row, col + 9, Particulars)
    worksheet.write(row, col + 10, Vch_Type)
    worksheet.write(row, col + 11, Amount_Varified)
    row += 1

workbook.close()
