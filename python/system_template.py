import pandas as pd
import numpy as np
import xlsxwriter
import shutil
import json

import warnings
warnings.filterwarnings("ignore", 'This pattern has match groups')
warnings.filterwarnings("ignore", 'divide by zero encountered in true_divide')
warnings.filterwarnings("ignore", 'invalid value encountered in multiply')

# SHEET COLUMNS
columns_sample_item_pricing = ['CURRENCY', 'ITEMNO', 'PRICELIST', 'DESC']

columns_sample_price_list_tax_authorities = [
    'CURRENCY', 'ITEMNO', 'PRICELIST', 'AUTHORITY', 'TAXINCL', 'TAXCLASS',
    'TXCLSDESC', 'TXAUTHDESC'
]

columns_sample_pricing_price_checks = [
    'CURRENCY', 'ITEMNO', 'PRICELIST', 'USERID', 'EXISTS', 'CGTPERCENT',
    'CLTPERCENT', 'CGTAMOUNT', 'CLTAMOUNT'
]

columns_sample_pricing_details = [
    'CURRENCY', 'ITEMNO', 'PRICELIST', 'DPRICETYPE', 'QTYUNIT', 'WEIGHTUNIT',
    'UNITPRICE'
]


# TEMPLATE FUNCTION
def system_template_fn(directory, customer_number, customer_pricelist,
                       server_path):

    file_directory = f'{directory}/{customer_number.strip()}_system.xlsx'

    # ITEM_PRICING
    IPcols = columns_sample_item_pricing
    IP = pd.DataFrame(index=np.arange(0, len(customer_pricelist)),
                      columns=IPcols)
    IP.set_index(np.arange(0, len(customer_pricelist)))
    IP['CURRENCY'] = 'ZAR'
    IP['ITEMNO'] = customer_pricelist['ITEMNO']
    IP['PRICELIST'] = customer_pricelist['PRICELIST']
    IP['DESC'] = customer_pricelist['DESC']

    # PRICE_LIST_TAX_AUTHORITIES
    PLTAcols = columns_sample_price_list_tax_authorities
    PLTA = pd.DataFrame(index=np.arange(0, len(customer_pricelist)),
                        columns=PLTAcols)
    PLTA.set_index(np.arange(0, len(customer_pricelist)))
    PLTA['CURRENCY'] = pd.NA
    PLTA['ITEMNO'] = pd.NA
    PLTA['PRICELIST'] = pd.NA
    PLTA['AUTHORITY'] = pd.NA
    PLTA['TAXINCL'] = pd.NA
    PLTA['TAXCLASS'] = pd.NA
    PLTA['TXCLSDESC'] = pd.NA
    PLTA['TXAUTHDESC'] = pd.NA

    # PRICING_PRICE_CHECKS
    PPCcols = columns_sample_pricing_price_checks
    PPC = pd.DataFrame(index=np.arange(0, len(customer_pricelist)),
                       columns=PPCcols)
    PPC.set_index(np.arange(0, len(customer_pricelist)))
    PPC['CURRENCY'] = pd.NA
    PPC['ITEMNO'] = pd.NA
    PPC['PRICELIST'] = pd.NA
    PPC['UID'] = pd.NA
    PPC['EXISTS'] = pd.NA
    PPC['CGTPERCENT'] = pd.NA
    PPC['CLTPERCENT'] = pd.NA
    PPC['CGTAMOUNT'] = pd.NA
    PPC['CLTAMOUNT'] = pd.NA

    # ITEM_PRICING_DETAILS
    PDcols = columns_sample_pricing_details
    PD = pd.DataFrame(index=np.arange(0, len(customer_pricelist)),
                      columns=PDcols)
    PD.set_index(np.arange(0, len(customer_pricelist)))
    PD['CURRENCY'] = 'ZAR'
    PD['ITEMNO'] = customer_pricelist['ITEMNO']
    PD['PRICELIST'] = customer_pricelist['PRICELIST']
    PD['DPRICETYPE'] = 1
    PD['QTYUNIT'] = 'M3'
    PD['WEIGHTUNIT'] = pd.NA
    PD['UNITPRICE'] = customer_pricelist['UNITPRICE']

    with pd.ExcelWriter(file_directory, engine='xlsxwriter') as writer:

        IP.to_excel(writer, sheet_name='Item_Pricing', index=False)
        PLTA.to_excel(writer,
                      sheet_name='Price_List_Tax_Authorities',
                      index=False)
        PPC.to_excel(writer, sheet_name='Pricing_Price_Checks', index=False)
        PD.to_excel(writer, sheet_name='Item_Pricing_Details', index=False)
        # Define name of sheets

        workbook = writer.book
        workbook.define_name("Item_Pricing",
                             f'=Item_Pricing!$A$1:$K${IP.shape[0]+4}')
        workbook.define_name(
            "Price_List_Tax_Authorities",
            f'=Price_List_Tax_Authorities!$A$1:$K${PLTA.shape[0]+4}')
        workbook.define_name(
            "Pricing_Price_Checks",
            f'=Pricing_Price_Checks!$A$1:$K${PPC.shape[0]+4}')
        workbook.define_name("Item_Pricing_Details",
                             f'=Item_Pricing_Details!$A$1:$K${PD.shape[0]+4}')

        writer.save()

    if server_path == 'none':
        pass
    else:
        try:
            shutil.copyfile(
                f'{directory}/{customer_number.strip()}_system.xlsx',
                f'{server_path}/{customer_number.strip()}_system.xlsx')
        except:
            pass