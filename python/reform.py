# MODULES
import pandas as pd
import numpy as np
import json
import warnings
import os
warnings.filterwarnings("ignore", 'This pattern has match groups')

import s5_ordersheet

# GET WORKING DIRECTORY #
# ///////////////////// #
workdir = os.getcwd()

#  IMPORT MASTER JSON AND CHANGE TO DF #
with open(f'{workdir}/python/templates/masterPricelistTemplate.json',
          'r') as jsonfile:
    master_json = json.load(jsonfile)
# COLUMNS/INDEX/DATA
columns = master_json['columns']
index = list(master_json['data'].keys())
data = master_json['data'].values()

# CREATE MASTER PRICELIST DF
master_price = pd.DataFrame(data, columns=columns, index=index)


# REFORMAT THE LAYMAN PRICELIST FUNCTION #
##########################################
def reformat_layman(pricelist_number, customer_number, df):
    # DICTIONARY TO SEPERATE PRODUCT NUMBERS
    D = {}
    for i, d in df.iterrows():
        D[str(i) + 'U'] = (d['IC_UNTREATED'], d['PRICE UNTREATED'],
                           d['R_FACTOR'], d['BUNDLE SIZE'])
        D[str(i) + 'T'] = (d['IC_TREATED'], d['PRICE TREATED'], d['R_FACTOR'],
                           d['BUNDLE SIZE'])

    customer_pricelist = master_price.copy()
    customer_pricelist['R/METER UNTREATED'] = 0
    customer_pricelist['R/METER TREATED'] = 0
    customer_pricelist['BUNDLE SIZE'] = 'NA'
    customer_pricelist['M3 TREATED'] = 0
    customer_pricelist['M3 UNTREATED'] = 0
    customer_pricelist['CURRENCY'] = 'ZAR'
    customer_pricelist['PRICELIST'] = pricelist_number

    # LOOP THOUGHT THE SEPERATED TREATED/UNTREATED PRODUCTS AN DALOWCATE PRICES ON THE MASTERPRICELIST FILE
    for k, v in D.items():

        # ADD THE PRICES TO THE CORRECT ITEMNO UNTREATED
        if k.endswith('U'):

            customer_pricelist.loc[v[0], ['UNITPRICE']] = v[1]
            customer_pricelist.loc[v[0], ['R/METER UNTREATED']] = round(
                int(v[1]) * v[2], 2)
            customer_pricelist.loc[v[0], ['M3 UNTREATED']] = v[1]
            customer_pricelist.loc[v[0], ['BUNDLE SIZE']] = v[3]

        # ADD THE PRICES TO THE CORRECT ITEMNO TREATED
        elif k.endswith('T'):

            customer_pricelist.loc[v[0], ['UNITPRICE']] = v[1]
            customer_pricelist.loc[v[0], ['R/METER TREATED']] = round(
                int(v[1]) * v[2], 2)
            customer_pricelist.loc[v[0], ['M3 TREATED']] = v[1]
            customer_pricelist.loc[v[0], ['BUNDLE SIZE']] = v[3]

    # REPLACE ALL THE ZERO WITH NAN VALUES AND SORT THE DF UNITPRICE WITH VALUES
    customer_pricelist.replace(0, np.nan, inplace=True)
    customer_pricelist = customer_pricelist.dropna(axis=0,
                                                   subset=['UNITPRICE'])
    # customer_pricelist = customer_pricelist.dropna(axis=0, subset=['DESC'])

    # RETURN ITEMS TO PROCESS S5 AND SYSTEM TEMPLATE
    return {
        'customer_number': customer_number,
        'customer_pricelist': customer_pricelist
    }
