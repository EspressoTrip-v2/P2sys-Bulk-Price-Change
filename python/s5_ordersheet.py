import pandas as pd
import numpy as np
import xlsxwriter
import warnings
import shutil
import platform
warnings.filterwarnings("ignore",
                        'This pattern has match grouserver_filepathps')
warnings.filterwarnings("ignore", 'divide by zero encountered in true_divide')
warnings.filterwarnings("ignore", 'invalid value encountered in multiply')


def create_s5_ordersheet(directory, customer_number, customer_pricelist,
                         server_path):

    # CREATE THE COLUMNS TO BE USED IN THE ORDERSHEET #
    ###################################################

    # TREATED
    columnsT = [
        'ITEM NUMBER', 'DESCRIPTION', 'BUNDLE \n SIZE', 'M3 TREATED \n PRICE',
        'TREATED BUNDLE \n PRICE', 'R/METER TREATED \n PRICE',
        '(BUNDLE) ORDER \n QUANTITY', 'TOTAL \n AMOUNT'
    ]
    # UNTREATED
    columnsU = [
        'ITEM NUMBER', 'DESCRIPTION', 'BUNDLE \n SIZE',
        'M3 UNTREATED \n PRICE', 'UNTREATED BUNDLE \n PRICE',
        'R/METER UNTREATED \n PRICE', '(BUNDLE) ORDER \n QUANTITY',
        'TOTAL \n AMOUNT'
    ]

    # LISTS TO BE ABLE TO INSERT BLANK LINES BETWEEN PRODUCT LINES #
    ################################################################
    desc_038 = [
        '(038 x 038)', '(038 x 050)', '(038 x 076)', '(038 x 114)',
        '(038 x 152)', '(038 x 228)'
    ]
    desc_050 = ['(050 x 076)', '(050 x 114)', '(050 x 152)', '(050 x 228)']
    desc_076 = ['(076 x 114)', '(076 x 152)', '(076 x 228)']
    nan_row = ['', '', '', '', np.nan, np.nan, np.nan, np.nan]

    # RESET THE INDEX
    customer_pricelist.reset_index(inplace=True)
    customer_pricelist.rename(columns={'index': 'ITEMNO'}, inplace=True)

    # 'ITEMNO' ,'DESC', 'UNITPRICE', 'CONVERSION', 'R/METER UNTREATED','R/METER TREATED', 'BUNDLE SIZE', 'M3 TREATED', 'M3 UNTREATED', 'CURRENCY','PRICELIST'

    # SET ALL NUMERIC NAN TO 0 & COLUMNS TO INTEGER
    customer_pricelist['M3 TREATED'].fillna(0, inplace=True)
    customer_pricelist['M3 TREATED'] = customer_pricelist['M3 TREATED'].astype(
        int)
    customer_pricelist['M3 UNTREATED'].fillna(0, inplace=True)
    customer_pricelist['M3 UNTREATED'] = customer_pricelist[
        'M3 UNTREATED'].astype(int)

    # CONVERT THE VALUES TO INTEGER/FLOATS
    customer_pricelist['UNITPRICE'] = customer_pricelist['UNITPRICE'].astype(
        int)

    # ADD THE BUNDLE PRICE CONVERSION
    customer_pricelist['BUNDLE PRICE'] = (customer_pricelist['UNITPRICE'].values\
                                            *  (1/customer_pricelist['CONVERSION'].values)).round(2)

    # customer_pricelist = customer_pricelist[~customer_pricelist['DESC'].str.
    #                                         contains('XXX')]

    # 038 PRODUCT LINE
    _038 = customer_pricelist[customer_pricelist['DESC'].str.contains(
        '(PINE: 038)')]

    # TREATED
    _038T = _038[_038['ITEMNO'].str.endswith('T')].copy()
    _038T = _038T[[
        'ITEMNO', 'DESC', 'BUNDLE SIZE', 'M3 TREATED', 'BUNDLE PRICE',
        'R/METER TREATED'
    ]]
    _038T['ORDER QTY M3'] = np.nan
    _038T['AMOUNT R'] = np.nan
    _038T = _038T[_038T['M3 TREATED'] > 0]

    _038T.sort_values(by=['DESC'], inplace=True, axis=0)
    _038T['DESC'] = _038T['DESC'].str.replace('PINE: ', '').str.replace(
        'LONG SABS S5', '').str.replace('SHORT SABS S5', '')

    # BLANK INSERTION LOOP TO SEPERATE PRODUCTS
    for i in desc_038:
        _038T.reset_index(inplace=True, drop=True)
        idx = list(_038T[_038T['DESC'].str.contains(i)].index)
        if len(idx) < 1:
            continue
        else:
            df1 = _038T.iloc[:idx[-1] + 1, :].copy()
            df1.loc[idx[-1] + 1] = nan_row

            df2 = _038T.iloc[idx[-1] + 1:, :].copy()
            _038T = pd.concat([df1, df2])

    # ADD THE TREATED COLUMN NAMES
    _038T.columns = columnsT

    # ROW LENGTH
    _038T_rownum = _038T.shape[0]

    # UNTREATED
    _038U = _038[~_038['ITEMNO'].str.endswith('T')].copy()
    _038U = _038U[_038U['DESC'].str.contains('SABS S5')]

    _038U = _038U[[
        'ITEMNO', 'DESC', 'BUNDLE SIZE', 'M3 UNTREATED', 'BUNDLE PRICE',
        'R/METER UNTREATED'
    ]]
    _038U['ORDER QTY M3'] = np.nan
    _038U['AMOUNT R'] = np.nan
    _038U = _038U[_038U['M3 UNTREATED'] > 0]
    _038U.sort_values(by=['DESC'], inplace=True, axis=0)

    _038U.sort_values(by=['DESC'], inplace=True, axis=0)
    _038U['DESC'] = _038U['DESC'].str.replace('PINE: ','').str.replace('LONG SABS S5','')\
    .str.replace('SHORT SABS S5','')

    # BLANK INSERTION LOOP TO SEPERATE PRODUCTS
    for i in desc_038:
        _038U.reset_index(inplace=True, drop=True)
        idx = list(_038U[_038U['DESC'].str.contains(i)].index)
        if len(idx) < 1:
            continue
        else:
            df1 = _038U.iloc[:idx[-1] + 1, :].copy()
            df1.loc[idx[-1] + 1] = nan_row

            df2 = _038U.iloc[idx[-1] + 1:, :].copy()
            _038U = pd.concat([df1, df2])
    # ADD UNTREATED COLUMN NAMES
    _038U.columns = columnsU

    # ROW LENGTH
    _038U_rownum = _038U.shape[0]

    # 050 PRODUCT LINE
    _050 = customer_pricelist[customer_pricelist['DESC'].str.contains(
        '(PINE: 050)')]

    # TREATED
    _050T = _050[_050['ITEMNO'].str.endswith('T')].copy()
    _050T = _050T[[
        'ITEMNO', 'DESC', 'BUNDLE SIZE', 'M3 TREATED', 'BUNDLE PRICE',
        'R/METER TREATED'
    ]]
    _050T['ORDER QTY M3'] = np.nan
    _050T['AMOUNT R'] = np.nan
    _050T = _050T[_050T['M3 TREATED'] > 0]
    _050T.sort_values(by=['DESC'], inplace=True, axis=0)

    _050T.sort_values(by=['DESC'], inplace=True, axis=0)
    _050T['DESC'] = _050T['DESC'].str.replace('PINE: ','').str.replace('LONG SABS S5','')\
    .str.replace('SHORT SABS S5','')

    # BLANK INSERTION LOOP TO SEPERATE PRODUCTS
    for i in desc_050:
        _050T.reset_index(inplace=True, drop=True)
        idx = list(_050T[_050T['DESC'].str.contains(i)].index)
        if len(idx) < 1:
            continue
        else:
            df1 = _050T.iloc[:idx[-1] + 1, :].copy()
            df1.loc[idx[-1] + 1] = nan_row

            df2 = _050T.iloc[idx[-1] + 1:, :].copy()

            _050T = pd.concat([df1, df2])
    # ADD TREATED COLUMN NAMES
    _050T.columns = columnsT

    # ROW LENGTH
    _050T_rownum = _050T.shape[0]

    # UNTREATED
    _050U = _050[~_050['ITEMNO'].str.endswith('T')].copy()
    _050U = _050U[_050U['DESC'].str.contains('SABS S5')]

    _050U = _050U[[
        'ITEMNO', 'DESC', 'BUNDLE SIZE', 'M3 UNTREATED', 'BUNDLE PRICE',
        'R/METER UNTREATED'
    ]]
    _050U['ORDER QTY M3'] = np.nan
    _050U['AMOUNT R'] = np.nan
    _050U = _050U[_050U['M3 UNTREATED'] > 0]
    _050U.sort_values(by=['DESC'], inplace=True, axis=0)

    _050U.sort_values(by=['DESC'], inplace=True, axis=0)
    _050U['DESC'] = _050U['DESC'].str.replace('PINE: ','').str.replace('LONG SABS S5','')\
    .str.replace('SHORT SABS S5','')

    # BLANK INSERTION LOOP TO SEPERATE PRODUCTS
    for i in desc_050:
        _050U.reset_index(inplace=True, drop=True)
        idx = list(_050U[_050U['DESC'].str.contains(i)].index)
        if len(idx) < 1:
            continue
        else:

            df1 = _050U.iloc[:idx[-1] + 1, :].copy()
            df1.loc[idx[-1] + 1] = nan_row

            df2 = _050U.iloc[idx[-1] + 1:, :].copy()
            _050U = pd.concat([df1, df2])
    # ADD UNTREATED COLUMN NAMES
    _050U.columns = columnsU

    # ROW LENGTH
    _050U_rownum = _050U.shape[0]

    # 076 PRODUCT LINE
    _076 = customer_pricelist[customer_pricelist['DESC'].str.contains(
        '(PINE: 076)')]

    # TREATED
    _076T = _076[_076['ITEMNO'].str.endswith('T')].copy()
    _076T = _076T[[
        'ITEMNO', 'DESC', 'BUNDLE SIZE', 'M3 TREATED', 'BUNDLE PRICE',
        'R/METER TREATED'
    ]]
    _076T['ORDER QTY M3'] = np.nan
    _076T['AMOUNT R'] = np.nan
    _076T = _076T[_076T['M3 TREATED'] > 0]
    _076T.sort_values(by=['DESC'], inplace=True, axis=0)

    _076T.sort_values(by=['DESC'], inplace=True, axis=0)
    _076T['DESC'] = _076T['DESC'].str.replace('PINE: ','').str.replace('LONG SABS S5','')\
    .str.replace('SHORT SABS S5','')

    # BLANK INSERTION LOOP TO SEPERATE PRODUCTS
    for i in desc_076:
        _076T.reset_index(inplace=True, drop=True)
        idx = list(_076T[_076T['DESC'].str.contains(i)].index)
        if len(idx) < 1:
            continue
        else:

            df1 = _076T.iloc[:idx[-1] + 1, :].copy()
            df1.loc[idx[-1] + 1] = nan_row

            df2 = _076T.iloc[idx[-1] + 1:, :].copy()

            _076T = pd.concat([df1, df2])
    # ADD TREATED COLUMN NAMES
    _076T.columns = columnsT

    # ROW LENGTH
    _076T_rownum = _076T.shape[0]

    # UNTREATED
    _076U = _076[~_076['ITEMNO'].str.endswith('T')].copy()
    _076U = _076U[_076U['DESC'].str.contains('SABS S5')]

    _076U = _076U[[
        'ITEMNO', 'DESC', 'BUNDLE SIZE', 'M3 UNTREATED', 'BUNDLE PRICE',
        'R/METER UNTREATED'
    ]]
    _076U['ORDER QTY M3'] = np.nan
    _076U['AMOUNT R'] = np.nan
    _076U = _076U[_076U['M3 UNTREATED'] > 0]
    _076U.sort_values(by=['DESC'], inplace=True, axis=0)

    _076U.sort_values(by=['DESC'], inplace=True, axis=0)
    _076U['DESC'] = _076U['DESC'].str.replace('PINE: ','').str.replace('LONG SABS S5','')\
    .str.replace('SHORT SABS S5','')

    # BLANK INSERTION LOOP TO SEPERATE PRODUCTS
    for i in desc_076:
        _076U.reset_index(inplace=True, drop=True)
        idx = list(_076U[_076U['DESC'].str.contains(i)].index)
        if len(idx) < 1:
            continue
        else:

            df1 = _076U.iloc[:idx[-1] + 1, :].copy()
            df1.loc[idx[-1] + 1] = nan_row

            df2 = _076U.iloc[idx[-1] + 1:, :].copy()

            _076U = pd.concat([df1, df2])
    # ADD UNTREATED COLUMN NAMES
    _076U.columns = columnsU

    # ROW LENGTH
    _076U_rownum = _076U.shape[0]

    # CREATE XLSX WRITER
    writer = pd.ExcelWriter(f'{directory}/S5_{customer_number.strip()}.xlsx',
                            engine='xlsxwriter')

    # CONVERT THE DATAFRAME TO EXCEL
    _038T.to_excel(writer, sheet_name='Treated 38mm', startrow=3, index=False)
    _038U.to_excel(writer,
                   sheet_name='Untreated 38mm',
                   startrow=3,
                   index=False)

    _050T.to_excel(writer, sheet_name='Treated 50mm', startrow=3, index=False)
    _050U.to_excel(writer,
                   sheet_name='Untreated 50mm',
                   startrow=3,
                   index=False)

    _076T.to_excel(writer, sheet_name='Treated 76mm', startrow=3, index=False)
    _076U.to_excel(writer,
                   sheet_name='Untreated 76mm',
                   startrow=3,
                   index=False)

    # GET THE WRITER WORKBOOK
    workbook = writer.book

    # ASSIGN THE WORKSHEET NAMES
    worksheet1 = writer.sheets['Treated 38mm']
    worksheet2 = writer.sheets['Untreated 38mm']
    worksheet3 = writer.sheets['Treated 50mm']
    worksheet4 = writer.sheets['Untreated 50mm']
    worksheet5 = writer.sheets['Treated 76mm']
    worksheet6 = writer.sheets['Untreated 76mm']

    # COLUMN ATTRIBUTES AND FORMATTING
    column_format1 = workbook.add_format()
    column_format1.set_align('center')
    column_format1.set_align('vcenter')
    column_format2 = workbook.add_format()
    column_format2.set_align('center')
    column_format2.set_align('vcenter')
    column_format2.set_text_wrap()
    column_format2.set_num_format('_(###0.00_);_(\(###0.00\);_(" "??_);_(@_)')

    # WORKSHEET 1
    worksheet1.set_column(0, 2, 20, column_format1)
    worksheet1.set_column(3, 7, 28, column_format2)
    worksheet1.set_column(7, 8, 20, column_format2)

    # ADD FORMULA TO ROWS
    r = np.arange(5, _038T_rownum + 5)
    for i in r:
        formula = f'=SUM(E{i}:E{i}*G{i}:G{i})'
        worksheet1.write_formula(f'H{i}:H{i}', formula)

    # WORKSHEET 2
    worksheet2.set_column(0, 2, 20, column_format1)
    worksheet2.set_column(3, 7, 28, column_format2)
    worksheet2.set_column(7, 8, 20, column_format2)

    # ADD FORMULA TO ROWS
    r = np.arange(5, _038U_rownum + 5)
    for i in r:
        formula = f'=SUM(E{i}:E{i}*G{i}:G{i})'
        worksheet2.write_formula(f'H{i}:H{i}', formula)

    # WORKSHEET 3
    worksheet3.set_column(0, 2, 20, column_format1)
    worksheet3.set_column(3, 7, 28, column_format2)
    worksheet3.set_column(7, 8, 20, column_format2)

    # ADD FORMULA TO ROWS
    r = np.arange(5, _050T_rownum + 5)
    for i in r:
        formula = f'=SUM(E{i}:E{i}*G{i}:G{i})'
        worksheet3.write_formula(f'H{i}:H{i}', formula)

    # WORKSHEET 4
    worksheet4.set_column(0, 2, 20, column_format1)
    worksheet4.set_column(3, 7, 28, column_format2)
    worksheet4.set_column(7, 8, 20, column_format2)

    # ADD FORMULA TO ROWS
    r = np.arange(5, _050U_rownum + 5)
    for i in r:
        formula = f'=SUM(E{i}:E{i}*G{i}:G{i})'
        worksheet4.write_formula(f'H{i}:H{i}', formula)

    # WORKSHEET 5
    worksheet5.set_column(0, 2, 20, column_format1)
    worksheet5.set_column(3, 7, 28, column_format2)
    worksheet5.set_column(7, 8, 20, column_format2)

    # ADD FORMULA TO ROWS
    r = np.arange(5, _076T_rownum + 5)
    for i in r:
        formula = f'=SUM(E{i}:E{i}*G{i}:G{i})'
        worksheet5.write_formula(f'H{i}:H{i}', formula)

    # WORKSHEET 6
    worksheet6.set_column(0, 2, 20, column_format1)
    worksheet6.set_column(3, 7, 28, column_format2)
    worksheet6.set_column(7, 8, 20, column_format2)

    # ADD FORMULA TO ROWS
    r = np.arange(5, _076U_rownum + 5)
    for i in r:
        formula = f'=SUM(E{i}:E{i}*G{i}:G{i})'
        worksheet6.write_formula(f'H{i}:H{i}', formula)

    # SET FORMATTING FOR THE MERGED CELLS AC WHITCHER
    merge_formatA = workbook.add_format({
        'bold': .5,
        'left': 1,
        'right': 1,
        'font_name': 'Monotype Corsiva',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 15,
        'bg_color': 'white'
    })
    # SET FORMATTING FOR THE MERGED CELLS ESTABLISHED 1902
    merge_formatB = workbook.add_format({
        'bold': .5,
        'left': 1,
        'right': 1,
        'bottom': 1,
        'font_name': 'Times New Roman',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'bg_color': 'white'
    })

    # SET THE FORMATTING FOR THE FOOTER S7 ADDITION
    merge_formatC = workbook.add_format({
        'bold': .5,
        'top': 1,
        'left': 1,
        'right': 1,
        'bottom': 1,
        'font_name': 'Calibri',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': 'white'
    })

    # CELL COLOR FORMAT FOR ORDER NUMBER
    color_format = workbook.add_format()
    color_format.set_align('center')
    color_format.set_align('vcenter')
    color_format.set_border()

    # MERGE CELLS AND ENTER WRITING
    worksheet1.merge_range('A1:H1', 'A.C. Whitcher (PTY) Ltd', merge_formatA)
    worksheet2.merge_range('A1:H1', 'A.C. Whitcher (PTY) Ltd', merge_formatA)
    worksheet3.merge_range('A1:H1', 'A.C. Whitcher (PTY) Ltd', merge_formatA)
    worksheet4.merge_range('A1:H1', 'A.C. Whitcher (PTY) Ltd', merge_formatA)
    worksheet5.merge_range('A1:H1', 'A.C. Whitcher (PTY) Ltd', merge_formatA)
    worksheet6.merge_range('A1:H1', 'A.C. Whitcher (PTY) Ltd', merge_formatA)

    worksheet1.merge_range('C3:F3', '', merge_formatA)
    worksheet2.merge_range('C3:F3', '', merge_formatA)
    worksheet3.merge_range('C3:F3', '', merge_formatA)
    worksheet4.merge_range('C3:F3', '', merge_formatA)
    worksheet5.merge_range('C3:F3', '', merge_formatA)
    worksheet6.merge_range('C3:F3', '', merge_formatA)

    worksheet1.write_string(2, 6, 'ORDER NO:', color_format)
    worksheet2.write_string(2, 6, 'ORDER NO:', color_format)
    worksheet3.write_string(2, 6, 'ORDER NO:', color_format)
    worksheet4.write_string(2, 6, 'ORDER NO:', color_format)
    worksheet5.write_string(2, 6, 'ORDER NO:', color_format)
    worksheet6.write_string(2, 6, 'ORDER NO:', color_format)

    worksheet1.write_string(2, 0, 'CUSTOMER:')
    worksheet2.write_string(2, 0, 'CUSTOMER:')
    worksheet3.write_string(2, 0, 'CUSTOMER:')
    worksheet4.write_string(2, 0, 'CUSTOMER:')
    worksheet5.write_string(2, 0, 'CUSTOMER:')
    worksheet6.write_string(2, 0, 'CUSTOMER:')

    worksheet1.write_string(2, 1, customer_number)
    worksheet2.write_string(2, 1, customer_number)
    worksheet3.write_string(2, 1, customer_number)
    worksheet4.write_string(2, 1, customer_number)
    worksheet5.write_string(2, 1, customer_number)
    worksheet6.write_string(2, 1, customer_number)

    worksheet1.merge_range('A2:H2', 'ESTABLISHED 1902', merge_formatB)
    worksheet2.merge_range('A2:H2', 'ESTABLISHED 1902', merge_formatB)
    worksheet3.merge_range('A2:H2', 'ESTABLISHED 1902', merge_formatB)
    worksheet4.merge_range('A2:H2', 'ESTABLISHED 1902', merge_formatB)
    worksheet5.merge_range('A2:H2', 'ESTABLISHED 1902', merge_formatB)
    worksheet6.merge_range('A2:H2', 'ESTABLISHED 1902', merge_formatB)

    # USE ROW NUM TO ADD 7 11% TEXT
    t38_row = _038T_rownum + 5
    worksheet1.merge_range(f'A{t38_row}:H{t38_row}',
                           '''S7 AVAILABLE AT AN ADDITIONAL 11%''',
                           merge_formatC)
    u38_row = _038U_rownum + 5
    worksheet2.merge_range(f'A{u38_row}:H{u38_row}',
                           '''S7 AVAILABLE AT AN ADDITIONAL 11%''',
                           merge_formatC)

    t50_row = _050T_rownum + 5
    worksheet3.merge_range(f'A{t50_row}:H{t50_row}',
                           '''S7 AVAILABLE AT AN ADDITIONAL 11%''',
                           merge_formatC)
    u50_row = _050U_rownum + 5
    worksheet4.merge_range(f'A{u50_row}:H{u50_row}',
                           '''S7 AVAILABLE AT AN ADDITIONAL 11%''',
                           merge_formatC)

    t76_row = _076T_rownum + 5
    worksheet5.merge_range(f'A{t76_row}:H{t76_row}',
                           '''S7 AVAILABLE AT AN ADDITIONAL 11%''',
                           merge_formatC)
    u76_row = _076U_rownum + 5
    worksheet6.merge_range(f'A{u76_row}:H{u76_row}',
                           '''S7 AVAILABLE AT AN ADDITIONAL 11%''',
                           merge_formatC)

    # DEFAULT ROW
    worksheet1.set_default_row(16)
    worksheet2.set_default_row(16)
    worksheet3.set_default_row(16)
    worksheet4.set_default_row(16)
    worksheet5.set_default_row(16)
    worksheet6.set_default_row(16)

    # HEADING ROWS
    worksheet1.set_row(3, 28, column_format2)
    worksheet2.set_row(3, 28, column_format2)
    worksheet3.set_row(3, 28, column_format2)
    worksheet4.set_row(3, 28, column_format2)
    worksheet5.set_row(3, 28, column_format2)
    worksheet6.set_row(3, 28, column_format2)
    # AC WHITCHER ROW
    worksheet1.set_row(0, 30)
    worksheet2.set_row(0, 30)
    worksheet3.set_row(0, 30)
    worksheet4.set_row(0, 30)
    worksheet5.set_row(0, 30)
    worksheet6.set_row(0, 30)

    # WORKSHEET PROTECTION
    worksheet1.protect('acwhitcher1234')
    worksheet2.protect('acwhitcher1234')
    worksheet3.protect('acwhitcher1234')
    worksheet4.protect('acwhitcher1234')
    worksheet5.protect('acwhitcher1234')
    worksheet6.protect('acwhitcher1234')

    # CELL COLOR FORMAT FOR ORDER NUMBER
    color_format_unlocked = workbook.add_format()
    color_format_unlocked.set_bg_color('yellow')
    color_format_unlocked.set_border()
    color_format_unlocked.set_locked(False)
    color_format_unlocked.set_align('center')
    color_format_unlocked.set_align('vcenter')

    worksheet2.write_string(2, 7, 'ENTER ORDER #', color_format_unlocked)
    worksheet3.write_string(2, 7, 'ENTER ORDER #', color_format_unlocked)
    worksheet4.write_string(2, 7, 'ENTER ORDER #', color_format_unlocked)
    worksheet5.write_string(2, 7, 'ENTER ORDER #', color_format_unlocked)
    worksheet6.write_string(2, 7, 'ENTER ORDER #', color_format_unlocked)
    worksheet1.write_string(2, 7, 'ENTER ORDER #', color_format_unlocked)

    unlocked = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'locked': False
    })

    len_38T = np.arange(5, _038T_rownum + 5)
    for i in len_38T:
        worksheet1.write(f'G{i}', '', unlocked)

    len_38U = np.arange(5, _038U_rownum + 5)
    for i in len_38U:
        worksheet2.write(f'G{i}', '', unlocked)

    len_50T = np.arange(5, _050T_rownum + 5)
    for i in len_50T:
        worksheet3.write(f'G{i}', '', unlocked)

    len_50U = np.arange(5, _050U_rownum + 5)
    for i in len_50U:
        worksheet4.write(f'G{i}', '', unlocked)

    len_76T = np.arange(5, _076T_rownum + 5)
    for i in len_76T:
        worksheet5.write(f'G{i}', '', unlocked)

    len_76U = np.arange(5, _076U_rownum + 5)
    for i in len_76U:
        worksheet6.write(f'G{i}', '', unlocked)

    writer.save()

    if server_path == 'none':
        pass
    else:
        try:
            shutil.copyfile(
                f'{directory}/S5_{customer_number.strip()}.xlsx',
                f'{server_path}/S5_{customer_number.strip()}.xlsx')
        except:
            pass
