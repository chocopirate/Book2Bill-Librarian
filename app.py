import pandas as pd
import numpy as np
import ibm_db
import ibm_db_dbi
import xlsxwriter
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

pd.options.display.float_format = '${:,.2f}'.format


class Application(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()

    def init_window(self):
        self.master.title('Librarian GUI')
        self.pack(fill=BOTH, expand=1)
        button_fiw_sql = Button(self, text='Load FIW SQL', width=16, command=self.open_fiw_sql)
        button_fiw_sql.pack(fill=X)
        button_bms_sql = Button(self, text='Load BMS SQL', width=16, command=self.open_bms_sql)
        button_bms_sql.pack(fill=X)
        button_customers = Button(self, text='Load customer data', width=16, command=self.load_customers)
        button_customers.pack(fill=X)
        button_retrieve_fiw = Button(self, text='Retrieve FIW data', width=16, command=self.retrieve_fiw)
        button_retrieve_fiw.pack(fill=X)
        button_run_bms = Button(self, text='Retrieve BMS data', width=16, command=self.retrieve_bms)
        button_run_bms.pack(fill=X)
        button_compare = Button(self, text='Compare data', width=16, command=self.compare_data)
        button_compare.pack(fill=X)
        button_save = Button(self, text='Save data', width=16, command=self.saver)
        button_save.pack(fill=X)
        window_menu = Menu(self.master)
        self.master.config(menu=window_menu)

    @staticmethod
    def client_exit():
        root.destroy()

    @staticmethod
    def open_fiw_sql():
        global fiw_sql
        file = filedialog.askopenfile(parent=root, mode='r', title='Choose a text file with SQL for FIW')
        if file is not None:
            fiw_sql = file.read()
            messagebox.showinfo(title='Status message', message='SQL loaded successfully.')
            return fiw_sql

    @staticmethod
    def open_bms_sql():
        global bms_sql
        file = filedialog.askopenfile(parent=root, mode='r', title='Choose a text file with SQL for BMS')
        if file is not None:
            bms_sql = file.read()
            messagebox.showinfo(title='Status message', message='SQL loaded successfully.')
            return bms_sql

    @staticmethod
    def load_customers():
        global customers
        file = filedialog.askopenfilename(parent=root, title='Select a spreadsheet containing customer mapping')
        if file is not '':
            customers = pd.read_excel(r'{}'.format(file), sheet_name='Customers', encoding='utf-8')
            messagebox.showinfo(title='Status message', message='Mapping data loaded successfully.')
            return customers

    @staticmethod
    def retrieve_fiw():
        global fiw, fiw_uid
        if 'fiw_sql' in globals() or 'fiw_sql' in locals():
            while True:
                try:
                    driver = 'IBM DB2 ODBC DRIVER'
                    database = 'EUHADBM0'
                    hostname = 'MEUHC.s390.emea.ibm.com'
                    port = '3210'
                    protocol = 'TCPIP'
                    security = 'SSL'
                    keydb = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.kdb'
                    keysth = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.sth'
                    fiw_uid = input('Enter your FIW-LR user ID: ').strip()
                    pwd = input('Enter your password: ').strip()
                    dsn_fiw = (
                        f'DRIVER={driver};'
                        f'DATABASE={database};'
                        f'HOSTNAME={hostname};'
                        f'PORT={port};'
                        f'PROTOCOL={protocol};'
                        f'UID={fiw_uid};'
                        f'PWD={pwd};'
                        f'SECURITY={security};'
                        f'SSL_keystoredb={keydb};'
                        f'SSL_keystash={keysth};')
                    conn_engine_fiw = ibm_db.connect(dsn_fiw, '', '')
                    conn_fiw = ibm_db_dbi.Connection(conn_engine_fiw)
                except Exception:
                    continue
                else:
                    print('Retrieving FIW data...')
                    fiw = pd.read_sql(fiw_sql, conn_fiw)
                    print('FIW data retrieved successfully')
                    break
            fiw = fiw.merge(customers, how='left', on='CONTRACT')
            return fiw
        else:
            messagebox.showwarning(title='SQL not loaded', message='A window will be opened now for selection')
            Application.open_fiw_sql()

    @staticmethod
    def retrieve_bms():
        global bms, bms_uid
        if 'bms_sql' in globals() or 'bms_sql' in locals():
            while True:
                try:
                    driver = 'IBM DB2 ODBC DRIVER'
                    database = 'MWNCDSNB'
                    hostname = 'bldbmsa.boulder.ibm.com'
                    port = '5508'
                    protocol = 'TCPIP'
                    security = 'SSL'
                    keydb = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.kdb'
                    keysth = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.sth'
                    bms_uid = input('Enter your BMSIW user ID: ').strip()
                    pwd = input('Enter your password: ').strip()
                    dsn_bms = (
                        f'DRIVER={driver};'
                        f'DATABASE={database};'
                        f'HOSTNAME={hostname};'
                        f'PORT={port};'
                        f'PROTOCOL={protocol};'
                        f'UID={bms_uid};'
                        f'PWD={pwd};'
                        f'SECURITY={security};'
                        f'SSL_keystoredb={keydb};'
                        f'SSL_keystash={keysth};')
                    conn_engine_bms = ibm_db.connect(dsn_bms, '', '')
                    conn_bms = ibm_db_dbi.Connection(conn_engine_bms)
                except Exception:
                    continue
                else:
                    print('Retrieving BMS data...')
                    bms = pd.read_sql(bms_sql, conn_bms)
                    print('BMS data retrieved successfully')
                    break
            bms = bms.merge(customers, how='left', on='CONTRACT')
            bms.loc[bms['INVOICENUMBER'].str.contains('X'), 'INV_TYPE'] = 'INT'
            bms.loc[bms['INVOICENUMBER'].str.contains('MAN'), 'INV_TYPE'] = 'MAN'
            bms.loc[bms['INVOICENUMBER'].str.contains('X|MAN') == False, 'INV_TYPE'] = 'EXT'
            return bms
        else:
            messagebox.showwarning(title='SQL not loaded',
                                   message='A window will be opened now for selection.')
            Application.open_bms_sql()

    @staticmethod
    def compare_data():
        global level1, level2, customers_df, ytd_delta

        div_extract = bms[['CONTRACT', 'MAJOR', 'BMDIV']]
        div_extract = div_extract.rename(columns={'BMDIV': 'DIV'})
        div_extract.drop_duplicates(inplace=True)

        fiw1 = fiw[['MONTH', 'CUSTOMER', 'CONTRACT', 'AMOUNT']].groupby(by=['MONTH', 'CUSTOMER', 'CONTRACT']).sum()
        fiw1['FIW AMOUNT'] = fiw1['AMOUNT']
        bms1 = bms[['MONTH', 'CUSTOMER', 'CONTRACT', 'AMOUNT']].groupby(by=['MONTH', 'CUSTOMER', 'CONTRACT']).sum()
        bms1['BMS AMOUNT'] = bms1['AMOUNT']

        level1 = fiw1.subtract(bms1, axis='columns', fill_value=0)
        level1['Near_Zero'] = level1['AMOUNT'].between(-1, 1)
        level1 = level1[level1['Near_Zero'] == False]
        level1.drop(columns='Near_Zero', inplace=True)
        level1.reset_index(inplace=True)
        level1.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        level1['BMS AMOUNT'] = level1['BMS AMOUNT'] * -1
        level1.fillna(0, inplace=True)

        fiw2 = fiw[
            ['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'AMOUNT']].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM']).sum()
        fiw2['FIW AMOUNT'] = fiw2['AMOUNT']

        bms2 = bms[
            ['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'AMOUNT', 'INV_TYPE']].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'INV_TYPE']).sum()
        bms2['BMS AMOUNT'] = bms2['AMOUNT']

        # use extract of INV_TYP from BMS data
        # merge it to level 3 comparison
        # invoices found in FIW, will be matched to INV_type too
        # inv_typ = bms[['INVOICE', 'MAJOR', 'BMDIV']]

        level2 = fiw2.subtract(bms2, axis='columns', fill_value=0)
        level2.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        level2['BMS AMOUNT'] = level2['BMS AMOUNT'] * -1
        level2.reset_index(inplace=True)
        level2.fillna(0, inplace=True)
        level2 = level2.merge(div_extract, how='left', on=['CONTRACT', 'MAJOR'])
        level2.loc[level2['DIV'].isnull(), 'DIV'] = level2['BMDIV']
        # level3 = level3.merge(inv_typ, how='left', on=['INVsOICE', ''])

        fiw0 = fiw[['CUSTOMER', 'CONTRACT', 'AMOUNT']].groupby(by=['CUSTOMER', 'CONTRACT']).sum()
        fiw0['FIW AMOUNT'] = fiw0['AMOUNT']
        bms0 = bms[['CUSTOMER', 'CONTRACT', 'AMOUNT']].groupby(by=['CUSTOMER', 'CONTRACT']).sum()
        bms0['BMS AMOUNT'] = bms0['AMOUNT']

        ytd_delta = fiw0.subtract(bms0, axis='columns', fill_value=0)
        ytd_delta['Near_Zero'] = ytd_delta['AMOUNT'].between(-1, 1)
        ytd_delta = ytd_delta[ytd_delta['Near_Zero'] == False]
        ytd_delta.drop(columns='Near_Zero', inplace=True)
        ytd_delta.reset_index(inplace=True)
        ytd_delta.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        ytd_delta['BMS AMOUNT'] = ytd_delta['BMS AMOUNT'] * -1
        ytd_delta.fillna(0, inplace=True)

        customers_df = fiw2.subtract(bms2, axis='columns', fill_value=0)
        customers_df.reset_index(inplace=True)
        customers_df.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        customers_df['BMS AMOUNT'] = customers_df['BMS AMOUNT'] * -1
        customers_df.fillna(0, inplace=True)
        customers_df = customers_df.merge(div_extract, how='left', on=['CONTRACT', 'MAJOR'])

    @staticmethod
    def saver():
        global workbook
        workbook = filedialog.asksaveasfilename(filetypes=(('Excel files', '*.xlsx'),
                                                           ('All files', '*.*')))
        if workbook[-5:] != '.xlsx':
            workbook = workbook + '.xlsx'
        else:
            pass

        print('Saving data...')
        writer = pd.ExcelWriter(workbook, engine='xlsxwriter')
        # info_sheet = workbook.add_worksheets('Information')
        # info_sheet['B2'] = 'FIW ID'
        # info_sheet['C2'] = fiw_uid
        # info_sheet['D2'] = 'Run date: ' + fiw.loc[0, 'RUN_DATE']
        # info_sheet['B3'] = 'BMS ID'
        # info_sheet['C3'] = bms_uid
        # info_sheet['D3'] = 'Run date: ' + bms.loc[0, 'RUN_DATE']
        fiw.to_excel(writer, sheet_name='FIW', index=False)
        bms.to_excel(writer, sheet_name='BMS', index=False)
        ytd_delta.to_excel(writer, sheet_name='YTD Overview', index=False)
        level1.to_excel(writer, sheet_name='Level 1', index=False)
        level2.to_excel(writer, sheet_name='Level 2', index=False)

        for customer in list(customers_df['CUSTOMER']):
            individual_view = customers_df[['CUSTOMER', 'MONTH', 'CONTRACT', 'DIV', 'MAJOR', 'INVOICE', 'DELTA', 'BMS AMOUNT', 'FIW AMOUNT']][
                customers_df['CUSTOMER'] == f'{str(customer)}']
            individual_view = individual_view.groupby(by=['CUSTOMER', 'MONTH', 'CONTRACT', 'DIV', 'MAJOR', 'INVOICE']).sum()
            individual_view.reset_index(inplace=True)
            individual_view.to_excel(writer, sheet_name=f'{customer}', index=False)
        writer.save()

        print(workbook)
        print('Data has been saved')


# initialize tkinter class interface
root = Tk()
# defines the size of the frame
root.geometry("240x160")
# fills the window frame with buttons
app = Application(root)
# root (button functions) are looped; the window closes only explicitly
app.mainloop()
