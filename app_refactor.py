import pandas as pd
import ibm_db
import ibm_db_dbi
from tkinter import *
from tkinter import filedialog

cols0 = ['CUSTOMER', 'CONTRACT', 'AMOUNT']
cols1 = ['MONTH', 'CUSTOMER', 'CONTRACT', 'AMOUNT']
cols2 = ['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'AMOUNT']
cols3 = ['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'AMOUNT']
group0 = ['CUSTOMER', 'CONTRACT']
group1 = ['MONTH', 'CUSTOMER', 'CONTRACT']
group2 = ['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE']
group3 = ['MONTH', 'CUSTOMER', 'CONTRACT', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM']


class Window(Frame):

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

    # maybe should not be static
    @staticmethod
    def client_exit():
        root.destroy()

    @staticmethod
    def open_fiw_sql():
        global fiw_sql
        file = filedialog.askopenfile(parent=root, mode='r', title='Choose a text file with FIW SQL')
        if file is not None:
            fiw_sql = file.read()
            return fiw_sql

    @staticmethod
    def open_bms_sql():
        global bms_sql
        file = filedialog.askopenfile(parent=root, mode='r', title='Choose a text file with BMS SQL')
        if file is not None:
            bms_sql = file.read()
            return bms_sql

    @staticmethod
    def load_customers():
        global customers
        file = filedialog.askopenfilename(parent=root, title='Load an excel file with customer mapping')
        if file is not None:
            customers = pd.read_excel(file, sheet_name='Customers', converters={'MAJOR': str})
            return customers

    @staticmethod
    def retrieve_fiw():
        global fiw
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
                uid = input('Enter your FIW-LR user ID: ').strip()
                pwd = input('Enter your password: ').strip()

                dsn_fiw = (
                    f'DRIVER={driver};'
                    f'DATABASE={database};'
                    f'HOSTNAME={hostname};'
                    f'PORT={port};'
                    f'PROTOCOL={protocol};'
                    f'UID={uid};'
                    f'PWD={pwd};'
                    f'SECURITY={security};'
                    f'SSL_keystoredb={keydb};'
                    f'SSL_keystash={keysth};')

                conn_engine_fiw = ibm_db.connect(dsn_fiw, '', '')
                conn_fiw = ibm_db_dbi.Connection(conn_engine_fiw)
                # add fiw_sql is None check, if None then call function

            except Exception:
                continue
            else:
                print('Retrieving FIW data...')
                fiw = pd.read_sql(fiw_sql, conn_fiw)
                print('FIW data retrieved successfully')
                break
        fiw = fiw.merge(customers, how='left', on='CONTRACT')
        return fiw

    @staticmethod
    def retrieve_bms():
        global bms
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
                uid = input('Enter your BMSIW user ID: ').strip()
                pwd = input('Enter your password: ').strip()

                dsn_bms = (
                    f'DRIVER={driver};'
                    f'DATABASE={database};'
                    f'HOSTNAME={hostname};'
                    f'PORT={port};'
                    f'PROTOCOL={protocol};'
                    f'UID={uid};'
                    f'PWD={pwd};'
                    f'SECURITY={security};'
                    f'SSL_keystoredb={keydb};'
                    f'SSL_keystash={keysth};')

                conn_engine_bms = ibm_db.connect(dsn_bms, '', '')
                conn_bms = ibm_db_dbi.Connection(conn_engine_bms)
                # add fiw_sql is None check, if None then call function
            except Exception:
                continue
            else:
                print('Retrieving BMS data...')
                bms = pd.read_sql(bms_sql, conn_bms)
                print('BMS data retrieved successfully')
                break
        bms = bms.merge(customers, how='left', on='CONTRACT')
        return bms

    @staticmethod
    def create_views():
        global fiw0, bms0, fiw1, bms1, fiw2, bms2, fiw3, bms3
        fiw0 = fiw[cols0].groupby(by=[group0]).sum()
        fiw0['FIW AMOUNT'] = fiw0['AMOUNT']
        bms0 = bms[cols0].groupby(by=[group0]).sum()

        bms0['BMS AMOUNT'] = bms0['AMOUNT']
        fiw1 = fiw[cols1].groupby(by=group1).sum()
        fiw1['FIW AMOUNT'] = fiw1['AMOUNT']
        bms1 = bms[cols1].groupby(by=group1).sum()
        bms1['BMS AMOUNT'] = bms1['AMOUNT']

        fiw2 = fiw[cols2].groupby(by=group2).sum()
        fiw2['FIW AMOUNT'] = fiw2['AMOUNT']
        bms2 = bms[cols2].groupby(by=group2).sum()
        bms2['BMS AMOUNT'] = bms2['AMOUNT']

        fiw3 = fiw[cols3].groupby(by=group3).sum()
        fiw3['FIW AMOUNT'] = fiw3['AMOUNT']
        bms3 = bms[cols3].groupby(by=group3).sum()
        bms3['BMS AMOUNT'] = bms3['AMOUNT']
        return fiw0, bms0, fiw1, bms1, fiw2, bms2, fiw3, bms3

    @staticmethod
    def compare_data(fiw_df=None, bms_df=None):
        global ytd, level1, level2, level3
        comparison = fiw_df.subtract(bms_df, axis='columns', fill_value=0)
        comparison.reset_index(inplace=True)
        comparison.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        comparison['BMS AMOUNT'] = comparison['BMS AMOUNT'] * -1
        comparison.fillna(0, inplace=True)

        return ytd, level1, level2, level3

    @staticmethod
    def saver():
        savefile = filedialog.asksaveasfilename(filetypes=(('Excel files', '*.xlsx'),
                                                           ('All files', '*.*')))
        savefile = savefile + '.xlsx'
        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
        fiw.to_excel(writer, sheet_name='FIW', index=False)
        bms.to_excel(writer, sheet_name='BMS', index=False)
        ytd_delta.to_excel(writer, sheet_name='YTD Delta', index=False)
        level1.to_excel(writer, sheet_name='Level 1', index=False)
        level2.to_excel(writer, sheet_name='Level 2', index=False)
        level3.to_excel(writer, sheet_name='Level 3', index=False)
        print('Saving data, please wait...')
        writer.save()
        print('Data has been saved')


# initialize tkinter class interface
root = Tk()
# defines the size of the frame
root.geometry("240x160")
# fills the frame with buttons, creating an interactive window
app = Window(root)
# buttons in windows can be activated multiple times
# the process terminates by closing the application window
root.mainloop()
