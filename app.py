from pandas import read_excel, read_sql, ExcelWriter, options
import ibm_db
import ibm_db_dbi
from tkinter import Tk, filedialog, messagebox, Frame, Button, Label, Entry, StringVar

options.display.float_format = '${:,.2f}'.format

cols0 = ['CUSTOMER', 'CONTRACT', 'AMOUNT', 'VOUCHER']
cols1 = ['CUSTOMER', 'CONTRACT', 'MONTH', 'AMOUNT', 'VOUCHER']
cols2 = ['CUSTOMER', 'CONTRACT', 'MONTH', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'AMOUNT', 'VOUCHER']
cust_cols = ['CUSTOMER', 'CONTRACT', 'MONTH', 'DIV', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'DELTA', 'BMS BILLING',
             'FIW BILLING', 'VOUCHER']
group0 = ['CUSTOMER', 'CONTRACT']
group1 = ['CUSTOMER', 'CONTRACT', 'MONTH']
group2 = ['CUSTOMER', 'CONTRACT', 'MONTH', 'BMDIV', 'MAJOR', 'INVOICE', 'PROJECTNUM']
cust_group = ['CUSTOMER', 'CONTRACT', 'MONTH', 'DIV', 'MAJOR', 'INVOICE', 'PROJECTNUM']


class Application(Frame):

    def __init__(self, parent):
        # global fiw_uid_field, fiw_pwd_field, bms_uid_field, bms_pwd_field
        Frame.__init__(self, parent)
        self.parent = parent
        self.parent.title('Librarian GUI')
        self.parent.geometry('440x290')
        self.parent.resizable(False, False)
        self.fiw_sql = None
        self.bms_sql = None
        self.customers = None
        self.fiw = None
        self.bms = None
        self.level1 = None
        self.level2 = None
        self.customers_df = None
        self.ytd_delta = None
        self.fiw_uid_field = StringVar()
        self.fiw_pwd_field = StringVar()
        self.bms_uid_field = StringVar()
        self.bms_pwd_field = StringVar()

        Button(self, relief='groove', text='1) Load customer data', width=18, command=self.load_customers).place(x=154,
                                                                                                                 y=10)
        Button(self, relief='groove', text='2a) Load FIW SQL', width=16, command=self.open_fiw_sql).place(x=20, y=50)
        Button(self, relief='groove', text='2b) Load BMS SQL', width=16, command=self.open_bms_sql).place(x=300, y=50)

        Label(self, text='FIW ID').place(x=20, y=80)
        Entry(self, relief='groove', borderwidth='2', textvariable=self.fiw_uid_field, width=14).place(x=20, y=100)
        Label(self, text='FIW Password').place(x=20, y=120)
        Entry(self, relief='groove', borderwidth='2', textvariable=self.fiw_pwd_field, width=14, show='*').place(x=20,
                                                                                                                 y=140)

        Label(self, text='BMS ID').place(x=300, y=80)
        Entry(self, relief='groove', borderwidth='2', textvariable=self.bms_uid_field, width=14).place(x=300, y=100)
        Label(self, text='BMS Password').place(x=300, y=120)
        Entry(self, relief='groove', borderwidth='2', textvariable=self.bms_pwd_field, width=14, show='*').place(x=300,
                                                                                                                 y=140)

        Button(self, relief='groove', text='3a) Retrieve FIW', width=16, command=self.retrieve_fiw).place(x=20, y=180)
        Button(self, relief='groove', text='3b) Retrieve BMS', width=16, command=self.retrieve_bms).place(x=300, y=180)
        Button(self, relief='groove', text='4) Compare data', width=18, command=self.compare_data).place(x=154, y=220)
        Button(self, relief='groove', text='5) Save data', width=18, command=self.saver).place(x=154, y=250)

    @staticmethod
    def busy():
        """Changes cursor within the application to hourglass"""
        root.config(cursor="wait")

    @staticmethod
    def not_busy():
        """Resets cursor within the application to default pointer"""
        root.config(cursor="")

    @staticmethod
    def client_exit():
        """Terminates the mainloop without crashing python runtime"""
        root.destroy()

    def open_fiw_sql(self):
        """Reads lines from text file to parse SQL"""
        file = filedialog.askopenfile(parent=root, mode='r', title='Select a text file with FIW SQL:')
        try:
            if file is not None:
                self.fiw_sql = file.read()
                messagebox.showinfo(title='Status message', message='SQL loaded successfully.')
                return self.fiw_sql
        except Exception:
            messagebox.showerror(title='Wrong format', message='File format not supported, use .txt format.')

    def open_bms_sql(self):
        """Reads lines from text file to parse SQL"""
        file = filedialog.askopenfile(parent=root, mode='r', title='Select a text file with BMS SQL:')
        try:
            if file is not None:
                self.bms_sql = file.read()
                messagebox.showinfo(title='Status message', message='SQL loaded successfully.')
                return self.bms_sql
        except Exception:
            messagebox.showerror(title='Wrong format', message='File format not supported, use .txt format.')

    def load_customers(self):
        """Reads excel table from workbook into pandas dataframe for data mapping"""
        file = filedialog.askopenfilename(parent=root, title='Select a spreadsheet with customer mapping:')
        try:
            if file is not '':
                self.customers = read_excel(r'{}'.format(file), sheet_name='Customers')#, encoding='utf-8')
                messagebox.showinfo(title='Status message', message='Mapping data loaded successfully.')
                return self.customers
        except Exception:
            messagebox.showerror(title='Wrong format', message='File format not supported, use .xlsx format.')

    def retrieve_fiw(self):
        """Retrieves data from FIW server using provided credentials and SQL"""
        if self.fiw_sql is not None:
            try:
                driver = 'IBM DB2 ODBC DRIVER'
                database = 'EUHADBM0'
                hostname = 'MEUHC.s390.emea.ibm.com'
                port = '3210'
                protocol = 'TCPIP'
                security = 'SSL'
                keydb = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.kdb'
                keysth = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.sth'
                fiw_uid = self.fiw_uid_field.get().strip()
                pwd = self.fiw_pwd_field.get().strip()
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
                messagebox.showerror(title='Authentication failed', message='Invalid user ID or password!')
            else:
                self.busy()
                self.fiw = read_sql(self.fiw_sql, conn_fiw)
                self.fiw = self.fiw.merge(self.customers, how='left', on='CONTRACT')
                self.not_busy()
                messagebox.showinfo(title='Status message', message='FIW data retrieved successfully.')
                return self.fiw

        else:
            messagebox.showwarning(title='SQL not loaded', message='A window will be opened now for selection')
            self.open_fiw_sql()

    def retrieve_bms(self):
        """Retrieves data from BMS server using provided credentials and SQL"""
        if self.bms_sql and self.customers is not None:
            try:
                driver = 'IBM DB2 ODBC DRIVER'
                database = 'MWNCDSNB'
                hostname = 'bldbmsa.boulder.ibm.com'
                port = '5508'
                protocol = 'TCPIP'
                security = 'SSL'
                keydb = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.kdb'
                keysth = r'C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.sth'
                bms_uid = self.bms_uid_field.get().strip()
                pwd = self.bms_pwd_field.get().strip()
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
                messagebox.showerror(title='Authentication failed', message='Invalid user ID or password!')
            else:
                self.busy()
                self.bms = read_sql(self.bms_sql, conn_bms)
                self.bms = self.bms.merge(self.customers, how='left', on='CONTRACT')
                self.not_busy()
                messagebox.showinfo(title='Status message', message='BMS data retrieved successfully.')
                return self.bms
            # ######################## INCLUDE VALIDATION IF OUTPUT CONTAINS INV_TYPE FIELD ##############
            # bms.loc[bms['INVOICENUMBER'].str.contains('X'), 'INV_TYPE'] = 'INT'
            # bms.loc[bms['INVOICENUMBER'].str.contains('MAN'), 'INV_TYPE'] = 'MAN'
            # bms.loc[bms['INVOICENUMBER'].str.contains('X|MAN') == False, 'INV_TYPE'] = 'EXT'

        else:
            messagebox.showwarning(title='SQL or customer datanot loaded',
                                   message='Proceed by loading customer mapping and SQLs')
            # self.open_bms_sql()

    def compare_data(self):
        """Compares data on defined view to create custom levels of detail"""
        if self.fiw and self.bms is not None:
            self.bms['VOUCHER'] = 0
            div_extract = self.bms[['CONTRACT', 'MAJOR', 'BMDIV']]
            div_extract = div_extract.rename(columns={'BMDIV': 'DIV'})
            div_extract.drop_duplicates(inplace=True)

            fiw1 = self.fiw[cols1].groupby(by=group1).sum()
            fiw1['FIW BILLING'] = fiw1['AMOUNT']
            bms1 = self.bms[cols1].groupby(by=group1).sum()
            bms1['BMS BILLING'] = bms1['AMOUNT']

            self.level1 = fiw1.subtract(bms1, axis='columns', fill_value=0)
            self.level1.reset_index(inplace=True)
            self.level1.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
            self.level1['BMS BILLING'] = self.level1['BMS BILLING'] * -1
            self.level1.fillna(0, inplace=True)

            fiw2 = self.fiw[cols2].groupby(by=group2).sum()
            fiw2['FIW BILLING'] = fiw2['AMOUNT']

            bms2 = self.bms[cols2].groupby(by=group2).sum()
            bms2['BMS BILLING'] = bms2['AMOUNT']

            # use extract of INV_TYP from BMS data
            # merge it to level 3 comparison
            # invoices found in FIW, will be matched to INV_type too
            # inv_typ = bms[['INVOICE', 'MAJOR', 'BMDIV']]

            self.level2 = fiw2.subtract(bms2, axis='columns', fill_value=0)
            self.level2.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
            self.level2['BMS BILLING'] = self.level2['BMS BILLING'] * -1
            self.level2.reset_index(inplace=True)
            self.level2.fillna(0, inplace=True)
            self.level2 = self.level2.merge(div_extract, how='left', on=['CONTRACT', 'MAJOR'])
            self.level2.loc[self.level2['DIV'].isnull(), 'DIV'] = self.level2['BMDIV']
            # level3 = level3.merge(inv_typ, how='left', on=['INVOICE', ''])

            fiw0 = self.fiw[cols0].groupby(by=group0).sum()
            fiw0['FIW BILLING'] = fiw0['AMOUNT']
            bms0 = self.bms[cols0].groupby(by=group0).sum()
            bms0['BMS BILLING'] = bms0['AMOUNT']

            self.ytd_delta = fiw0.subtract(bms0, axis='columns', fill_value=0)
            self.ytd_delta.reset_index(inplace=True)
            self.ytd_delta.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
            self.ytd_delta['BMS BILLING'] = self.ytd_delta['BMS BILLING'] * -1
            self.ytd_delta.fillna(0, inplace=True)

            self.customers_df = fiw2.subtract(bms2, axis='columns', fill_value=0)
            self.customers_df.reset_index(inplace=True)
            self.customers_df.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
            self.customers_df['BMS BILLING'] = self.customers_df['BMS BILLING'] * -1
            self.customers_df.fillna(0, inplace=True)
            self.customers_df = self.customers_df.merge(div_extract, how='left', on=['CONTRACT', 'MAJOR'])
        else:
            messagebox.showwarning(title='Missing data', message='FIW or BMS data was not retrieved.')
            if self.fiw_sql is not None:
                messagebox.showinfo(title='Query', message='FIW data will be retrieved')
                self.retrieve_fiw()
            elif self.bms_sql is not None:
                messagebox.showinfo(title='Query', message='BMS data will be retrieved')
                self.retrieve_bms()

    def saver(self):
        """Saves the dataframes from memory to local directory in xlsx format. Customer specific views are stored within
        same folder as the main output which user names during save"""
        save_location = filedialog.asksaveasfilename(filetypes=(('Excel files', '*.xlsx'),
                                                                ('All files', '*.*')))
        if save_location[-5:] != '.xlsx':
            save_location = save_location + '.xlsx'
        else:
            pass

        # messagebox.showinfo(title='Status message', message='Saving data...')
        writer = ExcelWriter(save_location, engine='xlsxwriter')
        self.fiw.to_excel(writer, sheet_name='FIW', index=False)
        self.bms.to_excel(writer, sheet_name='BMS', index=False)
        self.ytd_delta.to_excel(writer, sheet_name='YTD Overview', index=False)
        self.level1.to_excel(writer, sheet_name='Level 1', index=False)
        self.level2.to_excel(writer, sheet_name='Level 2', index=False)

        self.busy()
        for customer in set(self.customers_df['CUSTOMER']):
            individual_view = self.customers_df[cust_cols][self.customers_df['CUSTOMER'] == f'{str(customer)}']
            individual_view = individual_view.groupby(by=cust_group).sum()
            individual_view.reset_index(inplace=True)
            # print(customer)
            # print(save_location[0:int(save_location.rfind('/') + 1)] + f'{customer}' + '.xlsx')
            individual_view.to_excel(save_location[0:int(save_location.rfind('/') + 1)] + f'{customer}.xlsx',
                                     sheet_name=f'{customer}', index=False)

        writer.save()
        self.not_busy()
        messagebox.showinfo(title='Status message', message='Data has been saved.')

        # info_sheet = save_location.add_worksheets('Information')
        # info_sheet['B2'] = 'FIW ID'
        # info_sheet['C2'] = fiw_uid
        # info_sheet['D2'] = 'Run date: ' + fiw.loc[0, 'RUN_DATE']
        # info_sheet['B3'] = 'BMS ID'
        # info_sheet['C3'] = bms_uid
        # info_sheet['D3'] = 'Run date: ' + bms.loc[0, 'RUN_DATE']


if __name__ == "__main__":
    root = Tk()
    Application(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
