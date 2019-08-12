import pandas as pd
import ibm_db
import ibm_db_dbi
from tkinter import *
from tkinter import filedialog


class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()

    def init_window(self):

        self.master.title("Librarian GUI")
        self.pack(fill=BOTH, expand=1)

        # button_fiw_sql.grid(row=2, column=1)  # , columnspan=2)
        # button_bms_sql.place(x=6, y=6)
        button_fiw_sql = Button(self, text="Load FIW SQL", width=16, command=self.open_fiw_sql)
        button_fiw_sql.pack()
        button_bms_sql = Button(self, text="Load BMS SQL", width=16, command=self.open_bms_sql)
        button_bms_sql.pack()
        button_customers = Button(self, text="Load customer data", width=16, command=self.load_customers)
        button_customers.pack()
        button_retrieve_fiw = Button(self, text="Retrieve FIW data", width=16, command=self.retrieve_fiw)
        button_retrieve_fiw.pack()
        button_run_bms = Button(self, text="Retrieve BMS data", width=16, command=self.retrieve_bms)
        button_run_bms.pack()
        button_compare = Button(self, text="Compare data", width=16, command=self.compare_data)
        button_compare.pack()
        button_save = Button(self, text="Save data", width=16, command=self.saver)
        button_save.pack()

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
        global customers, bus_line, major_map
        file = filedialog.askopenfilename(parent=root, title="Load an excel file with customer mapping")
        if file is not None:
            major_map = pd.read_excel(file, sheet_name='Major map', converters={'MAJOR': str})
            customers = pd.read_excel(file, sheet_name='Customers', converters={'MAJOR': str})
            bus_line = pd.read_excel(file, sheet_name='BusLine map', converters={'MAJOR': str})
            return customers, bus_line, major_map

    @staticmethod
    def retrieve_fiw():
        global fiw
        while True:
            try:
                driver = "IBM DB2 ODBC DRIVER"
                database = "EUHADBM0"
                hostname = "MEUHC.s390.emea.ibm.com"
                port = "3210"
                protocol = "TCPIP"
                security = "SSL"
                keydb = r"C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.kdb"
                keysth = r"C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.sth"
                uid = input("Enter your FIW-LR user ID: ").strip()
                pwd = input("Enter your password: ").strip()

                dsn_fiw = (
                    f"DRIVER={driver};"
                    f"DATABASE={database};"
                    f"HOSTNAME={hostname};"
                    f"PORT={port};"
                    f"PROTOCOL={protocol};"
                    f"UID={uid};"
                    f"PWD={pwd};"
                    f"SECURITY={security};"
                    f"SSL_keystoredb={keydb};"
                    f"SSL_keystash={keysth};")

                conn_engine_fiw = ibm_db.connect(dsn_fiw, "", "")
                conn_fiw = ibm_db_dbi.Connection(conn_engine_fiw)
                # add fiw_sql is None check, if None then call function
                fiw = pd.read_sql(fiw_sql, conn_fiw)
            except Exception:
                continue
            else:
                break
        fiw = fiw.merge(bus_line, how='left', on='MAJOR')
        fiw = fiw.merge(customers, how='left', on='CONTRACT')
        return fiw

    @staticmethod
    def retrieve_bms():
        global bms
        while True:
            try:
                driver = "IBM DB2 ODBC DRIVER"
                database = "MWNCDSNB"
                hostname = "bldbmsa.boulder.ibm.com"
                port = "5508"
                protocol = "TCPIP"
                security = "SSL"
                keydb = r"C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.kdb"
                keysth = r"C:\ProgramData\IBM\DB2\DB2COPY1\DB2\ibmca.sth"
                uid = input("Enter your BMSIW user ID: ").strip()
                pwd = input("Enter your password: ").strip()

                dsn_bms = (
                    f"DRIVER={driver};"
                    f"DATABASE={database};"
                    f"HOSTNAME={hostname};"
                    f"PORT={port};"
                    f"PROTOCOL={protocol};"
                    f"UID={uid};"
                    f"PWD={pwd};"
                    f"SECURITY={security};"
                    f"SSL_keystoredb={keydb};"
                    f"SSL_keystash={keysth};")

                conn_engine_bms = ibm_db.connect(dsn_bms, "", "")
                conn_bms = ibm_db_dbi.Connection(conn_engine_bms)
                # add fiw_sql is None check, if None then call function
                bms = pd.read_sql(bms_sql, conn_bms)
            except Exception:
                continue
            else:
                break
        bms = bms.merge(major_map, how='left', on='BUSINESSTYPE')
        bms = bms.merge(bus_line, how='left', on='MAJOR')
        bms = bms.merge(customers, how='left', on='CONTRACT')
        print(bms.shape)
        return bms

    @staticmethod
    def compare_data():
        global level1, level2, level3, bms_ytd
        fiw1 = fiw[['MONTH', 'CUSTOMER', 'CONTRACT', 'AMOUNT']][
            fiw['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT']).sum()
        fiw1['FIW AMOUNT'] = fiw1['AMOUNT']
        bms1 = bms[['MONTH', 'CUSTOMER', 'CONTRACT', 'AMOUNT']][
            bms['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT']).sum()
        bms1['BMS AMOUNT'] = bms1['AMOUNT']

        level1 = fiw1.subtract(bms1, axis='columns', fill_value=0)
        level1['Near_Zero'] = level1['AMOUNT'].between(-1, 1)
        level1 = level1[level1['Near_Zero'] == False]
        level1.drop(columns='Near_Zero', inplace=True)
        level1.reset_index(inplace=True)
        level1.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        level1['BMS AMOUNT'] = level1['BMS AMOUNT'] * -1
        level1.fillna(0, inplace=True)

        fiw2 = fiw[['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE', 'AMOUNT']][
            fiw['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE']).sum()
        fiw2['FIW AMOUNT'] = fiw2['AMOUNT']

        bms2 = bms[['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE', 'AMOUNT']][
            bms['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE']).sum()
        bms2['BMS AMOUNT'] = bms2['AMOUNT']

        level2 = fiw2.subtract(bms2, axis='columns', fill_value=0)
        level2['Near_Zero'] = level2['AMOUNT'].between(-1, 1)
        level2 = level2[level2['Near_Zero'] == False]
        level2.drop(columns='Near_Zero', inplace=True)
        level2.reset_index(inplace=True)
        level2.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        level2['BMS AMOUNT'] = level2['BMS AMOUNT'] * -1
        level2.fillna(0, inplace=True)

        fiw3 = fiw[['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'AMOUNT']][
            fiw['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE', 'PROJECTNUM']).sum()
        fiw3['FIW AMOUNT'] = fiw3['AMOUNT']

        bms3 = bms[['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE', 'PROJECTNUM', 'AMOUNT']][
            bms['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['MONTH', 'CUSTOMER', 'CONTRACT', 'BUSINESSLINE', 'MAJOR', 'INVOICE', 'PROJECTNUM']).sum()
        bms3['BMS AMOUNT'] = bms3['AMOUNT']

        level3 = fiw3.subtract(bms3, axis='columns', fill_value=0)
        level3['Near_Zero'] = level3['AMOUNT'].between(-1, 1)
        level3 = level3[level3['Near_Zero'] == False]
        level3.drop(columns='Near_Zero', inplace=True)
        level3.rename(columns={'AMOUNT': 'DELTA'}, inplace=True)
        level3['BMS AMOUNT'] = level3['BMS AMOUNT'] * -1
        level3.reset_index(inplace=True)
        level3.fillna(0, inplace=True)

        bms_ytd = bms[
            ['CONTRACT', 'MONTH', 'BILLINGDATE', 'INVOICE', 'INVOICETEXT', 'MAJOR', 'PROJECTNUM', 'INVOICEDATE',
             'BILLFROMDATE', 'BILLTHRUDATE', 'AMOUNT']][
            bms['MAJOR'].isin(['326', '323', '352', '366', '346', '374', '365'])].groupby(
            by=['CONTRACT', 'MONTH', 'BILLINGDATE', 'INVOICE', 'INVOICETEXT', 'MAJOR', 'PROJECTNUM', 'INVOICEDATE',
                'BILLFROMDATE', 'BILLTHRUDATE']).sum()
        bms_ytd.reset_index(inplace=True)

    @staticmethod
    def saver():
        savefile = filedialog.asksaveasfilename(filetypes=(('Excel files', '*.xlsx'),
                                                           ('All files', '*.*')))
        savefile = savefile + '.xlsx'
        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
        fiw.to_excel(writer, sheet_name='FIW', index=False)
        bms_ytd.to_excel(writer, sheet_name='BMS', index=False)
        level1.to_excel(writer, sheet_name='Level1', index=False)
        level2.to_excel(writer, sheet_name='Level2', index=False)
        level3.to_excel(writer, sheet_name='Level3', index=False)
        writer.save()


root = Tk()
root.geometry("300x200")
app = Window(root)
root.mainloop()
