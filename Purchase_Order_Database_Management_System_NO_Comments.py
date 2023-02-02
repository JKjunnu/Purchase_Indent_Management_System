from cmath import exp
from operator import index
import pandas as pd  # to read and write excel sheets from database
from tkinter import*  # for the frontend gui
from tkinter import ttk, filedialog  # for treeview , combobox and filedialog
from tkinter.font import BOLD  # for calendar font
from tkcalendar import Calendar  # for the calendar entry widget
import datetime  # for date formating
from tkinter import messagebox  # for info and error messages
import psycopg2  # for connecting postgreSQL database
from psycopg2 import Error  # for error handeling
import os  # foe reading current directory
# for creating engine to read and write excel files
from sqlalchemy import create_engine

# The root window of tkinter


class homePage():
    def __init__(self, root):
        self.root = root
        self.root.title("HOME")
        self.root.geometry("1540x800+0+0")

        lbltitle = Label(self.root, bd=20, relief=RIDGE, text="PURCHASE INDENT MANAGEMENT SYSTEM",
                         fg="red", bg="white", font=("times new roman", 40, "bold"))
        lbltitle.pack(side=TOP, fill=X)

        dataFrame = Frame(self.root)
        dataFrame.pack(fill=BOTH, expand=True, padx=10, pady=(50, 10))

        recordFrame1 = Frame(dataFrame)
        recordFrame1.pack(side=LEFT, fill=BOTH, expand=True,
                          padx=(10, 5), pady=(10, 0))

        recordFrame2 = Frame(dataFrame)
        recordFrame2.pack(side=LEFT, fill=BOTH,
                          expand=True, padx=(5, 10), pady=(10, 0))

        recordFrame3 = Frame(self.root)
        recordFrame3.pack(side=TOP, fill=BOTH,
                          expand=True, padx=(5, 10), pady=(0, 10))

        self.btnNameNewRecord = Button(
            recordFrame1, text="NEW RECORD", font=("arial", 20, "bold"), command=self.funcNewRecord)
        self.btnNameNewRecord.pack(side=TOP,
                                   expand=True)

        self.btnNameUpdateRecord = Button(
            recordFrame1, text="UPDATE RECORD", font=("arial", 20, "bold"), command=self.funcUpdateRecord)
        self.btnNameUpdateRecord.pack(side=TOP,
                                      expand=True)

        self.btnNameSearchRecord = Button(
            recordFrame2, text="SEARCH RECORD", font=("arial", 20, "bold"), command=self.funcSearchRecord)
        self.btnNameSearchRecord.pack(side=TOP,
                                      expand=True)

        self.btnNameDeleteRecord = Button(
            recordFrame2, text="DELETE RECORD", font=("arial", 20, "bold"), command=self.funcDeleteRecord)
        self.btnNameDeleteRecord.pack(side=TOP,
                                      expand=True)

        self.btnNameUploadCSVRecord = Button(
            recordFrame3, text="INSERT MULTIPLE RECORDS", font=("arial", 20, "bold"), command=self.funcUploadCSVFile)
        self.btnNameUploadCSVRecord.pack(side=TOP,
                                         expand=True)

    def funcNewRecord(self):
        self.btnNameNewRecord.config(command="")

        self.newRecordTopLevel = Toplevel(self.root)
        self.newRecordTopLevel.geometry("1540x800+0+0")
        self.newRecordTopLevel.title("NEW RECORD")

        self.varIndentNo = StringVar()
        self.varItemDesc = StringVar()
        self.varDivision = StringVar()
        self.varIndentorName = StringVar()
        self.varModeOfProc = StringVar()
        self.varAmountEstimate = DoubleVar()
        self.varStatus = StringVar()
        self.varActualAmount = DoubleVar()
        self.varAdditionalInfo = StringVar()
        self.uploadSpecsBinaryData = ""
        self.varDateRaised = ""
        self.varDeliveryDateShow = ""

        lbltitle = Label(self.newRecordTopLevel, bd=20, relief=RIDGE, text="NEW RECORD",
                         fg="red", bg="white", font=("times new roman", 40, "bold"))
        lbltitle.pack(side=TOP, fill=X)

        dataFrame = LabelFrame(self.newRecordTopLevel, bd=20, relief=RIDGE, font=(
            "arial", 30, "bold"), text="ENTER NEW RECORD DETAILS")
        dataFrame.pack(fill=BOTH, expand=True, padx=10, pady=(20, 10))

        self.newRecordFrame1 = Frame(dataFrame, bd=20, relief=RIDGE)
        self.newRecordFrame1.pack(side=LEFT, fill=BOTH, expand=True,
                                  padx=(10, 5), pady=(10, 10))

        newRecordFrame2 = Frame(dataFrame, bd=20, relief=RIDGE)
        newRecordFrame2.pack(side=LEFT, fill=BOTH,
                             expand=True, padx=(5, 10), pady=(10, 10))

        newRecordFrame3 = Frame(self.newRecordTopLevel, bd=20, relief=RIDGE)
        newRecordFrame3.pack(fill=BOTH,
                             expand=True, padx=(10, 10), pady=(0, 10))

        lblNameIndent = Label(
            self.newRecordFrame1, text="INDENT NO", font=("arial", 20, "bold"))
        lblNameIndent.grid(row=0, column=0, sticky=W)

        newRecordentryNameIndent = Entry(
            self.newRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varIndentNo)
        newRecordentryNameIndent.grid(row=0, column=1, sticky=E)

        lblNameDateRaised = Label(
            self.newRecordFrame1, text="DATE RAISED", font=("arial", 20, "bold"))
        lblNameDateRaised.grid(row=1, column=0, sticky=W)

        btnNameDateRaised = Button(self.newRecordFrame1, text="SELECT", font=(
            "arial", 10, "bold"), command=self.funcNewRecordDateRaised)
        btnNameDateRaised.grid(row=1, column=2, padx=20)

        lblNameItemDesc = Label(
            self.newRecordFrame1, text="ITEM DESCRP", font=("arial", 20, "bold"))
        lblNameItemDesc.grid(row=2, column=0, sticky=W)

        newRecordentryNameItemDesc = Entry(
            self.newRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varItemDesc)
        newRecordentryNameItemDesc.grid(row=2, column=1, sticky=E)

        lblNameDivName = Label(
            self.newRecordFrame1, text="DIVISION", font=("arial", 20, "bold"))
        lblNameDivName.grid(row=3, column=0, sticky=W)

        newRecordentryNameDivName = Entry(
            self.newRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varDivision)
        newRecordentryNameDivName.grid(row=3, column=1, sticky=E)

        lblNameIndentorName = Label(
            self.newRecordFrame1, text="INDENTOR NAME", font=("arial", 20, "bold"))
        lblNameIndentorName.grid(row=4, column=0, sticky=W)

        newRecordentryNameIndentorName = Entry(
            self.newRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varIndentorName)
        newRecordentryNameIndentorName.grid(row=4, column=1, sticky=E)

        lblNameDeliveryDate = Label(
            self.newRecordFrame1, text="DELIVERY DATE", font=("arial", 20, "bold"))
        lblNameDeliveryDate.grid(row=5, column=0, sticky=W)

        btnNameDeliveryDate = Button(self.newRecordFrame1, text="SELECT", font=(
            "arial", 10, "bold"), command=self.funcNewRecordDeliveryDate)
        btnNameDeliveryDate.grid(row=5, column=2, padx=20)

        lblNameModeOfProc = Label(
            newRecordFrame2, text="MODE OF PROCUREMENT", font=("arial", 20, "bold"))
        lblNameModeOfProc.grid(row=0, column=0, sticky=W)

        newRecordcomboboxNameModeOfProc = ttk.Combobox(
            newRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varModeOfProc)
        newRecordcomboboxNameModeOfProc.grid(row=0, column=1, sticky=E)

        newRecordcomboboxNameModeOfProc['value'] = (
            "SELECT", "GEM", "OPEN TENDER", "LIMITED TENDER")

        newRecordcomboboxNameModeOfProc.current(0)

        lblNameSpecs = Label(
            newRecordFrame2, text="SPECS", font=("arial", 20, "bold"))
        lblNameSpecs.grid(row=1, column=0, sticky=W)

        self.btnNameUploadSpecs = Button(newRecordFrame2, text="UPLOAD FILE", font=(
            "arial", 10, "bold"), command=self.funcNewRecordUploadSpecsFileFirst)
        self.btnNameUploadSpecs.grid(row=1, column=1)

        lblNameAmountEstimate = Label(
            newRecordFrame2, text="AMOUNT ESTIMATE", font=("arial", 20, "bold"))
        lblNameAmountEstimate.grid(row=2, column=0, sticky=W)

        newRecordentryNameAmountEstimate = Entry(
            newRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varAmountEstimate)
        newRecordentryNameAmountEstimate.grid(row=2, column=1, sticky=E)

        lblNameStatus = Label(
            newRecordFrame2, text="STATUS", font=("arial", 20, "bold"))
        lblNameStatus.grid(row=3, column=0, sticky=W)

        newRecordentryNameStatus = Entry(
            newRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varStatus)
        newRecordentryNameStatus.grid(row=3, column=1, sticky=E)

        lblNameActualAmount = Label(
            newRecordFrame2, text="ACTUAL AMOUNT", font=("arial", 20, "bold"))
        lblNameActualAmount.grid(row=4, column=0, sticky=W)

        newRecordentryNameActualAmount = Entry(
            newRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varActualAmount)
        newRecordentryNameActualAmount.grid(row=4, column=1, sticky=E)

        lblNameAdditionalInfo = Label(
            newRecordFrame2, text="ADDITIONAL INFO", font=("arial", 20, "bold"))
        lblNameAdditionalInfo.grid(row=5, column=0, sticky=W)

        newRecordentryNameAdditionalInfo = Entry(
            newRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varAdditionalInfo)
        newRecordentryNameAdditionalInfo.grid(row=5, column=1, sticky=E)

        btnNameNewRecordSave = Button(
            newRecordFrame3, text="SAVE", font=("arial", 20, "bold"), command=self.funcNewRecordDbExecute)
        btnNameNewRecordSave.pack(side=LEFT,
                                  expand=True, pady=10, ipadx=70, ipady=5)

        btnNameNewRecordSave = Button(
            newRecordFrame3, text="CLEAR FIELDS", font=("arial", 20, "bold"), command=self.funcNewRecordClearFields)
        btnNameNewRecordSave.pack(side=LEFT,
                                  expand=True, pady=10, ipadx=70, ipady=5)

        def quit_window():
            self.btnNameNewRecord.config(command=self.funcNewRecord)
            self.newRecordTopLevel.destroy()

        self.newRecordTopLevel.protocol("WM_DELETE_WINDOW", quit_window)

    def funcNewRecordDateRaised(self):
        self.newRecordDateRaisedTopLevel = Toplevel(self.root)
        self.newRecordDateRaisedTopLevel.geometry("400x400+0+0")
        self.newRecordDateRaisedTopLevel.title("SELECT DATE")

        today = datetime.date.today()

        self.DateRaisedCal = Calendar(self.newRecordDateRaisedTopLevel, selectmode='day',
                                      year=today.year, month=today.month,
                                      day=today.day)
        self.DateRaisedCal.pack(padx=20, pady=20, expand=True)

        btnDateRaisedOk = Button(self.newRecordDateRaisedTopLevel, text="SAVE", font=(
            "arial", 20, "bold"), command=self.funcNewRecordDateRaisedFinal)
        btnDateRaisedOk.pack(padx=20, pady=20, expand=True)

    def funcNewRecordDateRaisedFinal(self):

        self.newRecordDateRaisedTopLevel.destroy()

        dt = self.DateRaisedCal.get_date()
        dt1 = datetime.datetime.strptime(dt, '%m/%d/%y')
        self.varDateRaisedShow = dt1.strftime("%d-%m-%Y")
        self.varDateRaised = dt1.strftime("%Y-%m-%d")

        entryNameDateRaised = Label(self.newRecordFrame1, text=self.varDateRaisedShow, font=(
            "arial", 20, "bold"))
        entryNameDateRaised.grid(row=1, column=1)

    def funcNewRecordDeliveryDate(self):
        self.newRecordDeliveryDateTopLevel = Toplevel(self.root)
        self.newRecordDeliveryDateTopLevel.geometry("400x400+0+0")
        self.newRecordDeliveryDateTopLevel.title("SELECT DATE")

        today = datetime.date.today()

        self.DeliveryDateCal = Calendar(self.newRecordDeliveryDateTopLevel, selectmode='day',
                                        year=today.year, month=today.month,
                                        day=today.day)
        self.DeliveryDateCal.pack(padx=20, pady=20, expand=True)

        btnDeliveryDateOk = Button(self.newRecordDeliveryDateTopLevel, text="SAVE", font=(
            "arial", 20, "bold"), command=self.funcNewRecordDeliveryDateFinal)
        btnDeliveryDateOk.pack(padx=20, pady=20, expand=True)

    def funcNewRecordDeliveryDateFinal(self):

        self.newRecordDeliveryDateTopLevel.destroy()

        dt = self.DeliveryDateCal.get_date()
        dt1 = datetime.datetime.strptime(dt, '%m/%d/%y')
        self.varDeliveryDateShow = dt1.strftime("%d-%m-%Y")
        self.varDeliveryDate = dt1.strftime("%Y-%m-%d")

        entryNameDeliveryDate = Label(self.newRecordFrame1, text=self.varDeliveryDateShow, font=(
            "arial", 20, "bold"))
        entryNameDeliveryDate.grid(row=5, column=1)

    def funcNewRecordUploadSpecsFileFirst(self):

        self.btnNameUploadSpecs.config(command="")
        self.varNewRecordSelectFileType = StringVar()

        self.newRecordUploadSpecsFileTopLevel = Toplevel(self.root)
        self.newRecordUploadSpecsFileTopLevel.geometry("400x600+0+0")
        self.newRecordUploadSpecsFileTopLevel.title("SELECT FILE TYPE")

        lblNameSelectExtension = Label(
            self.newRecordUploadSpecsFileTopLevel, text="SELECT FILE TYPE", font=("arial", 20, "bold"))
        lblNameSelectExtension.pack(side=TOP, expand=True)

        newRecordcomboboxNameModeOfProc = ttk.Combobox(
            self.newRecordUploadSpecsFileTopLevel, font=("arial", 20, "bold"), textvariable=self.varNewRecordSelectFileType)
        newRecordcomboboxNameModeOfProc.pack(side=TOP, expand=True)

        newRecordcomboboxNameModeOfProc['value'] = (
            "SELECT", ".pdf", ".docx", ".txt")

        btnNameNewRecordUploadExtensionSpecsFile = Button(
            self.newRecordUploadSpecsFileTopLevel, text="SELECT FILE TO UPLOAD", font=("arial", 20, "bold"), command=self.funcNewRecordUploadSpecsFileSecond)
        btnNameNewRecordUploadExtensionSpecsFile.pack(side=TOP, expand=True)

        def quit_window():
            self.btnNameUploadSpecs.config(
                command=self.funcNewRecordUploadSpecsFileFirst)
            self.newRecordUploadSpecsFileTopLevel.destroy()

        self.newRecordUploadSpecsFileTopLevel.protocol(
            "WM_DELETE_WINDOW", quit_window)

    def funcNewRecordUploadSpecsFileSecond(self):
        self.btnNameUploadSpecs.config(
            command=self.funcNewRecordUploadSpecsFileFirst)
        self.newRecordUploadSpecsFileTopLevel.destroy()

        if self.varNewRecordSelectFileType.get() == ".pdf" or self.varNewRecordSelectFileType.get() == ".docx" or self.varNewRecordSelectFileType.get() == ".txt":
            if self.varNewRecordSelectFileType.get() == ".pdf":
                fn = filedialog.askopenfilename(title="Select File", filetypes=[
                    ("Pdf File", "*.pdf")], parent=self.newRecordTopLevel)
            if self.varNewRecordSelectFileType.get() == ".docx":
                fn = filedialog.askopenfilename(title="Select File", filetypes=[
                    ("Word File", "*.docx")], parent=self.newRecordTopLevel)
            if self.varNewRecordSelectFileType.get() == ".txt":
                fn = filedialog.askopenfilename(title="Select File", filetypes=[
                    ("Text File", "*.txt")], parent=self.newRecordTopLevel)

            if fn != "":
                with open(fn, "rb") as f:
                    self.uploadSpecsBinaryData = f.read()

        else:
            messagebox.showinfo("INFO", "SELECT A FILE TYPE",
                                parent=self.newRecordTopLevel)

    def funcNewRecordClearFields(self):

        self.newRecordTopLevel.destroy()
        self.funcNewRecord()

    def funcNewRecordDbExecute(self):

        if self.varActualAmount.get() == "" or self.varAdditionalInfo.get() == "" or self.varAmountEstimate.get() == 0.0 or self.varAmountEstimate.get() == "" or self.varDateRaised == "" or self.varDeliveryDate == "" or self.varDivision.get() == "" or self.varIndentNo.get() == "" or self.varIndentorName.get() == "" or self.varItemDesc.get() == "" or self.varModeOfProc.get() == "" or self.varModeOfProc.get() == "SELECT" or self.uploadSpecsBinaryData == "" or self.varStatus.get() == "":
            messagebox.showerror(
                "Error", "All fields are required", parent=self.newRecordTopLevel)
        else:
            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "INSERT INTO purchase_details_tbl (indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,specs,amount_estimate,status,actual_amount,additional_info,extension) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"

                cur.execute(sql, (

                    self.varIndentNo.get(),
                    self.varDateRaised,
                    self.varItemDesc.get(),
                    self.varDivision.get(),
                    self.varIndentorName.get(),
                    self.varDeliveryDate,
                    self.varModeOfProc.get(),
                    (self.uploadSpecsBinaryData, ),
                    self.varAmountEstimate.get(),
                    self.varStatus.get(),
                    self.varActualAmount.get(),
                    self.varAdditionalInfo.get(),
                    self.varNewRecordSelectFileType.get()



                ))

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.newRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "DATA SAVED SUCCESSFULLY", parent=self.newRecordTopLevel)

    def funcUpdateRecord(self):

        self.btnNameUpdateRecord.config(command="")

        self.updateRecordTopLevel = Toplevel(self.root)
        self.updateRecordTopLevel.geometry("1540x950+0+0")
        self.updateRecordTopLevel.title("UPDATE RECORD")

        self.varGetUpdateRecord = StringVar()

        self.varUpdateRecordDateRaised = StringVar()
        self.varUpdateRecordItemDesc = StringVar()
        self.varUpdateRecordDivision = StringVar()
        self.varUpdateRecordIndentorName = StringVar()
        self.varUpdateRecordDeliveryDate = StringVar()
        self.varUpdateRecordModeOfProc = StringVar()
        self.varUpdateRecordAmountEstimate = DoubleVar()
        self.varUpdateRecordStatus = StringVar()
        self.varUpdateRecordActualAmount = DoubleVar()
        self.varUpdateRecordAdditionalInfo = StringVar()

        lbltitle = Label(self.updateRecordTopLevel, bd=20, relief=RIDGE, text="UPDATE RECORD",
                         fg="red", bg="white", font=("times new roman", 40, "bold"))
        lbltitle.pack(side=TOP, fill=X)

        self.UpdateRecordenterIndentNoFrame = Frame(
            self.updateRecordTopLevel, bd=20, relief=RIDGE)
        self.UpdateRecordenterIndentNoFrame.pack(fill=BOTH, expand=True,
                                                 padx=(10, 10), pady=(20, 10))

        updateRecordTreeViewFrame = LabelFrame(
            self.updateRecordTopLevel, bd=20, relief=RIDGE)
        updateRecordTreeViewFrame.pack(
            fill=BOTH, expand=True, padx=(10, 10), pady=(10, 10))

        UpdateRecorddataFrame = LabelFrame(self.updateRecordTopLevel, bd=20, relief=RIDGE, font=(
            "arial", 30, "bold"), text="UPDATE RECORD DETAILS")
        UpdateRecorddataFrame.pack(
            fill=BOTH, expand=True, padx=10, pady=(10, 10))

        self.UpdateRecordnewRecordFrame1 = Frame(
            UpdateRecorddataFrame, bd=20, relief=RIDGE)
        self.UpdateRecordnewRecordFrame1.pack(side=LEFT, fill=BOTH, expand=True,
                                              padx=(10, 5), pady=(10, 10))

        UpdateRecordnewRecordFrame2 = Frame(
            UpdateRecorddataFrame, bd=20, relief=RIDGE)
        UpdateRecordnewRecordFrame2.pack(side=LEFT, fill=BOTH,
                                         expand=True, padx=(5, 10), pady=(10, 10))

        UpdateRecordnewRecordFrame3 = Frame(
            self.updateRecordTopLevel, bd=20, relief=RIDGE)
        UpdateRecordnewRecordFrame3.pack(fill=BOTH,
                                         expand=True, padx=(10, 10), pady=(0, 10))

        self.UpdateRecordlblNameEnterIndent = Label(
            self.UpdateRecordenterIndentNoFrame, text="ENTER INDENT NO", font=("arial", 20, "bold"))
        self.UpdateRecordlblNameEnterIndent.pack(side=LEFT, expand=True)

        self.UpdateRecordentryNameEnterIndent = Entry(
            self.UpdateRecordenterIndentNoFrame, font=("arial", 20, "bold"), textvariable=self.varGetUpdateRecord)
        self.UpdateRecordentryNameEnterIndent.pack(side=LEFT, expand=True)

        self.UpdateRecordbtnSearchEnterIndent = Button(
            self.UpdateRecordenterIndentNoFrame, text="SEARCH", font=("arial", 20, "bold"), command=self.funcGetUpdateRecord)
        self.UpdateRecordbtnSearchEnterIndent.pack(side=LEFT, expand=True)

        self.updateRecordTreeView = ttk.Treeview(updateRecordTreeViewFrame, column=("date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                 "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"), height=2)

        self.updateRecordTreeView.pack(fill=X, expand=1)

        style = ttk.Style()
        style.configure("Treeview.Heading",
                        font=(None, 10, BOLD))

        scroll_x = ttk.Scrollbar(
            updateRecordTreeViewFrame, orient=HORIZONTAL, command=self.updateRecordTreeView.xview)

        scroll_x.pack(side=BOTTOM, fill=X)

        self.updateRecordTreeView.configure(xscrollcommand=scroll_x.set)

        self.updateRecordTreeView.heading("date_raised", text="DATE RAISED")
        self.updateRecordTreeView.heading(
            "item_descrp", text="ITEM DESCRIPTION")
        self.updateRecordTreeView.heading("division", text="DIVISION")
        self.updateRecordTreeView.heading(
            "indentor_name", text="INDENTOR NAME")
        self.updateRecordTreeView.heading(
            "delivery_date", text="DELIVERY DATE")
        self.updateRecordTreeView.heading(
            "mode_of_procurement", text="MODE OF PROCUREMENT")
        self.updateRecordTreeView.heading(
            "amount_estimate", text="AMOUNT ESTIMATE")
        self.updateRecordTreeView.heading("status", text="STATUS")
        self.updateRecordTreeView.heading(
            "actual_amount", text="ACTUAL AMOUNT")
        self.updateRecordTreeView.heading(
            "additional_info", text="ADDITIONAL INFO")

        self.updateRecordTreeView.column(
            "date_raised", width=140, anchor=CENTER)
        self.updateRecordTreeView.column(
            "item_descrp", width=140, anchor=CENTER)
        self.updateRecordTreeView.column("division", width=140, anchor=CENTER)
        self.updateRecordTreeView.column(
            "indentor_name", width=140, anchor=CENTER)
        self.updateRecordTreeView.column(
            "delivery_date", width=140, anchor=CENTER)
        self.updateRecordTreeView.column(
            "mode_of_procurement", width=170, anchor=CENTER)
        self.updateRecordTreeView.column(
            "amount_estimate", width=140, anchor=CENTER)
        self.updateRecordTreeView.column("status", width=140, anchor=CENTER)
        self.updateRecordTreeView.column(
            "actual_amount", width=140, anchor=CENTER)
        self.updateRecordTreeView.column(
            "additional_info", width=140, anchor=CENTER)

        self.updateRecordTreeView["show"] = "headings"
        self.updateRecordTreeView.bind(
            "<ButtonRelease-1>", self.funcUpdateRecordGetCursor)

        UpdateRecordlblNameDateRaised = Label(
            self.UpdateRecordnewRecordFrame1, text="DATE RAISED", font=("arial", 20, "bold"))
        UpdateRecordlblNameDateRaised.grid(row=1, column=0, sticky=W)

        self.UpdateRecordbtnNameDateRaised = Button(self.UpdateRecordnewRecordFrame1, text="EDIT", font=(
            "arial", 10, "bold"), command=self.funcSelectUpdateRecordDateRaised)
        self.UpdateRecordbtnNameDateRaised.grid(row=1, column=2, padx=20)

        self.UpdateRecordbtnNameDateRaised["state"] = "disabled"

        UpdateRecordlblNameItemDesc = Label(
            self.UpdateRecordnewRecordFrame1, text="ITEM DESCRP", font=("arial", 20, "bold"))
        UpdateRecordlblNameItemDesc.grid(row=2, column=0, sticky=W)

        UpdateRecordentryNameItemDesc = Entry(
            self.UpdateRecordnewRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordItemDesc)
        UpdateRecordentryNameItemDesc.grid(row=2, column=1)

        UpdateRecordlblNameDivName = Label(
            self.UpdateRecordnewRecordFrame1, text="DIVISION", font=("arial", 20, "bold"))
        UpdateRecordlblNameDivName.grid(row=3, column=0, sticky=W)

        UpdateRecordentryNameDivName = Entry(
            self.UpdateRecordnewRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordDivision)
        UpdateRecordentryNameDivName.grid(row=3, column=1)

        UpdateRecordlblNameIndentorName = Label(
            self.UpdateRecordnewRecordFrame1, text="INDENTOR NAME", font=("arial", 20, "bold"))
        UpdateRecordlblNameIndentorName.grid(row=4, column=0, sticky=W)

        UpdateRecordentryNameIndentorName = Entry(
            self.UpdateRecordnewRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordIndentorName)
        UpdateRecordentryNameIndentorName.grid(row=4, column=1)

        UpdateRecordlblNameDeliveryDate = Label(
            self.UpdateRecordnewRecordFrame1, text="DELIVERY DATE", font=("arial", 20, "bold"))
        UpdateRecordlblNameDeliveryDate.grid(row=5, column=0, sticky=W)

        self.UpdateRecordbtnNameDeliveryDate = Button(self.UpdateRecordnewRecordFrame1, text="EDIT", font=(
            "arial", 10, "bold"), command=self.funcSelectUpdateRecordDeliveryDate)
        self.UpdateRecordbtnNameDeliveryDate.grid(row=5, column=2, padx=20)

        self.UpdateRecordbtnNameDeliveryDate["state"] = "disabled"

        UpdateRecordlblNameModeOfProc = Label(
            UpdateRecordnewRecordFrame2, text="MODE OF PROCUREMENT", font=("arial", 20, "bold"))
        UpdateRecordlblNameModeOfProc.grid(row=0, column=0, sticky=W)

        newRecordcomboboxNameModeOfProc = ttk.Combobox(
            UpdateRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordModeOfProc)
        newRecordcomboboxNameModeOfProc.grid(row=0, column=1, sticky=E)

        newRecordcomboboxNameModeOfProc['value'] = (
            "SELECT", "GEM", "OPEN TENDER", "LIMITED TENDER")

        UpdateRecordlblNameSpecs = Label(
            UpdateRecordnewRecordFrame2, text="SPECS", font=("arial", 20, "bold"))
        UpdateRecordlblNameSpecs.grid(row=1, column=0, sticky=W)

        self.UpdateRecordbtnSpecsView = Button(
            UpdateRecordnewRecordFrame2, text="VIEW FILE", font=("arial", 10, "bold"), command=self.funcUpdateRecordSpecsViewFile)
        self.UpdateRecordbtnSpecsView.grid(row=1, column=1, sticky=W)

        self.UpdateRecordbtnSpecsView["state"] = "disabled"

        self.UpdateRecordbtnSpecsEdit = Button(
            UpdateRecordnewRecordFrame2, text="UPLOAD NEW FILE", font=("arial", 10, "bold"), command=self.funcUpdateRecordSpecsUploadNewFileFirst)
        self.UpdateRecordbtnSpecsEdit.grid(row=1, column=1, sticky=E)

        self.UpdateRecordbtnSpecsEdit["state"] = "disabled"

        UpdateRecordlblNameAmountEstimate = Label(
            UpdateRecordnewRecordFrame2, text="AMOUNT ESTIMATE", font=("arial", 20, "bold"))
        UpdateRecordlblNameAmountEstimate.grid(row=2, column=0, sticky=W)

        UpdateRecordentryNameAmountEstimate = Entry(
            UpdateRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordAmountEstimate)
        UpdateRecordentryNameAmountEstimate.grid(row=2, column=1)

        UpdateRecordlblNameStatus = Label(
            UpdateRecordnewRecordFrame2, text="STATUS", font=("arial", 20, "bold"))
        UpdateRecordlblNameStatus.grid(row=3, column=0, sticky=W)

        UpdateRecordentryNameStatus = Entry(
            UpdateRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordStatus)
        UpdateRecordentryNameStatus.grid(row=3, column=1)

        UpdateRecordlblNameActualAmount = Label(
            UpdateRecordnewRecordFrame2, text="ACTUAL AMOUNT", font=("arial", 20, "bold"))
        UpdateRecordlblNameActualAmount.grid(row=4, column=0, sticky=W)

        UpdateRecordentryNameActualAmount = Entry(
            UpdateRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordActualAmount)
        UpdateRecordentryNameActualAmount.grid(row=4, column=1)

        UpdateRecordlblNameAdditionalInfo = Label(
            UpdateRecordnewRecordFrame2, text="ADDITIONAL INFO", font=("arial", 20, "bold"))
        UpdateRecordlblNameAdditionalInfo.grid(row=5, column=0, sticky=W)

        UpdateRecordentryNameAdditionalInfo = Entry(
            UpdateRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordAdditionalInfo)
        UpdateRecordentryNameAdditionalInfo.grid(row=5, column=1)

        self.UpdateRecordbtnNameNewRecordSave = Button(
            UpdateRecordnewRecordFrame3, text="SAVE", font=("arial", 20, "bold"), command=self.funcUpdateRecordDbExecute)
        self.UpdateRecordbtnNameNewRecordSave.pack(side=TOP,
                                                   expand=True, pady=10, ipadx=70, ipady=5)

        self.UpdateRecordbtnNameNewRecordSave["state"] = "disabled"

        def quit_window():
            self.btnNameUpdateRecord.config(command=self.funcUpdateRecord)
            self.updateRecordTopLevel.destroy()

        self.updateRecordTopLevel.protocol("WM_DELETE_WINDOW", quit_window)

    def funcUpdateRecordGetCursor(self, event=""):
        cursor_row = self.updateRecordTreeView.focus()
        content = self.updateRecordTreeView.item(cursor_row)
        row = content["values"]
        self.varUpdateRecordDateRaised = row[0]
        self.varUpdateRecordDeliveryDate = row[4]
        self.varUpdateRecordItemDesc.set(row[1])
        self.varUpdateRecordDivision.set(row[2])
        self.varUpdateRecordIndentorName.set(row[3])
        self.varUpdateRecordModeOfProc.set(row[5])
        self.varUpdateRecordAmountEstimate.set(row[6])
        self.varUpdateRecordStatus.set(row[7])
        self.varUpdateRecordActualAmount.set(row[8])
        self.varUpdateRecordAdditionalInfo.set(row[9])

        dt1 = self.varUpdateRecordDateRaised
        dt2 = datetime.datetime.strptime(dt1, '%Y-%m-%d')
        self.varUpdateRecordDateRaisedShow = dt2.strftime("%d-%m-%Y")

        UpdateRecordentryNameDateRaised = Label(self.UpdateRecordnewRecordFrame1, text=self.varUpdateRecordDateRaisedShow, font=(
            "arial", 20, "bold"))
        UpdateRecordentryNameDateRaised.grid(row=1, column=1)

        dt3 = self.varUpdateRecordDeliveryDate
        dt4 = datetime.datetime.strptime(dt3, '%Y-%m-%d')
        self.varUpdateRecordDeliveryDateShow = dt4.strftime("%d-%m-%Y")

        UpdateRecordentryNameDeliveryDate = Label(self.UpdateRecordnewRecordFrame1, text=self.varUpdateRecordDeliveryDateShow, font=(
            "arial", 20, "bold"))
        UpdateRecordentryNameDeliveryDate.grid(row=5, column=1)

        self.UpdateRecordbtnSpecsView["state"] = "normal"
        self.UpdateRecordbtnSpecsEdit["state"] = "normal"
        self.UpdateRecordbtnNameNewRecordSave["state"] = "normal"
        self.UpdateRecordbtnNameDateRaised["state"] = "normal"
        self.UpdateRecordbtnNameDeliveryDate["state"] = "normal"

    def funcSelectUpdateRecordDateRaised(self):

        self.UpdateRecordDateRaisedTopLevel = Toplevel(self.root)
        self.UpdateRecordDateRaisedTopLevel.geometry("400x400+0+0")
        self.UpdateRecordDateRaisedTopLevel.title("SELECT DATE")

        today = datetime.date.today()

        self.UpdateDateRaisedCal = Calendar(self.UpdateRecordDateRaisedTopLevel, selectmode='day',
                                            year=today.year, month=today.month,
                                            day=today.day)
        self.UpdateDateRaisedCal.pack(padx=20, pady=20, expand=True)

        btnDateRaisedOk = Button(self.UpdateRecordDateRaisedTopLevel, text="SAVE", font=(
            "arial", 20, "bold"), command=self.funcSelectUpdateRecordDateRaisedFinal)
        btnDateRaisedOk.pack(padx=20, pady=20, expand=True)

    def funcSelectUpdateRecordDateRaisedFinal(self):

        self.UpdateRecordDateRaisedTopLevel.destroy()

        dt = self.UpdateDateRaisedCal.get_date()
        dt1 = datetime.datetime.strptime(dt, '%m/%d/%y')
        self.varUpdateRecordDateRaisedShow = dt1.strftime("%d-%m-%Y")
        self.varUpdateRecordDateRaised = dt1.strftime("%Y-%m-%d")

        UpdateRecordentryNameDateRaised = Label(self.UpdateRecordnewRecordFrame1, text=self.varUpdateRecordDateRaisedShow, font=(
            "arial", 20, "bold"))
        UpdateRecordentryNameDateRaised.grid(row=1, column=1)

    def funcSelectUpdateRecordDeliveryDate(self):

        self.UpdateRecordDeliveryDateTopLevel = Toplevel(self.root)
        self.UpdateRecordDeliveryDateTopLevel.geometry("400x400+0+0")
        self.UpdateRecordDeliveryDateTopLevel.title("SELECT DATE")

        today = datetime.date.today()

        self.UpdateRecordDeliveryDateCal = Calendar(self.UpdateRecordDeliveryDateTopLevel, selectmode='day',
                                                    year=today.year, month=today.month,
                                                    day=today.day)
        self.UpdateRecordDeliveryDateCal.pack(padx=20, pady=20, expand=True)

        btnDeliveryDateOk = Button(self.UpdateRecordDeliveryDateTopLevel, text="SAVE", font=(
            "arial", 20, "bold"), command=self.funcSelectUpdateRecordDeliveryDateFinal)
        btnDeliveryDateOk.pack(padx=20, pady=20, expand=True)

    def funcSelectUpdateRecordDeliveryDateFinal(self):

        self.UpdateRecordDeliveryDateTopLevel.destroy()

        dt = self.UpdateRecordDeliveryDateCal.get_date()
        dt1 = datetime.datetime.strptime(dt, '%m/%d/%y')
        self.varUpdateRecordDeliveryDateShow = dt1.strftime("%d-%m-%Y")
        self.varUpdateRecordDeliveryDate = dt1.strftime("%Y-%m-%d")

        UpdateRecordentryNameDeliveryDate = Label(self.UpdateRecordnewRecordFrame1, text=self.varUpdateRecordDeliveryDateShow, font=(
            "arial", 20, "bold"))
        UpdateRecordentryNameDeliveryDate.grid(row=5, column=1)

    def funcUpdateRecordSpecsViewFile(self):

        if 1 == 1:
            try:
                err1 = False
                err2 = False

                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT specs , extension FROM purchase_details_tbl WHERE indent_no=%s"

                cur.execute(sql, (self.varGetUpdateRecord.get(),))
                r = cur.fetchall()
                for i in r:
                    data = i[0]
                for i in r:
                    extensionFile = i[1]
                if data != None:
                    if extensionFile == '.pdf':

                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{self.varGetUpdateRecord.get()}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".pdf", filetypes=[
                                                          ("Pdf File", "*.pdf")], parent=self.updateRecordTopLevel)
                    if extensionFile == '.docx':

                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{self.varGetUpdateRecord.get()}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".docx", filetypes=[
                                                          ("Word File", "*.docx")], parent=self.updateRecordTopLevel)
                    if extensionFile == '.txt':

                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{self.varGetUpdateRecord.get()}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".txt", filetypes=[
                                                          ("Text File", "*.txt")], parent=self.updateRecordTopLevel)
                    if fn != "":
                        with open(fn, "wb") as f:
                            f.write(data)
                        f.close()
                if data == None:
                    err2 = True
                    messagebox.showinfo("INFO",
                                        "NO FILE FOUND", parent=self.updateRecordTopLevel)

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.updateRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False and err2 == False and fn != ""):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "FILE DOWNLOADED SUCCESSFULLY", parent=self.updateRecordTopLevel)

    def funcUpdateRecordSpecsUploadNewFileFirst(self):

        self.UpdateRecordbtnSpecsEdit.config(command="")
        self.varUpdateRecordSelectFileType = StringVar()

        self.updateRecordUploadSpecsFileTopLevel = Toplevel(self.root)
        self.updateRecordUploadSpecsFileTopLevel.geometry("400x600+0+0")
        self.updateRecordUploadSpecsFileTopLevel.title("SELECT FILE TYPE")

        lblNameSelectExtension = Label(
            self.updateRecordUploadSpecsFileTopLevel, text="SELECT FILE TYPE", font=("arial", 20, "bold"))
        lblNameSelectExtension.pack(side=TOP, expand=True)

        updateRecordcomboboxNameModeOfProc = ttk.Combobox(
            self.updateRecordUploadSpecsFileTopLevel, font=("arial", 20, "bold"), textvariable=self.varUpdateRecordSelectFileType)
        updateRecordcomboboxNameModeOfProc.pack(side=TOP, expand=True)

        updateRecordcomboboxNameModeOfProc['value'] = (
            "SELECT", ".pdf", ".docx", ".txt")

        btnNameUpdateRecordUploadExtensionSpecsFile = Button(
            self.updateRecordUploadSpecsFileTopLevel, text="SELECT FILE TO UPLOAD", font=("arial", 20, "bold"), command=self.funcUpdateRecordSpecsUploadNewFileSecond)
        btnNameUpdateRecordUploadExtensionSpecsFile.pack(side=TOP, expand=True)

        def quit_window():
            self.UpdateRecordbtnSpecsEdit.config(
                command=self.funcUpdateRecordSpecsUploadNewFileFirst)
            self.updateRecordUploadSpecsFileTopLevel.destroy()

        self.updateRecordUploadSpecsFileTopLevel.protocol(
            "WM_DELETE_WINDOW", quit_window)

    def funcUpdateRecordSpecsUploadNewFileSecond(self):
        self.UpdateRecordbtnSpecsEdit.config(
            command=self.funcUpdateRecordSpecsUploadNewFileFirst)
        self.updateRecordUploadSpecsFileTopLevel.destroy()

        if self.varUpdateRecordSelectFileType.get() == ".pdf" or self.varUpdateRecordSelectFileType.get() == ".docx" or self.varUpdateRecordSelectFileType.get() == ".txt":
            if self.varUpdateRecordSelectFileType.get() == ".pdf":
                fn = filedialog.askopenfilename(title="Select File", filetypes=[
                    ("Pdf File", "*.pdf")], parent=self.updateRecordTopLevel)
            if self.varUpdateRecordSelectFileType.get() == ".docx":
                fn = filedialog.askopenfilename(title="Select File", filetypes=[
                    ("Word File", "*.docx")], parent=self.updateRecordTopLevel)
            if self.varUpdateRecordSelectFileType.get() == ".txt":
                fn = filedialog.askopenfilename(title="Select File", filetypes=[
                    ("Text File", "*.txt")], parent=self.updateRecordTopLevel)

            if fn != "":
                with open(fn, "rb") as f:
                    data = f.read()
                try:
                    err1 = False

                    conn = psycopg2.connect(user="postgres",
                                            password="9729",
                                            host="localhost",
                                            port="5432",
                                            database="po_nal_db")

                    cur = conn.cursor()

                    sql = "UPDATE purchase_details_tbl SET specs=%s , extension=%s WHERE indent_no=%s"

                    cur.execute(
                        sql, ((data, ), self.varUpdateRecordSelectFileType.get(), self.varGetUpdateRecord.get()))

                except (Error) as error:
                    err1 = True
                    messagebox.showerror(
                        "DATABASE ERROR", error, parent=self.updateRecordTopLevel)
                    print("Error while connecting to PostgreSQL", error)
                finally:
                    if (conn and err1 == False):
                        cur.close()
                        conn.commit()
                        conn.close()
                        print("PostgreSQL connection is closed")
                        messagebox.showinfo(
                            "INFO", "FILE UPLOADED SUCCESSFULLY", parent=self.updateRecordTopLevel)

        else:
            messagebox.showinfo("INFO", "SELECT A FILE TYPE",
                                parent=self.updateRecordTopLevel)

    def funcGetUpdateRecord(self):

        if self.varGetUpdateRecord.get() == "":
            self.updateRecordTreeView.delete(
                *self.updateRecordTreeView.get_children())
            messagebox.showerror(
                "Error", "Indent No is Required", parent=self.updateRecordTopLevel)

        if self.varGetUpdateRecord.get() != "":
            self.updateRecordTreeView.delete(
                *self.updateRecordTreeView.get_children())
            try:

                err1 = False

                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no FROM purchase_details_tbl WHERE indent_no=%s"

                cur.execute(sql, (self.varGetUpdateRecord.get(),))
                row1 = cur.fetchall()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.updateRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
            if len(row1) != 0:

                if row1[0][0] == self.varGetUpdateRecord.get():
                    self.UpdateRecordlblNameEnterIndent.pack_forget()
                    self.UpdateRecordentryNameEnterIndent.pack_forget()
                    self.UpdateRecordbtnSearchEnterIndent.pack_forget()

                    lblNameUpdateRecordIndentNo = Label(
                        self.UpdateRecordenterIndentNoFrame, text="INDENT NO :", font=("arial", 20, "bold"))
                    lblNameUpdateRecordIndentNo.pack(side=LEFT, expand=True)

                    lblNameUpdateRecordIndentNoShow = Label(
                        self.UpdateRecordenterIndentNoFrame, text=self.varGetUpdateRecord.get(), font=("arial", 20, "bold"))
                    lblNameUpdateRecordIndentNoShow.pack(
                        side=LEFT, expand=True)

                    try:

                        err2 = False

                        conn = psycopg2.connect(user="postgres",
                                                password="9729",
                                                host="localhost",
                                                port="5432",
                                                database="po_nal_db")

                        cur = conn.cursor()

                        sql = "SELECT date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE indent_no=%s"

                        cur.execute(sql, (self.varGetUpdateRecord.get(),))
                        row2 = cur.fetchall()
                        if len(row2) != 0:
                            for i in row2:
                                self.updateRecordTreeView.insert(
                                    "", END, values=i)
                            conn.commit()
                        conn.close()

                    except (Error) as error:
                        err2 = True
                        messagebox.showerror(
                            "DATABASE ERROR", error, parent=self.updateRecordTopLevel)
                        print("Error while connecting to PostgreSQL", error)
                    finally:
                        if (conn and err2 == False):
                            cur.close()
                            print("PostgreSQL connection is closed")
            else:
                messagebox.showinfo(
                    "Info", "indent no. not found", parent=self.updateRecordTopLevel)

    def funcUpdateRecordDbExecute(self):
        if self.varUpdateRecordActualAmount.get() == "" or self.varUpdateRecordAdditionalInfo.get() == "" or self.varUpdateRecordAmountEstimate.get() == 0.0 or self.varUpdateRecordAmountEstimate.get() == "" or self.varUpdateRecordDateRaised == "" or self.varUpdateRecordDeliveryDate == "" or self.varUpdateRecordDivision.get() == "" or self.varGetUpdateRecord.get() == "" or self.varUpdateRecordIndentorName.get() == "" or self.varUpdateRecordItemDesc.get() == "" or self.varUpdateRecordModeOfProc.get() == "" or self.varUpdateRecordModeOfProc.get() == "SELECT" or self.varUpdateRecordStatus.get() == "":
            messagebox.showerror("Error", "Fields Missing",
                                 parent=self.updateRecordTopLevel)
        else:
            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "UPDATE purchase_details_tbl SET date_raised=%s,item_descrp=%s,division=%s,indentor_name=%s,delivery_date=%s,mode_of_procurement=%s,amount_estimate=%s,status=%s,actual_amount=%s,additional_info=%s WHERE indent_no=%s "

                cur.execute(sql, (



                    self.varUpdateRecordDateRaised,
                    self.varUpdateRecordItemDesc.get(),
                    self.varUpdateRecordDivision.get(),
                    self.varUpdateRecordIndentorName.get(),
                    self.varUpdateRecordDeliveryDate,
                    self.varUpdateRecordModeOfProc.get(),
                    self.varUpdateRecordAmountEstimate.get(),
                    self.varUpdateRecordStatus.get(),
                    self.varUpdateRecordActualAmount.get(),
                    self.varUpdateRecordAdditionalInfo.get(),
                    self.varGetUpdateRecord.get()



                ))

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.updateRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "DATA SAVED SUCCESSFULLY", parent=self.updateRecordTopLevel)

    def funcSearchRecord(self):
        self.btnNameSearchRecord.config(command="")
        self.searchRecordTopLevel = Toplevel(self.root)
        self.searchRecordTopLevel.geometry("1540x950+0+0")
        self.searchRecordTopLevel.title("SEARCH RECORD")

        self.varSearchRecordDateRaisedShow = ""
        self.varSearchRecordDateRaised = ""
        self.varSearchRecordDeliveryDateShow = ""
        self.varSearchRecordDeliveryDate = ""

        self.varSearchRecordChooseOptionMenu = StringVar()

        self.varEntryNameSearchRecordEnterFieldIndentNo = StringVar()
        self.varEntryNameSearchRecordEnterFieldItemDesc = StringVar()
        self.varEntryNameSearchRecordEnterFieldDivision = StringVar()
        self.varEntryNameSearchRecordEnterFieldIndentorName = StringVar()
        self.varEntryNameSearchRecordEnterFieldModeOfProc = StringVar()
        self.varEntryNameSearchRecordEnterFieldStatus = StringVar()
        self.dupvarSearchRecordDateRaisedShow = StringVar()
        self.dupvarSearchRecordDeliveryDateShow = StringVar()
        self.varEntryNameSearchRecordEnterFieldDateRaisedYear = StringVar()
        self.varEntryNameSearchRecordEnterFieldDeliveryDateYear = StringVar()
        self.dupvarSearchRecordDateRaisedShow.set("")
        self.dupvarSearchRecordDeliveryDateShow.set("")

        lbltitle = Label(self.searchRecordTopLevel, bd=20, relief=RIDGE, text="SEARCH RECORD",
                         fg="red", bg="white", font=("times new roman", 40, "bold"))
        lbltitle.pack(side=TOP, fill=X)

        SearchRecordChooseFieldFrame1 = Frame(
            self.searchRecordTopLevel, bd=20, relief=RIDGE)
        SearchRecordChooseFieldFrame1.pack(fill=BOTH, expand=True,
                                           padx=(10, 10), pady=(20, 10))

        lblnameSearchRecordChooseField = Label(
            SearchRecordChooseFieldFrame1, text="CHOOSE PARAMETER", font=("arial", 20, "bold"))
        lblnameSearchRecordChooseField.pack(side=LEFT, expand=True)

        self.varSearchRecordChooseOptionMenu.set("SELECT")
        optionMenuSearchRecordChooseField = OptionMenu(SearchRecordChooseFieldFrame1, self.varSearchRecordChooseOptionMenu, "INDENT NO",
                                                       "DATE RAISED", "ITEM DESCRIPTION", "DIVISION", "INDENTOR NAME", "DELIVERY DATE", "MODE OF PROCUREMENT", "STATUS", "INDENT RAISED YEAR", "DELIVERY YEAR", command=self.funcSearchRecordWidgetEnable)
        optionMenuSearchRecordChooseField.config(
            font=("arial", 20, "bold"))
        optionMenuSearchRecordChooseField.pack(side=LEFT, expand=True)

        self.SearchRecordChooseFieldFrame2 = Frame(
            self.searchRecordTopLevel, bd=20, relief=RIDGE)
        self.SearchRecordChooseFieldFrame2.pack(
            fill=BOTH, expand=True, padx=(10, 10), pady=(10, 10))

        lblnameSearchRecordIndentor = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER INDENT NO", font=("arial", 20, "bold"))
        lblnameSearchRecordIndentor.grid(row=0, column=0, sticky=W)
        self.entrynameSearchRecordIndentor = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldIndentNo)
        self.entrynameSearchRecordIndentor.grid(row=0, column=1)

        self.entrynameSearchRecordIndentor["state"] = "disabled"

        lblnameSearchRecordDateRaised = Label(
            self.SearchRecordChooseFieldFrame2, text="SELECT DATE RAISED", font=("arial", 20, "bold"))
        lblnameSearchRecordDateRaised.grid(row=1, column=0, sticky=W)

        self.lblNameSearchRecordDateRaisedShowDate = Entry(
            self.SearchRecordChooseFieldFrame2, textvariable=self.dupvarSearchRecordDateRaisedShow, font=("arial", 20, "bold"), justify=CENTER)
        self.lblNameSearchRecordDateRaisedShowDate.grid(row=1, column=1)

        self.lblNameSearchRecordDateRaisedShowDate["state"] = "disabled"

        self.SearchRecordbtnNameDateRaised = Button(self.SearchRecordChooseFieldFrame2, text="SELECT", font=(
            "arial", 20, "bold"), command=self.funcSearchRecordChooseDateRaised)
        self.SearchRecordbtnNameDateRaised.grid(row=1, column=2)

        self.SearchRecordbtnNameDateRaised["state"] = "disabled"

        lblnameSearchRecordItemDesc = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER ITEM DESC", font=("arial", 20, "bold"))
        lblnameSearchRecordItemDesc.grid(row=2, column=0, sticky=W)
        self.entrynameSearchRecordItemDesc = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldItemDesc)
        self.entrynameSearchRecordItemDesc.grid(row=2, column=1)

        self.entrynameSearchRecordItemDesc["state"] = "disabled"

        lblnameSearchRecordDivision = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER DIVISION", font=("arial", 20, "bold"))
        lblnameSearchRecordDivision.grid(row=3, column=0, sticky=W)
        self.entrynameSearchRecordDivision = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldDivision)
        self.entrynameSearchRecordDivision.grid(row=3, column=1)

        self.entrynameSearchRecordDivision["state"] = "disabled"

        lblnameSearchRecordIndentorName = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER INDENTOR NAME", font=("arial", 20, "bold"))
        lblnameSearchRecordIndentorName.grid(row=4, column=0, sticky=W)
        self.entrynameSearchRecordIndentorName = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldIndentorName)
        self.entrynameSearchRecordIndentorName.grid(row=4, column=1)

        self.entrynameSearchRecordIndentorName["state"] = "disabled"

        lblnameSearchRecordDeliveryDate = Label(
            self.SearchRecordChooseFieldFrame2, text="SELECT DELIVERY DATE", font=("arial", 20, "bold"))
        lblnameSearchRecordDeliveryDate.grid(row=5, column=0, sticky=W)

        self.lblNameSearchRecordDeliveryDateShowDate = Entry(
            self.SearchRecordChooseFieldFrame2, textvariable=self.dupvarSearchRecordDeliveryDateShow, font=("arial", 20, "bold"), justify=CENTER)
        self.lblNameSearchRecordDeliveryDateShowDate.grid(row=5, column=1)

        self.lblNameSearchRecordDeliveryDateShowDate["state"] = "disabled"

        self.SearchRecordbtnNameDeliveryDate = Button(self.SearchRecordChooseFieldFrame2, text="SELECT", font=(
            "arial", 20, "bold"), command=self.funcSearchRecordChooseDeliveryDate)
        self.SearchRecordbtnNameDeliveryDate.grid(row=5, column=2)

        self.SearchRecordbtnNameDeliveryDate["state"] = "disabled"

        lblnameSearchRecordModeOfProc = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER MODE OF PROCUREMENT", font=("arial", 20, "bold"))
        lblnameSearchRecordModeOfProc.grid(row=6, column=0, sticky=W)
        self.entrynameSearchRecordModeOfProc = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldModeOfProc)
        self.entrynameSearchRecordModeOfProc.grid(row=6, column=1)

        self.entrynameSearchRecordModeOfProc["state"] = "disabled"

        lblnameSearchRecordStatus = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER STATUS", font=("arial", 20, "bold"))
        lblnameSearchRecordStatus.grid(row=7, column=0, sticky=W)
        self.entrynameSearchRecordStatus = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldStatus)
        self.entrynameSearchRecordStatus.grid(row=7, column=1)

        self.entrynameSearchRecordStatus["state"] = "disabled"

        lblnameSearchRecordDateRaisedYear = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER INDENT RAISED YEAR", font=("arial", 20, "bold"))
        lblnameSearchRecordDateRaisedYear.grid(row=8, column=0, sticky=W)
        self.entrynameSearchRecordDateRaisedYear = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldDateRaisedYear)
        self.entrynameSearchRecordDateRaisedYear.grid(row=8, column=1)

        self.entrynameSearchRecordDateRaisedYear["state"] = "disabled"

        lblnameSearchRecordDeliveryDateYear = Label(
            self.SearchRecordChooseFieldFrame2, text="ENTER DELIVERY YEAR", font=("arial", 20, "bold"))
        lblnameSearchRecordDeliveryDateYear.grid(row=9, column=0, sticky=W)
        self.entrynameSearchRecordDeliveryDateYear = Entry(self.SearchRecordChooseFieldFrame2, font=(
            "arial", 20, "bold"), textvariable=self.varEntryNameSearchRecordEnterFieldDeliveryDateYear)
        self.entrynameSearchRecordDeliveryDateYear.grid(row=9, column=1)

        self.entrynameSearchRecordDeliveryDateYear["state"] = "disabled"

        self.SearchRecordChooseFieldFrame3 = Frame(
            self.searchRecordTopLevel, bd=20, relief=RIDGE)
        self.SearchRecordChooseFieldFrame3.pack(
            fill=BOTH, expand=True, padx=(10, 10), pady=(10, 10))

        self.btnNameSearchRecordDisplayTreeView = Button(self.SearchRecordChooseFieldFrame3, text="SEARCH", font=(
            "arial", 20, "bold"), command=self.funcSearchRecordExecuteDbFirst)
        self.btnNameSearchRecordDisplayTreeView.pack(expand=True)

        def quit_window():
            self.btnNameSearchRecord.config(command=self.funcSearchRecord)
            self.searchRecordTopLevel.destroy()

        self.searchRecordTopLevel.protocol("WM_DELETE_WINDOW", quit_window)

    def funcSearchRecordWidgetEnable(self, event=""):
        self.varEntryNameSearchRecordEnterFieldIndentNo.set("")
        self.varEntryNameSearchRecordEnterFieldItemDesc.set("")
        self.varEntryNameSearchRecordEnterFieldDivision.set("")
        self.varEntryNameSearchRecordEnterFieldIndentorName.set("")
        self.varEntryNameSearchRecordEnterFieldModeOfProc.set("")
        self.varEntryNameSearchRecordEnterFieldStatus.set("")
        self.dupvarSearchRecordDateRaisedShow.set("")
        self.dupvarSearchRecordDeliveryDateShow.set("")
        self.varEntryNameSearchRecordEnterFieldDateRaisedYear.set("")
        self.varEntryNameSearchRecordEnterFieldDeliveryDateYear.set("")

        self.varSearchRecordDateRaisedShow = ""
        self.varSearchRecordDateRaised = ""
        self.varSearchRecordDeliveryDateShow = ""
        self.varSearchRecordDeliveryDate = ""

        self.entrynameSearchRecordIndentor["state"] = "disabled"
        self.SearchRecordbtnNameDateRaised["state"] = "disabled"
        self.entrynameSearchRecordItemDesc["state"] = "disabled"
        self.entrynameSearchRecordDivision["state"] = "disabled"
        self.entrynameSearchRecordIndentorName["state"] = "disabled"
        self.SearchRecordbtnNameDeliveryDate["state"] = "disabled"
        self.entrynameSearchRecordModeOfProc["state"] = "disabled"
        self.entrynameSearchRecordStatus["state"] = "disabled"
        self.entrynameSearchRecordDateRaisedYear["state"] = "disabled"
        self.entrynameSearchRecordDeliveryDateYear["state"] = "disabled"

        if self.varSearchRecordChooseOptionMenu.get() == "INDENT NO":
            self.entrynameSearchRecordIndentor["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "DATE RAISED":
            self.SearchRecordbtnNameDateRaised["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "ITEM DESCRIPTION":
            self.entrynameSearchRecordItemDesc["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "DIVISION":
            self.entrynameSearchRecordDivision["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "INDENTOR NAME":
            self.entrynameSearchRecordIndentorName["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY DATE":
            self.SearchRecordbtnNameDeliveryDate["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "MODE OF PROCUREMENT":
            self.entrynameSearchRecordModeOfProc["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "STATUS":
            self.entrynameSearchRecordStatus["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR":
            self.entrynameSearchRecordDateRaisedYear["state"] = "normal"

        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":
            self.entrynameSearchRecordDeliveryDateYear["state"] = "normal"

    def funcSearchRecordChooseDateRaised(self):
        self.SearchRecordDateRaisedTopLevel = Toplevel(self.root)
        self.SearchRecordDateRaisedTopLevel.geometry("400x400+0+0")
        self.SearchRecordDateRaisedTopLevel.title("SELECT DATE")

        today = datetime.date.today()

        self.SearchDateRaisedCal = Calendar(self.SearchRecordDateRaisedTopLevel, selectmode='day',
                                            year=today.year, month=today.month,
                                            day=today.day)
        self.SearchDateRaisedCal.pack(padx=20, pady=20, expand=True)

        btnDateRaisedOk = Button(self.SearchRecordDateRaisedTopLevel, text="SAVE", font=(
            "arial", 20, "bold"), command=self.funcSelectSearchRecordDateRaisedFinal)
        btnDateRaisedOk.pack(padx=20, pady=20, expand=True)

    def funcSelectSearchRecordDateRaisedFinal(self):

        self.SearchRecordDateRaisedTopLevel.destroy()

        dt = self.SearchDateRaisedCal.get_date()
        dt1 = datetime.datetime.strptime(dt, '%m/%d/%y')
        self.varSearchRecordDateRaisedShow = dt1.strftime("%d-%m-%Y")
        self.varSearchRecordDateRaised = dt1.strftime("%Y-%m-%d")

        self.dupvarSearchRecordDateRaisedShow.set(
            self.varSearchRecordDateRaisedShow)

        self.lblNameSearchRecordDateRaisedShowDate = Entry(
            self.SearchRecordChooseFieldFrame2, textvariable=self.dupvarSearchRecordDateRaisedShow, font=("arial", 20, "bold"))
        self.lblNameSearchRecordDateRaisedShowDate.grid(row=1, column=1)
        self.lblNameSearchRecordDateRaisedShowDate["state"] = "disabled"

    def funcSearchRecordChooseDeliveryDate(self):
        self.SearchRecordDeliveryDateTopLevel = Toplevel(self.root)
        self.SearchRecordDeliveryDateTopLevel.geometry("400x400+0+0")
        self.SearchRecordDeliveryDateTopLevel.title("SELECT DATE")

        today = datetime.date.today()

        self.SearchDeliveryDateCal = Calendar(self.SearchRecordDeliveryDateTopLevel, selectmode='day',
                                              year=today.year, month=today.month,
                                              day=today.day)
        self.SearchDeliveryDateCal.pack(padx=20, pady=20, expand=True)

        btnDeliveryDateOk = Button(self.SearchRecordDeliveryDateTopLevel, text="SAVE", font=(
            "arial", 20, "bold"), command=self.funcSelectSearchRecordDeliveryDateFinal)
        btnDeliveryDateOk.pack(padx=20, pady=20, expand=True)

    def funcSelectSearchRecordDeliveryDateFinal(self):

        self.SearchRecordDeliveryDateTopLevel.destroy()

        dt = self.SearchDeliveryDateCal.get_date()
        dt1 = datetime.datetime.strptime(dt, '%m/%d/%y')
        self.varSearchRecordDeliveryDateShow = dt1.strftime("%d-%m-%Y")
        self.varSearchRecordDeliveryDate = dt1.strftime("%Y-%m-%d")

        self.dupvarSearchRecordDeliveryDateShow.set(
            self.varSearchRecordDeliveryDateShow)

        self.lblNameSearchRecordDeliveryDateShowDate = Entry(
            self.SearchRecordChooseFieldFrame2, textvariable=self.dupvarSearchRecordDeliveryDateShow, font=("arial", 20, "bold"))
        self.lblNameSearchRecordDeliveryDateShowDate.grid(row=5, column=1)
        self.lblNameSearchRecordDeliveryDateShowDate["state"] = "disabled"

    def funcSearchRecordExecuteDbFirst(self):
        flagSearchRecord = False

        if self.entrynameSearchRecordIndentor.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "INDENT NO":
            flagSearchRecord = False
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)

        if self.varSearchRecordDateRaised == "" and self.varSearchRecordChooseOptionMenu.get() == "DATE RAISED":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if self.entrynameSearchRecordItemDesc.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "ITEM DESCRIPTION":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if self.entrynameSearchRecordDivision.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "DIVISION":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if self.entrynameSearchRecordIndentorName.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "INDENTOR NAME":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if self.varSearchRecordDeliveryDate == "" and self.varSearchRecordChooseOptionMenu.get() == "DELIVERY DATE":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if self.entrynameSearchRecordModeOfProc.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "MODE OF PROCUREMENT":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if self.entrynameSearchRecordStatus.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "STATUS":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)

        if self.entrynameSearchRecordDateRaisedYear.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)

        if self.entrynameSearchRecordDeliveryDateYear.get() == "" and self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":
            flagSearchRecord = True
            messagebox.showerror(
                "ERROR", "SELECTED FIELD IS EMPTY", parent=self.searchRecordTopLevel)
        if flagSearchRecord == False:
            self.funcSearchRecordExecuteDbSecond()

    def funcSearchRecordExecuteDbSecond(self):

        self.btnNameSearchRecordDisplayTreeView.config(command="")
        self.searchRecordExecuteDbTopLevel = Toplevel(self.root)
        self.searchRecordExecuteDbTopLevel.geometry("1540x950+0+0")
        self.searchRecordExecuteDbTopLevel.title("VIEW DATA")

        lbltitle = Label(self.searchRecordExecuteDbTopLevel, bd=20, relief=RIDGE, text="DATA",
                         fg="red", bg="white", font=("times new roman", 40, "bold"))
        lbltitle.pack(side=TOP, fill=X)

        SearchRecordChooseFieldFrame1 = Frame(
            self.searchRecordExecuteDbTopLevel, bd=20, relief=RIDGE)
        SearchRecordChooseFieldFrame1.pack(fill=BOTH, expand=True,
                                           padx=(10, 10), pady=(20, 10))

        self.SearchRecordChooseFieldFrame2 = Frame(
            self.searchRecordExecuteDbTopLevel, bd=20, relief=RIDGE)
        self.SearchRecordChooseFieldFrame2.pack(fill=BOTH, expand=True,
                                                padx=(10, 10), pady=(10, 10))

        if self.varSearchRecordChooseOptionMenu.get() == "INDENT NO":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE indent_no=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordIndentor.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")

        if self.varSearchRecordChooseOptionMenu.get() == "DATE RAISED":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE date_raised=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.varSearchRecordDateRaised,
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")

        if self.varSearchRecordChooseOptionMenu.get() == "ITEM DESCRIPTION":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE item_descrp=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordItemDesc.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")
        if self.varSearchRecordChooseOptionMenu.get() == "DIVISION":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE division=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordDivision.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")
        if self.varSearchRecordChooseOptionMenu.get() == "INDENTOR NAME":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE indentor_name=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordIndentorName.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")
        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY DATE":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE delivery_date=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.varSearchRecordDeliveryDate,
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")
        if self.varSearchRecordChooseOptionMenu.get() == "MODE OF PROCUREMENT":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE mode_of_procurement=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordModeOfProc.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")
        if self.varSearchRecordChooseOptionMenu.get() == "STATUS":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE status=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordStatus.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")

        if self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE date_part('year',date_raised)=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordDateRaisedYear.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")

        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE date_part('year',delivery_date)=%s ORDER BY date_raised DESC "

                cur.execute(sql, (

                    self.entrynameSearchRecordDeliveryDateYear.get(),
                ))

                rows = cur.fetchall()

                self.searchRecordTreeView = ttk.Treeview(SearchRecordChooseFieldFrame1, column=("indent_no", "date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                                                                "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"))

                self.searchRecordTreeView.pack(fill=X, expand=1)

                style = ttk.Style()
                style.configure("Treeview.Heading",
                                font=(None, 10, BOLD))

                scroll_x = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=HORIZONTAL, command=self.searchRecordTreeView.xview)

                scroll_y = ttk.Scrollbar(
                    SearchRecordChooseFieldFrame1, orient=VERTICAL, command=self.searchRecordTreeView.yview)

                scroll_x.pack(side=BOTTOM, fill=X)

                scroll_y.pack(side=RIGHT, fill=Y)

                self.searchRecordTreeView.configure(
                    xscrollcommand=scroll_x.set)

                self.searchRecordTreeView.configure(
                    yscrollcommand=scroll_y.set)

                self.searchRecordTreeView.heading(
                    "indent_no", text="INDENT NO")
                self.searchRecordTreeView.heading(
                    "date_raised", text="DATE RAISED")
                self.searchRecordTreeView.heading(
                    "item_descrp", text="ITEM DESCRIPTION")
                self.searchRecordTreeView.heading("division", text="DIVISION")
                self.searchRecordTreeView.heading(
                    "indentor_name", text="INDENTOR NAME")
                self.searchRecordTreeView.heading(
                    "delivery_date", text="DELIVERY DATE")
                self.searchRecordTreeView.heading(
                    "mode_of_procurement", text="MODE OF PROCUREMENT")
                self.searchRecordTreeView.heading(
                    "amount_estimate", text="AMOUNT ESTIMATE")
                self.searchRecordTreeView.heading("status", text="STATUS")
                self.searchRecordTreeView.heading(
                    "actual_amount", text="ACTUAL AMOUNT")
                self.searchRecordTreeView.heading(
                    "additional_info", text="ADDITIONAL INFO")

                self.searchRecordTreeView.column(
                    "indent_no", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "date_raised", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "item_descrp", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "division", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "indentor_name", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "delivery_date", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "mode_of_procurement", width=170, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "amount_estimate", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "status", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "actual_amount", width=140, anchor=CENTER)
                self.searchRecordTreeView.column(
                    "additional_info", width=140, anchor=CENTER)

                self.searchRecordTreeView["show"] = "headings"

                if len(rows) != 0:
                    self.searchRecordTreeView.delete(
                        *self.searchRecordTreeView.get_children())
                    for i in rows:
                        self.searchRecordTreeView.insert("", END, values=i)
                    cur.close()
                    conn.commit()
                conn.close()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    print("PostgreSQL connection is closed")

        btnNameSearchRecordPrintData = Button(self.SearchRecordChooseFieldFrame2, text="PRINT", font=(
            "arial", 20, "bold"), command=self.funcNameSearchRecordPrintData)
        btnNameSearchRecordPrintData.grid(row=0, column=0)

        btnNameSearchRecordSumOfActualAmount = Button(self.SearchRecordChooseFieldFrame2, text="GET SUM OF ACTUAL AMOUNT", font=(
            "arial", 20, "bold"), command=self.funcNameSearchRecordSumOfActualAmount)
        btnNameSearchRecordSumOfActualAmount.grid(row=0, column=1)

        self.searchRecordTreeView.bind(
            "<ButtonRelease-1>", self.funcNameSearchRecordShowDownloadFileButton)

        def quit_window():
            self.btnNameSearchRecordDisplayTreeView.config(
                command=self.funcSearchRecordExecuteDbSecond)
            self.searchRecordExecuteDbTopLevel.destroy()

        self.searchRecordExecuteDbTopLevel.protocol(
            "WM_DELETE_WINDOW", quit_window)

    def funcNameSearchRecordPrintData(self):

        if self.varSearchRecordChooseOptionMenu.get() == "INDENT NO":
            columnName = 'indent_no'
            var1 = self.entrynameSearchRecordIndentor.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DATE RAISED":
            columnName = 'date_raised'
            var1 = self.varSearchRecordDateRaised
        if self.varSearchRecordChooseOptionMenu.get() == "ITEM DESCRIPTION":
            columnName = 'item_descrp'
            var1 = self.entrynameSearchRecordItemDesc.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DIVISION":
            columnName = 'division'
            var1 = self.entrynameSearchRecordDivision.get()
        if self.varSearchRecordChooseOptionMenu.get() == "INDENTOR NAME":
            columnName = 'indentor_name'
            var1 = self.entrynameSearchRecordIndentorName.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY DATE":
            columnName = 'delivery_date'
            var1 = self.varSearchRecordDeliveryDate
        if self.varSearchRecordChooseOptionMenu.get() == "MODE OF PROCUREMENT":
            columnName = 'mode_of_procurement'
            var1 = self.entrynameSearchRecordModeOfProc.get()
        if self.varSearchRecordChooseOptionMenu.get() == "STATUS":
            columnName = 'status'
            var1 = self.entrynameSearchRecordStatus.get()
        if self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR":
            columnName = 'date_raised'
            var1 = self.entrynameSearchRecordDateRaisedYear.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":
            columnName = 'delivery_date'
            var1 = self.entrynameSearchRecordDeliveryDateYear.get()

        fn = filedialog.asksaveasfilename(
            initialdir=os.getcwd(), initialfile="PO_DETAILS "+datetime.datetime.now().strftime("%d-%m-%Y %H%M%S"), defaultextension=".xlsx", filetypes=[("Excel file", '.xlsx')])

        if fn != "":

            try:

                err1 = False
                conn_string = 'postgresql://postgres:9729@localhost/po_nal_db'
                db = create_engine(conn_string)
                conn = db.connect()

                if self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR" or self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":
                    sql_query = f"SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE date_part('year',{columnName})='{var1}' ORDER BY date_raised DESC"

                else:

                    sql_query = f"SELECT indent_no,date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE {columnName}='{var1}' ORDER BY date_raised DESC"

                df = pd.read_sql(sql_query, conn)

                df['date_raised'] = df['date_raised'].apply(
                    lambda x: pd.Timestamp(x).strftime('%d-%m-%Y'))

                df['delivery_date'] = df['delivery_date'].apply(
                    lambda x: pd.Timestamp(x).strftime('%d-%m-%Y'))

                df.to_excel(fn, index=False)

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if err1 == False:

                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "DATA SAVED SUCCESFULLY", parent=self.searchRecordExecuteDbTopLevel)

    def funcNameSearchRecordShowDownloadFileButton(self, event=""):
        btnNameSearchRecordDownloadSelectedFile = Button(self.SearchRecordChooseFieldFrame2, text="DOWNLOAD SELECTED SPECS FILE", font=(
            "arial", 20, "bold"), command=self.funcNameSearchRecordDownloadSelectedFile)
        btnNameSearchRecordDownloadSelectedFile.grid(row=0, column=2)

    def funcNameSearchRecordDownloadSelectedFile(self):
        cursor_row = self.searchRecordTreeView.focus()
        content = self.searchRecordTreeView.item(cursor_row)
        row = content["values"]

        if 1 == 1:

            try:

                err1 = False
                err2 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql_query = "SELECT specs , extension FROM purchase_details_tbl WHERE indent_no=%s"

                cur.execute(sql_query, (

                    row[0],
                ))

                r = cur.fetchall()
                for i in r:
                    data = i[0]
                for i in r:
                    extensionFile = i[1]

                if data != None:
                    if extensionFile == '.pdf':
                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{row[0]}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".pdf", filetypes=[
                            ("Pdf File", "*.pdf")], parent=self.searchRecordExecuteDbTopLevel)

                        if fn != "":
                            with open(fn, "wb") as f:
                                f.write(data)
                            f.close()

                    if extensionFile == '.docx':
                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{row[0]}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".docx", filetypes=[
                            ("Word File", "*.docx")], parent=self.searchRecordExecuteDbTopLevel)

                        if fn != "":
                            with open(fn, "wb") as f:
                                f.write(data)
                            f.close()

                    if extensionFile == '.txt':
                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{row[0]}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".txt", filetypes=[
                            ("Text File", "*.txt")], parent=self.searchRecordExecuteDbTopLevel)

                        if fn != "":
                            with open(fn, "wb") as f:
                                f.write(data)
                            f.close()
                if data == None:
                    err2 = True
                    messagebox.showinfo("INFO",
                                        "NO FILE FOUND", parent=self.searchRecordExecuteDbTopLevel)

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False and err2 == False and fn != ""):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "FILE SAVED SUCCESFULLY", parent=self.searchRecordExecuteDbTopLevel)

    def funcNameSearchRecordSumOfActualAmount(self):

        if self.varSearchRecordChooseOptionMenu.get() == "INDENT NO":
            columnName = 'indent_no'
            var1 = self.entrynameSearchRecordIndentor.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DATE RAISED":
            columnName = 'date_raised'
            var1 = self.varSearchRecordDateRaised
        if self.varSearchRecordChooseOptionMenu.get() == "ITEM DESCRIPTION":
            columnName = 'item_descrp'
            var1 = self.entrynameSearchRecordItemDesc.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DIVISION":
            columnName = 'division'
            var1 = self.entrynameSearchRecordDivision.get()
        if self.varSearchRecordChooseOptionMenu.get() == "INDENTOR NAME":
            columnName = 'indentor_name'
            var1 = self.entrynameSearchRecordIndentorName.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY DATE":
            columnName = 'delivery_date'
            var1 = self.varSearchRecordDeliveryDate
        if self.varSearchRecordChooseOptionMenu.get() == "MODE OF PROCUREMENT":
            columnName = 'mode_of_procurement'
            var1 = self.entrynameSearchRecordModeOfProc.get()
        if self.varSearchRecordChooseOptionMenu.get() == "STATUS":
            columnName = 'status'
            var1 = self.entrynameSearchRecordStatus.get()
        if self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR":
            columnName = 'date_raised'
            var1 = self.entrynameSearchRecordDateRaisedYear.get()
        if self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":
            columnName = 'delivery_date'
            var1 = self.entrynameSearchRecordDeliveryDateYear.get()

        try:

            err1 = False
            conn = psycopg2.connect(user="postgres",
                                    password="9729",
                                    host="localhost",
                                    port="5432",
                                    database="po_nal_db")

            cur = conn.cursor()

            if self.varSearchRecordChooseOptionMenu.get() == "INDENT RAISED YEAR" or self.varSearchRecordChooseOptionMenu.get() == "DELIVERY YEAR":
                sql_query = f"SELECT SUM(actual_amount) FROM purchase_details_tbl WHERE date_part('year',{columnName})='{var1}'"

                cur.execute(sql_query)

            else:
                sql_query = "SELECT SUM(actual_amount) FROM purchase_details_tbl WHERE "+str(
                    columnName)+"=%s"

                cur.execute(sql_query, (

                    var1,
                ))

            r = cur.fetchall()
            for i in r:
                data1 = i[0]

        except (Error) as error:
            err1 = True
            messagebox.showerror(
                "DATABASE ERROR", error, parent=self.searchRecordExecuteDbTopLevel)
            print("Error while connecting to PostgreSQL", error)
        finally:
            if (conn and err1 == False):
                cur.close()
                conn.commit()
                conn.close()
                print("PostgreSQL connection is closed")
                SearchRecordChooseFieldFrame3 = Frame(
                    self.searchRecordExecuteDbTopLevel, bd=20, relief=RIDGE)
                SearchRecordChooseFieldFrame3.pack(fill=BOTH, expand=True,
                                                   padx=(10, 10), pady=(10, 10))

                labelNameSearchRecordSumOfActualAmountShow = Label(SearchRecordChooseFieldFrame3, text="INR "+str(data1), font=(
                    "arial", 20, "bold"))
                labelNameSearchRecordSumOfActualAmountShow.pack(expand=True)

    def funcDeleteRecord(self):

        self.btnNameDeleteRecord.config(command="")

        self.deleteRecordTopLevel = Toplevel(self.root)
        self.deleteRecordTopLevel.geometry("1540x950+0+0")
        self.deleteRecordTopLevel.title("DELETE RECORD")

        self.varGetDeleteRecord = StringVar()

        self.varDeleteRecordDateRaised = StringVar()
        self.varDeleteRecordItemDesc = StringVar()
        self.varDeleteRecordDivision = StringVar()
        self.varDeleteRecordIndentorName = StringVar()
        self.varDeleteRecordDeliveryDate = StringVar()
        self.varDeleteRecordModeOfProc = StringVar()
        self.varDeleteRecordAmountEstimate = DoubleVar()
        self.varDeleteRecordStatus = StringVar()
        self.varDeleteRecordActualAmount = DoubleVar()
        self.varDeleteRecordAdditionalInfo = StringVar()

        lbltitle = Label(self.deleteRecordTopLevel, bd=20, relief=RIDGE, text="DELETE RECORD",
                         fg="red", bg="white", font=("times new roman", 40, "bold"))
        lbltitle.pack(side=TOP, fill=X)

        self.DeleteRecordenterIndentNoFrame = Frame(
            self.deleteRecordTopLevel, bd=20, relief=RIDGE)
        self.DeleteRecordenterIndentNoFrame.pack(fill=BOTH, expand=True,
                                                 padx=(10, 10), pady=(20, 10))

        deleteRecordTreeViewFrame = LabelFrame(
            self.deleteRecordTopLevel, bd=20, relief=RIDGE)
        deleteRecordTreeViewFrame.pack(
            fill=BOTH, expand=True, padx=(10, 10), pady=(10, 10))

        DeleteRecorddataFrame = LabelFrame(self.deleteRecordTopLevel, bd=20, relief=RIDGE, font=(
            "arial", 30, "bold"), text="DELETE RECORD DETAILS")
        DeleteRecorddataFrame.pack(
            fill=BOTH, expand=True, padx=10, pady=(10, 10))

        self.DeleteRecordnewRecordFrame1 = Frame(
            DeleteRecorddataFrame, bd=20, relief=RIDGE)
        self.DeleteRecordnewRecordFrame1.pack(side=LEFT, fill=BOTH, expand=True,
                                              padx=(10, 5), pady=(10, 10))

        DeleteRecordnewRecordFrame2 = Frame(
            DeleteRecorddataFrame, bd=20, relief=RIDGE)
        DeleteRecordnewRecordFrame2.pack(side=LEFT, fill=BOTH,
                                         expand=True, padx=(5, 10), pady=(10, 10))

        DeleteRecordnewRecordFrame3 = Frame(
            self.deleteRecordTopLevel, bd=20, relief=RIDGE)
        DeleteRecordnewRecordFrame3.pack(fill=BOTH,
                                         expand=True, padx=(10, 10), pady=(0, 10))

        self.DeleteRecordlblNameEnterIndent = Label(
            self.DeleteRecordenterIndentNoFrame, text="ENTER INDENT NO", font=("arial", 20, "bold"))
        self.DeleteRecordlblNameEnterIndent.pack(side=LEFT, expand=True)

        self.DeleteRecordentryNameEnterIndent = Entry(
            self.DeleteRecordenterIndentNoFrame, font=("arial", 20, "bold"), textvariable=self.varGetDeleteRecord)
        self.DeleteRecordentryNameEnterIndent.pack(side=LEFT, expand=True)

        self.DeleteRecordbtnSearchEnterIndent = Button(
            self.DeleteRecordenterIndentNoFrame, text="SEARCH", font=("arial", 20, "bold"), command=self.funcGetDeleteRecord)
        self.DeleteRecordbtnSearchEnterIndent.pack(side=LEFT, expand=True)

        self.deleteRecordTreeView = ttk.Treeview(deleteRecordTreeViewFrame, column=("date_raised", "item_descrp", "division", "indentor_name", "delivery_date",
                                                 "mode_of_procurement", "amount_estimate", "status", "actual_amount", "additional_info"), height=2)

        self.deleteRecordTreeView.pack(fill=X, expand=1)

        style = ttk.Style()
        style.configure("Treeview.Heading",
                        font=(None, 10, BOLD))

        scroll_x = ttk.Scrollbar(
            deleteRecordTreeViewFrame, orient=HORIZONTAL, command=self.deleteRecordTreeView.xview)

        scroll_x.pack(side=BOTTOM, fill=X)

        self.deleteRecordTreeView.configure(xscrollcommand=scroll_x.set)

        self.deleteRecordTreeView.heading("date_raised", text="DATE RAISED")
        self.deleteRecordTreeView.heading(
            "item_descrp", text="ITEM DESCRIPTION")
        self.deleteRecordTreeView.heading("division", text="DIVISION")
        self.deleteRecordTreeView.heading(
            "indentor_name", text="INDENTOR NAME")
        self.deleteRecordTreeView.heading(
            "delivery_date", text="DELIVERY DATE")
        self.deleteRecordTreeView.heading(
            "mode_of_procurement", text="MODE OF PROCUREMENT")
        self.deleteRecordTreeView.heading(
            "amount_estimate", text="AMOUNT ESTIMATE")
        self.deleteRecordTreeView.heading("status", text="STATUS")
        self.deleteRecordTreeView.heading(
            "actual_amount", text="ACTUAL AMOUNT")
        self.deleteRecordTreeView.heading(
            "additional_info", text="ADDITIONAL INFO")

        self.deleteRecordTreeView.column(
            "date_raised", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "item_descrp", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column("division", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "indentor_name", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "delivery_date", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "mode_of_procurement", width=170, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "amount_estimate", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column("status", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "actual_amount", width=140, anchor=CENTER)
        self.deleteRecordTreeView.column(
            "additional_info", width=140, anchor=CENTER)

        self.deleteRecordTreeView["show"] = "headings"
        self.deleteRecordTreeView.bind(
            "<ButtonRelease-1>", self.funcDeleteRecordGetCursor)

        DeleteRecordlblNameDateRaised = Label(
            self.DeleteRecordnewRecordFrame1, text="DATE RAISED", font=("arial", 20, "bold"))
        DeleteRecordlblNameDateRaised.grid(row=1, column=0, sticky=W)

        DeleteRecordlblNameItemDesc = Label(
            self.DeleteRecordnewRecordFrame1, text="ITEM DESCRP", font=("arial", 20, "bold"))
        DeleteRecordlblNameItemDesc.grid(row=2, column=0, sticky=W)

        DeleteRecordentryNameItemDesc = Entry(
            self.DeleteRecordnewRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordItemDesc, justify=CENTER)
        DeleteRecordentryNameItemDesc.grid(row=2, column=1)

        DeleteRecordentryNameItemDesc["state"] = "disabled"

        DeleteRecordlblNameDivName = Label(
            self.DeleteRecordnewRecordFrame1, text="DIVISION", font=("arial", 20, "bold"))
        DeleteRecordlblNameDivName.grid(row=3, column=0, sticky=W)

        DeleteRecordentryNameDivName = Entry(
            self.DeleteRecordnewRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordDivision, justify=CENTER)
        DeleteRecordentryNameDivName.grid(row=3, column=1)

        DeleteRecordentryNameDivName["state"] = "disabled"

        DeleteRecordlblNameIndentorName = Label(
            self.DeleteRecordnewRecordFrame1, text="INDENTOR NAME", font=("arial", 20, "bold"))
        DeleteRecordlblNameIndentorName.grid(row=4, column=0, sticky=W)

        DeleteRecordentryNameIndentorName = Entry(
            self.DeleteRecordnewRecordFrame1, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordIndentorName, justify=CENTER)
        DeleteRecordentryNameIndentorName.grid(row=4, column=1)

        DeleteRecordentryNameIndentorName["state"] = "disabled"

        DeleteRecordlblNameDeliveryDate = Label(
            self.DeleteRecordnewRecordFrame1, text="DELIVERY DATE", font=("arial", 20, "bold"))
        DeleteRecordlblNameDeliveryDate.grid(row=5, column=0, sticky=W)

        DeleteRecordlblNameModeOfProc = Label(
            DeleteRecordnewRecordFrame2, text="MODE OF PROCUREMENT", font=("arial", 20, "bold"))
        DeleteRecordlblNameModeOfProc.grid(row=0, column=0, sticky=W)

        newRecordcomboboxNameModeOfProc = Entry(
            DeleteRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordModeOfProc, justify=CENTER)
        newRecordcomboboxNameModeOfProc.grid(row=0, column=1, sticky=E)

        newRecordcomboboxNameModeOfProc["state"] = "disabled"

        DeleteRecordlblNameSpecs = Label(
            DeleteRecordnewRecordFrame2, text="SPECS", font=("arial", 20, "bold"))
        DeleteRecordlblNameSpecs.grid(row=1, column=0, sticky=W)

        self.DeleteRecordbtnSpecsView = Button(
            DeleteRecordnewRecordFrame2, text="VIEW FILE", font=("arial", 10, "bold"), command=self.funcDeleteRecordSpecsViewFile)
        self.DeleteRecordbtnSpecsView.grid(row=1, column=1, sticky=NSEW)

        self.DeleteRecordbtnSpecsView["state"] = "disabled"

        DeleteRecordlblNameAmountEstimate = Label(
            DeleteRecordnewRecordFrame2, text="AMOUNT ESTIMATE", font=("arial", 20, "bold"))
        DeleteRecordlblNameAmountEstimate.grid(row=2, column=0, sticky=W)

        DeleteRecordentryNameAmountEstimate = Entry(
            DeleteRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordAmountEstimate, justify=CENTER)
        DeleteRecordentryNameAmountEstimate.grid(row=2, column=1)

        DeleteRecordentryNameAmountEstimate["state"] = "disabled"

        DeleteRecordlblNameStatus = Label(
            DeleteRecordnewRecordFrame2, text="STATUS", font=("arial", 20, "bold"))
        DeleteRecordlblNameStatus.grid(row=3, column=0, sticky=W)

        DeleteRecordentryNameStatus = Entry(
            DeleteRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordStatus, justify=CENTER)
        DeleteRecordentryNameStatus.grid(row=3, column=1)

        DeleteRecordentryNameStatus["state"] = "disabled"

        DeleteRecordlblNameActualAmount = Label(
            DeleteRecordnewRecordFrame2, text="ACTUAL AMOUNT", font=("arial", 20, "bold"))
        DeleteRecordlblNameActualAmount.grid(row=4, column=0, sticky=W)

        DeleteRecordentryNameActualAmount = Entry(
            DeleteRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordActualAmount, justify=CENTER)
        DeleteRecordentryNameActualAmount.grid(row=4, column=1)

        DeleteRecordentryNameActualAmount["state"] = "disabled"

        DeleteRecordlblNameAdditionalInfo = Label(
            DeleteRecordnewRecordFrame2, text="ADDITIONAL INFO", font=("arial", 20, "bold"))
        DeleteRecordlblNameAdditionalInfo.grid(row=5, column=0, sticky=W)

        DeleteRecordentryNameAdditionalInfo = Entry(
            DeleteRecordnewRecordFrame2, font=("arial", 20, "bold"), textvariable=self.varDeleteRecordAdditionalInfo, justify=CENTER)
        DeleteRecordentryNameAdditionalInfo.grid(row=5, column=1)

        DeleteRecordentryNameAdditionalInfo["state"] = "disabled"

        self.DeleteRecordbtnNameNewRecordDelete = Button(
            DeleteRecordnewRecordFrame3, text="DELETE RECORD", font=("arial", 20, "bold"), command=self.funcDeleteRecordDbExecute)
        self.DeleteRecordbtnNameNewRecordDelete.pack(side=TOP,
                                                     expand=True, pady=10, ipadx=70, ipady=5)

        self.DeleteRecordbtnNameNewRecordDelete["state"] = "disabled"

        def quit_window():
            self.btnNameDeleteRecord.config(command=self.funcDeleteRecord)
            self.deleteRecordTopLevel.destroy()

        self.deleteRecordTopLevel.protocol("WM_DELETE_WINDOW", quit_window)

    def funcDeleteRecordGetCursor(self, event=""):
        cursor_row = self.deleteRecordTreeView.focus()
        content = self.deleteRecordTreeView.item(cursor_row)
        row = content["values"]
        self.varDeleteRecordDateRaised = row[0]
        self.varDeleteRecordDeliveryDate = row[4]
        self.varDeleteRecordItemDesc.set(row[1])
        self.varDeleteRecordDivision.set(row[2])
        self.varDeleteRecordIndentorName.set(row[3])
        self.varDeleteRecordModeOfProc.set(row[5])
        self.varDeleteRecordAmountEstimate.set(row[6])
        self.varDeleteRecordStatus.set(row[7])
        self.varDeleteRecordActualAmount.set(row[8])
        self.varDeleteRecordAdditionalInfo.set(row[9])

        dt1 = self.varDeleteRecordDateRaised
        dt2 = datetime.datetime.strptime(dt1, '%Y-%m-%d')
        self.varDeleteRecordDateRaisedShow = dt2.strftime("%d-%m-%Y")

        DeleteRecordentryNameDateRaised = Label(self.DeleteRecordnewRecordFrame1, text=self.varDeleteRecordDateRaisedShow, font=(
            "arial", 20, "bold"))
        DeleteRecordentryNameDateRaised.grid(row=1, column=1)

        dt3 = self.varDeleteRecordDeliveryDate
        dt4 = datetime.datetime.strptime(dt3, '%Y-%m-%d')
        self.varDeleteRecordDeliveryDateShow = dt4.strftime("%d-%m-%Y")

        DeleteRecordentryNameDeliveryDate = Label(self.DeleteRecordnewRecordFrame1, text=self.varDeleteRecordDeliveryDateShow, font=(
            "arial", 20, "bold"))
        DeleteRecordentryNameDeliveryDate.grid(row=5, column=1)

        self.DeleteRecordbtnSpecsView["state"] = "normal"
        self.DeleteRecordbtnNameNewRecordDelete["state"] = "normal"

    def funcDeleteRecordSpecsViewFile(self):

        if 1 == 1:
            try:
                err1 = False
                err2 = False

                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT specs , extension  FROM purchase_details_tbl WHERE indent_no=%s"

                cur.execute(sql, (self.varGetDeleteRecord.get(),))
                r = cur.fetchall()
                for i in r:
                    data = i[0]

                for i in r:
                    extensionFile = i[1]

                if data != None:
                    if extensionFile == '.pdf':
                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{self.varGetDeleteRecord.get()}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".pdf", filetypes=[
                            ("Pdf File", "*.pdf")], parent=self.deleteRecordTopLevel)

                        if fn != "":
                            with open(fn, "wb") as f:
                                f.write(data)
                            f.close()

                    if extensionFile == '.docx':
                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{self.varGetDeleteRecord.get()}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".docx", filetypes=[
                            ("Word File", "*.docx")], parent=self.deleteRecordTopLevel)

                        if fn != "":
                            with open(fn, "wb") as f:
                                f.write(data)
                            f.close()

                    if extensionFile == '.txt':
                        fn = filedialog.asksaveasfilename(initialdir=os.getcwd(), initialfile=f"{self.varGetDeleteRecord.get()}_Specs_File "+str(datetime.datetime.now().strftime("%d-%m-%Y %H%M%S")), title="Save File", defaultextension=".txt", filetypes=[
                            ("Text File", "*.txt")], parent=self.deleteRecordTopLevel)

                        if fn != "":
                            with open(fn, "wb") as f:
                                f.write(data)
                            f.close()
                if data == None:
                    err2 = True
                    messagebox.showinfo("INFO",
                                        "NO FILE FOUND", parent=self.deleteRecordTopLevel)

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.deleteRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False and err2 == False and fn != ""):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "DATA DOWNLOADED SUCCESSFULLY", parent=self.deleteRecordTopLevel)

    def funcGetDeleteRecord(self):

        if self.varGetDeleteRecord.get() == "":
            self.deleteRecordTreeView.delete(
                *self.deleteRecordTreeView.get_children())
            messagebox.showerror(
                "Error", "Indent No is Required", parent=self.deleteRecordTopLevel)

        if self.varGetDeleteRecord.get() != "":
            self.deleteRecordTreeView.delete(
                *self.deleteRecordTreeView.get_children())
            try:

                err1 = False

                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "SELECT indent_no FROM purchase_details_tbl WHERE indent_no=%s"

                cur.execute(sql, (self.varGetDeleteRecord.get(),))
                row1 = cur.fetchall()

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.deleteRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
            if len(row1) != 0:

                if row1[0][0] == self.varGetDeleteRecord.get():
                    self.DeleteRecordlblNameEnterIndent.pack_forget()
                    self.DeleteRecordentryNameEnterIndent.pack_forget()
                    self.DeleteRecordbtnSearchEnterIndent.pack_forget()

                    lblNameDeleteRecordIndentNo = Label(
                        self.DeleteRecordenterIndentNoFrame, text="INDENT NO :", font=("arial", 20, "bold"))
                    lblNameDeleteRecordIndentNo.pack(side=LEFT, expand=True)

                    lblNameDeleteRecordIndentNoShow = Label(
                        self.DeleteRecordenterIndentNoFrame, text=self.varGetDeleteRecord.get(), font=("arial", 20, "bold"))
                    lblNameDeleteRecordIndentNoShow.pack(
                        side=LEFT, expand=True)

                    try:

                        err2 = False

                        conn = psycopg2.connect(user="postgres",
                                                password="9729",
                                                host="localhost",
                                                port="5432",
                                                database="po_nal_db")

                        cur = conn.cursor()

                        sql = "SELECT date_raised,item_descrp,division,indentor_name,delivery_date,mode_of_procurement,amount_estimate,status,actual_amount,additional_info FROM purchase_details_tbl WHERE indent_no=%s"

                        cur.execute(sql, (self.varGetDeleteRecord.get(),))
                        row2 = cur.fetchall()
                        if len(row2) != 0:
                            for i in row2:
                                self.deleteRecordTreeView.insert(
                                    "", END, values=i)
                            conn.commit()
                        conn.close()

                    except (Error) as error:
                        err2 = True
                        messagebox.showerror(
                            "DATABASE ERROR", error, parent=self.deleteRecordTopLevel)
                        print("Error while connecting to PostgreSQL", error)
                    finally:
                        if (conn and err2 == False):
                            cur.close()
                            print("PostgreSQL connection is closed")
            else:
                messagebox.showinfo(
                    "Info", "indent no. not found", parent=self.deleteRecordTopLevel)

    def funcDeleteRecordDbExecute(self):

        deleteRecordPromptDeleteRow = messagebox.askyesno(
            "Attention", "ARE YOU SURE THAT YOU WANT TO PERMANENTLY DELETE THIS RECORD?", parent=self.deleteRecordTopLevel)

        if deleteRecordPromptDeleteRow == 1:

            try:

                err1 = False
                conn = psycopg2.connect(user="postgres",
                                        password="9729",
                                        host="localhost",
                                        port="5432",
                                        database="po_nal_db")

                cur = conn.cursor()

                sql = "DELETE FROM purchase_details_tbl WHERE indent_no=%s "

                cur.execute(sql, (




                    self.varGetDeleteRecord.get(),



                ))

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=self.deleteRecordTopLevel)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (conn and err1 == False):
                    cur.close()
                    conn.commit()
                    conn.close()
                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "DATA DELETED SUCCESSFULLY", parent=self.deleteRecordTopLevel)

    def funcUploadCSVFile(self):

        fn = filedialog.askopenfilename(filetypes=[("Excel file", '.xlsx')])

        if fn != "":

            try:

                err1 = False
                conn_string = 'postgresql://postgres:9729@localhost/po_nal_db'
                db = create_engine(conn_string)
                conn = db.connect()

                df = pd.read_excel(fn)

                df['date_raised'] = df['date_raised'].apply(
                    lambda x: pd.Timestamp(x).strftime('%Y-%m-%d'))

                df['delivery_date'] = df['delivery_date'].apply(
                    lambda x: pd.Timestamp(x).strftime('%Y-%m-%d'))

                df.to_sql('purchase_details_tbl',
                          con=conn, if_exists='append', index=False)

            except (Error) as error:
                err1 = True
                messagebox.showerror(
                    "DATABASE ERROR", error, parent=root)
                print("Error while connecting to PostgreSQL", error)
            finally:
                if (err1 == False):

                    print("PostgreSQL connection is closed")
                    messagebox.showinfo(
                        "INFO", "DATA INSERTED SUCCESSFULLY", parent=root)


if __name__ == "__main__":
    root = Tk()
    ob = homePage(root)
    root.mainloop()
