# to get current date and time
from datetime import datetime
from datetime import date

from tkinter import *
from tkinter import ttk
import tkinter.messagebox
import pymongo
from PIL import ImageTk, Image
from pymongo import MongoClient

# Connection to mongodb
client = pymongo.MongoClient(
    "mongodb+srv://Komalpreet:Dell%40123@komal.rbvrq.mongodb.net/test?retryWrites=true&w=majority")
db = client.python_project
collection = db.student_data

class student:
    def __init__(self, root):
        self.root = root
        blank_space = " "
        self.root.title(200 * blank_space + "Student")
        self.root.geometry("1367x617+0+0")

        MainFrame = Frame(self.root, bd=10, width=1350, height=700, relief=RIDGE, bg="orange2")
        MainFrame.grid()

        TopFrame1 = Frame(MainFrame, bd=5, width=1340, height=50, relief=RIDGE)
        TopFrame1.grid(row=2, column=0)

        TitleFrame = Frame(MainFrame, bd=5, width=1340, height=100, relief=RIDGE, bg="lightblue1")
        TitleFrame.grid(row=0, column=0)

        TopFrame2 = Frame(MainFrame, bd=5, width=1340, height=450, relief=RIDGE)
        TopFrame2.grid(row=1, column=0)
        ##########################################################################################

        LeftFrame = Frame(TopFrame2, bd=5, width=1340, height=100, padx=2, bg="orange2", relief=RIDGE)
        LeftFrame.pack(side=LEFT)

        LeftFrame1 = Frame(LeftFrame, bd=5, width=600, height=180, padx=2, pady=4, relief=RIDGE)
        LeftFrame1.pack(side=TOP, padx=0, pady=0)

        LeftFrame2 = Frame(LeftFrame, bd=5, width=600, height=180, padx=2, relief=RIDGE)
        LeftFrame2.pack(side=TOP, pady=4)

        LeftFrame2Left = Frame(LeftFrame2, bd=5, width=300, height=170, padx=2, relief=RIDGE)
        LeftFrame2Left.pack(side=LEFT, pady=4)

        LeftFrame2Right = Frame(LeftFrame2, bd=5, width=300, height=170, padx=2, relief=RIDGE)
        LeftFrame2Right.pack(side=RIGHT, pady=4)

        RightFrame1 = Frame(TopFrame2, bd=5, width=320, height=400, padx=2, bg="orange2", relief=RIDGE)
        RightFrame1.pack(side=RIGHT)

        RightFrame1a = Frame(RightFrame1, bd=5, width=310, height=300, padx=2, pady=2, bg="orange2", relief=RIDGE)
        RightFrame1a.pack(side=TOP)

        ##-----------------------TEXT VARIABLES-----------------------------------------------##
        StudentID = StringVar()
        Firstname = StringVar()
        Surname = StringVar()
        Address = StringVar()
        Gender = StringVar()
        Mobile = StringVar()

        ##-----------------------TITLE-----------------------------------------------##
        self.lblTitle = Label(TitleFrame, font=('arial', 56, 'bold'), text="Student Management System", bd=7,
                              fg='cyan4', bg="lightblue1")
        self.lblTitle.grid(row=0, column=0, padx=132)

        ##-----------------------FORM-----------------------------------------------##
        self.lblStudentID = Label(LeftFrame1, font=('arial', 12, 'bold'), text="Student ID", bd=7, anchor=W)
        self.lblStudentID.grid(row=0, column=0, sticky=W, padx=5)
        self.txtStudentID = Entry(LeftFrame1, textvariable=StudentID, font=('arial', 12, 'bold'), bd=7, width=40,
                                  justify=LEFT)
        self.txtStudentID.grid(row=0, column=1)

        self.lblFirstname = Label(LeftFrame1, font=('arial', 12, 'bold'), text="Firstname", bd=7, anchor=W)
        self.lblFirstname.grid(row=1, column=0, sticky=W, padx=5)
        self.txtFirstname = Entry(LeftFrame1, textvariable=Firstname, font=('arial', 12, 'bold'), bd=7, width=40,
                                  justify=LEFT)
        self.txtFirstname.grid(row=1, column=1)

        self.lblSurname = Label(LeftFrame1, font=('arial', 12, 'bold'), text="Surname", bd=7, anchor=W)
        self.lblSurname.grid(row=2, column=0, sticky=W, padx=5)
        self.txtSurname = Entry(LeftFrame1, textvariable=Surname, font=('arial', 12, 'bold'), bd=7, width=40,
                                justify=LEFT)
        self.txtSurname.grid(row=2, column=1)

        self.lblAddress = Label(LeftFrame1, font=('arial', 12, 'bold'), text="Address", bd=7, anchor=W)
        self.lblAddress.grid(row=3, column=0, sticky=W, padx=5)
        self.txtAddress = Entry(LeftFrame1, textvariable=Address, font=('arial', 12, 'bold'), bd=7, width=40,
                                justify=LEFT)
        self.txtAddress.grid(row=3, column=1)

        self.lblGender = Label(LeftFrame1, font=('arial', 12, 'bold'), text="Gender", bd=7, anchor=W)
        self.lblGender.grid(row=4, column=0, sticky=W, padx=5)
        self.cboGender = ttk.Combobox(LeftFrame1, textvariable=Gender, font=('arial', 12, 'bold'), state='readonly',
                                      width=39)

        self.cboGender['values'] = ('', 'Female', 'Male')
        self.cboGender.current(0)
        self.cboGender.grid(row=4, column=1)

        self.lblMobile = Label(LeftFrame1, font=('arial', 12, 'bold'), text="Mobile", bd=7, anchor=W)
        self.lblMobile.grid(row=5, column=0, sticky=W, padx=5)
        self.txtMobile = Entry(LeftFrame1, textvariable=Mobile, font=('arial', 12, 'bold'), bd=7, width=40,
                               justify=LEFT)
        self.txtMobile.grid(row=5, column=1)

        ##-----------------------TEXTVARIABLES for comboboxes-----------------------------------------------##
        Programming = StringVar()
        JavaScript = StringVar()
        GeneralMaths = StringVar()
        JavaEE = StringVar()
        Python = StringVar()
        Database = StringVar()
        Spreadsheet = StringVar()
        English = StringVar()

        ##-----------------------ALL METHODS-----------------------------------------------##
        ##-----------------------EXIT-----------------------------------------------##
        def iExit():
            iExit = tkinter.messagebox.askyesno("Student Management System", "Confirm if you want to exit")
            if iExit > 0:
                root.destroy()
                return

        ##-----------------------RESET-----------------------------------------------##
        def Reset():
            StudentID.set("")
            Firstname.set("")
            Surname.set("")
            Address.set("")
            Gender.set("")
            Mobile.set("")

            Programming.set("Completed")
            JavaScript.set("Ongoing")
            JavaEE.set("Upcoming")
            GeneralMaths.set("Ongoing")
            Python.set("Ongoing")
            Database.set("Completed")
            Spreadsheet.set("Completed")
            English.set("Upcoming")

        ##-----------------------ADD NEW DATA-----------------------------------------------##
        def addData():
            now = datetime.now()
            current_time = now.strftime("%H:%M:%S")
            current_date = date.today()
            try:
                StuId = self.txtStudentID.get().replace(" ", "")
                fname = self.txtFirstname.get().replace(" ", "")
                lname = self.txtSurname.get().replace(" ", "")
                addr = self.txtAddress.get()
                gend = self.cboGender.get().replace(" ", "")
                mob = self.txtMobile.get().replace(" ", "")
                if (StuId=="" or fname=="" or lname=="" or addr=="" or gend ==""or mob==""):
                    tkinter.messagebox.showerror("Error","Please fill in all fields")
                else:
                    if any(char.isdigit() for char in fname):
                        tkinter.messagebox.showerror("Error", "Please enter valid first name")
                    elif any(char.isdigit() for char in lname):
                        tkinter.messagebox.showerror("Error", "Please enter valid last name")
                    elif addr.isdigit():
                        tkinter.messagebox.showerror("Error", "Please enter valid address")
                    elif not mob.isdigit() or len(mob)!=10:
                        tkinter.messagebox.showerror("Error", "Please enter valid phone number")
                    else:
                        StudentData = {
                            "_id": int(StuId),
                            "FirstName": fname,
                            "LastName": lname,
                            "Address": addr,
                            "Gender": gend,
                            "Mobile": mob,
                            "Database": self.cboDatabase.get(),
                            "JavaScript": self.cboJavaScript.get(),
                            "Programming": self.cboProgramming.get(),
                            "GeneralMaths": self.cboGeneralMaths.get(),
                            "Python": self.cboPython.get(),
                            "English": self.cboEnglish.get(),
                            "JavaEE": self.cboJavaEE.get(),
                            "Spreadsheet": self.cboSpreadsheet.get()
                        }
                        collection.insert_one(StudentData)
                        StudentID.set("")
                        Firstname.set("")
                        Surname.set("")
                        Address.set("")
                        Gender.set("")
                        Mobile.set("")
                        Programming.set("Completed")
                        JavaScript.set("Ongoing")
                        JavaEE.set("Upcoming")
                        GeneralMaths.set("Ongoing")
                        Python.set("Ongoing")
                        Database.set("Completed")
                        Spreadsheet.set("Completed")
                        English.set("Upcoming")
                        DisplayData()
                        tkinter.messagebox.showinfo("Student Management System",
                                                    "Your data has been submitted on " + str(
                                                        current_date) + " at " + str(current_time))


            except ValueError:
                tkinter.messagebox.showerror("Error","Invalid Student Id")

            except Exception as E:
                tkinter.messagebox.showerror("Error","Student id already exists in the database.")

        ##-----------------------DISPLAY-----------------------------------------------##
        def DisplayData():
            array = list(collection.find())
            self.student_records.delete(*self.student_records.get_children())
            for doc in array:
                # access each document's "_id" key
                print("\ndoc _id:", doc["_id"])
                print(doc["FirstName"])
                self.student_records.insert('', END, values=(doc["_id"], doc["FirstName"], doc["LastName"],doc["Address"],
                                                             doc["Gender"], doc["Mobile"],doc["Database"],
                                                             doc["JavaScript"],doc["Python"]))

        ##-----------------------Getting info-----------------------------------------------##
        def getInfo(ev):
            viewInfo = self.student_records.focus()
            learnData = self.student_records.item(viewInfo)
            row = learnData['values']
            StudentID.set(row[0])
            Firstname.set(row[1])
            Surname.set(row[2])
            Address.set(row[3])
            Gender.set(row[4])
            Mobile.set(row[5])
            Database.set(row[6])
            JavaScript.set(row[7])
            Python.set(row[8])

        ##-----------------------UPDATE-----------------------------------------------##
        def update():
            now = datetime.now()
            current_time = now.strftime("%H:%M:%S")
            current_date = date.today()
            try:
                StuId = int(self.txtStudentID.get())
                array1 = collection.find_one({"_id": StuId})
                if array1 is None:
                    tkinter.messagebox.showerror("Erro", "Student with id " + str(StuId) + " does not exists.")
                    StudentID.set("")
                    Firstname.set("")
                    Surname.set("")
                    Address.set("")
                    Gender.set("")
                    Mobile.set("")
                else:
                    print(array1)

                    fname = self.txtFirstname.get().replace(" ", "")
                    lname = self.txtSurname.get().replace(" ", "")
                    addr = self.txtAddress.get()
                    gend = self.cboGender.get().replace(" ", "")
                    mob = self.txtMobile.get().replace(" ", "")
                    datab = self.cboDatabase.get().replace(" ", "")
                    javaS = self.cboJavaScript.get().replace(" ", "")
                    pyt = self.cboPython.get()
                    eng = self.cboEnglish.get()
                    javaE = self.cboJavaEE.get()
                    mat = self.cboGeneralMaths.get()
                    spreSh = self.cboSpreadsheet.get()
                    prog = self.cboProgramming.get()


                    if (StuId == "" or fname == "" or lname == "" or addr == "" or gend == "" or mob == ""):
                        tkinter.messagebox.showerror("Error", "Please fill in all fields")
                    else:
                        if any(char.isdigit() for char in fname):
                            tkinter.messagebox.showerror("Error", "Please enter valid first name")
                        elif any(char.isdigit() for char in lname):
                            tkinter.messagebox.showerror("Error", "Please enter valid last name")
                        elif addr.isdigit():
                            tkinter.messagebox.showerror("Error", "Please enter valid address")
                        elif not mob.isdigit() or len(mob) != 10:
                            tkinter.messagebox.showerror("Error", "Please enter valid phone number")
                        else:
                            collection.update_many({"_id": StuId},
                                                   {"$set": {"FirstName": fname, "LastName": lname, "Address": addr,
                                                             "Gender": gend, "Mobile": mob,"Database": datab,
                                                             "JavaScript": javaS, "Python": pyt, "English":eng,
                                                             "Programming":prog,"GeneralMaths":mat,
                                                             "Spreadsheet":spreSh,"JavaEE":javaE}})

                            tkinter.messagebox.showinfo("Data updated",
                                                        "Student with id " + str(StuId) + " has been updated on " +
                                                        str(current_date) + " at " + str(current_time))

                            DisplayData()
                            StudentID.set("")
                            Firstname.set("")
                            Surname.set("")
                            Address.set("")
                            Gender.set("")
                            Mobile.set("")
                            Programming.set("Completed")
                            JavaScript.set("Ongoing")
                            JavaEE.set("Upcoming")
                            GeneralMaths.set("Ongoing")
                            Python.set("Ongoing")
                            Database.set("Completed")
                            Spreadsheet.set("Completed")
                            English.set("Upcoming")

            except ValueError:
                tkinter.messagebox.showerror("Error", "Please enter a valid student id")

        ##-----------------------DELETE-----------------------------------------------##
        def delete():
            now = datetime.now()
            current_time = now.strftime("%H:%M:%S")
            current_date = date.today()
            try:
                StuId = int(self.txtStudentID.get())
                array1 = collection.find_one({"_id": StuId})
                print(array1)
                dlt = tkinter.messagebox.askyesno("Student Management System", "Are you sure you want to delete "+str(StuId))
                if dlt>0:
                    collection.delete_one({"_id": StuId})
                    tkinter.messagebox.showinfo("Data deleted", "Student with id "+str(StuId)+" has "
                                            "been deleted on "+str(current_date)+" at "+str(current_time))
                    DisplayData()
                    StudentID.set("")
                    Firstname.set("")
                    Surname.set("")
                    Address.set("")
                    Gender.set("")
                    Mobile.set("")
                    Programming.set("Completed")
                    JavaScript.set("Ongoing")
                    JavaEE.set("Upcoming")
                    GeneralMaths.set("Ongoing")
                    Python.set("Ongoing")
                    Database.set("Completed")
                    Spreadsheet.set("Completed")
                    English.set("Upcoming")
            except Exception as E:
                tkinter.messagebox.showerror("Error","Please select a document from display to delete")
                DisplayData()

        ##-----------------------Print-----------------------------------------------##
        def iPrint():
            DisplayData()
            screen2 = Toplevel(root)
            screen2.title("Print")
            screen2.geometry("1367x617+0+0")
            MainFrame = Frame(screen2, bd=10, width=1350, height=700, relief=RIDGE, bg="orange2")
            MainFrame.pack()

            TitleFrame = Frame(MainFrame, bd=5, width=1340, height=100, relief=RIDGE, bg="lightblue1")
            TitleFrame.pack()

            TopFrame2 = Frame(MainFrame, bd=5, width=1340, height=520, relief=RIDGE)
            TopFrame2.pack()
            self.lblTitle = Label(TitleFrame, font=('arial', 56, 'bold'), text="List of all Students", bd=7,
                                  fg='cyan4', bg="lightblue1")
            self.lblTitle.grid(row=0, column=0, padx=360)

            scroll_y = Scrollbar(TopFrame2, orient=VERTICAL)
            self.student_records = ttk.Treeview(TopFrame2, height=22,
                                                columns=("stdid", "Firstname", "Surname", "Address", "Gender",
                                                         "Mobile", "Database", "JavaScript", "Python","JavaEE",
                                                         "English","Programming","Spreadsheet"),
                                                yscrollcommand=scroll_y.set)

            scroll_y.pack(side=RIGHT, fill=Y)
            self.student_records.heading("stdid", text="Student ID")
            self.student_records.heading("Firstname", text="Firstname")
            self.student_records.heading("Surname", text="Surname")
            self.student_records.heading("Address", text="Address")
            self.student_records.heading("Gender", text="Gender")
            self.student_records.heading("Mobile", text="Mobile")
            self.student_records.heading("Database", text="Database")
            self.student_records.heading("JavaScript", text="JavaScript")
            self.student_records.heading("Python", text="Python")
            self.student_records.heading("JavaEE", text="JavaEE")
            self.student_records.heading("English", text="English")
            self.student_records.heading("Programming", text="Programming")
            self.student_records.heading("Spreadsheet", text="Spreadsheet")

            self.student_records['show'] = 'headings'

            self.student_records.column("stdid", width=100)
            self.student_records.column("Firstname", width=130)
            self.student_records.column("Surname", width=130)
            self.student_records.column("Address", width=130)
            self.student_records.column("Gender", width=100)
            self.student_records.column("Mobile", width=100)
            self.student_records.column("Database", width=90)
            self.student_records.column("JavaScript", width=95)
            self.student_records.column("Python", width=100)
            self.student_records.column("JavaEE", width=70)
            self.student_records.column("English", width=80)
            self.student_records.column("Programming", width=100)
            self.student_records.column("Spreadsheet", width=90)

            self.student_records.pack(fill=BOTH, expand=1)
            array = list(collection.find())

            for doc in array:
                # access each document's "_id" key
                print("\ndoc _id:", doc["_id"])
                print(doc["FirstName"])
                self.student_records.insert('', END, values=(
                    doc["_id"], doc["FirstName"], doc["LastName"], doc["Address"], doc["Gender"],
                    doc["Mobile"], doc["Database"], doc["JavaScript"], doc["Python"],doc['JavaEE'],doc['English'],
                    doc['Programming'],doc['Spreadsheet']))

            for child in self.student_records.get_children():
                print(self.student_records.item(child)["values"])

        ##-----------------------SEARCH-----------------------------------------------##
        def search():
            try:
                StuId = int(self.txtStudentID.get())
                array1 = collection.find_one({"_id": StuId})
                if array1 is None:
                    tkinter.messagebox.showerror("Erro","Student with id " + str(StuId) + " does not exists.")
                    StudentID.set(" ")

                else:
                    print(array1)
                    tkinter.messagebox.showinfo("Information",array1)
                    StudentID.set("")
                    Firstname.set("")
                    Surname.set("")
                    Address.set("")
                    Gender.set("")
                    Mobile.set("")
            except ValueError:
                tkinter.messagebox.showerror("Error","Please enter a valid student id")

        ##-----------------------ABOUT-----------------------------------------------##
        def about():
            screen1 = Toplevel(root)
            screen1.title("About Us")
            screen1.geometry("1367x500+0+0")
            screen1.resizable(False, False)
            MainFrame = Frame(screen1, bd=10, width=1350, height=900, bg = "orange2",relief=RIDGE)
            MainFrame.pack()

            TitleFrame = Frame(MainFrame, bd=5, width=1340, height=50, relief=RIDGE, bg="lightblue1")
            TitleFrame.pack()

            TopFrame2 = Frame(MainFrame, bd=2, width=400, height=200, relief=GROOVE)

            TopFrame2.pack(side = TOP, padx=40, pady=40)

            self.lblTitle = Label(TitleFrame, font=('arial', 56, 'bold'), text="ABOUT US", bd=7,
                                  fg='cyan4', bg="lightblue1")
            self.lblTitle.grid(row=0, column=0, padx=460)

            panel = Label(TopFrame2, text= "Our Experts Team Is Always Available to Help YOU. IF you have any Concerns \n"
                                           " Regarding the Execution of System, Face Any Issues, Feel Free to Contact Us.\n"
                                           " Issue Bugs and get their solution in 24-hours. You can call Us or Email US.\n"'\n'
                                           " Contact Team:\n"
                                           " - KomalPreet Singh , Team Leader, Extension- 753373\n"
                                           " - Monika Garg , Problem Solving Expert, Extension- 753176\n"
                                           " - KomalPreet Kaur , Problem Solving Expert, Extension- 752338\n"
                                           " - Esha Esha , Problem Solving Expert, Extension- 752486\n"'\n'
                                            "Email Us At:"'\n'"pythonpops@hotmail.com"'\n''\n'
                                            "Address : Suite 134,  Bayridge Road, Mississauga, ON, CA"
                                           , font= 20, fg = "black")

            # The Pack geometry manager packs widgets in rows or columns.
            panel.grid(row=1, column=0, padx=50, pady=10)



        ##-----------------------Comboboxes-----------------------------------------------##
        ##-----------------------for Database Subject-----------------------------------------------##
        self.lblDatabase = Label(LeftFrame2Left, font=('arial', 12, 'bold'), text="Database", bd=7, anchor=W)
        self.lblDatabase.grid(row=0, column=0, sticky=W)
        self.cboDatabase = ttk.Combobox(LeftFrame2Left, textvariable=Database, font=('arial', 12, 'bold'),
                                        state='readonly', width=9)
        self.cboDatabase['values'] = ('Completed', 'Ongoing', 'Retake', 'Upcoming')
        self.cboDatabase.current(0)
        self.cboDatabase.grid(row=0, column=1)

        ##-----------------------for JavaScript Subject-----------------------------------------------##
        self.lblJavaScript = Label(LeftFrame2Left, font=('arial', 12, 'bold'), text="JavaScript", bd=7, anchor=W)
        self.lblJavaScript.grid(row=1, column=0, sticky=W)
        self.cboJavaScript = ttk.Combobox(LeftFrame2Left, textvariable=JavaScript, font=('arial', 12, 'bold'),
                                          state='readonly', width=9)
        self.cboJavaScript['values'] = ('Ongoing', 'Completed', 'Retake', 'Upcoming')
        self.cboJavaScript.current(0)
        self.cboJavaScript.grid(row=1, column=1)

        ##-----------------------for Programming Subject-----------------------------------------------##
        self.lblProgramming = Label(LeftFrame2Left, font=('arial', 12, 'bold'), text="Programming", bd=7, anchor=W)
        self.lblProgramming.grid(row=2, column=0, sticky=W)
        self.cboProgramming = ttk.Combobox(LeftFrame2Left, textvariable=Programming, font=('arial', 12, 'bold'),
                                           state='readonly', width=9)
        self.cboProgramming['values'] = ('Completed', 'Ongoing', 'Retake', 'Upcoming')
        self.cboProgramming.current(0)
        self.cboProgramming.grid(row=2, column=1)

        ##-----------------------for general mathematics Subject-----------------------------------------------##
        self.lblGeneralMaths = Label(LeftFrame2Left, font=('arial', 12, 'bold'), text="General Maths", bd=7, anchor=W)
        self.lblGeneralMaths.grid(row=3, column=0, sticky=W)
        self.cboGeneralMaths = ttk.Combobox(LeftFrame2Left, textvariable=GeneralMaths, font=('arial', 12, 'bold'),
                                            state='readonly', width=9)
        self.cboGeneralMaths['values'] = ('Ongoing', 'Completed', 'Retake', 'Upcoming')
        self.cboGeneralMaths.current(0)
        self.cboGeneralMaths.grid(row=3, column=1, pady=6)

        ##-----------------------for Python Subject-----------------------------------------------##
        self.lblPython = Label(LeftFrame2Right, font=('arial', 12, 'bold'), text="Python", bd=7, anchor=W)
        self.lblPython.grid(row=0, column=0, sticky=W)
        self.cboPython = ttk.Combobox(LeftFrame2Right, textvariable=Python, font=('arial', 12, 'bold'),
                                      state='readonly', width=9)
        self.cboPython['values'] = ('Ongoing', 'Completed', 'Retake', 'Upcoming')
        self.cboPython.current(0)
        self.cboPython.grid(row=0, column=1)

        ##-----------------------for English Subject-----------------------------------------------##
        self.lblEnglish = Label(LeftFrame2Right, font=('arial', 12, 'bold'), text="English", bd=7, anchor=E)
        self.lblEnglish.grid(row=1, column=0, sticky=W)
        self.cboEnglish = ttk.Combobox(LeftFrame2Right, textvariable=English, font=('arial', 12, 'bold'),
                                       state='readonly', width=9)
        self.cboEnglish['values'] = ('Upcoming', 'Completed', 'Ongoing', 'Retake')
        self.cboEnglish.current(0)
        self.cboEnglish.grid(row=1, column=1)

        ##-----------------------for JavaEE Subject-----------------------------------------------##
        self.lblJavaEE = Label(LeftFrame2Right, font=('arial', 12, 'bold'), text="Java EE", bd=7, anchor=W)
        self.lblJavaEE.grid(row=2, column=0, sticky=W)
        self.cboJavaEE = ttk.Combobox(LeftFrame2Right, textvariable=JavaEE, font=('arial', 12, 'bold'),
                                      state='readonly', width=9)
        self.cboJavaEE['values'] = ('Upcoming', 'Completed', 'Ongoing', 'Retake')
        self.cboJavaEE.current(0)
        self.cboJavaEE.grid(row=2, column=1)

        ##-----------------------for Spreadsheet Subject-----------------------------------------------##
        self.lblSpreadsheet = Label(LeftFrame2Right, font=('arial', 12, 'bold'), text="Spreadsheet", bd=7,
                                    anchor=W, justify=LEFT)
        self.lblSpreadsheet.grid(row=3, column=0, sticky=W)
        self.cboSpreadsheet = ttk.Combobox(LeftFrame2Right, textvariable=Spreadsheet, font=('arial', 12, 'bold'),
                                           state='readonly', width=9)
        self.cboSpreadsheet['values'] = ('Completed', 'Ongoing', 'Retake', 'Upcoming')
        self.cboSpreadsheet.current(0)
        self.cboSpreadsheet.grid(row=3, column=1, pady=6)

        ##-----------------------TREEVIEW--------------------------------------------------------------##
        scroll_y = Scrollbar(RightFrame1a, orient=VERTICAL)
        self.student_records = ttk.Treeview(RightFrame1a, height=18,
                                            columns=("stdid", "Firstname", "Surname", "Address", "Gender",
                                                     "Mobile", "Database", "JavaScript", "Python"),yscrollcommand=scroll_y.set)

        scroll_y.pack(side=RIGHT, fill=Y)
        self.student_records.heading("stdid", text="Student ID")
        self.student_records.heading("Firstname", text="Firstname")
        self.student_records.heading("Surname", text="Surname")
        self.student_records.heading("Address", text="Address")
        self.student_records.heading("Gender", text="Gender")
        self.student_records.heading("Mobile", text="Mobile")
        self.student_records.heading("Database", text="Database")
        self.student_records.heading("JavaScript", text="JavaScript")
        self.student_records.heading("Python", text="Python")

        self.student_records['show'] = 'headings'

        self.student_records.column("stdid", width=70)
        self.student_records.column("Firstname", width=100)
        self.student_records.column("Surname", width=100)
        self.student_records.column("Address", width=100)
        self.student_records.column("Gender", width=70)
        self.student_records.column("Mobile", width=100)
        self.student_records.column("Database", width=70)
        self.student_records.column("JavaScript", width=95)
        self.student_records.column("Python", width=70)
        self.student_records.pack(fill=BOTH, expand=1)

        self.student_records.bind("<ButtonRelease-1>", getInfo)

        ##-----------------------ALL BUTTONS--------------------------------------------------------------##
        self.btnAddNew = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                                command = addData, width=8, bg='light sea green',
                                text="Add New Data").grid(row=0, column=0, padx=1)

        self.btnPrint = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                               command = iPrint, width=8,bg='light sea green',
                               text="Print").grid(row=0, column=1, padx=1)

        self.btnDisplay = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                                 command = DisplayData, width=8,bg='light sea green',
                                 text="Display").grid(row=0, column=2, padx=1)

        self.btnUpdate = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                                command = update, width=8,bg='light sea green',
                                text="Update").grid(row=0, column=3, padx=1)

        self.btnDelete = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                                command = delete, width=8,bg='light sea green',
                                text="Delete").grid(row=0, column=4, padx=1)

        self.btnSearch = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                                command = search, width=8,bg='light sea green',
                                text="Search").grid(row=0, column=5, padx=1)

        self.btnReset = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                               command = Reset, width=8,bg='green2',
                               text="Reset").grid(row=0, column=6, padx=1)

        self.btnExit = Button(TopFrame1, pady=1, bd=4, fg='black', font=('arial', 16, 'bold'), padx=24,
                              command = iExit, width=8,bg='red',
                              text="Exit").grid(row=0, column=7, padx=1)

        ##-----------------------Adding menu bar to our application--------------------------------------------------------------##
        menubar = Menu(root, bg="cadet blue")
        filemenu = Menu(menubar, tearoff=0, bg="black", fg="white")

        filemenu.add_command(label="Open", command=DisplayData)
        filemenu.add_command(label="Save", command=addData)
        filemenu.add_command(label="Reset", command=Reset)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=root.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        helpmenu = Menu(menubar, tearoff=0, bg="black", fg="white")
        helpmenu.add_command(label="About Us", command = about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        root.config(menu=menubar)

if __name__ == '__main__':
    root = Tk()
    application = student(root)
    root.resizable(False, False)
    root.mainloop()