from tkinter import Tk, Label, Button
import tkinter as tk
from tkinter import ttk
from tkinter import *
import pyodbc

from tkcalendar import Calendar, DateEntry

import datetime


class MyFirstGUI:
    def __init__(self, master, *args):

        try: 
            conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                                'Server=RES_ATELIER_1\SERVER4AUTO;'
                                'Database=DB4Auto;'
                                'UID=sa;'
                                'PWD=aaaaaa;'
                                'Trusted_Connection=No;'
                                )
            cursor = conn.cursor()
            print("QA SQL Connection : OK")
        except:
            print("QA SQL Connection impossible")

        sql = """\
                        SELECT TOP(1) CAST(timestamp AS DATE) FROM QAGATE_1_MainTable ORDER BY timeStamp ASC
                    """
        cursor.execute(sql)
        val = cursor.fetchval()

        self.master = master

        cal0 = DateEntry(self.master, maxdate = datetime.datetime.now(), width=12, background='darkblue',
                    foreground='white', borderwidth=2)
        cal0.set_date(datetime.datetime.now())
        cal0.grid(row=0, column = 1, sticky=W)
        
        cal1 = DateEntry(self.master,mindate = val, maxdate = cal0.get_date(), width=12, background='darkblue',
                    foreground='white', borderwidth=2)
        cal1.set_date(val)
        cal1.grid(row=0, sticky=W)


        cal0.configure(mindate=cal1.get_date())
        def callback(eventObject):
            cal1.configure(maxdate = cal0.get_date())
            print("ok")

        def callback1(eventObject):
            cal0.configure(mindate = cal1.get_date())
            print("ok")

        cal0.bind("<<DateEntrySelected>>", callback)
        cal1.bind("<<DateEntrySelected>>", callback1)

        


        self.var1 = tk.IntVar()
        self.check = tk.Checkbutton(self.master, text="490035-2000",command=self.text, variable=self.var1, indicatoron=0)
        self.check.grid(row=1, sticky=W)
        self.var2 = tk.IntVar()
        self.check1 = tk.Checkbutton(self.master, text="490035-3300", variable=self.var2, indicatoron=0)
        self.check1.grid(row=2, sticky=W)
        ttk.Button(self.master, text='Quit', command=self.master.destroy).grid(row=3, sticky=W, pady=4)
        ttk.Button(self.master, text='Show', command=self.var_states).grid(row=4, sticky=W, pady=4)
        ttk.Button(self.master, text='Export', command=self.select).grid(row=7, sticky=W, pady=4)
        self.check2 = ttk.Checkbutton(self.master, command=self.selectall, text="Select all", takefocus = 0)
        self.check2.grid(row=5, sticky=W)
        scrollbar = ttk.Scrollbar(root, orient=VERTICAL)
        self.tree = ttk.Treeview(root, height=4, yscrollcommand=scrollbar.set)
        style = ttk.Style(root)

        style.theme_use("winnative")
        self.tree['show'] = 'headings'
        self.tree['columns'] = ('numero_of', 'date_debut', 'date_fin')
        self.tree.heading('#1', text='Numero OF', anchor='w')
        self.tree.column("#1", stretch="no")
        self.tree.heading("#2", text='Date debut', anchor='w')
        self.tree.column("#2", stretch="no")
        self.tree.heading("#3", text='Date fin', anchor='w')
        self.tree.column("#3", stretch="no")
        self.tree.grid(row=6, column =0)
        scrollbar.config(command=self.tree.yview)
        scrollbar.grid(row=6, column =1 , sticky=N+S+E)

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_OFDate]
                    """
            cursor.execute(sql)
            valProcess = cursor.fetchall()

            for row in valProcess:
                self.tree.insert("", "end", values=[row.currentOF, row.dateDebut, row.dateFin])

        except:
            print("Aucune reference //!\\")


    def yview(self, *args):
        apply(self.tree.yview, args)


    def var_states(self):
        print("490035-2000: %d\n490035-3300: %d" % (self.var1.get(), self.var2.get()))

    def text(self):
        global state

        if(state == 1):
            self.check1['state']= tk.DISABLED
            state = 0
        else:
            self.check1['state']= tk.NORMAL
            state=1

    def select(self):

        reslist = list()
        selection = self.tree.selection()

        for i in selection:
            entry = self.tree.item(i)['values'][0]
            reslist.append(entry)
        for val in reslist:
            print(val)
    
    def selectall(self):
        global state1, i

        list_child = []

        if(state1 == 1):
            for child in self.tree.get_children():
                list_child.insert(i, child)
                print(str(list_child[i]))
            self.tree.selection_set(list_child)
            state1 = 0
            print('ok')
        else:
            for child in self.tree.get_children():
                list_child.insert(i, child)
                print(str(list_child[i]))
            self.tree.selection_remove(list_child)
            state1 = 1
            print('nok')

state = 1
state1 = 1
i=0
root = tk.Tk()
my_gui = MyFirstGUI(root)
root.mainloop()