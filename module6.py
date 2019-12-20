#!python3

import tkinter as tk
from tkinter import ttk
import time
import threading

class Splash(tk.Toplevel):
    def __init__(self, parent):
        tk.Toplevel.__init__(self, parent)
        self.title("Splash")
        progressbar = ttk.Progressbar(self,orient=tk.HORIZONTAL, length=100, mode='indeterminate') 
        progressbar.pack(side="bottom") 
        progressbar.start()
        ## required to make window show before the program gets to the mainloop
        self.update()

class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.withdraw()
        
        t = threading.Thread(target = lambda: Splash(self))
        splash = t.start()
        ## setup stuff goes here
        self.title("Main Window")
        ## simulate a delay while loading
        time.sleep(6)
        

        ## finished loading so destroy splash
        splash.destroy()

        ## show window again
        self.deiconify()

if __name__ == "__main__":
    app = App()
    app.mainloop()