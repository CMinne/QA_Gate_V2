import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
from tkinter import font
from tkinter import *


def fonts():

    CS_Button = font.Font(family='comicsans', size=12, weight='bold')
    CS_Label = font.Font(family='comicsans', size=20, weight='bold')
    CS_Label_U = font.Font(family='comicsans', size=25, weight='bold', underline = True)
    CS_Label_Start = font.Font(family='comicsans', size=30, weight='bold', underline = True)
    return CS_Button, CS_Label, CS_Label_U, CS_Label_Start

def style():

    style = Style()
    style.theme_use('alt')
    style.configure('TButton',
                    bd = 0,
                    font = ('comicsans', 18, 'bold'), 
                    foreground = 'black', 
                    background = 'slate gray'
                    )
    style.configure("TCheckbutton", background="#f4f4f4")
