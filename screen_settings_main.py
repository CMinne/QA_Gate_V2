import tkinter as tk
from tkinter import *
import os

try:
    global PATH_IMAGE_LOGO
    PATH_IMAGE_LOGO = os.path.join(os.path.dirname(__file__), "Image/Logo_QA_Gate_4.0.ico")               # Take the directory relative path 
except (FileNotFoundError):
        print("Wrong file or file path jtekt_logo.ico")

class setting_screen():
    def __init__(link, controller):                                                                 # Initialisation /!\ LINK est le lien entre tout.

        link.controller = controller
        link.i = 1

    def get_pc_screen_size(link):

        link.PcScreenWidth = link.controller.winfo_screenwidth()
        link.PcScreenHeight = link.controller.winfo_screenheight()


    def set_screen(link):

        proportion = (1)
        ScreenHeight = int(link.PcScreenHeight*proportion) - 63
        ScreenWidth = int(link.PcScreenWidth*proportion)
        #print(str(ScreenHeight) + 'x' + str(ScreenWidth))
        
        link.controller.geometry(str(ScreenWidth) + "x" + str(ScreenHeight)+'+-10+0')
        link.controller.configure(background='#f4f4f4')
        link.controller.state('zoomed')


    def set_large_screen(link):

        link.controller.geometry(str( self.PcScreenWidth) + "x" + str( self.PcScreenHeight))
        link.controller.configure(background='#f4f4f4')
        link.controller.wm_attributes("-fullscreen",True)
        link.controller.resizable(0, 0)


    def change_screen_format(link):

        link.controller.wm_attributes("-fullscreen",link.i)
        link.i = not link.i
        link.controller.resizable(0, 0)


    def set_icon(link):

        link.controller.wm_title("QA Gate 4.0 GUI")                                                 # Change le titre

        

        try:
            link.controller.iconbitmap(PATH_IMAGE_LOGO)                                             # Change l'icône
            print("Icon Main : OK")                                                                 # Marqueur OK
        except:
            print("Error loading icon Main")                                                        # Marqueur erreur