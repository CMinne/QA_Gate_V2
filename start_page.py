import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
import syntax
import password_page
from tkinter import *
import os
import threading

try:
    global PATH_IMAGE_GEAR
    PATH_IMAGE_GEAR = os.path.join(os.path.dirname(__file__), "Image/gear.png")                     # Take the directory relative path 
except (FileNotFoundError):
        print("Wrong file or file path gear.png")



class StartPage(tk.Frame):
    
    def __init__(self, parent, controller, processusPageJour, processusPageOF, frames, extractionPage, extractionKPage):
        tk.Frame.__init__(self, parent)

        

        # Controleur principal de la fenêtre
        self.controller = controller

        # Mappage des controlleurs de toutes les frames
        self.processusPageJour = processusPageJour
        self.processusPageOF = processusPageOF
        self.extractionPage = extractionPage
        self.extractionKPage = extractionKPage
        pagePassword = password_page.page_password(controller)

        # Mappage de la bibliothèque de frame
        self.frames = frames

        # Création frame principale
        framePrincipale = tk.Frame(self, background ="#f4f4f4")
        framePrincipale.pack(side = 'top',
                             fill = 'both', 
                             expand = True)                                                         # S'étend sur toute la fenêtre

        # Création frame gauche (insertion dans frame principale)
        self.frameG = tk.Frame(framePrincipale, 
                          height = 0.174*int(self.winfo_screenheight()), 
                          width = 0.167*int(self.winfo_screenwidth()))
        self.frameG.place(relx = 0.12, 
                          rely = 0.5, 
                          anchor = 'center')                                                        # Anchor est la référence de la frame
        self.frameG.pack_propagate(0)

        # Création frame milieu (insertion dans frame principale)
        self.frameM1 = tk.Frame(framePrincipale, 
                           height = 0.174*int(self.winfo_screenheight()), 
                           width = 0.167*int(self.winfo_screenwidth()))
        self.frameM1.place(relx = 0.379, 
                     rely = 0.5, 
                     anchor = 'center')                                                             # Anchor est la référence de la frame
        self.frameM1.pack_propagate(0)

        # Création frame milieu (insertion dans frame principale)
        self.frameM2 = tk.Frame(framePrincipale, 
                           height = 0.174*int(self.winfo_screenheight()), 
                           width = 0.167*int(self.winfo_screenwidth()))
        self.frameM2.place(relx = 0.62, 
                     rely = 0.5, 
                     anchor = 'center')                                                             # Anchor est la référence de la frame
        self.frameM2.pack_propagate(0)


        # Création frame droite (insertion dans frame principale)
        self.frameD = tk.Frame(framePrincipale, 
                          height = 0.174*int(self.winfo_screenheight()), 
                          width = 0.167*int(self.winfo_screenwidth()))
        self.frameD.place(relx = 0.87, 
                     rely = 0.5, 
                     anchor = 'center')                                                             # Anchor est la référence de la frame
        self.frameD.pack_propagate(0)

        # Création frame option (insertion dans frame principale)
        frameOption = tk.Frame(framePrincipale, background = "#f4f4f4")
        frameOption.place(relx = 1, 
                          rely = 1, 
                          anchor = 'se')                                                            # Anchor est la référence de la frame

        CS_Label_Start = ('comicsans', 45, 'bold')

        # Création du label Main
        LabelMain = tk.Label(framePrincipale, 
                             background ="#f4f4f4",
                             text = "QA Gate 4.0 Analyzer",
                             font = CS_Label_Start,
                             borderwidth=2, 
                             relief="solid")                                                        # Cadre noir
        LabelMain.place(relx=0.5, 
                        rely=0.2, 
                        anchor='center')

        # Création du bouton Analyser Jour (insertion dans frame gauche)
        analyserJourButton = ttk.Button(self.frameG, 
                                       text = "Processus Analyser\n            Jour", 
                                       command = lambda: self.starting_processus(self.processusPageJour))
        analyserJourButton.pack(fill = 'both', 
                                expand = True)

        # Création du bouton Analyser OF (insertion dans frame milieu)
        analyserOFButton = ttk.Button(self.frameM1, 
                                    text = "Processus Analyser\n       OF en cours", 
                                    command = lambda: self.starting_processus(self.processusPageOF))
        analyserOFButton.pack(fill = 'both', 
                            expand = True)

        # Création du bouton Research (insertion dans frame droite)
        extractionD = ttk.Button(self.frameM2, 
                                    text = "Extraction données\n        Production", 
                                    command = lambda: self.starting_extraction(self.extractionPage))
        extractionD.pack(fill = 'both', 
                            expand = True)

        # Création du bouton Research (insertion dans frame droite)
        extractionK = ttk.Button(self.frameD, 
                                    text = "Extraction données\n          Kogame", 
                                    command = lambda: self.starting_extraction(self.extractionKPage))
        extractionK.pack(fill = 'both', 
                            expand = True)


        # Préchargement de l'icône "gear.png"
        logoGear = PhotoImage(file = PATH_IMAGE_GEAR)
        logoGear_resize = logoGear.subsample(25, 25)                                                # Changement de taille

        # Création du bouton Option (insertion dans frame option)
        optionButton = tk.Button(frameOption, 
                             text = "",
                             image = logoGear_resize,
                             command = pagePassword.password_popup
                             )
        optionButton.image = logoGear_resize                                                        # Obligatoire, sinon l'image ne se charge pas correctement.
        optionButton.pack()

        self.bind('<Configure>',self.resize)

        # Marqueur
        print("Start Page : OK")

    def starting_processus(self, controllerFrame):
        
        self.controller.show_frame(controllerFrame)                                                 # Affichage de la fenêtre
        frame = self.frames[controllerFrame]                                                        # Recupération de l'objet frame OF ou Jour
        eventT = threading.Thread(target = frame.sql_event)                                         # Thread pour le démarrage de la loop SQL
        eventT.start()                                                                              # Démarrage du thread

    def starting_extraction(self, controllerFrame):
        self.controller.show_frame(controllerFrame)                                                 # Affichage de la fenêtre

    def resize(self, event):
        
        # Resize frame with window size
        self.frameG.configure(width = 0.167*int(self.winfo_width()), 
                              height = 0.184*int(self.winfo_height()))

        self.frameM1.configure(width = 0.167*int(self.winfo_width()), 
                               height = 0.184*int(self.winfo_height()))

        self.frameM2.configure(width = 0.167*int(self.winfo_width()), 
                               height = 0.184*int(self.winfo_height()))

        self.frameD.configure(width = 0.167*int(self.winfo_width()), 
                              height = 0.184*int(self.winfo_height()))

        # Get size of frame
        heightIni = self.frameG['height']
        widthIni = self.frameG['width']

        height = heightIni // 2
        width = widthIni // 2

        # Look up table for font size
        if height < 10 or width < 65:
            height = 10
        elif height < 20 or width < 97:
            height = 12
        elif height < 30 or width < 120:
            height = 15
        elif height < 40 or width < 140:
            height = 18
        else:
            height = 22
        
        # Resize the font
        style = Style()
        style.configure('TButton',
                        font = ('comicsans', height, 'bold'))

