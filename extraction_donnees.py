import sys
import datetime
import threading

import tkinter as tk
from tkinter import *
import tkinter.ttk as ttk
try:
    from tkcalendar import DateEntry
except ImportError:
    print("tkcalendar import impossible")


import win32com.client as win32

import pyodbc 

import SQL_Cursor

import babel.numbers


global PATH_FOLDER_TEMPLATE
PATH_FOLDER_TEMPLATE = '//SERV14/Public_new/IE/Public/Public 4.0/QA Gate 4.0/Rapport production QA Gate 4.0/Rapport_Template/Rapport_prod_template.xlsm'
global PATH_FOLDER_PROD
PATH_FOLDER_PROD = r'\\SERV14\Public_new\IE\Public\Public 4.0\QA Gate 4.0\Rapport production QA Gate 4.0\Rapport_Prod\Rapport_prod_'

# Variable d'état
statePG = 0
stateSG = 1
stateAll = 1
i=0

class ExtractionQAPage(tk.Frame):

    def __init__(self, parent, controller, startPage):

        
        tk.Frame.__init__(self, parent)

        self.controller = controller
        self.startPage = startPage

        self.configure(background="#f4f4f4")


        # Paramétrage des polices de caractère
        fontButton = ('comicsans', 18, 'bold')
        fontLabelMain = ('Segoe UI', 48, 'normal')
        fontLabelCat = ('Segoe UI', 24, 'normal')
        fontLabelDescrip = ('Segoe UI', 36, 'normal')

        # Paramétrage Label

        # Label principale
        self.LabelText = tk.Label(self)
        self.LabelText.place(relx = 0.5,                                                            # Placement relatif par rapport à la taille de la fenêtre en x
                             rely = 0.05,                                                           # Placement relatif par rapport à la taille de la fenêtre en y 
                             #height = 85,                                                          # Hauteur fixe 
                             #width = 974,                                                          # Largeur fixe 
                             anchor = 'center')                                                     # Point d'accroche pour le placement : le centre du label

        self.LabelText.configure(background = '#f4f4f4')                                            # Fond en gris clair
        self.LabelText.configure(font = fontLabelMain)                                              # Police de caractère
        self.LabelText.configure(text = 'Extraction données Production')                            # Texte
        
        # Label reference
        self.LabelReference = tk.Label(self)
        self.LabelReference.place(relx = 0.5,                                                       # Placement relatif par rapport à la taille de la fenêtre en x 
                                  rely = 0.288,                                                     # Placement relatif par rapport à la taille de la fenêtre en y  
                                  #height = 51,                                                     # Hauteur fixe  
                                  #width = 234,                                                     # Largeur fixe  
                                  anchor = 'center')                                                # Point d'accroche pour le placement : le centre du label

        self.LabelReference.configure(background = "#f4f4f4")                                       # Fond en gris clair
        self.LabelReference.configure(font = fontLabelDescrip)                                      # Police de caractère
        self.LabelReference.configure(text = 'Reference')                                           # Texte

        # Label Type de pièce
        self.LabelType = tk.Label(self)
        self.LabelType.place(relx = 0.193,                                                          # Placement relatif par rapport à la taille de la fenêtre en x 
                             rely = 0.495,                                                          # Placement relatif par rapport à la taille de la fenêtre en y 
                             #height = 51,                                                          # Hauteur fixe 
                             #width = 224,                                                          # Largeur fixe 
                             anchor = 'w')                                                          # Point d'accroche pour le placement : l'ouest

        self.LabelType.configure(anchor = 'w')                                                      # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelType.configure(background = "#f4f4f4")                                            # Fond en gris clair
        self.LabelType.configure(font = fontLabelCat)                                               # Police de caractère
        self.LabelType.configure(text = 'Type de pièce :')                                          # Texte

        # Label Réference pièce
        self.LabelRef = tk.Label(self)
        self.LabelRef.place(relx = 0.193,                                                           # Placement relatif par rapport à la taille de la fenêtre en x  
                            rely = 0.384,                                                           # Placement relatif par rapport à la taille de la fenêtre en y  
                            #height = 51,                                                            # Hauteur fixe  
                            #width = 254,                                                            # Largeur fixe 
                            anchor = 'w')                                                           # Point d'accroche pour le placement : l'ouest 
        
        self.LabelRef.configure(anchor='w')                                                         # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelRef.configure(background="#f4f4f4")                                               # Fond en gris clair
        self.LabelRef.configure(font=fontLabelCat)                                                  # Police de caractère
        self.LabelRef.configure(text='''Référence pièce :''')                                       # Texte

        # Label OF
        self.LabelOF = tk.Label(self)
        self.LabelOF.place(relx = 0.193,                                                            # Placement relatif par rapport à la taille de la fenêtre en x 
                           rely = 0.669,                                                            # Placement relatif par rapport à la taille de la fenêtre en y   
                           #height = 51,                                                             # Hauteur fixe 
                           #width = 224,                                                             # Largeur fixe 
                           anchor = 'w')                                                            # Point d'accroche pour le placement : l'ouest

        self.LabelOF.configure(anchor = 'w')                                                        # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelOF.configure(background = "#f4f4f4")                                              # Fond en gris clair
        self.LabelOF.configure(font = fontLabelCat)                                                 # Police de caractère
        self.LabelOF.configure(text = 'Numéro(s) d\'OF :')                                          # Texte

        # Label Date Début
        self.LabelDebut = tk.Label(self)
        self.LabelDebut.place(relx = 0.283,                                                         # Placement relatif par rapport à la taille de la fenêtre en x  
                              rely = 0.208,                                                         # Placement relatif par rapport à la taille de la fenêtre en y  
                              anchor = 'center')                                                    # Point d'accroche pour le placement : le centre

        self.LabelDebut.configure(anchor = 'w')                                                     # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelDebut.configure(background = "#f4f4f4")                                           # Fond en gris clair
        self.LabelDebut.configure(font = fontLabelCat)                                              # Police de caractère
        self.LabelDebut.configure(text = 'Date de début :')                                         # Texte

        # Label Date Fin
        self.LabelFin = tk.Label(self)
        self.LabelFin.place(relx = 0.648,                                                           # Placement relatif par rapport à la taille de la fenêtre en x   
                            rely = 0.208,                                                           # Placement relatif par rapport à la taille de la fenêtre en y  
                            anchor = 'center')                                                      # Point d'accroche pour le placement : le centre

        self.LabelFin.configure(anchor = 'w')                                                       # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelFin.configure(background = "#f4f4f4")                                             # Fond en gris clair
        self.LabelFin.configure(font = fontLabelCat)                                                # Police de caractère
        self.LabelFin.configure(text = 'Date de fin :')                                             # Texte

        # Paramétrage Calendrier

        # Calendrier Calendrier fin
        self.calFin = DateEntry(self, 
                                maxdate = datetime.datetime.now(),                                  # Date maximale acceptée par le calendrier
                                background = 'darkblue',                                            # Fond du calendrier en bleu foncé 
                                foreground = 'white',                                               # Ecriture en blanc 
                                borderwidth = 2)                                                    # Calendrier avec des bordure de 2

        self.calFin.place(relx = 0.775,                                                             # Placement relatif par rapport à la taille de la fenêtre en x    
                          rely = 0.2153,                                                            # Placement relatif par rapport à la taille de la fenêtre en y  
                          height = 30,                                                              # Hauteur fixe 
                          width = 100,                                                              # Largeur fixe 
                          anchor = "center")                                                        # Point d'accroche pour le placement : le centre

        self.calFin.set_date(datetime.datetime.now())                                               # Set de la date du calendrier

        # Calendrier Calendrier début
        self.calDebut = DateEntry(self, 
                                  maxdate = self.calFin.get_date(),                                 # Date maximale acceptée par le calendrier 
                                  background = 'darkblue',                                          # Fond du calendrier en bleu foncé  
                                  foreground = 'white',                                             # Ecriture en blanc  
                                  borderwidth = 2)                                                  # Calendrier avec des bordure de 2

        self.calDebut.place(relx = 0.435,                                                           # Placement relatif par rapport à la taille de la fenêtre en x     
                            rely = 0.2153,                                                          # Placement relatif par rapport à la taille de la fenêtre en y   
                            height = 30,                                                            # Hauteur fixe  
                            width = 100,                                                            # Largeur fixe  
                            anchor = "center")                                                      # Point d'accroche pour le placement : le centre

        # Radio Bouton (2 position)
        
        # Bouton PG 
        self.varPG = tk.BooleanVar()                                                                # Variable du bouton PG
        self.varPG.set(True)
        self.ButtonPG = tk.Checkbutton(self, 
                                       variable = self.varPG,                                       # Lien avec la variable
                                       indicatoron = 0,                                             # Forme du bouton (Bouton clicable classique)
                                       command = self.handlerPG)                                    # Action lors d'un clic
        self.ButtonPG.place(relx = 0.413,                                                           # Placement relatif par rapport à la taille de la fenêtre en x 
                            rely = 0.384,                                                           # Placement relatif par rapport à la taille de la fenêtre en y 
                            height = 45,                                                            # Hauteur fixe 
                            width = 166,                                                            # Largeur fixe
                            anchor = 'center')                                                      # Point d'accroche pour le placement : le centre

        self.ButtonPG.configure(bd = 2,                                                             # Bordure du bouton 
                                font = fontButton,                                                  # Police de caractère 
                                foreground = 'black',                                               # Police en noir 
                                background = 'slate gray')                                          # Fond en slate gray 
                                
        self.ButtonPG.configure(text = 'PG')                                                        # Texte

        # Bouton SG
        self.varSG = tk.BooleanVar()                                                                # Variable du bouton SG 
        self.ButtonSG = tk.Checkbutton(self, 
                                       variable = self.varSG,                                       # Lien avec la variable 
                                       indicatoron = 0,                                             # Forme du bouton (Bouton clicable classique) 
                                       command = self.handlerSG)                                    # Action lors d'un clic
        self.ButtonSG.place(relx = 0.585,                                                           # Placement relatif par rapport à la taille de la fenêtre en x 
                            rely = 0.384,                                                           # Placement relatif par rapport à la taille de la fenêtre en y 
                            height = 45,                                                            # Hauteur fixe 
                            width = 166,                                                            # Largeur fixe
                            anchor = 'center')                                                      # Point d'accroche pour le placement : le centre
        
        self.ButtonSG.configure(bd = 2,                                                             # Bordure du bouton
                                font = fontButton,                                                  # Police de caractère  
                                foreground = 'black',                                               # Police en noir  
                                background = 'slate gray')                                          # Fond en slate gray 
        
        self.ButtonSG.configure(text = 'SG')                                                        # Texte
        self.ButtonSG['state'] = tk.DISABLED                                                        # Descativation du bouton

        # Bouton PG Ref1
        self.varPGRef1 = tk.BooleanVar()                                                            # Variable du bouton PG Ref 1
        self.ButtonPGRef1 = tk.Checkbutton(self, 
                                           variable = self.varPGRef1,                               # Lien avec la variable 
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique) 
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonPGRef1.place(relx = 0.413,                                                       # Placement relatif par rapport à la taille de la fenêtre en x 
                                rely = 0.492,                                                       # Placement relatif par rapport à la taille de la fenêtre en y 
                                height = 45,                                                        # Hauteur fixe  
                                width = 166,                                                        # Largeur fixe
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre
        
        self.ButtonPGRef1.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir   
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonPGRef1.configure(text = '490035-2000')                                           # Texte

        # Bouton PG Ref2
        self.varPGRef2 = tk.BooleanVar()                                                            # Variable du bouton PG Ref 2 
        self.ButtonPGRef2 = tk.Checkbutton(self, 
                                           variable = self.varPGRef2,                               # Lien avec la variable  
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique)  
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonPGRef2.place(relx = 0.585,                                                       # Placement relatif par rapport à la taille de la fenêtre en x  
                                rely = 0.492,                                                       # Placement relatif par rapport à la taille de la fenêtre en y 
                                height = 45,                                                        # Hauteur fixe 
                                width = 166,                                                        # Largeur fixe
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre
        
        self.ButtonPGRef2.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir    
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonPGRef2.configure(text = '490035-2100')                                           # Texte

        # Bouton PG Ref3
        self.varPGRef3 = tk.BooleanVar()                                                            # Variable du bouton PG Ref 3 
        self.ButtonPGRef3 = tk.Checkbutton(self, 
                                           variable = self.varPGRef3,                               # Lien avec la variable 
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique) 
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonPGRef3.place(relx = 0.413,                                                       # Placement relatif par rapport à la taille de la fenêtre en x   
                                rely = 0.561,                                                       # Placement relatif par rapport à la taille de la fenêtre en y  
                                height = 45,                                                        # Hauteur fixe 
                                width = 166,                                                        # Largeur fixe
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre
        
        self.ButtonPGRef3.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir    
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonPGRef3.configure(text = '490035-3200')                                           # Texte

        # Bouton PG Ref4
        self.varPGRef4 = tk.BooleanVar()                                                            # Variable du bouton PG Ref 4 
        self.ButtonPGRef4 = tk.Checkbutton(self, 
                                           variable = self.varPGRef4,                               # Lien avec la variable  
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique) 
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonPGRef4.place(relx = 0.585,                                                       # Placement relatif par rapport à la taille de la fenêtre en x    
                                rely = 0.561,                                                       # Placement relatif par rapport à la taille de la fenêtre en y 
                                height = 45,                                                        # Hauteur fixe 
                                width = 166,                                                        # Largeur fixe
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre
        
        self.ButtonPGRef4.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir    
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonPGRef4.configure(text = '490035-3300')                                           # Texte

        # Bouton SG Ref1
        self.varSGRef1 = tk.BooleanVar()                                                            # Variable du bouton SG Ref 1
        self.ButtonSGRef1 = tk.Checkbutton(self, 
                                           variable = self.varSGRef1,                               # Lien avec la variable   
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique) 
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonSGRef1.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir    
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonSGRef1.configure(text = '490052-0600')                                           # Texte

        # Bouton SG Ref2
        self.varSGRef2 = tk.BooleanVar()                                                            # Variable du bouton SG Ref 2 
        self.ButtonSGRef2 = tk.Checkbutton(self, 
                                           variable = self.varSGRef2,                               # Lien avec la variable 
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique) 
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonSGRef2.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir    
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonSGRef2.configure(text = '490052-0700')                                           # Texte

        # Bouton SG Ref3
        self.varSGRef3 = tk.BooleanVar()                                                            # Variable du bouton SG Ref 3 
        self.ButtonSGRef3 = tk.Checkbutton(self, 
                                           variable = self.varSGRef3,                               # Lien avec la variable  
                                           indicatoron = 0,                                         # Forme du bouton (Bouton clicable classique) 
                                           command = self.threadHandlerRef)                         # Action lors d'un clic
        
        self.ButtonSGRef3.configure(bd = 2,                                                         # Bordure du bouton
                                    font = fontButton,                                              # Police de caractère   
                                    foreground = 'black',                                           # Police en noir    
                                    background = 'slate gray')                                      # Fond en slate gray 
        
        self.ButtonSGRef3.configure(text = '490012-7300')                                           # Texte

        # Treeview

        # Frame Liste OF
        self.Frame_List = tk.Frame(self)

        self.Frame_List.place(relx = 0.5,                                                           # Placement relatif par rapport à la taille de la fenêtre en x 
                              rely = 0.80,                                                          # Placement relatif par rapport à la taille de la fenêtre en y 
                              height = 271,                                                         # Hauteur fixe  
                              width = 450,                                                          # Largeur fixe  
                              anchor = 'center')                                                    # Point d'accroche pour le placement : le centre

        self.Frame_List.configure(background = "#f4f4f4")                                           # Fond en gris clair 
        
        # Radio Bouton Select All
        self.CheckSelect = ttk.Checkbutton(self.Frame_List, 
                                           command = self.selectall,                                # Action lors d'un clic 
                                           text = "Select all",                                     # Texte 
                                           takefocus = 0)                                           # Effet de style quand on clique sur le bouton

        self.CheckSelect.place(relx = 0,                                                            # Placement relatif par rapport à la taille de la fenêtre en x 
                               rely = 0,                                                            # Placement relatif par rapport à la taille de la fenêtre en y 
                               anchor = 'nw')                                                       # Point d'accroche pour le placement : nord ouest

        # Scrollbar liste OF
        self.scrollbar = tk.Scrollbar(self.Frame_List, 
                                      orient = VERTICAL)                                            # Orientation verticale du scrollbar

        self.scrollbar.place(relx = 0.95,                                                           # Placement relatif par rapport à la taille de la fenêtre en x  
                             rely = 0.07,                                                           # Placement relatif par rapport à la taille de la fenêtre en y
                             relheight = 0.93,                                                      # Hauteur relative
                             relwidth = 0.05)                                                       # Largeur relative

        self.scrollbar.config(bg = '#f4f4f4')                                                       # Fond en gris clair 
        
        # Liste OF
        self.tree = ttk.Treeview(self.Frame_List, 
                                 yscrollcommand = self.scrollbar.set)                               # Lien entre la liste d'OF - Scrollbar

        # Paramétrage des colonnes et index de colonnes
        self.tree['show'] = 'headings'
        self.tree['columns'] = ('numero_of', 'reference', 'date_debut', 'date_fin')
        self.tree.heading('#1', text = 'Numero OF', anchor = 'w')
        self.tree.column("#1", stretch = "YES", minwidth = 0, width = 40)
        self.tree.heading("#2", text = 'Reference', anchor = 'w')
        self.tree.column("#2", stretch = "YES", minwidth = 0, width = 40)
        self.tree.heading("#3", text = 'Date debut', anchor = 'w')
        self.tree.column("#3", stretch = "YES", minwidth = 0, width = 100)
        self.tree.heading("#4", text = 'Date fin', anchor = 'w')
        self.tree.column("#4", stretch = "YES", minwidth = 0, width = 100)

        self.tree.place(relx = 0,                                                                   # Placement relatif par rapport à la taille de la fenêtre en x  
                        rely = 0.07,                                                                # Placement relatif par rapport à la taille de la fenêtre en y 
                        relheight = 0.93,                                                           # Hauteur relative 
                        relwidth = 0.95)                                                            # Largeur relative

        # Configuration lien Scrollbar - Liste OF
        self.scrollbar.config(command = self.tree.yview)

        # Progress Bar export données
        self.progressbar = ttk.Progressbar(self)

        # Label Progress Bar
        self.LabelBar = tk.Label(self)
        self.LabelBar.configure(background = "#f4f4f4")                                             # Fond en gris clair 
        self.LabelBar['text'] = '(#/#)'                                                             # Texte

        # Création du bouton Extraction
        self.extractionButton = tk.Button(self, 
                                          text = "Extraction",                                      # Texte 
                                          command = self.select)                                    # Action lors d'un clic 

        self.extractionButton.configure(bd = 2,                                                     # Bordure du bouton
                                        font = fontButton,                                          # Police de caractère   
                                        foreground = 'black',                                       # Police en noir    
                                        background = 'slate gray')                                  # Fond en slate gray 

        self.extractionButton.place(relx = 0.7,                                                     # Placement relatif par rapport à la taille de la fenêtre en x 
                                    rely = 0.85,                                                    # Placement relatif par rapport à la taille de la fenêtre en y 
                                    height = 45,                                                    # Hauteur fixe 
                                    width = 140,                                                    # Largeur fixe 
                                    anchor = 'center')                                              # Point d'accroche pour le placement : le centre


        # Création du bouton Retour (Retour menu principal)
        retourButton = tk.Button(self, 
                                 text = "Retour",                                                   # Texte 
                                 command = self.closing_extraction,
                                 bd = 2,
                                 foreground = 'black',                                              # Police en noir 
                                 background = 'slate gray') 

        retourButton.place(relx = 0.971,                                                            # Placement relatif par rapport à la taille de la fenêtre en x 
                           rely = 0.941,                                                            # Placement relatif par rapport à la taille de la fenêtre en y  
                           height = 45,                                                             # Hauteur fixe  
                           width = 80,                                                              # Largeur fixe 
                           anchor = 'center')                                                       # Point d'accroche pour le placement : le centre


        try:
            #SQLT = threading.Thread(target = self.sql_event_date)                                   # Création du Thread pour obtenir la date du premier OF 
            #SQLT.start()                                                                            # Démarrage du Thread
            self.sql_event_date()
            self.calDebut.set_date(self.valDateDebut)                                               # Set de la date du calendrier
            self.calDebut.configure(mindate = self.valDateDebut)                                    # Date minimale acceptée par le calendrier
            self.calFin.configure(mindate = self.calDebut.get_date())                               # Date minimale acceptée par le calendrier

        except:
            print('Aucun calendrier')                                                               # Marqueur

        self.calFin.bind("<<DateEntrySelected>>", self.handlerCalFin)                               # Bind action selection date pour modification des min/max
        self.calDebut.bind("<<DateEntrySelected>>", self.handlerCalDebut)                           # Bind action selection date pour modification des min/max

    def handlerCalFin(self, eventObject):
        self.calDebut.configure(maxdate = self.calFin.get_date())                                   # Date maximale acceptée par le calendrier
        handlerRefT = threading.Thread(target = self.handlerRef)                            # On crée le Thread de l'import
        handlerRefT.start()                                                                 # On démarre le thread

    def handlerCalDebut(self, eventObject):
        self.calFin.configure(mindate = self.calDebut.get_date())                                   # Date minimale acceptée par le calendrier
        handlerRefT = threading.Thread(target = self.handlerRef)                            # On crée le Thread de l'import
        handlerRefT.start()                                                                 # On démarre le thread


    def select(self):                                                                               # Récupération des OF selectionner pour export de données

        self.list = []                                                                              # Création d'une liste pour les OF                          
        selection = self.tree.selection()                                                           # Récupération des OF sélectionner

        for i in selection:
            entry = self.tree.item(i)['values'][0:2]                                                # Récupération des données de l'OF et de la référence
            self.list.append(entry)                                                                 # Ajout des donnes dans la liste
        if(int(len(list(self.list))) > 0):                                                          # Si minimum un OF a été sélectionné
            self.importT = threading.Thread(target=self.sql_import)                                 # On crée le Thread de l'import
            self.importT.start()                                                                    # Démarrage du Thread
        
           

    def closing_extraction(self):
        self.controller.show_frame(self.startPage)                                                  # Retour à la page de démarrage

    def yview(self, *args):
        apply(self.tree.yview, args)                                                                # Obligatoire pour faire fonctionner le scrollbar

    def selectall(self):
        global stateAll, i

        self.list_child = []                                                                        # Création d'une liste pour les OF        

        if(stateAll == 1):                                                                          # Si bouton pas encore ON
            for child in self.tree.get_children():                                                  # On récupère tous les OF présents dans la liste
                self.list_child.insert(i, child)                                                    # On les insert dans la liste                                                    
            self.tree.selection_set(self.list_child)                                                # On sélectionne tous les OF
            stateAll = 0                                                                            # Flag 0 car bouton ON
        else:
            for child in self.tree.get_children():                                                  # On récupère tous les OF présents dans la liste
                self.list_child.insert(i, child)                                                    # On les insert dans la liste
            self.tree.selection_remove(self.list_child)                                             # On désélectionne tous les OF
            stateAll = 1                                                                            # Flag 0 car bouton OFF

    def sql_event_date(self):
        
        try:
            self.conn = SQL_Cursor.sql_connection_DB_QA()                                           # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.get_date()                                                                         # Obtention de la date du premirr OF

            print("Requête SQL QA GATE Date debut OF: OK")                                          # Marqueur                                                                    

            SQL_Cursor.sql_deconnection_DB_QA(self.cursor)                                          # Fermeture de la connection
        except:
            print("Error SQL QA GATE Date debut OF //!\\")

    def get_date(self):
        try:
            sql = """\
                            SELECT TOP(1) CAST(timestamp AS DATE) FROM QAGATE_1_MainTable ORDER BY timeStamp ASC
                        """
            self.cursor.execute(sql)
            self.valDateDebut = self.cursor.fetchval()                                              # Récupère toutes les données de la requête

        except:
            print("Aucune date //!\\")


    def get_value_of(self):
        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_OFDate] @2000 = ?, @2100 = ?, @3200 = ?, @3300 = ?, @Debut = ?, @Fin = ?
                    """
            self.cursor.execute(sql,self.varPGRef1.get(),                                           # Bit pour les 490035-2000
                                    self.varPGRef2.get(),                                           # Bit pour les 490035-2100 
                                    self.varPGRef3.get(),                                           # Bit pour les 490035-3200 
                                    self.varPGRef4.get(),                                           # Bit pour les 490035-3300
                                    self.calDebut.get_date(),
                                    self.calFin.get_date())
            
            valProcess = self.cursor.fetchall()
            self.tree.delete(*self.tree.get_children())                                             # Nettoie la liste des OF de l'affichage
            for row in valProcess:
                self.tree.insert("", "end", values=[row.currentOF, row.reference, row.dateDebut, row.dateFin])

        except:
            print("Aucune valeur OF //!\\")

    def sql_import(self):
        
        try:
            self.conn = SQL_Cursor.sql_connection_DB_QA()                                           # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.import_qa()                                                                        # On commence le traitement pour l'import

            print("Requête SQL QA GATE : OK")                                                       # Marqueur                                                                    

            SQL_Cursor.sql_deconnection_DB_QA(self.cursor)                                          # Fermeture de la connection
        except:
            print("Error SQL QA GATE Import //!\\")

    def import_qa(self):                                                                            # Import des données
        try:
            self.extractionButton['state']= tk.DISABLED                                             # Desactivation du bouton extraction pour éviter des problèmes

            j=2                                                                                     # Démarrage ligne 2
            k=2

            self.excel = win32.Dispatch('Excel.Application')                                        # Utilisation d'Excel
            self.wb = self.excel.Workbooks.Open(PATH_FOLDER_TEMPLATE)                               # Eciture dans :


            ws = self.wb.Worksheets("Données Production")                                            # Utilisation de la WorkSheet "Données"
            ws2 = self.wb.Worksheets("Données Evenements")

            self.progressbar.place(relx = 0.716,                                                    # Placement relatif par rapport à la taille de la fenêtre en x 
                                   rely = 0.92,                                                     # Placement relatif par rapport à la taille de la fenêtre en y 
                                   width = 200,                                                     # Largeur fixe 
                                   anchor = 'center')                                               # Point d'accroche pour le placement : le centre

            self.LabelBar.place(relx = 0.8,                                                         # Placement relatif par rapport à la taille de la fenêtre en x  
                                rely = 0.92,                                                        # Placement relatif par rapport à la taille de la fenêtre en y  
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre

            ofid = 1                                                                                # ID pour progress bar

            for val in self.list:

                ofNumb = int(len(list(self.list)))                                                  # Récupération du nombre de ligne pour progress bar
                
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ')'                    # Set du label progress bar 
                sql = """\
                            EXEC [dbo].[QAGATE_1_ExtractionQG] @OF = ?, @Reference = ?
                        """

                return_query = self.cursor.execute(sql, 
                                                   val[0],                                          # Valeur de l'OF
                                                   val[1])                                          # Valeur référence

                valProcess = self.cursor.fetchall()                                                 # Récupère toutes les données de la requête
                rowNumb = int(len(list(valProcess)))                                                # Compte le nombre de données pour l'OF
                progress=1                                                                          # Pour progress bar
                    
                for row in valProcess:

                    valProg = int((progress*100/rowNumb))

                    self.progressbar['value'] = valProg                                             # Update progress bar
                    self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                    ws.Range('A' + str(j)).Value = val[0]                                           # Numéro d'OF dans la colonne A
                    ws.Range('B' + str(j)).Value = row.reference                                    # Numéro de référence dans la colonne B
                    ws.Range('C' + str(j)).Value = row.dateJour                                     # Date du jour dans la colonne C
                    ws.Range('D' + str(j)).Value = row.run                                          # Temps de marche dans la colonne D
                    ws.Range('E' + str(j)).Value = row.arret                                        # Temps d'arrêt dans la colonne E
                    ws.Range('L' + str(j)).Value = row.pieceBonne                                   # Nombre pièces bonnes dans la colonne L
                    ws.Range('M' + str(j)).Value = row.kogame                                       # Nombre pièces mauvaises Kogame dans la colonne M
                    ws.Range('N' + str(j)).Value = row.keyence                                      # Nombre pièces mauvaises Keyence dans la colonne N

                    
                    progress += 1                                                                   # Pour progress bar
                    j += 1                                                                          # Pour passer de ligne en ligne dans l'excel


                self.progressbar['value'] = 0

                sql = """\
                            EXEC [dbo].[QAGATE_1_ExtractionQGEvent] @OF = ?
                        """

                return_query = self.cursor.execute(sql, 
                                                   val[0])                                          # Valeur de l'OF
                                                    

                valProcess = self.cursor.fetchall()                                                 # Récupère toutes les données de la requête
                rowNumb = int(len(list(valProcess)))                                                # Compte le nombre de données pour l'OF
                progress=1                                                                          # Pour progress bar
                    
                for row in valProcess:

                    valProg = int((progress*100/rowNumb))

                    self.progressbar['value'] = valProg                                             # Update progress bar
                    self.LabelBar['text'] = str(valProg) + '%'

                    ws2.Range('A' + str(k)).Value = row.idEvent                                     # Numéro d'OF dans la colonne A
                    ws2.Range('B' + str(k)).Value = row.reference                                   # Numéro de référence dans la colonne B
                    ws2.Range('C' + str(k)).Value = row.currentOF                                   # Date du jour dans la colonne C
                    ws2.Range('D' + str(k)).Value = row.mnemoniqueAlarme                            # Temps de marche dans la colonne D
                    ws2.Range('E' + str(k)).Value = row.etat                                        # Temps d'arrêt dans la colonne E
                    ws2.Range('F' + str(k)).Value = row.timeStamp                                   # Nombre pièces bonnes dans la colonne L


                    progress += 1                                                                   # Pour progress bar
                    k += 1                                                                          # Pour passer de ligne en ligne dans l'excel

                if(ofid == 1):
                    firstOF = val[0]
                if(ofid == ofNumb):
                    lastOF = val[0]
                ofid += 1                                                                           # Pour le label de la progress bar

            self.extractionButton['state']= tk.NORMAL                                               # Réactivation du bouton
            self.progressbar.place_forget()                                                         # Quand export fini, suppression progress bar
            self.progressbar['value'] = 0                                                           # Set à 0 progress bar pour prochain export
            self.LabelBar.place_forget()                                                            # Quand export fini, suppression label progress bar
            self.LabelBar['text'] = '(#/#)'                                                         # Set à (#/#) label progress bar pour prochain export
    
            now = datetime.datetime.now()
            dt_string = now.strftime("%Y%m%d_%H%M%S")

            if(firstOF == lastOF):
                self.excel.Visible = True                                                           # Affichage Excel
                self.wb.SaveAs(PATH_FOLDER_PROD + str(firstOF) + '_'+ dt_string + '.xlsm')          # Sauvegarde
            else:
                self.excel.Visible = True                                                           # Affichage Excel
                self.wb.SaveAs(PATH_FOLDER_PROD + str(firstOF) + '_' + str(lastOF) + '_'+ dt_string + '.xlsm')
                                                                                                    # Sauvegarde
            

        except:

            self.extractionButton['state']= tk.NORMAL                                               # Réactivation du bouton
            self.progressbar.place_forget()                                                         # Quand export fini, suppression progress bar
            self.progressbar['value'] = 0                                                           # Set à 0 progress bar pour prochain export
            self.LabelBar.place(relx = 0.8,                                                         # Placement relatif par rapport à la taille de la fenêtre en x  
                                rely = 0.92,                                                        # Placement relatif par rapport à la taille de la fenêtre en y  
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre
            self.LabelBar['text'] = 'ERROR'                                                         # Set à (#/#) label progress bar pour prochain export
            print("Aucun import QA GATE //!\\")

        

    def handlerPG(self, event = None):

        global stateSG

        if(self.varSG.get() == False):
            if(stateSG == 0):
                self.varSG.set(False)                                                               # Etat bouton SG à False
                self.ButtonSG['state'] = tk.DISABLED                                                # Désactivation du bouton SG
                self.ButtonPGRef1.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x     
                                        rely = 0.492,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe  
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonPGRef2.place(relx = 0.585,                                               # Placement relatif par rapport à la taille de la fenêtre en x     
                                        rely = 0.492,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe  
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonPGRef3.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x     
                                        rely = 0.561,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe  
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonPGRef4.place(relx = 0.585,                                               # Placement relatif par rapport à la taille de la fenêtre en x     
                                        rely = 0.561,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe  
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre
                stateSG = 1

            else:
                handlerRefT = threading.Thread(target = self.handlerRef)                            # On crée le Thread de l'import
                handlerRefT.start()                                                                 # On démarre le thread

                self.ButtonSG['state'] = tk.NORMAL                                                  # Activation du bouton SG                           

                self.ButtonPGRef1.place_forget()                                                    # Retire référence 1 
                self.varPGRef1.set(False)                                                           # Etat bouton référence 1 à False
                self.ButtonPGRef2.place_forget()                                                    # Retire référence 2 
                self.varPGRef2.set(False)                                                           # Etat bouton référence 2 à False
                self.ButtonPGRef3.place_forget()                                                    # Retire référence 3 
                self.varPGRef3.set(False)                                                           # Etat bouton référence 3 à False
                self.ButtonPGRef4.place_forget()                                                    # Retire référence 4 
                self.varPGRef4.set(False)                                                           # Etat bouton référence 4 à False

                stateSG=0


    def handlerSG(self, event = None):

        global statePG

        if(self.varPG.get() == False):
            if(statePG == 0):
                self.varPG.set(False)                                                               # Etat bouton SG à False
                self.ButtonPG['state'] = tk.DISABLED                                                # Désactivation du bouton SG
                self.ButtonSGRef1.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x  
                                        rely = 0.492,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
                                        height = 45,                                                # Hauteur fixe   
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonSGRef2.place(relx = 0.585,                                               # Placement relatif par rapport à la taille de la fenêtre en x  
                                        rely = 0.492,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe   
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonSGRef3.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x  
                                        rely = 0.561,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe   
                                        width = 166,                                                # Largeur fixe  
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre
                statePG = 1

            else:
                self.ButtonPG['state'] = tk.NORMAL                                                  # Activation du bouton PG 

                self.ButtonSGRef1.place_forget()                                                    # Retire référence 1
                self.varSGRef1.set(False)                                                           # Etat bouton référence 1 à False
                self.ButtonSGRef2.place_forget()                                                    # Retire référence 2
                self.varSGRef2.set(False)                                                           # Etat bouton référence 2 à False
                self.ButtonSGRef3.place_forget()                                                    # Retire référence 3
                self.varSGRef3.set(False)                                                           # Etat bouton référence 3 à False

                statePG=0

    def handlerRef(self):
        try:
            self.conn = SQL_Cursor.sql_connection_DB_QA()                                             # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.get_value_of()                                                                     # Obtention des valeurs des OF

            print("Requête SQL QA Gate Extraction OF : OK")                                         # Marqueur                                                                    

            SQL_Cursor.sql_deconnection_DB_QA(self.cursor)                                          # Fermeture de la connection
        except:
            print("Error SQL QA Gate Extraction OF //!\\")                                          # Marqueur

    def threadHandlerRef(self, event = None):
        handlerRefT = threading.Thread(target = self.handlerRef)
        handlerRefT.start()