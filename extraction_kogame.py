import sys
import datetime

import tkinter as tk
from tkinter import *
import tkinter.ttk as ttk

import pyodbc 

import SQL_Cursor

import win32com.client as win32
import threading

import pandas as pd

global PATH_FOLDER_TEMPLATE
PATH_FOLDER_TEMPLATE = '//SERV14/Public_new/IE/Public/Public 4.0/QA Gate 4.0/Rapport production QA Gate 4.0/Rapport_Template/Rapport_Kogame_template.xlsx'

global PATH_FOLDER_KOGAME
PATH_FOLDER_KOGAME = r'\\SERV14\Public_new\IE\Public\Public 4.0\QA Gate 4.0\Rapport production QA Gate 4.0\Rapport_Kogame\Rapport_Kogame_'

# Variable d'état
statePG = 0
stateSG = 1
stateAll = 1
i=0

class ExtractionKoPage(tk.Frame):

    def __init__(self, parent, controller, startPage):

        
        tk.Frame.__init__(self, parent)                                                             # Création de la frame contenant tous les éléments à affichier

        self.controller = controller                                                                # Controller de la fenêtre principale
        self.startPage = startPage                                                                  # Controller de la frame de démarrage 

        self.configure(background="#f4f4f4")                                                        # Configuration du fond en gris clair

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
                             #height = 85,                                                           # Hauteur fixe 
                             #width = 974,                                                           # Largeur fixe 
                             anchor = 'center')                                                     # Point d'accroche pour le placement : le centre du label

        self.LabelText.configure(background = '#f4f4f4')                                            # Fond en gris clair
        self.LabelText.configure(font = fontLabelMain)                                              # Police de caractère
        self.LabelText.configure(text = 'Extraction données Kogame')                                # Texte
        
        # Label reference
        self.LabelReference = tk.Label(self)
        self.LabelReference.place(relx = 0.5,                                                       # Placement relatif par rapport à la taille de la fenêtre en x 
                                  rely = 0.258,                                                     # Placement relatif par rapport à la taille de la fenêtre en y  
                                  #height = 51,                                                      # Hauteur fixe  
                                  #width = 234,                                                      # Largeur fixe  
                                  anchor = 'center')                                                # Point d'accroche pour le placement : le centre du label

        self.LabelReference.configure(background = "#f4f4f4")                                       # Fond en gris clair
        self.LabelReference.configure(font = fontLabelDescrip)                                      # Police de caractère
        self.LabelReference.configure(text = 'Reference')                                           # Texte

        # Label Type de pièce
        self.LabelType = tk.Label(self)
        self.LabelType.place(relx = 0.193,                                                          # Placement relatif par rapport à la taille de la fenêtre en x 
                             rely = 0.334,                                                          # Placement relatif par rapport à la taille de la fenêtre en y 
                             #height = 51,                                                           # Hauteur fixe 
                             #width = 224,                                                           # Largeur fixe 
                             anchor = 'w')                                                          # Point d'accroche pour le placement : l'ouest

        self.LabelType.configure(anchor = 'w')                                                      # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelType.configure(background = "#f4f4f4")                                            # Fond en gris clair
        self.LabelType.configure(font = fontLabelCat)                                               # Police de caractère
        self.LabelType.configure(text = 'Type de pièce :')                                          # Texte

        # Label Réference pièce
        self.LabelRef = tk.Label(self)
        self.LabelRef.place(relx = 0.193,                                                           # Placement relatif par rapport à la taille de la fenêtre en x  
                            rely = 0.442,                                                           # Placement relatif par rapport à la taille de la fenêtre en y  
                            #height = 51,                                                            # Hauteur fixe  
                            #width = 254,                                                            # Largeur fixe 
                            anchor = 'w')                                                           # Point d'accroche pour le placement : l'ouest 
        
        self.LabelRef.configure(anchor='w')                                                         # Point d'accroche pour le texte dans le label : l'ouest
        self.LabelRef.configure(background="#f4f4f4")                                               # Fond en gris clair
        self.LabelRef.configure(font=fontLabelCat)                                                  # Police de caractère
        self.LabelRef.configure(text='''Référence pièce :''')                                       # Texte

        # Label principale
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


        # Radio Bouton (2 position)
        
        # Bouton PG 
        self.varPG = tk.BooleanVar()                                                                # Variable du bouton PG
        self.varPG.set(True)
        self.ButtonPG = tk.Checkbutton(self, 
                                       variable = self.varPG,                                       # Lien avec la variable
                                       indicatoron = 0,                                             # Forme du bouton (Bouton clicable classique)
                                       command = self.handlerPG)                                    # Action lors d'un clic
        self.ButtonPG.place(relx = 0.413,                                                           # Placement relatif par rapport à la taille de la fenêtre en x 
                            rely = 0.334,                                                           # Placement relatif par rapport à la taille de la fenêtre en y 
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
                            rely = 0.334,                                                           # Placement relatif par rapport à la taille de la fenêtre en y 
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
                                rely = 0.442,                                                       # Placement relatif par rapport à la taille de la fenêtre en y 
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
                                rely = 0.442,                                                       # Placement relatif par rapport à la taille de la fenêtre en y 
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
                                rely = 0.511,                                                       # Placement relatif par rapport à la taille de la fenêtre en y  
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
                                rely = 0.511,                                                       # Placement relatif par rapport à la taille de la fenêtre en y 
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
        self.LabelBar['text'] = '(#/#) ###%'                                                        # Texte

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


    def sql_import(self):
        
        try:
            self.conn = SQL_Cursor.sql_connection_DB_Measurlink()                                   # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.import_kogame()                                                                    # On commence le traitement pour l'import

            print("Requête SQL Measurlink : OK")                                                    # Marqueur                                                                    

            SQL_Cursor.sql_deconnection_DB_Measurlink(self.cursor)                                  # Fermeture de la connection

        except:
            print("Error SQL Measurlink Import //!\\")


    def get_value_of(self):                                                                         # Récupère les OF de Measurlink
        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_Value_OF] @2000 = ?, @2100 = ?, @3200 = ?, @3300 = ?
                    """
            
            self.cursor.execute(sql, 
                                self.varPGRef1.get(),                                               # Bit pour les 490035-2000
                                self.varPGRef2.get(),                                               # Bit pour les 490035-2100 
                                self.varPGRef3.get(),                                               # Bit pour les 490035-3200 
                                self.varPGRef4.get())                                               # Bit pour les 490035-3300

            valProcess = self.cursor.fetchall()                                                     # Récupère toutes les données de la requête
            self.tree.delete(*self.tree.get_children())                                             # Nettoie la liste des OF de l'affichage
            for row in valProcess:
                self.tree.insert("", "end", values=[row.currentOF, row.reference, '1', '2'])        # Insert les données reçu


        except:
            print("Aucune valeur OF //!\\")                                                         # Marqueur

    def import_kogame(self):                                                                        # Import des données
        try:
            self.extractionButton['state']= tk.DISABLED                                             # Desactivation du bouton extraction pour éviter des problèmes
            j=2                                                                                     # Démarrage ligne 2
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')                         # Utilisation d'Excel
            self.wb = self.excel.Workbooks.Open(PATH_FOLDER_TEMPLATE)                               # Ecriture dans :

            ws = self.wb.Worksheets("Données")                                                      # Utilisation de la WorkSheet "Données"

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

                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') '                   # Set du label progress bar 
                
                sql = """\
                            EXEC [dbo].[QAGATE_1_Export_Data] @OF = ?, @Reference = ?
                        """

                return_query = self.cursor.execute(sql, 
                                                   val[0],                                          # Valeur de l'OF
                                                   val[1])                                          # Valeur de la référence

                valProcess = self.cursor.fetchall()                                                 # Récupère toutes les données de la requête

                final_result = [list(i) for i in valProcess]
                final_result = pd.DataFrame(final_result, columns = ['id', 
                                                                     'Hauteur', 
                                                                     'HauteurTolUp', 
                                                                     'HauteurTolLow', 
                                                                     'ParallelismeFace',
                                                                     'ParallelismeFaceTolUp', 
                                                                     'ParallelismeFaceTolLow', 
                                                                     'PlaneiteFace',
                                                                     'PlaneiteFaceTolUp',
                                                                     'PlaneiteFaceTolLow',
                                                                     'PlaneiteFaceHexag',
                                                                     'PlaneiteFaceHexagTolUp',
                                                                     'PlaneiteFaceHexagTolLow',
                                                                     'RectitudeFace1Hexag',
                                                                     'RectitudeFace1HexagTolUp', 
                                                                     'RectitudeFace1HexagTolLow',
                                                                     'RectitudeFace2',
                                                                     'RectitudeFace2TolUp',
                                                                     'RectitudeFace2TolLow'])

                progress=0                                                                          # Pour progress bar
                i = len(final_result) - 1

                valProg = int((progress*100/21))

                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'
                                                                                                    # Set du label progress bar
                    
                ws.Range('A' + str(j) + ':' + 'A' + str(int(j+i))).Value = val[0]                   # Numéro d'OF dans la colonne A
                progress += 1
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('B' + str(j) + ':' + 'B' + str(int(j+i))).Value = val[1]                   # Numéro de référence dans la colonne B
                progress += 1
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('C' + str(j) + ':' + 'C' + str(int(j+i))).Value = list(zip(final_result['id']))
                progress += 1                                                                       # Id dans la colonne C
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('D' + str(j) + ':' + 'D' + str(int(j+i))).Value = list(zip(final_result['Hauteur']))
                progress += 1                                                                       # Données Hauteur dans la colonne D
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('E' + str(j) + ':' + 'E' + str(int(j+i))).Value = list(zip(final_result['HauteurTolUp']))
                progress += 1                                                                       # Données Hauteur dans la colonne D
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('F' + str(j) + ':' + 'F' + str(int(j+i))).Value = list(zip(final_result['HauteurTolLow']))
                progress += 1                                                                       # Données Hauteur dans la colonne D
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('G' + str(j) + ':' + 'G' + str(int(j+i))).Value = list(zip(final_result['ParallelismeFace']))                             
                progress += 1                                                                       # Données Parallelisme Face dans la colonne G
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('H' + str(j) + ':' + 'H' + str(int(j+i))).Value = list(zip(final_result['ParallelismeFaceTolUp']))                             
                progress += 1                                                                       # Données Parallelisme Face dans la colonne G
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('I' + str(j) + ':' + 'I' + str(int(j+i))).Value = list(zip(final_result['ParallelismeFaceTolLow']))                             
                progress += 1                                                                       # Données Parallelisme Face dans la colonne G
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('J' + str(j) + ':' + 'J' + str(int(j+i))).Value = list(zip(final_result['PlaneiteFace']))                                 
                progress += 1                                                                       # Données Planeite Face dans la colonne J
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('K' + str(j) + ':' + 'K' + str(int(j+i))).Value = list(zip(final_result['PlaneiteFaceTolUp']))                                 
                progress += 1                                                                       # Données Planeite Face dans la colonne J
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('L' + str(j) + ':' + 'L' + str(int(j+i))).Value = list(zip(final_result['PlaneiteFaceTolLow']))                                 
                progress += 1                                                                       # Données Planeite Face dans la colonne J
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('M' + str(j) + ':' + 'M' + str(int(j+i))).Value = list(zip(final_result['PlaneiteFaceHexag']))                          
                progress += 1                                                                       # Données Planeite Face Hexag dans la colonne M
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('N' + str(j) + ':' + 'N' + str(int(j+i))).Value = list(zip(final_result['PlaneiteFaceHexagTolUp']))                          
                progress += 1                                                                       # Données Planeite Face Hexag dans la colonne M
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('O' + str(j) + ':' + 'O' + str(int(j+i))).Value = list(zip(final_result['PlaneiteFaceHexagTolLow']))                          
                progress += 1                                                                       # Données Planeite Face Hexag dans la colonne M
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('P' + str(j) + ':' + 'P' + str(int(j+i))).Value = list(zip(final_result['RectitudeFace1Hexag']))                          
                progress += 1                                                                       # Données Rectitude Face 1 Hexag dans la colonne P
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('Q' + str(j) + ':' + 'Q' + str(int(j+i))).Value = list(zip(final_result['RectitudeFace1HexagTolUp']))                          
                progress += 1                                                                       # Données Rectitude Face 1 Hexag dans la colonne P
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('R' + str(j) + ':' + 'R' + str(int(j+i))).Value = list(zip(final_result['RectitudeFace1HexagTolLow']))                          
                progress += 1                                                                       # Données Rectitude Face 1 Hexag dans la colonne P
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('S' + str(j) + ':' + 'S' + str(int(j+i))).Value = list(zip(final_result['RectitudeFace2']))                               
                progress += 1                                                                       # Données Rectitude Face 2 dans la colonne S
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('T' + str(j) + ':' + 'T' + str(int(j+i))).Value = list(zip(final_result['RectitudeFace2TolUp']))                               
                progress += 1                                                                       # Données Rectitude Face 2 dans la colonne S
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'

                ws.Range('U' + str(j) + ':' + 'U' + str(int(j+i))).Value = list(zip(final_result['RectitudeFace2TolLow']))                               
                progress += 1                                                                       # Données Rectitude Face 2 dans la colonne S
                valProg = int((progress*100/21))
                self.progressbar['value'] = valProg                                                 # Update progress bar
                self.LabelBar['text'] = '(' + str(ofid)+ '/' + str(ofNumb) + ') ' + str(valProg) + '%'
 
                j += i+1                                                                            # Pour passer de ligne en ligne dans l'excel
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
                self.wb.SaveAs(PATH_FOLDER_KOGAME + str(firstOF) + '_'+ dt_string + '.xlsx')                         # Sauvegarde
            else:
                self.excel.Visible = True                                                           # Affichage Excel
                self.wb.SaveAs(PATH_FOLDER_KOGAME + str(firstOF) + '_' + str(lastOF) + '_'+ dt_string + '.xlsx')     # Sauvegarde

        except:
            self.extractionButton['state']= tk.NORMAL                                               # Réactivation du bouton
            self.progressbar.place_forget()                                                         # Quand export fini, suppression progress bar
            self.progressbar['value'] = 0                                                           # Set à 0 progress bar pour prochain export
            self.LabelBar.place(relx = 0.8,                                                         # Placement relatif par rapport à la taille de la fenêtre en x  
                                rely = 0.92,                                                        # Placement relatif par rapport à la taille de la fenêtre en y  
                                anchor = 'center')                                                  # Point d'accroche pour le placement : le centre
            self.LabelBar['text'] = 'ERROR'                                                         # Set à ERROR label progress bar pour prochain export
            print("Aucun import Kogame //!\\")


    def handlerPG(self, event = None):

        global stateSG

        if(self.varSG.get() == False):
            if(stateSG == 0):
                self.varSG.set(False)                                                               # Etat bouton SG à False
                self.ButtonSG['state'] = tk.DISABLED                                                # Désactivation du bouton SG
                self.ButtonPGRef1.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x    
                                        rely = 0.442,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
                                        height = 45,                                                # Hauteur fixe 
                                        width = 166,                                                # Largeur fixe 
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre
                
                self.ButtonPGRef2.place(relx = 0.585,                                               # Placement relatif par rapport à la taille de la fenêtre en x    
                                        rely = 0.442,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
                                        height = 45,                                                # Hauteur fixe 
                                        width = 166,                                                # Largeur fixe 
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre
                
                self.ButtonPGRef3.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x   
                                        rely = 0.511,                                               # Placement relatif par rapport à la taille de la fenêtre en y  
                                        height = 45,                                                # Hauteur fixe 
                                        width = 166,                                                # Largeur fixe
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre
                
                self.ButtonPGRef4.place(relx = 0.585,                                               # Placement relatif par rapport à la taille de la fenêtre en x    
                                        rely = 0.511,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
                                        height = 45,                                                # Hauteur fixe 
                                        width = 166,                                                # Largeur fixe 
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre
                stateSG = 1

            else:
                self.handlerRefT = threading.Thread(target = self.handlerRef)                       # On crée le Thread de l'import
                self.handlerRefT.start()                                                            # On démarre le thread

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
                                        rely = 0.442,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
                                        height = 45,                                                # Hauteur fixe  
                                        width = 166,                                                # Largeur fixe 
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonSGRef2.place(relx = 0.585,                                               # Placement relatif par rapport à la taille de la fenêtre en x  
                                        rely = 0.442,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
                                        height = 45,                                                # Hauteur fixe  
                                        width = 166,                                                # Largeur fixe 
                                        anchor = 'center')                                          # Point d'accroche pour le placement : le centre

                self.ButtonSGRef3.place(relx = 0.413,                                               # Placement relatif par rapport à la taille de la fenêtre en x  
                                        rely = 0.511,                                               # Placement relatif par rapport à la taille de la fenêtre en y 
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
            self.conn = SQL_Cursor.sql_connection_DB_Measurlink()                                   # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.get_value_of()                                                                     # Obtention des OF en fonction des conditions

            print("Requête SQL Measurlink Extraction OF : OK")                                      # Marqueur                                                                    

            SQL_Cursor.sql_deconnection_DB_Measurlink(self.cursor)                                  # Fermeture de la connection
        except:
            print("Error SQL Measurlink Extraction OF //!\\")                                       # Marqueur
        
        if((self.varPGRef1.get() == 1 or self.varPGRef2.get() == 1) and self.varPG.get() == 1 ):
            self.ButtonPGRef3.place_forget()                                                        # Retire référence 3
            self.varPGRef3.set(False)
            self.ButtonPGRef4.place_forget()                                                        # Retire référence 4
            self.varPGRef4.set(False)
        elif((self.varPGRef3.get() == 1 or self.varPGRef4.get() == 1) and self.varPG.get() == 1):
            self.ButtonPGRef1.place_forget()                                                        # Retire référence 3
            self.varPGRef1.set(False)
            self.ButtonPGRef2.place_forget()                                                        # Retire référence 4
            self.varPGRef2.set(False)

        elif((self.varPGRef1.get() == 0 and self.varPGRef2.get() == 0) and self.varPG.get() == 1 and self.ButtonPGRef3.winfo_ismapped() == 0):
            self.ButtonPGRef3.place(relx = 0.413,                                                   # Placement relatif par rapport à la taille de la fenêtre en x   
                                    rely = 0.511,                                                   # Placement relatif par rapport à la taille de la fenêtre en y  
                                    height = 45,                                                    # Hauteur fixe 
                                    width = 166,                                                    # Largeur fixe
                                    anchor = 'center')                                              # Point d'accroche pour le placement : le centre

            self.ButtonPGRef4.place(relx = 0.585,                                                   # Placement relatif par rapport à la taille de la fenêtre en x    
                                    rely = 0.511,                                                   # Placement relatif par rapport à la taille de la fenêtre en y 
                                    height = 45,                                                    # Hauteur fixe 
                                    width = 166,                                                    # Largeur fixe
                                    anchor = 'center')                                              # Point d'accroche pour le placement : le centre

        elif((self.varPGRef3.get() == 0 and self.varPGRef4.get() == 0) and self.varPG.get() == 1 and self.ButtonPGRef1.winfo_ismapped() == 0):

            self.ButtonPGRef1.place(relx = 0.413,                                                   # Placement relatif par rapport à la taille de la fenêtre en x 
                                    rely = 0.442,                                                   # Placement relatif par rapport à la taille de la fenêtre en y 
                                    height = 45,                                                    # Hauteur fixe  
                                    width = 166,                                                    # Largeur fixe
                                    anchor = 'center')                                              # Point d'accroche pour le placement : le centre

            self.ButtonPGRef2.place(relx = 0.585,                                                   # Placement relatif par rapport à la taille de la fenêtre en x  
                                    rely = 0.442,                                                   # Placement relatif par rapport à la taille de la fenêtre en y 
                                    height = 45,                                                    # Hauteur fixe 
                                    width = 166,                                                    # Largeur fixe
                                    anchor = 'center')                                              # Point d'accroche pour le placement : le centre
    
    def threadHandlerRef(self, event = None):
        handlerRefT = threading.Thread(target = self.handlerRef)
        handlerRefT.start()