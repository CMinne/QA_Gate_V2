## Module Tkinter ##
import tkinter as tk
from tkinter import *

## Module Interface ##
import syntax
import SQL_Cursor

## Module Basique Python ##
from datetime import datetime
import decimal
import threading

## Module Matplotlib ##
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.ticker as plticker
import matplotlib.dates as mdates

## Module Pandas ##
import pandas as pd
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()

## Module SQL SERVER ##
import pyodbc

## Module Numpy ##
import numpy as np



class ProcessusJourPage(tk.Frame):

    def __init__(self, parent, controller, startPage, processusPageOF, frames):

        
        tk.Frame.__init__(self, parent)

        # Controleur principal 
        self.controller = controller

        # Mappage des controlleurs de toutes les frames
        self.startPage = startPage
        self.processusPageOF = processusPageOF

        # Mappage de la bibliothèque de frame
        self.frames = frames

        self.configure(background="#f4f4f4")                                                        # Configuration du fond en gris clair

        # Paramétrage des polices
        self.CS_Label_Title = font.Font(self,                                                       # Font d'une major partie des titres
                                        family = '{Arial}', 
                                        size = 36)
        self.CS_Label_OF = font.Font(self,                                                          # Font du label OF
                                     family = '{Arial}', 
                                     size = 48)
        self.CS_Label_cpk = font.Font(self,                                                         # Font des cpk 
                                      family = '{Arial}', 
                                      size = 22)
        self.CS_Label_Value = font.Font(self,                                                       # Font du reste 
                                        family = '{Arial}', 
                                        size = 28)
        self.CS_Label_Code = font.Font(self,                                                        # Font du code évènement 
                                       family = '{Arial}', 
                                       size = 18)

        # Création Frame et Label

        # Création frame début OF
        self.Frame_Debut_OF = tk.Frame(self)
        self.Frame_Debut_OF.place(relx = 0.0, 
                                  rely = 0.0, 
                                  relheight = 0.06, 
                                  relwidth = 0.333)
        self.Frame_Debut_OF.configure(background = "#f4f4f4")

        # Création label début OF
        self.Label_Debut_OF = tk.Label(self.Frame_Debut_OF)
        self.Label_Debut_OF.place(relx = 0.0, 
                                  rely = 0.0, 
                                  relheight = 1, 
                                  relwidth = 1)
        self.Label_Debut_OF.configure(background = "#f4f4f4")
        self.Label_Debut_OF.configure(font = self.CS_Label_Title)
        self.Label_Debut_OF.configure(text = 'Debut OF : ##/##/##')                                 # Titre par défaut

        # Création frame fin OF
        self.Frame_Fin_OF = tk.Frame(self)
        self.Frame_Fin_OF.place(relx = 0.0, 
                                  rely = 0.055, 
                                  relheight = 0.084, 
                                  relwidth = 0.333)
        self.Frame_Fin_OF.configure(background = "#f4f4f4")

        # Création label fin OF
        self.Label_Fin_OF = tk.Label(self.Frame_Fin_OF)
        self.Label_Fin_OF.place(relx = 0.0, 
                                  rely = 0.0, 
                                  relheight = 1, 
                                  relwidth = 1)
        self.Label_Fin_OF.configure(background = "#f4f4f4")
        self.Label_Fin_OF.configure(font = self.CS_Label_Title)
        self.Label_Fin_OF.configure(text = 'Fin OF : ##/##/##')                                 # Titre par défaut


        # Création frame OF
        self.Frame_OF = tk.Frame(self)
        self.Frame_OF.place(relx = 0.333, 
                            rely = 0.0, 
                            relheight = 0.075, 
                            relwidth = 0.333)
        self.Frame_OF.configure(background = "#f4f4f4")

        # Création label OF
        self.Label_OF = tk.Label(self.Frame_OF)
        self.Label_OF.place(relx = 0.0, 
                            rely = 0.0, 
                            relheight = 1, 
                            relwidth = 1)
        self.Label_OF.configure(background = "#f4f4f4")
        self.Label_OF.configure(font = self.CS_Label_OF)
        self.Label_OF.configure(text = 'OF : ######')                                               # Titre par défaut

        # Création frame reference
        self.Frame_Ref = tk.Frame(self)
        self.Frame_Ref.place(relx = 0.333, 
                            rely = 0.075, 
                            relheight = 0.075, 
                            relwidth = 0.333)
        self.Frame_Ref.configure(background = "#f4f4f4")

        # Création label reference
        self.Label_Ref = tk.Label(self.Frame_Ref)
        self.Label_Ref.place(relx = 0.0, 
                            rely = 0.0, 
                            relheight = 1, 
                            relwidth = 1)
        self.Label_Ref.configure(background = "#f4f4f4")
        self.Label_Ref.configure(font = self.CS_Label_Title)
        self.Label_Ref.configure(text = '######-####')                                              # Titre par défaut


        # Création frame avancement OF
        self.Frame_Avancement = tk.Frame(self)
        self.Frame_Avancement.place(relx = 0.667, 
                                    rely = 0.0, 
                                    relheight = 0.084, 
                                    relwidth = 0.333)
        self.Frame_Avancement.configure(background = "#f4f4f4")

        # Création label avancement OF
        self.Label_Avancement = tk.Label(self.Frame_Avancement)
        self.Label_Avancement.place(relx = 0.0, 
                                    rely = 0.0, 
                                    relheight = 1, 
                                    relwidth = 1)
        self.Label_Avancement.configure(background = "#f4f4f4")
        self.Label_Avancement.configure(font = self.CS_Label_Title)
        self.Label_Avancement.configure(text = 'Avancement : ##%')                                  # Titre par défaut


        # Création frame Bekido
        self.Frame_Bekido = tk.Frame(self)
        self.Frame_Bekido.place(relx = 0.155, 
                                rely = 0.199, 
                                relheight = 0.064, 
                                relwidth = 0.195)
        self.Frame_Bekido.configure(background = "#f4f4f4")

        # Création label Bekido
        self.Label_Bekido = tk.Label(self.Frame_Bekido)
        self.Label_Bekido.place(relx = 0.0, 
                                rely = 0.0, 
                                relheight = 1, 
                                relwidth = 1)
        self.Label_Bekido.configure(anchor = 'w')                                                   # Position de l'écriture à gauche du label
        self.Label_Bekido.configure(background = "#f4f4f4")
        self.Label_Bekido.configure(font = self.CS_Label_Title)
        self.Label_Bekido.configure(foreground = "#ffffff")                                         # Police en blanc (car fond rouge ou vert)
        self.Label_Bekido.configure(text = 'Bekido : >###%')                                        # Titre par défaut


        # Création frame Prévision pièces contrôlées
        self.Frame_Prevision = tk.Frame(self)
        self.Frame_Prevision.place(relx = 0.155, 
                                   rely = 0.268, 
                                   relheight = 0.064, 
                                   relwidth = 0.195)
        self.Frame_Prevision.configure(background = "#f4f4f4")

        # Création label Prévision pièces contrôlées
        self.Label_Prevision = tk.Label(self.Frame_Prevision)
        self.Label_Prevision.place(relx = 0.0, 
                                   rely = 0.0, 
                                   relheight = 1, 
                                   relwidth = 1)
        self.Label_Prevision.configure(anchor = 'w')                                                # Position de l'écriture à gauche du label
        self.Label_Prevision.configure(background = "#f4f4f4")
        self.Label_Prevision.configure(font = self.CS_Label_Title)
        self.Label_Prevision.configure(text = 'Prévision : ####')                                   # Titre par défaut

        # Création frame pièces contrôlées actuelles
        self.Frame_Actuel = tk.Frame(self)
        self.Frame_Actuel.place(relx = 0.155, 
                                rely = 0.337, 
                                relheight = 0.064, 
                                relwidth = 0.195)
        self.Frame_Actuel.configure(background = "#f4f4f4")

        # Création label pièces contrôlées actuelles
        self.Label_Actuel = tk.Label(self.Frame_Actuel)
        self.Label_Actuel.place(relx = 0.0, 
                                rely = 0.0, 
                                relheight = 1,
                                relwidth = 1)
        self.Label_Actuel.configure(anchor = 'w')                                                   # Position de l'écriture à gauche du label
        self.Label_Actuel.configure(background = "#f4f4f4")
        self.Label_Actuel.configure(font = self.CS_Label_Title)
        self.Label_Actuel.configure(text = 'Actuel : ####')                                         # Titre par défaut

        # Création frame Delta pièces contrôlées
        self.Frame_Delta = tk.Frame(self)
        self.Frame_Delta.place(relx = 0.155, 
                               rely = 0.406, 
                               relheight = 0.064, 
                               relwidth = 0.195)
        self.Frame_Delta.configure(background = "#f4f4f4")

        # Création label Delta pièces contrôlées
        self.Label_Delta = tk.Label(self.Frame_Delta)
        self.Label_Delta.place(relx = 0.0, 
                               rely = 0.0, 
                               relheight = 1, 
                               relwidth = 1)
        self.Label_Delta.configure(anchor = 'w')                                                    # Position de l'écriture à gauche du label
        self.Label_Delta.configure(background = "#f4f4f4")
        self.Label_Delta.configure(font = self.CS_Label_Title)
        self.Label_Delta.configure(foreground = "#ffffff")                                          # Police en blanc (car fond rouge ou vert)
        self.Label_Delta.configure(text = 'Delta : ####')                                           # Titre par défaut

        # Création frame Graphique évènements
        self.Frame_Chart_Event = tk.Frame(self)
        self.Frame_Chart_Event.place(relx = 0.052, 
                                     rely = 0.584, 
                                     relheight = 0.342, 
                                     relwidth = 0.393)
        self.Frame_Chart_Event.configure(background = "#f4f4f4")

        # Paramétrage Graphique évènements
        self.figure_Event = Figure(figsize = (8,4), dpi = 100)                                      # Création d'une figure de 8 et 4 inch de largeur hauteur avec 100 points par inch
        self.subplot_Event = self.figure_Event.add_subplot()                                        # On rajoute une parcelle de dessin 
        self.step_Event = FigureCanvasTkAgg(self.figure_Event, self.Frame_Chart_Event)              # On lie Matplotlib à Tkinter
        self.step_Event.get_tk_widget().pack()                                                      # On met la figure dans la frame Graphique évènement
        self.figure_Event.set_facecolor('#f4f4f4')                                                  # On set son fond en gris clair

        # Création frame Titre graphique évènements
        self.Frame_Title_Chart_Event = tk.Frame(self)
        self.Frame_Title_Chart_Event.place(relx = 0.052, 
                                           rely = 0.555, 
                                           relheight = 0.054, 
                                           relwidth = 0.393)
        self.Frame_Title_Chart_Event.configure(background = "#f4f4f4")
        
        # Création label Titre graphique évènements
        self.Label_Title_Chart_Event = tk.Label(self.Frame_Title_Chart_Event)
        self.Label_Title_Chart_Event.place(relx = 0.0, 
                                           rely = 0.0, 
                                           relheight = 1, 
                                           relwidth = 1)
        self.Label_Title_Chart_Event.configure(background = "#f4f4f4")
        self.Label_Title_Chart_Event.configure(font = self.CS_Label_Title)
        self.Label_Title_Chart_Event.configure(text = 'Timeline évènement')                         # Titre


        # Création frame total des évènements
        self.Frame_Total_Event = tk.Frame(self)
        self.Frame_Total_Event.place(relx = 0.0, 
                                     rely = 0.72, 
                                     relheight = 0.09, 
                                     relwidth = 0.06)
        self.Frame_Total_Event.configure(background = "#f4f4f4")

        # Création label total des évènements
        self.Label_Total_Event = tk.Label(self.Frame_Total_Event)
        self.Label_Total_Event.place(relx = 0.0, 
                                     rely = 0.0, 
                                     relheight = 1, 
                                     relwidth = 1)
        self.Label_Total_Event.configure(background = "#f4f4f4")
        self.Label_Total_Event.configure(font = self.CS_Label_Value)
        self.Label_Total_Event.configure(text = 'Total\n####')                                     # Titre par défaut


        # Création frame total des évènements
        self.Frame_Code_Event = tk.Frame(self)
        self.Frame_Code_Event.place(relx = 0.08, 
                                     rely = 0.931, 
                                     relheight = 0.045, 
                                     relwidth = 0.34)
        self.Frame_Code_Event.configure(background = "#f4f4f4")

        # Création label total des évènements
        self.Label_Code_Event = tk.Label(self.Frame_Code_Event)
        self.Label_Code_Event.place(relx = 0.0, 
                                     rely = 0.0, 
                                     relheight = 1, 
                                     relwidth = 1)
        self.Label_Code_Event.configure(background = "#f4f4f4")
        self.Label_Code_Event.configure(anchor = 'w')
        self.Label_Code_Event.configure(font = self.CS_Label_Code)
        self.Label_Code_Event.configure(text = 'Code : #')                                          # Titre par défaut


        # Création frame Graphique Keyence
        self.Frame_Pie_Keyence = tk.Frame(self)
        self.Frame_Pie_Keyence.place(relx = 0.48, 
                                     rely = 0.257, 
                                     relheight = 0.328, 
                                     relwidth = 0.25)
        self.Frame_Pie_Keyence.configure(background = "#f4f4f4")
        
        # Paramétrage Graphique Keyence
        self.figure_Keyence = Figure(figsize = (5,4), dpi = 100)                                    # Création d'une figure de 5 et 4 inch de largeur hauteur avec 100 points par inch
        self.subplot_Keyence = self.figure_Keyence.add_subplot()                                    # On rajoute une parcelle de dessin 
        self.subplot_Keyence.axis('equal')                                                          # Permet de normaliser les axes
        self.subplot_Keyence.set_title('KEYENCE')
        self.pie_Keyence = FigureCanvasTkAgg(self.figure_Keyence, self.Frame_Pie_Keyence)           # On lie Matplotlib à Tkinter
        self.pie_Keyence.get_tk_widget().pack()                                                     # On met la figure dans la frame Graphique Keyence
        self.figure_Keyence.set_facecolor('#f4f4f4')                                                # On set son fond en gris clair

        self.subplot_Keyence.clear()
        labels = 'NO DATA','Autre'                                                                  # Label erreur
        patches, texts, autotexts = self.subplot_Keyence.pie([1,0], 
                                                                labels = labels, 
                                                                colors = ['red'], 
                                                                autopct = '%1.1f%%',                # Affichage en pourcent
                                                                shadow = True, 
                                                                startangle = 90)
        self.figure_Keyence.canvas.draw_idle()                                                      # Redessine le tracé


        # Création frame Graphique Kogame
        self.Frame_Pie_Kogame = tk.Frame(self)
        self.Frame_Pie_Kogame.place(relx = 0.48, 
                                    rely = 0.614, 
                                    relheight = 0.328, 
                                    relwidth = 0.25)
        self.Frame_Pie_Kogame.configure(background = "#f4f4f4")

        # Paramétrage Graphique Kogame
        self.figure_Kogame = Figure(figsize = (5,4), dpi = 100)                                     # Création d'une figure de 5 et 4 inch de largeur hauteur avec 100 points par inch
        self.subplot_Kogame = self.figure_Kogame.add_subplot()                                      # On rajoute une parcelle de dessin  
        self.subplot_Kogame.axis('equal')                                                           # Permet de normaliser les axes
        self.subplot_Kogame.set_title('KOGAME')
        self.pie_Kogame = FigureCanvasTkAgg(self.figure_Kogame, self.Frame_Pie_Kogame)              # On lie Matplotlib à Tkinter
        self.pie_Kogame.get_tk_widget().pack()                                                      # On met la figure dans la frame Graphique Keyence
        self.figure_Kogame.set_facecolor('#f4f4f4')                                                 # On set son fond en gris clair

        self.subplot_Kogame.clear()

        labels = 'NO DATA','Autre'                                                                  # Label par défaut
        patches, texts, autotexts = self.subplot_Kogame.pie([1,0], 
                                                            labels = labels, 
                                                            colors = ['red'], 
                                                            autopct = '%1.1f%%',                    # Affichage en pourcent
                                                            shadow = True, 
                                                            startangle = 90)
        self.figure_Kogame.canvas.draw_idle()


        # Création frame Chokko
        self.Frame_Chokko = tk.Frame(self)
        self.Frame_Chokko.place(relx = 0.53, 
                                rely = 0.199, 
                                relheight = 0.064, 
                                relwidth = 0.18)
        self.Frame_Chokko.configure(background = "#f4f4f4")

        # Création label Chokko
        self.Label_Chokko = tk.Label(self.Frame_Chokko)
        self.Label_Chokko.place(relx = 0.0, 
                                rely = 0.0, 
                                relheigh = 1, 
                                relwidth = 1)
        self.Label_Chokko.configure(anchor = 'w')                                                   # Position de l'écriture à gauche du label
        self.Label_Chokko.configure(background = "#f4f4f4")
        self.Label_Chokko.configure(font = self.CS_Label_Title)
        self.Label_Chokko.configure(foreground = "#ffffff")                                         # Police en blanc (car fond rouge ou vert)
        self.Label_Chokko.configure(text = 'Chokko : ###%')                                         # Titre par défaut


        # Création frame Rebut total
        self.Frame_Rebut = tk.Frame(self)
        self.Frame_Rebut.place(relx = 0.766, 
                               rely = 0.199, 
                               relheight = 0.064, 
                               relwidth = 0.211)
        self.Frame_Rebut.configure(background = "#f4f4f4")

        # Création label Rebut total
        self.Label_Rebut = tk.Label(self.Frame_Rebut)
        self.Label_Rebut.place(relx = 0.0, 
                               rely = 0.0, 
                               relheight = 1, 
                               relwidth = 1)
        self.Label_Rebut.configure(anchor = 'w')                                                    # Position de l'écriture à gauche du label
        self.Label_Rebut.configure(background = "#f4f4f4")
        self.Label_Rebut.configure(font = self.CS_Label_Title)
        self.Label_Rebut.configure(text = 'Rebut total : ####')                                     # Titre par défaut


        # Création frame Rebut Keyence
        self.Frame_Keyence = tk.Frame(self)
        self.Frame_Keyence.place(relx = 0.75, 
                                 rely = 0.38, 
                                 relheight = 0.095, 
                                 relwidth = 0.216)
        self.Frame_Keyence.configure(background = "#f4f4f4")

        # Création label Rebut Keyence
        self.Label_Keyence = tk.Label(self.Frame_Keyence)
        self.Label_Keyence.place(relx = 0.0, 
                                 rely = 0.0, 
                                 relheight = 1, 
                                 relwidth = 1)
        self.Label_Keyence.configure(anchor = 'w')                                                  # Position de l'écriture à gauche du label
        self.Label_Keyence.configure(background = "#f4f4f4")
        self.Label_Keyence.configure(font = self.CS_Label_Value)
        self.Label_Keyence.configure(text = 'Keyence : ###% \n(####)')                              # Titre par défaut


        # Création frame Rebut Kogame
        self.Frame_Kogame = tk.Frame(self)
        self.Frame_Kogame.place(relx = 0.75, 
                                rely = 0.623, 
                                relheight = 0.095, 
                                relwidth = 0.216)
        self.Frame_Kogame.configure(background = "#f4f4f4")

        # Création label Rebut Kogame
        self.Label_Kogame = tk.Label(self.Frame_Kogame)
        self.Label_Kogame.place(relx = 0.0, 
                                rely = 0.0, 
                                relheight = 1, 
                                relwidth = 1)
        self.Label_Kogame.configure(anchor = 'w')                                                   # Position de l'écriture à gauche du label
        self.Label_Kogame.configure(background = "#f4f4f4")
        self.Label_Kogame.configure(font = self.CS_Label_Value)
        self.Label_Kogame.configure(text = 'Kogame : ###% \n(####)')                                # Titre par défaut

        # Création frame Titre cpk
        self.Frame_Cpk = tk.Frame(self)
        self.Frame_Cpk.place(relx = 0.818, 
                             rely = 0.733, 
                             relheight = 0.045, 
                             relwidth = 0.055)
        self.Frame_Cpk.configure(background="#f4f4f4")

        self.Label_Cpk = tk.Label(self.Frame_Cpk)
        self.Label_Cpk.place(relx = 0.0, 
                             rely = 0.0,
                             relheight = 1, 
                             relwidth = 1)
        self.Label_Cpk.configure(background="#f4f4f4")
        self.Label_Cpk.configure(font=self.CS_Label_cpk)
        self.Label_Cpk.configure(text='Cpk')                                                        # Titre


        # Création frame CPK
        self.Frame_Cpk_Val = tk.Frame(self)
        self.Frame_Cpk_Val.place(relx = 0.75, 
                                 rely = 0.782, 
                                 relheight = 0.184, 
                                 relwidth = 0.206)
        self.Frame_Cpk_Val.configure(background = "#f4f4f4")

        # Création label CPK hauteur
        self.Label_Hauteur = tk.Label(self.Frame_Cpk_Val)
        self.Label_Hauteur.place(relx = 0.0, 
                                 rely = 0.0, 
                                 relheight = 0.155, 
                                 relwidth=1)
        self.Label_Hauteur.configure(anchor = 'w')                                                  # Position de l'écriture à gauche du label
        self.Label_Hauteur.configure(background = "#f4f4f4")
        self.Label_Hauteur.configure(font = self.CS_Label_cpk)
        self.Label_Hauteur.configure(text = 'Hauteur : ##.####')                                    # Titre par défaut

        # Création label CPK parallellisme
        self.Label_Parallelisme = tk.Label(self.Frame_Cpk_Val)
        self.Label_Parallelisme.place(relx = 0.0, 
                                      rely = 0.161, 
                                      relheight = 0.155, 
                                      width = 395)
        self.Label_Parallelisme.configure(anchor = 'w')                                             # Position de l'écriture à gauche du label
        self.Label_Parallelisme.configure(background = "#f4f4f4")
        self.Label_Parallelisme.configure(font = self.CS_Label_cpk)
        self.Label_Parallelisme.configure(text = 'Parallélisme : ##.####')                          # Titre par défaut

        # Création label CPK planeite face Hexag
        self.Label_Planeite_f_h = tk.Label(self.Frame_Cpk_Val)
        self.Label_Planeite_f_h.place(relx = 0.0, 
                                      rely = 0.323, 
                                      relheight = 0.155, 
                                      width = 395)
        self.Label_Planeite_f_h.configure(anchor = 'w')                                             # Position de l'écriture à gauche du label
        self.Label_Planeite_f_h.configure(background = "#f4f4f4")
        self.Label_Planeite_f_h.configure(font = self.CS_Label_cpk)
        self.Label_Planeite_f_h.configure(text = 'Planéité face hex : ##.####')                     # Titre par défaut

        # Création label CPK planeite face
        self.Label_Planeite_f = tk.Label(self.Frame_Cpk_Val)
        self.Label_Planeite_f.place(relx = 0.0, 
                                    rely = 0.484, 
                                    relheight = 0.155, 
                                    width = 395)
        self.Label_Planeite_f.configure(anchor = 'w')                                               # Position de l'écriture à gauche du label
        self.Label_Planeite_f.configure(background = "#f4f4f4")
        self.Label_Planeite_f.configure(font = self.CS_Label_cpk)
        self.Label_Planeite_f.configure(text = 'Planéité face : ##.####')                           # Titre par défaut

        # Création label CPK Rectitude face 1
        self.Label_Rectitude1 = tk.Label(self.Frame_Cpk_Val)
        self.Label_Rectitude1.place(relx = 0.0, 
                                    rely = 0.645, 
                                    relheight = 0.155, 
                                    width = 395)
        self.Label_Rectitude1.configure(anchor = 'w')                                               # Position de l'écriture à gauche du label
        self.Label_Rectitude1.configure(background = "#f4f4f4")
        self.Label_Rectitude1.configure(font = self.CS_Label_cpk)
        self.Label_Rectitude1.configure(text = 'Rectitude face 1 : ##.####')                        # Titre par défaut

        # Création label CPK Rectitude face 2
        self.Label_Rectitude2 = tk.Label(self.Frame_Cpk_Val)
        self.Label_Rectitude2.place(relx = 0.0, 
                                    rely = 0.806, 
                                    relheight = 0.155, 
                                    width = 395)
        self.Label_Rectitude2.configure(anchor = 'w')                                               # Position de l'écriture à gauche du label
        self.Label_Rectitude2.configure(background = "#f4f4f4")
        self.Label_Rectitude2.configure(font = self.CS_Label_cpk)
        self.Label_Rectitude2.configure(text = 'Rectitude face 2 : ##.####')                        # Titre par défaut


        # Création frame Information interface
        self.Frame_Info_Interface = tk.Frame(self)
        self.Frame_Info_Interface.place(relx = 0.432, 
                                        rely = 0.944, 
                                        relheight = 0.038, 
                                        relwidth = 0.143)
        self.Frame_Info_Interface.configure(background = "#f4f4f4")

        # Création label Information interface
        self.Label_Info_Interface = tk.Label(self.Frame_Info_Interface)
        self.Label_Info_Interface.place(relx = 0.0, 
                                        rely = 0.0, 
                                        relheight = 1, 
                                        relwidth = 1)
        self.Label_Info_Interface.configure(background = "#f4f4f4")
        self.Label_Info_Interface.configure(font = self.CS_Label_Value)
        self.Label_Info_Interface.configure(text = 'SUIVI JOUR')                                    # Titre


        # Création du bouton retour menu démarrage
        self.retourButton = tk.Button(self, 
                                      text = "Retour",                                              # Titre 
                                      command = self.closing_processus_jour,                        # Action lors d'un clique (fermeture)
                                      bd = 2,
                                      foreground = 'black',                                         # Police en noir 
                                      background = 'slate gray')                                    # Fond en slate gray
                                                                     
        self.retourButton.place(relx = 0.95, 
                                rely = 0.941, 
                                relheight = 0.042, 
                                relwidth = 0.042)

        # Création bouton changement de fenêtre suivi
        self.Button_OF = tk.Button(self,
                                   text = 'OF',                                                     # Titre
                                   command = self.closing_processus_jourBis,
                                   bd = 2,
                                   foreground = 'black',                                            # Police en noir 
                                   background = 'slate gray')                                       # Fond en slate gray)
        self.Button_OF.place(relx = 0.01, 
                             rely = 0.941, 
                             relheight = 0.042, 
                             relwidth = 0.066)

        # Progress Bar export données

        # Création frame progress bar
        self.Frame_Progress = tk.Frame(self)
        self.Frame_Progress.place(relx = 0.667, 
                                  rely = 0.084, 
                                  relheight = 0.04, 
                                  relwidth = 0.333)
        self.Frame_Progress.configure(background = "#f4f4f4")
        s = ttk.Style()
        s.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
        self.progressbar = ttk.Progressbar(self.Frame_Progress, 
                                           style="green.Horizontal.TProgressbar")                   # Ajout du style
        self.progressbar.place(relx = 0.5,                                                          # Placement relatif par rapport à la taille de la fenêtre en x 
                               rely = 0.5,                                                          # Placement relatif par rapport à la taille de la fenêtre en y 
                               relwidth = 0.8,
                               anchor = 'center')                                                   # Largeur fixe 


        self.bind('<Configure>',self.resize)

    # Script de fermture et d'ouverture de fenêtre
    def closing_processus_jour(self):

        self.controller.show_frame(self.startPage)                                                  # On ouvre la fenêtre de démarrage

        try:
            self.after_cancel(self.idAfter)                                                         # On annule la loop infini pour le rafraichissement des données

        except:
            print("(Jour) SQL Loop cancel : //!\\")

    def closing_processus_jourBis(self):

        self.controller.show_frame(self.processusPageOF)                                                 # Affichage de la fenêtre
        frame = self.frames[self.processusPageOF]                                                        # Recupération de l'objet frame OF ou Jour
        eventT = threading.Thread(target = frame.sql_event)                                         # Thread pour le démarrage de la loop SQL
        eventT.start()                                                                              # Démarrage du thread

        try:
            self.after_cancel(self.idAfter)                                                         # On annule la loop infini pour le rafraichissement des données

        except:
            print("(Jour) SQL Loop cancel : //!\\")


    # Gestion des requêtes SQL
    def sql_event(self=None):
        
        try:
            self.conn = SQL_Cursor.sql_connection_DB_QA()                                           # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.sql_get_reference()                                                                # Récupération de la référence de l'OF en cours                    

            self.sql_get_of()                                                                       # Récupération du numéro de l'OF en cours

            self.sql_get_bekido()                                                                   # Récupération du Bekido

            self.sql_get_chokko()                                                                   # Récupération du Chokko

            self.sql_get_prevision()                                                                # Récupération des prévisions, actuel, et delta de pièces

            self.sql_get_event()                                                                    # Récupération du nombre d'évènements pour l'OF

            self.sql_get_rebut()                                                                    # Récupération du nombre de rebut total, Keyence, Kogame

            self.sql_get_avancement()                                                               # Récupération de l'avancement

            self.sql_chart_keyence()                                                                # Récupération des données Keyence pour le graphique

            self.sql_chart_event()                                                                  # Récupération des données évènements pour le graphique

            #print("(JOUR) Requête SQL QA GATE : OK")                                                # Marqueur                                                                    

            SQL_Cursor.sql_deconnection_DB_QA(self.cursor)                                          # Fermeture de la connection

        except:
            print("(JOUR) Requête SQL QA GATE : //!\\")                                             # Marqueur

        try:
            self.conn = SQL_Cursor.sql_connection_DB_Measurlink()                                   # Connection à la base de donnée
            self.cursor = self.conn.cursor()                                                        # Récupération du curseur

            self.sql_chart_cpk_kogame()                                                             # Récupération des données Kogame pour le graphique

            #print("(JOUR) Requête SQL Measurlink : OK")                                             # Marqueur 

            SQL_Cursor.sql_deconnection_DB_Measurlink(self.cursor)                                  # Fermeture de la connection

        except:
            print("(JOUR) Requête SQL Measurlink : //!\\")                                          # Marqueur 

        self.idAfter = self.after(30000, self.sql_event)                                            # Démarrage loop mise à jour données (30s)     


    # Script récupération de la référence de l'OF en cours
    def sql_get_reference(self):
        try:
            sql = """\
                        SELECT reference FROM QAGATE_1_MainTable WHERE idPiece = (SELECT MAX(idPiece) FROM QAGATE_1_MainTable)
                  """
            self.cursor.execute(sql)                                                                # Exécute la requête
            self.reference = self.cursor.fetchval()                                                 # Récupère la valeur

            self.Label_Ref['text'] = str(self.reference)

            #print("(JOUR) Reference : OK")                                                          # Marqueur

        except:
            self.Label_Ref['text'] = '######-####'
            print("(JOUR) Reference : //!\\")                                                       # Marqueur

        
    # Script récupération du numéro de l'OF en cours
    def sql_get_of(self):

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_OF]
                  """
            self.cursor.execute(sql)                                                                # Exécute la requête
            self.valOF = self.cursor.fetchval()                                                     # Récupère la valeur
            
            self.Label_OF["text"] = "OF : " + str(self.valOF)                                       # Update du numéro de l'OF

            #print("(JOUR) Numéro OF : OK")                                                          # Marqueur

        except:
            self.Label_OF["text"] = "OF : ######"                                                   # Numéro de l'OF par défaut
            print("(JOUR) Numéro OF : //!\\")                                                       # Marqueur


    # Script récupération du Bekido
    def sql_get_bekido(self):

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_Bekido_Jour]
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valBekido = self.cursor.fetchval()                                                      # Récupère la valeur

            if(int(valBekido) > 100):                                                               # Si Bekido >100% alors
                self.Label_Bekido["text"] = "Bekido : >100%"                                        # Update du bekido
            else:                                                                                   # Sinon
                self.Label_Bekido["text"] = "Bekido : " + str(valBekido) + "%"                      # Update du bekido
        
            if (int(valBekido) <= 90):                                                              # Si Bekido <=90% alors
                self.Label_Bekido["background"] = "red"                                             # Fond en rouge
            else :                                                                                  # Sinon
                self.Label_Bekido["background"] = "green"                                           # Fond en vert

            #print("(JOUR) Bekido : OK")                                                             # Marqueur

        except:
            self.Label_Bekido["text"] = "Bekido : ###%"                                             # Bekido par défaut
            print("(JOUR) Bekido : //!\\")                                                          # Marqueur

    # Script récupération du Chokko
    def sql_get_chokko(self):

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_Chokko_Jour]
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valChokko = self.cursor.fetchval()                                                      # Récupère la valeur

            self.Label_Chokko["text"] = "Chokko : " + str(valChokko) + "%"                          # Update du chokko

            if (int(valChokko) <= 90):                                                              # Si Chokko <=90% alors
                self.Label_Chokko["background"] = "red"                                             # Fond en rouge
            else :                                                                                  # Sinon
                self.Label_Chokko["background"] = "green"                                           # Fond en vert
        
            #print("(JOUR) Chokko : OK")                                                             # Marqueur

        except:
            self.Label_Chokko["text"] = "Chokko : ###%"                                             # Chokko par défaut
            print("(JOUR) Bekido : //!\\")                                                          # Marqueur


    # Script récupération des prévisions, actuel, et delta de pièces
    def sql_get_prevision(self):

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_Prevision_Jour]
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valProcess = self.cursor.fetchall()                                                     # Récupère les 3 valeurs

            for row in valProcess:                                                                  # Système de pointeurs pour aller chercher chaque valeur
                valDelta = int(row.Delta)
                self.Label_Prevision["text"] = "Prévision : " + str(row.Prevision)                  # Update des prévisions
                self.Label_Actuel["text"] = "Actuel : " + str(row.Actuel)                           # Update de l'actuel
                self.Label_Delta["text"] = "Delta : " + str(valDelta)                               # Update du delta

            if(valDelta < 0):                                                                       # Si Delta négatif alors
                self.Label_Delta["background"] = "red"                                              # Fond en rouge
                self.Label_Delta["foreground"] = "#ffffff"                                          # Police en blanc
            elif(valDelta > 0):                                                                     # Si Delta positif alors
                self.Label_Delta["background"] = "green"                                            # Fond en vert
                self.Label_Delta["foreground"] = "#ffffff"                                          # Police en blanc
            else:                                                                                   # Si Delta = 0
                self.Label_Delta["background"] = "#f4f4f4"                                          # Fond en gris clair
                self.Label_Delta["foreground"] = "#000000"                                          # Police en noir

            #print("(JOUR) Prevision : OK")                                                          # Marqueur

        except:
            self.Label_Prevision["text"] = "Prévision : ####"                                       # Prevision par défaut
            self.Label_Actuel["text"] = "Actuel : ####"                                             # Actuel par défaut
            self.Label_Delta["text"] = "Delta : ####"                                               # Delta par défaut
            print("(JOUR) Prevision : //!\\")                                                        # Marqueur


    # Script récupération du nombre d'évènements pour l'OF
    def sql_get_event(self):

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_Event_Count_Jour]
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valEvent = self.cursor.fetchval()                                                       # Récupère la valeur

            self.Label_Total_Event["text"] = 'Total\n' + str(valEvent)                             # Update du nombre d'évènement

            sql = """\
                        SELECT TOP(1) MnemoniqueAlarme, QAGATE_1_EventData.etat
                        FROM QAGATE_1_EventData 
                        INNER JOIN QAGATE_1_EventInfo 
                        ON QAGATE_1_EventData.code = QAGATE_1_EventInfo.code 
                        ORDER BY idEvent DESC
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valCode = self.cursor.fetchall()                                                        # Récupère la valeur
            for row in valCode:
                self.Label_Code_Event["text"] = 'Code : ' + str(row.MnemoniqueAlarme)               # Update du nombre d'évènement
                if(int(row.etat) == 3):
                    self.Label_Code_Event["foreground"] = "green"                                   # Police en vert
                else:
                    self.Label_Code_Event["foreground"] = "red"                                     # Police en rouge

            #print("(JOUR) Total évènement : OK")                                                    # Marqueur

        except:
            self.Label_Total_Event["text"] = 'Total\n####'                                          # Nombre d'évènement par défaut
            self.Label_Code_Event["text"] = 'Code : #'                                              # Update du nombre d'évènement
            print("(JOUR) Total évènement : //!\\")                                                 # Marqueur


    # Script récupération du nombre de rebut total, Keyence, Kogame
    def sql_get_rebut(self):

        try:
            sql = """\
                    EXEC [dbo].[QAGATE_1_Rebut_Jour]
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valProcess = self.cursor.fetchall()                                                     # Récupère les 3 valeurs

            for row in valProcess:                                                                  # Système de pointeurs pour aller chercher chaque valeur
                self.Label_Rebut["text"] = "Rebut total : " + str(row.Total)                        # Update rebut total
                self.Label_Keyence["text"] = "Keyence : " + str(row.PKeyence) + '%\n(' + str(row.Keyence) + ')'
                                                                                                    # Update rebut Keyence
                self.Label_Kogame["text"] = "Kogame : " + str(row.PKogame) + '%\n(' + str(row.Kogame) + ')'
                                                                                                    # Update rebut Kogame
            #print("(JOUR) Rebut : OK")                                                              # Marqueur

        except:
            self.Label_Rebut["text"] = "Rebut total : ####"                                         # Rebut total par défaut
            self.Label_Keyence["text"] = "Keyence : ###% (####)"                                    # Rebut Keyence par défaut
            self.Label_Kogame["text"] = "Kogame : ###% (####)"                                      # Rebut Kogame par défaut
            print("(JOUR) Rebut : //!\\")                                                           # Marqueur


    # Script récupération de l'avancement
    def sql_get_avancement(self):

        try:
            sql = """\
                    EXEC [dbo].[QAGATE_1_Avancement] 
                    """

            self.cursor.execute(sql)                                                                # Exécute la requête
            valProcess = self.cursor.fetchall()                                                     # Récupère les 2 valeurs

            for row in valProcess:                                                                  # Système de pointeurs pour aller chercher chaque valeur
                self.Label_Avancement["text"] = "Avancement : " + str(row.Avancement) + "%"         # Update valeur avancement
                self.progressbar['value'] = int(row.Avancement)
                self.Label_Debut_OF["text"] = "Debut OF : " + str(row.Date)                         # Update date début OF

            sql = """\
                        EXEC [dbo].[QAGATE_1_PredictionFinOF] 
                        """

            self.cursor.execute(sql)                                                                # Exécute la requête
            valFin = self.cursor.fetchval()                                                     # Récupère les 2 valeurs
            self.Label_Fin_OF["text"] = "Fin OF : " + str(valFin)
            #print("(JOUR) Info avancement : OK")                                                    # Marqueur

        except:
            self.Label_Avancement["text"] = "Avancement : ###%"                                     # Valeur avancement par défaut
            self.Label_Debut_OF["text"] = "Debut OF : ##/##/##"                                     # Date début OF par défaut
            self.Label_Fin_OF["text"] = "Fin OF : ##/##/##"
            print("(JOUR) Info avancement : //!\\")                                                 # Marqueur


    # Récupération des données Keyence pour le graphique
    def sql_chart_keyence(self):

        try:
            sql = """\
                    EXEC [dbo].[QAGATE_1_Chart_Keyence_Jour]               	 
                    """
            self.cursor.execute(sql)                                                                # Exécute la requête
            valChartKeyence = self.cursor.fetchall()                                                # Récupère le top 3 + reste des outils défectueux

            self.subplot_Keyence.clear()                                                            # Clear l'ancien graphique pour l'update

            for row in valChartKeyence:                                                             # Système de pointeurs pour aller chercher chaque valeur
                labels = str(row.First_Name), str(row.Second_Name), str(row.Third_Name), 'Autres'   # Update des 4 labels
                pieSizes = [int(row.FirstVal), int(row.SecondVal), int(row.ThirdVal),int(row.OtherVal)]
                                                                                                    # Update des valeurs liées aux labels
                colors = ['#ff9999','#66b3ff','#99ff99','#ffcc99']                                  # Couleurs de chaque "tranches" du camembert

            if(pieSizes == [0,0,0,0] ):                                                             # Si aucun outil problématique
                labels = 'Aucun rebut','Autre'                                                      # Label par défaut
                patches, texts, autotexts = self.subplot_Keyence.pie([1,0], 
                                                                     labels = labels, 
                                                                     colors = 'green', 
                                                                     autopct = '%1.1f%%',           # Affichage en pourcent
                                                                     shadow = True, 
                                                                     startangle = 90)
            else:                                                                                   # Si il y a des valeurs, alors
                patches, texts, autotexts = self.subplot_Keyence.pie(pieSizes, 
                                                                     labels = labels, 
                                                                     colors = colors, 
                                                                     autopct = '%1.1f%%',           # Affichage en pourcent 
                                                                     shadow = True, 
                                                                     startangle = 90)
            self.subplot_Keyence.set_title('KEYENCE')
            self.figure_Keyence.canvas.draw_idle()                                                  # Affichage du graphique update

            for text in texts:
                text.set_color('#535353')                                                           # Paramétrage de la couleur des textes

            for autotext in autotexts:
                autotext.set_color('#535353')                                                       # Paramétrage de la couleur des valeurs

            #print("(JOUR) Chart Keyence : //!\\")                                                   # Marqueur

        except:
            self.subplot_Keyence.clear()
            labels = 'ERREUR','Autre'                                                               # Label erreur
            patches, texts, autotexts = self.subplot_Keyence.pie([1,0], 
                                                                    labels = labels, 
                                                                    colors = ['red'], 
                                                                    autopct = '%1.1f%%',            # Affichage en pourcent
                                                                    shadow = True, 
                                                                    startangle = 90)
            self.subplot_Keyence.set_title('KEYENCE')
            self.figure_Keyence.canvas.draw_idle()                                                  # Affichage du graphique update

            print("(JOUR) Chart Keyence : //!\\")                                                   # Marqueur


    # Récupération des données Kogame pour le graphique + CPK
    def sql_chart_cpk_kogame(self):

        try:
            sql = """\
                        EXEC [dbo].[QAGATE_1_Rejet_Cpk_Kogame_Jour] @Reference = ?              	 
                  """
            self.cursor.execute(sql, 
                                self.reference)                                                     # Exécute la requête
            valChartKogame = self.cursor.fetchall()

            self.subplot_Kogame.clear()                                                             # Clear l'ancien graphique pour l'update

            for row in valChartKogame:                                                              # Système de pointeurs pour aller chercher chaque valeur
                
                self.Label_Hauteur["text"] = "Hauteur : " + str(row.Cpk_Hauteur)                    # Update cpk hauteur

                if (float(row.Cpk_Hauteur) <= 1.33):                                                # Si cpk < 1.33, alors
                    self.Label_Hauteur["foreground"] = "red"                                        # Texte en rouge
                else :                                                                              # Sinon
                    self.Label_Hauteur["foreground"] = "green"                                      # Texte en vert

                self.Label_Parallelisme["text"] = "Parallélisme : " + str(row.Cpk_Parallelisme)     # Update cpk parallelisme

                if (float(row.Cpk_Parallelisme) <= 1.33):                                           # Si cpk < 1.33, alors
                    self.Label_Parallelisme["foreground"] = "red"                                   # Texte en rouge
                else :                                                                              # Sinon
                    self.Label_Parallelisme["foreground"] = "green"                                 # Texte en vert

                self.Label_Planeite_f_h["text"] = "Planéité face hex : " + str(row.Cpk_PlaneiteFaceHex)                    
                                                                                                    # Update cpk planeité face hex
                if (float(row.Cpk_PlaneiteFaceHex) <= 1.33):                                        # Si cpk < 1.33, alors
                    self.Label_Planeite_f_h["foreground"] = "red"                                   # Texte en rouge
                else :                                                                              # Sinon
                    self.Label_Planeite_f_h["foreground"] = "green"                                 # Texte en vert

                self.Label_Planeite_f["text"] = "Planéité face : " + str(row.Cpk_PlaneiteFace)      # Update cpkplaneité face

                if (float(row.Cpk_PlaneiteFace) <= 1.33):                                           # Si cpk < 1.33, alors
                    self.Label_Planeite_f["foreground"] = "red"                                     # Texte en rouge
                else :                                                                              # Sinon
                    self.Label_Planeite_f["foreground"] = "green"                                   # Texte en vert

                self.Label_Rectitude1["text"] = "Rectitude face 1 : " + str(row.Cpk_Rectitude1)     # Update cpk rectitude face 1

                if (float(row.Cpk_Rectitude1) <= 1.33):                                             # Si cpk < 1.33, alors
                    self.Label_Rectitude1["foreground"] = "red"                                     # Texte en rouge
                else :                                                                              # Sinon
                    self.Label_Rectitude1["foreground"] = "green"                                   # Texte en vert

                self.Label_Rectitude2["text"] = "Rectitude face 2 : " + str(row.Cpk_Rectitude2)     # Update cpk rectitude face 2

                if (float(row.Cpk_Rectitude2) <= 1.33):                                             # Si cpk < 1.33, alors
                    self.Label_Rectitude2["foreground"] = "red"                                     # Texte en rouge
                else :                                                                              # Sinon
                    self.Label_Rectitude2["foreground"] = "green"                                   # Texte en vert


                labels = str(row.First_Name), str(row.Second_Name), str(row.Third_Name), 'Autres'   # Update des 4 labels
                pieSizes = [int(row.FirstVal), int(row.SecondVal), int(row.ThirdVal),int(row.OtherVal)]
                                                                                                    # Update des valeurs liées aux labels
                colors = ['#ff9999','#66b3ff','#99ff99','#ffcc99']                                  # Couleurs de chaque "tranches" du camembert

            if(pieSizes == [0,0,0,0] ):                                                             # Si aucun outil problématique
                labels = 'Aucun rebut','Autre'                                                      # Label par défaut
                patches, texts, autotexts = self.subplot_Kogame.pie([1,0], 
                                                                    labels = labels, 
                                                                    colors = 'green', 
                                                                    autopct = '%1.1f%%',            # Affichage en pourcent
                                                                    shadow = True, 
                                                                    startangle = 90)
            else:                                                                                   # Si il y a des valeurs, alors
                patches, texts, autotexts = self.subplot_Kogame.pie(pieSizes, 
                                                                    labels = labels, 
                                                                    colors = colors, 
                                                                    autopct = '%1.1f%%',            # Affichage du graphique update
                                                                    shadow = True, 
                                                                    startangle = 90) 
            self.subplot_Kogame.set_title('KOGAME')
            self.figure_Kogame.canvas.draw_idle()                                                   # Affichage du graphique update

            for text in texts:
                text.set_color('#535353')                                                           # Paramétrage de la couleur des textes

            for autotext in autotexts:
                autotext.set_color('#535353')                                                       # Paramétrage de la couleur des valeurs

            #print("(JOUR) Chart/CPK Kogame : OK")                                                   # Marqueur

        except:
            self.Label_Hauteur["text"] = "Hauteur : ##.####"                                        # Cpk hauteur par défaut
            self.Label_Parallelisme["text"] = "Parallélisme : ##.####"                              # Cpk parallelisme par défaut
            self.Label_Planeite_f_h["text"] = "Planéité face hex : ##.####"                         # Cpk planeite face hex par défaut
            self.Label_Planeite_f["text"] = "Planéité face : ##.####"                               # Cpk planeite face par défaut
            self.Label_Rectitude1["text"] = "Rectitude face 1 : ##.####"                            # Cpk rectitude face 1 par défaut
            self.Label_Rectitude2["text"] = "Rectitude face 2 : ##.####"                            # Cpk rectitude face 2 par défaut

            self.subplot_Kogame.clear()

            labels = 'ERREUR','Autre'                                                               # Label par défaut
            patches, texts, autotexts = self.subplot_Kogame.pie([1,0], 
                                                                labels = labels, 
                                                                colors = ['red'], 
                                                                autopct = '%1.1f%%',                # Affichage en pourcent
                                                                shadow = True, 
                                                                startangle = 90)
            self.subplot_Kogame.set_title('KOGAME')
            self.figure_Kogame.canvas.draw_idle()                                                   # Affichage du graphique update

            print("(JOUR) Chart/CPK Kogame : //!\\")                                                # Marqueur



    def sql_chart_event(self):

        try:
            date = pd.read_sql_query(
            '''SELECT CAST(DATEADD(HOUR, -6, GETDATE()) AS DATE) AS currentDate, CAST(DATEADD(HOUR, 18, GETDATE()) AS DATE) AS nextDate, GETDATE() AS currentTimestamp''', self.conn)
                                                                                                      # Récupération de l'horodatage 06:00:00/06:00:00 du jour et lendemain
            SQL_Query = pd.read_sql_query(
            '''EXEC [dbo].[QAGATE_1_Event_Jour]''', self.conn)                                        # Récupération des évènements

            self.subplot_Event.clear()                                                              # Clear l'ancien graphique pour l'update

            df = pd.DataFrame(SQL_Query, columns=['currentOF','code','etat','timeStamp'])           # Récupération et ajout des données d'évènements dans un tableau 2D Pandas


            df1 = pd.DataFrame([[None, None, df[-1:]['etat'].to_string(index = False), date['currentTimestamp'].to_string(index = False)]], columns = ['currentOF','code','etat','timeStamp'])
                                                                                                    # Tri des colonnes utiles
            df = df.append(df1)

            self.subplot_Event.step(df['timeStamp'], df['etat'], where='post')                      # Création du graphique en escalier
            
            self.subplot_Event.xaxis_date()                                                         # Axe X format date
            
            loc = plticker.MultipleLocator(base = 1.0)                                              # Espacement entre chaque valeur d'axe 
            self.subplot_Event.yaxis.set_major_locator(loc)                                         # Paramétrage espacement pour l'axe Y

            self.subplot_Event.set_ylim(top = 3.2, bottom = -0.2)                                   # Limite en Y pour une meilleure vue des évènements

            labels = [item.get_text() for item in self.subplot_Event.get_yticklabels()]             # Récupération des noms de label actuel
           
            labels[1] = 'Stop'
            labels[2] = 'Setup'
            labels[3] = 'Pause'
            labels[4] = 'Run'

            self.subplot_Event.set_yticklabels(labels)                                              # Changement des noms de label


            xfmt = mdates.DateFormatter('%H:%M')                                                    # Format de la date
            self.subplot_Event.xaxis.set_major_formatter(xfmt)                                      # Paramétrage format date en axe X
            

            hours = mdates.HourLocator(interval = 3)                                                # Espacement entre chaque valeur d'axe 
            self.subplot_Event.xaxis.set_major_locator(hours)                                       # Paramétrage espacement pour l'axe X

            self.subplot_Event.set_xlim(str(date['currentDate'].to_string(index=False)) + ' 06:00:00', str(date['nextDate'].to_string(index=False)) + ' 06:00:00')
                                                                                                    # Limite en X pour une meilleure vue des évènements
            #self.subplot_Event.set_xticklabels(self.subplot_Event.get_xticklabels(), rotation=45, horizontalalignment='right',fontweight='light',fontsize='x-large')

            self.figure_Event.canvas.draw_idle()                                                    # Affichage du graphique

            #print("(JOUR) Chart Event : OK")                                                        # Marqueur

        except:
            print("(JOUR) Chart Event : //!\\")                                                     # Marqueur


    # Script de changement de taille de police + button
    def resize(self, event):

        # Changement taille de bouton
        self.retourButton.configure(height = int(0.042*self.winfo_height()), 
                                    width = int(0.042*self.winfo_width()))  
        self.Button_OF.configure(height = int(0.042*self.winfo_height()), 
                                 width = int(0.066*self.winfo_width()))

        # Get size of frame
        heightIni = self.winfo_height()
        widthIni = self.winfo_width()

        height = heightIni // 2
        width = widthIni // 2


        # Look up table for font size
        if height < 100 or width < 200:
            height = 7
        elif height < 200 or width < 400:
            height = 10
        elif height < 320 or width < 500:
            height = 13
        elif height < 400 or width < 600:
            height = 17
        elif height < 500 or width < 800:
            height = 19
        else:
            height = 22
        
        # Resize font
        self.CS_Label_Title.configure(size = str(int(height*1.6)))
        self.CS_Label_OF.configure(size = str(int(height*2.2)))
        self.CS_Label_cpk.configure(size = str(int(height)))
        self.CS_Label_Value.configure(size = str(int(height*1.3)))
        self.CS_Label_Code.configure(size = str(int(height*0.82)))
 
