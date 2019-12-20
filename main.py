import tkinter as tk
from tkinter import ttk
import start_page
import processus_page_jour
import processus_page_of
import screen_settings_main
import extraction_donnees
import extraction_kogame
import syntax
import threading

# Création de variable page

startPage = start_page.StartPage
processusPageJour = processus_page_jour.ProcessusJourPage
processusPageOF = processus_page_of.ProcessusOFPage
extractionPage = extraction_donnees.ExtractionQAPage
extractionKPage = extraction_kogame.ExtractionKoPage

# Fondation du GUI

class QAGate(tk.Tk):                                                                                # Heritage de Tkinter

    def __init__(self, *args, **kwargs):                                                            # Initialisation /!\ SELF est le lien entre tout. 
        tk.Tk.__init__(self, *args, **kwargs)                                                       # Initialisation de Tkinter

        # Set des dimensions du GUI
        screenSetting = screen_settings_main.setting_screen(self)                                   # Création de l'objet screenSetting pour configurer l'écran
        screenSetting.get_pc_screen_size()                                                          # Récupération des informations écrans
        screenSetting.set_screen()                                                                  # Paramétrage de l'écran

        #Set de l'icône JTEKT
        screenSetting.set_icon()    
        
        # Set du style des boutons (TTK theme)
        syntax.style()

        # Création du container
        self.container = tk.Frame(self, background="#f4f4f4")                                        # Créer le container où tout sera mis                                                                     
        self.container.pack(side = "top", fill = "both", expand = True)                              # Configuration 
        self.container.grid_rowconfigure(0, weight = 1)                                              # Configuration de ligne avec une priorité de 1
        self.container.grid_columnconfigure(0, weight = 1)                                           # Configuration de colonne avec une priorité de 1

        # Création de la bibliothèque des windows de l'application
        self.frames = {}

        # Intégration de la page de démarrage
        frame = startPage(self.container, 
                          self, 
                          processusPageJour, 
                          processusPageOF, 
                          self.frames, 
                          extractionPage, 
                          extractionKPage)                                                          # On crée la frame principale
        self.frames[startPage] = frame                                                              # On l'a stock dans la bibliothèque
        frame.grid(row = 0, column = 0, sticky = "nsew")                                            # Créer un cadrillage de 1x1 étendu sur toute la frame (fenêtre)

        # Création et démarrage des Thread
        startPageT = threading.Thread(target = lambda: self.show_frame(startPage))                  # On donne startPage pour faire apparaître la frame de démarrage
        startPageT.start()
        extractT = threading.Thread(target = self.extract)                                          # Frame d'extraction QA Gate
        extractT.start()
        extractKT = threading.Thread(target = self.extractK)                                        # Frame d'extraction Kogame
        extractKT.start()
        procJourT = threading.Thread(target = self.procJour)                                        # Frame suivi jour
        procJourT.start()
        procOFT = threading.Thread(target = self.procOF)                                            # Frame suivi OF
        procOFT.start()                                                          
        # /!\ permet à l'application de ne pas freezer

    def show_frame(self, cont):                                                                     # Le controleur est envoyé pour définir la frame qu'on veut afficher

        frame = self.frames[cont]                                                                   # On donne le nom de la frame a montré, pour que la bibliothèque prenne le bon contrôleur
        frame.tkraise()                                                                             # Affiche la frame

    def extract(self):

        frame = extractionPage(self.container, self, startPage)                                     # On crée la frame extraction QA Gate 4.0
        self.frames[extractionPage] = frame                                                         # On l'a stock dans la bibliothèque
        frame.grid(row = 0, column = 0, sticky = "nsew")                                            # Créer un cadrillage de 1x1 étendu sur toute la frame (fenêtre)
        self.show_frame(startPage)

    def extractK(self):

        frame = extractionKPage(self.container, self, startPage)                                    # On crée la frame extraction Kogame
        self.frames[extractionKPage] = frame                                                        # On l'a stock dans la bibliothèque
        frame.grid(row = 0, column = 0, sticky = "nsew")                                            # Créer un cadrillage de 1x1 étendu sur toute la frame (fenêtre)
        self.show_frame(startPage)

    def procJour(self):

        # Intégration de la page processus jour
        frame = processusPageJour(self.container, self, startPage, processusPageOF, self.frames)    # On crée la frame processus jour
        self.frames[processusPageJour] = frame                                                      # On l'a stock dans la bibliothèque
        frame.grid(row = 0, column = 0, sticky = "nsew")                                            # Créer un cadrillage de 1x1 étendu sur toute la frame (fenêtre)
        self.show_frame(startPage)

    def procOF(self):

        # Intégration de la page processus of
        frame = processusPageOF(self.container, self, startPage, processusPageJour, self.frames)                                    # On crée la frame processus of
        self.frames[processusPageOF] = frame                                                        # On l'a stock dans la bibliothèque
        frame.grid(row = 0, column = 0, sticky = "nsew")                                            # Créer un cadrillage de 1x1 étendu sur toute la frame (fenêtre)
        self.show_frame(startPage)

def main():

    rootMain = QAGate()
    rootMain.mainloop()

if __name__ == '__main__':
    main()
# Boucle while infini permettant de garder le GUI ouvert

# /!\ Ne rien mettre après