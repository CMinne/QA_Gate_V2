import tkinter as tk
from tkinter import *
import os
import settings_page
import modify_password_page
import SQL_Cursor
import sha256

class page_password():
    def __init__(self, controllerStart):

        self.controllerStart = controllerStart

    def password_popup(self):

        # Cache la fenêtre principale
        self.controllerStart.withdraw()

        # Créer la fenêtre de mot de passe
        self.rootPassword = tk.Tk()

        # Lié aux autres pages
        self.pageSettings = settings_page.page_settings(self.controllerStart)                       # Page de settings
        self.pageModify = modify_password_page.page_password_modify(self.rootPassword)              # Page de modification de mot de passe
        self.sha256Var = sha256.encrypt_decrypt                                                     # Algorithm de cryptage-décryptage

        # Bind la touche croix rouge pour faire réapparaitre la fenêtre principale
        self.rootPassword.protocol("WM_DELETE_WINDOW", 
                                   lambda : self.change_window(controllerDown = self.rootPassword,
                                                               controllerUp= self.controllerStart
                                                               )
                                    )
        # Bind Enter pour valider le mot de passe
        self.rootPassword.bind('<Return>', 
                               self.submit
                               )
        
        # Setting de l'écran
        PosX = int(self.rootPassword.winfo_screenwidth()/2 - 111.5)                                    
        PosY = int(self.rootPassword.winfo_screenheight()/2 - 46.5)
        self.rootPassword.geometry("223x93"+ "+" + str(PosX) + "+" + str(PosY))
        self.rootPassword.resizable(0, 0)
        self.rootPassword.title("Password")

        # Chargement de l'icône JTEKT
        try:
            dirname = os.path.dirname(__file__)                                                     # Chemin relatif du fichier  
            filename = os.path.join(dirname, 'Image/jtekt_logo.ico')                                # Rajoute le chemin de l'image
            self.rootPassword.iconbitmap(filename)                                                  # Change l'icône
            print("Icon password : OK")                                                             # Marqueur OK
        except:
            print("Error loading icon Settings")                                                    # Marqueur erreur

        # Création frame Password (frame principale)
        framePassword = tk.Frame(self.rootPassword, 
                                 background = 'light gray'
                                 )
        framePassword.pack(side = "top",
                           fill = 'both',
                           expand = True
                           )

        # Création du label Title (insertion dans la frame principale)
        labelTitle = tk.Label(framePassword, 
                              text = "Sécurité utilisateur",
                              font = "comicsans 12 bold underline",
                              background = 'light gray'
                              )
        labelTitle.grid(row = 0,
                        columnspan = 3,
                        sticky = 'nsew'
                        )

        # Création du label password (insertion dans la frame principale)
        labelPassword = tk.Label(framePassword,
                                 text = "Password : ", 
                                 background = 'light gray'
                                 )
        labelPassword.grid(row = 1,
                           sticky = 'w'
                           )

        # Création du label error (insertion dans la frame principale)
        self.labelError = tk.Label(framePassword, 
                                   text = "", 
                                   fg='red', 
                                   background = 'light gray'
                                   )
        self.labelError.grid(row = 2,
                             columnspan=2,
                             sticky = 'w'
                             )

        # Création de l'entrée password (insertion dans la frame principale)
        self.entryPassword = tk.Entry(framePassword, 
                                      show='*'
                                      )
        self.entryPassword.grid(row=1, 
                                column=1, 
                                sticky = 'w'
                                )

        # Création de du bouton submit (insertion dans la frame principale)
        submitButton = tk.Button(framePassword, 
                                 text='Submit', 
                                 command = self.submit
                                 )
        submitButton.grid(row=3, 
                          column=1, 
                          sticky = 'e'
                          )

        # Création du bouton quit (insertion dans la frame principale)
        quitButton = tk.Button(framePassword, 
                               text='Quit',
                               command = lambda: self.change_window(controllerDown = self.rootPassword,
                                                                    controllerUp = self.controllerStart
                                                                    )
                               )
        quitButton.grid(row=3, 
                        column=2, 
                        sticky = 'w'
                        )

        # Création du bouton modifier (insertion dans la frame principale)
        modifierButton = tk.Button(framePassword, 
                                   text='Modify',
                                   command = self.pageModify.modify_popup
                                   )
        modifierButton.grid(row=3, 
                            column=0,
                            sticky = 'w'
                            )

        
        self.rootPassword.mainloop()

    def submit(self, event=None):

        conn = SQL_Cursor.sql_connection_DB_QA()
        cursor = conn.cursor()
        try:

            sql = """\
                    SELECT nPassword FROM QAGATE_1_PasswordData WHERE idPassword = 1 
                  """
            cursor.execute(sql)
            valPassword = cursor.fetchval()

        except:

            print("Aucun Password")

        if (self.sha256Var.verify_password(valPassword, self.entryPassword.get())): # Vérification du mot de passe entré
            SQL_Cursor.sql_deconnection_DB_QA(cursor)
            self.rootPassword.destroy()
            self.pageSettings.option_popup()
        else:
            self.labelError['text'] = "Wrong password"                                              # Message d'erreur
            SQL_Cursor.sql_deconnection_DB_QA(cursor)


    def change_window(self, controllerDown, controllerUp):

        # Destruit la fenêtre visible
        controllerDown.destroy()

        # Fais réapparaître la fenêtre caché
        controllerUp.deiconify()
