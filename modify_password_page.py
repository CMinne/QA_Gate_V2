import tkinter as tk
from tkinter import *
import os
import syntax
import sha256
import SQL_Cursor


class page_password_modify():
    def __init__(self, controllerPassword):

        self.controllerPassword = controllerPassword

    def modify_popup(self):

        # Cache la fenêtre Password
        self.controllerPassword.withdraw()

        # Créer la fenêtre de modification du mot de passe
        self.rootModify = tk.Tk()

        # Lié au cryptage et décryptage du mot de passe
        self.sha256Var = sha256.encrypt_decrypt
        
        # Bind la touche croix rouge pour faire réapparaitre la fenêtre de modification du mot de passe
        self.rootModify.protocol("WM_DELETE_WINDOW", 
                                 lambda : self.change_window(controllerDown = self.rootModify, 
                                                             controllerUp= self.controllerPassword 
                                                             )
                                    )
        # Bind Enter pour valider la modification du mot de passe au clavier
        self.rootModify.bind('<Return>', 
                             self.save
                             )
        
        # Setting de l'écran
        PosX = int(self.rootModify.winfo_screenwidth()/2 - 135)                                    
        PosY = int(self.rootModify.winfo_screenheight()/2 - 67.5)
        self.rootModify.geometry("270x135"+ "+" + str(PosX) + "+" + str(PosY))
        self.rootModify.resizable(0, 0)
        self.rootModify.title("Settings Password")

        # Chargement de l'icône JTEKT
        try:
            dirname = os.path.dirname(__file__)                                                     # Chemin relatif du fichier  
            filename = os.path.join(dirname, 'Image/jtekt_logo.ico')                                # Rajoute le chemin de l'image
            self.rootModify.iconbitmap(filename)                                                    # Change l'icône
            print("Icon modify : OK")                                                               # Marqueur OK
        except:
            print("Error loading icon modify")                                                      # Marqueur erreur

        # Création frame Password (frame principale)
        framePassword = tk.Frame(self.rootModify, 
                                 background = 'light gray'
                                )
        framePassword.pack(side = "top",
                           fill = 'both',
                           expand = True
                           )

        # Création du label Title (insertion dans la frame principale)
        labelTitle = tk.Label(framePassword, 
                              text = "Changement mot de passe",
                              font = "comicsans 12 bold underline",
                              background = 'light gray'
                              )
        labelTitle.grid(row = 0,
                        columnspan = 3,
                        sticky = 'nsew'
                        )

        # Création du label old password (insertion dans la frame principale)
        labelOldPassword = tk.Label(framePassword, 
                                    text = "Password : ", 
                                    background = 'light gray' 
                                    )
        labelOldPassword.grid(row = 1,
                              sticky = 'w'
                              )

        # Création du label new password (insertion dans la frame principale)
        labelNewPassword = tk.Label(framePassword,
                                    text = "New password : ", 
                                    background = 'light gray'
                                    )
        labelNewPassword.grid(row = 2,
                              sticky = 'w'
                              )

        # Création du label confirm password (insertion dans la frame principale)
        labelConfirmPassword = tk.Label(framePassword,
                                 text = "Confirm password : ", 
                                 background = 'light gray'
                                 )
        labelConfirmPassword.grid(row = 3,
                           sticky = 'w'
                           )

        # Création du label error (insertion dans la frame principale)
        self.labelError = tk.Label(framePassword, 
                                   text = "", 
                                   fg='red', 
                                   background = 'light gray'
                                   )
        self.labelError.grid(row = 4,
                             columnspan=2,
                             sticky = 'w'
                             )

        # Création de l'entrée old password (insertion dans la frame principale)
        self.entryOldPassword = tk.Entry(framePassword, 
                                         show='*'
                                         )
        self.entryOldPassword.grid(row=1, 
                                   column=1, 
                                   sticky = 'w'
                                   )

         # Création de l'entrée new password (insertion dans la frame principale)
        self.entryNewPassword = tk.Entry(framePassword, 
                                         show='*'
                                         )
        self.entryNewPassword.grid(row=2, 
                                   column=1, 
                                   sticky = 'w'
                                   )

         # Création de l'entrée confirm password (insertion dans la frame principale)
        self.entryConfirmPassword = tk.Entry(framePassword, 
                                             show='*'
                                             )
        self.entryConfirmPassword.grid(row=3, 
                                       column=1, 
                                       sticky = 'w'
                                       )

        # Création de du bouton submit (insertion dans la frame principale)
        submitButton = tk.Button(framePassword, 
                                 text='Submit', 
                                 command = self.save
                                 )
        submitButton.grid(row=5, 
                          column=1, 
                          sticky = 'e'
                          )

        # Création du bouton quit (insertion dans la frame principale)
        quitButton = tk.Button(framePassword, 
                               text='Quit',
                               command = lambda: self.change_window(controllerDown = self.rootModify, 
                                                                    controllerUp = self.controllerPassword
                                                                    )
                               )
        quitButton.grid(row=5, 
                        column=2, 
                        sticky = 'w'
                        )

        # Début while infini
        self.rootModify.mainloop()

    def save(self, event=None):

        try:
            conn = SQL_Cursor.sql_connection_DB_QA()
            cursor = conn.cursor()

            sql = """\
                    SELECT nPassword FROM QAGATE_1_PasswordData WHERE idPassword = 1 
                  """
            cursor.execute(sql)
            valPassword = cursor.fetchval()


            if ((self.sha256Var.verify_password(valPassword, self.entryOldPassword.get())) 
                and 
                (self.entryNewPassword.get() == self.entryConfirmPassword.get())):                     # Vérification que l'ancien mdp correspond bien avec celui donné et que le nouveau et confirm sont identiques

                try:

                    sql = """\
                            UPDATE QAGATE_1_PasswordData SET nPassword = ? WHERE idPassword = 1
                        """
                    value = self.sha256Var.hash_password(self.entryConfirmPassword.get())               # Si oui, on crypte le nouveau mdp
                    cursor.execute(sql, value)
                    conn.commit()
                    SQL_Cursor.sql_deconnection_DB_QA(cursor)
                    self.change_window(controllerDown = self.rootModify,controllerUp = self.controllerPassword)

                except:
                    print("Error Password")
                    SQL_Cursor.sql_deconnection_DB_QA(cursor)

            else:
                self.labelError['text'] = "Wrong password"                                              # Message d'erreur 
        except:

            print("Aucun Password")

    def change_window(self, controllerDown, controllerUp):

        # Destruit la fenêtre visible
        controllerDown.destroy()

        # Fais réapparaître la fenêtre caché
        controllerUp.deiconify()

