import tkinter as tk
from tkinter import *
import os
import SQL_Cursor

class page_settings():

    def __init__(self, controllerStart):

        self.controllerStart = controllerStart

    def option_popup(self):

        # Création du pop-up
        rootOption = tk.Tk()
        self.controllerOption = rootOption
        self.controllerOption.protocol("WM_DELETE_WINDOW", lambda : self.change_window(controllerDown = self.controllerOption,controllerUp= self.controllerStart))
        self.controllerOption.bind('<Return>', self.save_data)
        #rootOption.overrideredirect(True)

        # Setting de l'écran
        PosX = int(rootOption.winfo_screenwidth()/2 - 225)                                    
        PosY = int(rootOption.winfo_screenheight()/2 - 57)
        rootOption.geometry("450x114"+ "+" + str(PosX) + "+" + str(PosY))
        rootOption.resizable(0, 0)
        rootOption.title("Settings QA Gate 4.0")

        try:
            dirname = os.path.dirname(__file__)                                                     # Take the directory relative path  
            filename = os.path.join(dirname, 'Image/jtekt_logo.ico')                                # Add the image path
            rootOption.iconbitmap(filename)                                                         # Set the icon
            print("Icon Settings : OK") 
        except:
            print("Error loading icon Settings")                                                    # If error print this

        # Création de la frame principale
        frameMain = Frame(rootOption, 
                          background = 'light gray')
        frameMain.pack(fill = 'both', 
                       expand = True)

        # Creation label cycle (insertion dans frame principale)
        labelCyclePG1 = tk.Label(frameMain, 
                              text = "Temps cycle UE21",
                              background = 'light gray')
        labelCyclePG1.grid(row = 0, 
                        sticky = 'w')

        labelCyclePG2 = tk.Label(frameMain, 
                              text = "Temps cycle UE24",
                              background = 'light gray')
        labelCyclePG2.grid(row = 1, 
                        sticky = 'w')

        # Creation label piece (insertion dans frame principale)
        labelPiecePG1 = tk.Label(frameMain, 
                              text = "Nombre pièces UE21", 
                              background = 'light gray'
                              )
        labelPiecePG1.grid(row = 2, 
                        sticky = 'w')

        labelPiecePG2 = tk.Label(frameMain, 
                              text = "Nombre pièces UE24", 
                              background = 'light gray'
                              )
        labelPiecePG2.grid(row = 3, 
                        sticky = 'w')

        # Creation label cycle moyen (insertion dans frame principale)
        labelCyclePG1 = tk.Label(frameMain, 
                              text = "Temps cycle moyen UE21 : #### (s)",
                              background = 'light gray')
        labelCyclePG1.grid(row = 0,
                           column = 2,
                           sticky = 'w')

        labelCyclePG2 = tk.Label(frameMain, 
                              text = "Temps cycle moyen UE24 : #### (s)",
                              background = 'light gray')
        labelCyclePG2.grid(row = 1,
                           column = 2,
                           sticky = 'w')

        # Création des entrées (insertion dans frame principale)
        self.entryCycleUE21 = tk.Entry(frameMain)
        self.entryCycleUE21.grid(row=0, column=1)

        self.entryCycleUE24 = tk.Entry(frameMain)
        self.entryCycleUE24.grid(row=1, column=1)
        
        self.entryPieceUE21 = tk.Entry(frameMain)
        self.entryPieceUE21.grid(row=2, column=1)

        self.entryPieceUE24 = tk.Entry(frameMain)
        self.entryPieceUE24.grid(row=3, column=1)

        # Récupération des valeurs stockées

        conn = SQL_Cursor.sql_connection_DB_QA()
        cursor = conn.cursor()

        try:

            sql1 = """\
                    SELECT tempsCycle FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE21';
                """

            sql2 = """\
                    SELECT tempsCycle FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE24';
                """

            sql3 = """\
                    SELECT nombre FROM QAGATE_1_NombrePiece INNER JOIN QAGATE_1_Client ON QAGATE_1_NombrePiece.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE21';
                """
            
            sql4 = """\
                    SELECT nombre FROM QAGATE_1_NombrePiece INNER JOIN QAGATE_1_Client ON QAGATE_1_NombrePiece.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE24';
                """
            sql5 = """\
                    EXEC[dbo].[QAGATE_1_CycleMoyen]
                """

            # Insertion des valeurs dans chacunes des entrées
            cursor.execute(sql1)
            self.cycleUE21 = cursor.fetchval()
            self.entryCycleUE21.insert(10, self.cycleUE21)

            cursor.execute(sql2)
            self.cycleUE24 = cursor.fetchval()
            self.entryCycleUE24.insert(10, self.cycleUE24)

            cursor.execute(sql3)
            self.entryPieceUE21.insert(10, cursor.fetchval())

            cursor.execute(sql4)
            self.entryPieceUE24.insert(10, cursor.fetchval())

            cursor.execute(sql5)
            valProcess = cursor.fetchall()                                                 # Récupère toutes les données de la requête                         
            
            for row in valProcess:
                labelCyclePG1['text'] = "Temps cycle moyen UE21 : " + str(row.UE21) + " (s)"
                labelCyclePG2['text'] = "Temps cycle moyen UE24 : " + str(row.UE24) + " (s)"

            SQL_Cursor.sql_deconnection_DB_QA(cursor)

        except:
            print("Error get entry")
            SQL_Cursor.sql_deconnection_DB_QA(cursor)

        
        

        # Création du bouton Quit (insertion dans frame principale)
        buttonQuit = tk.Button(frameMain, 
                               text = 'Quit', 
                               command = lambda : self.change_window(controllerDown = self.controllerOption, controllerUp= self.controllerStart)
                               )
        buttonQuit.grid(row = 4, 
                        column = 0, 
                        sticky = 'w', 
                        pady = 4
                        )

        # Création du bouton Save (insertion dans frame principale)
        buttonSave = tk.Button(frameMain, 
                               text='Save', 
                               command=self.save_data                                               # Sauvegarde des données dans un fichier .json
                               )
        buttonSave.grid(row = 4, 
                        column = 1, 
                        sticky = 'w', 
                        pady = 4
                        )

        # Création du label error (insertion dans la frame principale)
        self.labelError = tk.Label(frameMain, 
                                   text = "", 
                                   fg='red', 
                                   background = 'light gray'
                                   )
        self.labelError.grid(row = 5,
                             columnspan=2,
                             sticky = 'w'
                             )

        # Démarrage de la boucle while infini de la fenêtre d'option
        rootOption.mainloop()

    def save_data(self, event=None):
 
        # Copie des valeurs entrées dans la bibliothèque de données du fichier .json
        conn = SQL_Cursor.sql_connection_DB_QA()
        cursor = conn.cursor()

        try:

            sql = """\
                    UPDATE QAGATE_1_Cycle SET tempsCycle = ? WHERE idClient = (SELECT QAGATE_1_Client.idClient FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE21');
                    UPDATE QAGATE_1_Cycle SET tempsCycle = ? WHERE idClient = (SELECT QAGATE_1_Client.idClient FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE24');
                    UPDATE QAGATE_1_NombrePiece SET nombre = ? WHERE idClient = (SELECT QAGATE_1_Client.idClient FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE21');
                    UPDATE QAGATE_1_NombrePiece SET nombre = ? WHERE idClient = (SELECT QAGATE_1_Client.idClient FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE24');
                """

            cursor.execute(sql, self.entryCycleUE21.get(), self.entryCycleUE24.get(), self.entryPieceUE21.get(), self.entryPieceUE24.get())

            if(float(self.cycleUE21) != float(self.entryCycleUE21.get())):
                sql = """\
                    INSERT INTO QAGATE_1_ArchiveTempsCycle (cycle, idClient) VALUES ( ?, (SELECT QAGATE_1_Client.idClient FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE21'));;
                """
                cursor.execute(sql, self.entryCycleUE21.get())

            if(float(self.cycleUE24) != float(self.entryCycleUE24.get())):
                sql = """\
                    INSERT INTO QAGATE_1_ArchiveTempsCycle (cycle, idClient) VALUES ( ?, (SELECT QAGATE_1_Client.idClient FROM QAGATE_1_Cycle INNER JOIN QAGATE_1_Client ON QAGATE_1_Cycle.idClient = QAGATE_1_Client.idClient WHERE QAGATE_1_Client.nameClient = 'UE24'));;
                """
                cursor.execute(sql, self.entryCycleUE24.get())

            conn.commit()
            SQL_Cursor.sql_deconnection_DB_QA(cursor)
            # Fermeture de la fenêtre d'option
            self.change_window(controllerDown= self.controllerOption,controllerUp= self.controllerStart)

        except:
           print("Error Entry")
           self.labelError['text'] = "Wrong entry"                                                  # Message d'erreur
           SQL_Cursor.sql_deconnection_DB_QA(cursor)
        
        

    def change_window(self, controllerDown, controllerUp):

        # Destruit la fenêtre visible
        controllerDown.destroy()

        # Fais réapparaître la fenêtre caché
        controllerUp.deiconify()




