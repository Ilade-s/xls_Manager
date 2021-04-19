"""
xlsManagerGUI (interface graphique des trois modules)
------------------

MODULES UTILISABLES : 
    - xlsPlot : création de graphiques à partir d'un fichier
    - xlsWriter : édition de fichier xls
    - xlsReader : lecture de fichier xls

FONCTIONNEMENT :
    - Tout se passe dans l'interface graphique (ni console, ni python)
    - Le programme fonctionne en plusieurs étapes :
        - FENETRE 1 :
            - 1 : Choix du module à utiliser
            - 2 : Choix du fichier à utiliser
            - 3 : Choix de la fonction à utiliser
        - FENETRE 2 :
            - 4 : Entrée des paramètres nécessaires
        - FENETRE 3 :
            - 5 : Affichage du résultat (dépend de la fonction utilisée)
            - 6 : Demandes éventuelles (sauvegarde...)

MODULES UTILISES : (en plus des trois modules)
    - tkinter (interface graphique)
    - matplotlib (graphiques)
    - pandas
    - numpy
    - xlrd, xlwd et xlutils (gestion de fichiers xls)
"""

from tkinter import * # interface graphique
import xlsPlot # Création de graphiques
import xlsReader # Edtition de fichiers xls
import xlsWriter # Lecture de fichier xls
import tkinter.filedialog as fldialog # Choix du fichier
import os # Pour trouver le répertoire courant (os.getcwd)
from tkinter.ttk import *


class window(Tk):
    def __init__(self, master=None, titlefont=("Arial",13), font=("Arial",11)) -> None:
        super().__init__(master)
        self.titlefont = titlefont
        self.font = font
        self.master = master
        #print(self.ModuleChoice.__name__) # print le nom de fonction
    
    def ModuleChoice(self):
        """
        Fenêtre initiale, permet de choisir le module à utiliser
        """
        self.funcs = []
        self.OpenButton = None
        funcchoice = StringVar()
        self.warninglabel = None

        def OpenFile():
            self.FilePath = fldialog.askopenfilename(initialdir=os.getcwd(),title="Tableur à utiliser",filetypes=(("xls files","*.xls"),("all files","*.*")))
            if self.FilePath!="":
                self.OpenButton["text"] = self.FilePath.split("/")[-1][:-4]
                ExitButton["state"] = "normal"
        
        def IsChecked():
            """
            Action lors du choix de classe
            Affiche la liste des fonctions disponibles dans la classe choisie
            Si des fonctions d'une classe précedemment choisie sont présentes, elles seront préalablement effacées
            """
            if len(self.funcs)!=0:
                for f in self.funcs:
                    f.destroy()
                self.funcs = []
            if self.OpenButton!=None:
                self.OpenButton.destroy()
            ExitButton["state"] = "disabled"
            self.OpenButton = Button(self, text="Choix du fichier", command=OpenFile, state="normal", width=20)
            # récupération de la classe souhaitée
            if value.get() == xlsPlot.__name__: # xlsPlot
                if self.warninglabel!=None:
                    self.warninglabel.destroy()
                classe = xlsPlot.xlsDB
            elif value.get() == xlsReader.__name__: # xlsReader
                if self.warninglabel!=None:
                    self.warninglabel.destroy()
                classe = xlsReader.xlsData
            elif value.get() == xlsWriter.__name__: # xlsWriter
                classe = xlsWriter.xlsWriter
                self.FilePath = ""
                ExitButton["state"] = "normal"
                self.warninglabel = Label(self, text="Ne choisir un fichier uniquement si vous souhaitez le modifier \nPour créer un nouveau fichier, laisser le choix vide",font=self.font)
                self.warninglabel.pack(pady=5)
            else:
                classe = xlsReader.xlsData
            self.OpenButton.pack()
            # Ajout des fonctions dans la fenêtre
            for func in [method for method in dir(classe) if method[0]!="_" and method!="SaveFile"]:
                self.funcs.append(Radiobutton(self, text=func, variable=funcchoice, value=func))
                self.funcs[-1].pack(anchor="w",padx=10)
                funcchoice.set(func)
            
                    

        self.geometry("550x400")
        self.title("xlsManager : choix fonction")

        # texte de présentation
        Label(self, text="Bienvenue dans cette interface de gestion de tableurs xls",font=self.titlefont).pack(pady=5,anchor=CENTER)
        Label(self, text="Pour commencer, merci de choisir le module que vous souhaitez utiliser. \n Ses fonctions vous seront ensuite proposées",font=self.font).pack(pady=5)

        # Choix de module
        value = StringVar()
        value.set("xlsReader.__name__")
        # Boutons de choix de classe
        Radiobutton(self, text="xlsReader : lecture de fichier xls", command=IsChecked, variable=value, value=xlsReader.__name__).pack(anchor="w",padx=10)
        Radiobutton(self, text="xlsWriter : Edition de fichier xls", command=IsChecked, variable=value, value=xlsWriter.__name__).pack(anchor="w",padx=10)
        Radiobutton(self, text="xlsPlot : création de graphique", command=IsChecked, variable=value, value=xlsPlot.__name__).pack(anchor="w",padx=10)
        
        Label(self, text="Fonctions disponibles :", font=self.titlefont).pack(pady=5,anchor=CENTER)
        # Bouton de confirmation (ferme la fenêtre)
        ExitButton = Button(self, text="Confirmer", command=self.destroy, state="disabled",width=20)
        ExitButton.place(x=210,y=350)

        self.mainloop()

        self.fonction = funcchoice.get()
        self.Module = value.get()

    def ModuleUse(self):
        """
        Fenêtre principale, permet l'uilisation du module choisi précedemment
        Elle s'adapte dynamiquement par la décoration de la classe Module
        """

if __name__=='__main__': # Exécution
    win = window()

    win.ModuleChoice()

    print("classe/module choisie :",win.Module)
    print("fonction choisie :",win.fonction)
    print("Fichier choisi :",win.FilePath)
    #win.mainloop()

    
