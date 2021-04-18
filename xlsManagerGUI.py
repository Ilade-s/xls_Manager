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
        funcchoice = StringVar()

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
            # récupération de la classe souhaitée
            if value.get() == xlsPlot.__name__: # xlsPlot
                classe = xlsPlot.xlsDB
            elif value.get() == xlsReader.__name__: # xlsReader
                classe = xlsReader.xlsData
            elif value.get() == xlsWriter.__name__: # xlsWriter
                classe = xlsWriter.xlsWriter
            else:
                classe = xlsReader.xlsData
            # Ajout des fonctions dans la fenêtre
            for func in [method for method in dir(classe) if method[0]!="_"]:
                self.funcs.append(Radiobutton(self, text=func, variable=funcchoice, value=func, font=self.font))
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
        Radiobutton(self, text="xlsReader : lecture de fichier xls", command=IsChecked, variable=value, value=xlsReader.__name__, font=self.font).pack(anchor="w",padx=10)
        Radiobutton(self, text="xlsWriter : Edition de fichier xls", command=IsChecked, variable=value, value=xlsWriter.__name__, font=self.font).pack(anchor="w",padx=10)
        Radiobutton(self, text="xlsPlot : création de graphique", command=IsChecked, variable=value, value=xlsPlot.__name__, font=self.font).pack(anchor="w",padx=10)
        
        Label(self, text="Fonctions disponibles :", font=self.titlefont).pack(pady=5,anchor=CENTER)
        # Bouton de confirmation (ferme la fenêtre)
        Button(self, text="Confirmer", command=self.destroy).place(x=240,y=350)

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
    #win.mainloop()

    
