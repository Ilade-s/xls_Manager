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

from tkinter import *  # interface graphique
import xlsPlot  # Création de graphiques
import xlsReader  # Edtition de fichiers xls
import xlsWriter  # Lecture de fichier xls
import tkinter.filedialog as fldialog  # Choix du fichier
from tkinter.simpledialog import askstring  # récupéaration nom pour sauvergarde
import os  # Pour trouver le répertoire courant (os.getcwd)
from tkinter.ttk import *  # meilleurs widgets
import xlrd  # récupération des feuilles/sheets
import tkinter.messagebox as msgbox  # Messages d'information ou d'avertissement


class window(Tk):
    def __init__(self, master=None, titlefont=("Arial", 13), font=("Arial", 11)):
        """
        Classe représentant la fenêtre tkinter
        Initialise la classe et la fenêtre
        SubClass de Tk

        PARAMETRES : (modification non recommandée)
        -------------
            - master : None || fenêtre maîtresse 
                - default = None
            - titlefont : tuple(str,int)
                - style du texte pour les titres (police et taille)
                - default = ("Arial",13)
            - font : tuple(str,int)
                - style du texte normal (police et taille)
                - default = ("Arial",11)
        """
        super().__init__(master)
        self.retval = [[]]
        self.titlefont = titlefont
        self.font = font
        self.master = master
        self.fonction = ""
        self.Module = ""
        self.FilePath = ""

    def ModuleChoice(self):
        """
        Fenêtre initiale, permet de choisir le module à utiliser
        Une fois les choix confirmés, la fenêtre de choix des paramètres est automatiquement appelée
        """
        self.funcs = []
        self.OpenButton = None
        funcchoice = StringVar()
        self.warninglabel = None

        def Confirmation():
            """
            Action lors de la confirmation des choix
            """
            # récupération choix
            self.fonction = funcchoice.get()
            self.Module = value.get()
            # debug
            print("classe/module choisi :", self.Module)
            print("fonction choisie :", self.fonction)
            print("Fichier choisi :", self.FilePath)
            # appel fenêtre secondiaire (choix des paramètres)
            self.WinParam()

        def OpenFile():
            NewPath = fldialog.askopenfilename(initialdir=os.getcwd(
            ), title="Tableur à utiliser", filetypes=(("xls files", "*.xls"), ("all files", "*.*")))
            if NewPath != "":
                self.FilePath = NewPath
                self.OpenButton["text"] = self.FilePath.split("/")[-1][:-4]
                ExitButton["state"] = "normal"

        def IsChecked():
            """
            Action lors du choix de classe
            Affiche la liste des fonctions disponibles dans la classe choisie
            Si des fonctions d'une classe précedemment choisie sont présentes, elles seront préalablement effacées
            """
            if len(self.funcs) != 0:
                for f in self.funcs:
                    f.destroy()
                self.funcs = []
            if self.OpenButton != None:
                self.OpenButton.destroy()
            ExitButton["state"] = "disabled"
            self.OpenButton = Button(
                self, text="Choix du fichier", command=OpenFile, state="normal", width=20)
            # récupération de la classe souhaitée
            if value.get() == xlsPlot.__name__:  # xlsPlot
                if self.warninglabel != None:
                    self.warninglabel.destroy()
                classe = xlsPlot.xlsDB
            elif value.get() == xlsReader.__name__:  # xlsReader
                if self.warninglabel != None:
                    self.warninglabel.destroy()
                classe = xlsReader.xlsData
            elif value.get() == xlsWriter.__name__:  # xlsWriter
                classe = xlsWriter.xlsWriter
                self.FilePath = ""
                ExitButton["state"] = "normal"
                self.warninglabel = Label(
                    self, text="Ne choisir un fichier uniquement si vous souhaitez le modifier \nPour créer un nouveau fichier, laisser vide", font=self.font, anchor=CENTER)
                self.warninglabel.pack(pady=5)
            else:
                classe = xlsReader.xlsData
            self.OpenButton.pack()
            # Ajout des fonctions dans la fenêtre
            for func in [method for method in dir(classe) if method[0] != "_" and method != "SaveFile"]:
                self.funcs.append(Radiobutton(
                    self, text=func, variable=funcchoice, value=func))
                self.funcs[-1].pack(anchor="w", padx=10)
                funcchoice.set(func)

        self.geometry("550x400")
        self.title("xlsManager : choix fonction")

        # texte de présentation
        Label(self, text="Bienvenue dans cette interface de gestion de tableurs xls",
              font=self.titlefont).pack(pady=5, anchor=CENTER)
        Label(self, text="Pour commencer, merci de choisir le module que vous souhaitez utiliser. \n Ses fonctions vous seront ensuite proposées", font=self.font).pack(pady=5)

        # Choix de module
        value = StringVar()
        value.set(xlsReader.__name__)
        # Boutons de choix de classe
        Radiobutton(self, text="xlsReader : lecture de fichier xls", command=IsChecked,
                    variable=value, value=xlsReader.__name__).pack(anchor="w", padx=10)
        Radiobutton(self, text="xlsWriter : Edition de fichier xls", command=IsChecked,
                    variable=value, value=xlsWriter.__name__).pack(anchor="w", padx=10)
        Radiobutton(self, text="xlsPlot : création de graphique", command=IsChecked,
                    variable=value, value=xlsPlot.__name__).pack(anchor="w", padx=10)

        Label(self, text="Fonctions disponibles :",
              font=self.titlefont).pack(pady=5, anchor=CENTER)
        # Bouton de confirmation (ferme la fenêtre)
        ExitButton = Button(self, text="Confirmer",
                            command=Confirmation, state="disabled", width=20)
        ExitButton.place(x=210, y=350)

    def WinParam(self):
        """
        Fenêtre d'entrée des paramètres
        Fonction mère de WinXlsReader, WinXlsWriter, WinXlsPlot
        """
        self.ClearWindow()
        # Ajout widgets communs
        Label(self, text="Classe/Module : "+self.Module,
              font=self.titlefont).pack(anchor=CENTER)
        Label(self, text="Fonction : "+self.fonction,
              font=self.titlefont).pack(anchor=CENTER, pady=20)
        Label(self, text="Fichier choisi : "+self.FilePath.split("/")
              [-1][:-4]+"\n("+self.FilePath+")", font=self.font).pack(anchor="w")
        # appel de la fenêtre correspondante à la classe demandée/au module demandé
        if self.Module == "xlsReader":
            self.WinXlsReader()
        elif self.Module == "xlsWriter":
            self.WinXlsWriter()
        elif self.Module == "xlsPlot":
            self.WinXlsPlot()

    def WinXlsReader(self):
        """
        Fenêtre pour le module d'xlsReader
        """
        self.title("xlsReader : initialisation de la classe")

        def WinRetFunc():
            """
            Fenêtre qui affiche le résultat de la fonction
            Permet aussi de sauvegarder les résultats, par xlsWriter
            """
            AffChoice = IntVar()
            AffChoice.set(1)

            def Affichage():
                """
                Action lors du choix d'affichage
                """
                Aff = AffChoice.get()  # récupération type d'affichage
                if Aff == 1:  # Affichage normal
                    self.ClearWindow()
                    self.title("Affichage simple")
                    self.geometry("{}x{}".format(1000, 400))
                    Label(self, text="Affichage résultat :",
                          font=self.titlefont).pack(anchor=CENTER)
                    if self.formatage == "rowmat" or self.formatage == "colmat":
                        for i in self.retval:
                            Label(self, text=str(i)).pack(
                                pady=5, padx=5)
                    else:
                        for i in self.retval.items():
                            for j in i:
                                Label(self, text=str(j)).pack(
                                    pady=5, padx=5)

                    ExitButton = Button(self, text="Retour",
                                        command=WinRetFunc, width=20)
                    ExitButton.place(x=290, y=350)
                    QuitButton = Button(self, text="Quitter",
                                        command=self.destroy, width=20)
                    QuitButton.place(x=580, y=350)

                elif Aff == 3:  # sauvegarde (xlsWriter)
                    path = fldialog.asksaveasfilename(initialdir=os.getcwd(), title="Sauvegarde résultat", filetypes=(
                        ("xls files (.xls)", "*.xls"), ("all files", "*.*")), defaultextension=".xls", initialfile="save")
                    filename = path.split("/")[-1][:-4]
                    print(filename)
                    sheetname = askstring(
                        "Nom de feuille/sheet", "Nom de la feuille de tableur:")
                    try:
                        xls = xlsWriter.xlsWriter(SheetName=sheetname)
                        if self.formattage != "dict":  # "Conversion" résultat en dictionnaire pour xlsWriter
                            dictretval = self.xls.Lecture(
                                self.Rowstart, self.Rowstop, self.Colstart, self.Colstop, "dict")
                        else:
                            dictretval = self.retval
                        xls.AddData(dictretval, self.Colstart,
                                    self.Rowstart, autoSave=(True, filename))
                        msgbox.showinfo(
                            "Sauvergarde réussie", "le résultat a été sauvegardé avec succès sous le nom "+filename)
                    except Exception as e:
                        print("Exeception reçue :", e)
                        msgbox.showerror(
                            "Erreur sauvegarde", "Une erreur a été rencontrée durant la sauvegarde, veuillez réessayer")

            self.ClearWindow()
            self.title("xlsReader : Résultat")
            self.geometry("{}x{}".format(600, 200))
            # Widgets
            Label(self, text="Vous pouvez choisir comment afficher les résultats :").pack(
                pady=10, padx=5, anchor=CENTER)
            Radiobutton(self, text="Affichage normal (type print)", variable=AffChoice, value=1,
                        command=lambda: ExitButton.configure(state="normal")).pack(anchor="w", padx=30)
            #Radiobutton(self, text="Aperçu tableau", variable=AffChoice, value=2, command=lambda: ExitButton.configure(state="normal")).pack(anchor="w",padx=30)
            Radiobutton(self, text="Sauvegarde dans un tableur xls", variable=AffChoice, value=3,
                        command=lambda: ExitButton.configure(state="normal")).pack(anchor="w", padx=30)
            Label(self, text='Le résultat de la fonction peut aussi être retrouvé dans l\'attribut "retval" de la classe \nLorsque vous avez fini, vous pouvez fermer la fenêtre ou cliquer sur quitter').pack(
                pady=10, padx=5, anchor=CENTER)
            # bouton de confirmation
            ExitButton = Button(self, text="Confirmer",
                                command=Affichage, width=20, state="disabled")
            ExitButton.place(x=140, y=150)
            QuitButton = Button(self, text="Quitter",
                                command=self.destroy, width=20)
            QuitButton.place(x=350, y=150)

        def ConfirmationArgs():
            """
            Action lors de la confirmation des paramètres
            """
            # Récupération variables + conversion en int
            try:
                self.Rowstart = int(self.Rowstart.get())
            except:
                msgbox.showwarning(
                    "Ligne de départ invalide", "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            try:
                self.Rowstop = int(self.Rowstop.get())
            except:
                self.Rowstop = None
            try:
                self.Colstart = int(self.Colstart.get())
            except:
                msgbox.showwarning("Colonne de départ invalide",
                                   "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            try:
                self.Colstop = int(self.Colstop.get())
            except:
                msgbox.showwarning(
                    "Colonne de fin invalide", "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            self.formatage = self.formatage.get()
            # appel fonction
            if self.fonction == xlsReader.xlsData.Lecture.__name__:
                try:
                    self.retval = self.xls.Lecture(
                        self.Rowstart, self.Rowstop, self.Colstart, self.Colstop, self.formatage)
                    print("Résultat :", self.retval)
                    WinRetFunc()  # Affichage fenêtre finale (résultat)
                except:
                    msgbox.showerror(
                        "Erreur lecture", "la lecture du fichier à rencontré une erreur (certainement dûe à un mauvais paramètre). \nVeuillez réessayer")
                    WinArgs()  # réinitialisation fenêtre
                    return 0

        def WinArgs():
            """
            Deuxième partie de la fenêtre, permet de récupérer les paramètres nécessaires à la fonction choisie
            """
            self.Rowstart = StringVar()
            self.Rowstart.set("0")
            self.Rowstop = StringVar()
            self.Colstart = StringVar()
            self.Colstart.set("0")
            self.Colstop = StringVar()
            self.Colstop.set("0")
            self.formatage = StringVar()

            def IntValidate(text):
                """
                vérifie si l'entrée est une string convertible en integer
                """
                if text == "":
                    return True
                try:
                    int(text)+1
                except:
                    return False
                return True

            IntValid = self.register(IntValidate)

            self.ClearWindow()  # nettoyage fenêtre
            self.title(self.fonction+" : récupération des arguments")
            # Widgets de récupération
            # Valeurs numériques (zone de lecture)
            Label(self, text="Ligne de départ :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Rowstart, justify=LEFT, validate="key",
                  validatecommand=(IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Ligne de fin (laisser vide pour ligne maximale) :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Rowstop, justify=LEFT, validate="key",
                  validatecommand=(IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Colonne de départ :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Colstart, justify=LEFT, validate="key",
                  validatecommand=(IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Colonne de fin :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Colstop, justify=LEFT, validate="key",
                  validatecommand=(IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            # formatage
            self.formatage.set("colmat")
            Label(self, text="Formatage").pack(pady=15, padx=10, anchor="w")
            Radiobutton(self, text="- Matrice en colonne : cols[col[rows],...]", variable=self.formatage,
                        value="colmat", command=lambda: ExitButton.configure(state="normal")).pack(anchor="w", padx=30)
            Radiobutton(self, text="- Matrice en ligne : rows[row[col],...]", variable=self.formatage,
                        value="rowmat", command=lambda: ExitButton.configure(state="normal")).pack(anchor="w", padx=30)
            Radiobutton(self, text="- Dictionnaire des colonnes : cols{col[0] :[col[1:]],...}", variable=self.formatage,
                        value="dict", command=lambda: ExitButton.configure(state="normal")).pack(anchor="w", padx=30)
            # Bouton de confirmation
            ExitButton = Button(
                self, text="Confirmer", command=ConfirmationArgs, width=20, state="disabled")
            ExitButton.place(x=210, y=350)

        def ConfirmationInit():
            """
            Action lors de la confirmation des paramètres pour l'initialisation
            """
            if sheetChoice.get() in feuilles:  # feuille choisie
                # Récupération paramètres
                self.feuille = sheetChoice.get()
                print("feuille choisie :", self.feuille)
                # initialisation de la classe
                try:
                    self.xls = xlsReader.xlsData(
                        fullPath=self.FilePath, sheet=feuilles.index(self.feuille))
                except:  # erreur init
                    msgbox.showerror("Erreur initialisation classe",
                                     "la classe n'a pas pu être initialisée, veuillez réessayer")
                    self.destroy()
                print("Classe initialisée")
            else:  # feuille non spécifiée
                msgbox.showinfo("Feuille indéfinie",
                                "La feuille à lire n'a pas été spécifiée")

            WinArgs()  # Récupération arguments (2e fenêtre)

        # Placement widgets args initialisation
        feuilles = self.GetSheets()
        sheetChoice = StringVar()
        sheetChoice.set("")

        Label(self, text="Feuille à lire (sheet) :",
              font=self.font).pack(padx=10, anchor="w")
        # Donne la liste des feuilles du fichier, permettant à l'utilisateur d'en choisir une
        Combobox(self, values=feuilles, width=max(
            [len(f) for f in feuilles]), state="readonly", textvariable=sheetChoice, height=30).pack(padx=20, anchor="w")

        # Bouton de confirmation
        ExitButton = Button(self, text="Confirmer",
                            command=ConfirmationInit, width=20)
        ExitButton.place(x=210, y=350)

    def WinXlsWriter(self):
        """
        Fenêtre pour le module d'xlsWriter
        """
        self.title("xlsReader : paramètres")
        # Placement widgets

    def WinXlsPlot(self):
        """
        Fenêtre pour le module d'xlsPlot
        """
        self.title("xlsReader : paramètres")
        # Placement widgets

    def ClearWindow(self):
        """
        Efface tous les widgets de la fenêtre
        """
        for w in self.winfo_children():
            w.destroy()
        self.pack_propagate(0)

    def GetSheets(self):
        """
        Permet de récupérer la liste des feuilles
        """
        with xlrd.open_workbook(self.FilePath, on_demand=True) as file:
            return (file._sheet_names)


def main():
    win = window()
    win.ModuleChoice()  # fenêtre initiale
    win.mainloop()


if __name__ == '__main__':  # Exécution
    main()
