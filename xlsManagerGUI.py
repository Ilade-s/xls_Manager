"""
xlsManagerGUI (interface graphique des trois modules)
------------------

MODULES UTILISABLES : 
    - xlsPlot : création de graphiques à partir d'un fichier
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
    - tkinter (interface graphique) :
        - ttk : widgets plus modernes
        - filedialog : ouverture/sauvegarde de fichiers (I/O)
        - simpledialog : demandes d'infos simples
        - messagebox : messages d'avertissement et d'erreur
    - matplotlib (graphiques)
    - Utilisés par les sous modules :
        - pandas
        - numpy
        - xlrd, xlwd et xlutils (gestion de fichiers xls)
"""

from tkinter import *  # interface graphique
import xlsPlot  # Création de graphiques
import xlsReader  # Edition de fichiers xls
import xlsWriter  # Lecture de fichier xls
import tkinter.filedialog as fldialog  # Choix du fichier
from tkinter.simpledialog import askstring  # récupéaration nom pour sauvergarde
import os  # Pour trouver le répertoire courant (os.getcwd)
from tkinter.ttk import *  # meilleurs widgets
import tkinter.messagebox as msgbox  # Messages d'information ou d'avertissement


def IntValidate(text):
    """
    vérifie si l'entrée est une string convertible en integer ou vide
    """
    if text == "":
        return True
    try:
        int(text)+1
    except Exception as e:
        # print(e)
        return False
    return True


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
        self.IntValid = self.register(IntValidate)

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
            else:
                classe = xlsReader.xlsData
            self.OpenButton.pack()
            # Ajout des fonctions dans la fenêtre
            for func in [method for method in dir(classe) if method[0] != "_" and method != "SaveFile"]:
                self.funcs.append(Radiobutton(
                    self, text=func, variable=funcchoice, value=func))
                self.funcs[-1].pack(anchor="w", padx=10)
                funcchoice.set(func)

        self.geometry("{}x{}".format(550, 400))
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
                    sheetname = askstring(
                        "Nom de feuille/sheet", "Nom de la feuille de tableur:")
                    try:
                        xls = xlsWriter.xlsWriter(SheetName=sheetname)
                        if self.formatage != "dict":  # "Conversion" résultat en dictionnaire pour xlsWriter
                            dictretval = self.xls.Lecture(
                                self.Rowstart, self.Rowstop, self.Colstart, self.Colstop, "dict")
                        else:
                            dictretval = self.retval
                        xls.AddData(dictretval, self.Colstart,
                                    self.Rowstart, autoSave=(True, filename))
                        msgbox.showinfo(
                            "Sauvergarde réussie", f"le résultat a été sauvegardé avec succès sous le nom {filename}")
                    except Exception as e:
                        print(e)
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
            except Exception as e:
                print(e)
                msgbox.showwarning(
                    "Ligne de départ invalide", "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            try:
                self.Rowstop = int(self.Rowstop.get())
            except Exception as e:
                # print(e)
                self.Rowstop = None
            try:
                self.Colstart = int(self.Colstart.get())
            except Exception as e:
                print(e)
                msgbox.showwarning("Colonne de départ invalide",
                                   "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            try:
                self.Colstop = int(self.Colstop.get())
            except Exception as e:
                print(e)
                msgbox.showwarning(
                    "Colonne de fin invalide", "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            self.formatage = self.formatage.get()
            # affichage paramètres console (debug)
            print("=======================================")
            print("Paramètres :")
            print(f"\tLigne de départ : {self.Rowstart}")
            print(f"\tLigne de fin : {self.Rowstop}")
            print(f"\tColonne des clés : {self.Colstart}")
            print(f"\tColonnes de donnée : {self.Colstop}")
            print(f"\tFormatage : {self.formatage}")
            print("=======================================")
            # appel fonction
            if self.fonction == xlsReader.xlsData.Lecture.__name__:
                try:
                    self.retval = self.xls.Lecture(
                        self.Rowstart, self.Rowstop, self.Colstart, self.Colstop, self.formatage)
                    print(f"Résultat : {self.retval}")
                    WinRetFunc()  # Affichage fenêtre finale (résultat)
                except Exception as e:
                    print(e)
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

            self.ClearWindow()  # nettoyage fenêtre
            self.title(f"{self.fonction} : récupération des arguments")
            # Widgets de récupération
            # Valeurs numériques (zone de lecture)
            Label(self, text="Ligne de départ :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Rowstart, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Ligne de fin (laisser vide pour ligne maximale) :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Rowstop, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Colonne de départ :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Colstart, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Colonne de fin :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Colstop, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
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
                except Exception as e:  # erreur init
                    print(e)
                    msgbox.showerror("Erreur initialisation classe",
                                     "la classe n'a pas pu être initialisée, veuillez réessayer")
                    self.destroy()
                print("Classe initialisée")
            else:  # feuille non spécifiée
                msgbox.showinfo("Feuille indéfinie",
                                "La feuille à lire n'a pas été spécifiée")

            WinArgs()  # Récupération arguments (2e fenêtre)

        # Placement widgets args initialisation
        feuilles = xlsReader.xlsData._GetSheets(self.FilePath)
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

    def WinXlsPlot(self):
        """
        Fenêtre pour le module d'xlsPlot
        """
        self.SortCol = None
        self.ttlwidgets = []
        self.sortwidgets = []
        ttlvalue = IntVar()
        ttlvalue.set(1)
        tx = StringVar()
        ty = StringVar()
        tx.set("0")
        ty.set("0")
        ttlvar = StringVar()
        TypeAffichage = BooleanVar()

        def WinRetFunc():
            """
            Permet de choisir entre l'affichage du plot et la sauvegarde de la figure sous format .png
            """
            def Affichage():
                """
                Action après confirmation du choix d'affichage/de sauvegarde (appui du bouton)

                TODO : Ajouter titleOffset à choisir (pour l'instant par défaut 2)
                """
                filename = ""
                if TypeAffichage.get():  # Sauvegarde demandée
                    path = fldialog.asksaveasfilename(initialdir=os.getcwd(), title="Sauvegarde résultat", filetypes=(
                        ("Image", "*.png"), ("all files", "*.*")), defaultextension=".png", initialfile="plot")
                    filename = path.split("/")[-1][:-4]
                    if filename=="":
                        msgbox.showwarning("Fichier non choisi", "Le choix de fichier a été annulé")
                # Exécution fonction demandée
                try:
                    if self.fonction == "DiagrammeMultiBarres":
                        self.xls.DiagrammeMultiBarres(self.SortOptions, self.DataCols, self.KeyCol,
                                                        self.Start, self.Stop, self.ttlOffset, 
                                                        PlotSave=(TypeAffichage.get(), filename))
                    elif self.fonction == "DiagrammeMultiCirculaire":
                        self.xls.DiagrammeMultiCirculaire(self.SortOptions, self.DataCols, self.KeyCol,
                                                        self.Start, self.Stop, self.ttlOffset, 
                                                        PlotSave=(TypeAffichage.get(), filename))
                except Exception as e:
                    print(e)
                    msgbox.showerror(
                        "Affichage impossible", "Une erreur à été rencontrée lors de l'exécution de la fonction\nVeuillez réessayer")

                if TypeAffichage.get():  # Sauvegarde
                    msgbox.showinfo("Sauvegarde terminée", f"Le plot a été sauvegardé sous le nom {filename}")
                
                WinRetFunc() # Retour à la fenêtre précédente (choix de récupération)

            self.ClearWindow()  # Efface la fenêtre précédente
            self.title("Retour du résultat de la fonction")
            self.geometry("{}x{}".format(550, 200))
            Label(self, text="Veuillez choisir un moyen de récupérer le plot/graphique :",
                  font=self.titlefont).pack(padx=10, pady=5, anchor=CENTER)
            Label(self, text="(Vous pourrez retourner à cette fenêtre après avoir fait un choix)",
                  font=self.font).pack(padx=10, anchor=CENTER)
            # Demande la facon de récupérer le plot
            Radiobutton(self, text="Affichage direct",
                        command=lambda: ExitButton.configure(state="normal"),
                        variable=TypeAffichage, value=False).pack(padx=10, pady=5, anchor=CENTER)
            Radiobutton(self, text="Sauvegarde image du plot",
                        command=lambda: ExitButton.configure(state="normal"),
                        variable=TypeAffichage, value=True).pack(padx=10, pady=5, anchor=CENTER)
            # Bouton de confirmation
            ExitButton = Button(self, text="Confirmer",
                                command=Affichage, width=20, state="disabled")
            ExitButton.place(x=140, y=150)
            # Bouton pour quitter l'interface
            Button(self, text="Quitter", command=exit,
                   width=20).place(x=280, y=150)

        def ConfirmationArgs():
            """
            Action lors de la confirmation des paramètres
            """
            # Title offset
            try:
                self.ttlOffset = int(self.ttlOffset.get())
            except Exception as e:
                print(e)
                msgbox.showwarning(
                    "Offset titre invalide", "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            # ligne de départ
            try:
                self.Start = int(self.Start.get())
            except Exception as e:
                print(e)
                msgbox.showwarning(
                    "Ligne de départ invalide", "La valeur indiquée est invalide (laissé vide ?)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            # ligne de fin
            try:
                self.Stop = int(self.Stop.get())
            except Exception as e:
                # print(e)
                self.Stop = None
            # colonne des clés
            try:
                assert int(self.KeyColbox.get(
                )) not in self.DataCols, "Erreur : Colonne de clé dans les colonnes de données"
                self.KeyCol = int(self.KeyColbox.get())
            except Exception as e:
                print(e)
                msgbox.showwarning(
                    "Colonne des clés invalide", "La valeur indiquée est invalide (laissée vide ou déjà une colonne de donnée)")
                WinArgs()  # réinitialisation fenêtre
                return 0
            # colonnes de données
            if not self.DataCols:  # pas de colonnes de données choisies
                msgbox.showwarning(
                    "Colonnes de données invalides", "Aucune colonne de donnée n'a été chosie")
                WinArgs()  # réinitialisation fenêtre
                return 0
            # Options de Tri
            self.Sort = self.Sort.get()
            if self.Sort:  # Tri demandé
                try:
                    assert int(self.SortCol.get()) >= 0 and int(
                        self.SortCol.get()) in self.DataCols, "Colonne de tri invalide"
                    self.SortOptions = (
                        self.Sort, self.SortingType.get(), int(self.SortCol.get()))
                except Exception as e:
                    print(e)
                    msgbox.showwarning("Colonne de Tri invalide",
                                       "L'index de la colonne de tri est invalide (non spécifié ou inférieur à 0)")
            else:  # Pas de tri demandé
                # création paramètre qui sera ignoré car self.Sort == False
                self.SortOptions = (self.Sort, False, 0)
            # Affichage console des paramètres (debug)
            print("=======================================")
            print("Paramètres :")
            print(f"\tLigne de départ : {self.Start}")
            print(f"\tLigne de fin : {self.Stop}")
            print(f"\tColonne des clés : {self.KeyCol}")
            print(f"\tColonnes de donnée : {self.DataCols}")
            print(f"\tOptions de Tri : {self.SortOptions}")
            print("=======================================")
            # Appel fenêtre de retour/d'affichage
            WinRetFunc()

        def WinArgs():
            """
            Deuxième partie de la fenêtre, permet de récupérer les paramètres nécessaires à la fonction choisie
            """
            # Rows
            self.ttlOffset = StringVar()
            self.ttlOffset.set("1")
            self.Start = StringVar()
            self.Start.set("1")
            self.Stop = StringVar()
            # Sorting options
            self.Sort = BooleanVar()
            self.SortOrder = BooleanVar()
            self.ColSort = StringVar()
            self.ColSort.set("1")
            # DataColumns
            self.DataCols = []

            def ConfirmDataCol():
                """
                Action de sauvegarde de la colonne sélectionnée après l'appui du bouton
                """
                # Vérifs Index valide
                if IntValidate(DataColbox.get()):
                    if DataColbox.get() == self.KeyColbox.get():
                        msgbox.showwarning(
                            "Colonne de donnée invalide", "L'index choisi est le même que celui de la colonne de clé")
                        return 0
                    if int(DataColbox.get()) >= 0 and int(DataColbox.get()) not in self.DataCols:
                        self.DataCols.append(int(DataColbox.get()))
                        ApercuDataCol['text'] = f"Apreçu colonnes : {str(self.DataCols)}"
                        if self.SortCol != None:
                            self.SortCol['values'] = self.DataCols
                        print(
                            f"Colonne de donnée {int(DataColbox.get())} ajoutée")
                    else:
                        msgbox.showwarning(
                            "Colonne de donnée invalide", "L'index est invalide (inférieur à 0 ou déjà choisi)")
                else:
                    msgbox.showwarning(
                        "Colonne de donnée invalide", "L'index est invalide (pas int)")
                DataColbox.set("1")
                # Dévérouillage options de tri
                SortCheck['text'] = "Tri des valeurs"
                SortCheck['state'] = "normal"

            def SortingOptions():
                """
                Action d'affichage, ou non, des options de tri des données
                """
                ExitButton['state'] = "normal"
                if self.Sort.get():  # Affichage des options de tri
                    self.SortingType = BooleanVar()
                    # ajout widgets
                    # Choix croissant ou décroissant
                    self.sortwidgets.append(
                        Label(self, text="Type de tri :", font=self.font))
                    self.sortwidgets[-1].pack(padx=10, anchor="w", pady=10)

                    self.sortwidgets.append(
                        Radiobutton(self, text="Tri croissant",
                                    variable=self.SortingType, value=False))
                    self.sortwidgets[-1].pack(anchor="w", padx=30)

                    self.sortwidgets.append(
                        Radiobutton(self, text="Tri décroissant",
                                    variable=self.SortingType, value=True))
                    self.sortwidgets[-1].pack(anchor="w", padx=30)
                    # Choix de l'index de la colonne de tri
                    self.sortwidgets.append(
                        Label(self, text="Colonne de tri (choix dans la liste des colonnes de donnée) :",
                              font=self.font))
                    self.sortwidgets[-1].pack(padx=10, anchor="w", pady=10)
                    self.SortCol = Combobox(self, justify=LEFT, values=self.DataCols,
                                            state="readonly")
                    self.SortCol.pack(pady=5, padx=10, anchor="w")
                    self.sortwidgets.append(self.SortCol)
                else:  # suppression des options de tri
                    for f in self.sortwidgets:
                        f.destroy()
                    self.sortwidgets = []

            self.ClearWindow()  # nettoyage fenêtre
            self.geometry("{}x{}".format(600, 650))
            self.title(f"{self.fonction} : récupération des arguments")
            # Widgets de récupération
            # Récupération intervalle des lignes
            Label(self, text="Ligne de départ :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Start, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            Label(self, text="Ligne de fin (laisser vide pour ligne maximale) :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.Stop, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            # Récupération Offset des titres
            Label(self, text="Offset des titres (écart par rapport à la ligne de départ) :").pack(
                pady=5, padx=10, anchor="w")
            Entry(self, textvariable=self.ttlOffset, justify=LEFT, validate="key",
                  validatecommand=(self.IntValid, "%P")).pack(pady=5, padx=10, anchor="w")
            # Récupération colonnes à exploiter
            # Colonne des clés
            Label(self, text="Colonne des clés :").pack(
                pady=5, padx=10, anchor="w")
            self.KeyColbox = Spinbox(self, justify=LEFT, validate="key",
                                     validatecommand=(self.IntValid, "%P"), from_=.0, to=10000)
            self.KeyColbox.set("0")
            self.KeyColbox.pack(pady=5, padx=10, anchor="w")
            # Colonnes de donnée
            Label(self, text="Colonne de données \n(entrer une colonne, puis appuyer sur Valider : recommencer autant de fois que nécessaire) :").pack(
                pady=5, padx=10, anchor="w")
            DataColbox = Spinbox(self, justify=LEFT, validate="key",
                                 validatecommand=(self.IntValid, "%P"), from_=.0, to=10000)
            DataColbox.set("1")
            DataColbox.pack(pady=5, padx=10, anchor="w")
            # Bouton de confirmation
            Button(self,
                   text="Valider colonne", command=ConfirmDataCol, width=20, state="normal"
                   ).pack(padx=10, pady=5, anchor="w")
            ApercuDataCol = Label(
                self, text=f"Apreçu colonnes : {str(self.DataCols)}")
            ApercuDataCol.pack(padx=10, pady=5, anchor="w")
            # Options de tri
            SortCheck = Checkbutton(self, text="Tri des valeurs (entrer au moins une colonne de donnée avant) :",
                                    variable=self.Sort, onvalue=True, offvalue=False, command=SortingOptions, state="disabled")
            SortCheck.pack(anchor="w", padx=30)
            # Bouton de confirmation
            ExitButton = Button(
                self, text="Confirmer", command=ConfirmationArgs, width=20, state="disabled")
            ExitButton.place(x=240, y=600)

        def ConfirmationInit():
            """
            Action lors de la confirmation des paramètres pour l'initialisation
            """
            if ttlvalue.get() == 2:  # vérification et récupération des coordonnées
                if IntValidate(tx.get()) and IntValidate(ty.get()):  # Vérif coords valides
                    if int(tx.get()) >= 0 and int(ty.get()) >= 0:
                        self.TitleCell = (int(tx.get()), int(ty.get()))
                    else:  # si erreur, réinitialisation
                        msgbox.showwarning("Coords titre invalides",
                                           "Les coordonnées du titre ne sont pas valides, veuillez réessayer")
                        self.WinXlsPlot()
                        return 0
            else:  # récupération du titre
                self.TitleCell = ttlvar.get()

            if sheetChoice.get() in feuilles:  # feuille choisie
                # Récupération paramètres
                self.feuille = sheetChoice.get()
                print("feuille choisie :", self.feuille)
                # initialisation de la classe
                try:
                    self.xls = xlsPlot.xlsDB(
                        fullPath=self.FilePath, sheet=feuilles.index(self.feuille), TitleCell=self.TitleCell)
                    print(f"Titre graphique : {self.xls.Title}")
                except Exception as e:  # erreur init
                    print(e)
                    msgbox.showerror("Erreur initialisation classe",
                                     "la classe n'a pas pu être initialisée, veuillez réessayer")
                    self.destroy()
                print("Classe initialisée")
            else:  # feuille non spécifiée
                msgbox.showwarning("Feuille indéfinie",
                                   "La feuille à lire n'a pas été spécifiée")

            WinArgs()  # Récupération arguments (2e fenêtre)

        def TitleChoice():
            """
            Action lors du choix d'une manière de choisir le titre du graphique
            Affiche les entrées nécessaires et supprime les anciennes
            """
            ExitButton['state'] = "normal"
            if len(self.ttlwidgets) != 0:
                for f in self.ttlwidgets:
                    f.destroy()
                self.ttlwidgets = []
            if ttlvalue.get() == 1:  # Choix direct du titre (str)
                # ajout widgets
                self.ttlwidgets.append(
                    Label(self, text="Titre :", font=self.font))
                self.ttlwidgets.append(
                    Entry(self, textvariable=ttlvar, justify=LEFT))
            else:  # donc 2 : Choix d'une cellule (tuple)
                # Demande les coordonnées de la cellule contenant le titre
                self.ttlwidgets.append(
                    Label(self, text="Placement titre :\ncoord x :", font=self.font))
                self.ttlwidgets.append(
                    Entry(self, textvariable=tx, justify=LEFT, validate="key",
                          validatecommand=(self.IntValid, "%P")))
                self.ttlwidgets.append(
                    Label(self, text="coord y :", font=self.font))
                self.ttlwidgets.append(
                    Entry(self, textvariable=ty, justify=LEFT, validate="key",
                          validatecommand=(self.IntValid, "%P")))

            for w in self.ttlwidgets:  # boucle affichage
                w.pack(pady=5, padx=10, anchor="w")

        self.title("xlsReader : paramètres")
        # Placement widgets args initialisation
        feuilles = xlsReader.xlsData._GetSheets(self.FilePath)
        sheetChoice = StringVar()
        sheetChoice.set("")

        Label(self, text="Feuille à lire (sheet) :",
              font=self.font).pack(padx=10, anchor="w")
        # Donne la liste des feuilles du fichier, permettant à l'utilisateur d'en choisir une
        Combobox(self, values=feuilles, width=max(
            [len(f) for f in feuilles]), state="readonly", textvariable=sheetChoice, height=30
        ).pack(padx=20, anchor="w")
        # Demande la facon de choisir un titre pour le graphique
        Radiobutton(self, text="Entrée d'un titre personnalisé", command=TitleChoice,
                    variable=ttlvalue, value=1).pack(anchor="w", padx=10, pady=10)
        Radiobutton(self, text="Choix d'une cellule contenant le titre", command=TitleChoice,
                    variable=ttlvalue, value=2).pack(anchor="w", padx=10)
        # Bouton de confirmation
        ExitButton = Button(self, text="Confirmer", state="disabled",
                            command=ConfirmationInit, width=20)
        ExitButton.place(x=210, y=350)

    def ClearWindow(self):
        """
        Efface tous les widgets de la fenêtre
        """
        for w in self.winfo_children():
            w.destroy()
        self.pack_propagate(0)


def main():
    win = window()
    win.ModuleChoice()  # fenêtre initiale
    win.mainloop()


if __name__ == '__main__':  # Exécution
    main()
