"""
xlsWriter (édition de tableur xls)
---------------
Ce module secondaire permet d'écrire des données sous forme de dictionnaire de colonnes, ainsi que de supprimer des données d'un fichier xls.
Celles-ci sont ensuite sauvegardées dans un tableur au format xls, qui sera exploitable par xlsPlot.py

FONCTIONNEMENT :
---------------
    - 1 : Initialiser la classe xlsWriter permet de créer le Workbook (pour édition)
    - 2 : Ajout ou suppression de données :
        - La méthode AddData permet d'ajouter des colonnes de données au Workbook
        - la méthode DeleteData permet de supprimer une zone de données du tableur 
    - 3 : la méthode SaveFile permet enfin de sauvegarder le Workbook dans un fichier xls, avec nom personnalisable

Attention à bien lire les docstrings des méthodes afin de ne pas faire d'erreurs de paramètres (néanmoins, des "assert" sont présents pour protéger)
--------------

"""
import xlwt  # écriture de fichier xls
# lecture de fichier xls (pour édition de fichier preéxistant)
from xlrd import open_workbook
import xlutils.copy  # joint entre xlrd et xlwt
import copy  # Copie des données (AddData)


class xlsWriter:
    def __init__(self, FileName="", SheetName="DataSheet", fullPath=""):
        """
        Quand appelée, créé un Workbook object avec une feuille ("DataSheet") qui pourra ensuite être modifié puis sauvegardé

        PARAMETRE :
            - FileName : str
                - Nom du fichier existant à éditer, si vide, crééra un nouvel objet pour édition.
                - SANS EXTENSION DE FICHIER
                - Default = "" (vide)
            - SheetName : str
                - Nom de la feuille/sheet à éditer ou créer
                - Default = "DataSheet"
            - fullPath : str
                - si différent de "", remplace fileName pour l'ouverture de fichier
                - default = "" (désactivé)
        """
        if FileName != "" or fullPath != "":  # Fichier existant
            self.FileName = FileName
            if fullPath == "":
                FileReader = open_workbook(
                    "./"+FileName+'.xls', formatting_info=True, on_demand=True)
            else:
                FileReader = open_workbook(
                    fullPath, formatting_info=True, on_demand=True)
            self.File = xlutils.copy.copy(FileReader)
            try:
                self.Sheet = self.File.get_sheet(SheetName)
            except:
                self.Sheet = self.File.add_sheet(
                    SheetName, True)  # ajout d'une feuille
            self.NewFile = False
        else:  # Nouveau fichier
            self.FileName = "ExtractedData"
            self.NewFile = True
            self.File = xlwt.Workbook()  # création tableur
            self.Sheet = self.File.add_sheet(
                SheetName, True)  # ajout d'une feuille

    def AddData(self, data, ColStart=0, RowStart=0, KeysCol=None, Title=(None, 0, 0), autoSave=(False, None)):
        """
        Ajoute les données en paramètre Data à la feuille instancée dans __init__

        PARAMETRES :
        -------------
        Les index commencent tous à 0
        -------------
            - Data : dict{colName:[rows],...}
                - dictionnaire contenant les colonnes de données à ajouter 
                    - Seront ajoutées dans leur entiéreté
                    - La colonne de clé est renseignée par le paramètre "KeysCol"
            - ColStart : int
                - Colonne de départ pour écriture : si des clés sont données (Data), elles seront écrites sur cette colonne
                - Default = 0
            - RowStart : int
                - Ligne de départ des données à ajouter
                - Default = 0
            - KeysCol : None || str (optionnel)
                - Nom de la colonne (clé du dictionnaire Data) contenant les clés
                - Si None, le programme assume qu'il n'existe pas de clé (les données seront mises sans recherche de clés)
            - Title : tuple(str || None,int,int)
                - dans l'ordre :
                    - Title[0] : str || None : Titre de la feuille/sheet (si None, la paramètre sera ignoré)
                    - Title[1] : int : coord x (ligne/row)
                    - Title[2] : int : coord y (colonne/column)
            - AutoSave : tuple(bool, FileName: str || None)
                - AutoSave[0] : Si true, SaveFile sera immédiatement appellé (si false, AutoSave sera ignoré)
                - AutoSave[1] : FileName : paramètre de SaveFile(), se reporter à la docstring de SaveFile pour plus d'infos

        SORTIE :
        -------------
        (indirectement) Les données sont ajoutées à la feuille, qui peut ensuite être sauvegardée
        """
        Data = copy.deepcopy(data)
        # Vérification paramètres
        assert len(Title) == 3, "Paramètre Title invalide (format incorrect)"
        assert ColStart >= 0, "Colonne de départ invalide"
        assert RowStart >= 0, "Ligne de départ invalide"
        assert KeysCol in Data.keys() or KeysCol == None, "Paramètres KeysCol invalide"
        # Extraction clés de colonnes
        ColumnKeys = [key for key in Data.keys()]
        # Calcul nrows data
        for j in Data.values():
            lenData = len(j)
        # Extraction données dict en liste
        if KeysCol != None:
            KeyColumn = Data.pop(KeysCol)
            KeyCol = ColumnKeys.pop(ColumnKeys.index(KeysCol))
        DataColumns = [data for data in Data.values()]
        # Debug
        # print(ColumnKeys)
        # print(KeyColumn)
        # print(DataColumns)
        KeyOffset = 0
        # Ajout clés au Workbook
        if KeysCol != None:
            KeyOffset += 1
            self.Sheet.write(RowStart, ColStart, label=KeyCol)
            for nrow in range(lenData):
                self.Sheet.write(RowStart+nrow+1, ColStart,
                                 label=KeyColumn[nrow])
        # Ajout données au Workbook
        for ncol in range(len(ColumnKeys)):
            self.Sheet.write(RowStart, ColStart+ncol +
                             KeyOffset, label=ColumnKeys[ncol])
            for nrow in range(lenData):
                self.Sheet.write(RowStart+nrow+1, ColStart +
                                 ncol+KeyOffset, label=DataColumns[ncol][nrow])
        # Ajout Title (si non None)
        if Title[0] != None:
            self.Sheet.write(Title[1], Title[2], label=Title[0])

        (Save, filename) = autoSave
        if Save:
            self.SaveFile(filename)

    def DeleteData(self, ColStart=0, RowStart=0, ColEnd=10, RowEnd=10, autoSave=(False, None)):
        """
        Permet de supprimer des données d'une feuille, selon une zone préétablie

        PARAMETRES :
        --------------
        Les indexs commencent tous à 0
        --------------
            - ColStart : int
                - Index de la colonne de départ, incluse (côté gauche)
                - Default = 0
            - RowStart : int
                - Index de la ligne de départ, incluse (côté haut)
                - Default = 0
            - ColEnd : int
                - Index de la dernière colonne, non incluse (coté droite)
                - Default = 10
            - RowEnd : int
                - Index de la dernière ligne, non incluse (côté bas)
                - Default = 10

        SORTIE : 
        -------------
        Aucune
        """
        # Vérification paramètres
        assert ColStart >= 0, "Colonne de départ invalide"
        assert RowStart >= 0, "Ligne de départ invalide"
        assert ColEnd-ColStart >= 1, "Dernière colonne invalide"
        assert RowEnd-RowStart >= 1, "Dernière ligne invalide"
        # Suppression de l'intervalle 2D
        for col in range(ColStart, ColEnd):
            for row in range(RowStart, RowEnd):
                self.Sheet.write(row, col, label=None)

        (Save, filename) = autoSave
        if Save:
            self.SaveFile(filename)

    def SaveFile(self, FileName=None):
        """
        Sauvegarde le fichier xls

        PARAMETRES :
        -------------
            - FileName : str || None
                - Nom du fichier à enregistrer
                - default = None (le nom sera : {OriginalFileName}_Edited.xls )
        """
        if self.NewFile:  # Création d'un nouveau fichier
            if FileName == None:
                self.File.save(self.FileName+".xls")  # Sauvegarde
            else:
                # Conversion en string (si int ou float)
                FileName = str(FileName)
                self.File.save(FileName+".xls")  # Sauvegarde
        else:  # Edition de fichier preéxistant
            if FileName == None:
                self.File.save(self.FileName+"_Edited"+".xls")  # Sauvegarde
            else:
                self.File.save(FileName+".xls")  # Sauvegarde


if __name__ == "__main__":  # test
    print("=============================================")
    print("Bienvenue dans mon programme/module d'édition de fichier xls")
    print("Vous pouvez lancer des test pour créer, dans l'ordre :")
    print("\t- ExtractedData : créée, contient data")
    print("\t- ExtractedData_Edited : A partir de ExtractedData, ajoute une feuille NewData")
    print("\t- Data_Del.xls : A partir de ExtractedData, supprime la dernière colonne de données")
    print("\tAttention, ces trois fichiers seront crées dans le dossier courant")
    print("=============================================")
    print("Appuyez sur entrée pour continuer...")
    input()

    data = {"keys": [chr(65+i) for i in range(10)], "data": [i for i in range(10)],
            "d": [i for i in range(10)], "da": [i for i in range(10)]}
    print(data)

    # Création du fichier ExtractedData.xls
    xls = xlsWriter()  # création workbook
    xls.AddData(data, KeysCol="keys", Title=("Données originales", 6, 6))
    xls.SaveFile()  # sauvegarde xls

    # Edition du fichier ExtractedData.xls
    xls = xlsWriter(FileName="ExtractedData",
                    SheetName="NewSheet")  # création workbook
    xls.AddData(data, KeysCol="keys", ColStart=2, Title=("EditedData", 0, 0))
    xls.SaveFile()  # sauvegarde xls

    # Suppression ligne données 4 de ExtractedData.xls sous le nom Data_Del.xls
    xls = xlsWriter(FileName="ExtractedData")
    xls.DeleteData(ColStart=3, ColEnd=4, RowEnd=11)
    xls.SaveFile(FileName="Data_Del")
