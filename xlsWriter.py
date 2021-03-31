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
import xlwt # écriture de fichier xls
from xlrd import open_workbook # lecture de fichier xls (pour édition de fichier preéxistant)
import xlutils.copy # joint entre xlrd et xlwt
import copy # Copie des données (AddData)

class xlsWriter:
    def __init__(self,FileName="",SheetName="DataSheet"):
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
        """
        if FileName=="": # Nouveau fichier
            self.NewFile = True
            self.File = xlwt.Workbook() # création tableur
            self.Sheet = self.File.add_sheet(SheetName,True) # ajout d'une feuille
        else: # Fichier existant
            FileReader = open_workbook(FileName+'.xls', formatting_info=True, on_demand=True)
            self.File = xlutils.copy.copy(FileReader)
            try:
                self.Sheet = self.File.get_sheet(SheetName)
            except:
                self.Sheet = self.File.add_sheet(SheetName,True) # ajout d'une feuille
            self.NewFile = False

    def AddData(self,data,ColStart=0,RowStart=0,KeysCol=None):
        """
        Ajoute les données en paramètre Data à la feuille instancée dans __init__

        PARAMETRES :
        -------------
        Les indexs commencent tous à 0
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
                - Si None, le programme assume qu'il n'existe pas de clé
        SORTIE :
        -------------
        (indirectement) Les données sont ajoutées à la feuille, qui peut ensuite être sauvegardée
        """
        Data = copy.deepcopy(data)
        # Vérification paramètres
        assert ColStart >= 0, "Colonne de départ invalide"
        assert RowStart >= 0, "Ligne de départ invalide"
        assert KeysCol in Data.keys() or KeysCol==None, "Paramètres KeysCol invalide"
        # Extraction clés de colonnes
        ColumnKeys = [key for key in Data.keys()]
        # Calcul nrows data
        for j in Data.values():
            lenData = len(j)
        # Extraction données dict en liste
        if KeysCol!=None:
            KeyColumn = Data.pop(KeysCol)
            KeyCol = ColumnKeys.pop(ColumnKeys.index(KeysCol))
        else:
            KeyColumn = [i for i in range(lenData)]
            KeyCol = "keys"
        DataColumns = [data for data in Data.values()]

        # Debug
        #print(ColumnKeys)
        #print(KeyColumn)
        #print(DataColumns)

        # Ajout clés au Workbook
        self.Sheet.write(RowStart,ColStart,label=KeyCol)
        for nrow in range(lenData):
            self.Sheet.write(RowStart+nrow+1,ColStart,label=KeyColumn[nrow])
        # Ajout données au Workbook
        for ncol in range(len(ColumnKeys)):
            self.Sheet.write(RowStart,ColStart+ncol+1,label=ColumnKeys[ncol])
            for nrow in range(lenData):
                self.Sheet.write(RowStart+nrow+1,ColStart+ncol+1,label=DataColumns[ncol][nrow])

    def DeleteData(self,ColStart=0,RowStart=0,ColEnd=10,RowEnd=10):
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
        assert ColStart>=0, "Colonne de départ invalide"
        assert RowStart>=0, "Ligne de départ invalide"
        assert ColEnd-ColStart>=1, "Dernière colonne invalide"
        assert RowEnd-RowStart>=1, "Dernière ligne invalide"
        # Suppression de l'intervalle 2D

    def SaveFile(self,FileName="ExtractedData"):
        """
        Sauvegarde le fichier xls

        Si un fichier preéxistant à été modifié, le fichier édité sera enregistré sous le nom {OriginalFileName}_Edited.xls"

        PARAMETRES :
        -------------
            - FileName : str (inutile si édition d'un fichier preéxistant)
                - Nom du fichier à enregistrer
                - default = "ExtractedData" (feuille/sheet "DataSheet")
        """
        if self.NewFile: # Création d'un nouveau fichier
            FileName = str(FileName) # Conversion en string (si int ou float)
            self.File.save(FileName+".xls") # Sauvegarde
        else: # Edition de fichier preéxistant
            self.File.save(FileName+"_Edited"+".xls") # Sauvegarde

if __name__=="__main__": # test
    data = {"keys":[chr(65+i) for i in range(10)],"data":[i for i in range(10)],"d":[i for i in range(10)],"da":[i for i in range(10)]}
    print(data)

    # Création du fichier ExtractedData.xls
    xls = xlsWriter() # création workbook
    xls.AddData(data, KeysCol="keys")
    xls.SaveFile() # sauvegarde xls

    # Edition du fichier ExtractedData.xls
    xls = xlsWriter(FileName="ExtractedData",SheetName="NewSheet") # création workbook
    xls.AddData(data, KeysCol="keys", ColStart=2)
    xls.SaveFile() # sauvegarde xls
