"""
DataExtractor
---------------
Ce module secondaire permet d'écrire des données sous forme de dictionnaire de colonnes.
Celles-ci sont ensuite sauvegardées dans un tableur au format xls, qui sera exploitable par xlsPlot.py
Ce module sera directement inclus dans xlsPlot.py dans un second temps.

FONCTIONNEMENT :
---------------
    La liste en paramètre peut contenir des clés ou non : si trouvées, seront mises dans la colonne 0 du tableur, sinon les indexs seront mis à leur place

"""
import xlwt # écriture de fichier xls

class xlsWriter:
    def __init__(self,FileName=""):
        """
        Quand appelée, créé un Workbook object avec une feuille ("DataSheet") qui pourra ensuite être modifié puis sauvegardé
        
        PARAMETRE :
            - FileName : str
                - Nom du fichier existant à éditer, si vide, crééra un nouvel objet pour édition.
                - SANS EXTENSION DE FICHIER
                - Default = "" (vide)
        """
        if FileName=="": # Nouveau fichier
            self.NewFile = True
            self.File = xlwt.Workbook() # création tableur
            self.Sheet = self.File.add_sheet("DataSheet",True) # ajout d'une feuille
        else: # Fichier existant
            self.FileToEdit = open(FileName+".xls","w")
            self.NewFile = False

    def AddData(self,Data,ColStart=0,RowStart=0,KeysCol=None,ColsOrder=None):
        """
        Ajoute les données en paramètre Data à la feuille instancée dans __init__

        PARAMETRES :
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
            - ColsOrder : None || list[keys] (optionnel)
                - Liste des clés de colonnes dans l'ordre souhaité
                - Si None, les colonnes seront entrées arbitrairement (par Data.keys())

        SORTIE :
        -------------
        (indirectement) Les données sont ajoutées à la feuille, qui peut ensuite être sauvegardée
        """
        # Vérification paramètres
        assert ColStart >= 0, "Colonne de départ invalide"
        assert RowStart >= 0, "Ligne de départ invalide"
        assert KeysCol in Data.keys() or KeysCol==None, "Paramètres KeysCol invalide"
        if ColsOrder!=None:
            for c in ColsOrder:
                assert ColsOrder[c] in Data.keys(), "ColsOrder invalide"
            ColumnKeys = ColsOrder
        else:
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
        DataColumns = [data for data in Data.values()]
        # Debug
        print(ColumnKeys)
        print(KeyColumn)
        print(DataColumns)
        # Ajout clés au Workbook
        self.Sheet.write(RowStart,ColStart,label=KeyCol)
        for nrow in range(lenData):
            self.Sheet.write(RowStart+nrow+1,ColStart,label=KeyColumn[nrow])
        # Ajout données au Workbook
        for ncol in range(len(ColumnKeys)):
            self.Sheet.write(RowStart,ColStart+ncol+1,label=ColumnKeys[ncol])
            for nrow in range(lenData):
                self.Sheet.write(RowStart+nrow+1,ColStart+ncol+1,label=DataColumns[ncol][nrow])


    def SaveFile(self,FileName="ExtractedData"):
        """
        Sauvegarde le fichier xls

        PARAMETRES :
        -------------
            - FileName : str
                - Nom du fichier à enregistrer
                - default = "ExtractedData" (feuille/sheet "DataSheet")
        """
        FileName = str(FileName) # Conversion en string (si int ou float)
        self.File.save(FileName+".xls") # Sauvegarde

if __name__=="__main__": # test
    data = {"keys":[chr(65+i) for i in range(10)],"data":[i for i in range(10)],"d":[i for i in range(10)],"da":[i for i in range(10)]}
    print(data)
    xls = xlsWriter() # création workbook
    xls.AddData(data, KeysCol="keys")
    xls.SaveFile() # sauvegarde xls
