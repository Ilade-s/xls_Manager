"""
xlsReader (contient la classe xlsData avec initialisation et fonctions)
-------
Module de lecture un fichier xls
----------
- Module pouvant être utlisé dans d'autres programmes

- Si lancé en main, proposera de lancer un test de chaque fonction

UTILISATION :
----------
    La classe, quand initialisée, ouvre le fichier xls, puis peut lire le fichier

FONCTIONS :
----------
    - Lecture : Lit le fichier xls, puis renvoie les données en matrice
    - _GetSheets : permet d'obtenir la liste des feuilles du fichier ciblé par le chemin
        - la classe xlsData n'a pas besoin d'être initialisée pour son utilisation
        - méthode normalement utilisée uniquement pour l'affichage des feuilles disponibles (GUI)
"""
import xlrd # Module de gestion mère xls

class xlsData:
    def __init__(self, sheet=10, fileName="pop-16ans-dipl6817", fullPath=""):
        """
        Initialisation de la base de données xls (ouverture et extraction)
        
        PARAMETRES :
        --------
        - sheet : int
            - Index de la feuille de tableur à extraire
            - default = 10 (11-1)
        
        - fileName : str
            - nom du fichier xls à ouvrir
            - default = "pop-16ans-dipl6817"
        
        - fullPath : str
            - si différent de "", remplace fileName pour l'ouverture de fichier
            - default = "" (désactivé)
        """
        # Vérification paramètres
        assert sheet >= 0

        # Ouverture fichier xls
        if fullPath=="":
            with xlrd.open_workbook("./"+fileName+".xls", on_demand=True) as file: 
                self.Data = file.get_sheet(sheet)
        else:
            with xlrd.open_workbook(fullPath, on_demand=True) as file: 
                self.Data = file.get_sheet(sheet)

    def Lecture(self,rowstart=13,rowstop=None,colstart=2,colstop=3,formatage="colmat"):
        """
        Lit le fichier xls, puis renvoie les données en matrice

        PARAMETRES :
        -----------
        Les index commencent tous à 0
        -------------
            - rowstart : int (incluse)
                - ligne de départ (coord x)
                - default = 0
            - rowstop : int || None (incluse)
                - ligne de fin (coord x)
                - default = 0
            - colstart : int (incluse)
                - colonne de départ (coord y)
                - default = 0
            - colstop : int (incluse)
                - colonne de fin (coord y)
                - default = 0
            - formattage : str
                - "colmat" : format cols[col[rows],...]
                - "rowmat" : format rows[row[col],...]
                - "dict" : format cols{col[0]:[col[1:]],...}
        
        SORTIE : 
        -----------
            - MatData : list[list[any]] || cols{col[0]:[col[1:]],...}
                - Matrice contenant les données 
                - format selon le paramètre "format"


        """
        # Vérification des paramètres
        assert rowstart>=0, "ligne de départ invalide (rowstart)"
        assert colstart>=0, "colonne de départ invalide (colstart)"
        assert rowstop==None or rowstop>=0, "ligne de fin invalide (rowstop)"
        assert colstop>=0, "colonne de fin invalide (colstop)"
        assert formatage=="colmat" or formatage=="rowmat" or formatage=="dict", "formatage invalide"

        # Extraction des données en matrice des colonnes
        MatData = [self.Data.col_values(col, rowstart, rowstop) for col in range(colstart,colstop+1)]
        # Conversion des données en matrice des lignes
        if formatage=="rowmat":
            MatData = [[col[i] for col in MatData] for i in range(len(MatData[0]))]
        # Extraction en dictionnaire
        if formatage=="dict":
            MatData = {col[0]:col[1:] for col in MatData}

        # Renvoi de la matrice
        return MatData
    
    def _GetSheets(FilePath):
        """
        Permet de récupérer la liste des feuilles
        Ne sera pas affichée dans la GUI (méthode utilisée uniquement pour affichage) 
        Sera peut être ajoutée dans la GUI si nécessaire/utile

        PARAMETRE :
            - FilePath : str
                - chemin complet ou seulement le nom (si dans dossier courant/de travail)
        
        SORTIE :
            - sheet names : list[str]
                - liste des noms des feuilles de tableur contenue dans le fichier
        """
        with xlrd.open_workbook(FilePath, on_demand=True) as file:
            return (file._sheet_names)

if __name__=='__main__': # Test
    # Lecture de ExtractedData.xls
    xls = xlsData()
    mat = xls.Lecture()
    xls = xlsData(0, "ExtractedData")

    mat = xls.Lecture(rowstart=0,colstart=0,colstop=3)
    print(mat)
    mat = xls.Lecture(rowstart=0,colstart=0,colstop=3,formatage="rowmat")
    print(mat)
    mat = xls.Lecture(rowstart=0,colstart=0,colstop=3,formatage="dict")
    print(mat)