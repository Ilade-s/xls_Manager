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
"""
import xlrd # Module de gestion mère xls

class xlsData:
    def __init__(self, sheet=10, fileName="pop-16ans-dipl6817", TitleCell=(0,0)):
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
        
        - TitleCell : tuple(int,int)
            - coordonnées de la cellule contenant le titre de la feuille souhaité
            - default = (0,0)
        """
        # Vérification paramètres
        for i in TitleCell:
            assert i >= 0
        assert sheet >= 0

        # Ouverture fichier xls
        with xlrd.open_workbook("./"+fileName+".xls", on_demand=True) as file: 
            self.Data = file.get_sheet(sheet)

        # Extraction titre feuille
        (rowx, columnx) = TitleCell
        self.Title = self.Data.cell_value(rowx,columnx)

    def Lecture(self,rowstart=0,rowstop=0,colstart=0,colstop=0):
        """
        Lit le fichier xls, puis renvoie les données en matrice

        PARAMETRES :
            - rowstart : int
                - ligne de départ 
        """
        pass