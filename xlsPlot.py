"""
xlsPlot
-------
Module de création de plots (matplotlib) à partir de données d'un fichier xls
----------
Module pouvant être utlisé dans d'autres programmes, utilisant matplotlib afin de créer des graphiques sur les données d'un fichier xls, lu avec le module xlrd

MODULES UTILISES (A INSTALLER) :
----------
    - xlrd (lecture de fichier xls)
    - matplotlib (graphiques)

UTILISATION :
----------
    La classe, quand initialisée, ouvre le fichier xls, puis peut exploiter toutes les fonctions

FONCTIONS :
----------
    - DiagrammeBarres : Utlisant une seule colonne, va créer une graphique en barres
    - GrapheAxes : utlisant deux colonnes (x et y) créé un graphique y(x)
    - DiagrammeCirculaire : Utlisant une seule colonne, permet de comparer leur part dans la somme avec un camembert
"""

import xlrd # Module de gestion mère xls
import matplotlib # Création de graphiques
import sys # Messages d'erreur

class xlsDB:
    def __init__(self, sheet=10, fileName="pop-16ans-dipl6817"):
        """
        Initialisation de la base de données xls (ouverture et extraction)
        
        PARAMETRES :
        --------
        sheet : int
            Index de la feuille de tableur à extraire
                default = 10 (11-1)
        
        fileName : str
            nom du fichier xls à ouvrir
                default = "pop-16ans-dipl6817"
        """
        # Ouverture fichier xls
        with xlrd.open_workbook(fileName+".xls", on_demand=True) as file: 
            self.Data = file.get_sheet(sheet)

        # Extraction titre feuille
        self.Title = self.Data.cell_value(0,0)
        print("Test :",self.Title)

    def DiagrammeBarres(self, DataColumn=3, KeyColumn=2, Start=24, Stop="auto"):
        """
        Permet de créer des diagrammes en barres pour comparer les éléments d'une seule colonne

        
        PARAMETRES :
        --------
        Commencent tous à 0
        --------
        DataColumn : int
            index de la colonne contenant les valeurs à comparer
                default = 3

        KeyColumn : int
            index de la colonne contenant les clés (noms) liées aux données
                default = 2
        
        Start : int
            index de la ligne de départ (inclue) des éléments à étudier
                default = 24
        
        Stop : int || str
            index de la dernière ligne (exclue) des éléments à étudier ou "auto" pour exploiter toutes les données (après start)
                default = "auto"

        SORTIE :
        --------
        ExitCode : int
            0 : Erreur lors de l'exécution
            1 : Exécution réussie
        """
        if Stop=="auto":   
            DataList = self.Data.col_values(DataColumn, Start)
        else:
            KeyList = self.Data.col_values(KeyColumn, Start, Stop)

    def GrapheAxes(self):
        """
        ...
        
        PARAMETRES :
        --------
        key : str
            clé à rechercher

        dataID : str
            nom de colonne de la donnée souhaitée liée à la clé recherchée (si trouvée)

        SORTIE :
        --------
        data : int ou str
            donnée liée, integer si possible, sinon en string
            (renvoie 0 si clée non trouvée ou si donnée non trouvée)
        """

        pass
    
    def DiagrammeCirculaire(self):
        """
        ...
        
        PARAMETRES :
        --------
        key : str
            clé à rechercher

        dataID : str
            nom de colonne de la donnée souhaitée liée à la clé recherchée (si trouvée)

        SORTIE :
        --------
        data : int ou str
            donnée liée, integer si possible, sinon en string
            (renvoie 0 si clée non trouvée ou si donnée non trouvée)
        """

        pass

# Tests des fonctions
if __name__=='__main__':
    # feuille = int(input("feuille à ouvrir : "))
    # xls = xlsDB(feuille)

    xls = xlsDB()