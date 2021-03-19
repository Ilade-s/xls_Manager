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
    - GrapheBarres : utlisant une seule colonne, va créer une graphique en barres
    - func2
"""

import xlrd # Module de gestion mère xls
import matplotlib # Création de graphiques
import sys # Messages d'erreur

class xlsDB:
    def __init__(self, fileName="pop-16ans-dipl6817"):
        """
        Initialisation de la base de données xls (ouverture et extraction)
        
        PARAMETRES :
        --------
        fileName : str
            nom du fichier xls à ouvrir
                default = "pop-16ans-dipl6817"
        """
        pass

    def DiagrammeBarres(self):
        """
        
        """
        pass

    def GrapheAxes(self):
        """
        
        """
        pass
    
    def DiagrammeCirculaire(self):
        """
        
        """
        pass

# Tests des fonctions
if __name__=='__main__':
    pass