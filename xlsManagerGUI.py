"""
xlsManagerGUI (interface graphique des trois modules)
------------------

MODULES UTILISABLES : 
    - xlsPlot : création de graphiques à partir d'un fichier
    - xlsWriter : édition de fichier xls
    - xlsReader : lecture de fichier xls

FONCTIONNEMENT :
    - Tout se passe dans l'interface graphique (ni console, ni python)
    - Le programme fonctionne en plusieurs étapes (séparées en fenêtres) :
        - 1 : Choix du module à utiliser
        - 2 : Choix du fichier à utiliser
        - 3 : Choix de la fonction à utliser
        - 4 : Entrée des paramètres nécessaires
        - 5 : Affichage du résultat (dépend de la fonction utilisée)
        - 6 : Demandes éventuelles (sauvegarde...)

MODULES UTILISES : (en plus des trois modules)
    - tkinter (interface graphique)
    - matplotlib (graphiques)
    - pandas
    - numpy
    - xlrd, xlwd et xlutils (gestion de fichiers xls)
"""

import xlsPlot # Création de graphiques
import xlsReader # Edtition de fichiers xls
import xlsWriter # Lecture de fichier xls