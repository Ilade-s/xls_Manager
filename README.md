# xls_Manager
xlsPlot (création de graphiques)
----------
(créateur de plots (matplotlib) à partir de données d'un fichier xls / édition et création de fichier xls) :

xlsWriter (édition de tableurs xls)
----------
Modules pouvant être utilisés dans d'autres programmes, utilisant matplotlib afin de créer des graphiques sur les données d'un fichier xls, lu avec le module xlrd, et xlwt/xlutils pour l'édition

xlsReader (Lecture de fichier xls)
----------
Permet de lire simplement les cases d'un tableur au format xls

MODULES UTILISES (A INSTALLER) :
----------
    - xlrd (lecture de fichier xls)
    - matplotlib (graphiques)
    - panda (DataFrame : pour données graphiques)
    - numpy (Calculs : graphique)
    - xlwt (écriture de fichier xls)
    - xlutils (lien entre xlrd et xlwt)
FONCTIONS :
----------
    - xlsPlot :
        - DiagrammeMultiBarres : Utlisant une colonne de clé, va créer un graphique en barres avec plusieurs colonnes de données
        - DiagrammeMultiCirculaire : Utlisant une ou plusieurs colonnes de données, permet de les comparer dans un ou plusieurs camembert (un pour chaque colonne de données)
    - xlsWriter :
        - AddData : ajout de colonnes de données et de clés en dictionnaire, à un fichier xls preéxistant ou nouveau
        - DeleteData : suppression d'un intervalle/zone de données d'un tableur preéxistant 
    - xlsReader :
        - Lecture : lecture d'un fichier xls, selon une zone de tableur et une feuille/sheet donnée
Un exemple d'importation et d'utilisation des modules est également disponible : ImportationExemple.py
-----------
____________________________________________
English Version :
-----------
xlsPlot.py (graphs creator)
-----------
Plot creator (matplotlib) from data in an xls file

xlsWriter.py (xls file edition)
----------
Program/Module that can be used in other programs, using matplotlib to create plots on data from an xls file, read with the xlrd module, and edited with xlwt and xlutils

xlsReader.py (xls file reader)
----------

MODULES USED (HAVE TO BE INSTALLED):
----------
    - xlrd (xls file reading)
    - matplotlib (graphics)
    - panda (DataFrame : for graphic data)
    - numpy (Calculations : graphics)
    - xlwt (xls file editing)
    - xlutils (link between xlrd and xlwt)
FUNCTIONS : 
----------
    - xlsPlot :
        - DiagrammeMultiBarres : Using a key column, will create a bar chart with several columns of data
        - DiagrammeMultiCirculaire : Using one or more columns of data, will compare them in one or more pie charts (one for each column of data)
    - xlsWriter :
        - AddData : add data columns and keys in dictionary, to a preexisting or new xls file
        - DeleteData : delete a range/data field from a pre-existing spreadsheet 
    - xlsReader :
        - Lecture : allow reading of a spectified zone of a sheet
