# xls_PlotCreator
Créateur de plots (matplotlib) à partir de données d'un fichier xls
----------
Module pouvant être utlisé dans d'autres programmes, utilisant matplotlib afin de créer des graphiques sur les données d'un fichier xls, lu avec le module xlrd
----------
MODULES UTILISES (A INSTALLER) :
----------
    - xlrd (lecture de fichier xls)
    - matplotlib (graphiques)
    - panda (DataFrame : pour données graphiques)
    - numpy (Calculs : graphique)
    (- sys : Messages d'erreur -included by default in python-)
FONCTIONS :
----------
    - DiagrammeBarres : Utlisant une colonne de clé, va créer un graphique en barres avec plusieurs colonnes de données
    - DiagrammeCirculaire : Utlisant une seule colonne, permet de comparer leur part dans la somme avec un camembert