"""
xlsPlot (contient la classe xlsDB avec initialisation et fonctions)
-------
Module de création de plots (matplotlib) à partir de données d'un fichier xls
----------
- Module pouvant être utlisé dans d'autres programmes, utilisant matplotlib afin de créer des graphiques sur les données d'un fichier xls, lu avec le module xlrd

- Si lancé en main, proposera de lancer un test de chaque fonction

MODULES UTILISES (A INSTALLER) :
----------
    - xlrd (lecture de fichier xls)
    - matplotlib (graphiques)
    - panda (DataFrame : graphique)
    - numpy (Calculs : graphique)
    (- sys : Messages d'erreur -included by default in python-)

UTILISATION :
----------
    La classe, quand initialisée, ouvre le fichier xls, puis peut exploiter toutes les fonctions

FONCTIONS :
----------
    - DiagrammeBarres : Utlisant une colonne de clé, va créer un graphique en barres avec plusieurs colonnes de données
    - DiagrammeCirculaire : Utlisant une seule colonne, permet de comparer leur part dans la somme avec un camembert
"""

import xlrd # Module de gestion mère xls
import matplotlib.pyplot as plt # Création de graphiques
import pandas as pd # Pour utilisation DataFrame (graphiques)
import sys # Messages d'erreur
import numpy as np # Calculs shares DiagrammeCirculaire

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

    def DiagrammeMultiBarres(self, SortedElements=(False, False, 0), DataColumns=[3], KeyColumn=2, Start=15, Stop=None, TitleOffset=2, figSize=(20.0,20.0)):
        """
        Permet de créer des diagrammes en barres pour comparer les éléments de une ou plusieurs colonnes de données

        
        PARAMETRES :
        --------
        Attention, cette focntion assume que le tableau est sous forme verticale et ne supportera pas les formes horizonales
        --------
        SortedElements : tuple(bool, bool, int)
            SortedElements[0] :
                Indique si les données doivent être triées ou non (ordre croissant)
                    default = False
            SortedElements[1] :
                Indique si les données doivent être triées en ordre croissant (False) ou décroissant (True)
                    default = False
            SortedElements[2] :
                Indique l'index de la colonne de données servant à trier les éléments (index dans DataColumns)
                    default = 0

        DataColumns : list[int]
            liste des index de colonnes contenant les valeurs à comparer
                default = [3]

        KeyColumn : int
            index de la colonne contenant les clés (noms) liées aux données
                default = 2
        
        Start : int
            index de la ligne de départ (inclue) des éléments à étudier
                default = 24
        
        Stop : int || None
            index de la dernière ligne (exclue) des éléments à étudier ou "auto" pour exploiter toutes les données (après start)
                default = "auto"

        TitleOffset : int
            Indique l'écart entre le Start et le titre (permet de trouver les titres d'axes)
                default = 2

        figSize : tuple(float, float)
            Indique la taille du diagramme (x, y)
                default (recommandé pour lecture) = (20.0,20.0)

        SORTIE :
        --------
        None
        """  
        # Vérification des paramètres
        for c in DataColumns:
            assert c!=KeyColumn, "Erreur : Les colonnes des données et des clés/noms sont les mêmes"
        assert SortedElements[2]<=len(DataColumns), "Erreur : l'index de la colonne choisie n'existe pas"
        assert Stop==None or Stop>Start, "Erreur, choix d'intervalle impossible (stop<=start)"
        assert SortedElements[0] or not SortedElements[0], "Le paramètre SortedElements est invalide (non boléen)"
        assert SortedElements[1] or not SortedElements[1], "Le paramètre SortedElements est invalide (non boléen)"
        
        # Extraction données et clés de la feuille
        DataLists = [self.Data.col_values(c, Start, Stop) for c in DataColumns]
        KeyList = self.Data.col_values(KeyColumn, Start, Stop)

        # Arrondi des valeurs des données
        DataLists =  [[round(float(i)) for i in DataList] for DataList in DataLists]
        KeyList = [str(int(float(k))) for k in KeyList]

        # Vidage cases vides
        DataLists = [
            [DataList.pop(DataList.index(i)) for i in DataList if i!=""]
            for DataList in DataLists]
        KeyList = [KeyList.pop(KeyList.index(i)) for i in KeyList if i!=""]

        # Création liste éléments (non merged)
        ElementList = [[KeyList[i]]+[DataList[i] for DataList in DataLists] for i in range(len(KeyList))]
        
        # Merge data with same key (fix) with a dictionnary
        KeyList = list(set(KeyList))
        ElementDict = {}
        for key in KeyList:
            ElementDict[key] = [sum([e[c+1] for e in ElementList if key in e]) for c in range(len(DataColumns))]

        # Reconversion in List
        ElementList = [[key]+ElementDict[key] for key in KeyList]

        if SortedElements[0]:
            # Tri des éléments par données
            def getKey(element):
                return element[SortedElements[2]+1]

            ElementList.sort(key=getKey, reverse=SortedElements[1])
                
        # Création figure
        df = pd.DataFrame(ElementList,columns=[self.Data.cell_value(Start-TitleOffset, KeyColumn)]+[self.Data.cell_value(Start-TitleOffset, DataColumn) for DataColumn in DataColumns])

        df.plot(x=self.Data.cell_value(Start-TitleOffset, KeyColumn),
                y=[self.Data.cell_value(Start-TitleOffset, DataColumn) for DataColumn in DataColumns],
                kind="bar", figsize=figSize)
        
        plt.legend(bbox_to_anchor=(0.8,1.0))

        # Ajout titre
        plt.title(self.Title)
        # Affichage diagramme
        plt.show()
    
    def DiagrammeCirculaire(self, DataColumn=3, KeyColumn=2, Start=15, Stop=None, TitleOffset=2, figSize=(20.0,20.0)):
        """
        Permet de créer un diagramme ciculaire afin de comparer des parts de valeur de clés

        PARAMETRES :
        --------
        Attention, cette focntion assume que le tableau est sous forme verticale et ne supportera pas les formes horizonales
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
        
        Stop : int || None
            index de la dernière ligne (exclue) des éléments à étudier ou "auto" pour exploiter toutes les données (après start)
                default = "auto"

        TitleOffset : int
            Indique l'écart entre le Start et le titre (permet de trouver les titres d'axes)
                default = 2

        figSize : tuple(float, float)
            Indique la taille du diagramme (x, y)
                default (recommandé pour lecture) = (20.0,20.0)

        SORTIE :
        --------
        None
        """  
        # Vérification des paramètres
        assert DataColumn!=KeyColumn, "Erreur : Les colonnes des données et des clés/noms sont les mêmes"
        assert Stop==None or Stop>Start, "Erreur, choix d'intervalle impossible (stop<=start)"

        # Extraction données de la feuille
        DataList = self.Data.col_values(DataColumn, Start, Stop)
        KeyList = self.Data.col_values(KeyColumn, Start, Stop)

        # Création graphique
        def func(pct, allvals):
            absolute = int(pct/100.*np.sum(allvals))
            return "{:.1f}%\n({:d} pers.)".format(pct, absolute)

        fig, ax = plt.subplots(figsize=figSize, subplot_kw=dict(aspect="equal"))

        ax.pie(DataList, labels=KeyList, autopct=lambda pct: func(pct, DataList))   
        
        # Ajout titre graphique
        plt.title(self.Title)


        # Création légende graphique
        ax.legend(title=self.Data.cell_value(Start-TitleOffset, KeyColumn),
          loc="best",
          bbox_to_anchor=(1, 0, 0.5, 1))

        # Affichage graphique
        plt.show()

# Tests des fonctions
if __name__=='__main__':
    # feuille = int(input("feuille à ouvrir : "))
    # xls = xlsDB(feuille)

    xls = xlsDB()
    print("=============================================")
    print("Bienvenue dans mon programme/module de gestion et de visualisation de données au format xls")
    print("Vous pouvez lancer un test pour chacune de ces deux fonctions :")
    print("\t- 1 : DiagrammeMultiBarres")
    print("\t- 2 : DiagrammeCirculaire")
    print("=============================================")

    Choix = input("Choix (1 ou 2) : ")

    if Choix=="1":
        print("Test DiagrammeMultiBarres :")
        # Affichage hommes et femmes sans diplôme, de 16 à 24 ans
        xls.DiagrammeMultiBarres((True,True,0),[3,5],0) 
    
    elif Choix=="2":
        print("Test DiagrammeCirculaire :")
        # Affichage données de 15 (inclu) à 20 (exclu)
        xls.DiagrammeCirculaire(Stop=20) 
    
    else:
        print("Choix incorrect")
        sys.exit("\tArrêt...")