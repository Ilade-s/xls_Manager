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
    - DiagrammeMultiBarres : Utlisant une colonne de clé, va créer un graphique en barres avec plusieurs colonnes de données
    - DiagrammeMultiCirculaire : Utlisant une ou plusieurs colonnes de données, permet de les comparer dans un ou plusieurs camembert (un pour chaque colonne de données)
"""
import xlrd # Module de gestion mère xls
import matplotlib.pyplot as plt # Création de graphiques
import pandas as pd # Pour utilisation DataFrame (graphiques)
import sys # Messages d'erreur
import numpy as np # Calculs shares DiagrammeCirculaire

class xlsDB:
    def __init__(self, sheet=10, fileName="pop-16ans-dipl6817", TitleCell=(0,0)):
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
        
        TitleCell : tuple(int,int)
            coordonnées de la cellule contenant le titre de la feuille souhaité
                default = (0,0)
        """
        # Vérification paramètres
        for i in TitleCell:
            assert i >= 0
        assert sheet >= 0

        # Ouverture fichier xls
        with xlrd.open_workbook(fileName+".xls", on_demand=True) as file: 
            self.Data = file.get_sheet(sheet)

        # Extraction titre feuille
        (rowx, columnx) = TitleCell
        self.Title = self.Data.cell_value(rowx,columnx)

    def DiagrammeMultiBarres(self, SortedElements=(False, False, 0), DataColumns=[3], KeyColumn=2, Start=15, Stop=None, TitleOffset=2, figSize=(20.0,20.0)):
        """
        Permet de créer des diagrammes en barres pour comparer les éléments de une ou plusieurs colonnes de données

        PARAMETRES :
        --------
        Attention, cette fonction part du principe que le tableau est sous forme verticale et ne supportera pas les formes horizonales
        --------
        SortedElements : tuple(bool, bool, int)
            SortedElements[0] :
                Indique si les données doivent être triées ou non
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
        assert SortedElements[0] or not SortedElements[0], "Le paramètre SortedElements[0] est invalide (non boléen)"
        assert SortedElements[1] or not SortedElements[1], "Le paramètre SortedElements[1] est invalide (non boléen)"
        
        # Extraction données et clés de la feuille
        DataLists = [self.Data.col_values(c, Start, Stop) for c in DataColumns]
        KeyList = self.Data.col_values(KeyColumn, Start, Stop)

        # Arrondi des valeurs des données et clés
        DataLists =  [[round(float(i)) for i in DataList] for DataList in DataLists]
        try:    
            KeyList = [str(int(float(k))) for k in KeyList]
        except:
            pass

        # Vidage cases vides
        DataLists = [[i for i in DataList if i!=""] for DataList in DataLists]
        KeyList = [i for i in KeyList if i!=""]
        
        # Création liste éléments (non merged)
        ElementList = [[KeyList[i]]+[DataList[i] for DataList in DataLists] for i in range(len(DataLists[0]))]
        
        # Merge data with same key (fix) with a dictionary
        KeyList = list(set(KeyList))
        ElementDict = {}
        for key in KeyList:
            ElementDict[key] = [sum([e[c+1] for e in ElementList if key in e]) for c in range(len(DataColumns))]

        # Reconversion in List
        ElementList = [[key]+ElementDict[key] for key in KeyList]

        # Tri des éléments par données
        if SortedElements[0]:
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
    
    def DiagrammeMultiCirculaire(self, SortedElements=(False, False, 0), DataColumns=[3], KeyColumn=2, Start=15, Stop=None, TitleOffset=2, figSize=(20.0,20.0)):
        """
        Permet de créer un diagramme ciculaire afin de comparer des parts de valeur de clés

        Si il y a trois colonnes de données à visualiser, lors de l'affichage, le subplot en bas à droite sera une copie de celui en bas à gauche 
        
        PARAMETRES :
        --------
        Attention, cette fonction part du principe que le tableau est sous forme verticale et ne supportera pas les formes horizonales
        --------
        SortedElements : tuple(bool, bool, int)
            SortedElements[0] :
                Indique si les données doivent être triées ou non
                    default = False
            SortedElements[1] :
                Indique si les données doivent être triées dand l'ordre des aiguilles d'une montre/clockwise(False) ou l'ordre inverse/conterclockwise (True)
                    default = False
            SortedElements[2] :
                Indique l'index de la colonne de données servant à trier les éléments (index dans DataColumns)
                    default = 0

        DataColumns : list[int]
            liste des index de colonnes contenant les valeurs à comparer
                default = [3]
                taille max : 4 éléments (si la liste en contient plus, n'afffichera que les 4 premiers)

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
        assert SortedElements[0] or not SortedElements[0], "Le paramètre SortedElements[0] est invalide (non boléen)"
        assert SortedElements[1] or not SortedElements[1], "Le paramètre SortedElements[1] est invalide (non boléen)"
        if len(DataColumns)>4:    
            DataColumns = DataColumns[:4] # limit of 4 data column to be displayed

        # Extraction données de la feuille
        DataLists = [self.Data.col_values(c, Start, Stop) for c in DataColumns]
        KeyList = self.Data.col_values(KeyColumn, Start, Stop)

        # Arrondi des valeurs des données et clés
        DataLists =  [[round(float(i)) for i in DataList] for DataList in DataLists]
        try:    
            KeyList = [str(int(float(k))) for k in KeyList]
        except:
            pass

        # Création liste éléments (non merged)
        ElementList = [[KeyList[i]]+[DataList[i] for DataList in DataLists] for i in range(len(DataLists[0]))]
        
        # Merge data with same key (fix) with a dictionary
        KeyList = list(set(KeyList))
        ElementDict = {}
        for key in KeyList:
            ElementDict[key] = [sum([e[c+1] for e in ElementList if key in e]) for c in range(len(DataColumns))]

        # Reconversion in List
        ElementList = [[key]+ElementDict[key] for key in KeyList]

        # Tri des éléments par données
        if SortedElements[0]:
            def getKey(element):
                return element[SortedElements[2]+1]

            ElementList.sort(key=getKey, reverse=SortedElements[1])

        # Data recovery from ElementList
        DataLists = [[e[1+c] for e in ElementList] for c in range(len(DataColumns))]
        
        # Calcul nombre de lignes et colonnes
        if len(DataColumns)<=2:
            rows = 1
            cols = len(DataColumns)
        else:
            rows = 2
            cols = 2

        # Création graphique
        fig, ax = plt.subplots(figsize=figSize, subplot_kw=dict(aspect="equal"), nrows=rows, ncols=cols, constrained_layout=True)

        # Création pie charts + titres individuels
        def func(pct, allvals):
            absolute = int(pct/100.*np.sum(allvals))
            return "{:.1f}%\n({:d} pers.)".format(pct, absolute)

        c = 0
        if rows>1 and cols>1: # 2x2
            for row in range(rows):
                for col in range(cols):
                    ax[row][col].pie(DataLists[c], autopct=lambda pct: func(pct, DataLists[c]))   
                    ax[row][col].set_title(self.Data.cell_value(Start-TitleOffset, DataColumns[c]))
                    if c<2: c += 1
        elif cols>1: #1x2
            for col in range(cols):
                ax[col].pie(DataLists[c], autopct=lambda pct: func(pct, DataLists[c]))   
                ax[col].set_title(self.Data.cell_value(Start-TitleOffset, DataColumns[c]))
                c += 1
        else: # 1x1
            ax.pie(DataLists[c], autopct=lambda pct: func(pct, DataLists[c]))   
            ax.set_title(self.Data.cell_value(Start-TitleOffset, DataColumns[c]))

        # Ajout titre graphique
        plt.suptitle(self.Title)

        # Création légende graphique
        plt.legend(title=self.Data.cell_value(Start-TitleOffset, KeyColumn),
          loc="best",
          bbox_to_anchor=(1, 0, 0.5, 1),
          labels=KeyList)

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
    print("\t- 2 : DiagrammeMultiCirculaire")
    print("=============================================")

    Choix = input("Choix (1 ou 2) : ")

    if Choix=="1":
        print("Test DiagrammeMultiBarres :")
        # Affichage hommes et femmes sans diplôme, de 16 à 24 ans, par region
        xls.DiagrammeMultiBarres((True,True,0),[3,5]) 
        # Affichage hommes et femmes sans diplôme, de 16 à 24 ans, par département
        #xls.DiagrammeMultiBarres((True,True,0),[3,5]) 
    
    elif Choix=="2":
        print("Test DiagrammeMultiCirculaire :")
        # Affichage données de 15 (inclu) à 20 (exclu) de quatres colonnes de données : 3,4,6,5, dans l'ordre inverse des aiguilles d'une montre
        xls.DiagrammeMultiCirculaire(Stop=20, DataColumns=[3,4,5], SortedElements=(True, True, 0)) 
    
    else:
        print("Choix incorrect")
        sys.exit("\tArrêt...")