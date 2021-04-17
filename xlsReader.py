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

    def Lecture(self,rowstart=13,rowstop=None,colstart=2,colstop=3,compatibility=False):
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
            - compatibility : bool
                - Si True, MatData sera au format dictionnaire, pour être utilisé avec xlsWriter
                - Si False, MatData sera au format matrice, pour une lecture pour aisée
                - default = False
        
        SORTIE : 
        -----------
            - MatData : list[list[any]]
                - Matrice contenant les données 
                - format :
                    - compatibility = True : cols{col[0]:[col[1:]],...}
                    - compatibility = False : cols[col[rows],...]

        """
        # Vérification des paramètres
        assert rowstart>=0, "ligne de départ invalide (rowstart)"
        assert colstart>=0, "colonne de départ invalide (colstart)"
        assert rowstop==None or rowstop>=0, "ligne de fin invalide (rowstop)"
        assert colstop>=0, "colonne de fin invalide (colstop)"

        # Extraction des données de la zone souhaitée
        MatData = [self.Data.col_values(col, rowstart, rowstop) for col in range(colstart,colstop+1)]

        # Conversion en dictionnaire si demandé
        if compatibility:
            MatData = {col[0]:col[1:] for col in MatData}

        # Renvoi de la matrice
        return MatData

if __name__=='__main__': # Test
    xls = xlsData()
    mat = xls.Lecture()

    # affichage de chaque élément
    for col in mat:
        for row in col:
            print(row,end=' ')
        print(' ')
    
    # affichage de la matrice en entier
    print(mat)