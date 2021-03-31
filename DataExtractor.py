"""
DataExtractor
---------------
Ce module secondaire permet d'extraire des données sous forme de liste (simple ou imbriquée).
Celles-ci sont ensuite sauvegardée dans un tableur au format xls, qui sera exploitable par xlsPlot.py
Il sera directement inclus dans xlsPlot.py dans un second temps.

FONCTIONNEMENT :
---------------
    La liste en paramètre peut contenir des clés ou non : si trouvées, seront mises dans la colonne 0 du tableur, sinon les indexs seront mis à leur place

"""
import xlwt # écriture de fichier xls

class xlsData:
    def __init__(self,NewFile=True):
        """
        Quand appelée, créé un Workbook object avec une feuille ("DataSheet") qui pourra ensuite être modifié puis sauvegardé
        """
        if NewFile:
            File = xlwt.Workbook() # création tableur
            self.Sheet = File.add_sheet("DataSheet") # ajout d'une feuille
    
    def AddData(Data,ColStart=0,RowStart=0):
        """
        Ajoute les données en paramètre à la feuille instancée dans __init__

        PARAMETRES :
        -------------
            - Data : list || list[list[key,...]]
                liste contenant les données à ajouter et éventuellement les clés à ajouter au début
            - ColStart : int
                Colonne de départ pour écriture : si des clés sont données (Data), elles seront écrites sur cette colonne
                - Default = 0
            - RowStart : int
                Ligne de départ des données à ajouter
                - Default = 0
        SORTIE :
        -------------
        (indirectement) Les données sont ajoutées à la feuille, qui peut ensuite être sauvegardée
        """
        pass

    def SaveFile(self,FileName="ExtractedData"):
        """
        Sauvegarde le fichier xls

        PARAMETRES :
        -------------
            - FileName : str
                Nom du fichier à enregistrer
        """
        self.File.save(FileName)

if __name__=="__main__": # test
    data = [[chr(65+i),i] for i in range(10)]
    print(data)
    xlsData()