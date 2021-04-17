import xlsWriter
import xlsPlot
import xlsReader

# importation de xlsWriter
data = {"keys":[chr(65+i) for i in range(10)],"data":[i for i in range(10)],"d":[i+1 for i in range(10)],"da":[i+2 for i in range(10)]}

xls = xlsWriter.xlsWriter()

xls.AddData(data, KeysCol="keys", Title=("Données",4,4))

xls.SaveFile() # sauvegarde xls

# importation de xlsReader
xls = xlsReader.xlsData(0, "ExtractedData", TitleCell=(4,4))

ReadData = xls.Lecture(rowstart=0,colstart=0,colstop=3,compatibility=True)

print(ReadData) # affichage données récupérées

# exemple d'utilisation de xlsWriter à la suite de xlsReader (copie de données)
xls = xlsWriter.xlsWriter()

xls.AddData(ReadData, KeysCol="keys", Title=("Données copiées",4,4))

xls.SaveFile("DataCopy") # sauvegerde en DataCopy.xls

# importation de xlsPlot
xls = xlsPlot.xlsDB(0, "ExtractedData", TitleCell=(4,4))

xls.DiagrammeMultiBarres(DataColumns=[1,2,3],KeyColumn=0,Start=1,TitleOffset=1,SortedElements=(True, True, 0))