import xlsWriter
import xlsPlot

data = {"keys":[chr(65+i) for i in range(10)],"data":[i for i in range(10)],"d":[i+1 for i in range(10)],"da":[i+2 for i in range(10)]}

xls = xlsWriter.xlsWriter()

xls.AddData(data, KeysCol="keys", Title=("Données",4,4))

xls.SaveFile() # sauvegarde xls

xls = xlsPlot.xlsDB(0, "ExtractedData", TitleCell=(4,4))

xls.DiagrammeMultiBarres(DataColumns=[1,2,3],KeyColumn=0,Start=1,TitleOffset=1,SortedElements=(True, True, 0))