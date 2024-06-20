'Запрашиваем файл QTN
Dim excelFile
excelFile = selectExcel()

'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open (excelFile)

qtn   = ArticlesExcel.Cells(22, 4).Value
plant = ArticlesExcel.Cells(21, 4).Value
sorg  = ArticlesExcel.Cells(3, 4).Value
template = "BUY-" & sorg

Dim arrSerno()

' Считаем, что в 25 строке - начало таблицы для обработки
intRow = 25
' Цикл для каждой строки
On Error Resume Next
Do Until ArticlesExcel.Cells(intRow,10).Value = ""
	ReDim arrSerno(intRow - 24)
	arrSerno(intRow - 24) = ArticlesExcel.Cells(intRow,10).Value
	intRow = intRow + 1
Loop

WScript.Echo Join(arrSerno)
Dim arrUniqSerno : arrUniqSerno= uniqFE(arrSerno)
WScript.Echo Join(arrUniqSerno)


MsgBox "The script finished!", vbSystemModal Or vbInformation


' returns an array of the unique items in for-each-able collection fex
Function uniqFE(fex)
  Dim dicTemp : Set dicTemp = CreateObject("Scripting.Dictionary")
  Dim xItem
  For Each xItem In fex
      dicTemp(xItem) = 0
  Next
  uniqFE = dicTemp.Keys()
End Function