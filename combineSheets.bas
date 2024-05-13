Attribute VB_Name = "combineSheets"
Option Explicit

Sub Combine()
    
    'define variables
    Dim wbPaths As Variant
    Dim wb As Workbook
    Dim i As Long
    Dim destinationWorksheet As Worksheet
    Dim sourceTable As ListObject
    Dim firstOpenRow As Long
    Const startRow As Byte = 1
    
    'create new sheet to hold combined data
    Sheets.Add(After:=Sheets("Key")).Name = "CombineData"
    
    ' create headers
    ' select tab to add headers
    Sheets("CombineData").Select
    
    ' add headers
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Column1"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Column2"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Column3"
    Application.CutCopyMode = False
    
    ' set destination worksheet
    Set destinationWorksheet = ThisWorkbook.Worksheets("CombineData")
    
    ' move back to key sheet
    Sheets("Key").Select
    
    ' define workbook paths array
    wbPaths = Array(Range("B7").Value, _
                    Range("B8").Value, _
                    Range("B9").Value)
              
    ' For next loop to iterate through wookbook paths
    For i = LBound(wbPaths) To UBound(wbPaths)
        
        ' open workbook(i)
        Set wb = Workbooks.Open(wbPaths(i))
        
        ' set source table
        Set sourceTable = wb.Worksheets("Data").ListObjects("Data")
        
        'calculate fist open row on destination worksheet
        With destinationWorksheet
            firstOpenRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
        'copy sourceTable column 1 to destination column A
        sourceTable.ListColumns("Column1").DataBodyRange.Copy _
            Destination:=destinationWorksheet.Range("A" & firstOpenRow)
            
        'copy sourceTable column 2 to destination column B
        sourceTable.ListColumns("Column2").DataBodyRange.Copy _
            Destination:=destinationWorksheet.Range("B" & firstOpenRow)
        
        'copy sourceTable column 3 to destination column C
        sourceTable.ListColumns("Column3").DataBodyRange.Copy _
            Destination:=destinationWorksheet.Range("C" & firstOpenRow)
        
        ' close workbook(i)
        wb.Close SaveChanges:=False
    Next i

End Sub


