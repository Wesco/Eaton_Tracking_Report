Attribute VB_Name = "Imports"
Option Explicit

Sub ImportMaster()
    Dim Path As String
    
    Path = "\\br3615gaps\gaps\Billy Mac-Master Lists\Eaton Master List.xls"
    
    Workbooks.Open Path
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
    ActiveWorkbook.Close
    
    Sheets("Master").Select
    Columns(1).ClearContents
    Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Value = Range(Cells(1, 4), Cells(ActiveSheet.UsedRange.Rows.Count, 4)).Value
End Sub
