Attribute VB_Name = "Program"
Option Explicit
Public Const RepositoryName As String = "Eaton_Tracking_Report"
Public Const VersionNumber As String = "1.0.0"

Sub Main()
    Dim Path As String: Path = "\\br3615gaps\gaps\Eaton\Tracking Report\"
    Dim FileName As String: FileName = "Tracking Report " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    Dim TotalRows As Long
    Dim Addr As String
    Dim Col As Integer

    Application.ScreenUpdating = False

    Clean
    UserImportFile Sheets("Tracking").Range("A1"), FileFilter:="Tracking (*.*), *.*"
    UserImportFile Sheets("POH").Range("A1")

    ReOrderCols

    Sheets("Tracking").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Eaton PO#
    Columns(4).Insert
    Range("D1").Value = "EATON PO#"
    Col = FindColumn("PO Number")
    Addr = Cells(2, Col).Address(False, False)
    Range("D2:D" & TotalRows).Formula = "=IFERROR(VLOOKUP(VALUE(RIGHT(" & Addr & ",LEN(" & Addr & ")-5)),POH!B:S,18,FALSE),"""")"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value

    'Eaton Part
    Columns(5).Insert
    Range("E1").Value = "EATON PART"
    Col = FindColumn("Material number")
    Addr = Cells(2, Col).Address(False, False)
    Range("E2:E" & TotalRows).Formula = "=IFERROR(VLOOKUP(" & Addr & ",Master!A:C,3,FALSE),"""")"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value

    Sheets("Tracking").Copy

    'Fix Number Formats
    ActiveSheet.UsedRange.NumberFormat = "General"
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
    ActiveSheet.UsedRange.HorizontalAlignment = xlLeft

    'Date Shipped
    Columns(FindColumn("Date Shipped")).NumberFormat = "m/d/yyyy"

    ActiveSheet.UsedRange.Columns.AutoFit

    On Error GoTo RenameFile
    ActiveWorkbook.SaveAs Path & FileName, xlOpenXMLWorkbook
    On Error GoTo 0
    ActiveWorkbook.Close

    Email SendTo:="WMaclellan@wescodist.com", _
          Subject:="Tracking Report", _
          Body:="<a href=""file:///" & Path & FileName & """>" & Path & FileName & "</a>"
    MsgBox "Complete!"

    Application.ScreenUpdating = True
    ThisWorkbook.Saved = True
    Exit Sub

RenameFile:
    If Err.Number = 1004 Then
        FileName = "Tracking Report " & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "ss") & ".xlsx"
        Resume
    Else
        MsgBox "Error # " & Err.Number & "  - " & Err.Description
        Exit Sub
    End If
End Sub

Sub ReOrderCols()
    Dim ColList As Variant
    Dim ColNum As Integer
    Dim i As Integer

    ColList = Array("PO Number", _
                    "Sales Document", _
                    "Material Number", _
                    "Freight Carrier text", _
                    "Tracking / PRO number", _
                    "Date Shipped", _
                    "Ordered", _
                    "Invoice Number")

    Sheets("Tracking").Select

    On Error GoTo ColNotFound
    For i = 0 To UBound(ColList)
        ColNum = FindColumn(ColList(i))
        If ColNum <> i + 1 Then
            Columns(ColNum).Cut
            Columns(i + 2).Insert
        End If
    Next
    On Error GoTo 0
    Exit Sub

ColNotFound:
    ColNum = i + 1
    Resume Next
End Sub

Sub Clean()
    Dim s As Worksheet
    Dim PrevDispAlerts As Boolean

    PrevDispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" And s.Name <> "Master" Then
            s.Select
            Range("A1").Select
            Cells.Delete
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    Application.DisplayAlerts = PrevDispAlerts
End Sub
