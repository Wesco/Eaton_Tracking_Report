Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim Path As String: Path = "\\br3615gaps\gaps\Eaton\Tracking Report\"
    Dim FileName As String: FileName = "Tracking Report " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    Dim TotalRows As Long

    Application.ScreenUpdating = False

    Clean
    UserImportFile Sheets("Tracking").Range("A1"), FileFilter:="Tracking (*.*), *.*"
    UserImportFile Sheets("POH").Range("A1")

    Sheets("Tracking").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Eaton PO#
    Columns(4).Insert
    Range("D1").Value = "EATON PO#"
    Range("D2:D" & TotalRows).Formula = "=IFERROR(VLOOKUP(VALUE(RIGHT(A2,LEN(A2)-5)),POH!B:S,18,FALSE),"""")"
    Range("D2:D" & TotalRows).Value = Range("D2:D" & TotalRows).Value

    'Eaton Part
    Columns(5).Insert
    Range("E1").Value = "EATON PART"
    Range("E2:E" & TotalRows).Formula = "=IFERROR(VLOOKUP(C2,Master!A:C,3,FALSE),"""")"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value

    'Date Shipped
    Range("H2:H" & TotalRows).NumberFormat = "m/d/yyyy"

    Sheets("Tracking").Copy
    ActiveSheet.UsedRange.Columns.AutoFit
    ActiveWorkbook.SaveAs Path & FileName, xlOpenXMLWorkbook
    ActiveWorkbook.Close

    Email SendTo:="WMaclellan@wescodist.com", _
          Subject:="Tracking Report", _
          Body:="<a href=""file:///" & Path & FileName & """>" & Path & FileName & "</a>"
    MsgBox "Complete!"

    Application.ScreenUpdating = True
    ThisWorkbook.Saved = True
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
