Attribute VB_Name = "modMain"
' ------------------------------------------
'
' Get the list of worksheets in the workbook and
' add a sheet called "Table of Content" as the first
' sheet in the workbook. Add hyperlinks to sheets.
'
' ------------------------------------------
Option Explicit

Private Const cTOC = "Table of Content"

Public Sub Process(ByVal sInputFolder As String, ByVal sOutputFolder As String) 

    Dim sh As Worksheet 
    Dim wRow As Long

    On Error Resume Next

    ' Delete the previous sheet if already present
    Worksheets(cTOC).Delete
    If (Err.Number <> 0) Then 
        Err.clear
    End If 

    On Error GoTo 0

    ActiveWorkbook.Sheets.Add Before:=ActiveWorkbook.Worksheets(1)

    With ActiveSheet
        .Name = cTOC
        .Cells(2, 1).Value = "Table of contents"
        .Cells(2, 1).Font.Size = 18
        .Cells(2, 1).Font.Bold = True
    End With

    wRow = 4

    For Each sh In ActiveWorkbook.Worksheets

        If Not (ActiveSheet Is sh) Then

            With ActiveSheet
                .Hyperlinks.Add _
                    Anchor:=ActiveSheet.Cells(wRow, 1), _
                    Address:="", _
                    SubAddress:="'" & sh.Name & "'!A1", _
                    ScreenTip:=sh.Name, _
                    TextToDisplay:=sh.Name
            End With

            wRow = wRow + 1

        End if

    Next

    ActiveWindow.DisplayGridlines = False

End sub
