Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Version: 14 February 2018 :* :* :*
''
'' Contents
''
'' ICanHazSheet: check if a sheet exists (by name)
'' ClearAndResizeTable: clear and resize a named table on the active sheet then move the selection to the first cell of the table
'' AI_SendRangeAsHTML: Send a range as an HTML mail
'' AI_GetLastShitName: Get a name of the last sheet in the active workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function ICanHazSheet(cSheetName As String, _
                   Optional oWorkBook As Excel.Workbook) As Boolean

    ICanHazSheet = False
    Dim wb

    If oWorkBook Is Nothing Then
        Set oWorkBook = ActiveWorkbook
    End If

    For Each wb In oWorkBook.Worksheets
        If wb.Name = cSheetName Then
            ICanHazSheet = True
            Exit Function
        End If
    Next wb
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Clear and resize a named table on the active sheet then move the selection to the first cell of the table
'' Arguments: Table name
'' Alexander Ivashkin, 19 January 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearAndResizeTable(TableName As String)
    Dim tblData As ListObject
    Set tblData = ActiveSheet.ListObjects(TableName)
    With tblData.DataBodyRange
        If .Rows.Count > 1 Then
            .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
        End If
    End With
    tblData.DataBodyRange.ClearContents
    tblData.DataBodyRange.Cells(1, 1).Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Send a range as an HTML mail
'' Arguments:
''   Range
''   Email address(es)
''   Optional CC address(es)
'' From Ron de Bruin's website (http://www.rondebruin.nl/win/s1/outlook/bmail2.htm)
'' Adapted and slightly amended by Alexander Ivashkin, 24 January 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AI_SendRangeAsHTML(BodyRange As Range, ToAddress As String, Subj As String, Optional BodyHeader As String = "", Optional BodyFooter As String = "", Optional CCAddress As String = "", Optional BCCAddress As String = "")
    'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    'Don't forget to copy the function RangetoHTML in the module.
    'Working in Excel 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object

    'On Error Resume Next
    'Only the visible cells in the selection
    'Set BodyRange = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set BodyRange = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    'On Error GoTo 0

    If BodyRange Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = ToAddress
        .CC = CCAddress
        .BCC = BCCAddress
        .Subject = Subj
        .HTMLBody = BodyHeader & RangeToHTML(BodyRange) & BodyFooter
        .Display   'or use .Display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangeToHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' http://www.rondebruin.nl/win/s1/outlook/bmail2.htm
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangeToHTML = ts.ReadAll
    ts.Close
    RangeToHTML = Replace(RangeToHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function


Function AI_GetLastShitName() As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Get a name of the last sheet in the active workbook.
'' If you will use in a worksheet formula, be sure to use INDIRECT
''
'' Alexander Ivashkin, 6 February 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    AI_GetLastShitName = Sheets(Sheets.Count).Name

End Function
