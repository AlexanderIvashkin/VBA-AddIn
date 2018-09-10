''''''''''''''''''''''''''''''''''''''''''''''
' Import / Export modules
' Adapted from: https://www.rondebruin.nl/win/s9/win002.htm
'' Version: 14 February 2018 :* :* :*
''
''''''''''''''''''''''''''''''''''''''''''''''

' If IsDebug is turned on, we're working on this XLAM and using our own subroutines to save our own source code.
' Turning IsDebug on would save the source files in the same folder as the AddIn
Const IsDebug = True
'Const cSourceWorkbook = "AI.xlam"

Option Explicit

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the current folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    If IsDebug Then
        szSourceWorkbook = ThisWorkbook.Name
    Else
        szSourceWorkbook = ActiveWorkbook.Name
    End If
    
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                szFileName = szFileName & ".bas"
            Case Else
                szFileName = szFileName & ".txt"
        End Select
        
        ''' Export the component to a text file.
        cmpComponent.Export szExportPath & szFileName
        
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
   
    Next cmpComponent

End Sub


Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    If IsDebug Then
        SpecialPath = ThisWorkbook.Path
    Else
        SpecialPath = ActiveWorkbook.Path
    End If
    'SpecialPath = WshShell.SpecialFolders("MyDocuments")
    

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function
