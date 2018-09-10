Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Version: 14 February 2018 :* :* :*
''
'' Contents
'' AI_ParseMagicSymbols: Replace magic symbols (placeholders) with dynamic data
'' AI_MATCH_Regexp: Regexp version of MATCH (match a regexp against a range of strings)
'' AI_MATCH_Regexps:Regexp version of MATCH - another version: match a string against an array of regexps
'' AI_RegExp_IsMatch: Check if a regexp matches
'' AI_RegExp_GetSubMatch: Get a submatch from a regexp
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Replace magic symbols (placeholders) with dynamic data.
''
'' Arguments: a string full of magic.
''
'' Placeholders consist of one symbol prepended with a %:
''    %d - current date
''    %t - current time
''    %u - username (user ID)
''    %n - full user name (usually name and surname)
''    %% - literal % (placeholder escape)
''    Using an unsupported magic symbol will treat the % literally, as if it had been escaped.
''    A single placeholder terminating the string will also be treated literally.
''    Magic symbols are case-sensitive.
''
'' Returns:   A string with no magic but with lots of beauty.
''
'' Examples:
'' "Today is %d" becomes "Today is 2018-01-26"
'' "Beautiful time: %%%t%%" yields "Beautiful time: %16:10:51%"
'' "There are %zero% magic symbols %here%.", true to its message, outputs "There are %zero% magic symbols %here%."
'' "%%% looks lovely %%%" would show "%% looks lovely %%" - one % for the escaped "%%" and the second one for the unused "%"!
''
'' Alexander Ivashkin, 26 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_ParseMagicSymbols(ByVal TextToParse As String) As String

Dim sFinalResult As String
Dim aTokenizedString() As String
Dim sTempString As String
Dim sPlaceholder As String
Dim sCurrentString As String
Dim iIterator As Integer
Dim iTokenizedStringSize As Integer
Dim bThisStringHasPlaceholder As Boolean

' Default placeholder is "%"
Const cPlaceholderSymbol As String = "%"

aTokenizedString = Split(Expression:=TextToParse, Delimiter:=cPlaceholderSymbol)
iTokenizedStringSize = UBound(aTokenizedString())
bThisStringHasPlaceholder = False
sFinalResult = ""

For iIterator = 0 To iTokenizedStringSize
    sCurrentString = aTokenizedString(iIterator)
    
    If bThisStringHasPlaceholder Then
        If sCurrentString <> "" Then
            sPlaceholder = Left(sCurrentString, 1)
            sTempString = Right(sCurrentString, Len(sCurrentString) - 1)
            
            ' This is the place where the MAGIC happens
            Select Case sPlaceholder
                Case "d":
                    sCurrentString = Date & sTempString
                Case "t":
                    sCurrentString = Time & sTempString
                Case "u":
                    sCurrentString = Environ$("Username") & sTempString
                Case "n":
                    sCurrentString = Environ$("fullname") & sTempString
                Case Else:
                    sCurrentString = cPlaceholderSymbol & sCurrentString
            End Select
        Else
            ' We had two placeholders in a row, meaning that somebody tried to escape!
            sCurrentString = cPlaceholderSymbol
            bThisStringHasPlaceholder = False
        End If
    End If
    
    sFinalResult = sFinalResult & sCurrentString
    
    If sCurrentString = "" Or (iIterator + 1 <= iTokenizedStringSize And sCurrentString <> cPlaceholderSymbol) Then
        ' Each string in the array has been split at the placeholders. If we do have a next string, then it must contain a magic symbol.
        
        bThisStringHasPlaceholder = True
        ' Even though it is called "...ThisString...", it concerns the NEXT string.
        ' The logic is correct as we will check this variable on the next iteration, when the next string will become ThisString.
    Else
        bThisStringHasPlaceholder = False
    End If
    
Next iIterator

AI_ParseMagicSymbols = sFinalResult

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Regexp version of MATCH
'' Returns row of the first match of LookupRegexp in the LookupArray range
'' Accepts a single column only (throws xlErrValue if given a two-dimensional range)
'' Alexander Ivashkin, 17 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AI_MATCH_Regexp(ByVal LookupRegexp As String, ByVal LookupArray As Range, Optional ByVal IsCaseInSensitive As Boolean = False) As Variant
    Dim vLookupRange As Variant
    Dim i As Long
    
On Error GoTo Hell
    vLookupRange = LookupArray.Value

    ' Accept only a single column
    If UBound(vLookupRange, 2) > 1 Then
        AI_MATCH_Regexp = CVErr(xlErrValue)
        Exit Function
    End If

    For i = 1 To UBound(vLookupRange, 1)
        If AI_RegExp_IsMatch(vLookupRange(i, 1), LookupRegexp, IsCaseInSensitive) Then
            AI_MATCH_Regexp = i
            Exit Function
        End If
    Next i

    AI_MATCH_Regexp = CVErr(xlErrNA)
    Exit Function

Hell:
    Debug.Print "AI_MATCH_Regexp Something went wrong: ", Err.Number, "  ", Err.Description, "  ", Err.Source
    AI_MATCH_Regexp = CVErr(xlErrValue)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Regexp version of MATCH - another version: match a string against an array of regexps
'' Returns row of the first match of LookupString against the LookupRegexp range
'' Accepts a single column only (throws xlErrValue if given a two-dimensional range)
'' Alexander Ivashkin, 17 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AI_MATCH_Regexps(ByVal LookupString As String, ByVal LookupRegexp As Range, Optional ByVal IsCaseInSensitive As Boolean = False) As Variant
    Dim vLookupRange As Variant
    Dim i As Long
    
On Error GoTo Hell
    vLookupRange = LookupRegexp.Value

    ' Accept only a single column
    If UBound(vLookupRange, 2) > 1 Then
        AI_MATCH_Regexps = CVErr(xlErrValue)
        Exit Function
    End If

    For i = 1 To UBound(vLookupRange, 1)
        If AI_RegExp_IsMatch(LookupString, vLookupRange(i, 1), IsCaseInSensitive) Then
            AI_MATCH_Regexps = i
            Exit Function
        End If
    Next i

    AI_MATCH_Regexps = CVErr(xlErrNA)
    Exit Function

Hell:
    Debug.Print "AI_MATCH_Regexps Something went wrong: ", Err.Number, "  ", Err.Description, "  ", Err.Source
    AI_MATCH_Regexps = CVErr(xlErrValue)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Check if a regexp matches
'' Alexander Ivashkin, 14 Nov 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AI_RegExp_IsMatch(ByVal TextToParse As String, ByVal RegularExpression As String, Optional ByVal bIgnoreCase As Boolean = False) As Boolean
    
    Dim rgxRegex As Regexp
    Set rgxRegex = New Regexp
    Dim rgxResults As Object
    Dim rgxMatch As Variant

    With rgxRegex
        .Pattern = RegularExpression
        .Global = True
        .IgnoreCase = bIgnoreCase
    End With

On Error GoTo Hell
    Set rgxResults = rgxRegex.Execute(TextToParse)
    
    If rgxResults.Count > 0 Then
        AI_RegExp_IsMatch = True
    Else
        AI_RegExp_IsMatch = False
    End If
    
    Exit Function

Hell:
    Debug.Print "AI_RegExp_IsMatch Something went wrong: ", Err.Number, "  ", Err.Description, "  ", Err.Source
    AI_RegExp_IsMatch = CVErr(xlErrNA)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Get a submatch from a regexp
'' Alexander Ivashkin, 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AI_RegExp_GetSubMatch(ByVal TextToParse As String, ByVal RegularExpression As String, Optional ByVal bIgnoreCase As Boolean = False, Optional ByVal iMatchIndex As Integer, Optional ByVal iSubMatchIndex As Integer) As String
    
    Dim rgxRegex As Regexp
    Set rgxRegex = New Regexp
    Dim rgxResults As Object
    Dim rgxMatch As Variant

    With rgxRegex
        .Pattern = RegularExpression
        .Global = Not (iSubMatchIndex = 0 And iMatchIndex = 0)
        .IgnoreCase = bIgnoreCase
    End With

On Error GoTo Hell
    Set rgxResults = rgxRegex.Execute(TextToParse)
    
    'For Each rgxMatch In rgxResults
    '     Debug.Print rgxMatch.Value
    'Next rgxMatch

    If rgxResults.Count > 0 Then
        'Debug.Print rgxResults(iMatchIndex).SubMatches(iSubMatchIndex)
        AI_RegExp_GetSubMatch = rgxResults(iMatchIndex).SubMatches(iSubMatchIndex)
    Else
        AI_RegExp_GetSubMatch = ""
    End If
    
    Exit Function

Hell:
    Debug.Print "AI_RegExp_GetSubMatch Something went wrong: ", Err.Number, "  ", Err.Description, "  ", Err.Source
    AI_RegExp_GetSubMatch = CVErr(xlErrNA)
End Function


