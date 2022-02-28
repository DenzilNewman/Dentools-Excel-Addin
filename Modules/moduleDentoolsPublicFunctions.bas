Attribute VB_Name = "moduleDentoolsPublicFunctions"
Option Explicit

Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Public Sub OpenWebUrl(ByVal thisURL As String, Optional ByVal defaultProtocol As String = "https")
    Dim lSuccess As Long
    Dim cPos As Integer
    Dim protocol As String
       
    
    cPos = InStr(1, thisURL, "://")
    If cPos Then
        protocol = LCase$(Mid$(thisURL, 1, cPos - 1))
        If Not (protocol = "http" Or protocol = "https") Then
            thisURL = ""
        End If
    Else
        protocol = defaultProtocol
        cPos = InStr(1, thisURL, "://")
        If cPos = 0 Then protocol = protocol & "://"
        thisURL = protocol & thisURL
    End If
    
    
    If (thisURL <> "") Then
        lSuccess = ShellExecute(0, "Open", thisURL)
    End If

End Sub

Public Function TrimCharacters(ByVal strTrimValue As String, Optional ByVal trimChars As String = "", Optional ByVal trimDirection As enmTrimDirection = enmTrimDirection.Both)
    Dim leftPos As Long, rightPos As Long
    Dim cLoop As Long
    leftPos = 1
    rightPos = Len(strTrimValue)
    
    If trimChars = "" Then
        trimChars = vbNewLine & Chr$(0) & "     "
    End If
    
    If trimDirection = Both Or trimDirection = Left Then
        For cLoop = leftPos To rightPos
            If InStrB(1, trimChars, Mid$(strTrimValue, cLoop, 1)) Then
                leftPos = cLoop + 1
            Else
                Exit For
            End If
        Next cLoop
    End If
    
    If trimDirection = Both Or trimDirection = Right Then
        For cLoop = rightPos To leftPos Step -1
            If InStrB(1, trimChars, Mid$(strTrimValue, cLoop, 1)) Then
                rightPos = cLoop - 1
            Else
                Exit For
            End If
        Next cLoop
    End If
    
    TrimCharacters = Mid$(strTrimValue, leftPos, rightPos + 1 - leftPos)
End Function

Public Function RemoveCharacters(ByVal strTrimValue As String, Optional ByVal charsToRemove As String = "", Optional ByVal ReplaceChar As String = "")
    Dim cLoop As Long
    Dim retVal() As String
    
    If charsToRemove = "" Then
        charsToRemove = vbNewLine & Chr$(0) & "     "
    End If
    
    If Len(strTrimValue) > 0 Then
        ReDim retVal(Len(strTrimValue))
        For cLoop = 1 To Len(strTrimValue)
            If InStrB(1, charsToRemove, Mid$(strTrimValue, cLoop, 1)) = 0 Then
                retVal(cLoop) = Mid$(strTrimValue, cLoop, 1)
            Else
                retVal(cLoop) = ReplaceChar
            End If
        Next cLoop
    End If
    RemoveCharacters = Join(retVal, "")
End Function

Public Function RemoveRepeatingCharacters(ByVal strTrimValue As String, Optional ByVal charsToRemove As String = "")
    Dim cLoop As Long
    Dim retVal() As String
    
    If Len(strTrimValue) > 0 Then
        ReDim retVal(Len(strTrimValue))
        retVal(0) = Mid$(strTrimValue, 1, 1)
        For cLoop = 2 To Len(strTrimValue)
            If (charsToRemove = "" Or InStrB(1, charsToRemove, Mid$(strTrimValue, cLoop, 1)) > 0) And Mid$(strTrimValue, cLoop, 1) = Mid$(strTrimValue, cLoop - 1, 1) Then
                retVal(cLoop) = ""
            Else
                retVal(cLoop) = Mid$(strTrimValue, cLoop, 1)
            End If
        Next cLoop
    End If
    RemoveRepeatingCharacters = Join(retVal, "")
End Function


Public Function getWorkbookDictionary(Optional ByVal DictionaryWorkbook As Workbook = Nothing, Optional ByVal autoCreate As Boolean = True) As Worksheet
    Dim retVal As Worksheet
    If DictionaryWorkbook Is Nothing Then Set DictionaryWorkbook = ActiveWorkbook
    
    On Error Resume Next
    Set retVal = DictionaryWorkbook.Worksheets("_Dictionary")
    If Err Then Set retVal = Nothing
    On Error GoTo 0
        
    If autoCreate And (retVal Is Nothing) Then
        Set retVal = DictionaryWorkbook.Worksheets.Add
        With retVal
            .Name = "_Dictionary"
            .Cells(1, 1) = "KEY"
            .Cells(1, 2) = "VALUE"
            .Visible = xlSheetVeryHidden
        End With
    End If
    
    Set getWorkbookDictionary = retVal
    Set retVal = Nothing
    Set DictionaryWorkbook = Nothing
End Function


Public Function getWBDValue(ByVal keyName As String, Optional ByVal valueName As String = "Value", Optional ByVal defaultValue As Variant, Optional ByVal DictionaryWorkbook As Workbook = Nothing) As Variant
    Dim wsDictionary As Worksheet
    Dim idxMatch As Integer, valCol As Integer
    Dim maxCols As Integer, maxRows As Integer
    Dim curRange As Range
    
    Set wsDictionary = getWorkbookDictionary(DictionaryWorkbook, False)
    If wsDictionary Is Nothing Then
        If Not IsMissing(defaultValue) Then getWBDValue = defaultValue
    Else
        valueName = UCase$(Trim$(valueName))
        With wsDictionary.Cells.SpecialCells(xlCellTypeLastCell)
            maxCols = .Column
            maxRows = .Row
        End With
        On Error Resume Next
        idxMatch = Application.WorksheetFunction.Match(valueName, wsDictionary.Range(wsDictionary.Cells(1, 1), wsDictionary.Cells(1, maxCols)), 0)
        If Err Then idxMatch = -1
        On Error GoTo 0
                
        If idxMatch = -1 Then
            If Not IsMissing(defaultValue) Then getWBDValue = defaultValue
        Else
            valCol = idxMatch
            keyName = UCase$(Trim$(keyName))
            
            On Error Resume Next
            idxMatch = Application.WorksheetFunction.Match(keyName, wsDictionary.Range(wsDictionary.Cells(2, 1), wsDictionary.Cells(maxRows, 1)), 0)
            If Err Then idxMatch = -1
            On Error GoTo 0
            If idxMatch = -1 Then
                If Not IsMissing(defaultValue) Then getWBDValue = defaultValue
            Else
                getWBDValue = wsDictionary.Cells(1 + idxMatch, valCol)
            End If
            
        End If
        
    End If
    Set curRange = Nothing

End Function

Public Function setWBDValue(ByVal keyName As String, ByVal setValue As Variant, Optional ByVal valueName As String = "Value", Optional ByVal DictionaryWorkbook As Workbook = Nothing) As Variant
    Dim wsDictionary As Worksheet
    Dim idxMatch As Integer, valCol As Integer
    Dim maxCols As Integer, maxRows As Integer
    Dim curRange As Range
    
    Set wsDictionary = getWorkbookDictionary(DictionaryWorkbook, True)
    valueName = UCase$(Trim$(valueName))
    With wsDictionary.Cells.SpecialCells(xlCellTypeLastCell)
        maxCols = .Column
        maxRows = .Row
    End With
    On Error Resume Next
    idxMatch = Application.WorksheetFunction.Match(valueName, wsDictionary.Range(wsDictionary.Cells(1, 1), wsDictionary.Cells(1, maxCols)), 0)
    If Err Then idxMatch = -1
    On Error GoTo 0
    
    If idxMatch = -1 Then
        idxMatch = maxCols
        maxCols = maxCols + 1
        wsDictionary.Cells(1, idxMatch).Value = valueName
    End If
    
    valCol = idxMatch
    keyName = UCase$(Trim$(keyName))

    On Error Resume Next
    idxMatch = Application.WorksheetFunction.Match(keyName, wsDictionary.Range(wsDictionary.Cells(2, 1), wsDictionary.Cells(maxRows, 1)), 0)
    If Err Then idxMatch = -1
    On Error GoTo 0
    
    If idxMatch = -1 Then
        idxMatch = maxRows
        maxRows = maxRows + 1
        wsDictionary.Cells(idxMatch, 1).Value = keyName
    End If
    
    
    wsDictionary.Cells(idxMatch, valCol + 1) = setValue

       
    
    Set curRange = Nothing

End Function

Function SerialiseValue(ByVal newValue As Variant) As String
    Dim strVal As String, vType As String
    Dim arrY As Long
    vType = LCase$(TypeName$(newValue))
    If Mid$(vType, Len(vType) - 1, 2) = "()" Then vType = "array"
    
    Select Case vType
    Case "array"
        arrY = 0
        On Error Resume Next
        arrY = UBound(newValue, 1)
        On Error GoTo 0
        If arrY = 0 Then
            strVal = "[]"
        ElseIf arrY = 1 Then
            Stop
        Else
            Debug.Print "Cannot Serialize Multi-dimentional Array"
        End If
    Case "string"
        strVal = "$" & newValue
    Case "integer", "long", "single", "double", "boolean"
        strVal = Mid$(vType, 1, 1) & CStr(newValue)
    Case "decimal"
        strVal = "x" & CStr(newValue)
    Case "date"
        strVal = "{" & (CStr(DateDiff("s", "1970-01-01", newValue))) & "}"
    Case Else
        Debug.Print "Cannot Serialize Type: " & LCase$(TypeName$(newValue))
    End Select
    SerialiseValue = strVal
End Function

Function testdumbArray()
    Dim x(1, 1) As String
    Debug.Print SerialiseValue(x)
End Function


Function DeSerialiseValue(ByVal serialised As String) As Variant
    Dim sType As String, retVal As Variant
    sType = Mid$(serialised, 1, 1)
    serialised = Mid$(serialised, 2)
    Select Case sType
    Case "$"
        retVal = serialised
    Case "{"
        retVal = DateAdd("s", CLng(Mid$(serialised, 1, Len(serialised) - 1)), "1970-01-01")
    Case "i"
        retVal = CInt(serialised)
    Case "l"
        retVal = CLng(serialised)
    Case "s"
        retVal = CSng(serialised)
    Case "d"
        retVal = CDbl(serialised)
    Case "b"
        retVal = CBool(serialised)
    Case "x"
        retVal = CDec(serialised)
    Case Else
        Debug.Print "Unknown type: " & sType
    End Select
    DeSerialiseValue = retVal
End Function
