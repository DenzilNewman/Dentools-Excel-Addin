VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTidyCells 
   Caption         =   "Tidy Cell Options"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5760
   OleObjectBlob   =   "frmTidyCells.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTidyCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDo_Change()
    If cboDo.ListIndex = 0 Then
        cboThenDo1.ListIndex = 0
        lblThen1.Visible = False
        cboThenDo1.Visible = False
    Else
        lblThen1.Visible = True
        cboThenDo1.Visible = True
    End If
End Sub

Private Sub cboThenDo1_Change()
    If cboThenDo1.ListIndex = 0 Then
        cboThenDo2.ListIndex = 0
        lblThen2.Visible = False
        cboThenDo2.Visible = False
    Else
        lblThen2.Visible = True
        cboThenDo2.Visible = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub formApply()
    Dim trimVals As String
    Dim i As Long, cPos As Long
    Dim strChar As String
    Dim applyRange As Range, curCell As Range
    Dim trimDirection As enmTrimDirection
    Dim thenDo(2) As String
    
    Set applyRange = Nothing
    If RefRange.Text <> "" Then
        On Error Resume Next
        Set applyRange = Range(RefRange.Text)
        On Error GoTo 0
    End If
    
    If Not applyRange Is Nothing Then
        trimVals = ""
        If lstTrim.ListIndex > 0 Then
            For i = 0 To lstTrimChars.ListCount - 1
                If lstTrimChars.Selected(i) = True Then
                    strChar = lstTrimChars.List(i)
                    cPos = InStr(strChar, "{")
                    If cPos Then
                        strChar = Mid$(strChar, cPos + 1)
                        cPos = InStr(strChar, "}")
                        If cPos Then strChar = Mid$(strChar, 1, cPos - 1)
                    End If
                    
                    Select Case UCase$(strChar)
                    Case "SPACE": strChar = " "
                    Case "CR": strChar = vbCr
                    Case "LF": strChar = vbLf
                    Case "TAB": strChar = vbTab
                    Case Len(strChar) > 1
                        strChar = ""
                    End Select
                    trimVals = trimVals & strChar
                End If
            Next i
        End If
    End If
    
    If cboDo.ListIndex = 0 Then
        thenDo(0) = ""
    Else
        thenDo(0) = cboDo.List(cboDo.ListIndex)
    End If
    
    If ((cboThenDo1.ListIndex = 0) Or (cboThenDo1.Visible = False)) Then
        thenDo(1) = ""
    Else
        thenDo(1) = cboThenDo1.List(cboThenDo1.ListIndex)
    End If
    
    If ((cboThenDo2.ListIndex = 0) Or (cboThenDo2.Visible = False)) Then
        thenDo(2) = ""
    Else
        thenDo(2) = cboThenDo2.List(cboThenDo2.ListIndex)
    End If
      
    For Each curCell In applyRange.Cells
      If trimVals <> "" Then curCell.Value = TrimCharacters(curCell.Value, trimVals)
      
      For i = 0 To 2
        Select Case thenDo(i)
        Case "Empty cells with the value ""NULL"""
            If curCell.Value = "NULL" Then curCell.Value = ""
        Case "Replace all whitespace with spaces"
            curCell.Value = RemoveCharacters(curCell.Value, "", " ")
        Case "Reduce multiple spaces to a single space"
            curCell.Value = RemoveRepeatingCharacters(curCell.Value, " ")
        Case "Remove CrLf, Cr, Lf"
            curCell.Value = RemoveCharacters(curCell.Value, vbNewLine)
        Case "Remove spaces"
            curCell.Value = RemoveCharacters(curCell.Value, " ")
        End Select
      Next i
      
    Next curCell
    
End Sub

Private Sub cmdOK_Click()
    formApply
    Unload Me
End Sub

Private Sub lstTrim_Change()
    If lstTrim.ListIndex = 0 Then
        lstTrimChars.Enabled = False
    Else
        lstTrimChars.Enabled = True
    End If

End Sub

Private Sub lstTrimChars_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim i As Long, curSelection As Range
    
    lstTrim.Clear
    lstTrim.AddItem "No Trim"
    lstTrim.AddItem "Left & Right"
    lstTrim.AddItem "Left"
    lstTrim.AddItem "Right"
    lstTrim.ListIndex = 1
    
    
    lstTrimChars.Clear
    lstTrimChars.AddItem "Spaces {SPACE}"
    lstTrimChars.AddItem "Tabs {TAB}"
    lstTrimChars.AddItem "Carriage-Return {Cr}"
    lstTrimChars.AddItem "Line-Feed {Lf}"
    
    For i = 0 To lstTrimChars.ListCount - 1
        lstTrimChars.Selected(i) = True
    Next i
    
    cboDo.AddItem "Nothing"
    cboDo.AddItem "Empty cells with the value ""NULL"""
    cboDo.AddItem "Replace all whitespace with spaces"
    cboDo.AddItem "Reduce multiple spaces to a single space"
    cboDo.AddItem "Remove CrLf, Cr, Lf"
    cboDo.AddItem "Remove spaces"
    
    For i = 0 To cboDo.ListCount - 1
        cboThenDo1.AddItem (cboDo.List(i))
        cboThenDo2.AddItem (cboDo.List(i))
    Next i
    
    cboDo.ListIndex = 1
    cboThenDo1.ListIndex = 2
    cboThenDo2.ListIndex = 3
    
    
    Set curSelection = Selection
    If Not curSelection Is Nothing Then
        RefRange.Text = curSelection.Address(True, True, xlA1, True)
        Set curSelection = Nothing
    End If
    'lstTrimChars.AddItem "Double-Spaces {SPACE}{SPACE}"

End Sub

