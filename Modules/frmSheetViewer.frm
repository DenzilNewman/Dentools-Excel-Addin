VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetViewer 
   Caption         =   "Sheet Viewer"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "frmSheetViewer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSheetViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refreshDetail()
    Dim ws As Worksheet
    lblWorkbook.Caption = ActiveWorkbook.Name
    lblWorkbookPath.Caption = ActiveWorkbook.Path
    
    lstSheets.Clear
    
    For Each ws In ActiveWorkbook.Sheets
        lstSheets.AddItem ws.Name
    Next
    
    Set ws = Nothing
    lstSheets.ListIndex = 0
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function countVisibleSheets() As Long
    Dim lCount As Integer
    Dim ws As Worksheet
    lCount = 0
    For Each ws In ActiveWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then lCount = lCount + 1
    Next
    Set ws = Nothing
    countVisibleSheets = lCount
End Function

Private Sub cmdOK_Click()
    Dim selSheet As String
    Dim ws As Worksheet
        
    If ActiveWorkbook.Name = lblWorkbook.Caption Then
        If lstSheets.ListIndex > -1 Then
            selSheet = lstSheets.List(lstSheets.ListIndex)
            Set ws = Nothing
            On Error Resume Next
            Set ws = ActiveWorkbook.Sheets(selSheet)
            On Error GoTo 0
                        
            If opSheetVisible.Value Then
                ws.Visible = xlSheetVisible
            Else
                If countVisibleSheets() < 2 Then
                    MsgBox "At least one sheet must remain visible."
                Else
                    If opSheetHidden.Value Then
                        ws.Visible = xlSheetHidden
                    ElseIf opSheetVeryHidden.Value Then
                        ws.Visible = xlSheetVeryHidden
                    End If
                End If
            End If
            
            Set ws = Nothing
            
            Unload Me
        End If
    Else
        refreshDetail
    End If
End Sub

Private Sub cmdRefresh_Click()
    refreshDetail
End Sub

Private Sub lstSheets_Change()
    Dim selSheet As String
    Dim ws As Worksheet
    
    If ActiveWorkbook.Name = lblWorkbook.Caption Then
        If lstSheets.ListIndex > -1 Then
            selSheet = lstSheets.List(lstSheets.ListIndex)
            Set ws = Nothing
            On Error Resume Next
            Set ws = ActiveWorkbook.Sheets(selSheet)
            On Error GoTo 0
            grpVisability.Caption = ws.Name
            If ws.Visible = xlSheetHidden Then
                opSheetHidden.Value = True
            ElseIf ws.Visible = xlSheetVeryHidden Then
                opSheetVeryHidden.Value = True
            Else
                opSheetVisible.Value = True
            End If
            
        End If
    Else
        Set ws = Nothing
        refreshDetail
    End If
End Sub

Private Sub lstSheets_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    refreshDetail
    
    
    
End Sub
