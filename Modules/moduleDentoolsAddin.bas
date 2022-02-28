Attribute VB_Name = "moduleDentoolsAddin"
Option Explicit
Public Const DentoolsAddinVersionString As String = "0.100"

'http://supportingtech.blogspot.co.uk/2011/03/microsoft-faceid-numbers-for-vba.html

'Stuff to look at for possible later inclusion
' https://excelchamps.com/blog/camera-tool/
'https://excelchamps.com/blog/useful-macro-codes-for-vba-newcomers/
' Being able to copy a selection as comma seperated values

Private Const keyChars As String = "ZDYQ1jyk7tCBsSNXPMF2a6br3wuWRVvcq9EndULxp5e8mloATGH0O4fghIiJzK"
Private Const addinMenuCaption As String = "Dentools"
Private menuBar As CommandBarControl
Private macroPath As String
Private securityKey As String

Sub DentoolsAddinEventManager(ByVal eventName As String, Optional ByVal strContextInfo As String)
    Select Case LCase$(eventName)
    Case "open"
        installBar
    '--------------------------------------------
    Case "install"
        installBar True
    Case "uninstall"
        unInstallBar
    End Select
End Sub



Private Sub installBar(Optional ByVal forceUpdate As Boolean = True)
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
    Dim strAddinMenuCaption As String
    
    macroPath = "'" & ThisWorkbook.FullName & "'!moduleDentoolsPublicMethods."
    
    
    If InStr(strAddinMenuCaption, "&") = 0 Then
        strAddinMenuCaption = "&" & addinMenuCaption
    Else
        strAddinMenuCaption = addinMenuCaption
    End If
     
    Set cmbBar = Application.CommandBars("Worksheet Menu Bar")
    Set menuBar = Nothing
    
    For Each cmbControl In cmbBar.Controls
        'Debug.Print cmbControl.Caption & ", " & strAddinMenuCaption
        If cmbControl.Caption = strAddinMenuCaption Then
            If ((forceUpdate = False) And (cmbControl.Tag = DentoolsAddinVersionString)) Then
                Set menuBar = cmbControl
            Else
                cmbControl.Delete
            End If
        End If
    Next
    
    If menuBar Is Nothing Then
        Set menuBar = cmbBar.Controls.Add(Type:=msoControlPopup, temporary:=True) 'adds a menu item to the Menu Bar
        With menuBar
            .Caption = strAddinMenuCaption 'names the menu item
            .Tag = DentoolsAddinVersionString
            With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
                .Caption = "Help" 'adds a description to the menu item
                .OnAction = macroPath & "helpDentoolsAddin" 'runs the specified macro
                .FaceId = 487 'assigns an icon to the dropdown
            End With
            
            
            With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
                .Caption = "Hide/Show Sheets" 'adds a description to the menu item
                .TooltipText = "Hide or show sheets, even very hidden"
                .OnAction = macroPath & "hideShowSheets" 'runs the specified macro
                .FaceId = 2556 'assigns an icon to the dropdown
            End With
            
            
            
            With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
                .Caption = "Crunch Rows" 'adds a description to the menu item
                .TooltipText = "Crunch multiple rows (including merged cells) into a single row"
                .OnAction = macroPath & "crunchRows" 'runs the specified macro
                .FaceId = 3177 'assigns an icon to the dropdown
            End With
        
            
            With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
                .Caption = "Tidy Cell Values" 'adds a description to the menu item
                .TooltipText = "Trim Cells, tidy junk values"
                .OnAction = macroPath & "tidyCellValues" 'runs the specified macro
                .FaceId = 1964 'assigns an icon to the dropdown
            End With
            
        End With
    End If
    
    Set cmbControl = Nothing
    Set cmbBar = Nothing
End Sub

Private Sub unInstallBar()
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
    Dim strAddinMenuCaption As String
    
    If InStr(strAddinMenuCaption, "&") = 0 Then
        strAddinMenuCaption = "&" & addinMenuCaption
    Else
        strAddinMenuCaption = addinMenuCaption
    End If
    
    If menuBar Is Nothing Then
        Set cmbBar = Application.CommandBars("Worksheet Menu Bar")
        For Each cmbControl In cmbBar.Controls
            'Debug.Print cmbControl.Caption & ", " & strAddinMenuCaption
            If cmbControl.Caption = strAddinMenuCaption Then
                Set menuBar = cmbControl
            End If
        Next
        Set cmbBar = Nothing
    End If
    
    If Not menuBar Is Nothing Then
        menuBar.Delete
        Set menuBar = Nothing
    End If
End Sub


Private Function generateKey() As String
    Dim retString As String
    Dim rSeed As Double, cLoop As Long
    Randomize
    rSeed = Rnd
    For cLoop = 1 To 20
        retString = retString & Mid$(keyChars, CInt((rSeed * Len(keyChars)) + 1), 1)
        Randomize rSeed
        rSeed = Rnd
    Next cLoop
    generateKey = retString
End Function

Public Sub DentoolsAddinSettingDelete(ByVal SettingName As String)
    DeleteSetting "DentoolsAddin", "AddinConfig", SettingName
End Sub

Function getUserSecKey()
    If securityKey = "" Then
        securityKey = GetSetting("DentoolsAddin", "AddinConfig", "SecUserKey", "")
        If securityKey = "" Then
            securityKey = generateKey()
            SaveSetting "DentoolsAddin", "AddinConfig", "SecUserKey", securityKey
        End If
    End If
    getUserSecKey = securityKey
End Function

Function keyNums() As Variant
    Dim keyNumVals() As Long, midPoint As Integer
    Dim keyString As String, kLoop As Long
    keyString = getUserSecKey()
    ReDim keyNumVals(Len(keyString))
    midPoint = CInt(Len(keyString) / 2)
    For kLoop = 0 To UBound(keyNumVals) - 1
        keyNumVals(kLoop) = InStr(keyChars, Mid$(keyString, kLoop + 1, 1)) - midPoint
    Next kLoop
    keyNums = keyNumVals
End Function

Function Encrypt(ByVal strValue As String) As String
    Dim kRot() As Long, kIdx As Long, cIdx As Long
    Dim strOut() As String
    kRot = keyNums()
    kIdx = 0
    If strValue <> "" Then
        ReDim strOut(Len(strValue))
        For cIdx = 1 To Len(strValue)
            kIdx = kIdx + 1
            strOut(cIdx - 1) = Chr$(Asc(Mid$(strValue, cIdx, 1)) + kRot(kIdx))
            If kIdx >= UBound(kRot) Then kIdx = 0
        Next
    End If
    Encrypt = Join(strOut, "")
End Function

Function Decrypt(ByVal strValue As String) As String
    Dim kRot() As Long, kIdx As Long, cIdx As Long
    Dim strOut() As String
    kRot = keyNums()
    kIdx = 0
    If strValue <> "" Then
        ReDim strOut(Len(strValue))
        For cIdx = 1 To Len(strValue)
            kIdx = kIdx + 1
            strOut(cIdx - 1) = Chr$(Asc(Mid$(strValue, cIdx, 1)) + (kRot(kIdx) * -1))
            If kIdx >= UBound(kRot) Then kIdx = 0
        Next
    End If
    Decrypt = Join(strOut, "")
End Function

Public Property Get DentoolsAddinSetting(ByVal SettingName As String) As Variant
  DentoolsAddinSetting = DeSerialiseValue(Decrypt(GetSetting("DentoolsAddin", "AddinSetting", SettingName, "")))
End Property

Public Property Let DentoolsAddinSetting(ByVal SettingName As String, ByVal vNewValue As Variant)
  SaveSetting "DentoolsAddin", "AddinSetting", SettingName, Encrypt(SerialiseValue(vNewValue))
End Property
