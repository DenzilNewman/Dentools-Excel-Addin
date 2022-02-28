VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "Dentools Excel Addin - Help"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   OleObjectBlob   =   "frmHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TitleHeight As Integer = 29
Const BorderWidth As Integer = 6

Private Sub helpPages_Change()

End Sub

Private Sub helpPages_Scroll(ByVal Index As Long, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

End Sub



Private Sub lblDentoolsLink_Click()
    moduleDentoolsPublicFunctions.OpenWebUrl lblDentoolsLink.Caption
End Sub

Private Sub UserForm_Click()
    frameHelpFeatures.Top = 0 - helpPages.Pages("hpFeatures").ScrollTop
End Sub

Private Sub UserForm_Resize()
    Dim gapSize As Integer
    gapSize = cmdClose.Height * 0.2

    cmdClose.Top = Me.Height - (TitleHeight + (cmdClose.Height + gapSize))
    helpPages.Width = Me.Width - (BorderWidth * 2)
    helpPages.Height = cmdClose.Top - gapSize
    

    
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
    DentoolsAddinEventManager "install"
End Sub



Private Sub UserForm_Initialize()
    lblVersion.Caption = DentoolsAddinVersionString
    lblOSVersion.Caption = Application.OperatingSystem
    lblExcelVersion.Caption = Application.Version
   
    UserForm_Resize
End Sub
