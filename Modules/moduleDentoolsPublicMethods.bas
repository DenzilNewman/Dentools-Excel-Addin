Attribute VB_Name = "moduleDentoolsPublicMethods"
Option Explicit

Public Enum enmTrimDirection
    Both = 0
    Left = 1
    Right = 2
End Enum


Public Sub hideShowSheets()
    'MsgBox "Dentools Addin" & vbNewLine & "  Version: " & DentoolsAddinVersionString, vbInformation
    frmSheetViewer.Show 1
    
End Sub

Public Sub helpDentoolsAddin()
    'MsgBox "Dentools Addin" & vbNewLine & "  Version: " & DentoolsAddinVersionString, vbInformation
    frmHelp.Show 1
    
End Sub



Public Sub crunchRows()
    Dim selectionRange As Range
    Dim topRow As Integer, bottomRow As Integer
    Dim columns As Long
    Dim cLoop As Long, rLoop As Long
    Dim cellValue, newCellValue As Variant
    Dim selectionSheet As Worksheet
    Set selectionRange = Selection
    Set selectionSheet = selectionRange.Worksheet
    columns = selectionSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    topRow = selectionRange.Rows(1).Row
    bottomRow = selectionRange.Rows(selectionRange.Rows.Count).Row
    If topRow < bottomRow Then

        For cLoop = 1 To columns
            cellValue = selectionSheet.Cells(topRow, cLoop).Value
            For rLoop = (topRow + 1) To bottomRow
                newCellValue = selectionSheet.Cells(rLoop, cLoop).Value
                If Not IsEmpty(newCellValue) Then
                    cellValue = cellValue & vbNewLine & newCellValue
                End If
            Next rLoop
        
            selectionSheet.Range(selectionSheet.Cells(topRow, cLoop), selectionSheet.Cells(bottomRow, cLoop)).MergeCells = False
            selectionSheet.Cells(topRow, cLoop).Value = cellValue
            selectionSheet.Range(selectionSheet.Cells(topRow + 1, cLoop), selectionSheet.Cells(bottomRow, cLoop)).ClearContents
            
        Next
        selectionSheet.Range(selectionSheet.Cells(topRow + 1, cLoop), selectionSheet.Cells(bottomRow, cLoop)).EntireRow.Delete Shift:=xlUp
    End If
    selectionSheet.Range(selectionSheet.Cells(topRow, 1), selectionSheet.Cells(topRow, 1)).EntireRow.Select
    Set selectionSheet = Nothing
    Set selectionRange = Nothing
End Sub

Sub tidyCellValues()
    frmTidyCells.Show 1
End Sub

