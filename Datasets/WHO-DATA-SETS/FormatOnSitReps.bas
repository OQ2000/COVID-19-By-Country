Attribute VB_Name = "Module1"
Sub UnmergeAllCells()
    ActiveSheet.Cells.UnMerge
End Sub
Sub DeleteRegions()
    Dim lRow As Long
    Dim iCntr As Long
    lRow = 1000
    For iCntr = lRow To 1 Step -1
        If Cells(iCntr, 1) = "Western Pacific Region" Or Cells(iCntr, 1) = "Territories" Or Cells(iCntr, 1) = "European Region" Or Cells(iCntr, 1) = "South-East Asia Region" Or Cells(iCntr, 1) = "Eastern Mediterranean Region" Or Cells(iCntr, 1) = "Territories**" Or Cells(iCntr, 1) = "Region of the Americas" Or Cells(iCntr, 1) = "African Region" Or Cells(iCntr, 1) = "Subtotal" Then
            Rows(iCntr).Delete
        End If
    Next
End Sub
Sub Format()
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete
    Call FindReplaceAll
    Call UnmergeAllCells
    Call DeleteRegions
End Sub
Sub FindReplaceAll()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "*Subtotal*"
rplc = "Subtotal"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub


