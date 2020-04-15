Attribute VB_Name = "Module1"
Sub Multi_FindReplace()
'PURPOSE: Find & Replace a list of text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim sht As Worksheet
Dim fndList As Variant
Dim rplcList As Variant
Dim x As Long

fndList = Array("Republic of Korea", "Viet*", "*Lao People's*", "The United Kingdom", "Russian Federation", "Kosovo*", "Iran*", "United States of*", "United States Virgin*", "Turks and Caicos*", "Democratic Republic*", "United Republic of*", "Central African*", "Cabo Verde", "*Diamond Princess*", "Northern Mariana Islands*", "*?nternational ?onveyance*", "Bolivia*", "Bosnia*", "Cote_D’ivoire", "Cote_D_Ivoire", "Curacao", "Saint_Barthelemy", "Venezuela*")
rplcList = Array("South Korea", "Vietnam", "Laos", "United Kingdom", "Russia", "Kosovo", "Iran", "United States of America", "United States Virgin Islands", "Turks and Caicos Islands", "Congo", "Tanzania", "Central African Republic", "Cape Verde", "Diamond Princess", "Northern Mariana Islands", "Diamond Princess", "Bolivia", "Bosnia", "Côte_D’ivoire", "Côte_D’ivoire", "Curaçao", "Saint_Barthélemy", "Venezuela")

'Loop through each item in Array lists
  For x = LBound(fndList) To UBound(fndList)
    'Loop through each worksheet in ActiveWorkbook
      For Each sht In ActiveWorkbook.Worksheets
        sht.Cells.Replace What:=fndList(x), Replacement:=rplcList(x), _
          LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
          SearchFormat:=False, ReplaceFormat:=False
      Next sht
  
  Next x

End Sub

