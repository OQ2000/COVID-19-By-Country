Attribute VB_Name = "Module2"
Sub ColourFixer()
Dim sh As Worksheet
Dim rw As Range
Dim RowCount As Integer

RowCount = 0

Set sh = ActiveSheet
For Each rw In sh.Rows

  If sh.Cells(rw.Row, 1).Value = "" Then
    Exit For
  End If
    a = sh.Cells(rw.Row, 1).Value
    If a = "No Of Sit_Rep" Then
    Rows(rw.Row).Interior.Color = RGB(155, 200, 230)
    Else
    If a Mod 2 = 0 Then
    Rows(rw.Row).Interior.Color = RGB(220, 235, 247)
    Else
    Rows(rw.Row).Interior.Color = RGB(189, 215, 238)
    End If
    End If
  RowCount = RowCount + 1

Next rw

MsgBox (RowCount)
End Sub
