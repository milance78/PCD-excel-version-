Attribute VB_Name = "Sheet2"
Private Sub Worksheet_Change(ByVal Target As Range)
If Intersect(Target, Me.Range("B5:H22")) Is Nothing Then Exit Sub
Module1.ParseMagicImportText CStr(Target.Value), False
End Sub
