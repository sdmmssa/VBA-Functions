Attribute VB_Name = "Functions"
Function LastRow(Sht As Worksheet, ColumnID As String)
'function returns the last row of a specified worksheet in the specified column

LastRow = Sht.Cells(Sht.Rows.Count, ColumnID).End(xlUp).Row

End Function
