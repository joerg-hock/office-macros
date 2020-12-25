' Function:		TableColumnMerge
' Description:	merge table cell with column span
' Parameter:
'	- tbl	Table		table-object
'	- row	Integer		row number
'	- col	Integer		column number to start
'	- span	Integer		number of columns to merge
Private Sub TableColumnMerge(tbl As Table, row As Integer, col As Integer, span As Integer)
    tbl.Cell(row, col).Merge tbl.Cell(row, col + span - 1)
End Sub

' Function:		TableRowMerge
' Description:	merge table cell with column span
' Parameter:
'	- tbl	Table		table-object
'	- row	Integer		row number
'	- col	Integer		column number to start
'	- span	Integer		number of columns to merge
Private Sub TableRowMerge(tbl As Table, row As Integer, col As Integer, span As Integer)
    tbl.Cell(row, col).Merge tbl.Cell(row + span - 1, col)
End Sub
