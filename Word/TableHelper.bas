' Function:     GetByTitle
' Description:  Get the table by it's title
' Parameters:
'   -   title   String      table title
' Return:
'   - Table-Object if the table could found, otherwise Null
Public Function GetByTitle(title As String)
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        If tbl.title = title Then
            Set GetByTitle = tbl
            Exit Function
        End If
    Next
    GetByTitle = Null
End Function

' Function:     MergeColumn
' Description:  merge table cell with column span
' Parameter:
'   - tbl   Table       table-object
'   - row   Integer     row number
'   - col   Integer     column number to start
'   - span  Integer     number of columns to merge
Public Sub MergeColumn(tbl As Table, row As Integer, col As Integer, span As Integer)
    tbl.Cell(row, col).Merge tbl.Cell(row, col + span - 1)
End Sub

' Function:     MergeRow
' Description:  merge table cell with column span
' Parameter:
'   - tbl   Table       table-object
'   - row   Integer     row number
'   - col   Integer     column number to start
'   - span  Integer     number of columns to merge
Public Sub MergeRow(tbl As Table, row As Integer, col As Integer, span As Integer)
    tbl.Cell(row, col).Merge tbl.Cell(row + span - 1, col)
End Sub
