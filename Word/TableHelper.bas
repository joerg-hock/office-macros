
' Function:     Exists
' Description:  Check if Table exists
' Parameters:
'   -   title   String      table title
' Return:       Boolean
Public Function Exists(title As String) As Boolean
    Dim Tbl As Table
    For Each Tbl In ActiveDocument.Tables
        If Tbl.title = title Then
            Exists = True
            Exit Function
        End If
    Next Tbl
    
    Exists = False
End Function

' Function:     GetByTitle
' Description:  Get the table by it's title
' Parameters:
'   -   title   String      table title
' Return:
'   - Table-Object if the table could found, otherwise Null
Public Function GetByTitle(title As String)
    Dim Tbl As Table
    For Each Tbl In ActiveDocument.Tables
        If Tbl.title = title Then
            Set GetByTitle = Tbl
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
Public Sub MergeColumn(Tbl As Table, row As Integer, col As Integer, span As Integer)
    Tbl.Cell(row, col).Merge Tbl.Cell(row, col + span - 1)
End Sub

' Function:     MergeRow
' Description:  merge table cell with column span
' Parameter:
'   - tbl   Table       table-object
'   - row   Integer     row number
'   - col   Integer     column number to start
'   - span  Integer     number of columns to merge
Public Sub MergeRow(Tbl As Table, row As Integer, col As Integer, span As Integer)
    Tbl.Cell(row, col).Merge Tbl.Cell(row + span - 1, col)
End Sub
