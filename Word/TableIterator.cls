
' Only private variables were used, to get one of the
' use the getter or setter method if its present.
Private tbl As Table            ' Table-object to iterate

Private head_count As Integer   ' Number of head cells
Private body_count As Integer   ' Number of body cells per row
Private row_index  As Integer   ' Current row index
Private row_count  As Integer   ' Maximal number of rows
Private cell_count As Integer   ' Maximal number of cells


' Function:     Init
' Description:  Initializes the class and set's default values
' Parameter:
'   -   t       Table       the table-object to iterate
'   -   head    Integer     the number of head cells
'   -   body    Integer     the number of cells that represent a row
Public Sub Init(t As Table, head As Integer, body As Integer)
    Set tbl = t
    cell_count = tbl.Range.Cells.count
    head_count = head
    body_count = body
    row_index = 0
    row_count = (cell_count - head) / body
End Sub

' Function:     Valid
' Description:  condition for the loop, checks if the end of the table is reached
' Parameter:    none
' Return:       boolean
'   - False     if the end is reached
'   - True      otherwise
Public Function Valid() As Boolean
    If row_index < row_count Then
        Valid = True
    Else
        Valid = False
    End If
End Function

' Function:     Rewind
' Description:  Sets the iterator to the beginning
' Parameter:    none
Public Sub Rewind()
    row_index = 0
End Sub

' Function:     NextRow
' Description:  Counts up the row_index counter
' Parameter:    none
Public Sub NextRow()
    row_index = row_index + 1
End Sub

' Function:     Cell
' Description:  get the cell number x form the current row
' Parameter:
'   -   Index   Integer     the index of the Cell
' Return:       Cell-Object
Public Function Cell(Index As Integer) As Cell
    Set Cell = tbl.Range.Cells(Index + row_index * body_count + head_count)
End Function

' Function:     Index
' Description:  returns the current row number
' Parameter:    none
' Return:       Integer
Public Function Index() As Integer
    Index = row_index
End Function
