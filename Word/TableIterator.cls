
' Only private variables were used, to get one of the
' use the getter or setter method if its present.
Private Tbl As Table            ' Table-object to iterate

Private HeadCount As Integer   ' Number of head cells
Private BodyCount As Integer   ' Number of body cells per row
Private RowIndex  As Integer   ' Current row index
Private RowCount  As Integer   ' Maximal number of rows
Private CellCount As Integer   ' Maximal number of cells


' Function:     Init
' Description:  Initializes the class and set's default values
' Parameter:
'   -   t       Table       the table-object to iterate
'   -   head    Integer     the number of head cells
'   -   body    Integer     the number of cells that represent a row
Public Sub Init(t As Table, head As Integer, body As Integer)
    Set Tbl = t
    CellCount = Tbl.Range.Cells.count
    HeadCount = head
    BodyCount = body
    RowIndex = 0
    RowCount = (CellCount - HeadCount) / BodyCount
End Sub

' Function:     Valid
' Description:  condition for the loop, checks if the end of the table is reached
' Parameter:    none
' Return:       boolean
'   - False     if the end is reached
'   - True      otherwise
Public Function Valid() As Boolean
    If RowIndex < RowCount Then
        Valid = True
    Else
        Valid = False
    End If
End Function

' Function:     Rewind
' Description:  Sets the iterator to the beginning
' Parameter:    none
Public Sub Rewind()
    RowIndex = 0
End Sub

' Function:     NextRow
' Description:  Counts up the row_index counter
' Parameter:    none
Public Sub NextRow()
    RowIndex = RowIndex + 1
End Sub

' Function:     Cell
' Description:  get the cell number x form the current row
' Parameter:
'   -   Index   Integer     the index of the Cell
' Return:       Cell-Object
Public Function Cell(Index As Integer) As Cell
    Set Cell = Tbl.Range.Cells(Index + RowIndex * BodyCount + HeadCount)
End Function

' Function:     Index
' Description:  returns the current row number
' Parameter:    none
' Return:       Integer
Public Function Index() As Integer
    Index = RowIndex
End Function
