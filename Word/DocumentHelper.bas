' Function:		GetTableByTitle
' Description:	Get the table by it's title
' Parameters:
'	-	title	String		table title
' Return:
'	- Table-Object if the table could found, otherwise Null
Private Function GetTableByTitle(title As String)
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        If tbl.title = title Then
            Set GetTableByTitle = tbl
            Exit Function
        End If
    Next
    GetTableByTitle = Null
End Function