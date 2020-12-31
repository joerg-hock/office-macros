 
' Function:     DeleteByTitle
' Description:  delete all shaps by it's title
' Parameters:
'   -   Title   String      shape title
Public Sub DeleteByTitle(Title As String)
    Dim ads As Shape

    For Each ads In ActiveDocument.Shapes
        If ads.Title = Title Then
            ads.Delete
        End If
    Next
End Sub
