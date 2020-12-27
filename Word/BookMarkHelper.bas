' Function:     GetByName
' Description:  Get a bookmark by it's name
' Parameters:
'   -   Name    String      Name of the bookmark
' Return:
'   - if Found      Bookmark-Object
'   - otherwise     Nothing
Public Function GetByName(Name As String) As Bookmark
    Dim bm As Bookmark
    For Each bm In ActiveDocument.Bookmarks
        If bm.Name = Name Then
            Set GetByName = bm
            Exit Function
        End If
    Next
    Set GetByName = Nothing
End Function

' Function:     GetArrayByName
' Description:  Get a array of bookmarks by it's name counted up #Name#i
' Parameters:
'   -   Name    String      Name of the bookmark
'   -   Start   Integer     Start of name counter
' Return:
'   - System.Collections.ArrayList
Public Function GetArrayByName(Name As String, Optional Start As Integer = 0) As Object
    Dim Bms As Object
    Set Bms = CreateObject("System.Collections.ArrayList")
    Dim bm As Bookmark
    Dim i As Integer
    i = Start
    
    ' Check if the bookmark without index exists
    Set bm = GetByName(Name)
    If Not bm Is Nothing Then
        Bms.Add bm
    End If
    
    ' Do while loop runs until the last bookmark was found
    Do While True
        Set bm = GetByName(Name & i)
        If Not bm Is Nothing Then
            Bms.Add bm
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    Set GetArrayByName = Bms
End Function

' Function:     SetTextByName
' Description:  Set the range text of a bookmark without deleting it
' Parameters:
'   -   Name    String      Name of the bookmark
'   -   Text    String      Text to set
Public Sub SetTextByName(Name As String, Text As String)
    Dim bm As Bookmark
    Dim rg As Range
    Set bm = GetByName(Name)
    If Not bm Is Nothing Then
        Set rg = bm.Range
        rg.Text = Text
        ActiveDocument.Bookmarks.Add Name, rg
    End If
End Sub

' Function:     SetArrayTextByName
' Description:  Set the range text of a bookmark array without deleting it
' Parameters:
'   -   Name    String      Name of the bookmark
'   -   Text    String      Text to set
'   -   Start   Integer     Start of name counter
Public Sub SetArrayTextByName(Name As String, Text As String, Optional Start As Integer = 0)
    Dim bm As Bookmark
    For Each bm In GetArrayByName(Name, Start)
        SetTextByName bm.Name, Text
    Next
End Sub

