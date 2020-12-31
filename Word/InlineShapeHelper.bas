
' Function:     Crop
' Description:  crops a InlineShape and set new width
' Parameters:
'   -   IL      InlineShape
'   -   Left    Integer
'   -   Top     Integer
'   -   Right   Integer
'   -   Bottom  Integer
'   -   Width   Integer
Public Sub Crop(IL As InlineShape, Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, Width As Single)
    With IL.ConvertToShape
        .PictureFormat.CropLeft = Left
        .PictureFormat.CropTop = Top
        .PictureFormat.CropRight = Right
        .PictureFormat.CropBottom = Bottom
        
        .Width = Width
        
        .ConvertToInlineShape
    End With
End Sub

' Function:     CropSelection
' Description:  Crops a selection of InlineShapes
' Parameters:
'   -   Left    Integer
'   -   Top     Integer
'   -   Right   Integer
'   -   Bottom  Integer
'   -   Width   Integer
Public Sub CropSelection(Left As Integer, Top As Integer, Right As Integer, Bottom As Integer, Width As Single)
    Dim IL As InlineShape
    For Each IL In Selection.InlineShapes
        Crop IL, Left, Top, Right, Bottom, Width
    Next IL
End Sub

' Function:     Replace
' Description:  replaces a InlineShape by file path
' Parameters:
'   -   IL          InlineShape
'   -   FilePath    String          Path to new InlineShape
Sub Replace(IL As InlineShape, FilePath As String)
    Dim sngWdth As Single, SngHght As Single
    Dim rng As Range
    Set rng = IL.Range
    
    IL.Delete
    
    ActiveDocument.InlineShapes.AddPicture FilePath, LinkTofile:=False, savewithdocument:=True, Range:=rng
End Sub
