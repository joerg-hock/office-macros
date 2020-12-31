 
' Function:     FileDialog
' Description:  Shows a file select dialog an returns the selected files
' Parameters:
'   -   InitPath    String      the initial directory
' Return:       FileDialogSelectedItems
Public Function FileDialog( _
    InitPath As String, _
    Optional Title As String = "Select a File", _
    Optional Filters As FileDialogFilters, _
    Optional MultiSelect As Boolean = False, _
    Optional View As mosfileDialogview _
    ) As FileDialogSelectedItems
    
    Dim dlgOpen As FileDialog
    Dim filepath As String
    
    Set dlgOpen = Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
    
    With dlgOpen
        .InitialFileName = InitPath
        .InitialView = View
        .AllowMultiSelect = MultiSelect
        .Title = Title
        .Filters.Clear
        
        If Not Filters Is Nothing Then
            Dim f As FileDialogFilter
            For Each f In Filter
                .Filters.Add f.Description, f.Extensions, f.Position
            Next f
        End If
        
        .Show
        
        FileDialog = .SelectedItems
    End With
End Function

