
' Function:     UpdateAllStoryRanges
' Description:  Updates all Story-Ranges
' Parameters:   None
Public Sub UpdateAllStoryRanges()
    Dim objStory As Range
    For Each objStory In ActiveDocument.StoryRanges
        DocumentHelper.UpdateAllFieldsInStoryRange objStory
    Next objStory
End Sub

' Function:     UpdateAllFieldsInStoryRange
' Description:  Updates only the given range
' Parameters:
'   -   objStr  Range      Range Object to be updated
Public Sub UpdateAllFieldsInStoryRange(objStr As Range)
    Dim objShape As Shape

    With objStr
        .Fields.Update

        Select Case .StoryType
            Case wdMainTextStory, wdPrimaryHeaderStory, _
              wdPrimaryFooterStory, wdEvenPagesHeaderStory, _
              wdEvenPagesFooterStory, wdFirstPageHeaderStory, _
              wdFirstPageFooterStory

                For Each objShape In .ShapeRange
                    With objShape.TextFrame
                        If .HasText Then .TextRange.Fields.Update
                    End With
                Next
        End Select
    End With
End Sub

' Function:     UpdateAllFieldsFrom
' Description:  Object to be updated
' Parameters:
'   -   obj     Object      An collection of object with update function
Public Sub UpdateAllFieldsFrom(obj As Object)
    Dim elem As Object
    For Each elem In obj
        If Not elem.Update Is Nothing Then elem.Update
    Next
End Sub

' Function:     UpdateAll
' Description:  Updates all Elements
' Parameters:   None
Sub UpdateAll()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    DocumentHelper.UpdateAllStoryRanges
    DocumentHelper.UpdateAllFieldsFrom ActiveDocument.Indexes
    DocumentHelper.UpdateAllFieldsFrom ActiveDocument.TablesOfAuthorities
    DocumentHelper.UpdateAllFieldsFrom ActiveDocument.TablesOfFigures
    DocumentHelper.UpdateAllFieldsFrom ActiveDocument.TablesOfContents

    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
End Sub

' Function:     InsertWatermarkAllSections
' Description:  Inserts a watermark in all sections
' Parameters:
'   -   Text    String      Watermark text
Public Sub InsertWatermarkAllSections(Optional Text As String = "DRAFT")
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        InsertWatermark sec, Text
    Next
End Sub

' Function:     InsertWatermark
' Description:  Inserts a watermark in one section
' Parameters:
'   -   sec     Section     Section to set the watermark
'   -   Text    String      Watermark text
Public Sub InsertWatermark(sec As Section, Optional Text As String = "DRAFT")
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    
    On Error GoTo ErrHandler
    Dim strWMName As String
     
    sec.Range.Select
    strWMName = sec.Index
    Selection.HeaderFooter.Shapes.AddTextEffect(msoTextEffect1, Text, "Arial", 1, False, False, 0, 0).Select
    
    With Selection.ShapeRange
         
        .Name = strWMName
        .TextEffect.NormalizedHeight = False
        .Line.Visible = False
         
        With .Fill
            .Visible = True
            .Solid
            .ForeColor.RGB = Gray
            .Transparency = 0.5
        End With
         
        .Rotation = 315
        .LockAspectRatio = True
        .Height = InchesToPoints(2.42)
        .Width = InchesToPoints(6.04)
         
        With .WrapFormat
            .AllowOverlap = True
            .Side = wdWrapNone
            .Type = 3
        End With
         
        .RelativeHorizontalPosition = wdRelativeVerticalPositionMargin
        .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
        .Left = wdShapeCenter
        .Top = wdShapeCenter
    End With
     
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
     
    Exit Sub
     
ErrHandler:
    MsgBox "An error occured trying to insert the watermark." & Chr(13) & _
    "Error Number: " & Err.Number & Chr(13) & _
    "Decription: " & Err.Description, vbOKOnly + vbCritical, "Error"
End Sub

' Function:     GetEntireDocumentContent
' Description:  Get the text from a DOC-File
' Parameters:
'   -   FileName    String      full file path
Function GetEntireDocumentText(FileName As String) As String
    Dim Text As String
    
    With Documents.OpenNoRepairDialog(FileName:=FileName, Visible:=False, ReadOnly:=True)
        With .Range(Start:=0, End:=.Range.End)
            Text = .Text
        End With
        .Close False
    End With
    
    GetEntireDocumentText = Text
End Function

' Function:     GetVariable
' Description:  Get a document variable
' Parameters:
'   -   Name    String      variable name
Public Function GetVariable(Name As String) As String
    Dim adv As Variable
    Dim res As String
    res = ""
    For Each adv In ActiveDocument.Variables
        If adv.Name = Name Then
          res = adv.Value
        End If
    Next adv
    
    GetVariable = res
End Function

' Function:     DeleteVariable
' Description:  deletes a document variable
' Parameters:
'   -   Name    String      variable name
Public Sub DeleteVariable(Name As String)
    Dim adv As Variable
    For Each adv In ActiveDocument.Variables
        If adv.Name = Name Then
          adv.Delete
          Exit Sub
        End If
    Next adv
End Sub

' Function:     SetVariable
' Description:  sets a document variable
' Parameters:
'   -   Name    String      variable name
'   -   Value   String
Public Sub SetVariable(Name As String, Value As String)
    Dim adv As Variable
    For Each adv In ActiveDocument.Variables
        If adv.Name = Name Then
          adv.Value = Value
          Exit Sub
        End If
    Next adv
    ActiveDocument.Variables.Add Name, Value
End Sub

' Function:     SaveAsDraft
' Description:  saves the active document as a draft version and activates revision tracking
' Parameters:   none
Public Sub SaveAsDraft()
    ActiveDocument.SaveAs2 Replace(ActiveDocument.FullName, ".docx", "_DRAFT.docx")
    
    DocumentHelper.InsertWatermarkAllSections "DRAFT"
    
    ActiveDocument.TrackRevisions = True
    ActiveDocument.Save
End Sub
