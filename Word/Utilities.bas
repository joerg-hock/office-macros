
' Function:     Wait
' Description:  Sleep / Wait Sub
' Parameters:
'   -   n       Long      Seconds to wait
Public Sub Wait(n As Long)
    Dim T As Date
    T = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, T)
End Sub

' Function:     ArraySize
' Description:  returns the size of an array
' Parameters:
'   -   arr     Variant
Public Function ArraySize(arr() As Variant) As Integer
    ArraySize = UBound(arr) - LBound(arr) + 1
End Function
