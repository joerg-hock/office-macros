
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

