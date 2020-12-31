
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

' Function:     StartsWith
' Description:  Checks if the string starts with a defined prefix
' Parameters:
'   -   str     String      String to check
'   -   prefix  String      needly string
Public Function StartsWith(str As String, prefix As String) As Boolean
    StartsWith = Left(str, Len(prefix)) = prefix
End Function
