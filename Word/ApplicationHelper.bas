
Private Const WM_CLOSE = &H10
Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2


' Function:     CaptureWindow
' Description:  takes a screenshot of a window by application name
' Parameters:
'   -   AppName   String      Name of the application
' Return:       boolean
Public Function CaptureWindow(AppName As String) As Boolean
    On Error GoTo ErrorHandler
    
    AppActivate AppName
    
    Dim AltScan As Double
    
    AltScan = user32.MapVirtualKey(VK_MENU, 0)
    user32.KeybdEvent VK_MENU, AltScan, 0, 0
    
    user32.KeybdEvent VK_SNAPSHOT, 0, 0, 0
    user32.KeybdEvent VK_MENU, AltScan, KEYEVENTF_KEYUP, 0
    
    Utilities.Wait 1
    CaptureWindow = True
    Exit Function
    
ErrorHandler:
    CaptureWindow = False
End Function
