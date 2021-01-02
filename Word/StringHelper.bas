' Function:     MD5
' Description:  Get the table by it's title
' Parameters:
'   -   sIn    String
'   -   bB64   Boolean
' Return:
'   - Hash string
Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
        
    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc
        
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
        
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = True Then
       MD5 = ConvToBase64(bytes)
    Else
       MD5 = ConvToHex(bytes)
    End If
        
    Set oT = Nothing
    Set oMD5 = Nothing

End Function

' Function:     ConvToBase64
' Description:  Converts a string to a base64 string
' Parameters:
'   -   vIn    Variant
' Return:
'   - converted string
Public Function ConvToBase64(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
   
   Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64 = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing
End Function

' Function:     ConvToHex
' Description:  Converts a string to a hex string
' Parameters:
'   -   vIn    Variant
' Return:
'   - converted string
Public Function ConvToHex(vIn As Variant) As Variant
    'Check that Net Framework 3.5 (includes .Net 2 and .Net 3 is installed in windows
    'and not just Net Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHex = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing
End Function

' Function:     StartsWith
' Description:  Checks if the string starts with a defined prefix
' Parameters:
'   -   str     String      String to check
'   -   prefix  String      needly string
Public Function StartsWith(str As String, prefix As String) As Boolean
    StartsWith = (Left(UCase(str), Len(prefix)) = UCase(prefix))
End Function

' Function:     EndsWith
' Description:  Checks if the string ends with a defined ending
' Parameters:
'   -   str     String      String to check
'   -   ending  String      needly string
Public Function EndsWith(str As String, ending As String) As Boolean
     EndsWith = (Right(Trim(UCase(str)), Len(ending)) = UCase(ending))
End Function
