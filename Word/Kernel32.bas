
Public Declare PtrSafe Function GetPrivateProfileString32 Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Public Declare PtrSafe Function WritePrivateProfileString32 Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer

