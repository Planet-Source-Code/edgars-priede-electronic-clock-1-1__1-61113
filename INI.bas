Attribute VB_Name = "INI"
Private Declare Function GetPrivateProfileInt _
  Lib "kernel32" _
  Alias "GetPrivateProfileIntA" _
  (ByVal strSection As String, _
   ByVal strKeyName As String, _
   ByVal lngDefault As Long, _
   ByVal strFileName As String) _
As Long
  
Private Declare Function GetPrivateProfileString _
  Lib "kernel32" _
  Alias "GetPrivateProfileStringA" _
  (ByVal strSection As String, _
   ByVal strKeyName As String, _
   ByVal strDefault As String, _
   ByVal strReturned As String, _
   ByVal lngSize As Long, _
   ByVal strFileName As String) _
As Long
   
Private Declare Function WritePrivateProfileString _
  Lib "kernel32" _
  Alias "WritePrivateProfileStringA" _
  (ByVal strSection As String, _
   ByVal strKeyNam As String, _
   ByVal strValue As String, _
   ByVal strFileName As String) _
As Long


Public Function INIGetSettingInteger( _
  strSection As String, _
  strKeyName As String, _
  strFile As String) _
  As Integer
  Dim intValue As Integer

  On Error GoTo PROC_ERR
  
  intValue = GetPrivateProfileInt(strSection, strKeyName, 0, strFile)

  INIGetSettingInteger = intValue

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIGetSettingInteger"
  Resume PROC_EXIT
  
End Function

Public Function INIGetSettingString( _
  strSection As String, _
  strKeyName As String, _
  strFile As String) _
  As String
  Dim strBuffer As String * 256
  Dim intSize As Integer

  On Error GoTo PROC_ERR
  
  intSize = GetPrivateProfileString(strSection, strKeyName, "", strBuffer, 256, strFile)

  INIGetSettingString = Left$(strBuffer, intSize)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIGetSettingString"
  Resume PROC_EXIT
  
End Function

Public Function INIWriteSetting( _
  strSection As String, _
  strKeyName As String, _
  strValue As String, _
  strFile As String) _
  As Integer
  Dim intStatus As Integer

  On Error GoTo PROC_ERR
  
  intStatus = WritePrivateProfileString( _
    strSection, _
    strKeyName, _
    strValue, _
    strFile)

  INIWriteSetting = (intStatus <> 0)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "INIWriteSetting"
  Resume PROC_EXIT
  
End Function
