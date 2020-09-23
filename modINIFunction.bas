Attribute VB_Name = "modINIFunction"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal strSection As String, ByVal strKeyname As Any, ByVal strValue As Any, ByVal lpFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal strSection As String, ByVal strKeyname As String, ByVal strDefault As String, ByVal ReturnedValue As String, ByVal ValueSize As Long, ByVal lpFilename As String) As Long

Public Sub WriteINI(Filename As String, strSection As String, strKeyname As String, strValue As String)
  WritePrivateProfileString strSection, strKeyname, strValue, Filename
End Sub
Public Function GetINI(Filename As String, strSection As String, strKeyname As String, Optional MaxLength As Long = 32) As Variant
  Dim strReturnValue As String * 255
  MaxLength = (MaxLength Mod 255) + 1
  
  GetPrivateProfileString strSection, strKeyname, "", strReturnValue, MaxLength, Filename
  CutPoint = InStr(1, strReturnValue, Chr$(0))
  If CutPoint = 0 Then
    GetINI = strReturnValue
  Else
    CutPoint = CutPoint - 1
    GetINI = Left$(strReturnValue, CutPoint)
  End If
End Function

