Attribute VB_Name = "Utilities"
'declarations for working with Ini files
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias _
    "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
 
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias _
    "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long
 
'// INI CONTROLLING PROCEDURES
'reads an Ini string
Public Function ReadIni(FileName As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
  v = GetPrivateProfileString(Section, Key, "", RetVal, 255, FileName)
  ReadIni = Left(RetVal, v)
End Function
 
'reads an Ini section
Public Function ReadIniSection(FileName As String, Section As String) As String
Dim RetVal As String * 255, v As Long
  v = GetPrivateProfileSection(Section, RetVal, 255, FileName)
  ReadIniSection = Left(RetVal, v)
End Function
 
'writes an Ini string
Public Sub WriteIni(FileName As String, Section As String, Key As String, Value As String)
  WritePrivateProfileString Section, Key, Value, FileName
End Sub
 
'writes an Ini section
Public Sub WriteIniSection(FileName As String, Section As String, Value As String)
  WritePrivateProfileSection Section, Value, FileName
End Sub

