Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal nDefault As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileSection Lib "kernel32.dll" Alias "WritePrivateProfileSectionA" ( _
    ByVal lpAppName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long


Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long



Public Function WriteINI(Section As String, Variable As String, Value As String, File As String) As Long
    WriteINI = WritePrivateProfileString(Section, Variable, Value, File)
End Function

Public Function GetINIString(Section As String, Variable As String, File As String) As String
    Dim temp As String * 255
    
    GetINIString = Left(temp, GetPrivateProfileString(Section, Variable, "", temp, 255, File))
End Function

Public Function GetINILong(Section As String, Variable As String, File As String, Optional ByVal nDefault As Integer) As Long
    GetINILong = GetPrivateProfileInt(Section, Variable, nDefault, File)
End Function
