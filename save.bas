Attribute VB_Name = "Module1"
'Declarations of INI File

Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
      "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
      String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
      "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public strServer As String
Public intPort As Integer
Public strNickRoot As String
Public strNick As String
Public strPass As String
Public strOwner As String
Public strChannel As String
Public strKey As String
Public charPre As String
Public strMessages() As String



Function WriteIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
   Dim i As Integer
   On Error GoTo sWriteIniFileError

   i = WritePrivateProfileString(sSection, sItem, sText, sIniFileName)
   WriteIniFile = True

   Exit Function
sWriteIniFileError:
   WriteIniFile = False
End Function






Function ReadIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sDefault As String) As String
   Dim iRetAmount As Integer
   Dim sTemp As String

   sTemp = String$(10000, 0)
   iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 10000, sIniFileName)
   sTemp = Left$(sTemp, iRetAmount)
   ReadIniFile = sTemp
End Function




