Attribute VB_Name = "ModINIWG"
' ModINIWG.bas

Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
 ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
 ByVal lpDefault As String, ByVal lpReturnedString As String, _
 ByVal nSize As Long, ByVal lpFileName As String) As Long
 ' lpDefault is the string return if no ini file found.
'------------------------------------------------------

' File  stuff  ini & RecentFiles
Public IniTitle$, IniSpec$

Public Function WriteINI(Title$, TheKey$, Info$, ISpec$) As Boolean
   WritePrivateProfileString Title$, TheKey$, Info$, ISpec$
End Function

Public Function GetINI(Title$, TheKey$, Ret$, ISpec$) As Boolean
Dim n As Long
   On Error GoTo NoINI
   Ret$ = String(255, 0)
   n = GetPrivateProfileString(Title$, TheKey$, "", Ret$, 255, ISpec$)
   'N is the number of characters copied to Ret$
   If n <> 0 Then
     GetINI = True
     Ret$ = Left$(Ret$, n)
   Else
     GetINI = False
     Ret$ = ""
   End If
   On Error GoTo 0
   Exit Function
'==========
NoINI:
GetINI = False
Ret$ = ""
End Function

