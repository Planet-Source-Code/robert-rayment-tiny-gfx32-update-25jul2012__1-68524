VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ulli's PSC Code56112

' To make QueryEnvironment the Default
' Goto Tools, Procedure Attributes, Advanced
' Name = QueryEnvironment
' Set Procedure ID: to (Default)
' Apply.

' On Form1 (If IDE/Environment tested there) place
' Dim Environment As New cEnvironment  'For IDE or ENVIRONMENT

' Used in this program to select the Dropper
' cursor, normal for IDE (DROPPER) but
' DROPPER32 (32bpp cursor) for Environment.
' 32bpp cursors are particularly bad in IDE!
' See code in Sub Optools_MouseUp(...) where Tools = [Dropper].

Option Explicit

Public Enum eEnvironment
    EnvironIDE = 1
    EnvironCompiled = 2
End Enum

Public Property Get QueryEnvironment() As eEnvironment
Attribute QueryEnvironment.VB_UserMemId = 0
    QueryEnvironment = EnvironCompiled
    Debug.Assert Not SetToIDE(QueryEnvironment)
End Property

Private Function SetToIDE(Env As eEnvironment) As Boolean
    Env = EnvironIDE
End Function

