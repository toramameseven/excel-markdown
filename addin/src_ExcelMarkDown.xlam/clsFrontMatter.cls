VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrontMatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public title As String
Public sideBarLabel As String
Public sideBarOrder As long

Public Function CreateFrontMatter()
  Dim s As String
  s = "---" & vbCrLf
  s = s & "title: " & title & vbCrLf
  s = s & "---" & vbCrLf
  Create = s
End Function
