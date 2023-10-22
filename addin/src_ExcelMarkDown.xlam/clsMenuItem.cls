VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Caption As String
Public OnAction As String
Public FaceId As Long
Public KeyA As String
Public IsNoDialog As Boolean
Public IsMenu As Boolean
Public IsTool As Boolean

Private Sub Class_Initialize()
    IsNoDialog = True
    IsMenu = True
    IsTool = True
End Sub
