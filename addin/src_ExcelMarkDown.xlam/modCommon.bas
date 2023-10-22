Attribute VB_Name = "modCommon"

Option Explicit

Sub CreateDirectory_Test()
   CreateDirectory "c:\sss\bbbb\"
End Sub


Function CreateDirectory(fullPath As String) As Long
  ''On Error GoTo ErrExit



  Dim myPath() As String
  myPath = Split(fullPath, "\")

  
  Dim pathCurrent As String
  pathCurrent = myPath(0)
  Dim i As Long
  For i = 1 To UBound(myPath)
    If myPath(i) <> "" Then
        pathCurrent = pathCurrent & "\" & myPath(i)
        If Dir(pathCurrent, vbDirectory) = "" Then
            MkDir pathCurrent
        End If
    End If
  Next
  
  CreateDirectory = 0
  Exit Function
  
ErrExit:
  CreateDirectory = 9
End Function


Function ReplaceNbsp2Sp(str As String) As String
    Dim nbsp(0 To 1) As Byte
    nbsp(0) = 160
    nbsp(1) = 0
    ReplaceNbsp2Sp = Replace(str, nbsp, " ")
End Function

Public Sub OpenFolder(ByVal folderName As String)
    On Error GoTo EH
    folderName = """" & folderName & """"
    Shell "C:\Windows\Explorer.exe " & folderName, vbNormalFocus
    
    Exit Sub
EH:
    MsgBox Err.Description
End Sub


''' ######################### fso
Public Function GetFileName(ByVal fullPath As String) As String
    ' name and extension
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetFileName = FSO.GetFileName(fullPath)
    Set FSO = Nothing
End Function


Public Function GetParentFolderName(ByVal fullPath As String) As String
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName = FSO.GetParentFolderName(fullPath)
    Set FSO = Nothing
End Function

Public Function getFileExtension(ByVal fileName As String) As String
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo EH
    getFileExtension = FSO.GetExtensionName(fileName)
    GoTo NE
EH:
    getFileExtension = ""
    Debug.Print fileName & "::_getFileExtension" & "err"
NE:
    Set FSO = Nothing
End Function


Public Function FsoFolderExists(ByVal fullPath1 As String) As Boolean
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FsoFolderExists = FSO.FolderExists(fullPath1)
    Set FSO = Nothing
End Function


Public Function FileExists(fileName As String)
    Dim FSO As Object, target As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileExists = FSO.FileExists(fileName)
End Function

Public Function LenMbcs(ByVal strVal As String) As Long
    LenMbcs = LenB(StrConv(strVal, vbFromUnicode))
End Function

Public Function LeftMbcs(ByVal strVal As String, ByVal leftlen As Long) As String
    LeftMbcs = strVal
    If leftlen <= 0 Then
        Exit Function
    End If

    Dim i As Long
    Dim strName As String
    For i = Len(strVal) To 1 Step -1
        strName = Left$(strVal, i)
        If LenMbcs(strName) <= leftlen Then
            If LenMbcs(strName) = leftlen Then
                'continue
            Else
                strName = strName & " "
            End If
            Exit For
        End If
    Next i
    LeftMbcs = strName
End Function
