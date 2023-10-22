VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAllAdoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'''testme

'dictionary
Private documents As Object
Private fileNames As Object

Private Sub InitParam()
    Set documents = CreateObject("Scripting.Dictionary")
    documents.CompareMode = vbTextCompare 'vbTextCompare(1) … ignore case。 vbBinaryCompare(0) … case sensitive.
    Set fileNames = CreateObject("Scripting.Dictionary")
    fileNames.CompareMode = vbTextCompare 'vbTextCompare(1) … ignore case。 vbBinaryCompare(0) … case sensitive.
End Sub

Public Sub AddDocument(ByRef document As clsString, ByVal lang As String, outputFileName As String)
    If documents.Exists(lang) Then
        Dim t As clsString
        Set t = documents.Item(lang)
        t.AddClsString document
        documents.Remove lang
        documents.Add lang, t
    Else
        documents.Add lang, document
    End If

    If fileNames.Exists(lang) Then
        '' do nothing
    Else
        fileNames.Add lang, outputFileName
    End If
End Sub

Public Function saveAll(ByVal tType As String, ByVal OutputDir As String, Optional ByVal OutputDirSub As String = "") As String
    Dim langs() As Variant
    langs = documents.keys()
    
    Dim i As Long
    Dim savePath As String
    Dim bookName As String
    bookName = Split(ActiveWorkbook.Name, ".")(0)

    'Save combined sheet.
    Dim basePathWithBS As String
    Dim basePathSubWithBS As String

    If OutputDir = "" Then
        basePathWithBS = ActiveWorkbook.Path & "\"
    Else
        basePathWithBS = OutputDir & "\"
    End If

    If OutputDirSub = "" Then
        basePathSubWithBS = ActiveWorkbook.Path & "\"
    Else
        basePathSubWithBS = OutputDirSub & "\"
    End If

    Dim fileName As String
    For i = 0 To UBound(langs)
        fileName = fileNames.Item(langs(i))

        If fileName = "" Then
            fileName = bookName
        End If

        ''
        If langs(i) = "" Then
            ' do nothing
        Else
            savePath = basePathWithBS & CStr(langs(i)) & "\"
        End If

        ''
        If OutputDirSub = "" Then
            ' do nothing
        Else
            savePath = savePath & CStr(langs(i)) & "\" & OutputDirSub & "\"
        End If

        CreateDirectory savePath

        savePath = savePath & fileName & "." & tType

        documents.Item(langs(i)).SaveToFileUTF8 savePath
    Next i

    saveAll = savePath
End Function

Private Sub Class_Initialize()
    Call InitParam
End Sub
