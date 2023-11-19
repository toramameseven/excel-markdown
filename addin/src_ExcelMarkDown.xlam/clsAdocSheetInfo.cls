VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAdocSheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dicPicType As Object
Private dicSheetParams As Object
Private dicLangs As Object
Private LangRow As Long

Private Sub InitParam()
    Set dicPicType = CreateObject("Scripting.Dictionary")
    dicPicType.CompareMode = vbTextCompare 'vbTextCompare(1) … ignore case。 vbBinaryCompare(0) … case sensitive。

    dicPicType.Add "PNG", "PNG"
    dicPicType.Add "JPG", "JPG"
    dicPicType.Add "BMP", "BMP"
End Sub

'VALUECOL default lang(if VALUECOL2 is empty, fall back to this), VALUECOL2 target lang
Public Sub SetSheetInfo(ByRef sht As Worksheet, ByVal FLGCOL As Long, ByVal VALUECOL As Long, ByVal VALUECOL2 As Long)
    
    Set dicSheetParams = CreateObject("Scripting.Dictionary")
    dicSheetParams.CompareMode = vbTextCompare 'vbTextCompare(1) … ignore case。 vbBinaryCompare(0) … case sensitive。
    
    Dim i As Long
    Dim myKey As String
    Dim myValue As String
    Dim myValue2 As String
    Dim theValue As String

    
    For i = 1 To 1000
        myKey = sht.Cells(i, FLGCOL)
        myValue = sht.Cells(i, VALUECOL)
        myValue2 = sht.Cells(i, VALUECOL2)
        theValue = IIf(myValue2 <> "", myValue2, myValue)
        
        
        If myKey <> "" Then
            dicSheetParams.Add UCase(myKey), theValue
            
            If UCase(myKey) = "LANG" Then
                LangRow = i
            End If
        Else
            Exit For
        End If
    Next i
    
    
    '' for split document by language or etc.
    Set dicLangs = CreateObject("Scripting.Dictionary")
    dicLangs.CompareMode = vbTextCompare 'vbTextCompare(1) … ignore case。 vbBinaryCompare(0) … case sensitive。
    
    Dim endcol As Long
    If LangRow > 0 Then
        endcol = sht.Cells(LangRow, columns.Count).End(xlToLeft).Column
        For i = VALUECOL To endcol
            myKey = sht.Cells(LangRow, i)
            If myKey <> "" Then
                dicLangs.Add CStr(i), myKey
            End If
        Next i
    End If
    
    If dicLangs.Count = 0 Then
       'column 3(default), "" empty, or 3,jp/   10,en/     15,cn/
        dicLangs.Add "3", ""
    End If
End Sub

Public Sub SetInformation(ByVal myKey As String, ByVal theValue As String)
    dicSheetParams.Add UCase(myKey), theValue
End Sub

Public Function langs() As Variant()
    langs = dicLangs.Items()
End Function

Public Function langsColumns() As Variant()
    langsColumns = dicLangs.keys
End Function

Public Function lang(ByVal n As String) As String
    lang = dicLangs.Item(n)
End Function

Public Function dicColLang() As Object
    Set dicColLang = dicLangs
End Function

Public Function LangColumn(ByVal myKey As String) As Long
    If dicLangs.Exists(myKey) Then
        LangColumn = CLng(dicLangs.Item(myKey))
    Else
        LangColumn = 0
    End If
End Function



Public Function Item(ByVal myKey As String) As String
    If dicSheetParams.Exists(myKey) Then
        Item = dicSheetParams.Item(myKey)
    Else
        Item = ""
        Debug.Print Now, "No Sheet Item:" & myKey
    End If
End Function


Public Function ExistsLangColumn(ByVal mycol As Long) As Boolean
    ExistsLangColumn = dicLangs.Exists(CStr(mycol))
End Function


Public Function GetDicPic() As Object
    Set GetDicPic = dicPicType
End Function

Private Sub Class_Initialize()
    Call InitParam
End Sub

Public Function makeLanguageFolders(ByVal ppathfull As String) As String()

    Dim langs() As Variant
    Dim LangColumns() As Variant
    
    Dim folders As clsString
    Set folders = New clsString
    
    makeLanguageFolders = folders.MakeArray()

    langs = Me.langs()
    LangColumns = Me.langsColumns()

    Dim il As Long
    Dim langPath As String
    Dim imagePath As String
    For il = 0 To UBound(langs)
        langPath = ppathfull & "\" & langs(il)
        If modCommon.FsoFolderExists(langPath) Then
            ' no ope
        Else
            CreateDirectory langPath
        End If
        folders.Add langPath
    Next il
    
    makeLanguageFolders = folders.MakeArray()
End Function

Public Function makePicFilePath(ByVal curpath As String, picColumn As String) As String
    Dim t As String
    If dicLangs.Exists(picColumn) Then
        t = curpath & "\" & Me.dicColLang.Item(picColumn)
    Else
        t = curpath
    End If
    
    Dim s As String
    s = Me.Item("subFolder")
    If s <> "" Then
        t = t & "\" & s
    End If

    s = Me.Item("imagefolder")
    If s <> "" Then
        t = t & "\" & s
    End If
    
    makePicFilePath = t
End Function

Public Function GetOutDirFullPath() As String
    Dim s As String
    Dim outS As String
    s = Me.Item("outDir")
    If s = "" Then
        outS = ActiveWorkbook.Path
    End If

    If InStr(s, ":") Then
        outS = s
    Else
        outS = ActiveWorkbook.Path & "\" & s
    End If

    GetOutDirFullPath = outS
End Function

Public Function GetOutPicDirFullPath() As String
    Dim s As String
    s = Me.Item("outPicDir")
    If s = "" Then
        s = Me.GetOutDirFullPath()
    End If
    GetOutPicDirFullPath = s
End Function

Public Function ImageDirFullPathWithSlash(picColumn As Long) As String
    Dim s As String
    s = GetOutPicDirFullPath

    Dim picDir As String
    picDir = ActiveSheet.Cells(LangRow, picColumn)
    
    Dim imgDir As String
    imgDir = Me.Item("imagefolder")

    ImageDirFullPathWithSlash = s & IIf(picDir = "", "", "\" & picDir) & "\" & IIf(imgDir = "", "", imgDir & "\")

End Function
