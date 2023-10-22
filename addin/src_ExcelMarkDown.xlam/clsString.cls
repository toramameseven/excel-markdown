VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private stringCollection As Collection
Public PathToSave As String
Public SectionName As String
Public CurrentWeight As Long

Private Sub Class_Initialize()
    Set stringCollection = New Collection
End Sub

Public Sub Add(ByRef rData As Variant)
    If stringCollection.Count = 0 And CStr(rData) = "" Then
        Exit Sub
    End If
    stringCollection.Add CStr(rData)
End Sub

Public Sub AddArray(ByRef vData() As String)
    Dim i As Long
    For i = 0 To UBound(vData)
        stringCollection.Add vData(i)
    Next i
End Sub

Public Sub AddArrayV(ByRef vData())
    Dim i As Long
    For i = 0 To UBound(vData)
        stringCollection.Add CStr(vData(i))
    Next i
End Sub

Public Function Item(ByRef vItem As Long) As String
    If vItem > 0 Then
        Item = stringCollection.Item(vItem)
    Else
        Item = "##################################### item error########################################"
    End If
End Function

Public Sub Remove(ByRef vItem As Long)
    stringCollection.Remove (vItem)
End Sub

Public Function Count() As Long
    Count = stringCollection.Count
End Function

Public Function AddClsString(ByRef c As clsString)
    Dim i As Long
    For i = 1 To c.Count
        stringCollection.Add c.Item(i)
    Next i
End Function

Public Function MakeArray() As String()
    Dim myStrings() As String
    Dim i As Long

    If stringCollection.Count > 0 Then
        ReDim myStrings(0 To stringCollection.Count - 1)
        For i = 0 To UBound(myStrings)
            myStrings(i) = stringCollection.Item(i + 1)
        Next i
    Else
        ReDim myStrings(0 To 0)
    End If
    MakeArray = myStrings
End Function


Public Sub ClearAll()
    Call Class_Initialize
End Sub


Public Function Joins(Optional ByVal vSep As String = vbCrLf) As String
    Joins = Join(Me.MakeArray, vSep)
End Function

Public Function LoadFromFile(fn As String) As Boolean

    Dim LineNum&, ItemCap As Integer
    Dim fs As Variant
    Dim FSO As Object

    On Error GoTo OpenError
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fs = FSO.OpenTextFile(fn, 1, False)

    Set stringCollection = New Collection
    Do While fs.AtEndOfStream <> True
        stringCollection.Add fs.ReadLine
    Loop
    Call fs.Close
    LoadFromFile = True
    GoTo ExitThis

OpenError:
    LoadFromFile = False

ExitThis:
    Set fs = Nothing
    Set FSO = Nothing
End Function

Public Function LoadUtf8(fn As String) As Boolean

    On Error GoTo LoadUtf8_Error
    Dim buf() As String
    Dim obj As Object
    Set obj = CreateObject("ADODB.Stream")

    With obj
        .Charset = "UTF-8"
        .Open
        .LoadFromFile fn
        buf() = Split(.ReadText, vbCrLf)
        .Close
    End With
    Me.AddArray buf

    On Error GoTo 0
    LoadUtf8 = True
    Exit Function

LoadUtf8_Error:
    LoadUtf8 = False
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadUtf8, line " & Erl & "."
    obj.Close
End Function

Public Function SaveToFile(fn As String) As Boolean
    Dim i&, LineNum&
    Dim FSO As Object
    Dim fs As Variant

    On Error GoTo FileErr
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fs = FSO.OpenTextFile(fn, 2, True)
    LineNum = 0
    For i = 1 To stringCollection.Count
        Call fs.WriteLine(stringCollection.Item(i))
    Next i
    Call fs.Close
    SaveToFile = True
    GoTo ExitThis

FileErr:
    fs.Close
    SaveToFile = False

ExitThis:
    Set fs = Nothing
    Set FSO = Nothing
End Function

Public Function SaveToFileUTF8(fn As String) As Boolean
    'Dim writeStream As ADODB.Stream
    'Microsoft ActiveX Data Objects 2.5 Library
    Dim writeStream As Object
    On Error GoTo saveUtf8_Error

    If stringCollection.Count = 0 Then
        Exit Function
    End If

    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Charset = "UTF-8"
    writeStream.Open

    ' output to buffer.
    Dim i As Long
    For i = 1 To stringCollection.Count
        writeStream.WriteText stringCollection.Item(i), 1 ' 1:add new line.
    Next i

    'bom
    writeStream.Position = 0 '
    writeStream.Type = 1 'adTypeBinary:1    default, Charset text data:2
    writeStream.Position = 3 ' set stream position 3.

    Dim byteData() As Byte ' temporary
    byteData = writeStream.Read 'set stream to temporary
    writeStream.Close 'close stream(reset)

    writeStream.Open
    writeStream.Write byteData

    writeStream.SaveToFile fn, 2 'adSaveCreateOverWrite:2

    On Error GoTo 0
    GoTo saveUtf8_Normal_exit

saveUtf8_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveUtf8 of Class Module clsStrings file:" & fn

saveUtf8_Normal_exit:
    writeStream.Close
    Set writeStream = Nothing
End Function



Public Sub convertArabicNumSubStep()
    Dim ss() As String
    ss = Split(Me.Joins(vbCrLf), vbCrLf)

    Dim outs As clsString
    Set outs = New clsString

    Dim i As Long
    Dim j As Long
    Dim cnt As Long

    For i = 0 To UBound(ss)
        If Left(trim(ss(i)), 3) = ".. " Then
            ss(i) = "[arabic]" & vbCrLf & ss(i)
        End If
        outs.Add ss(i)
    Next i


    Me.ClearAll
    Me.AddClsString outs
End Sub

