VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private logs As Collection
Private errs As Collection

Public sheetName As String
Public Address As String


Private Sub Class_Initialize()
    Set logs = New Collection
    Set errs = New Collection
End Sub

Public Sub AddLog(ByVal stname As String, ByVal address1 As String, ByVal msg1 As String)
    If stname = "" Then stname = sheetName
    If address1 = "" Then address1 = Address
    logs.Add msg1 & "|" & address1 & "|" & stname
End Sub

Public Sub AddErr(ByVal stname As String, ByVal address1 As String, ByVal msg1 As String)
    If stname = "" Then stname = sheetName
    If address1 = "" Then address1 = Address
    errs.Add msg1 & "|" & address1 & "|" & stname
End Sub

Public Sub OutputLogs(Optional ByVal txtFormat = "|1", Optional ByVal isHyperlink As Boolean = False)
    If txtFormat = "" Then
        txtFormat = "|1"
    End If
    outputcommon 0, txtFormat, isHyperlink
End Sub

Public Sub OutputErrs(Optional ByVal txtFormat = "|3-|2 : |1", Optional ByVal isHyperlink As Boolean = True)
    If txtFormat = "" Then
        txtFormat = "|3-|2 : |1"
    End If
    outputcommon 1, txtFormat, isHyperlink
End Sub


'############################################################################################

Private Sub outputcommon(ByVal logtype As Long, ByVal txtFormat, ByVal isHyperlink As Boolean)
    Dim stname As String
    Dim logclass As Collection

    If logtype = 0 Then
        stname = "logs"
        Set logclass = logs
    ElseIf logtype = 1 Then
        stname = "errs"
        Set logclass = errs
    Else
        Exit Sub
    End If

    If isExistSheet(stname) Then
        '
    Else
        ActiveWorkbook.Sheets.Add.Name = stname
    End If

    '' output errors and logs
    Dim ii As Long

    '' output errs to worksheet
    Worksheets(stname).Cells.Clear
    Dim msglist() As String
    Dim disptxt As String
    For ii = 1 To logclass.Count
        msglist = Split(logclass.Item(ii), "|")

        disptxt = Replace(txtFormat, "|1", msglist(0))
        disptxt = Replace(disptxt, "|2", msglist(1))
        disptxt = Replace(disptxt, "|3", msglist(2))

        If isHyperlink Then
            ActiveSheet.Hyperlinks.Add Anchor:=Worksheets(stname).Cells(ii, 1), Address:="", SubAddress:="'" & msglist(2) & "'!" & msglist(1), TextToDisplay:=disptxt
        Else
            Worksheets(stname).Cells(ii, 1) = disptxt
        End If
    Next

    If errs.Count > 0 Then
        Worksheets(stname).Select
        Cells(1, 1).Select
    End If
End Sub

Private Function isExistSheet(ByVal tsheetname) As Boolean
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        If UCase(st.Name) = UCase(tsheetname) Then
            isExistSheet = True
            Exit Function
        End If
    Next
    isExistSheet = False
End Function





