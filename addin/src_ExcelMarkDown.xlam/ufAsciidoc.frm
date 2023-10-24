VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAsciidoc 
   Caption         =   "ShapeExcelMarkDown2"
   ClientHeight    =   6576
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   5976
   OleObjectBlob   =   "ufAsciidoc.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufAsciidoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit

Private Sub chacolorZero()
    If TypeName(Selection) = "Range" Then
       ''
    Else
        Exit Sub
    End If
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

Private Sub chacolor(ByVal chc As Long)
    If TypeName(Selection) = "Range" Then
       ''
    Else
        Exit Sub
    End If
    With Selection.Font
        .Color = chc
        .TintAndShade = 0
    End With


'    With Selection.Font
'        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'    End With
'    With Selection.Font
'        .Color = -16776961
'        .TintAndShade = 0
'    End With
'    Range("G20").Select
'    With Selection.Font
'        .Color = -1003520
'        .TintAndShade = 0
'    End With
'    Range("G23").Select
'    With Selection.Font
'        .Color = -16711681
'        .TintAndShade = 0
'    End With
End Sub


Public Sub AddSpaceUpDownSection()

  
    Dim i As Long
    
    Dim DataEndRow As Long
    DataEndRow = modAsciidoc.getEndLine(1, 20)
    

    Dim FlagCell As String
    Dim ValueCellNext As String
    Dim ValueCellBefore As String
    Dim PicCellBefore As String
    Dim PicCellNext As String
    
    
    
    For i = 10 To DataEndRow + 100
        FlagCell = Cells(i, 1)
        PicCellNext = Cells(i + 1, 2)
        PicCellBefore = Cells(i - 1, 2)
        ValueCellNext = Cells(i + 1, 3)
        ValueCellBefore = Cells(i - 1, 3)
        
        If Left(FlagCell, 1) = "=" Then
            If PicCellNext <> "" Or ValueCellNext <> "" Then
                rows(i + 1).Insert
            End If
            
            If PicCellBefore <> "" Or ValueCellBefore <> "" Then
                rows(i).Insert
                i = i + 1
            End If
        End If
        
    Next i
End Sub

Private Sub btnAddSpaceSection_Click()
  AddSpaceUpDownSection
End Sub

Private Sub btnbkcolorblue_Click()
    bkcolor 15773696
End Sub

Private Sub btnbkcolorred_Click()
    bkcolor 255
End Sub

Private Sub btnbkcoloryellow_Click()
    bkcolor 65535
End Sub

Private Sub btnbkcolorzero_Click()
    Call bkcolorZero
End Sub

Private Sub btnBoldLine_Click()
    With Selection.ShapeRange.Line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .weight = 1.5
    End With
End Sub



Private Sub btnCalcCols_Click()
    Dim s As String
    s = GetColumnsWidthsForAdoc(Selection)
    Dim r As Range
    Set r = Selection
    If r.Resize(1, 1).Offset(-1) = "" Then
        r.Resize(1, 1).Offset(-1) = s
    End If
End Sub

Private Sub btnchrblue_Click()
    chacolor -1003520
End Sub

Private Sub btnchrred_Click()
    chacolor -16776961
End Sub

Private Sub btnchryellow_Click()
    chacolor -16711681
End Sub

Private Sub btnchrzero_Click()
    chacolorZero
End Sub

Public Sub DeleteSectionNos(rRng As Range)
    Dim x As Long
    Dim y As Long

    For x = 1 To rRng.rows.Count
        For y = 1 To rRng.columns.Count
            If rRng.Cells(x, y) <> "" Then
                rRng.Cells(x, y) = modReg.DeleteSectionNo(rRng.Cells(x, y))
            End If
        Next y
    Next x
End Sub


Private Sub btnDelSecNo_Click()
    DeleteSectionNos Selection
End Sub

Private Sub btnNormalLine_Click()
   
    Selection.ShapeRange.Line.Visible = msoFalse
End Sub

Private Sub btnOutText_Click()
    Dim rRng As Range
    
    Set rRng = Selection.Resize(2, 3)
    
    Dim data() As Variant
    
    data = rRng.value
    
    Dim i As Long
    Dim j As Long
    
    Dim t As clsString
    Set t = New clsString
    Dim a As clsString
    Set a = New clsString
    
    Dim endfile As clsString
    Set endfile = New clsString
    
    
    For i = 1 To UBound(data, 1)
        Set t = New clsString
        For j = 1 To UBound(data, 2)
            't.add """" & data(i, j) & """"
            t.Add data(i, j)
        Next j
        a.Add t.Joins(",")
    Next i
    
    
    
    Dim fileName As String
    fileName = ActiveWorkbook.Path & "\" & Format(Now, "YYYYMMDDHHMMSS")
    
    a.SaveToFileUTF8 fileName & ".csv"
    a.SaveToFileUTF8 fileName & ".end"
    
End Sub

Private Sub btnPutPicName_Click()
    modAsciidoc.SetSetPicNameCellFromCellAddress TextPrefixPic.Text
End Sub

Private Sub btnSelAllobj_Click()
    '' modAsciidoc.selectallpic
    ActiveSheet.shapes.SelectAll
End Sub

Private Sub btnUpdateIndex_Click()
    ''
End Sub

Private Sub buttonSetChapter_Click()
    modAsciidocTools.SetChapter
End Sub

Private Sub cmdAddTableTags_Click()
    Call InsertTableTags
End Sub


Private Sub cmdBeginTable_Click()
    modAsciidocTools.InsertTableTagBegin
End Sub

Private Sub cmdCreateAdoc_Click()
    modAsciidoc.MakeDocument "adoc"
End Sub

Private Sub cmdCreateMD_Click()
    modAsciidoc.MakeDocument "md"
End Sub

Private Sub cmdCreateSD_Click()
    modAsciidoc.MakeDocument "sp"
End Sub

Private Sub cmdCreateTextile_Click()
    modAsciidoc.MakeDocument "textile"
End Sub

Private Sub cmdEndTable_Click()
    modAsciidocTools.InsertTableTagEnd
End Sub

Private Sub cmdEditProp_Click()
'    Dim shapeType As Long
'    shapeType = IIf(IsNumeric(Me.txtShapeType), Me.txtShapeType, 0)
'    modChartExcel.repalacePropMain shapeType, Me.txtShapeColor, 0
End Sub

Private Sub cmdExcelChart_Click()
    Dim shapeType As String
    shapeType = Me.txtChartNo.Text
    Dim shapeTypeNum As String
    shapeTypeNum = Split(shapeType, ",")(0)
    
    Dim shapeTypeLng As Long
    
    shapeTypeLng = IIf(IsNumeric(shapeTypeNum), CLng(shapeTypeNum), 0)
    modChartExcel.repalacePropMain shapeTypeLng, Me.txtShapeColor, 0
    
    
   
    If IsNumeric(shapeTypeNum) Then
        Call modChartExcel.makeShapes(shapeTypeLng)
    Else
        Call modChartExcel.makeShapes
    End If
End Sub

Private Sub cmdIndentDown_Click()
    modAsciidoc.selectionHeadingUpDown (1)
End Sub

Private Sub cmdIndentUp_Click()
   modAsciidoc.selectionHeadingUpDown (-1)
End Sub

Private Sub cmdIndex_Click()
    'frmIndex.Show vbModeless
End Sub

Private Sub cmdListNum_Click()
    modAsciidocTools.InsertFalgInsertLineUpDown "."
End Sub



Private Sub cmdLineAbove_Click()
    Call modAsciidocTools.InsertLineAboveBelow
End Sub

Private Sub cmdOpenChrome_Click()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = ActiveWorkbook.Path
        .AllowMultiSelect = False
        .Filters.Add "docment", "*.adoc; *.md; *.textile", 1
    
        .title = "select open file"
        If .Show = True Then
            filePath = .SelectedItems(1)
        End If
    End With
    
 'open Google Chrome and go to the URL
 CreateObject("WScript.Shell").Run _
 ("chrome.exe -url " & "file://" & filePath)
    
End Sub

Private Sub cmdPrePost_Click()
    Selection.Text = txtPre.Text & Selection.Text & txtPost.Text
End Sub

Private Sub cmdR2R_Click()

End Sub

Private Sub cmdRedBox_Click()
    modAsciidoc.CreateRedBox
End Sub

Private Sub cmdGroup_Click()
    Call ShapeGrouping
End Sub

Private Sub cmdLint_Click()
    Call modAsciidoc.Lint
End Sub

Private Sub cmdLoadPic_Click()
    
    Dim fname As String
    fname = Selection
    
    Dim loadrow As Long
    
    If fname = "" Then
        loadrow = 0
    Else
        loadrow = Selection.row
    End If
    
    
    If MsgBox("Load pictures", vbOKCancel) = vbOK Then
        Call modAsciidoc.LoadPictures(loadrow)
    End If
End Sub



Private Sub cmdMakeTemplate_Click()
    modAsciidoc.CreateTemplate
End Sub

Private Sub cmdOpenCode_Click()
    Dim vscode As String
    'vscode = """C:\Program Files\Microsoft VS Code\Code.exe"" "
    vscode = """Code.exe"" "
    Dim fileName As String
    
    If Cells(1, 3) <> "" Then
        fileName = ActiveWorkbook.Path & "\" & Cells(2, 3) & "\" & ActiveSheet.Name & ".adoc"
    Else
        fileName = ActiveWorkbook.Path & "\" & ActiveSheet.Name & ".adoc"
    End If
    
    Dim s As Long
    s = Shell(vscode + ActiveWorkbook.Path, vbNormalFocus)
End Sub

Private Sub cmdOpenfolder_Click()
    modCommon.OpenFolder ActiveWorkbook.Path
End Sub


Private Sub cmdRefreshSize_Click()
    modAsciidoc.RefreshImageSize
End Sub

Private Sub cmdResize_Click()
    Call PicResizeHight250
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PicResize
' Author    : toramame
' Date      : 2017/01/31
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub PicResizeHight250()
    On Error GoTo PicResize_Error
    
    Selection.Placement = xlMove
    Selection.ShapeRange.LockAspectRatio = msoTrue
    Selection.ShapeRange.Height = 250

    On Error GoTo 0
    GoTo PicResize_Normal_exit

PicResize_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PicResize of Module modPicture"

PicResize_Normal_exit:
End Sub

Private Sub cmdSavePic_Click()
    On Error GoTo cmdSavePic_Click_Error
    modAsciidoc.SavePicturesWork False
    
    
    On Error GoTo 0
    Exit Sub

cmdSavePic_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSavePic_Click, line " & Erl & "."
End Sub



Private Sub cmdSection2_Click()
    modAsciidocTools.InsertSectionTag 2
End Sub

Private Sub cmdSection3_Click()
    modAsciidocTools.InsertSectionTag 3
End Sub

Private Sub cmdSection4_Click()
    modAsciidocTools.InsertSectionTag 4
End Sub

Private Sub cmdSection5_Click()
    modAsciidocTools.InsertSectionTag 5
End Sub

Private Sub cmdSetLinePic_Click()
    modPicture.SetShapeLineRound
End Sub

Private Sub cmdSpaceArr_Click()
    modAsciidoc.AddSpaceUnderPicture
End Sub

Private Sub cmdNote_Click()
    modAsciidocTools.InsertFalgInsertLineUpDown "NOTE:"
End Sub

Private Sub cmdSplitKuten_Click()
    modAsciidocTools.splitKuten
End Sub

Private Sub cmdTenDot_Click()
    modAsciidocTools.InsertFalgInsertLineUpDown "*"
End Sub

Private Sub cmdxlmove_Click()
    modAsciidoc.SetImageAddxlMove
End Sub

Private Sub cmdxlsArrange_Click()
    modAsciidoc.ArrangeFormatXls
End Sub

Private Sub CommandButton1_Click()
    'modKakariuke.outKakariukePic
End Sub



Private Sub cmdOpenFolder2_Click()
    modCommon.OpenFolder Selection.Text
End Sub

Private Sub cmdSetRowColors_Click()
    modAsciidocTools.SetRowsColor
End Sub



Private Sub bkcolorZero()
    If TypeName(Selection) = "Range" Then
       ''
    Else
        Exit Sub
    End If
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Private Sub bkcolor(ByVal colorid As Long)

    If TypeName(Selection) = "Range" Then
       ''
    Else
        Exit Sub
    End If
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = colorid
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


'    Range("H7").Select
'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 255
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'    Range("H9").Select
'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 5287936
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'    Range("H11").Select
'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 65535
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
End Sub


Private Sub cmdLonginput_Click()
    LongInput
End Sub
Private Sub LongInput()
    Dim collectionCells As Collection

    Set collectionCells = GetCells(Selection, 5)
    Dim i As Long
    For i = 1 To collectionCells.Count
        Dim myForm As frmLongInput2
        Set myForm = New frmLongInput2
        Set myForm.myRng = collectionCells.Item(i)
        myForm.Show vbModeless
    Next i
End Sub

Private Function GetCells(ByRef rRng As Range, Optional maxCount = 1000) As Collection
    Dim rngSelect As Range
    Set rngSelect = rRng

    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim xMax As Long
    Dim yMax As Long

    Dim cellcollection As Collection
    Set cellcollection = New Collection


    For i = 1 To rngSelect.Areas.Count
        If i > maxCount Then
            Exit For
        End If

        xMax = rngSelect.Areas(i).columns.Count
        yMax = rngSelect.Areas(i).rows.Count

        For x = 1 To xMax
            For y = 1 To yMax
                If rngSelect.Areas(i).Cells(y, x) <> "" Then
                    cellcollection.Add (rngSelect.Areas(i).Cells(y, x))
                End If
            Next y
        Next x
    Next i
    Set GetCells = cellcollection
End Function


Private Sub CommandButton6_Click()

End Sub





Private Sub UserForm_Activate()
    txtFile.Text = Application.ActiveWorkbook.Name
End Sub

Private Sub UserForm_Initialize()
    Call txtChartNo.Clear
    Call txtChartNo.AddItem("5, square")
    Call txtChartNo.AddItem("74,homeBase")
    Call txtChartNo.AddItem("51, |=>")
    Call txtChartNo.AddItem("9,oval")
    Call txtChartNo.AddItem("71, manualInput")
    Call txtChartNo.AddItem("67, document")
'    Call txtChartNo.AddItem("")
'    Call txtChartNo.AddItem("")
'    Call txtChartNo.AddItem("")
'    Call txtChartNo.AddItem("")
'    Call txtChartNo.AddItem("")
End Sub

