Attribute VB_Name = "modAsciidoc"
Option Explicit


Private Const lang As String = "LANG"
Private Const IMAGE_PATH = "imageFolder"
Private Const NORMAL_FONT_SIZE As Long = 11
Private Const FLG_COL As Long = 1
Private Const PIC_COL As Long = 2
Private Const VALUE_COL As Long = 3

Private Const MAX_ROW As Long = 500

Private Const PIC_OFFSET_X As Long = 1
Private Const PIC_OFFSET_Y As Long = 0

Public Const SECTION_TAG As String = "#"
Public Const SECTION_TAG2 As String = "##"
Public Const SECTION_TAG3 As String = "###"
Public Const SECTION_TAG4 As String = "####"
Public Const SECTION_TAG5 As String = "#####"

'(section 0 to 9, language 0 to 2)
Private m_sections(1 To 10, 0 To 2) As Long
Private m_numList(0 To 2) As Long

Private globalBeforeParam As String
Private moduleLogs As clsLogs


Private Sub ResetSections()
    m_numList(0) = 0
    m_numList(1) = 0
    m_numList(2) = 0
    
    Dim i As Long
    For i = 1 To UBound(m_sections)
        m_sections(i, 0) = 0
        m_sections(i, 1) = 0
        m_sections(i, 2) = 0
    Next i
End Sub

'' sectionPosition: 1.2.3.4...10
'' index: language id
Private Function NewSectionNumber(ByVal sectionPosition As Long, ByVal index As Long) As String
    Dim i As Long
    
    m_sections(sectionPosition, index) = m_sections(sectionPosition, index) + 1
    
    '' clear under sectionPosition
    For i = sectionPosition + 1 To UBound(m_sections, 1)
        m_sections(i, index) = 0
    Next i
    
    '' clear list index
    m_numList(index) = 0
    
    '' make section No
    Dim s As String
    s = ""
    For i = 1 To sectionPosition
       s = s & m_sections(i, index) & "."
    Next i
    NewSectionNumber = s & " "
End Function

'' i: language number, mum: when 0, reset
Private Function NewListNum(ByVal i As Long, Optional ByVal num As Long = 1) As String
    If num = 0 Then
        m_numList(i) = 0
    End If
    
    m_numList(i) = m_numList(i) + 1
    NewListNum = m_numList(i) & ". "
End Function

Public Sub CreateTemplate()
    'select first sheet
    Dim sh As Worksheet
    For Each sh In ActiveWindow.SelectedSheets
        sh.Select
        Exit For
    Next
       
    If WorksheetFunction.CountA(Range(Cells(1, 1), Cells(100, VALUE_COL))) > 0 Then
        MsgBox "cells are not empty for the template"
        Exit Sub
    End If

    
    Dim infoRow As Long
    infoRow = infoRow + 1
    Cells(infoRow, FLG_COL) = "FLG"
    Cells(infoRow, PIC_COL) = "PIC"
    Cells(infoRow, VALUE_COL) = "VALUE_COL"
    
    infoRow = infoRow + 1
    Cells(infoRow, 1) = "LANG"
    Cells(infoRow, 3) = "ja"

    infoRow = infoRow + 1
    Cells(infoRow, 1) = "subFolder"
    Cells(infoRow, 3) = ""
    
    infoRow = infoRow + 1
    Cells(infoRow, 1) = "imageFolder"
    Cells(infoRow, 3) = ""
    
    infoRow = infoRow + 1
    Cells(infoRow, 1) = "IsSectionNo"
    Cells(infoRow, 3) = "true"
    
    infoRow = infoRow + 1
    Cells(infoRow, 1) = "IsSectionSplit"
    Cells(infoRow, 3) = "false"
    
    infoRow = infoRow + 1
    Cells(infoRow, 1) = "mdOffset"
    Cells(infoRow, 3) = "1"
    
    
    infoRow = infoRow + 1
    Cells(infoRow, 1) = "OutDir"
    Cells(infoRow, 3) = ""

    infoRow = infoRow + 1
    Cells(infoRow, 1) = "OutPicDir"
    Cells(infoRow, 3) = ""

    infoRow = infoRow + 1
    Cells(infoRow, 1) = "fileName"
    Cells(infoRow, 3) = ""

    infoRow = infoRow + 1
    Cells(infoRow, 1) = "Enable"
    Cells(infoRow, 3) = "TRUE"
    
    'AB set to string format, and ime off
    columns("A:B").Select
    columns("A:B").NumberFormatLocal = "@"
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeOff
        .ShowInput = True
        .ShowError = True
    End With
End Sub


Public Sub GetSectionListOnlyTitle_Test()
    Call GetSectionListOnlyTitle
End Sub

Public Function GetSectionListOnlyTitle(Optional ByVal upTo As Long = 3) As String

        Dim sh As Worksheet
        Set sh = ActiveSheet

        '' Create sheet info
        Dim sheetInfo As clsAdocSheetInfo
        Set sheetInfo = New clsAdocSheetInfo
        Call sheetInfo.SetSheetInfo(sh, FLG_COL, VALUE_COL, VALUE_COL)
        
        
        Dim langsColumns() As Variant
        langsColumns = sheetInfo.langsColumns
        
        
        
        Dim rowStart As Long
        Dim rowEnd As Long
        rowStart = GetContentStartRow(sh)
        rowEnd = getEndLine(FLG_COL, VALUE_COL)
        
        
        Dim sections As clsString
        Set sections = New clsString
        
        Dim sectionFileName As String

        
        Dim i As Long
        Dim linkTitle As String
        
        For i = rowStart To rowEnd
            linkTitle = ""
            Dim flg As String
            Dim pic As String
            Dim cellValue As String
            
            ' Dim flgB As String
            ' Dim picB As String
            ' Dim cellValueB As String

            flg = Cells(i, FLG_COL)
            If Left(flg, 1) = SECTION_TAG Then
                ' flgB = Cells(i - 1, FLG_COL)
                If getSectionIndexes(flg) <= upTo Then
                    cellValue = TrimLeft(Cells(i, VALUE_COL + 0))
                    sections.Add flg & cellValue & "\\" & Cells(i, FLG_COL).Address
                End If
            End If
        Next
    GetSectionListOnlyTitle = sections.Joins(vbCr)
End Function


Private Function getSectionIndexes(sectionTags As String) As Long
    Dim s As String
    s = Replace(sectionTags, SECTION_TAG, "")
    getSectionIndexes = Len(sectionTags) - Len(s)
End Function


Private Function has2byteString(s As String) As Boolean
  If Len(s) <> LenB(StrConv(s, vbFromUnicode)) Then
    has2byteString = True
  Else
    has2byteString = False
  End If
End Function



Public Function Lint() As String

    Dim myLogs As clsLogs
    Set myLogs = New clsLogs

    Dim sh As Worksheet
    Set sh = ActiveSheet

    '' Create sheet info
    Dim sheetInfo As clsAdocSheetInfo
    Set sheetInfo = New clsAdocSheetInfo
    Call sheetInfo.SetSheetInfo(sh, FLG_COL, VALUE_COL, VALUE_COL)
    
    
    Dim langsColumns() As Variant
    langsColumns = sheetInfo.langsColumns
    
    
    
    Dim rowStart As Long
    Dim rowEnd As Long
    rowStart = GetContentStartRow(sh)
    rowEnd = getEndLine(FLG_COL, VALUE_COL)
    
    
    Dim sections As clsString
    Set sections = New clsString
    
    Dim sectionFileName As String

    
    Dim i As Long
    For i = rowStart To rowEnd
        Dim flg As String
        Dim pic As String
        Dim cellValue As String
        

        ''/////////////////////
        flg = Cells(i, FLG_COL)
        Dim ii As Long
        Dim columnOffset As Long
        
        'text
        If Left(flg, 1) = SECTION_TAG Or Left(flg, 1) = "." Or Left(flg, 1) = "*" Or _
           ((UCase(flg) = "NOTE:" Or UCase(flg) = "TIP:" Or UCase(flg) = "IMPORTANT:" Or UCase(flg) = "CAUTION:" Or UCase(flg) = "WARNING:")) Then
            For ii = 0 To UBound(langsColumns)

                columnOffset = CLng(langsColumns(ii)) - VALUE_COL
                
                cellValue = TrimLeft(Cells(i, VALUE_COL + columnOffset))
                
                If cellValue = "" Then
                    myLogs.AddErr ActiveSheet.Name, Cells(i, VALUE_COL + columnOffset).Address, flg & ":Empty Cell."
                End If

            Next
        End If
        
        'Pic
        Dim noPicCell As clsString
        Set noPicCell = New clsString
        Dim picName As String
        
        For ii = 0 To UBound(langsColumns)
        
            columnOffset = CLng(langsColumns(ii)) - VALUE_COL
            picName = Cells(i, PIC_COL + columnOffset)
            If picName = "" Then
                noPicCell.Add CStr(PIC_COL + columnOffset)
            Else
                If modCommon.FileExists(sheetInfo.ImageDirFullPathWithSlash(PIC_COL + columnOffset + 1) & picName) = False Then
                    myLogs.AddErr ActiveSheet.Name, Cells(i, PIC_COL + columnOffset).Address, ":No Image " & picName
                End If
            End If
        Next
        
        'Pic Name check
        If noPicCell.Count = 0 Or noPicCell.Count = UBound(langsColumns) + 1 Then
            'ok
        Else
            For ii = 1 To noPicCell.Count
                myLogs.AddErr ActiveSheet.Name, Cells(i, CLng(noPicCell.Item(ii))).Address, ":No Image"
            Next
        End If
    ''/////////////////////
    Next
    
        'no pic name check
    
    Dim sp As Shape
    For Each sp In ActiveSheet.shapes
        If sp.Type <> msoComment Then
            Dim pic1 As String, pic2 As String
            pic1 = sp.TopLeftCell.Offset(-1, -1).Text
            pic2 = sp.TopLeftCell.Offset(0, -1).Text

            If pic1 = "" And pic2 = "" Then
                myLogs.AddErr ActiveSheet.Name, sp.TopLeftCell.Address, ":No Image File"
            ElseIf pic1 <> "" And pic2 <> "" Then
                myLogs.AddErr ActiveSheet.Name, sp.TopLeftCell.Address, ":Multi Image Name"
            Else
            
            End If
        End If
    Next
    myLogs.OutputErrs "", True
End Function


''##########################################################################################################################################################
Public Function MakeDocument(ByVal tType As String, Optional ByVal makeOption As String = "") As String
    Set moduleLogs = New clsLogs
    
    Dim allDocuments As clsAllAdoc
    Set allDocuments = New clsAllAdoc
        
    Dim isBeforeNull As Boolean
    isBeforeNull = False
    
    Dim isListContinue As Boolean  ''
    Dim flg As String
    Dim pic As String
    Dim cellValue As String
    
    Dim flgB As String
    Dim picB As String
    Dim cellValueB As String
    
    '' section no is for all sheet
    Call ResetSections
    Dim PathParentWithBS As String

    
    ''///////////////////sheets////////////////////////////////////////////////////////////////////////
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    
    Dim sh As Worksheet
    For Each sh In ActiveWindow.SelectedSheets
        sheetCollection.Add sh
    Next
    
    Dim sheetNo As Long
    sheetNo = 0

    ' for 3 language doc
    Dim langDocs(0 To 2) As clsString
    Set langDocs(0) = New clsString
    Set langDocs(1) = New clsString
    Set langDocs(2) = New clsString

    Dim splitFileName As String

    For Each sh In sheetCollection
        sheetNo = sheetNo + 1
    
        Dim ContentStartRow As Long
        ContentStartRow = GetContentStartRow(sh)
        sh.Select
        sh.Activate

        moduleLogs.sheetName = sh.Name
        
        '' Create sheet info
        Dim sheetInfo As clsAdocSheetInfo
        Set sheetInfo = New clsAdocSheetInfo
        Call sheetInfo.SetSheetInfo(sh, FLG_COL, VALUE_COL, VALUE_COL)
        
        
        Dim langsColumns() As Variant
        langsColumns = sheetInfo.langsColumns
        Dim langPath As String
        Dim sectionNoForInc As Long
        

        Dim isSplitSection As Boolean
        isSplitSection = (LCase(sheetInfo.Item("IsSectionSplit")) = "true")
        
        '' outdir\(lang)\subdir\xxx.md
        Dim outDir As String
        outDir = sheetInfo.GetOutDirFullPath()
        
        'make sub folders
        Dim folders() As String
        folders = sheetInfo.makeLanguageFolders(outDir)
        
        
        ' if  selection is range
        If TypeName(Selection) = "Range" Then
            '' go next
        Else
            '' ex. picture is selected.
            sh.Range("A1").Select
        End If
        
        ' if selections is multi rows
        Dim MaxRows As Long
        Dim rowStart As Long
        Dim rowEnd As Long
        Dim selectRange As Range
        Set selectRange = Selection
        rowStart = selectRange(1).row
        rowEnd = selectRange(selectRange.Count).row
        If rowStart = rowEnd Then
            rowStart = ContentStartRow
            MaxRows = getEndLine(FLG_COL, VALUE_COL)
            rowEnd = MaxRows
        Else
            ' do nothing
        End If
        

        ' Loop document in a sheets(langs) ///////////////////////////////////////////////////////////
        Dim ii As Long
        For ii = 0 To UBound(langsColumns)
            Set langDocs(ii) = New clsString
            '' Create sheet info
            Call sheetInfo.SetSheetInfo(sh, FLG_COL, VALUE_COL, langsColumns(ii))
            sectionNoForInc = 0
            Dim columnOffset As Long
            columnOffset = CLng(langsColumns(ii)) - VALUE_COL
            langPath = sheetInfo.lang(langsColumns(ii))
            
            PathParentWithBS = outDir & "\"
            
            If langPath <> "" Then
                PathParentWithBS = PathParentWithBS & langPath & "\"
            End If

            If sheetInfo.Item("subFolder") <> "" Then
                PathParentWithBS = PathParentWithBS & sheetInfo.Item("subFolder") & "\"
                CreateDirectory PathParentWithBS
            End If
            
            Dim fileNameForNoSplit As String
            fileNameForNoSplit = sh.Name
            If sheetInfo.Item("fileName") <> "" Then
                fileNameForNoSplit = sheetInfo.Item("fileName")
            End If

            isListContinue = False
        

            '' work at sheet
            Dim i As Long
            Dim sheetNum As Long
            Dim nextIndex As Long
            For i = rowStart To rowEnd
                moduleLogs.Address = Cells(i, VALUE_COL + columnOffset).Address

                flgB = Cells(i - 1, FLG_COL)
                picB = Cells(i - 1, PIC_COL + columnOffset)
                cellValueB = Cells(i - 1, VALUE_COL + columnOffset)

                flg = Cells(i, FLG_COL)

                If flg = "=" Then Cells(i, FLG_COL) = "#"
                If flg = "==" Then Cells(i, FLG_COL) = "##"
                If flg = "===" Then Cells(i, FLG_COL) = "###"
                If flg = "====" Then Cells(i, FLG_COL) = "####"
                If flg = "=====" Then Cells(i, FLG_COL) = "#####"
 
                pic = Cells(i, PIC_COL + columnOffset)
                cellValue = Cells(i, VALUE_COL + columnOffset)

                '' reload flag
                flg = Cells(i, FLG_COL)
                If Left(flg, 1) = "#" Then
                    flg = GetSectionFlg(flg)
                End If

                '' get split FileName suffix
                Dim isNotDoSuffix As Boolean
                isNotDoSuffix = False
                If LCase(Left(flg, 5)) = "file:" Then
                    isNotDoSuffix = True
                    splitFileName = Mid(flg, 6)
                    fileNameForNoSplit = ""
                End If

                Dim isNotDoTag2 As Boolean
                isNotDoTag2 = False
               
                '' splitSection
                If (isSplitSection) And flg = SECTION_TAG2 Then

                    isNotDoTag2 = True
                    nextIndex = 0
                    '' for asciidoc
                    '' langDocs(ii).convertArabicNumSubStep
                    If langDocs(ii).PathToSave <> "" Then
                        langDocs(ii).SaveToFileUTF8 langDocs(ii).PathToSave
                        fileNameForNoSplit = ""
                    End If
                    
                    allDocuments.AddDocument langDocs(ii), langPath, ""
                    sectionNoForInc = sectionNoForInc + 1

                    '' get fileName for split
                    If splitFileName = "" Then
                        splitFileName = LCase(FindAlphabetRegExp(Cells(i, VALUE_COL)))
                        fileNameForNoSplit = ""
                    End If
                    '' new file  '' langDocs(ii)
                    Set langDocs(ii) = New clsString
                    isListContinue = False
                    langDocs(ii).Add ("---")
                    langDocs(ii).Add ("title: " & cellValue)
                    langDocs(ii).Add ("sidebar:")
                    langDocs(ii).Add ("  label: " & cellValue)
                    langDocs(ii).Add ("  order: " & CStr(sectionNoForInc))
                    langDocs(ii).Add ("---")
                    langDocs(ii).Add ("")

                    langDocs(ii).PathToSave = PathParentWithBS & fileNameForNoSplit & splitFileName & "." & tType
                    splitFileName = ""
                    Debug.Print "filename", langDocs(ii).PathToSave
                End If '' splitSection
                
                
                '' create body  (adoc, textile, markdown)
                Dim resultMakeBody As String

                resultMakeBody = ""
                If isNotDoSuffix Or _
                    isNotDoTag2 Or _
                     (flg = "" And cellValue = "" And pic = "") Or _
                    IsCustomCurrentDocument(tType, flg) = False Or _
                    Left(flg, 2) = "//" Or _
                    LCase(cellValue) = "//empty" Or _
                    LCase(pic) = "//empty" Or _
                    ((UCase(flg) = "NOTE:" Or UCase(flg) = "TIP:" Or UCase(flg) = "IMPORTANT:" Or UCase(flg) = "CAUTION:" Or UCase(flg) = "WARNING:") And cellValue = "") Then
                    '' Do nothing
                Else
                    resultMakeBody = makeBody(i, tType, isListContinue, sheetInfo, columnOffset, ii, nextIndex)
                End If

                '' add results
                If resultMakeBody <> "" Then
                    langDocs(ii).Add resultMakeBody
                    isBeforeNull = False
                End If

                If resultMakeBody = "" And isBeforeNull = False Then
                    langDocs(ii).Add ""
                    isBeforeNull = True
                Else
                    ''Do nothing
                End If
            Next i '' end sheet parse (rows)

            If (LCase(sheetInfo.Item("IsSectionSplit")) = "false") Then
                '' langDocs(ii).convertArabicNumSubStep
                langDocs(ii).SaveToFileUTF8 PathParentWithBS & fileNameForNoSplit & splitFileName & "." & tType
                allDocuments.AddDocument langDocs(ii), langPath, sheetInfo.Item("fileName")
            End If
        Next ii '' language
    Next '' sheets

    
    '' for split files
    For ii = 0 To UBound(langsColumns)
        If langDocs(ii).PathToSave <> "" Then
            langDocs(ii).SaveToFileUTF8 langDocs(ii).PathToSave
        End If
    Next

    '' save sheets markdown
    If sheetCollection.Count > 1 Then
        MakeDocument = allDocuments.saveAll(tType, outDir, sheetInfo.Item("subFolder"))
    End If
    moduleLogs.OutputErrs
End Function

Function FindAlphabetRegExp(ByRef s As String) As String
    Dim reg As New RegExp
    reg.Pattern = "[^a-zA-Z01233456789.]"
    reg.Global = True
    FindAlphabetRegExp = reg.Replace(s, "-")
End Function

Sub test_FindAlphabetRegExp()
    Debug.Print FindAlphabetRegExp("eeeeee000Ç†Ç†Ç†Ç†rrrr\ttt\tttt\.tdc")
End Sub

Private Function IsCustomCurrentDocument(docType As String, flg As String) As Boolean
    If flg = "" Then
        IsCustomCurrentDocument = True
        Exit Function
    End If
    Dim cFlg As String, cType As String, cFlgSplit As String
    cFlgSplit = Split(flg, " ")(0)
    cFlg = UCase(cFlgSplit)
    cType = UCase(docType)

    If cFlg = "TEXTILE" Or cFlg = "ADOC" Or cFlg = "MD" Then
        IsCustomCurrentDocument = (cType = cFlg)
    Else
        IsCustomCurrentDocument = True
    End If
End Function

''' make translate line [iRow] of excel sheet
'''
''' [tType] shows document type; asciidoc, markdown or textile.
''' [isListContinue] show this line [iRow] is inside list.
''' SheetProperty is set in [sheetInfo].
''' [columnOffset] is the column offset for which language(column) use.
''' [langIndex] is  the column of each language (jp, en ,ch  ... here is language but anything ok etc. release or debug)
''' [columnOffset] is [langIndex] - COL_VALUE
Private Function makeBody(ByRef iRow As Long, ByVal tType As String, ByRef isListContinue As Boolean, ByRef sheetInfo As clsAdocSheetInfo, _
                    ByVal columnOffset As Long, ByVal langIndex As Long, ByRef nestIndex As Long) As String
    Select Case tType
        Case "textile"
        makeBody = makeBodyTextile(iRow, isListContinue, sheetInfo, columnOffset, langIndex, nestIndex)
        
        Case "md"
        makeBody = makeBodyMD(iRow, isListContinue, sheetInfo, columnOffset, langIndex, nestIndex)
        
        Case Else
        MsgBox "doctype error"
        Stop
    End Select
End Function


Private Function GetTableColumns(ByVal r As Long, ByVal c As Long, ByVal maxCols As Long, columns() As Long) As Long
    Dim i As Long
    ReDim columns(0 To maxCols)
    Dim rng As Range
    Set rng = Cells(r, c)
    For i = 0 To maxCols
        If rng.Text = "" Then
            GetTableColumns = i
            Exit Function
        Else
            columns(i) = rng.MergeArea.Column
        End If
        Set rng = rng.Offset(0, 1)
    Next i
End Function

Private Function MCell(ByVal r As Long, ByVal colOffset As Long, columns() As Long) As String
    MCell = Cells(r, columns(colOffset))
End Function


Private Function TrimLeft(cellValue As String) As String

    If cellValue = "" Then
       TrimLeft = ""
       Exit Function
    End If

    TrimLeft = cellValue
    Dim s() As String
    s = Split(cellValue, " ")
    Dim i As Long
    
    Dim left1 As String
    left1 = s(0)
    left1 = Replace(left1, ".", "")
    left1 = Replace(left1, "ÅE", "")
    left1 = Replace(left1, ")", "")
    left1 = Replace(left1, "á@", "1")
    left1 = Replace(left1, "áA", "1")
    left1 = Replace(left1, "áB", "1")
    left1 = Replace(left1, "áC", "1")
    left1 = Replace(left1, "áD", "1")
    left1 = Replace(left1, "áE", "1")
    left1 = Replace(left1, "áF", "1")
    left1 = Replace(left1, "áG", "1")
    left1 = Replace(left1, "áH", "1")
    
    If IsNumeric(left1) Then
        s(0) = ""
        TrimLeft = trim(Join(s, " "))
    End If
End Function

Private Function makeBodyMD(iRow As Long, ByRef isListContinue As Boolean, ByRef sheetInfo As clsAdocSheetInfo, _
                                 ByVal columnOffset As Long, ByVal langIndex As Long, ByRef nestIndex As Long) As String
    makeBodyMD = ""
    '' isListContinue makeBodyMD
    Dim flg As String
    Dim pic As String
    Dim cellValue As String
    
    Dim flgB As String
    Dim picB As String
    Dim cellValueB As String
    
    Dim flgN As String
    Dim picN As String
    Dim cellValueN As String
    Dim pictureColumn As Long
    Dim valueColumn As Long

    pictureColumn = PIC_COL + columnOffset
    valueColumn = VALUE_COL + columnOffset
    
    flg = trim((Cells(iRow, FLG_COL)))
    Dim flgSplit As String
    flgSplit = Split(flg + " ", " ")(0)
    pic = trim((Cells(iRow, pictureColumn)))
    cellValue = trim(Cells(iRow, valueColumn))
    cellValue = RegLinkToMarkDown(cellValue)
    cellValue = Replace(cellValue, "<", "\<")

    If UCase(flgSplit) = "MD" Then
       flg = Mid(flg, 4)
    End If
    
    flgN = trim((Cells(iRow + 1, FLG_COL)))
    picN = trim((Cells(iRow + 1, pictureColumn)))
    cellValueN = trim(Cells(iRow + 1, valueColumn))
    cellValueN = RegLinkToMarkDown(cellValueN)
    
    flgB = trim((Cells(iRow - 1, FLG_COL)))
    picB = trim((Cells(iRow - 1, pictureColumn)))
    cellValueB = trim(Cells(iRow - 1, valueColumn))
    cellValueB = RegLinkToMarkDown(cellValueB)
    
     
    Dim originalRow As Long
    originalRow = iRow
    
    Dim sep As String
    Dim tempBody As String
    
    sep = ""
    Dim sectionNum As Long
    Dim sectionNumString As String

    Dim isSectionNo As Boolean
    isSectionNo = UCase(sheetInfo.Item("IsSectionNo")) = "TRUE"
    
    
    Cells(iRow, VALUE_COL + columnOffset).Font.Color = rgb(0, 0, 0)
    Cells(iRow, VALUE_COL + columnOffset).Font.Size = NORMAL_FONT_SIZE
    
    If Left(flg, 2) = SECTION_TAG2 Then
        nestIndex = 0

        Dim rawFlg As String
        rawFlg = flg
        flg = GetSectionFlg(rawFlg)

        If cellValue = "" Then
            Exit Function
        End If
    
        Dim mdOffsetStr As String
        mdOffsetStr = sheetInfo.Item("mdOffset")
        Dim mdOffset As Long
        mdOffset = 0
        If mdOffsetStr <> "" Then
            mdOffset = CLng(mdOffsetStr)
        End If
        
        
        sectionNum = Len(flg) - 1
        sectionNumString = NewSectionNumber(sectionNum, langIndex)
        rawFlg = SetSectionFlg(flg, sectionNumString)
        If isSectionNo = False Then
            sectionNumString = ""
        End If

        If sectionNum > 4 Then
            sectionNumString = ""
        End If
        
        cellValue = TrimLeft(cellValue)
        

        tempBody = String(Len(flg) - 1 + mdOffset, "#") & " " & sectionNumString & cellValue & vbCrLf

        isListContinue = False
        nestIndex = 0
        Cells(iRow, VALUE_COL + columnOffset).Font.Color = rgb(255, 0, 0) 'red
        Cells(iRow, FLG_COL) = rawFlg
    ElseIf Left(flg, 1) = "*" Then
    
        If cellValue = "" Then
            Exit Function
        End If
      
        nestIndex = Len(flg)
        tempBody = MarkDownIndent(nestIndex - 1) & "-" & " " & cellValue
        
        Cells(iRow, VALUE_COL + columnOffset).Font.Color = rgb(0, 0, 255) 'blue

    ElseIf Left(flg, 1) = "." Then
    
        If cellValue = "" Then
            Exit Function
        End If
    
        Dim isList As Boolean
        isList = flgN = "" Or Left(flgN, 1) = "."
        
        If isList Then
            nestIndex = Len(flg)
            cellValue = TrimLeft(cellValue)
            tempBody = MarkDownIndent(nestIndex - 1) & "1. " & cellValue
            isListContinue = True
        Else
            '' This is Block Attribute, Do nothing
        End If
        Cells(iRow, VALUE_COL + columnOffset).Font.Color = rgb(0, 0, 255) 'blue
        
    ElseIf UCase(flg) = "MD" Then
        tempBody = cellValue

    ElseIf LCase(flg) = "code" Then
        tempBody = "```" + vbCrLf + RTrim(Cells(iRow, valueColumn))
        Dim j As Long
        For j = originalRow + 1 To originalRow + 1000
            If Cells(j, FLG_COL) <> "code" Then
                tempBody = tempBody + vbCrLf + RTrim(Cells(j, valueColumn))
            Else
                tempBody = tempBody + vbCrLf + RTrim(Cells(j, valueColumn)) + vbCrLf + "```"
                iRow = j
                Exit For
            End If
        Next j
        
        If iRow > originalRow + 999 Then
            MsgBox "code error"
            Stop
        End If
        
    ElseIf flg = "[[]]" Then
        tempBody = ""
    
    ElseIf flg = "[]" Then
        tempBody = ""
        
    ElseIf UCase(flg) = "NOTE:" Or UCase(flg) = "TIP:" Or UCase(flg) = "IMPORTANT:" Or UCase(flg) = "CAUTION:" Or UCase(flg) = "WARNING:" Then
       tempBody = MarkDownIndent(nestIndex) & "> " & UCase(flg) & " " & cellValue
        
    ElseIf pic <> "" Then
        Dim strPath As String
        strPath = sheetInfo.Item(IMAGE_PATH)
        
        If strPath <> "" Then
            strPath = strPath & "/"
        End If
        
        tempBody = MarkDownIndent(nestIndex) & "![](./" & strPath & GetFileName(pic) & ")"
        
    ElseIf Left(flg, 4) = "|===" Or LCase(flg) = "table" Then
        Debug.Print ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>", iRow
        Dim tableColTo As Long
        Dim tableRowTo As Long


        Dim trow As Long
        For trow = 1 To MAX_ROW
            If trim((Cells(iRow + trow, FLG_COL))) = "|===" Or LCase(trim((Cells(iRow + trow, FLG_COL)))) = "table" Then
                tableRowTo = iRow + trow
                Exit For
            End If
        Next trow
        
        Dim tableColumns() As Long
        tableColTo = GetTableColumns(iRow, valueColumn, 25, tableColumns)
        
        
        Dim rngTable As Range
        Set rngTable = Range(Cells(iRow, valueColumn), Cells(tableRowTo, tableColTo))
        
        Dim ii As Long
        Dim jj As Long
        
        Dim vv() As Variant
        ReDim vv(0 To tableRowTo - iRow + 1, 0 To tableColTo - 1)
        
        For ii = 0 To tableRowTo - iRow + 1
            For jj = 0 To tableColTo - 1
                If (Left(trim(Cells(iRow + ii, tableColumns(jj))), 2) <> "<!") Then
                    vv(ii, jj) = Replace(Cells(iRow + ii, tableColumns(jj)), "<", "\<")
                Else
                    vv(ii, jj) = Cells(iRow + ii, tableColumns(jj))
                End If
            Next jj
        Next ii


        '' get table columns info
        Dim columnInfo As String
        columnInfo = Cells(iRow - 1, tableColumns(0))
        If LCase(Left(columnInfo, 6)) = "[cols=" Then
            columnInfo = Mid(columnInfo, 8, Len(columnInfo) - 9)
            '' <!-- word cols 1,4  -->
            tempBody = "<!-- word cols " & columnInfo & " -->"
        Else
            '' do noting
        End If
        
        ' // merge info rows
        ' // <!-- wordDown rowMerge 3-4,5-6,7-9 -->
        ' // rowMerge = 3-4,5-6,7-9

        '' table merge info
        Dim rMergeFlags As Range
        Set rMergeFlags = Range(Cells(iRow, PIC_COL), Cells(tableRowTo, PIC_COL))
        Dim rowMergeStartEnd() As Long ''(1) = 2
        ReDim rowMergeStartEnd(0 To rMergeFlags.rows.Count)
        Dim inSideMergeArea As Boolean
        Dim endline As Long
        For ii = rMergeFlags.rows.Count To 0 Step -1
            If rMergeFlags.Cells(ii, 1) = "m" And inSideMergeArea = False Then
                inSideMergeArea = True
                endline = ii
            End If

            If inSideMergeArea And rMergeFlags.Cells(ii, 1) = "" Then
                rowMergeStartEnd(ii) = endline
                inSideMergeArea = False
                endline = 0
            End If
        Next

        Dim mergeInfo As clsString
        Set mergeInfo = New clsString
        For ii = 1 To rMergeFlags.rows.Count
            If rowMergeStartEnd(ii) > 0 Then
                mergeInfo.Add ii & "-" & CStr(rowMergeStartEnd(ii))
            End If
        Next
        Dim wordDownMerge

        '' add table information
        wordDownMerge = "<!-- word rowMerge " & mergeInfo.Joins(",") & " -->"

        If tempBody = "" Then
            tempBody = wordDownMerge
        Else
            tempBody = tempBody & vbCrLf & wordDownMerge
        End If
        
        Dim tableVales As String
        tableVales = makeGridMarkDown2(vv, "")
        If tempBody = "" Then
            tempBody = tableVales
        Else
            tempBody = tempBody & vbCrLf & tableVales
        End If

        iRow = tableRowTo
    
    ElseIf flg = "_" Then
        globalBeforeParam = cellValue
        
    ElseIf Left(flg, 5) = "_link" Then
        MsgBox "flg, _link err"
        Stop
        
    ElseIf flg = "//" Then
        tempBody = "<!--" & cellValue & "-->"
        
    ElseIf flg = "[#" Or Left(flg, 1) = "[" Then
        tempBody = ""

    ElseIf UCase(flg) = "LISTOFF" Then
        nestIndex = nestIndex - 1
        If nestIndex < 0 Then
            nestIndex = 0
        End If
    Else
        tempBody = MarkDownIndent(nestIndex) & IIf(flg = "", "", flg & " ") & sep & cellValue
    End If
    
    makeBodyMD = tempBody
End Function

Private Function MarkDownIndent(nestIndex As Long) As String
    MarkDownIndent = String(nestIndex * 4, " ")
End Function


Private Function GetSectionFlg(ByVal flg As String) As String
    Dim sectionFlgAndStr() As String
    sectionFlgAndStr = Split(flg, ",")
    GetSectionFlg = sectionFlgAndStr(0)
End Function

Private Function SetSectionFlg(ByVal flg As String, ByVal sectionString As String) As String
    SetSectionFlg = flg & "," & sectionString
End Function

Private Function makeBodyTextile(ByRef iRow As Long, ByRef isListContinue As Boolean, ByRef sheetInfo As clsAdocSheetInfo, _
        ByVal columnOffset As Long, ByVal langIndex As Long, ByRef nestIndex As Long) As String
    Dim flg As String
    Dim pic As String
    Dim cellValue As String
    
    Dim flgB As String
    Dim picB As String
    Dim cellValueB As String
    
    Dim flgN As String
    Dim picN As String
    Dim cellValueN As String

    Dim pictureColumn As Long
    Dim valueColumn As Long

    pictureColumn = PIC_COL + columnOffset
    valueColumn = VALUE_COL + columnOffset
    
    flg = trim((Cells(iRow, FLG_COL)))
    pic = trim((Cells(iRow, pictureColumn)))
    cellValue = (trim(Cells(iRow, valueColumn)))
    
    flgN = trim((Cells(iRow + 1, FLG_COL)))
    picN = trim((Cells(iRow + 1, pictureColumn)))
    cellValueN = (trim(Cells(iRow + 1, valueColumn)))
    
    flgB = trim((Cells(iRow - 1, FLG_COL)))
    picB = trim((Cells(iRow - 1, pictureColumn)))
    cellValueB = (trim(Cells(iRow - 1, valueColumn)))
    
    
    Dim sep As String
    Dim tempBody As String
    
    Dim originalRow As Long
    originalRow = iRow
        
    sep = ""
    Dim sectionNum As Long
    Dim sectionNumString As String
    Dim isSectionNo As Boolean
    isSectionNo = UCase(sheetInfo.Item("IsSectionNo")) = "TRUE"
    
    If Left(flg, 2) = SECTION_TAG2 Then
        Dim rawFlg As String
         rawFlg = flg
        flg = GetSectionFlg(rawFlg)
    
        ' If flg = SECTION_TAG2 Then
        '     gTextileFistSectionName = cellValue
        ' End If
    
        If sheetInfo.Item("RETURNTOC") <> "" Then
            tempBody = vbCrLf & sheetInfo.Item("RETURNTOC") & vbCrLf & vbCrLf
        End If

        sectionNumString = ""
        sectionNum = Len(flg) - 1
        sectionNumString = NewSectionNumber(sectionNum, langIndex)
        rawFlg = SetSectionFlg(flg, sectionNumString)
        If isSectionNo = False Then
            sectionNumString = String(Len(flg) - 1, "Å°") & " "
        End If
        
        cellValue = TrimLeft(cellValue)

        tempBody = tempBody & "h" & CStr(sectionNum) & ". " & sectionNumString & cellValue & vbCrLf
        Cells(iRow, VALUE_COL + columnOffset).Font.Color = rgb(255, 0, 0) 'red
        Cells(iRow, FLG_COL) = rawFlg
    ElseIf LCase(flg) = "code" Then
        tempBody = "<pre>" + vbCrLf + RTrim(Cells(iRow, VALUE_COL + columnOffset))
        Dim j As Long
        For j = originalRow + 1 To originalRow + 1000
            If Cells(j, FLG_COL) <> "code" Then
                tempBody = tempBody + vbCrLf + RTrim(Cells(j, VALUE_COL + columnOffset))
            Else
                tempBody = tempBody + vbCrLf + RTrim(Cells(j, VALUE_COL + columnOffset)) + vbCrLf + "</pre>"
                iRow = j
                Exit For
            End If
        Next j
        
        If iRow > originalRow + 999 Then
            MsgBox "code error"
            Stop
        End If
        
    ElseIf flg = "[[]]" Then
        tempBody = ""
    
    ElseIf flg = "[]" Then
        tempBody = ""
        
    ElseIf flg = "." Then
        Dim isList As Boolean
        isList = flgN = "" Or flgN = "."
        
        If isList Then
            cellValue = TrimLeft(cellValue)
            tempBody = NewListNum(langIndex) & cellValue
        End If

    ElseIf UCase(flg) = "TEXTILE" Then
        tempBody = cellValue
        
    ElseIf pic <> "" Then
        tempBody = "!{border: solid}" & GetFileName(pic) & "(" & GetFileName(pic) & ")!"

    ElseIf Left(flg, 4) = "|===" Or LCase(flg) = "table" Then
        Dim tableColTo As Long
        Dim tableRowTo As Long
        'tableColTo = Cells(iRow, VALUE_COL).End(xlToRight).Column
        'tableRowTo = Cells(iRow, VALUE_COL).End(xlDown).row
        Dim tableColumns() As Long
        tableColTo = GetTableColumns(iRow, VALUE_COL + columnOffset, 25, tableColumns)
        
        
        Dim trow As Long
        For trow = 1 To MAX_ROW
            If trim((Cells(iRow + trow, FLG_COL))) = "|===" Or LCase(trim((Cells(iRow + trow, FLG_COL)))) = "table" Then
                tableRowTo = iRow + trow
                Exit For
            End If
        Next trow
        
        If trow > MAX_ROW Then
            MsgBox "Table Rows is too long " & CStr(iRow)
            Stop
        End If
        Dim ii As Long
        Dim jj As Long
        
        
        ReDim vv(0 To tableRowTo - iRow, 0 To tableColTo - 1)
        
        For ii = 0 To tableRowTo - iRow
            For jj = 0 To tableColTo - 1
                vv(ii, jj) = Cells(iRow + ii, tableColumns(jj))
            Next jj
        Next ii
        

        Dim rngTable As Range
        'Set rngTable = Range(Cells(iRow, VALUE_COL + columnOffset), Cells(tableRowTo, tableColTo + columnOffset))

        tempBody = makeGridRedmine2(vv)
        iRow = tableRowTo
    ElseIf Left(flg, 1) = "." Then
        ' Dim isList As Boolean
        ' isList = flgN = "" Or flgN = "."
    
        ' If isList Then
        '     tempBody = "# " & cellValue
        ' End If
    ElseIf UCase(flg) = "LISTOFF" Then
        '' for asciidoctor
    Else
        tempBody = IIf(flg = "", "", flg & " ") & sep & cellValue
    End If
    
    makeBodyTextile = tempBody
    
    
End Function





Public Sub SetImageAddxlMove()
    Dim topRow As Long
    Dim sp As Shape
    Dim sh As Worksheet
    
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    For Each sh In ActiveWindow.SelectedSheets
        sheetCollection.Add sh
    Next
       
    For Each sh In sheetCollection
        sh.Select
        sh.Activate
        For Each sp In ActiveSheet.shapes
            If sp.Type <> msoComment Then
                sp.Placement = xlMove
                If sp.title = "" Then
                    sp.title = CStr(sp.Height) & "," & CStr(sp.Width)
                    ''Debug.Print CStr(sp.Height) & "," & CStr(sp.Width)
                Else
                    ''Debug.Print "shape titel is not empty."
                End If
            End If
        Next
    Next
End Sub

Public Sub RefreshImageSize()
    Dim pics As ShapeRange
    Set pics = Selection.ShapeRange
    Dim sp As Shape
    Set sp = pics(1)
    sp.title = CStr(sp.Height) & "," & CStr(sp.Width)
End Sub



Public Sub AddSpaceUnderPicture()

    Dim sp As Shape
    
   
    Dim topRow As Long
    Dim bottomRow As Long
    Dim dataExistRow As Long
    
    Dim i As Long
    
    Dim DataEndRow As Long
    DataEndRow = getEndLine(1, 20)
    
    Dim ismoved As Boolean
    
    Dim delrange As Range
    
    Dim sheetInfo As clsAdocSheetInfo
    Set sheetInfo = New clsAdocSheetInfo
    Call sheetInfo.SetSheetInfo(ActiveSheet, FLG_COL, VALUE_COL, VALUE_COL)
    
    For Each sp In ActiveSheet.shapes
        ismoved = False
        If sp.Type <> msoComment And sheetInfo.ExistsLangColumn(sp.TopLeftCell.Column) Then
            sp.Placement = xlMove
            ''Debug.Print sp.TopLeftCell.Address   '
            topRow = sp.TopLeftCell.row
            bottomRow = sp.BottomRightCell.row
            dataExistRow = 0
            For i = topRow To IIf(DataEndRow > bottomRow, DataEndRow, bottomRow)
                If WorksheetFunction.CountA(Cells(i, VALUE_COL).EntireRow) > 0 Then
                    dataExistRow = i
                    Exit For
                End If
            Next i
            
        
            If topRow = dataExistRow Then
                sp.Top = sp.TopLeftCell.Offset(-1, 0).Top
                ismoved = True
            End If
            
            If dataExistRow = 0 Then
                '' do nothing
                
            ElseIf dataExistRow <= bottomRow Then
                Cells(dataExistRow, 1).Resize(bottomRow - dataExistRow + 1, 1).EntireRow.Insert xlShiftDown
                
            ElseIf dataExistRow > bottomRow + 3 Then

                Set delrange = sp.BottomRightCell.Offset(1, 0).Resize(dataExistRow - bottomRow - 2, 1).EntireRow
                If WorksheetFunction.CountA(delrange) = 0 Then
                    delrange.Delete
                End If
            Else
                '
            End If
            
            
            If ismoved Then
                sp.Top = sp.TopLeftCell.Offset(1, 0).Top + 5
            End If
             
        End If
    Next
End Sub


'Sub TEST_getEndLine()
'    Debug.Print getEndLine(1, 3)
'End Sub


Public Function getEndLine(ByVal ncol1 As Long, Optional ByVal ncol2 As Long) As Long
    
    Dim rngTarget As Range
    Set rngTarget = Application.Intersect(Range(Cells(1, ncol1), Cells(rows.Count, ncol2)), ActiveSheet.UsedRange)
    getEndLine = rngTarget.row + rngTarget.rows.Count

End Function


Public Function GetLinkIndex() As String()
    ''Call InitParam
    
    Dim i As Long
    Dim asciidoc As clsString
    Set asciidoc = New clsString
    
    
    Dim flg As String
    Dim cellValue As String
    Dim isBeforeNull As Boolean
    isBeforeNull = False
    
    Dim sh As Worksheet
    
    
    Dim dicIndex As Object
    Set dicIndex = CreateObject("Scripting.Dictionary")
    Dim strIndexes() As String
    'dicR.CompareMode = vbTextCompare
    ' keys itmes
    
    
    Worksheets("Setting").Select
    Dim endline As Long
    endline = modAsciidoc.getEndLine(1, 1)
    
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    
    Dim ii As Long
    For ii = 1 To endline
        If Cells(ii, 2) = "1" Then
            sheetCollection.Add Worksheets(CStr(Cells(ii, 1)))
        End If
    Next ii
    
    For Each sh In sheetCollection 'ActiveWindow.SelectedSheets
        sh.Select
        sh.Activate
        
        Dim MaxRows As Long
        MaxRows = getEndLine(FLG_COL, VALUE_COL)
        
        Dim DATA_START_ROW As Long
        DATA_START_ROW = GetContentStartRow(sh)
        
    
        For i = DATA_START_ROW To MaxRows
            flg = Cells(i, FLG_COL)
            cellValue = Cells(i, VALUE_COL)
            
            If Left(flg, 2) = "[[" Or Left(flg, 2) = "[#" Then
                strIndexes = Split(Mid(Replace(flg, "]", ""), 3), ",")
                
                On Error Resume Next
                If UBound(strIndexes) > 0 Then
                    ''dicIndex.add Trim(strIndexes(0)), "<<" & Mid(Replace(flg, "]", ""), 3) & ">>"
                    dicIndex.Add trim(strIndexes(0)), strIndexes(0) & "," & strIndexes(1)
                Else
                    ''dicIndex.add Trim(strIndexes(0)), "<<" & strIndexes(0) & "," & Cells(i + 1, VALUE_COL) & ">>"
                    dicIndex.Add trim(strIndexes(0)), strIndexes(0) & "," & Cells(i + 1, VALUE_COL)
                End If
                On Error GoTo 0
            End If
        Next i
    Next
    
    Dim ctmp As clsString
    Set ctmp = New clsString
    Dim ctmpstr() As Variant
    ctmpstr = dicIndex.Items
    ctmp.AddArrayV ctmpstr
    
    GetLinkIndex = ctmp.MakeArray

'    With CreateObject("forms.TextBox.1")
'     .MultiLine = True
'     .Text = ctmp.Joins(vbCrLf)
'     .SelStart = 0
'     .SelLength = .TextLength
'     .Copy
'    End With
End Function


Public Sub LoadPictures(Optional ByVal loadrow As Long = 0)
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""

    On Error GoTo EH

    Dim sheetInfo As clsAdocSheetInfo
    Set sheetInfo = New clsAdocSheetInfo
    Call sheetInfo.SetSheetInfo(ActiveSheet, FLG_COL, VALUE_COL, VALUE_COL)

    Dim langPath As String
    langPath = sheetInfo.Item(lang)
    Dim imagePath As String
    imagePath = sheetInfo.Item(IMAGE_PATH)

    Dim i As Long
    
    Dim ppath As String

    ppath = ActiveWorkbook.Path & "\" & langPath & "\" & imagePath & "\"
    ppath = Replace(ppath, "\\", "\")
    ppath = Replace(ppath, "\\", "\")
    
    
    Dim MaxRows As Long
    MaxRows = getEndLine(PIC_COL, PIC_COL)
    
    Dim minrow As Long
    
    If loadrow > 0 Then
        MaxRows = loadrow
        minrow = loadrow
    Else
        minrow = 1
    End If
    
    Dim imageFileOrPath As String
    For i = MaxRows To minrow Step -1
        imageFileOrPath = Cells(i, PIC_COL)
        If UCase(getFileExtension(imageFileOrPath)) = "PNG" Or UCase(getFileExtension(imageFileOrPath)) = "GIF" Then
            LoadPicFile ppath & (imageFileOrPath), Cells(i, PIC_COL).Offset(PIC_OFFSET_Y, PIC_OFFSET_X)  ''
        End If
    Next i
    
EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = ""
End Sub


Public Sub SetSetPicNameCellFromCellAddress(Optional prefix As String)
    Dim sp As Shape
    
    Dim logs As clsLogs
    Set logs = New clsLogs
    
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    Dim sht As Worksheet
    
    For Each sht In ActiveWindow.SelectedSheets
        sheetCollection.Add sht
    Next
    
    Worksheets(1).Select
    
    
    Dim picNameRng As Range
    Dim imageName As String
    
    For Each sht In sheetCollection
        sht.Select
        sht.Activate
        
        If prefix = "" Then
            prefix = ActiveSheet.Name
        End If
        
        For Each sp In ActiveSheet.shapes
            If sp.Type <> msoComment Then
                Set picNameRng = GetPicNameCellForWrite(sp)
                If picNameRng = "" Then
                    imageName = prefix & "_" & Replace(picNameRng.Address, "$", "") & ".png"
                    ''picNameRng = "![PIC](" & imageName & ")"
                    picNameRng = imageName
                Else
                    ''
                End If
            End If
        Next
    Next
End Sub



Private Function GetPicNameCell(ByVal r As Shape) As Range
    Set GetPicNameCell = r.TopLeftCell.Offset(0, 0)
    
    If GetPicNameCell.Text <> "" Then
        
    Else
        Set GetPicNameCell = r.TopLeftCell.Offset(0, -1)
    End If
End Function


Private Function GetPicNameCellForWrite(ByVal r As Shape) As Range
    Dim pic1 As String
    Dim pic2 As String
    Dim pic3 As String
    
    pic1 = r.TopLeftCell.Offset(-1, -1).Text
    pic2 = r.TopLeftCell.Offset(0, -1).Text
    pic3 = r.TopLeftCell.Offset(0, 0).Text
    
    If pic1 = "" And pic2 = "" And pic3 = "" Then
        Set GetPicNameCellForWrite = r.TopLeftCell.Offset(0, -1)
    ElseIf pic1 <> "" Then
        Set GetPicNameCellForWrite = r.TopLeftCell.Offset(-1, -1)
    ElseIf pic2 <> "" Then
        Set GetPicNameCellForWrite = r.TopLeftCell.Offset(0, -1)
    ElseIf pic3 <> "" Then
        Set GetPicNameCellForWrite = r.TopLeftCell.Offset(0, 0)
    End If
End Function


Public Sub CheckPictureFileNames()

On Error GoTo CheckPictureFileNames_Error

    Dim sp As Shape

    Dim myLogs As clsLogs
    Set myLogs = New clsLogs
    
    Dim i As Long
    
    If MsgBox("Do you check the file name exists ", vbOKCancel) = vbOK Then
        'ActiveSheet.Shapes.SelectAll
    Else
        GoTo normalex
    End If
    
    
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    Dim sht As Worksheet
    
    For Each sht In ActiveWindow.SelectedSheets
        sheetCollection.Add sht
    Next
    
    Worksheets(1).Select
    
    
    Dim fname As String
    For Each sht In sheetCollection
    
        sht.Activate
        ActiveSheet.shapes.SelectAll
        If ActiveSheet.shapes.Count > 0 Then
            For Each sp In Selection.ShapeRange
                If sp.Type <> msoComment Then
                    fname = GetPicNameCell(sp)
                    If fname = "" Then
                        myLogs.AddErr ActiveSheet.Name, GetPicNameCell(sp).Address, "No image file name."
                    End If
                End If
            Next
        End If
        ActiveCell.Select
    Next
    
    
    myLogs.OutputErrs "", True

normalex:
    
    On Error GoTo 0
    Exit Sub

CheckPictureFileNames_Error:
    ActiveCell.Select
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckPictureFileNames, line " & Erl & "."


End Sub

Public Sub SavePictures()
    SavePicturesWork False
End Sub


Public Sub GetContentStartRow_Test()
    Dim l As Long
    l = GetContentStartRow(Worksheets("sheet4"))
    Debug.Print l
End Sub

Private Function GetContentStartRow(sht As Worksheet) As Long
    GetContentStartRow = sht.Range("A1").End(xlDown).row + 2  ' 2: next row and empty row
End Function


Public Sub SavePicturesWork(ByVal isOrgsize As Boolean)
    
    On Error GoTo SavePictures_Error
    
    Dim myLogs As clsLogs
    Set myLogs = New clsLogs
    
    Dim sheetInfo As clsAdocSheetInfo
    Set sheetInfo = New clsAdocSheetInfo
    sheetInfo.SetSheetInfo ActiveSheet, FLG_COL, VALUE_COL, VALUE_COL

    Dim outDir As String
    outDir = sheetInfo.GetOutPicDirFullPath()

    sheetInfo.makeLanguageFolders outDir
    
   
    Dim iCount As Long
    On Error Resume Next
    
    Dim t As String
    t = TypeName(Selection)
    
    If t = "Range" Then
        'not picture drawing
        iCount = 0
    Else
        iCount = Selection.ShapeRange.Count
    End If
    
'    iCount = Selection.ShapeRange.Count
'    If Err.Number <> 0 Then
'        iCount = 0
'    End If
    
    On Error GoTo SavePictures_Error
    If iCount = 0 Then
        If MsgBox("Save all images?", vbOKCancel) = vbOK Then
            If ActiveSheet.shapes.Count > 0 Then
                ActiveSheet.shapes.SelectAll
            Else
                GoTo normalex
            End If
        Else
            GoTo normalex
        End If
    End If
    
    Dim pics As ShapeRange
    Set pics = Selection.ShapeRange
    
    pics.Item(1).TopLeftCell.Select
    
    
    Dim fname As String
    Dim fnameCell As String
    Dim columnforpic As String
    Dim curpath As String
    Dim savePath As String
    

    curpath = sheetInfo.GetOutPicDirFullPath()
    
    Dim sp As Shape
    Dim sptmp As Shape
    
    Dim targetCell As String
    
    For Each sp In pics
    
        If sp.Type <> msoComment Then
            columnforpic = CStr(sp.TopLeftCell.Column)
            savePath = sheetInfo.makePicFilePath(curpath, columnforpic)
            fnameCell = GetPicNameCell(sp)
            '' get file name from ![PIC](Sheet4_C16.png)
            ''

            fname = RegFindCapture("!\[.*?\]\((.*?)\)", fnameCell)

            If fname = "" Then
               fname = fnameCell
            End If

            If sheetInfo.GetDicPic.Exists(getFileExtension(fname)) Then
                ''
            Else
                If fname <> "" Then
                    fname = fname & ".png"
                    GetPicNameCell(sp) = fname
                End If
            End If
            
            
            '' capture after copy
            targetCell = sp.TopLeftCell.Address
            Set sptmp = sp
            sptmp.Select
            

            Dim sizeText As String
            sizeText = sp.title
            Dim sizeArray()  As String
            sizeArray = Split(sizeText, ",")
            Dim sizeArraySingle(0 To 1) As String
            
            If UBound(sizeArray) > 0 Then
                If IsNumeric(sizeArray(0)) Then
                    sizeArraySingle(0) = CSng(sizeArray(0))
                    
                    If IsNumeric(sizeArray(1)) Then
                        sizeArraySingle(1) = CSng(sizeArray(1))
                        
                        sptmp.LockAspectRatio = msoFalse
                        sp.Height = sizeArraySingle(0)
                        sp.Width = sizeArraySingle(1)
                        sp.LockAspectRatio = msoTrue
                    End If
                End If
            End If
            
            sptmp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
            ''sp.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
            ''sptmp.Delete
            
        
            '' main picture save
            Dim retCode As Long
            If fname = "" Then
                myLogs.AddErr ActiveSheet.Name, GetPicNameCell(sp).Address, "No file name for Image."
            Else
                CreateDirectory savePath
                retCode = SaveCipPictureByName(savePath & "\" & fname, 100)  '100 is quality of jpeg
                If retCode <> 0 Then
                    myLogs.AddErr ActiveSheet.Name, GetPicNameCell(sp).Address, "Image Output Error retCode, " & retCode
                Else
                    '
                End If
            End If
        End If
    Next
    
    myLogs.OutputErrs "", True

normalex:
    
    On Error GoTo 0
    Exit Sub

SavePictures_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SavePictures, line " & Erl & " " & targetCell & "."

End Sub

Private Function makeGridRedmine2(ByRef tables()) As String
    ' empty cells are merged
    Dim rngSrcGrid As Range
    Dim myrows As Long
    Dim mycols As Long

    myrows = UBound(tables, 1) + 1
    mycols = UBound(tables, 2) + 1

    If myrows * mycols = 1 Then Exit Function

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim tmpLenb As Long
    Dim tmpLinesInRow As Long

    Dim LinesInRow() As Long
    ReDim LinesInRow(0 To myrows)

    For i = 1 To mycols
        For j = 1 To myrows
            tmpLinesInRow = UBound(Split(tables(j - 1, i - 1), vbLf)) + 1
            If LinesInRow(j) < tmpLinesInRow Then
                LinesInRow(j) = tmpLinesInRow
            End If
        Next j
    Next i

    'add string
    Dim dim3data() As String
    ReDim dim3data(0 To myrows, 0 To mycols, 0 To 10)  '(row, column, valueÅj 1o lines on one cell
    Dim tmparray() As String
    
    'rows in cell
    Dim maxLenInCol() As Long
    ReDim maxLenInCol(0 To mycols)

    For i = 1 To mycols
        For j = 1 To myrows
            tmparray = Split(tables(j - 1, i - 1), vbLf)

            For k = 1 To UBound(tmparray) + 1
                dim3data(j, i, k) = trim(tmparray(k - 1))

                tmpLenb = modCommon.LenMbcs(dim3data(j, i, k))
                If maxLenInCol(i) < tmpLenb Then
                    maxLenInCol(i) = tmpLenb
                End If
            Next k
        Next j
    Next i


    'make grid
    Dim GridOutput As Collection
    Set GridOutput = New Collection

    Dim strRowSep As String
    Dim strIndexSep As String
    Dim strRowSepInRow As String
    Dim strRowContent As String
    Dim strCellContent As String

    Dim countMargin As Long
    countMargin = 1


    strRowSep = "+"
    strRowSepInRow = "|"
    strIndexSep = "+"
    For i = 1 To mycols
        'strRowSep = strRowSep & String(maxLenInCol(i) + 2 * countMargin, "-") & "+"
        'strRowSepInRow = strRowSepInRow & String(maxLenInCol(i) + 2, " ") & "|"
        'strIndexSep = strIndexSep & String(maxLenInCol(i) + 2 * countMargin, "=") & "+"
    Next i
    
    Dim spaceRowsInCol() As Long
    ReDim spaceRowsInCol(0 To mycols) ' for merge rows

    ''GridOutput.Add strRowSep
    For j = myrows To 1 Step -1
        For k = LinesInRow(j) To 1 Step -1 'for multi rows
            strRowContent = "|"
            For i = 1 To mycols
                strCellContent = trim(dim3data(j, i, k))
                tmpLenb = LenMbcs(strCellContent)
                ''Debug.Print tmpLenb, strCellContent


                '' merge rows
                If strCellContent = "^" Then
                    spaceRowsInCol(i) = spaceRowsInCol(i) + 1
                Else
                    Dim strRowCombine As String
                    If spaceRowsInCol(i) > 0 Then
                        strRowCombine = "/" & CStr(spaceRowsInCol(i) + 1) & "."
                    Else
                        strRowCombine = ""
                    End If
                    
                    Dim addSpaces As Long
                    
                    '' wide column is not added extra spaces
                    addSpaces = IIf(maxLenInCol(i) > 26, 0, maxLenInCol(i) - tmpLenb)
                    addSpaces = 2
                    strRowContent = strRowContent & _
                                    strRowCombine & String(countMargin, " ") & _
                                        strCellContent & _
                                    String(addSpaces, " ") & String(countMargin, " ") & "|"
                                    
                    spaceRowsInCol(i) = 0
                End If
            Next i

            If GridOutput.Count = 0 Then
                GridOutput.Add strRowContent
            Else
                GridOutput.Add strRowContent, , 1
            End If
        Next k
        '        If j = 1 Then
        '            GridOutput.Add strIndexSep
        '        Else
        '            GridOutput.Add strRowSep
        '        End If
    Next j

    'output
    For j = 1 To GridOutput.Count
        makeGridRedmine2 = makeGridRedmine2 & GridOutput.Item(j) & vbCrLf
    Next j
End Function

Private Function makeGridMarkDown2(ByRef vv(), strIndent As String) As String
    ' nothing to do for empty line
    ' when merge upper cell, add ^

    Dim myrows As Long
    Dim mycols As Long

    myrows = UBound(vv, 1)
    mycols = UBound(vv, 2) + 1

    If myrows * mycols = 1 Then Exit Function

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim tmpLenb As Long
    Dim tmpLinesInRow As Long


    'read grid
    Dim LinesInRow() As Long
    ReDim LinesInRow(0 To myrows)

    ' check rows of the cells
    Dim maxLenInCol() As Long
    ReDim maxLenInCol(0 To mycols)


    For i = 1 To mycols
        For j = 1 To myrows

            'check the rows of column
            tmpLinesInRow = UBound(Split(vv(j - 1, i - 1), vbLf)) + 1
            If LinesInRow(j) < tmpLinesInRow Then
                LinesInRow(j) = tmpLinesInRow
            Else

            End If
        Next j
    Next i


    ' add string
    Dim dim3data() As String
    ReDim dim3data(0 To myrows, 0 To mycols)
    For i = 1 To mycols
        maxLenInCol(i) = 0
        For j = 1 To myrows
            dim3data(j, i) = Replace(vv(j - 1, i - 1), vbCrLf, "<BR>")
            tmpLenb = LenMbcs(dim3data(j, i))
            If maxLenInCol(i) < tmpLenb Then
                maxLenInCol(i) = tmpLenb
            End If
        Next j
    Next i


    'make grid
    Dim GridOutput As Collection
    Set GridOutput = New Collection

    Dim strRowSep As String
    Dim strIndexSep As String
    Dim strRowSepInRow As String
    Dim strRowContent As String
    Dim strCellContent As String

    Dim countMargin As Long
    countMargin = 1


    strRowSep = "+"
    strRowSepInRow = "|"
    strIndexSep = "+"
    
    For i = 1 To mycols
        strRowSep = strRowSep & String(maxLenInCol(i) + 2 * countMargin, "-") & "+"
        strRowSepInRow = strRowSepInRow & String(maxLenInCol(i) + 2, " ") & "|"
        strIndexSep = strIndexSep & String(maxLenInCol(i) + 2 * countMargin, "=") & "+"
    Next i

    ''GridOutput.Add strRowSep
    Dim colwidth As Long
    For j = 1 To myrows
            strRowContent = "|"
            For i = 1 To mycols
                ''strCellContent = Trim(dim3data(j, i))
                strCellContent = (dim3data(j, i))
                tmpLenb = LenMbcs(strCellContent)
                colwidth = maxLenInCol(i) - tmpLenb
                ''colwidth = 2
                strRowContent = strRowContent & _
                                String(countMargin, " ") & _
                                    strCellContent & _
                                String(colwidth, " ") & String(countMargin, " ") & "|"
            Next i
            
            GridOutput.Add strIndent & strRowContent
            
            If GridOutput.Count = 1 Then
                Dim c As clsString
                Set c = New clsString
                For i = 1 To mycols
                    c.Add String(maxLenInCol(i) + 2, "-")
                Next i
                GridOutput.Add strIndent & c.Joins("|")
            End If
    Next j

    'output
    For j = 1 To GridOutput.Count
        makeGridMarkDown2 = makeGridMarkDown2 & GridOutput.Item(j) & vbCrLf
    Next j
End Function


Public Sub DeleSectionNum__Main()

    'Call InitParam
    
    Dim i As Long
    Dim asciidoc As clsString
    Set asciidoc = New clsString
    
    
    Dim flg As String
    Dim pic As String
    Dim cellValue As String
    Dim isBeforeNull As Boolean
    isBeforeNull = False
    
    Dim sh As Worksheet
    
    
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    
    For Each sh In ActiveWindow.SelectedSheets
        sheetCollection.Add sh
    Next
    
    
    For Each sh In sheetCollection
        sh.Select
        sh.Activate
        
        Dim MaxRows As Long
        MaxRows = getEndLine(FLG_COL, VALUE_COL)
        
        Dim DATA_START_ROW As Long
        DATA_START_ROW = GetContentStartRow(sh)
    
        For i = DATA_START_ROW To MaxRows
            flg = Cells(i, FLG_COL)
            pic = Cells(i, PIC_COL)
            cellValue = Cells(i, VALUE_COL)
            
            If flg = "." Then
                Cells(i, VALUE_COL) = DeleSectionNumWork(cellValue)
            End If
        Next i
    Next
End Sub

Private Function DeleSectionNumWork(ByVal myValue As String) As String
    Dim tmp As String
    tmp = myValue
    Dim spacepos As Long
    spacepos = InStr(1, tmp, " ")
    
    Dim strSplit() As String
    strSplit = Split(tmp)
    If IsNumeric(Replace(strSplit(0), ".", "")) Then
        tmp = Mid(tmp, spacepos + 1)
    Else
    
    End If
    
    DeleSectionNumWork = tmp
End Function

Public Sub ArrangeFormatXls()

    'Call InitParam
    
    Dim i As Long
    
    
    Dim flg As String
    Dim pic As String
    Dim cellValue As String
    Dim isBeforeNull As Boolean
    isBeforeNull = False
    
    Dim sh As Worksheet
    
    Dim ss() As String
    
    
    Dim sheetCollection As Collection
    Set sheetCollection = New Collection
    
    For Each sh In ActiveWindow.SelectedSheets
        sheetCollection.Add sh
    Next
    
    
    For Each sh In sheetCollection
        sh.Select
        sh.Activate

        Dim MaxRows As Long
        MaxRows = getEndLine(FLG_COL, VALUE_COL)
        
        Dim DATA_START_ROW As Long
        DATA_START_ROW = GetContentStartRow(sh)
        
        For i = DATA_START_ROW To MaxRows
            flg = Cells(i, FLG_COL)
            pic = Cells(i, PIC_COL)
            cellValue = Cells(i, VALUE_COL)
            
            If cellValue <> "" And flg = "" Then
                ss = Split(cellValue, " ")
                If UBound(ss) > 0 Then
                    Select Case ss(0)
                        Case ".", "..", SECTION_TAG, SECTION_TAG2, SECTION_TAG3, SECTION_TAG4, SECTION_TAG5, "NOTE:", "TIP:", "*"
                            Cells(i, FLG_COL) = ss(0)
                            Cells(i, VALUE_COL) = Mid(cellValue, Len(ss(0)) + 2)
                    End Select
                Else
                    If Left(cellValue, 2) = "[[" Then
                            Cells(i, FLG_COL) = cellValue
                            Cells(i, VALUE_COL) = ""
                    End If
                End If
            End If
            
            flg = Cells(i, FLG_COL)
            pic = Cells(i, PIC_COL)
            cellValue = Cells(i, VALUE_COL)
            
            
            ''color
            If cellValue <> "" And flg <> "" Then
                Select Case flg
                    Case SECTION_TAG2, SECTION_TAG3, SECTION_TAG4
                        With Cells(i, 1).EntireRow.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                        End With
                    Case Else
                        With Cells(i, 1).EntireRow.Interior
                            .Pattern = xlNone
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                End Select
           Else
            
           End If
        Next i
    Next
End Sub








Public Function GetAsciiTableCols(ByRef colindex As String) As Long
    '' "[width=""100%"",cols=""1""G,]"
    Dim ss() As String
    ss = Split(colindex, "cols=""")
    
    If UBound(ss) > 0 Then
        GetAsciiTableCols = UBound(Split(Split(ss(1), """")(0), ",")) + 1
    Else
        GetAsciiTableCols = 0
    End If
End Function

Public Sub CreateRedBox()
    On Error GoTo EH

    Dim r As Range
    
    ' if  selection is range
    If TypeName(Selection) = "Range" Then
        '' go next
    Else
        MsgBox "select range or a cell. Not Picture or some."
        Exit Sub
    End If
    
    
    Set r = Selection
    ActiveSheet.shapes.AddShape(msoShapeRoundedRectangle, r.Left, r.Top, r.Width, r.Height).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .weight = 3
    End With
    Selection.ShapeRange.Fill.Visible = msoFalse
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.rgb = rgb(255, 0, 0)
        .Transparency = 0
    End With
    Exit Sub
    
EH:
    MsgBox Err.Description
End Sub

Public Sub selectionHeadingUpDown_TestUp()
    selectionHeadingUpDown -1
End Sub

Public Sub selectionHeadingUpDown_TestDown()
    selectionHeadingUpDown 1
End Sub

Public Sub selectionHeadingUpDown(upDown As Long)
    Dim r As Range
    Set r = Selection
    
    If r.Left > 0 Then
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To r.rows.Count
        r.Cells(i, 1) = headingUpDown(r.Cells(i, 1), upDown)
    Next
End Sub

Public Function headingUpDown(s As String, upDown As Long)
    If Left(s, 1) = "#" Then
        Dim ss() As String
        ss = Split(s, ",")
        Dim indent As Long
        indent = Len(ss(0))
        Dim indentUpdate As Long
        indentUpdate = indent + upDown
        If indentUpdate = 0 Then
            indentUpdate = 1
        End If
        headingUpDown = String(indentUpdate, "#")
    Else
        headingUpDown = s
    End If
End Function
