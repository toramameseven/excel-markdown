Attribute VB_Name = "modAsciidocTools"
Option Explicit


'Private Const SECTION_TAG As String = "#"
'Private Const SECTION_TAG2 As String = "##"
'Private Const SECTION_TAG3 As String = "###"
'Private Const SECTION_TAG4 As String = "####"
'Private Const SECTION_TAG5 As String = "#####"

Public Sub MakeMd()
    modAsciidoc.MakeDocument "md"
End Sub

Sub InsertLineAboveBelow()
    On Error GoTo EH
    Dim sizeInsert As Long
    sizeInsert = 1
    
    Dim r As Range
    Set r = Selection
    


   r.Offset(1).Resize(sizeInsert).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
   r.Offset(1).Resize(sizeInsert).EntireRow.ClearFormats
   r.Offset(0).Resize(sizeInsert).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
   r.Offset(-sizeInsert).Resize(sizeInsert).EntireRow.ClearFormats
   
   r.Select
    
EH:
End Sub


Public Sub SetRowsColor()
    
    Dim rng As Range
    
    
    ' if  selection is range
    If TypeName(Selection) = "Range" Then
        '' go next
    Else
        MsgBox "select range or a cell. Not Picture or some."
        Exit Sub
    End If
    
    Set rng = Selection
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0"
    rng.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With rng.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
    End With
    rng.FormatConditions(1).StopIfTrue = False
    rng.Resize(1, 1).Select
    
End Sub


' 1 convert to  ==
' 1.1.1 convert to  ====
Public Sub SetChapter()

    Dim s As String

    Dim rng As Range

    'On Error GoTo This_Error
    
    Dim thisPro As String
    
    thisPro = "SetChapter"
    Application.ScreenUpdating = False
    Application.StatusBar = thisPro

    Set rng = Selection
    Dim r As Long
    Dim c As Long

    r = rng.rows.Count
    c = rng.columns.Count
    
    Dim isRow As Boolean
    isRow = True
    Dim n As Long
    n = IIf(isRow, c, r)

    Dim nLines As Long
    nLines = IIf(isRow, r, c)

    Dim w() As String
    ReDim w(0 To n - 1)

    Dim i As Long
    Dim j As Long
    Dim sForClip As clsString
    Set sForClip = New clsString
    

    For j = 0 To nLines - 1
        coreModuleSetChapter rng.Offset(j, 0)
    Next j
    
    
    For j = nLines - 1 To 0
        coreModuleSetChapter rng.Offset(j, 0)
    Next j
       
    
    For j = nLines - 1 To 0 Step -1
       InsertLines rng.Offset(j, 0)
    Next j

'    save clipBoard
'    With New MSForms.DataObject
'        .SetText sForClip.Joins
'        .PutInClipboard
'    End With

   On Error GoTo 0
   GoTo normal_exit

This_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure " & thisPro
    
normal_exit:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub SetChapterEX()

    Dim s As String
    Dim rng As Range

    'On Error GoTo This_Error
    Application.ScreenUpdating = False

    Dim i As Long
    Dim j As Long

    For j = 0 To 1300
        coreModuleSetChapter Range("C1").Offset(j, 0)
    Next j
    
    
   On Error GoTo 0
   GoTo normal_exit

This_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure "
    
normal_exit:
    Application.ScreenUpdating = True
    Application.StatusBar = False

End Sub


Function coreModuleSetChapter(rng As Range)

    Dim s As String
    Dim outs As String
    Dim flg As String
    
    s = trim(rng.Cells(1, 1))
    
    flg = ""
    
    If RegMatch("^\d+[.-]d+[.-]d+[.-]\d+\s.+", s, outs) Then
        flg = SECTION_TAG5
    ElseIf RegMatch("^\d+[.-]\d+[.-]\d+(.+)", s, outs) Then
        flg = SECTION_TAG4
    ElseIf RegMatch("^\d+[.-]\d+(.+)", s, outs) Then
        flg = SECTION_TAG3
    ElseIf RegMatch("^\d+[\D](.+)", s, outs) Then
         flg = "."
    Else
    
    End If
    
    If flg <> "" Then
        rng.Offset(0, -2) = flg
        rng.Offset(0, 5) = s
        rng = trim(outs)
    End If
       
End Function






Sub RegFind_test()

    Debug.Print "start----------------"
    Debug.Print RegFind("^\d+\s.+", "2 line1")
    Debug.Print RegFind("^\d+\.\d+\s.+", "2.5 line2")
    Debug.Print RegFind("^\d+.\d+\.\d+\s.+", "2.5.6 line3")
    Debug.Print RegFind("^\d+.\d+.\d+\.\d+\s.+", "2.5.6.9 line4")
End Sub

Function InsertLines(rng As Range) As String
    If Left(rng.Offset(0, -2).Cells(1, 1), 1) = SECTION_TAG Then
        rng.Resize(1, 1).Offset(1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        rng.Resize(1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
End Function


Public Sub InsertSectionTag_Test()
    InsertSectionTag 2
End Sub


Public Sub InsertFalgInsertLineUpDown(ByVal flag As String)
    Dim rng As Range
    Set rng = Selection
    
    Cells(rng.row, 1) = flag

    rng.Resize(1, 1).Offset(1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    rng.Resize(1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    rng.Select
End Sub


Public Sub InsertSectionTag(ByVal sectionIndex As Long)
    Dim rng As Range
    Set rng = Selection
    
    Cells(rng.row, 1) = String(sectionIndex, SECTION_TAG)

    rng.Resize(1, 1).Offset(1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    rng.Resize(1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    rng.Select
End Sub

Public Sub InsertTableTag()
    Dim rng As Range
    Set rng = Selection
    
    Cells(1, 1) = "table"
    Cells(rng.row + rng.rows.Count - 1, 1) = "table"

    rng.Resize(1, 1).Offset(rng.rows.Count, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    rng.Resize(1, 1).Offset(0, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ''rng.Resize(1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    rng.Select
End Sub


Public Sub InsertTableTags()
    Dim rng As Range
    Set rng = Selection
    
    Cells(rng.row, 1) = "table"
    Cells(rng.row + rng.rows.Count - 1, 1) = "table"

    rng.Select
End Sub


Public Sub InsertTableTagBegin()
    Dim rng As Range
    Set rng = Selection
    
    Cells(rng.row, 1) = "table"

    rng.Resize(1, 1).Offset(0, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ''rng.Resize(1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    rng.Select
End Sub

Public Sub InsertTableTagEnd()
    Dim rng As Range
    Set rng = Selection
    
    Cells(rng.row, 1) = "table"

    rng.Resize(1, 1).Offset(1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ''rng.Resize(1, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    rng.Select
End Sub


Function RegFind(myPattern As String, myString As String) As Boolean
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp
    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = myPattern
    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True
    'Set global applicability.
    objRegExp.Global = True
    

    
    Dim retVal As Boolean
    retVal = False

    'Test whether the String can be compared.
    If (objRegExp.test(myString) = True) Then
        retVal = True

'        'Get the matches.
'        Set colMatches = objRegExp.Execute(myString)   ' Execute search.]
'
'
'        For Each objMatch In colMatches   ' Iterate Matches collection.
'            matchval = objMatch.Value
'            replaceVal = "<<" & objMatch.SubMatches(0) & "," & objMatch.SubMatches(1) & ">>"
'            retval = Replace(retval, matchval, replaceVal)
'        Next
    End If
    RegFind = retVal
End Function

Sub RegFindLink_Test()
    Dim r As Variant
    r = RegLinkToMarkDown("<<ddddd1,dddddd>>ddddddddddddd<<ddddd2,dddddd>>ddddddddddddd<<,ssあああss>>")
    Debug.Print r
        r = RegLinkToMarkDown("ssssssssssssssssssssssssssssssssssssssssssssssssssss")
    Debug.Print r
End Sub


Function RegLinkToMarkDown(myString As String) As String
    
    If InStr(myString, "<<") = False Then
        RegLinkToMarkDown = myString
    End If
    

    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp
    objRegExp.Pattern = "<<(.*?),(.*?)>>"
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    
    Dim retVal As String

    retVal = myString

    'Test whether the String can be compared.
    If (objRegExp.test(myString) = True) Then
        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.]
        
        '' <<ddddd1,dddddd>>
        '' objMatch.value =  <<ddddd1,dddddd>>
        '' objMatch.submatches(0) ddddd1
        '' objMatch.submatches(1) dddddd
        Dim matchAll As String
        Dim replaceVal  As String
        For Each objMatch In colMatches   ' Iterate Matches collection.
            matchAll = objMatch.value
            replaceVal = "[" & objMatch.submatches(1) & "](#" & Slugify(objMatch.submatches(0)) & ")"
            retVal = Replace(retVal, matchAll, replaceVal)
        Next
    End If
    RegLinkToMarkDown = retVal
End Function


Sub Slugify_Test()
    Debug.Print Slugify("ddd ddd  ddd（dd       \")
End Sub


Function Slugify(sLink As String) As String
    Dim s As String
    s = trim(sLink)
    s = LCase(s)
    s = RegReplace("[\]\[\!""\#\$\%\&\'\(\)\*\+\,\.\/\:\;\<\=\>\?\@\\\^\_\{\|\}\~（）［］]", s, "")
    s = RegReplace("\s+", s, "-")
    Slugify = RegReplace("\-+$", s, "")
End Function



Function RegReplace(myPattern As String, myString As String, myReplace As String) As String
    RegReplace = False

    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = myPattern

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True
    
    
    Dim matchval As String
    Dim replaceVal As String
    
    Dim retVal As String
    retVal = myString

    RegReplace = objRegExp.Replace(myString, myReplace)
End Function



Sub RegMatch_Test()
    Dim s As String
    Debug.Print RegMatch("^\d+.\d+.\d+\.\d+\.\s*(.+)", "2.5.6.9.    LinNN", s)
    Debug.Print s
End Sub

Function RegMatch(myPattern As String, myString As String, myMatch As String) As Boolean
    RegMatch = False

    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = myPattern

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True
    
    
    Dim matchval As String
    Dim replaceVal As String
    
    Dim retVal As String
    retVal = myString

    'Test whether the String can be compared.
    If (objRegExp.test(myString) = True) Then
        RegMatch = True

        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.]
        
        For Each objMatch In colMatches   ' Iterate Matches collection.
'            matchval = objMatch.Value
'            replaceVal = "<<" & objMatch.SubMatches(0) & "," & objMatch.SubMatches(1) & ">>"
'            retVal = Replace(retVal, matchval, replaceVal)
            myMatch = objMatch.submatches(0)
        Next
    End If

End Function


Sub MakeSampleCode()

    On Error GoTo EH

    Dim a As clsString
    Set a = New clsString

    a.Add "'= AsciiDoc Article Title"
    a.Add "Firstname Lastname <author@asciidoctor.org>"
    a.Add "1.0, July 29, 2014, Asciidoctor 1.5 article template"
    a.Add ":toc:"
    a.Add ":icons: font"
    a.Add ":quick-uri: https://asciidoctor.org/docs/asciidoc-syntax-quick-reference/"
    a.Add ""
    a.Add "Content entered directly below the header but before the first section heading is called the preamble."
    a.Add ""
    a.Add "'== First level heading"
    a.Add ""
    a.Add "This is a paragraph with a *bold* word and an _italicized_ word."
    a.Add ""
    a.Add ".Image caption"
    a.Add "image::image-file-name.png[I am the image alt text.]"
    a.Add ""
    a.Add "This is another paragraph.footnote:[I am footnote text and will be displayed at the bottom of the article.]"
    a.Add ""
    a.Add "'=== Second level heading"
    a.Add ""
    a.Add ".Unordered list title"
    a.Add "* list item 1"
    a.Add "** nested list item"
    a.Add "*** nested nested list item 1"
    a.Add "*** nested nested list item 2"
    a.Add "* list item 2"
    a.Add ""
    a.Add "This is a paragraph."
    a.Add ""
    a.Add ".Example block title"
    a.Add "'===="
    a.Add "Content in an example block is subject to normal substitutions."
    a.Add "'===="
    a.Add ""
    a.Add ".Sidebar title"
    a.Add "****"
    a.Add "Sidebars contain aside text and are subject to normal substitutions."
    a.Add "****"
    a.Add ""
    a.Add "'==== Third level heading"
    a.Add ""
    a.Add "[#id-for-listing-block]"
    a.Add ".Listing block title"
    a.Add "'----"
    a.Add "Content in a listing block is subject to verbatim substitutions."
    a.Add "Listing block content is commonly used to preserve code input."
    a.Add "'----"
    a.Add ""
    a.Add "'===== Fourth level heading"
    a.Add ""
    a.Add ".Table title"
    a.Add "|==="
    a.Add "|Column heading 1 |Column heading 2"
    a.Add ""
    a.Add "|Column 1, row 1"
    a.Add "|Column 2, row 1"
    a.Add ""
    a.Add "|Column 1, row 2"
    a.Add "|Column 2, row 2"
    a.Add "|==="
    a.Add ""
    
    'aas<<#primitives-nulls,aaa>>
    
    a.Add "[#mylabel]"
    a.Add "'====== Fifth level heading"
    a.Add ""
    a.Add "[quote, firstname lastname, movie title]"
    a.Add "____"
    a.Add "I am a block quote or a prose excerpt."
    a.Add "I am subject to normal substitutions."
    a.Add "____"
    a.Add ""
    a.Add "[verse, firstname lastname, poem title and more]"
    a.Add "____"
    a.Add "I am a verse block."
    a.Add "  Indents and endlines are preserved in verse blocks."
    a.Add "____"
    a.Add ""
    a.Add ""
    a.Add "<<#mylabel,DisplayName>>"
    a.Add ""
    a.Add ""
    a.Add "'== First level heading"
    a.Add ""
    a.Add "TIP: There are five admonition labels: Tip, Note, Important, Caution and Warning."
    a.Add ""
    a.Add "// I am a comment and won't be rendered."
    a.Add ""
    a.Add ". ordered list item"
    a.Add ".. nested ordered list item"
    a.Add ". ordered list item"
    a.Add ""
    a.Add "The text at the end of this sentence is cross referenced to <<_third_level_heading,the third level heading>>"
    a.Add ""
    a.Add "'== First level heading"
    a.Add ""
    a.Add "This is a link to the https://asciidoctor.org/docs/user-manual/[Asciidoctor User Manual]."
    a.Add "This is an attribute reference {quick-uri}[which links this text to the Asciidoctor Quick Reference Guide]."
    
    
    
    If isExistSheet("sample_adoc") Then
        MsgBox ("sample_adoc exists already.")
        Exit Sub
    End If
    
    Dim sht As Worksheet
    Set sht = Worksheets.Add
    sht.Name = "sample_adoc"
    
    Dim i As Long
    
    For i = 1 To a.Count
        Worksheets("sample_adoc").Cells(i, 1) = a.Item(i)
    Next i
EH:
    
End Sub



Sub execPsScript()
    Dim psFile As String
    Dim execCommand As String
    Dim wsh As Object
    Dim result As Integer
    

    psFile = "C:\Users\user\Desktop\test.ps1"
    
    Set wsh = CreateObject("WScript.Shell")
    

    execCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted " & psFile
    

    result = wsh.Run(command:=execCommand, WindowStyle:=0, WaitOnReturn:=True)
    
    If (result = 0) Then
        MsgBox ("Power Shell  success")
    Else
        MsgBox ("Power Shell  fail")
    End If
    
    Set wsh = Nothing
End Sub


Sub splitKuten()
    Dim target As String
    Dim r As Range
    Set r = Selection
    target = Selection
    
    Dim strs() As String
    strs = Split(target, "。")
    
    r.Offset(1).Resize(UBound(strs) + 2).EntireRow.Insert xlShiftDown
    
    Dim i As Long
    For i = 0 To UBound(strs)
        If strs(i) <> "" Then
            r.Offset(i + 1).Cells(1, 1) = strs(i) & "。"
        End If
        
    Next i
End Sub

Sub GetColumnsWidthsForAdoc_Test()
    Debug.Print GetColumnsWidthsForAdoc(Selection)

End Sub


Public Function GetColumnsWidthsForAdoc(rng As Range) As String
    Dim i As Long
    Dim ColumnsWidths() As String
    Dim ColumnsWidthsSingle() As Single
    Dim AdocCols() As String
    ReDim ColumnsWidths(0 To rng.columns.Count - 1)
    ReDim AdocCols(0 To rng.columns.Count - 1)
    ReDim ColumnsWidthsSingle(0 To rng.columns.Count - 1)

    Dim t As Long

    t = rng.rows.Count
    
    Dim sumOfColumns As Single
    sumOfColumns = 0

    For i = 1 To rng.columns.Count
        sumOfColumns = sumOfColumns + rng.columns(i).ColumnWidth
        ColumnsWidthsSingle(i - 1) = rng.columns(i).ColumnWidth
        ColumnsWidths(i - 1) = ColumnsWidthsSingle(i - 1)
    Next i
    
    For i = 1 To rng.columns.Count
        AdocCols(i - 1) = Int(ColumnsWidthsSingle(i - 1) / sumOfColumns * 40)
    Next

    GetColumnsWidthsForAdoc = "[cols=""" & Join(AdocCols, ",") & """]"
End Function
