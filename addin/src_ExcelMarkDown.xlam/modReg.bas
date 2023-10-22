Attribute VB_Name = "modReg"
Option Explicit




Private Function RegDeleteSectionNo(myPattern As String, myString As String) As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp
    objRegExp.Pattern = myPattern
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    
    Dim matchval As String
    Dim replaceVal As String
    
    Dim retVal As String
    retVal = Replace(myString, ChrW(160), "")  ' nbsp

    'Test whether the String can be compared.
    If (objRegExp.test(retVal) = True) Then

        'Get the matches.
        Set colMatches = objRegExp.Execute(retVal)   ' Execute search.]


        For Each objMatch In colMatches   ' Iterate Matches collection.
            matchval = objMatch.submatches(0)
            replaceVal = ""
            retVal = Replace(retVal, matchval, replaceVal)
        Next
    End If
    
    RegDeleteSectionNo = retVal
End Function


Sub RegDeleteSectionNo_TEST()
    ''Debug.Print RegDeleteSectionNo("((\d)+\.(\d)*\s+).*", "2.5 New features")
    Debug.Print "2.5.1 introduction"
    
    Selection = RegDeleteSectionNo("(((\d+)\.)+(\d)*\s+).*", Selection)
End Sub

Public Function DeleteSectionNo(message As String) As String
    Const regs As String = "(((\d+)\.)+(\d)*\s+).*"
    DeleteSectionNo = RegDeleteSectionNo(regs, message)
End Function




Sub RegFindCapture_Test()
    Debug.Print RegFindCapture("'(.*?)'", "'bbb'")
    '' >> bbb
    '
    Debug.Print RegFindCapture("'(.*?)'", "'bbb','bbbb'")
    '' >> bbb|bbbb

    Debug.Print RegFindCapture("{{(.+),.+}}", "<item index=""1"" data =""{{TestSample,SampleName}}"" />")
    '' >> bbb|bbbb
End Sub

Public Function RegFindCapture(myPattern As String, myString As String) As String
    On Error GoTo EH
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
    
    Dim retVal As clsString
    Set retVal = New clsString

    'Test whether the String can be compared.
    If (objRegExp.test(myString) = True) Then

        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.]


        For Each objMatch In colMatches   ' Iterate Matches collection.
            matchval = objMatch.value
            retVal.Add objMatch.submatches(0)
        Next
    Else
        '
    End If
    
    RegFindCapture = retVal.Joins("|")
    
    Exit Function
EH:
    RegFindCapture = "ERRRRRRRRRRRRR"
End Function







