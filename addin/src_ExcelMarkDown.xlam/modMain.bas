Attribute VB_Name = "modMain"
Option Explicit
#If VBA7 And Win64 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub ufExcelMarkDownShow()
    Dim myufAsciidoc As ufAsciidoc
    Set myufAsciidoc = New ufAsciidoc
    myufAsciidoc.Show vbModeless
End Sub

Public Sub LongInput()
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

Public Function GetCells(ByRef rRng As Range, Optional maxCount = 1000) As Collection
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

Public Function DelWsLeft(vstring As String) As String
    Dim sss As String
    sss = Mid$(vstring, 1, 1)
    If sss = vbCrLf Or sss = vbCr Or sss = vbLf Or sss = vbNewLine Or sss = " " Or sss = "　" Then
        DelWsLeft = DelWsLeft(Mid$(vstring, 2))
    Else
        DelWsLeft = vstring
    End If
End Function



Private Function GetColumn(strCellValue As String) As Long
    Dim icolumns As Long
    For icolumns = 1 To 255
        If ActiveSheet.Cells(1, icolumns) = strCellValue Then
            GetColumn = icolumns
            Exit Function
        End If
    Next
    GetColumn = 0
End Function

Public Function isExistSheet(ByVal tsheetname) As Boolean
    Dim st As Worksheet
    For Each st In ActiveWorkbook.Sheets
        If UCase(st.Name) = UCase(tsheetname) Then
            isExistSheet = True
            Exit Function
        End If
    Next
    isExistSheet = False
End Function


Public Function NumberingMultiString(ByVal strings As String, ByVal delimita As String, isNumbering As Boolean, seps As String) As String
    Dim iCount As Long
    iCount = 1

    Dim pos As Long
    Dim s As String
    Dim i As Long


    Dim SEPSTR As String
    SEPSTR = seps


    Dim w() As String

    w() = Split(strings, delimita)
    iCount = 1
    For i = 0 To UBound(w)
        s = trim(w(i))
        If s = "" Then
            ''
        Else
            pos = InStr(1, s, SEPSTR, vbTextCompare)

            If pos >= 1 And pos <= 3 Then
                s = Mid$(s, pos + 1)
            End If

            If isNumbering Then
                s = iCount & SEPSTR & " " & trim(s)
            Else
                s = trim(s)
            End If

            w(i) = s
            iCount = iCount + 1
        End If
    Next i
    NumberingMultiString = Join(w, delimita)
End Function


Public Function ItemiseMultiString(ByVal strings As String, ByVal delimita As String, isItemize As Boolean) As String
    Dim iCount As Long
    iCount = 1

    Dim pos As Long
    Dim s As String
    Dim i As Long


    Dim charItem As String
    charItem = "・"


    Dim w() As String

    w() = Split(strings, delimita)
    For i = 0 To UBound(w)
        s = trim(w(i))
        If s = "" Then
            ''
        Else
            If Left(s, 1) = charItem Then
                If isItemize Then
                    ''
                Else
                    s = trim(Mid(s, 2))
                End If
            Else
                If isItemize Then
                    s = charItem & " " & s
                Else
                    ''
                End If
            End If


            w(i) = s
        End If
    Next i
    ItemiseMultiString = Join(w, delimita)
End Function



Public Function AddOneLine(ByVal strings As String, isAdd As Boolean) As String

    Dim pos As Long
    Dim s As String
    Dim i As Long
        
    
    Dim retVal As String
    retVal = ""

    Dim w() As String

    w() = Split(strings, vbLf)   'vbLf Chr(10)
    
    
    If strings = "" Then
        AddOneLine = ""
        Exit Function
    End If
    
    If isAdd Then
        s = trim(w(UBound(w)))
        If s = "" Then
            retVal = strings
        Else
            retVal = strings & vbLf
        End If
    Else
        For i = UBound(w) To 0 Step -1
            s = trim(w(i))
            If s <> "" Then
            ReDim o(0 To i) As String
            Dim j As Long
            For j = 0 To i
                o(j) = w(j)
            Next j
                retVal = Join(o, vbLf)
            End If
        Next i
    End If
    
    
    If retVal = "" Then
        retVal = strings
    End If
    
    AddOneLine = retVal
End Function


Public Function MaruMultiString(ByVal strings As String, ByVal delimita As String, Optional isMaru As Boolean = True) As String
    ' add kuten
    Dim iCount As Long
    iCount = 1

    Dim pos As Long
    Dim s As String
    Dim i As Long


    Dim w() As String

    w() = Split(strings, delimita)
    iCount = 1
    For i = 0 To UBound(w)
        s = trim(w(i))
        If s = "" Then
            ''
        Else
            If Right(s, 1) = "。" Then
                If Not isMaru Then
                    s = Left(s, Len(s) - 1)
                End If
            Else
                If isMaru Then
                    s = s & "。"
                End If
            End If

            iCount = iCount + 1
        End If
        w(i) = s
    Next i
    MaruMultiString = Join(w, delimita)
End Function

Public Function DottingMultiLineString(ByVal strings As String, ByVal delimita As String, Optional isDotting As Boolean = True) As String
    Dim iCount As Long
    iCount = 1

    Dim pos As Long
    Dim s As String
    Dim i As Long


    Dim w() As String

    w() = Split(strings, delimita)
    iCount = 1
    For i = 0 To UBound(w)
        s = trim(w(i))
        If s = "" Then
            ''
        Else
            pos = InStr(1, s, "・", vbTextCompare)

            If pos >= 1 And pos <= 2 Then
                s = Mid$(s, pos + 1)
            End If

            If isDotting Then
                s = "・ " & trim(s)
            Else
                s = trim(s)
            End If

            w(i) = s
            iCount = iCount + 1
        End If
    Next i
    DottingMultiLineString = Join(w, delimita)
End Function







