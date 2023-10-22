Attribute VB_Name = "modChartExcel"
Option Explicit



Public Sub makeShapes(Optional defaultShapeType As MsoAutoShapeType = msoShapeRectangle)
    Dim rng As Range
    Set rng = Selection

    Dim i As Long
    Dim j As Long
    Dim CellText As String
    Dim isAct As Boolean

    For i = 1 To rng.rows.Count
        For j = 1 To rng.columns.Count
            isAct = True
            If rng.Cells(i, j).MergeCells Then
                If rng.Cells(i, j) <> rng.Cells(i, j).MergeArea.Cells(1, 1) Then
                    isAct = False
                End If
            End If

            ''
            If isAct Then
                CellText = rng.Cells(i, j).Text
                If CellText <> "" Then
                    Select Case CellText
                        Case "<--", "-->", "---", "^||", "||v", "|||", "<==", "==>", "===", "^##", "##v", "###"
                            '' make links
                            makeArrow rng.Cells(i, j), CellText
                        Case Else
                            makeShapeType rng.Cells(i, j), defaultShapeType
                    End Select
                End If
            End If
        Next
    Next
End Sub


Private Function makeShapeType(rng As Range, Optional defaultShapeType As MsoAutoShapeType = msoShapeRectangle)
    Dim CellText As String
    CellText = rng.Cells(1, 1).Text

    Dim shapeType As MsoAutoShapeType

    Dim dicPicType As Object
    Set dicPicType = CreateObject("Scripting.Dictionary")
    dicPicType.CompareMode = vbTextCompare 'vbTextCompare(1) Åc ignore case vbBinaryCompare(0) Åc case sensitive

    dicPicType.Add "\((.*)\)", msoShapeRoundedRectangle
    dicPicType.Add "\{\{(.*)\}\}", msoShapeOctagon
    dicPicType.Add "\{(.*)\}", msoShapeDiamond

    '' dicPicType.Add "\((.*)\)", msoShapeDiamond
    dicPicType.Add "\/(.*)\/", msoShapeParallelogram



    Dim TextInside As String

    Dim Key
    Dim strKey As String
    For Each Key In dicPicType.keys
        strKey = Key
        TextInside = RegfindSubPatString(strKey, CellText)
        If TextInside <> "" Then
            shapeType = dicPicType(Key)
            Exit For
        End If
    Next
    
    Dim rgbColor() As Long


    If TextInside = "" Then
        Dim tag() As String
        tag = Split(CellText, "|")
        If UBound(tag) > 0 Then
            If IsNumeric(trim(tag(1))) Then
                TextInside = tag(0)
                shapeType = CLng(trim(tag(1)))
            End If
        End If
        
        rgbColor = createRGB("")
        If UBound(tag) > 1 Then
            rgbColor = createRGB(tag(2))
        End If
    End If


    If TextInside = "" Then
        TextInside = CellText
        shapeType = defaultShapeType
    End If

    makeShape rng, TextInside, shapeType, rgbColor

End Function


Private Function createRGB(p As String) As Long()
    On Error GoTo Error_Exit

    Dim rgbColor() As Long
    ReDim rgbColor(0 To 2)
    rgbColor(0) = Val("&h" & Mid(p, 1, 2))
    rgbColor(1) = Val("&h" & Mid(p, 3, 2))
    rgbColor(2) = Val("&h" & Mid(p, 5, 2))
    createRGB = rgbColor
    Exit Function

Error_Exit:
    rgbColor(0) = 91
    rgbColor(1) = 155
    rgbColor(2) = 213
    createRGB = rgbColor
End Function

Private Function createLineRGB(p() As Long) As Long()
    Dim rgbColor() As Long
    ReDim rgbColor(0 To 2)
    Dim d As Long
    d = 20
    rgbColor(0) = IIf(p(0) - d >= 0, p(0) - d, 0)
    rgbColor(1) = IIf(p(1) - d >= 0, p(1) - d, 0)
    rgbColor(2) = IIf(p(2) - d >= 0, p(2) - d, 0)


    createLineRGB = rgbColor
End Function



Sub repalaceType_Test()
    repalacePropMain 1, "255,255,255", 10
End Sub

Public Sub repalacePropMain(defaultShapeType As Long, colorTo As String, Size As Single)
    Dim rng As Range
    Set rng = Selection

    Dim i As Long
    Dim j As Long
    Dim CellText As String
    Dim isAct As Boolean

    For i = 1 To rng.rows.Count
        For j = 1 To rng.columns.Count
            isAct = True
            If rng.Cells(i, j).MergeCells Then
                If rng.Cells(i, j) <> rng.Cells(i, j).MergeArea.Cells(1, 1) Then
                    isAct = False
                End If
            End If

            ''
            If isAct Then
                repalaceType rng.Offset(i - 1, j - 1), defaultShapeType
                repalaceColor rng.Offset(i - 1, j - 1), colorTo
            End If
        Next
    Next
End Sub
' string| shapeType| Color| fontSize
Public Sub repalaceType(rng As Range, shapeTypeReplace As Long)
    If shapeTypeReplace = 0 Then
        Exit Sub
    End If
    
    If rng.Cells(1, 1).Text = "" Then
        Exit Sub
    End If
    
    Dim tags() As String
    tags = Split(rng.Cells(1, 1).Text, "|")
    
    If UBound(tags) < 1 Then
       ReDim Preserve tags(0 To 1)
    End If
    
    tags(1) = shapeTypeReplace
    
    Dim s
    s = Join(tags, "|")
    rng.Cells(1, 1).value = s
End Sub

Sub repalaceColor_Test()
    repalaceColor Selection, "0,255,0"
End Sub

Public Sub repalaceColor(rng As Range, shapeColor As String)
    If shapeColor = "" Then
        Exit Sub
    End If
    
    If rng.Cells(1, 1).Text = "" Then
        Exit Sub
    End If
    
    Dim tags() As String
    tags = Split(rng.Cells(1, 1).Text, "|")
    
    If UBound(tags) < 2 Then
       ReDim Preserve tags(0 To 2)
    End If
    
    tags(2) = shapeColor
    
    Dim s
    s = Join(tags, "|")
    rng.Cells(1, 1).value = s
End Sub

Public Sub repalaceFontSize(rng As Range, fontSize As Single)
    If fontSize = 0 Then
        Exit Sub
    End If
    Dim tags() As String
    tags = Split(rng.Cells(1, 1).Text, "|")
    
    If UBound(tags) < 3 Then
       ReDim Preserve tags(0 To 3)
    End If
    
    tags(3) = fontSize
    
    Dim s
    s = Join(tags, "|")
    rng.Cells(1, 1).value = s
End Sub


Private Sub makeShape(rng As Range, insideText As String, shapeType As MsoAutoShapeType, colors() As Long)

    Dim CellText As String
    CellText = rng.Cells(1, 1).Text

    Dim Left As Single
    Dim Top As Single
    Dim Width As Single
    Dim Height As Single

    Left = rng.MergeArea.Left
    Top = rng.MergeArea.Top
    Width = rng.MergeArea.Width
    Height = rng.MergeArea.Height

    Dim tmpShape As Shape
    Set tmpShape = ActiveSheet.shapes.AddShape(shapeType, Left, Top, Width, Height)
    tmpShape.TextFrame2.TextRange.Characters.Text = insideText
    tmpShape.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    tmpShape.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    

    tmpShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    tmpShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    tmpShape.Placement = xlFreeFloating
    
    With tmpShape.Fill
        .Visible = msoTrue
        .ForeColor.rgb = rgb(colors(0), colors(1), colors(2))
        .Transparency = 0
        .Solid
    End With

    Dim rgbLine() As Long
    rgbLine = createLineRGB(colors)
   With tmpShape.Line
       .Visible = msoTrue
       .ForeColor.rgb = rgb(rgbLine(0), rgbLine(1), rgbLine(2))
       .Transparency = 0
   End With
End Sub

Private Sub makeArrow(rng As Range, arrowType As String)

    Dim CellText As String
    CellText = rng.Cells(1, 1).Text

    Dim Left As Single
    Dim Top As Single
    Dim Width As Single
    Dim Height As Single

    Left = rng.MergeArea.Left
    Top = rng.MergeArea.Top
    Width = rng.MergeArea.Width
    Height = rng.MergeArea.Height

    Dim startX As Single
    Dim startY As Single
    Dim endX As Single
    Dim endY As Single

    Dim startStyle As MsoArrowheadStyle
    Dim endStyle As MsoArrowheadStyle

    startStyle = msoArrowheadNone
    endStyle = msoArrowheadNone


    Dim LineWeight As Single
    LineWeight = 1



    Select Case arrowType
        Case "<--"
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5
            startStyle = msoArrowheadTriangle
        Case "-->"
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5
            endStyle = msoArrowheadTriangle
        Case "---"
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5

        Case "^||"
            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height
            startStyle = msoArrowheadTriangle

        Case "||v"
            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height
            endStyle = msoArrowheadTriangle
        Case "|||"
            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height

        '''
        Case "<=="
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5
            startStyle = msoArrowheadTriangle
            LineWeight = 6
        Case "==>"
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5
            endStyle = msoArrowheadTriangle
            LineWeight = 6
        Case "==="
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5
            LineWeight = 6

        Case "^##"
            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height
            startStyle = msoArrowheadTriangle
            LineWeight = 6

        Case "##v"
            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height
            endStyle = msoArrowheadTriangle
            LineWeight = 6
        Case "###"
            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height
            LineWeight = 6
    End Select



    Dim tmpShape As Shape
    Set tmpShape = ActiveSheet.shapes.AddConnector(msoConnectorStraight, startX, startY, endX, endY)
    tmpShape.Line.EndArrowheadStyle = endStyle
    tmpShape.Line.BeginArrowheadStyle = startStyle
    tmpShape.Line.weight = LineWeight
    tmpShape.Placement = xlFreeFloating


End Sub

