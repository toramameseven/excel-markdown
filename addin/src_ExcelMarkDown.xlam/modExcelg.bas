Attribute VB_Name = "modExcelg"
Option Explicit

Public Sub makeShapes()
    Dim rng As Range
    Set rng = Selection

    Dim i As Long
    Dim j As Long
    Dim CellText As String
    Dim isAct As Boolean
    
    Dim shapes As Object
    Set shapes = MakeShapeDictionary
    
    Dim keys
    keys = MakeShapeDictionary.keys
    Dim ShapeTypeReg
    
    Dim TextInside As String
    Dim shapeType As MsoAutoShapeType
    
    Dim FrontColor As Long
    Dim InteriorColor As Long
    
    
    For i = 1 To rng.rows.Count
        For j = 1 To rng.columns.Count
            isAct = True
            If rng.Cells(i, j).MergeCells Then
                If rng.Cells(i, j) <> rng.Cells(i, j).MergeArea.Cells(1, 1) Then
                    isAct = False
                End If
            End If

            If isAct Then
                CellText = rng.Cells(i, j).Text
                FrontColor = rng.Cells(i, j).Font.Color
                InteriorColor = rng.Cells(i, j).Interior.Color
                
                Dim l As Long

                
                
                If CellText <> "" Then
                    If RegfindSubPatString("([\<|\^|\-|\=|\||\#][\-|\=|\||\#][\>|v|\-|\=|\||\#])", CellText) <> "" Then
                        makeArrow rng.Cells(i, j), CellText
                        
                    Else
                        '' node
                        shapeType = msoShapeFlowchartProcess
                        
                        For Each ShapeTypeReg In keys
                            TextInside = RegfindSubPatString(CStr(ShapeTypeReg), CellText)
                            If TextInside <> "" Then
                                shapeType = MakeShapeDictionary.Item(ShapeTypeReg)
                                Exit For
                            End If
                        Next
                        
                        If TextInside = "" Then
                            TextInside = CellText
                            shapeType = msoShapeFlowchartProcess '[]
                        End If
                        makeShape rng.Cells(i, j), TextInside, shapeType, FrontColor, InteriorColor
                        
                    End If
                End If
            End If
        Next
    Next
End Sub


'(a)
'[(f)]
'([d])
'((a))
'[[s]]
'{{s}}
'/a/
'{s}
'[s]

Private Function MakeShapeDictionary() As Object
    
    Dim dic As Object
    
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare 'vbTextCompare(1)
    
    dic.Add "\(\((.*)\)\)", msoShapeFlowchartConnector
    dic.Add "\[\((.*)\)\]", msoShapeFlowchartMagneticDisk
    dic.Add "\(\[(.*)\]\)", msoShapeFlowchartTerminator
    dic.Add "\((.*)\)", msoShapeFlowchartAlternateProcess
    dic.Add "\[\[(.*)\]\]", msoShapeFlowchartPredefinedProcess
    dic.Add "\{\{(.*)\}\}", msoShapeFlowchartPreparation
    dic.Add "\{(.*)\}", msoShapeFlowchartDecision
    dic.Add "\/(.*)\/", msoShapeFlowchartData
    dic.Add "\[(.*)\]", msoShapeFlowchartProcess
    
    Set MakeShapeDictionary = dic
End Function




Private Sub makeShape(rng As Range, CellText As String, shapeType As MsoAutoShapeType, Optional FrontColor As Long, Optional InteriorColor As Long)

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
    tmpShape.TextFrame2.TextRange.Characters.Text = CellText
    'tmpShape.TextFrame.TextRange.Characters.text.Font.Color = FrontColor
    
    With tmpShape.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.rgb = FrontColor
        .Transparency = 0
        .Solid
    End With
    
    
    tmpShape.Fill.ForeColor.rgb = InteriorColor
    

    tmpShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    tmpShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub



' - =  ,left right
' | #  , top bottom
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
    
    Dim isLR As Boolean
    Dim isTB As Boolean
    
    Dim LineWidht As Single
    LineWidht = 1
    
    
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    
    s1 = Mid(arrowType, 1, 1)
    s2 = Mid(arrowType, 2, 1)
    s3 = Mid(arrowType, 3, 1)
    
    
    If s2 = "=" Or s2 = "#" Then
        LineWidht = 6
    End If
    
    If s2 = "|" Or s2 = "#" Then
        isTB = True
    Else
        isLR = True
    End If
    
    If s1 = "<" Or s1 = "^" Then
        startStyle = msoArrowheadTriangle
    End If
    
    If s3 = ">" Or s3 = "v" Then
        endStyle = msoArrowheadTriangle
    End If

    If isLR Then
            startX = Left
            endX = Left + Width
            startY = Top + Height * 0.5
            endY = Top + Height * 0.5
    Else

            startX = Left + Width * 0.5
            endX = Left + Width * 0.5
            startY = Top
            endY = Top + Height
    End If



    Dim tmpShape As Shape
    Set tmpShape = ActiveSheet.shapes.AddConnector(msoConnectorStraight, startX, startY, endX, endY)
    tmpShape.Line.EndArrowheadStyle = endStyle
    tmpShape.Line.BeginArrowheadStyle = startStyle
    tmpShape.Line.weight = LineWidht
End Sub


Sub RegfindSubPatStringTest()

    'Debug.Print RegfindSubPatString("\[(.*)\]", "[aa]")
    Debug.Print RegfindSubPatString("([\<|\^|\-|\=|\||\#][\-|\=|\||\#][\>|\v|\-|\=|\||\#])", "W==")
    
End Sub


'''
'''
'''
'''
Function RegfindSubPatString(myPattern As String, myString As String) As String
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
    
    RegfindSubPatString = retVal.Joins("|")
End Function



