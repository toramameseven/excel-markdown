VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLongInput2 
   Caption         =   "UserForm1"
   ClientHeight    =   6780
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12972
   OleObjectBlob   =   "frmLongInput2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLongInput2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private m_myRng As Range
Private m_OriginalText As String
Private m_OriginalF As String
Private m_ProjectWorkbookName As String
Private Const REPLACEWORD As String = "||||"

Public Property Set myRng(ByVal vRange As Range)
    Set m_myRng = vRange.Resize(1, 1)
    Me.Caption = m_myRng.Worksheet.Name & " " & m_myRng.Address

    m_OriginalText = m_myRng.Text
    Me.txtInput.Text = m_OriginalText


    On Error GoTo EH

    Dim sheetName
    sheetName = ActiveSheet.Name
    
    Dim linklist() As String
    linklist = modAsciidoc.GetLinkIndex

    Worksheets(sheetName).Select


    Dim i As Long
    For i = 0 To UBound(linklist)
         Me.lstLinkFiles.AddItem linklist(i)
    Next i
EH:
    '
End Property

Private Sub CopyTo(offsetRow As Long, offsetColumn As Long, isTrim As Boolean)
    Dim trng As Range
    Set trng = m_myRng.Offset(offsetRow, offsetColumn)

    If trng.Text = "" Then
        trng.Cells(1, 1) = Me.txtInput.SelText
    Else
        Exit Sub
    End If

    Me.txtInput.SelText = ""

    If isTrim Then
        Me.txtInput.Text = DelWsLeft(Me.txtInput.Text)
    Else
        ''
    End If
End Sub

Private Sub bntSubscript_Click()
        inlinecommand "~" & REPLACEWORD & "~"
End Sub

Private Sub btnBold_Click()
    inlinecommand "*" & REPLACEWORD & "*"
End Sub
Private Sub btnItalic_Click()
    inlinecommand "_" & REPLACEWORD & "_"
End Sub


Private Sub inlinecommand(ByVal vCommand As String)
    Me.txtInput.SelText = Replace(vCommand, REPLACEWORD, Me.txtInput.SelText)
End Sub

Private Sub btnStrikethrough_Click()
    inlinecommand "[line-through]#" & REPLACEWORD & "#"
End Sub

Private Sub btnSuperscript_Click()
        inlinecommand "^" & REPLACEWORD & "^"
End Sub

Private Sub btnToBottom_Click()
    Call CopyTo(1, 0, chkTrim.value)
End Sub

Private Sub btnToLeft_Click()
    Call CopyTo(0, -1, chkTrim.value)
End Sub

Private Sub btnToRight_Click()
    Call CopyTo(0, 1, chkTrim.value)
End Sub

Private Sub btnToTop_Click()
    Call CopyTo(-1, 0, chkTrim.value)
End Sub


Private Sub btnUnderline_Click()
        inlinecommand "[underline]#" & REPLACEWORD & "#"
End Sub

Private Sub cmdAddref_Click()

    If trim(txtLinkfile.Text) = "" Then
        MsgBox "Select Reference."
        Exit Sub
    End If
    Dim LinkHtml As String
    Dim LinkHtmls() As String
    LinkHtmls = Split(txtLinkfile.Text, "|")
    LinkHtml = LinkHtmls(2)
    Dim linkTitle As String
    Dim LinkTitles() As String
    LinkTitles = Split(txtLinkfile.Text, "|")
    linkTitle = LinkTitles(0)

    If LinkHtml <> "" Then
        If Me.txtInput.SelText <> "" Then
            Me.txtInput.SelText = Replace(Replace("`REF::####::$$$$`", "####", LinkHtml), "$$$$", Me.txtInput.SelText)
        Else
            Me.txtInput.SelText = Replace(Replace("`REF::####::$$$$`", "####", LinkHtml), "$$$$", "") 'auto set title
        End If
    Else
        '
    End If
End Sub

Private Sub cmdBox_Click()
    Me.txtInput.SelText = Replace("`box::####`", "####", Me.txtInput.SelText)
End Sub

Private Sub cmdbutton_Click()
    Me.txtInput.SelText = Replace("`button::####`", "####", Me.txtInput.SelText)
End Sub

Private Sub cmdCheckBox_Click()
    Me.txtInput.SelText = Replace("`check::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdcmdTab_Click()
    Me.txtInput.SelText = Replace("`tab::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdImgInline_Click()
    Me.txtInput.SelText = Replace("`img01::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdImgline_Click()
    Me.txtInput.SelText = Replace("`img::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdImgNote_Click()
    Me.txtInput.SelText = "`imgnote::`"
    txtInput.SetFocus
End Sub

Private Sub cmdInput_Click()
    Me.txtInput.SelText = Replace("`input::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdItemWindow_Click()
    Me.txtInput.SelText = Replace("`window::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdMdRef_Click()
    ReplaceMarkdownRef
End Sub

Private Sub cmdMenuItem_Click()
    Me.txtInput.SelText = Replace("`menu::####`", "####", Me.txtInput.SelText)
    txtInput.SetFocus
End Sub

Private Sub cmdDotting_Click()
    Dim s As String
    s = Me.txtInput.SelText
    s = modMain.DottingMultiLineString(s, vbCr)
    Me.txtInput.SelText = s
End Sub

Private Sub cmdDottingClear_Click()
    Dim s As String
    s = Me.txtInput.SelText
    s = modMain.DottingMultiLineString(s, vbCr, False)
    Me.txtInput.SelText = s
End Sub

Private Sub cmdItem_Click()
    Me.txtInput.SelText = Replace("`item::####`", "####", Me.txtInput.SelText)
End Sub

Private Sub cmdKakko1_Click()
    Call ReplaceKakko("［", "］")
End Sub

Private Sub cmdKakko2_Click()
    Call ReplaceKakko("「", "」")
End Sub
Private Sub cmdKakko3_Click()
    Call ReplaceKakko("(", ")")
End Sub

Private Sub ReplaceKakko(kakko1 As String, kakko2 As String)
    Dim s As String
    s = kakko1 & Me.txtInput.SelText & kakko2
    Me.txtInput.SelText = s
End Sub

Private Sub ReplaceMarkdownRef()
    Dim s As String
    s = "[" & Me.txtInput.SelText & "](#" & Me.txtInput.SelText & ")"
    Me.txtInput.SelText = s
End Sub

Private Sub cmdMnueClick_Click()
    Me.txtInput.SelText = "`menuclick::`"
End Sub

Private Sub cmdNumber_Click()
    Dim s As String
    s = Me.txtInput.SelText
    s = modMain.NumberingMultiString(s, vbCr, True, ".")
    Me.txtInput.SelText = s
End Sub

Private Sub cmdNumClear_Click()
    Dim s As String
    s = Me.txtInput.SelText
    s = modMain.NumberingMultiString(s, vbCr, False, ".")
    Me.txtInput.SelText = s
End Sub

Private Sub cmdOK_Click()

    m_OriginalText = Replace(Me.txtInput.Text, vbCrLf, vbLf)
    m_myRng.Cells(1, 1) = m_OriginalText

    Unload Me
End Sub

Private Sub lstLinkFiles_Click()
    Me.txtLinkfile = Me.lstLinkFiles.Text ''Split(Me.lstLinkFiles.Text, "|")(2)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If (m_OriginalText = Me.txtInput.Text) Then
        '
    Else
        '        If MsgBox("Modified. Do you close?", vbOKCancel) = vbCancel Then
        '            Cancel = 1
        '        End If
    End If
End Sub
