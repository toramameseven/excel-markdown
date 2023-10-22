Attribute VB_Name = "modPicture"
Option Explicit
'' http://excel.syogyoumujou.com/memorandum/without_picture.html


'
Private Const QUALITY_PARAMS As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const ENCODER_BMP    As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_JPG    As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_GIF    As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_TIF    As String = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_PNG    As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 And Win64 Then
    'for clip board
    Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" ( _
            ByVal hwnd As LongPtr) As Long
            
    Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" ( _
            ByVal wFormat As Long) As Long
            
    Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
    'GDI+
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus.dll" ( _
            ByRef token As Long, _
            ByRef inputBuf As GdiplusStartupInput, _
            ByVal outputBuf As Long) As Long
            
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus.dll" ( _
            ByVal token As Long)
            
    Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" ( _
            ByVal hbm As LongPtr, _
            ByVal hpal As LongPtr, _
            bitmap As LongPtr) As Long
            
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus.dll" ( _
            ByVal image As LongPtr) As Long
            
    Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus.dll" ( _
            ByVal image As LongPtr, _
            ByVal fileName As LongPtr, _
            ByRef clsidEncoder As GUID, _
            ByVal encoderParams As Any) As Long
            
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" ( _
            ByVal lpszCLSID As LongPtr, _
            ByRef pCLSID As GUID) As Long
            
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus.dll" ( _
            ByVal image As LongPtr, _
            Height As Long) As Long
            
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus.dll" ( _
            ByVal image As LongPtr, _
            Width As Long) As Long
            
    Private Type EncoderParameter
        GUID As GUID
        NumberOfValues As Long
        TypeAPI As Long
        value As LongPtr
    End Type
    Private Type EncoderParameters
        Count As Long
        Parameter(0 To 15) As EncoderParameter
    End Type
#Else
    Private Declare Function OpenClipboard Lib "user32.dll" ( _
            ByVal hwnd As Long) As Long
    Private Declare Function GetClipboardData Lib "user32.dll" ( _
            ByVal wFormat As Long) As Long
    Private Declare Function CloseClipboard Lib "user32.dll" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function GdiplusStartup Lib "gdiplus.dll" ( _
            ByRef token As Long, _
            ByRef inputBuf As GdiplusStartupInput, _
            ByVal outputBuf As Long) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" ( _
            ByVal token As Long)
    Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" ( _
            ByVal hbm As Long, _
            ByVal hpal As Long, _
            bitmap As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus.dll" ( _
            ByVal image As Long) As Long
    Private Declare Function GdipSaveImageToFile Lib "gdiplus.dll" ( _
            ByVal image As Long, _
            ByVal fileName As Long, _
            ByRef clsidEncoder As GUID, _
            ByVal encoderParams As Any) As Long
    Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
            ByVal lpszCLSID As Long, _
            ByRef pCLSID As GUID) As Long
    Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" ( _
            ByVal image As Long, _
            Height As Long) As Long
    Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" ( _
            ByVal image As Long, _
            Width As Long) As Long
    Private Type EncoderParameter
        GUID As GUID
        NumberOfValues As Long
        TypeAPI As Long
        value As Long
    End Type
    Private Type EncoderParameters
        Count As Long
        Parameter(0 To 15) As EncoderParameter
    End Type
#End If

Public Enum GDIPlusStatusConstants
    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

Private m_GDIplusToken
Private Const CF_BITMAP As Long = 2 ' data type of clipboard


Public Sub SetShapeLineRound()
    If modPicture.IsSelectionShape Then
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText2
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.400000006
            .Transparency = 0
            .weight = 3#
        End With
    End If
End Sub

Public Function IsSelectionShape() As Boolean
    On Error GoTo ErrHandl
      Dim shp As ShapeRange
      Set shp = Selection.ShapeRange
      IsSelectionShape = True
    Exit Function
    
ErrHandl:
      Err.Clear
      IsSelectionShape = False
End Function


Public Function SaveCipPictureByName(ByVal filePath As String, ByVal lngQ As Long) As Long
    'lngQ quality of jpeg
    
    Dim strExt As String
    strExt = getFileExtension(filePath)
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strReceive As String
    Dim objGdipBmp
    Dim hBmp As OLE_HANDLE
    If GDIplus_Initialize() = False Then 'GDI+ initialize
        MsgBox "we can not initialize GDI+ ", vbCritical: Exit Function
    End If
    hBmp = pvGetHBitmapFromClipboard() ' get clipboard image handle
    If hBmp = 0 Then GoSub FINGDIP
    'create bitmap object from Bitmap handle. objGdipBmp is created image.
    If GdipCreateBitmapFromHBITMAP(hBmp, 0&, objGdipBmp) = 0 Then
        Call GdipGetImageWidth(objGdipBmp, lngWidth) 'get width
        Call GdipGetImageHeight(objGdipBmp, lngHeight) 'get height
        If Not (lngWidth <= 3200 And lngHeight <= 3200) Then
            MsgBox "Too Big Image Size.", vbExclamation + vbOKOnly
            GoSub TERMINATE
        End If
        SaveCipPictureByName = SaveImageToFile(objGdipBmp, filePath, strExt, lngQ)
    End If
TERMINATE: 'dispose image
    Call GdipDisposeImage(objGdipBmp)
FINGDIP: 'End GDI+
    Call Gdiplus_Shutdown
End Function


'' private --------------
Private Function GDIplus_Initialize() As Boolean
    Dim lngStatus As Long
    Dim uGdiStartupInput As GdiplusStartupInput
    If m_GDIplusToken <> 0 Then Call Gdiplus_Shutdown
    With uGdiStartupInput
        .GdiplusVersion = 1
        .DebugEventCallback = 0
        .SuppressBackgroundThread = 0
        .SuppressExternalCodecs = 0
    End With
    GDIplus_Initialize = CBool(GdiplusStartup(m_GDIplusToken, uGdiStartupInput, 0&) = 0)
End Function
'-------------------------------------------------------------------------------------------------------

Private Function Gdiplus_Shutdown() 'end GDI+
    Dim retCode As Long
    retCode = 0
    If m_GDIplusToken <> 0 Then
        GdiplusShutdown (m_GDIplusToken): m_GDIplusToken = 0
    End If
End Function
'-------------------------------------------------------------------------------------------------------

'GDI+ to file
Private Function SaveImageToFile(ByVal objBmp, ByVal sFilename As String, _
                                ByVal sFormat As String, ByVal nQuarity As Long) As Long
    Dim strEncoder As String
    Dim uEncoderParams As EncoderParameters
    Select Case UCase$(sFormat)
        Case "JPG": strEncoder = ENCODER_JPG
        Case "GIF": strEncoder = ENCODER_GIF
        Case "TIF": strEncoder = ENCODER_TIF
        Case "PNG": strEncoder = ENCODER_PNG
        Case Else: strEncoder = ENCODER_BMP
    End Select
    
    If UCase$(sFormat) = "JPG" Then
        nQuarity = Abs(nQuarity)
        If nQuarity = 0 Or 100 < nQuarity Then nQuarity = 100
        uEncoderParams.Count = 1
        With uEncoderParams.Parameter(0)
            .GUID = pvToCLSID(QUALITY_PARAMS)
            .TypeAPI = 4
            .NumberOfValues = 1
            .value = VarPtr(nQuarity)
        End With
    End If
    
    Dim retCode As Long
    If UCase$(sFormat) = "JPG" Then
        retCode = GdipSaveImageToFile(objBmp, StrPtr(sFilename), _
                pvToCLSID(strEncoder), VarPtr(uEncoderParams))
    Else
        retCode = GdipSaveImageToFile(objBmp, StrPtr(sFilename), _
                pvToCLSID(strEncoder), ByVal 0&)
    End If
    SaveImageToFile = retCode
End Function
'-------------------------------------------------------------------------------------------------------

' get a bitmap in the clipboard
Private Function pvGetHBitmapFromClipboard() As OLE_HANDLE
    If OpenClipboard(0&) <> 0 Then
        pvGetHBitmapFromClipboard = GetClipboardData(CF_BITMAP)
    Else
        pvGetHBitmapFromClipboard = 0
    End If
    Call CloseClipboard
End Function


'for jpeg, get classid form string
Private Function pvToCLSID(ByVal s As String) As GUID
    Call CLSIDFromString(StrPtr(s), pvToCLSID)
End Function

Public Sub ShapeGrouping(Optional ByVal nonv As Long = 1)
    On Error GoTo ShapeGrouping_Error

    On Error GoTo EH
    Selection.ShapeRange.Group.Select
    With Selection
        .Placement = xlMove
        .PrintObject = True
    End With

    Exit Sub
EH:
    MsgBox "Err"

    On Error GoTo 0
    GoTo ShapeGrouping_Normal_exit

ShapeGrouping_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShapeGrouping of Module modPicture"

ShapeGrouping_Normal_exit:
End Sub


Public Function LoadPicFile(ByRef fname As String, ByRef rng1 As Range) As Boolean
    Dim myShape As Shape
    
    On Error GoTo LoadPicFile_Error
    
    
    
    'insert image
    Set myShape = ActiveSheet.shapes.AddPicture( _
          fileName:=fname, _
          LinkToFile:=False, _
          SaveWithDocument:=True, _
          Left:=rng1.Cells(1, 1).Left, _
          Top:=rng1.Cells(1, 1).Top, _
          Width:=0, _
          Height:=0)
          
    '--(2) same hight and width for the inserted image.
    With myShape
        .Left = .Left + 5  ' shift some pixel
        .Top = .Top + 5 ' shift some pixel
        .ScaleHeight 1, msoTrue
        .ScaleWidth 1, msoTrue
        .title = fname
    End With
    
    myShape.LockAspectRatio = msoTrue
    myShape.Height = myShape.Height
     
    'add empty lines under the image
    Dim i As Long
    Dim rowMax As Long
    Dim rowsempty As Long
    rowsempty = 0
    rowMax = Range(myShape.TopLeftCell.Offset(1, 0), myShape.BottomRightCell.Offset(1, 0)).EntireRow.Count
    For i = 2 To rowMax
        If WorksheetFunction.CountA(myShape.TopLeftCell.Offset(1, 0).Resize(i).EntireRow) > 0 Then
            rowsempty = i
            Exit For
        Else
        
        End If
    Next i
    If rowsempty > 0 Then
        myShape.TopLeftCell.Resize(rowMax - i, 1).Offset(1, 0).EntireRow.Insert xlShiftDown
        
    End If
    
            
    Set myShape = Nothing

    On Error GoTo 0
    GoTo LoadPicFile_Normal_exit

LoadPicFile_Error:
    ''MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadPicFile of Module modPicture"
    LoadPicFile = False
    Exit Function

LoadPicFile_Normal_exit:
    LoadPicFile = True
    
End Function









