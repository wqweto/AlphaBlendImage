VERSION 5.00
Begin VB.UserControl AlphaBlendImage 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   DrawStyle       =   2  'Dot
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
End
Attribute VB_Name = "AlphaBlendImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' AlphaBlendImage (c) 2019 by wqweto@gmail.com
'
' Poor Man's Transparent Image Control
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "AlphaBlendImage"

'=========================================================================
' Public events
'=========================================================================

Event Click()
Event OwnerDraw(ByVal hGraphics As Long, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, ByVal hPicture As Long)
Event DblClick()
Event ContextMenu()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'=========================================================================
' API
'=========================================================================

'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppARGB          As Long = &H26200A
Private Const PixelFormat32bppPARGB         As Long = &HE200B
'--- for GdipDrawImageXxx
Private Const UnitPixel                     As Long = 2
'--- DIB Section constants
Private Const DIB_RGB_COLORS                As Long = 0 '  color table in RGBs
'--- for GdipSetInterpolationMode
Private Const InterpolationModeHighQualityBicubic As Long = 7
'--- for Gdip*WorldTransform
Private Const MatrixOrderAppend             As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, pIconInfo As ICONINFO) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Function ApiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateIconIndirect Lib "user32" (pIconInfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (ByRef pInit As Any, ByVal cbInit As Long) As IUnknown
'--- gdi+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, ByVal Scan0 As Long, hBitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (hImgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, clrMatrix As Any, grayMatrix As Any, ByVal lFlags As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal hImgAttr As Long, ByVal lAdjustType As Long, ByVal fAdjustEnabled As Long, ByVal clrLow As Long, ByVal clrHigh As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImgAttr As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lPixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, hBtmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As Long, hBitmap As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, nWidth As Single, nHeight As Single) As Long   '
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal sFileName As Long, mImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal pStream As IUnknown, ByRef mImage As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal nDx As Single, ByVal nDy As Single, ByVal lOrder As Long) As Long
Private Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal nSx As Single, ByVal nSy As Single, ByVal lOrder As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal hGraphics As Long, ByVal nRotation As Single, ByVal lOrder As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lMode As Long) As Long

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

Private Type PICTDESC
    lSize               As Long
    lType               As Long
    hBmp                As Long
    hPal                As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_OPACITY           As Single = 1
Private Const DEF_ROTATION          As Single = 0
Private Const DEF_ZOOM              As Single = 1
Private Const DEF_MASKCOLOR         As Long = vbMagenta
Private Const DEF_AUTOREDRAW        As Boolean = False
Private Const DEF_STRETCH           As Boolean = False

Private m_oPicture              As StdPicture
Private m_clrMask               As OLE_COLOR
Private m_bAutoRedraw           As Boolean
Private m_sngOpacity            As Single
Private m_sngRotation           As Single
Private m_sngZoom               As Single
Private m_bStretch              As Boolean
'--- run-time
Private m_eContainerScaleMode   As ScaleModeConstants
Private m_bShown                As Boolean
Private m_hAttributes           As Long
Private m_hBitmap               As Long
Private m_hPictureBitmap        As Long
Private m_hPictureAttributes    As Long
Private m_hRedrawDib            As Long
Private m_nDownButton           As Integer
Private m_nDownShift            As Integer
Private m_sngDownX              As Single
Private m_sngDownY              As Single

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get Picture() As StdPicture
    Set Picture = m_oPicture
End Property

Property Set Picture(oValue As StdPicture)
    If Not m_oPicture Is oValue Then
        Set m_oPicture = oValue
        pvPreparePicture m_oPicture, m_clrMask, m_hPictureBitmap, m_hPictureAttributes
        If Not m_bStretch And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender m_hPictureBitmap, Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get MaskColor() As OLE_COLOR
    MaskColor = m_clrMask
End Property

Property Let MaskColor(ByVal clrValue As OLE_COLOR)
    If m_clrMask <> clrValue Then
        m_clrMask = clrValue
        pvPreparePicture m_oPicture, m_clrMask, m_hPictureBitmap, m_hPictureAttributes
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get AutoRedraw() As Boolean
    AutoRedraw = m_bAutoRedraw
End Property

Property Let AutoRedraw(ByVal bValue As Boolean)
    If m_bAutoRedraw <> bValue Then
        m_bAutoRedraw = bValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Opacity() As Single
    Opacity = m_sngOpacity
End Property

Property Let Opacity(ByVal sngValue As Single)
    If m_sngOpacity <> sngValue Then
        m_sngOpacity = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Rotation() As Single
    Rotation = m_sngRotation
End Property

Property Let Rotation(ByVal sngValue As Single)
    If m_sngRotation <> sngValue Then
        m_sngRotation = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Zoom() As Single
    Zoom = m_sngZoom
End Property

Property Let Zoom(ByVal sngValue As Single)
    If m_sngZoom <> sngValue Then
        m_sngZoom = sngValue
        If Not m_bStretch And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender m_hPictureBitmap, Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Stretch() As Boolean
    Stretch = m_bStretch
End Property

Property Let Stretch(ByVal bValue As Boolean)
    If m_bStretch <> bValue Then
        m_bStretch = bValue
        If Not m_bStretch And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender m_hPictureBitmap, Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Refresh()
    Const FUNC_NAME     As String = "Refresh"
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(0)
        If hMemDC = 0 Then
            GoTo QH
        End If
        If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
            GoTo QH
        End If
        hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        pvPaintControl hMemDC
    End If
    UserControl.Refresh
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Repaint()
    Const FUNC_NAME     As String = "Repaint"
    
    On Error GoTo EH
    If m_bShown Then
        pvPrepareBitmap m_hBitmap
        pvPrepareAttribs m_sngOpacity, m_hAttributes
        Refresh
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Function GdipLoadPicture(sFileName As String) As StdPicture
    Const FUNC_NAME     As String = "GdipLoadPicture"
    Dim hBitmap         As Long
    
    On Error GoTo EH
    If GdipLoadImageFromFile(StrPtr(sFileName), hBitmap) <> 0 Then
        GoTo QH
    End If
    Set GdipLoadPicture = pvLoadPicture(hBitmap)
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GdipLoadPictureArray(baBuffer() As Byte) As StdPicture
    Const FUNC_NAME     As String = "GdipLoadPictureArray"
    Dim pStream         As IUnknown
    Dim hBitmap         As Long
    
    On Error GoTo EH
    Set pStream = SHCreateMemStream(baBuffer(LBound(baBuffer)), UBound(baBuffer) - LBound(baBuffer) + 1)
    If pStream Is Nothing Then
        GoTo QH
    End If
    If GdipLoadImageFromStream(pStream, hBitmap) <> 0 Then
        GoTo QH
    End If
    Set GdipLoadPictureArray = pvLoadPicture(hBitmap)
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'= private ===============================================================

Private Function pvLoadPicture(hBitmap As Long) As StdPicture
    Const FUNC_NAME     As String = "pvLoadPicture"
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    Dim hMemDC          As Long
    Dim hDib            As Long
    Dim hPrevDib        As Long
    Dim hGraphics       As Long
    Dim uInfo           As ICONINFO
    Dim hIcon           As Long
    Dim uDesc           As PICTDESC
    Dim aGUID(0 To 3)   As Long
    
    On Error GoTo EH
    If GdipGetImageDimension(hBitmap, sngWidth, sngHeight) <> 0 Then
        GoTo QH
    End If
    hMemDC = CreateCompatibleDC(0)
    If hMemDC = 0 Then
        GoTo QH
    End If
    If Not pvCreateDib(hMemDC, sngWidth, sngHeight, hDib) Then
        GoTo QH
    End If
    hPrevDib = SelectObject(hMemDC, hDib)
    If GdipCreateFromHDC(hMemDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipDrawImageRectRect(hGraphics, hBitmap, 0, 0, sngWidth, sngHeight, 0, 0, sngWidth, sngHeight) <> 0 Then
        GoTo QH
    End If
    Call SelectObject(hMemDC, hPrevDib)
    hPrevDib = 0
    With uInfo
        .fIcon = 1
        .hbmColor = hDib
        .hbmMask = CreateBitmap(sngWidth, sngHeight, 1, 1, ByVal 0)
    End With
    hIcon = CreateIconIndirect(uInfo)
    With uDesc
        .lSize = Len(uDesc)
        .lType = vbPicTypeIcon
        .hBmp = hIcon
    End With
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    If OleCreatePictureIndirect(uDesc, aGUID(0), 1, pvLoadPicture) <> 0 Then
        GoTo QH
    End If
    hIcon = 0
QH:
    If hBitmap <> 0 Then
        Call GdipDisposeImage(hBitmap)
        hBitmap = 0
    End If
    If hMemDC <> 0 Then
        If hPrevDib <> 0 Then
            Call SelectObject(hMemDC, hPrevDib)
        End If
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    If hDib <> 0 Then
        Call DeleteObject(hDib)
        hDib = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    If hIcon <> 0 Then
        Call DestroyIcon(hIcon)
        hIcon = 0
    End If
    If uInfo.hbmMask <> 0 Then
        Call DeleteObject(uInfo.hbmMask)
        uInfo.hbmMask = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvRefresh()
    m_bShown = False
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    UserControl.Refresh
End Sub

Private Function pvPaintControl(ByVal hDC As Long) As Boolean
    Const FUNC_NAME     As String = "pvPaintControl"
    Dim hGraphics       As Long
    
    On Error GoTo EH
    If Not m_bShown Then
        m_bShown = True
        pvPrepareBitmap m_hBitmap
        pvPrepareAttribs m_sngOpacity, m_hAttributes
    End If
    If m_hBitmap <> 0 Then
        If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
            GoTo QH
        End If
        If GdipDrawImageRectRect(hGraphics, m_hBitmap, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, , m_hAttributes) <> 0 Then
            GoTo QH
        End If
        '--- success
        pvPaintControl = True
    End If
QH:
    On Error Resume Next
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareBitmap(hBitmap As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareBitmap"
    Dim hGraphics       As Long
    Dim hNewBitmap      As Long
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim sngPicWidth     As Single
    Dim sngPicHeight    As Single
    Dim sngZoom         As Single
    
    On Error GoTo EH
    If GdipCreateBitmapFromScan0(ScaleWidth, ScaleHeight, ScaleWidth * 4, PixelFormat32bppPARGB, 0, hNewBitmap) <> 0 Then
        GoTo QH
    End If
    If GdipGetImageGraphicsContext(hNewBitmap, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipSetInterpolationMode(hGraphics, InterpolationModeHighQualityBicubic) <> 0 Then
        GoTo QH
    End If
    lWidth = ScaleWidth
    lHeight = ScaleHeight
    RaiseEvent OwnerDraw(hGraphics, lLeft, lTop, lWidth, lHeight, m_hPictureBitmap)
    If lWidth > 0 And lHeight > 0 Then
        If m_hPictureBitmap <> 0 Then
            If GdipGetImageDimension(m_hPictureBitmap, sngPicWidth, sngPicHeight) <> 0 Then
                GoTo QH
            End If
            sngZoom = IIf(Not m_bStretch, m_sngZoom, 1)
            If GdipRotateWorldTransform(hGraphics, m_sngRotation, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
            If GdipTranslateWorldTransform(hGraphics, lWidth / 2 / sngZoom, lHeight / 2 / sngZoom, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
            If GdipScaleWorldTransform(hGraphics, sngZoom, sngZoom, MatrixOrderAppend) <> 0 Then
                GoTo QH
            End If
            If m_bStretch Then
                lLeft = lLeft - lWidth / 2
                lTop = lTop - lHeight / 2
                If GdipDrawImageRectRect(hGraphics, m_hPictureBitmap, lLeft, lTop, lWidth, lHeight, 0, 0, sngPicWidth, sngPicHeight, , m_hPictureAttributes) <> 0 Then
                    GoTo QH
                End If
            Else
                lLeft = lLeft - sngZoom * sngPicWidth / 2
                lTop = lTop - sngZoom * sngPicHeight / 2
                If GdipDrawImageRectRect(hGraphics, m_hPictureBitmap, lLeft + (lWidth - sngPicWidth) / 2, lTop + (lHeight - sngPicHeight) / 2, sngPicWidth, sngPicHeight, 0, 0, sngPicWidth, sngPicHeight, , m_hPictureAttributes) <> 0 Then
                    GoTo QH
                End If
            End If
        ElseIf Not Ambient.UserMode Then
            Call GdipDisposeImage(hNewBitmap)
            hNewBitmap = 0
        End If
    End If
    '--- commit
    If hNewBitmap <> hBitmap Then
        If hBitmap <> 0 Then
            Call GdipDisposeImage(hBitmap)
            hBitmap = 0
        End If
        hBitmap = hNewBitmap
    End If
    hNewBitmap = 0
    '-- success
    pvPrepareBitmap = True
QH:
    On Error Resume Next
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareAttribs(ByVal sngAlpha As Single, hAttributes As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrepareAttribs"
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim hNewAttributes  As Long
    
    On Error GoTo EH
    If GdipCreateImageAttributes(hNewAttributes) <> 0 Then
        GoTo QH
    End If
    clrMatrix(0, 0) = 1
    clrMatrix(1, 1) = 1
    clrMatrix(2, 2) = 1
    clrMatrix(3, 3) = sngAlpha
    clrMatrix(4, 4) = 1
    If GdipSetImageAttributesColorMatrix(hNewAttributes, 0, 1, clrMatrix(0, 0), clrMatrix(0, 0), 0) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hAttributes)
        hAttributes = 0
    End If
    hAttributes = hNewAttributes
    hNewAttributes = 0
    '--- success
    pvPrepareAttribs = True
QH:
    On Error Resume Next
    If hNewAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hNewAttributes)
        hNewAttributes = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPreparePicture(oPicture As StdPicture, ByVal clrMask As OLE_COLOR, hPictureBitmap As Long, hPictureAttributes As Long) As Boolean
    Const FUNC_NAME     As String = "pvPreparePicture"
    Dim hTempBitmap     As Long
    Dim hNewBitmap      As Long
    Dim hNewAttributes  As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim uHdr            As BITMAPINFOHEADER
    Dim hMemDC          As Long
    Dim uInfo           As ICONINFO
    Dim baColorBits()   As Byte
    Dim bHasAlpha       As Boolean
    Dim hDib            As Long
    Dim lpBits          As Long
    Dim hPrevDib        As Long
    Dim lIdx            As Long
    Dim pPic            As IPicture
    
    On Error GoTo EH
    If Not oPicture Is Nothing Then
        If oPicture.Handle <> 0 Then
            Select Case oPicture.Type
            Case vbPicTypeBitmap
                If GdipCreateBitmapFromHBITMAP(oPicture.Handle, 0, hNewBitmap) <> 0 Then
                    GoTo QH
                End If
                If clrMask <> -1 Then
                    If GdipCreateImageAttributes(hNewAttributes) <> 0 Then
                        GoTo QH
                    End If
                    If GdipSetImageAttributesColorKeys(hNewAttributes, 0, 1, TranslateColor(clrMask), TranslateColor(clrMask)) <> 0 Then
                        GoTo QH
                    End If
                End If
            Case Else
                lWidth = HM2Pix(oPicture.Width)
                lHeight = HM2Pix(oPicture.Height)
                hMemDC = CreateCompatibleDC(0)
                If hMemDC = 0 Then
                    GoTo QH
                End If
                With uHdr
                    .biSize = Len(uHdr)
                    .biPlanes = 1
                    .biBitCount = 32
                    .biWidth = lWidth
                    .biHeight = -lHeight
                    .biSizeImage = (4 * lWidth) * lHeight
                End With
                If oPicture.Type = vbPicTypeIcon Then
                    If GetIconInfo(oPicture.Handle, uInfo) = 0 Then
                        GoTo QH
                    End If
                    ReDim baColorBits(0 To uHdr.biSizeImage - 1) As Byte
                    If GetDIBits(hMemDC, uInfo.hbmColor, 0, lHeight, baColorBits(0), uHdr, DIB_RGB_COLORS) = 0 Then
                        GoTo QH
                    End If
                    For lIdx = 3 To UBound(baColorBits) Step 4
                        If baColorBits(lIdx) <> 0 Then
                            bHasAlpha = True
                            Exit For
                        End If
                    Next
                    If Not bHasAlpha Then
                        '--- note: GdipCreateBitmapFromHICON working ok for old-style (single-bit) transparent icons only
                        If GdipCreateBitmapFromHICON(oPicture.Handle, hNewBitmap) <> 0 Then
                            GoTo QH
                        End If
                    Else
                        If GdipCreateBitmapFromScan0(lWidth, lHeight, 4 * lWidth, PixelFormat32bppPARGB, VarPtr(baColorBits(0)), hTempBitmap) <> 0 Then
                            GoTo QH
                        End If
                        If GdipCloneBitmapAreaI(0, 0, lWidth, lHeight, PixelFormat32bppARGB, hTempBitmap, hNewBitmap) <> 0 Then
                            GoTo QH
                        End If
                    End If
                Else
                    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
                    If hDib = 0 Then
                        GoTo QH
                    End If
                    hPrevDib = SelectObject(hMemDC, hDib)
                    Set pPic = oPicture
                    pPic.Render hMemDC, 0, 0, lWidth, lHeight, 0, oPicture.Height, oPicture.Width, -oPicture.Height, ByVal 0
                    If GdipCreateBitmapFromScan0(lWidth, lHeight, 4 * lWidth, PixelFormat32bppPARGB, lpBits, hTempBitmap) <> 0 Then
                        GoTo QH
                    End If
                    If GdipCloneBitmapAreaI(0, 0, lWidth, lHeight, PixelFormat32bppARGB, hTempBitmap, hNewBitmap) <> 0 Then
                        GoTo QH
                    End If
                End If
            End Select
        End If
    End If
    '--- commit
    If hPictureBitmap <> 0 Then
        Call GdipDisposeImage(hPictureBitmap)
        hPictureBitmap = 0
    End If
    hPictureBitmap = hNewBitmap
    hNewBitmap = 0
    If hPictureAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hPictureAttributes)
        hPictureAttributes = 0
    End If
    hPictureAttributes = hNewAttributes
    hNewAttributes = 0
    '--- success
    pvPreparePicture = True
QH:
    On Error Resume Next
    If hNewBitmap <> 0 Then
        Call GdipDisposeImage(hNewBitmap)
        hNewBitmap = 0
    End If
    If hNewAttributes <> 0 Then
        Call GdipDisposeImageAttributes(hNewAttributes)
        hNewAttributes = 0
    End If
    If hTempBitmap <> 0 Then
        Call GdipDisposeImage(hTempBitmap)
        hTempBitmap = 0
    End If
    If hPrevDib <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        hPrevDib = 0
    End If
    If hDib <> 0 Then
        Call DeleteObject(hDib)
        hDib = 0
    End If
    If uInfo.hbmColor <> 0 Then
        Call DeleteObject(uInfo.hbmColor)
        uInfo.hbmColor = 0
    End If
    If uInfo.hbmMask <> 0 Then
        Call DeleteObject(uInfo.hbmMask)
        uInfo.hbmMask = 0
    End If
    If hMemDC <> 0 Then
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvSizeExtender(ByVal hBitmap As Long, oExt As VBControlExtender)
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    
    If hBitmap = 0 Then
        GoTo QH
    End If
    If GdipGetImageDimension(m_hPictureBitmap, sngWidth, sngHeight) <> 0 Then
        GoTo QH
    End If
    oExt.Width = ScaleX(sngWidth * m_sngZoom, vbPixels, m_eContainerScaleMode)
    oExt.Height = ScaleY(sngHeight * m_sngZoom, vbPixels, m_eContainerScaleMode)
QH:
End Sub

Private Sub pvHandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_nDownButton = Button
    m_nDownShift = Shift
    m_sngDownX = X
    m_sngDownY = Y
End Sub

'= common ================================================================

Private Function pvCreateDib(ByVal hMemDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, hDib As Long) As Boolean
    Const FUNC_NAME     As String = "pvCreateDib"
    Dim uHdr            As BITMAPINFOHEADER
    Dim lpBits          As Long
    
    On Error GoTo EH
    With uHdr
        .biSize = Len(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lWidth
        .biHeight = -lHeight
        .biSizeImage = 4 * lWidth * lHeight
    End With
    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
    If hDib = 0 Then
        GoTo QH
    End If
    '--- success
    pvCreateDib = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function TranslateColor(ByVal clrValue As OLE_COLOR, Optional ByVal Alpha As Single = 1) As Long
    Dim uQuad           As UcsRgbQuad
    Dim lTemp           As Long
    
    Call OleTranslateColor(clrValue, 0, VarPtr(uQuad))
    lTemp = uQuad.R
    uQuad.R = uQuad.B
    uQuad.B = lTemp
    lTemp = Alpha * &HFF
    If lTemp > 255 Then
        uQuad.A = 255
    ElseIf lTemp < 0 Then
        uQuad.A = 0
    Else
        uQuad.A = lTemp
    End If
    Call CopyMemory(TranslateColor, uQuad, 4)
End Function

Private Function HM2Pix(ByVal Value As Single) As Long
   HM2Pix = Int(Value * 1440 / 2540 / Screen.TwipsPerPixelX + 0.5!)
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Function C_ArrayByte(Value As Variant) As Byte()
    On Error GoTo QH
    If Not IsEmpty(Value) Then
        C_ArrayByte = Value
    End If
QH:
End Function

Private Function PictureFromBuffer(baBuffer() As Byte) As StdPicture
    Dim sFile           As String
    Dim nFile           As Integer
    
    On Error GoTo QH
    If Peek(ArrPtr(baBuffer)) = 0 Then
        GoTo QH
    End If
    If UBound(baBuffer) < 0 Then
        GoTo QH
    End If
    sFile = Environ$("TMP") & "\$~" & STR_MODULE_NAME & ".tmp"
    Call ApiDeleteFile(sFile)
    nFile = FreeFile
    Open sFile For Binary Access Write Shared As nFile
    Put nFile, , baBuffer
    Close nFile
    Set PictureFromBuffer = LoadPicture(sFile)
    Call ApiDeleteFile(sFile)
QH:
End Function

Private Function PictureToBuffer(oPic As StdPicture) As Byte()
    Dim sFile           As String
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    
    On Error GoTo QH
    If oPic Is Nothing Then
        GoTo QH
    End If
    sFile = Environ$("TMP") & "\$~" & STR_MODULE_NAME & ".tmp"
    Call ApiDeleteFile(sFile)
    SavePicture oPic, sFile
    baBuffer = vbNullString
    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    Close nFile
    Call ApiDeleteFile(sFile)
    PictureToBuffer = baBuffer
QH:
End Function

Private Function ToScaleMode(sScaleUnits As String) As ScaleModeConstants
    Select Case sScaleUnits
    Case "Twip"
        ToScaleMode = vbTwips
    Case "Point"
        ToScaleMode = vbPoints
    Case "Pixel"
        ToScaleMode = vbPixels
    Case "Character"
        ToScaleMode = vbCharacters
    Case "Centimeter"
        ToScaleMode = vbCentimeters
    Case "Millimeter"
        ToScaleMode = vbMillimeters
    Case "Inch"
        ToScaleMode = vbInches
    Case Else
        ToScaleMode = vbTwips
    End Select
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    pvHandleMouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseUp"
    
    On Error GoTo EH
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    If Button = -1 Then
        GoTo QH
    End If
    If Button <> 0 And X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight Then
        If (m_nDownButton And Button And vbLeftButton) <> 0 Then
            RaiseEvent Click
        ElseIf (m_nDownButton And Button And vbRightButton) <> 0 Then
            RaiseEvent ContextMenu
        End If
    End If
    m_nDownButton = 0
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_DblClick()
    pvHandleMouseDown vbLeftButton, m_nDownShift, m_sngDownX, m_sngDownY
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Resize()
    pvRefresh
End Sub

Private Sub UserControl_Hide()
    m_bShown = False
End Sub

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"
    Const AC_SRC_ALPHA  As Long = 1
    Const Opacity       As Long = &HFF
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(hDC)
        If hMemDC = 0 Then
            GoTo DefPaint
        End If
        If m_hRedrawDib = 0 Then
            If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
                GoTo DefPaint
            End If
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
            If Not pvPaintControl(hMemDC) Then
                GoTo DefPaint
            End If
        Else
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        End If
        If AlphaBlend(hDC, 0, 0, ScaleWidth, ScaleHeight, hMemDC, 0, 0, ScaleWidth, ScaleHeight, AC_SRC_ALPHA * &H1000000 + Opacity * &H10000) = 0 Then
            GoTo DefPaint
        End If
    Else
        If Not pvPaintControl(hDC) Then
            GoTo DefPaint
        End If
    End If
    If False Then
DefPaint:
        If m_hRedrawDib <> 0 Then
            '--- note: before deleting DIB try de-selecting from dc
            Call SelectObject(hMemDC, hPrevDib)
            Call DeleteObject(m_hRedrawDib)
            m_hRedrawDib = 0
        End If
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), vbBlack, B
    End If
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    AutoRedraw = DEF_AUTOREDRAW
    Opacity = DEF_OPACITY
    Rotation = DEF_ROTATION
    Zoom = DEF_ZOOM
    MaskColor = DEF_MASKCOLOR
    Stretch = DEF_STRETCH
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    With PropBag
        AutoRedraw = .ReadProperty("AutoRedraw", DEF_AUTOREDRAW)
        Opacity = .ReadProperty("Opacity", DEF_OPACITY)
        Rotation = .ReadProperty("Rotation", DEF_ROTATION)
        Zoom = .ReadProperty("Zoom", DEF_ZOOM)
        MaskColor = .ReadProperty("MaskColor", DEF_MASKCOLOR)
        Stretch = .ReadProperty("Stretch", DEF_STRETCH)
        Set Picture = PictureFromBuffer(C_ArrayByte(.ReadProperty("Picture", vbNullString)))
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "AutoRedraw", AutoRedraw, DEF_AUTOREDRAW
        .WriteProperty "Opacity", Opacity, DEF_OPACITY
        .WriteProperty "Rotation", Rotation, DEF_ROTATION
        .WriteProperty "Zoom", Zoom, DEF_ZOOM
        .WriteProperty "MaskColor", MaskColor, DEF_MASKCOLOR
        .WriteProperty "Stretch", Stretch, DEF_STRETCH
        .WriteProperty "Picture", PictureToBuffer(m_oPicture), vbNullString
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "ScaleUnits" Then
        m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Initialize()
    Dim aInput(0 To 3)  As Long
    
    If GetModuleHandle("gdiplus") = 0 Then
        aInput(0) = 1
        Call GdiplusStartup(0, aInput(0))
    End If
    m_eContainerScaleMode = vbTwips
End Sub

Private Sub UserControl_Terminate()
    If m_hAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hAttributes)
        m_hAttributes = 0
    End If
    If m_hBitmap <> 0 Then
        Call GdipDisposeImage(m_hBitmap)
        m_hBitmap = 0
    End If
    If m_hPictureBitmap <> 0 Then
        Call GdipDisposeImage(m_hPictureBitmap)
        m_hPictureBitmap = 0
    End If
    If m_hPictureAttributes <> 0 Then
        Call GdipDisposeImageAttributes(m_hPictureAttributes)
        m_hPictureAttributes = 0
    End If
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
End Sub
