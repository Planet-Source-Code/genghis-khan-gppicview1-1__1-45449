VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' From cDIBSection.cls
' Copyright � 1999 Steve McMahon

' ======================================================================================
' Constants
' ======================================================================================

' constants for the biCompression field
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const BI_BITFIELDS = 3&
Private Const BI_JPEG = 4&
Private Const BI_PNG = 5&

Private Const JPEG_Y = 1               ' /* Y only component of YCbCr */
Private Const JPEG_YCbCr = 2          ' /* YCbCr as define by CCIR 601 */
Private Const JPEG_RGB = 3           ' /* 3 component RGB */

' DIB color table identifiers
Private Const DIB_RGB_COLORS              As Long = 0    ' BITMAPINFO������һ����ɫ�������б�����32λ��ɫ��RGBQUAD��
Private Const DIB_PAL_COLORS              As Long = 1    ' BITMAPINFO������һ��16λ��ɫ������������

Private Const STRETCH_ANDSCANS = 1
Private Const STRETCH_ORSCANS = 2
Private Const STRETCH_DELETESCANS = 3
Private Const STRETCH_HALFTONE = 4

' ��������Ԫ��դ��������
Private Const BLACKNESS = &H42
Private Const DSTINVERT = &H550009
Private Const MERGECOPY = &HC000CA
Private Const MERGEPAINT = &HBB0226
Private Const NOTSRCCOPY = &H330008
Private Const NOTSRCERASE = &H1100A6
Private Const PATCOPY = &HF00021
Private Const PATINVERT = &H5A0049
Private Const PATPAINT = &HFB0A09
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const WHITENESS = &HFF0062

' ======================================================================================
' Enums
' ======================================================================================
' ======================================================================================
' Enums
' ======================================================================================

'ָ����װ��ͼ������,LoadImage,CopyImage
Private Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

' LoadImage
Private Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/*
    LR_COPYDELETEORG = &H8                       '/*
    LR_CREATEDIBSECTION = &H2000                 '/*
    LR_DEFAULTCOLOR = &H0                        '/*
    LR_DEFAULTSIZE = &H40                        '/*
    LR_LOADFROMFILE = &H10                       '/*
    LR_LOADMAP3DCOLORS = &H1000                  '/*
    LR_LOADTRANSPARENT = &H20                    '/*
    LR_MONOCHROME = &H1                          '/*
    LR_SHARED = &H8000                           '/*
    LR_COPYFROMRESOURCE = &H4000                 '/*
    LR_VGACOLOR = &H80                           '/*
End Enum

Public Enum UseAPIPaintPicture
    APIBitBlt = 0
    APIStretchBlt = 3
End Enum

' ======================================================================================
' Types
' ======================================================================================

Private Type BITMAP
    bmType                  As Long              '/* Type of bitmap */
    bmWidth                 As Long              '/* Pixel width */
    bmHeight                As Long              '/* Pixel height */
    bmWidthBytes            As Long              '/* Byte width = 3 x Pixel width */
    bmPlanes                As Integer           '/* Color depth of bitmap */
    bmBitsPixel             As Integer           '/* Bits per pixel, must be 16 or 24 */
    bmBits                  As Long              '/* This is the pointer to the bitmap data */
End Type

'/* �豸�޹�λͼ (DIB)�Ĵ�С����ɫ��Ϣ  (��λ�� bmp �ļ��Ŀ�ͷ��) 40 bytes */
Private Type BITMAPINFOHEADER
    biSize                                 As Long    '/* �ṹ���� */
    biwidth                                As Long    '/* ָ��λͼ�Ŀ��ȣ�������Ϊ��λ */
    biheight                               As Long    '/* ָ��λͼ�ĸ߶ȣ�������Ϊ��λ */
    biPlanes                               As Integer '/* ָ��Ŀ���豸�ļ���(����Ϊ 1 ) */
    biBitCount                             As Integer '/* λͼ����ɫλ��,ÿһ�����ص�λ(1��4��8��16��24��32) */
    biCompression                          As Long    '/* ָ��ѹ������(BI_RGB Ϊ��ѹ��) */
    biSizeImage                            As Long    '/* ͼ��Ĵ�С,���ֽ�Ϊ��λ,����BI_RGB��ʽ��,������Ϊ0 */
    biXPelsPerMeter                        As Long    '/* ָ���豸ˮ׼�ֱ��ʣ���ÿ�׵�����Ϊ��λ */
    biYPelsPerMeter                        As Long    '/* ��ֱ�ֱ��ʣ�����ͬ�� */
    biClrUsed                              As Long    '/* ˵��λͼʵ��ʹ�õĲ�ɫ���е���ɫ������,��Ϊ0�Ļ�,˵��ʹ�����е�ɫ���� */
    biClrImportant                         As Long    '/* ˵����ͼ����ʾ����ҪӰ�����ɫ��������Ŀ�������0����ʾ����Ҫ */
End Type

'/* �������ɺ졢�̡�����ɵ���ɫ��� */
Private Type RGBQUAD
    rgbBlue                                As Byte
    rgbGreen                               As Byte
    rgbRed                                 As Byte
    rgbReserved                            As Byte    '/* '����������Ϊ 0 */
End Type

Private Type BITMAPINFO
    bmiHeader                              As BITMAPINFOHEADER
    bmiColors                              As RGBQUAD
End Type

' ======================================================================================
' API declares:
' ======================================================================================
'## �豸�������� ###############################
'== Fun ========================================
'CreateCompatibleDC ����һ�����ض��豸����һ�µ��ڴ��豸����������ֵ��Long�����豸�����������������Ϊ��
'DeleteDC           ɾ��ר���豸��������Ϣ�������ͷ�������ش�����Դ����Ҫ��������GetDC����ȡ�ص��豸����������ֵ��Long��ִ�гɹ���Ϊ���㣬ʧ����Ϊ��
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'## ��ͼ���� ###################################
'== Fun ========================================
'DeleteObject ���������ɾ��GDI���󣬱��续�ʡ�ˢ�ӡ����塢λͼ�������Լ���ɫ��ȵȡ�����ʹ�õ�����ϵͳ��Դ���ᱻ�ͷš�����ֵ��Long�������ʾ�ɹ������ʾʧ��
'GetObject    ȡ�ö�ָ���������˵����һ���ṹ��windows�ֲὨ����GetObject������������øú�����GetObjectAPI��vb�����ڱ�����GetObject�ؼ��ֻ���������ֵ��Long����lpObject��Ϊ�㣨��ByVal As Long����������´���һ��NULL����������������û������ĳ��ȡ���ִ�гɹ�����������ṹ�ڲ���ʵ���ֽ�������ʧ�ܣ�������ֵ
'SelectObject ÿ���豸������������ѡ�����е�ͼ�ζ������а���λͼ��ˢ�ӡ����塢�����Լ�����ȵȡ�һ��ѡ���豸������ֻ����һ������ѡ���Ķ�������豸�����Ļ�ͼ������ʹ�á����磬��ǰѡ���Ļ��ʾ��������豸�����������߶���ɫ����ʽ������ֵ��Long������ǰѡ���豸��������ͬhObject���͵�һ������ľ�������ʾ��������ѡ���Ķ�����һ������Region��������������г���֮һ��SIMPLEREGION�� COMPLEXREGION �� NULLREGION ��������������� GDI_ERROR ��ʾ����
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'## ��դ���㺯�� ###############################
'== Fun ========================================
'BitBlt            ��һ��λͼ��һ���豸�������Ƶ���һ����Դ��Ŀ��DC�໥�������ݡ�����ֵ��Long�������ʾ�ɹ������ʾʧ�ܡ�������GetLastError
'SetStretchBltMode ָ��StretchBlt �� StretchDIBits����������ģʽ����������ģʽ������Windows��ζ������������޳���ɨ���߽��п��ơ�����VB����Ϳؼ���������API��ͼ������ʹ���������������ָ�ԭ����StretchBltģʽ
'StretchBlt        ��һ��λͼ��һ���豸�������Ƶ���һ����Դ��Ŀ��DC�໥�������ݡ�������������豸�����ж���һ��Ŀ����Σ�����λͼ�ж���һ��Դͼ��Դ���λ������Ҫ�����������Ա���Ŀ����εĴ�С���������ֵ��Long�������ʾ�ɹ������ʾʧ�ܡ�������GetLastError
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long

'## λͼ���� ###################################
'== Fun ========================================
'CreateDIBSection ����һ��DIBSection������һ��GDI���󣬿���һ�����豸�й�λͼ����ʹ�á����ǣ������ڲ���Ϊһ�����豸�޹�λͼ���档����ֵ��Long��ִ�гɹ�����DIBSectionλͼ�ľ�������ʾʧ�ܡ�������GetLastError
'LoadImage        ����һ��λͼ��ͼ���ָ��
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

' ======================================================================================
' Private variables:
' ======================================================================================
' Handle to the current DIBSection:
Private m_hDIB As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO

' Scans must align on dword boundaries:
Public Property Get BytesPerScanLine() As Long
    Dim lngNum As Long
    
    If Me.Width <= 0 Then
       BytesPerScanLine = 0
       Exit Function
    End If
    BytesPerScanLine = ((m_tBI.bmiHeader.biwidth * m_tBI.bmiHeader.biBitCount + 31) \ 32) * 4
End Property

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hBmpOld)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    m_hDC = 0: m_hDIB = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Property Get ColorBit() As Long
    ColorBit = m_tBI.bmiHeader.biBitCount
End Property

Public Function Create(ByVal Width As Long, ByVal Height As Long) As Boolean
    Call ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
       If CreateDIB(m_hDC, Width, Height, m_hDIB) Then
          m_hBmpOld = SelectObject(m_hDC, m_hDIB)
          Create = True
       Else
          DeleteDC m_hDC
          m_hDC = 0
       End If
    End If
End Function

Public Function CreateDIB(ByVal lhDC As Long, _
                          ByVal Width As Long, _
                          ByVal Height As Long, _
                          ByRef hDIB As Long) As Boolean
    Dim lngI              As Long
    
    With m_tBI.bmiHeader
         .biBitCount = 24
         .biCompression = BI_RGB
         .biClrImportant = 0
         .biClrUsed = 0
         .biheight = Height
         .biPlanes = 1
         .biSize = Len(m_tBI.bmiHeader)
         .biwidth = Width
         .biXPelsPerMeter = 0
         .biYPelsPerMeter = 0
         .biSizeImage = Me.BytesPerScanLine * .biheight
    End With
    hDIB = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
    CreateDIB = (hDIB <> 0)
End Function

Public Function CreateFromFile(ByVal FileName As String) As Boolean
    Dim hMap                   As Long
    Dim TempPic                As IPictureDisp
    Dim lhDC                   As Long
    Dim lhDCDesktop            As Long
    Dim lhBmpOld               As Long
    Dim tBMP                   As BITMAP
    Dim lhWnd                  As Long
    Dim blnOk                  As Boolean
    
    If Not ExistFile(FileName) Then
        MsgBox "File does not exist:" & vbCrLf & FileName, vbCritical, App.Title
        Exit Function
    End If
    
    hMap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_CREATEDIBSECTION Or LR_LOADFROMFILE)
    If hMap Then
        '
    Else
        On Error Resume Next
        Set TempPic = LoadPicture(FileName)
        On Error GoTo 0
        If TempPic Is Nothing Then
        Else
            If TempPic.Type = vbPicTypeBitmap Then
                hMap = TempPic.handle
            Else
                Set TempPic = Nothing
            End If
        End If
    End If
    
    If hMap Then
       Call GetObjectAPI(hMap, Len(tBMP), tBMP)
       blnOk = Create(tBMP.bmWidth, tBMP.bmHeight)
       If blnOk Then
          lhDC = CreateCompatibleDC(0)
          If (lhDC <> 0) Then
              lhBmpOld = SelectObject(lhDC, hMap)
              Call LoadPictureBlt(lhDC)
              Call SelectObject(lhDC, lhBmpOld)
              Call DeleteDC(lhDC)
          End If
       End If
    End If
    If hMap <> 0 Then Call DeleteObject(hMap)
    CreateFromFile = blnOk 'indicate success
End Function

Public Function CreateFromHandle(ByRef hImage As Long, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1) As Boolean
    Dim lhDC                      As Long
    Dim lhBmpOld                  As Long
    Dim blnOk                     As Boolean
    Dim byBGR()                   As Byte
    Dim tBMP                      As BITMAP
    Dim tBmi                      As BITMAPINFO
    
    If lSrcWidth < 0 Or lSrcHeight < 0 Then
       Call GetObjectAPI(hImage, Len(tBMP), tBMP)
       blnOk = Create(tBMP.bmWidth, tBMP.bmHeight)
    Else
       blnOk = Create(lSrcWidth, lSrcHeight)
    End If
    
    If blnOk Then
       lhDC = CreateCompatibleDC(0)
       If (lhDC <> 0) Then
           lhBmpOld = SelectObject(lhDC, hImage)
           With tBmi.bmiHeader
                .biBitCount = 32
                .biClrImportant = 0
                .biClrUsed = 0
                .biCompression = BI_RGB
                .biheight = m_tBI.bmiHeader.biheight
                .biPlanes = 1
                .biSize = Len(tBmi.bmiHeader)
                .biwidth = m_tBI.bmiHeader.biwidth
                .biXPelsPerMeter = 0
                .biYPelsPerMeter = 0
                .biSizeImage = (((.biwidth * 3) + 3) And &HFFFFFFFC) * .biheight
           End With
           ReDim byBGR(1 To 4, 1 To m_tBI.bmiHeader.biwidth, 1 To m_tBI.bmiHeader.biheight)
           Call GetDIBits(lhDC, hImage, 0, m_tBI.bmiHeader.biheight, byBGR(1, 1, 1), tBmi, DIB_RGB_COLORS)
           Call SetDIBits(m_hDC, m_hDIB, 0, m_tBI.bmiHeader.biheight, byBGR(1, 1, 1), tBmi, DIB_RGB_COLORS)
           Call SelectObject(lhDC, lhBmpOld)
           Call DeleteDC(lhDC)
           CreateFromHandle = True
       End If
    End If
End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1) As Boolean
    Dim lhDC                      As Long
    Dim lhBmpOld                  As Long
    Dim tBMP                      As BITMAP
    Dim blnOk                     As Boolean
    
    If lSrcWidth < 0 Or lSrcHeight < 0 Then
       GetObjectAPI picThis.handle, Len(tBMP), tBMP
       blnOk = Create(tBMP.bmWidth, tBMP.bmHeight)
    Else
       blnOk = Create(lSrcWidth, lSrcHeight)
    End If
    
    If blnOk Then
       lhDC = CreateCompatibleDC(0)
       If (lhDC <> 0) Then
           lhBmpOld = SelectObject(lhDC, picThis.handle)
           Call LoadPictureBlt(lhDC, 0, 0, lSrcWidth, lSrcHeight, lSrcLeft, lSrcTop, vbSrcCopy)
           Call SelectObject(lhDC, lhBmpOld)
           Call DeleteDC(lhDC)
           CreateFromPicture = True
       End If
    End If
End Function

Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property

' gets a thumbnail DIB.
Public Function GetThumbnailDIB(ByVal thumbWidth As Long, ByVal thumbHeight As Long, _
                                Optional ByVal AutoFitSize As Boolean = True) As cDIBSection
    Dim sngScale                As Long
    Dim lngWidth                As Long
    Dim lngHeight               As Long
    Dim clsThumb                As cDIBSection
    
    If AutoFitSize Then
       If m_tBI.bmiHeader.biwidth > thumbWidth Or m_tBI.bmiHeader.biheight > thumbHeight Then
          If m_tBI.bmiHeader.biwidth > m_tBI.bmiHeader.biheight Then
             sngScale = m_tBI.bmiHeader.biwidth / thumbWidth
             If m_tBI.bmiHeader.biheight / sngScale > thumbHeight Then sngScale = m_tBI.bmiHeader.biheight / thumbHeight
          Else
             sngScale = m_tBI.bmiHeader.biheight / thumbHeight
             If m_tBI.bmiHeader.biwidth / sngScale > thumbWidth Then sngScale = m_tBI.bmiHeader.biwidth / thumbWidth
          End If
          lngWidth = CLng(m_tBI.bmiHeader.biwidth / sngScale)
          lngHeight = CLng(m_tBI.bmiHeader.biheight / sngScale)
       Else
          lngWidth = m_tBI.bmiHeader.biwidth
          lngHeight = m_tBI.bmiHeader.biheight
       End If
    Else
       lngWidth = thumbWidth
       lngHeight = thumbHeight
    End If
    Set clsThumb = New cDIBSection
    If Not clsThumb.Create(lngWidth, lngHeight) Then GoTo ErrorHandle
    Me.PaintPicture clsThumb.hDC, APIStretchBlt, 0, 0, lngWidth, lngHeight
    Set GetThumbnailDIB = clsThumb
    Exit Function
ErrorHandle:
    Set clsThumb = Nothing
End Function

Public Property Get DataSize() As Long
    DataSize = m_tBI.bmiHeader.biSizeImage
End Property

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biheight
End Property

Public Sub LoadPictureBlt(ByVal lSrcDC As Long, _
                          Optional ByVal lDestLeft As Long = 0, _
                          Optional ByVal lDestTop As Long = 0, _
                          Optional ByVal lDestWidth As Long = -1, _
                          Optional ByVal lDestHeight As Long = -1, _
                          Optional ByVal lSrcLeft As Long = 0, _
                          Optional ByVal lSrcTop As Long = 0, _
                          Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    If lDestWidth < 0 Then lDestWidth = m_tBI.bmiHeader.biwidth
    If lDestHeight < 0 Then lDestHeight = m_tBI.bmiHeader.biheight
    Call BitBlt(m_hDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, lSrcDC, lSrcLeft, lSrcTop, eRop)
End Sub

Public Function PaintPicture(ByVal lhDC As Long, _
                             Optional ByVal APIProc As UseAPIPaintPicture = APIBitBlt, _
                             Optional ByVal lDestLeft As Long = 0, _
                             Optional ByVal lDestTop As Long = 0, _
                             Optional ByVal lDestWidth As Long = -1, _
                             Optional ByVal lDestHeight As Long = -1, _
                             Optional ByVal lSrcLeft As Long = 0, _
                             Optional ByVal lSrcTop As Long = 0, _
                             Optional ByVal lSrcWidth As Long = -1, _
                             Optional ByVal lSrcHeight As Long = -1, _
                             Optional ByVal eRop As RasterOpConstants = vbSrcCopy) As Boolean
    Dim lngPtr               As Long
    Dim blnOk                As Boolean
    
    blnOk = False
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biwidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biheight
    If (lSrcWidth < 0) Then lSrcWidth = m_tBI.bmiHeader.biwidth
    If (lSrcHeight < 0) Then lSrcHeight = m_tBI.bmiHeader.biheight
    Select Case APIProc
           Case APIBitBlt
             blnOk = BitBlt(lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, _
                            m_hDC, lSrcLeft, lSrcTop, eRop)
           Case APIStretchBlt
             Call SetStretchBltMode(lhDC, STRETCH_HALFTONE)
             blnOk = StretchBlt(lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, _
                                m_hDC, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, eRop)
    End Select
    PaintPicture = blnOk
End Function

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biwidth
End Property

Private Sub Class_Initialize()
    'load all those mutthas...
    Call ClearUp
End Sub

Private Sub Class_Terminate()
    Call ClearUp
End Sub


