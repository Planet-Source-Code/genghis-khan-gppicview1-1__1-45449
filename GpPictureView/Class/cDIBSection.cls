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
' Copyright © 1999 Steve McMahon

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
Private Const DIB_RGB_COLORS              As Long = 0    ' BITMAPINFO°üº¬ÁËÒ»¸öÑÕÉ«±í£¬ÆäÖÐ±£´æÓÐ32Î»ÑÕÉ«£¨RGBQUAD£©
Private Const DIB_PAL_COLORS              As Long = 1    ' BITMAPINFO°üº¬ÁËÒ»¸ö16Î»µ÷É«°åË÷ÒýµÄÊý×é

Private Const STRETCH_ANDSCANS = 1
Private Const STRETCH_ORSCANS = 2
Private Const STRETCH_DELETESCANS = 3
Private Const STRETCH_HALFTONE = 4

' ³£¼ûµÄÈýÔª¹âÕ¤²Ù×÷´úÂë
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

'Ö¸¶¨±»×°ÔØÍ¼ÏñÀàÐÍ,LoadImage,CopyImage
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

'/* Éè±¸ÎÞ¹ØÎ»Í¼ (DIB)µÄ´óÐ¡¼°ÑÕÉ«ÐÅÏ¢  (ËüÎ»ÓÚ bmp ÎÄ¼þµÄ¿ªÍ·´¦) 40 bytes */
Private Type BITMAPINFOHEADER
    biSize                                 As Long    '/* ½á¹¹³¤¶È */
    biwidth                                As Long    '/* Ö¸¶¨Î»Í¼µÄ¿í¶È£¬ÒÔÏñËØÎªµ¥Î» */
    biheight                               As Long    '/* Ö¸¶¨Î»Í¼µÄ¸ß¶È£¬ÒÔÏñËØÎªµ¥Î» */
    biPlanes                               As Integer '/* Ö¸¶¨Ä¿±êÉè±¸µÄ¼¶Êý(±ØÐëÎª 1 ) */
    biBitCount                             As Integer '/* Î»Í¼µÄÑÕÉ«Î»Êý,Ã¿Ò»¸öÏñËØµÄÎ»(1£¬4£¬8£¬16£¬24£¬32) */
    biCompression                          As Long    '/* Ö¸¶¨Ñ¹ËõÀàÐÍ(BI_RGB Îª²»Ñ¹Ëõ) */
    biSizeImage                            As Long    '/* Í¼ÏóµÄ´óÐ¡,ÒÔ×Ö½ÚÎªµ¥Î»,µ±ÓÃBI_RGB¸ñÊ½ÊÇ,¿ÉÉèÖÃÎª0 */
    biXPelsPerMeter                        As Long    '/* Ö¸¶¨Éè±¸Ë®×¼·Ö±æÂÊ£¬ÒÔÃ¿Ã×µÄÏñËØÎªµ¥Î» */
    biYPelsPerMeter                        As Long    '/* ´¹Ö±·Ö±æÂÊ£¬ÆäËûÍ¬ÉÏ */
    biClrUsed                              As Long    '/* ËµÃ÷Î»Í¼Êµ¼ÊÊ¹ÓÃµÄ²ÊÉ«±íÖÐµÄÑÕÉ«Ë÷ÒýÊý,ÉèÎª0µÄ»°,ËµÃ÷Ê¹ÓÃËùÓÐµ÷É«°åÏî */
    biClrImportant                         As Long    '/* ËµÃ÷¶ÔÍ¼ÏóÏÔÊ¾ÓÐÖØÒªÓ°ÏìµÄÑÕÉ«Ë÷ÒýµÄÊýÄ¿£¬Èç¹ûÊÇ0£¬±íÊ¾¶¼ÖØÒª */
End Type

'/* ÃèÊöÁËÓÉºì¡¢ÂÌ¡¢À¶×é³ÉµÄÑÕÉ«×éºÏ */
Private Type RGBQUAD
    rgbBlue                                As Byte
    rgbGreen                               As Byte
    rgbRed                                 As Byte
    rgbReserved                            As Byte    '/* '±£Áô£¬±ØÐëÎª 0 */
End Type

Private Type BITMAPINFO
    bmiHeader                              As BITMAPINFOHEADER
    bmiColors                              As RGBQUAD
End Type

' ======================================================================================
' API declares:
' ======================================================================================
'## Éè±¸³¡¾°º¯Êý ###############################
'== Fun ========================================
'CreateCompatibleDC ´´½¨Ò»¸öÓëÌØ¶¨Éè±¸³¡¾°Ò»ÖÂµÄÄÚ´æÉè±¸³¡¾°¡£·µ»ØÖµ£ºLong£¬ÐÂÉè±¸³¡¾°¾ä±ú£¬Èô³ö´íÔòÎªÁã
'DeleteDC           É¾³ý×¨ÓÃÉè±¸³¡¾°»òÐÅÏ¢³¡¾°£¬ÊÍ·ÅËùÓÐÏà¹Ø´°¿Ú×ÊÔ´¡£²»Òª½«ËüÓÃÓÚGetDCº¯ÊýÈ¡»ØµÄÉè±¸³¡¾°¡£·µ»ØÖµ£ºLong£¬Ö´ÐÐ³É¹¦ÔòÎª·ÇÁã£¬Ê§°ÜÔòÎªÁã
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'## »æÍ¼º¯Êý ###################################
'== Fun ========================================
'DeleteObject ÓÃÕâ¸öº¯ÊýÉ¾³ýGDI¶ÔÏó£¬±ÈÈç»­±Ê¡¢Ë¢×Ó¡¢×ÖÌå¡¢Î»Í¼¡¢ÇøÓòÒÔ¼°µ÷É«°åµÈµÈ¡£¶ÔÏóÊ¹ÓÃµÄËùÓÐÏµÍ³×ÊÔ´¶¼»á±»ÊÍ·Å¡£·µ»ØÖµ£ºLong£¬·ÇÁã±íÊ¾³É¹¦£¬Áã±íÊ¾Ê§°Ü
'GetObject    È¡µÃ¶ÔÖ¸¶¨¶ÔÏó½øÐÐËµÃ÷µÄÒ»¸ö½á¹¹¡£windowsÊÖ²á½¨ÒéÓÃGetObjectÕâ¸öÃû×ÖÀ´ÒýÓÃ¸Ãº¯Êý¡£GetObjectAPIÔÚvbÖÐÓÃÓÚ±ÜÃâÓëGetObject¹Ø¼ü×Ö»ìÏý¡£·µ»ØÖµ£ºLong£¬ÈçlpObjectÉèÎªÁã£¨ÓÃByVal As LongÔÚÕâÖÖÇé¿öÏÂ´«µÝÒ»¸öNULL²ÎÊý£©£¬Ôò±ØÐëÉèÖÃ»º³åÇøµÄ³¤¶È¡£ÈçÖ´ÐÐ³É¹¦£¬·µ»ØÔØÈë½á¹¹ÄÚ²¿µÄÊµ¼Ê×Ö½ÚÊý£»ÈçÊ§°Ü£¬·µ»ØÁãÖµ
'SelectObject Ã¿¸öÉè±¸³¡¾°¶¼¿ÉÄÜÓÐÑ¡ÈëÆäÖÐµÄÍ¼ÐÎ¶ÔÏó¡£ÆäÖÐ°üÀ¨Î»Í¼¡¢Ë¢×Ó¡¢×ÖÌå¡¢»­±ÊÒÔ¼°ÇøÓòµÈµÈ¡£Ò»´ÎÑ¡ÈëÉè±¸³¡¾°µÄÖ»ÄÜÓÐÒ»¸ö¶ÔÏó¡£Ñ¡¶¨µÄ¶ÔÏó»áÔÚÉè±¸³¡¾°µÄ»æÍ¼²Ù×÷ÖÐÊ¹ÓÃ¡£ÀýÈç£¬µ±Ç°Ñ¡¶¨µÄ»­±Ê¾ö¶¨ÁËÔÚÉè±¸³¡¾°ÖÐÃè»æµÄÏß¶ÎÑÕÉ«¼°ÑùÊ½¡£·µ»ØÖµ£ºLong£¬ÓëÒÔÇ°Ñ¡ÈëÉè±¸³¡¾°µÄÏàÍ¬hObjectÀàÐÍµÄÒ»¸ö¶ÔÏóµÄ¾ä±ú£¬Áã±íÊ¾³ö´í¡£ÈçÑ¡¶¨µÄ¶ÔÏóÊÇÒ»¸öÇøÓò£¨Region£©£¬½á¹û¾ÍÊÇÏÂÁÐ³£ÊýÖ®Ò»£ºSIMPLEREGION£¬ COMPLEXREGION »ò NULLREGION ¶ÔÇøÓò½øÐÐÃèÊö£¬ GDI_ERROR ±íÊ¾³ö´í
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'## ¹âÕ¤ÔËËãº¯Êý ###############################
'== Fun ========================================
'BitBlt            ½«Ò»·ùÎ»Í¼´ÓÒ»¸öÉè±¸³¡¾°¸´ÖÆµ½ÁíÒ»¸ö¡£Ô´ºÍÄ¿±êDCÏà»¥¼ä±ØÐë¼æÈÝ¡£·µ»ØÖµ£ºLong£¬·ÇÁã±íÊ¾³É¹¦£¬Áã±íÊ¾Ê§°Ü¡£»áÉèÖÃGetLastError
'SetStretchBltMode Ö¸¶¨StretchBlt ºÍ StretchDIBitsº¯ÊýµÄÉìËõÄ£Ê½¡£ÕâÖÖÉìËõÄ£Ê½¶¨ÒåÁËWindowsÈçºÎ¶ÔÉìËõ¹ý³ÌÖÐÌÞ³ýµÄÉ¨ÃèÏß½øÐÐ¿ØÖÆ¡£¶ÔÓÚVB´°ÌåºÍ¿Ø¼þ£¬ÌÈÈôÔÚAPI»æÍ¼¹ý³ÌÖÐÊ¹ÓÃÕâ¸öº¯Êý£¬½¨Òé»Ö¸´Ô­À´µÄStretchBltÄ£Ê½
'StretchBlt        ½«Ò»·ùÎ»Í¼´ÓÒ»¸öÉè±¸³¡¾°¸´ÖÆµ½ÁíÒ»¸ö¡£Ô´ºÍÄ¿±êDCÏà»¥¼ä±ØÐë¼æÈÝ¡£Õâ¸öº¯Êý»áÔÚÉè±¸³¡¾°ÖÐ¶¨ÒåÒ»¸öÄ¿±ê¾ØÐÎ£¬²¢ÔÚÎ»Í¼ÖÐ¶¨ÒåÒ»¸öÔ´Í¼Ïó¡£Ô´¾ØÐÎ»á¸ù¾ÝÐèÒª½øÐÐÉìËõ£¬ÒÔ±ãÓëÄ¿±ê¾ØÐÎµÄ´óÐ¡Ïà·û¡£·µ»ØÖµ£ºLong£¬·ÇÁã±íÊ¾³É¹¦£¬Áã±íÊ¾Ê§°Ü¡£»áÉèÖÃGetLastError
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long

'## Î»Í¼´´½¨ ###################################
'== Fun ========================================
'CreateDIBSection ´´½¨Ò»¸öDIBSection¡£ÕâÊÇÒ»¸öGDI¶ÔÏó£¬¿ÉÏóÒ»·ùÓëÉè±¸ÓÐ¹ØÎ»Í¼ÄÇÑùÊ¹ÓÃ¡£µ«ÊÇ£¬ËüÔÚÄÚ²¿×÷ÎªÒ»·ùÓëÉè±¸ÎÞ¹ØÎ»Í¼±£´æ¡£·µ»ØÖµ£ºLong£¬Ö´ÐÐ³É¹¦·µ»ØDIBSectionÎ»Í¼µÄ¾ä±ú£¬Áã±íÊ¾Ê§°Ü¡£»áÉèÖÃGetLastError
'LoadImage        ÔØÈëÒ»¸öÎ»Í¼¡¢Í¼±ê»òÖ¸Õë
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


