Attribute VB_Name = "modDeclares"
Option Explicit

' ======================================================================================
' Constants
' ======================================================================================

#Const DEBUGMODE = 0

'Set of bit flags that indicate which common control classes will be loaded
'from the DLL. The dwICC value of tagINITCOMMONCONTROLSEX can
'be a combination of the following:
Public Const ICC_LISTVIEW_CLASSES = &H1          '/* listview, header
Public Const ICC_TREEVIEW_CLASSES = &H2          '/* treeview, tooltips
Public Const ICC_BAR_CLASSES = &H4               '/* toolbar, statusbar, trackbar, tooltips
Public Const ICC_TAB_CLASSES = &H8               '/* tab, tooltips
Public Const ICC_UPDOWN_CLASS = &H10             '/* updown
Public Const ICC_PROGRESS_CLASS = &H20           '/* progress
Public Const ICC_HOTKEY_CLASS = &H40             '/* hotkey
Public Const ICC_ANIMATE_CLASS = &H80            '/* animate
Public Const ICC_WIN95_CLASSES = &HFF            '/* loads everything above
Public Const ICC_DATE_CLASSES = &H100            '/* month picker, date picker, time picker, updown
Public Const ICC_USEREX_CLASSES = &H200          '/* ComboEx
Public Const ICC_COOL_CLASSES = &H400            '/* Rebar (coolbar) control


' Ö¸¶¨´°¿ÚµÄ½á¹¹ÖÐÈ¡µÃÐÅÏ¢£¬ÓÃÓÚGetWindowLong¡¢SetWindowLongº¯Êý
Public Const GWL_EXSTYLE = (-20)                 '/* À©Õ¹´°¿ÚÑùÊ½ */
Public Const GWL_HINSTANCE = (-6)                '/* ÓµÓÐ´°¿ÚµÄÊµÀýµÄ¾ä±ú */
Public Const GWL_HWNDPARENT = (-8)               '/* ¸Ã´°¿ÚÖ®¸¸µÄ¾ä±ú¡£²»ÒªÓÃSetWindowWordÀ´¸Ä±äÕâ¸öÖµ */
Public Const GWL_ID = (-12)                      '/* ¶Ô»°¿òÖÐÒ»¸ö×Ó´°¿ÚµÄ±êÊ¶·û */
Public Const GWL_STYLE = (-16)                   '/* ´°¿ÚÑùÊ½ */
Public Const GWL_USERDATA = (-21)                '/* º¬ÒåÓÉÓ¦ÓÃ³ÌÐò¹æ¶¨ */
Public Const GWL_WNDPROC = (-4)                  '/* ¸Ã´°¿ÚµÄ´°¿Úº¯ÊýµÄµØÖ· */
Public Const DWL_DLGPROC = 4                     '/* Õâ¸ö´°¿ÚµÄ¶Ô»°¿òº¯ÊýµØÖ· */
Public Const DWL_MSGRESULT = 0                   '/* ÔÚ¶Ô»°¿òº¯ÊýÖÐ´¦ÀíµÄÒ»ÌõÏûÏ¢·µ»ØµÄÖµ */
Public Const DWL_USER = 8                        '/* º¬ÒåÓÉÓ¦ÓÃ³ÌÐò¹æ¶¨ */


' GetDeviceCapsË÷Òý±í£¬ÓÃÓÚGetDeviceCapsº¯Êý
Public Const DRIVERVERSION = 0                   '/* ±¸Çý¶¯³ÌÐò°æ±¾
Public Const BITSPIXEL = 12                      '/*
Public Const LOGPIXELSX = 88                     '/*  Logical pixels/inch in X
Public Const LOGPIXELSY = 90                     '/*  Logical pixels/inch in Y

' Windows¶ÔÏó³£Êý±í£¬º¯ÊýGetSysColor
Public Const COLOR_ACTIVEBORDER = 10             '/* »î¶¯´°¿ÚµÄ±ß¿ò
Public Const COLOR_ACTIVECAPTION = 2             '/* »î¶¯´°¿ÚµÄ±êÌâ
Public Const COLOR_ADJ_MAX = 100                 '/*
Public Const COLOR_ADJ_MIN = -100                '/*
Public Const COLOR_APPWORKSPACE = 12             '/* MDI×ÀÃæµÄ±³¾°
Public Const COLOR_BACKGROUND = 1                '/*
Public Const COLOR_BTNDKSHADOW = 21              '/*
Public Const COLOR_BTNLIGHT = 22                 '/*
Public Const COLOR_BTNFACE = 15                  '/* °´Å¥
Public Const COLOR_BTNHIGHLIGHT = 20             '/* °´Å¥µÄ3D¼ÓÁÁÇø
Public Const COLOR_BTNSHADOW = 16                '/* °´Å¥µÄ3DÒõÓ°
Public Const COLOR_BTNTEXT = 18                  '/* °´Å¥ÎÄ×Ö
Public Const COLOR_CAPTIONTEXT = 9               '/* ´°¿Ú±êÌâÖÐµÄÎÄ×Ö
Public Const COLOR_GRAYTEXT = 17                 '/* »ÒÉ«ÎÄ×Ö£»ÈçÊ¹ÓÃÁË¶¶¶¯¼¼ÊõÔòÎªÁã
Public Const COLOR_HIGHLIGHT = 13                '/* Ñ¡¶¨µÄÏîÄ¿±³¾°
Public Const COLOR_HIGHLIGHTTEXT = 14            '/* Ñ¡¶¨µÄÏîÄ¿ÎÄ×Ö
Public Const COLOR_INACTIVEBORDER = 11           '/* ²»»î¶¯´°¿ÚµÄ±ß¿ò
Public Const COLOR_INACTIVECAPTION = 3           '/* ²»»î¶¯´°¿ÚµÄ±êÌâ
Public Const COLOR_INACTIVECAPTIONTEXT = 19      '/* ²»»î¶¯´°¿ÚµÄÎÄ×Ö
Public Const COLOR_MENU = 4                      '/* ²Ëµ¥
Public Const COLOR_MENUTEXT = 7                  '/* ²Ëµ¥ÕýÎÄ
Public Const COLOR_SCROLLBAR = 0                 '/* ¹ö¶¯Ìõ
Public Const COLOR_WINDOW = 5                    '/* ´°¿Ú±³¾°
Public Const COLOR_WINDOWFRAME = 6               '/* ´°¿ò
Public Const COLOR_WINDOWTEXT = 8                '/* ´°¿ÚÕýÎÄ
Public Const COLORONCOLOR = 3

' º¯ÊýCombineRgnµÄ·µ»ØÖµ£¬ÀàÐÍLong
Public Const COMPLEXREGION = 3                   '/* ÇøÓòÓÐ»¥Ïà½»µþµÄ±ß½ç */
Public Const SIMPLEREGION = 2                    '/* ÇøÓò±ß½çÃ»ÓÐ»¥Ïà½»µþ */
Public Const NULLREGION = 1                      '/* ÇøÓòÎª¿Õ */
Public Const ERRORAPI = 0                        '/* ²»ÄÜ´´½¨×éºÏÇøÓò */

' ×éºÏÁ½ÇøÓòµÄ·½·¨£¬º¯ÊýCombineRgnµÄµÄ²ÎÊýnCombineModeËùÊ¹ÓÃµÄ³£Êý
Public Const RGN_AND = 1                         '/* hDestRgn±»ÉèÖÃÎªÁ½¸öÔ´ÇøÓòµÄ½»¼¯ */
Public Const RGN_COPY = 5                        '/* hDestRgn±»ÉèÖÃÎªhSrcRgn1µÄ¿½±´ */
Public Const RGN_DIFF = 4                        '/* hDestRgn±»ÉèÖÃÎªhSrcRgn1ÖÐÓëhSrcRgn2²»Ïà½»µÄ²¿·Ö */
Public Const RGN_OR = 2                          '/* hDestRgn±»ÉèÖÃÎªÁ½¸öÇøÓòµÄ²¢¼¯ */
Public Const RGN_XOR = 3                         '/* hDestRgn±»ÉèÖÃÎª³ýÁ½¸öÔ´ÇøÓòORÖ®ÍâµÄ²¿·Ö */

' Missing Draw State constants declarations£¬²Î¿´DrawStateº¯Êý
'/* Image type */
Public Const DST_COMPLEX = &H0                   '/* »æÍ¼ÔÚÓÉlpDrawStateProc²ÎÊýÖ¸¶¨µÄ»Øµ÷º¯ÊýÆÚ¼äÖ´ÐÐ¡£lParamºÍwParam»á´«µÝ¸ø»Øµ÷ÊÂ¼þ
Public Const DST_TEXT = &H1                      '/* lParam´ú±íÎÄ×ÖµÄµØÖ·£¨¿ÉÊ¹ÓÃÒ»¸ö×Ö´®±ðÃû£©£¬wParam´ú±í×Ö´®µÄ³¤¶È
Public Const DST_PREFIXTEXT = &H2                '/* ÓëDST_TEXTÀàËÆ£¬Ö»ÊÇ & ×Ö·ûÖ¸³öÎªÏÂ¸÷×Ö·û¼ÓÉÏÏÂ»®Ïß
Public Const DST_ICON = &H3                      '/* lParam°üÀ¨Í¼±ê¾ä±ú
Public Const DST_BITMAP = &H4                    '/* lParamÖÐµÄ¾ä±ú
' /* State type */
Public Const DSS_NORMAL = &H0                    '/* ÆÕÍ¨Í¼Ïó
Public Const DSS_UNION = &H10                    '/* Í¼Ïó½øÐÐ¶¶¶¯´¦Àí
Public Const DSS_DISABLED = &H20                 '/* Í¼Ïó¾ßÓÐ¸¡µñÐ§¹û
Public Const DSS_MONO = &H80                     '/* ÓÃhBrushÃè»æÍ¼Ïó
Public Const DSS_RIGHT = &H8000                  '/*

' Built in ImageList drawing methods:
Public Const ILD_NORMAL = 0&
Public Const ILD_TRANSPARENT = 1&
Public Const ILD_BLEND25 = 2&
Public Const ILD_SELECTED = 4&
Public Const ILD_FOCUS = 4&
Public Const ILD_MASK = &H10&
Public Const ILD_IMAGE = &H20&
Public Const ILD_ROP = &H40&
Public Const ILD_OVERLAYMASK = 3840&
Public Const ILC_MASK = &H1&
Public Const ILCF_MOVE = &H0&
Public Const ILCF_SWAP = &H1&

Public Const CLR_DEFAULT = -16777216
Public Const CLR_HILIGHT = -16777216
Public Const CLR_NONE = -1

' General windows messages:
Public Const WM_COMMAND = &H111
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETFONT = &H30
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const WM_NOTIFY = &H4E&

' Show window styles
Public Const SW_SHOWNORMAL = 1
Public Const SW_ERASE = &H4
Public Const SW_HIDE = 0
Public Const SW_INVALIDATE = &H2
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_OTHERUNZOOM = 4
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTCLOSING = 1
Public Const SW_RESTORE = 9
Public Const SW_PARENTOPENING = 3
Public Const SW_SHOW = 5
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4

' ³£¼ûµÄ¹âÕ¤²Ù×÷´úÂë
Public Const BLACKNESS = &H42                    '/* ±íÊ¾Ê¹ÓÃÓëÎïÀíµ÷É«°åµÄË÷Òý0Ïà¹ØµÄÉ«²ÊÀ´Ìî³äÄ¿±ê¾ØÐÎÇøÓò£¬£¨¶ÔÈ±Ê¡µÄÎïÀíµ÷É«°å¶øÑÔ£¬¸ÃÑÕÉ«ÎªºÚÉ«£©¡£
Public Const DSTINVERT = &H550009                '/* ±íÊ¾Ê¹Ä¿±ê¾ØÐÎÇøÓòÑÕÉ«È¡·´¡£
Public Const MERGECOPY = &HC000CA                '/* ±íÊ¾Ê¹ÓÃ²¼¶ûÐÍµÄAND£¨Óë£©²Ù×÷·û½«Ô´¾ØÐÎÇøÓòµÄÑÕÉ«ÓëÌØ¶¨Ä£Ê½×éºÏÒ»Æð¡£
Public Const MERGEPAINT = &HBB0226               '/* Í¨¹ýÊ¹ÓÃ²¼¶ûÐÍµÄOR£¨»ò£©²Ù×÷·û½«·´ÏòµÄÔ´¾ØÐÎÇøÓòµÄÑÕÉ«ÓëÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«ºÏ²¢¡£
Public Const NOTSRCCOPY = &H330008               '/* ½«Ô´¾ØÐÎÇøÓòÑÕÉ«È¡·´£¬ÓÚ¿½±´µ½Ä¿±ê¾ØÐÎÇøÓò¡£
Public Const NOTSRCERASE = &H1100A6              '/* Ê¹ÓÃ²¼¶ûÀàÐÍµÄOR£¨»ò£©²Ù×÷·û×éºÏÔ´ºÍÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«Öµ£¬È»ºó½«ºÏ³ÉµÄÑÕÉ«È¡·´¡£
Public Const PATCOPY = &HF00021                  '/* ½«ÌØ¶¨µÄÄ£Ê½¿½±´µ½Ä¿±êÎ»Í¼ÉÏ¡£
Public Const PATINVERT = &H5A0049                '/* Í¨¹ýÊ¹ÓÃ²¼¶ûOR£¨»ò£©²Ù×÷·û½«Ô´¾ØÐÎÇøÓòÈ¡·´ºóµÄÑÕÉ«ÖµÓëÌØ¶¨Ä£Ê½µÄÑÕÉ«ºÏ²¢¡£È»ºóÊ¹ÓÃOR£¨»ò£©²Ù×÷·û½«¸Ã²Ù×÷µÄ½á¹ûÓëÄ¿±ê¾ØÐÎÇøÓòÄÚµÄÑÕÉ«ºÏ²¢¡£
Public Const PATPAINT = &HFB0A09                 '/* Í¨¹ýÊ¹ÓÃXOR£¨Òì»ò£©²Ù×÷·û½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòÄÚµÄÑÕÉ«ºÏ²¢¡£
Public Const SRCAND = &H8800C6                   '/* Í¨¹ýÊ¹ÓÃAND£¨Óë£©²Ù×÷·ûÀ´½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòÄÚµÄÑÕÉ«ºÏ²¢
Public Const SRCCOPY = &HCC0020                  '/* ½«Ô´¾ØÐÎÇøÓòÖ±½Ó¿½±´µ½Ä¿±ê¾ØÐÎÇøÓò¡£
Public Const SRCERASE = &H440328                 '/* Í¨¹ýÊ¹ÓÃAND£¨Óë£©²Ù×÷·û½«Ä¿±ê¾ØÐÎÇøÓòÑÕÉ«È¡·´ºóÓëÔ´¾ØÐÎÇøÓòµÄÑÕÉ«ÖµºÏ²¢¡£
Public Const SRCINVERT = &H660046                '/* Í¨¹ýÊ¹ÓÃ²¼¶ûÐÍµÄXOR£¨Òì»ò£©²Ù×÷·û½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«ºÏ²¢¡£
Public Const SRCPAINT = &HEE0086                 '/* Í¨¹ýÊ¹ÓÃ²¼¶ûÐÍµÄOR£¨»ò£©²Ù×÷·û½«Ô´ºÍÄ¿±ê¾ØÐÎÇøÓòµÄÑÕÉ«ºÏ²¢¡£
Public Const WHITENESS = &HFF0062                '/* Ê¹ÓÃÓëÎïÀíµ÷É«°åÖÐË÷Òý1ÓÐ¹ØµÄÑÕÉ«Ìî³äÄ¿±ê¾ØÐÎÇøÓò¡££¨¶ÔÓÚÈ±Ê¡ÎïÀíµ÷É«°åÀ´Ëµ£¬Õâ¸öÑÕÉ«¾ÍÊÇ°×É«£©¡£

'--- for mouse_event
Public Const MOUSE_MOVED = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000       '/*
Public Const MOUSEEVENTF_LEFTDOWN = &H2          '/* Ä£ÄâÊó±ê×ó¼ü°´ÏÂ
Public Const MOUSEEVENTF_LEFTUP = &H4            '/* Ä£ÄâÊó±ê×ó¼üÌ§Æð
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20       '/* Ä£ÄâÊó±êÖÐ¼ü°´ÏÂ
Public Const MOUSEEVENTF_MIDDLEUP = &H40         '/* Ä£ÄâÊó±êÖÐ¼ü°´ÏÂ
Public Const MOUSEEVENTF_MOVE = &H1              '/* ÒÆ¶¯Êó±ê */
Public Const MOUSEEVENTF_RIGHTDOWN = &H8         '/* Ä£ÄâÊó±êÓÒ¼ü°´ÏÂ
Public Const MOUSEEVENTF_RIGHTUP = &H10          '/* Ä£ÄâÊó±êÓÒ¼ü°´ÏÂ
Public Const MOUSETRAILS = 39                    '/*

Public Const BMP_MAGIC_COOKIE = 19778            '/* this is equivalent to ascii string "BM" */
' constants for the biCompression field
Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const BI_BITFIELDS = 3&
'Public Const BITSPIXEL = 12                     '/* Number of bits per pixel
' DIB color table identifiers
Public Const DIB_PAL_COLORS = 1                  '/* ÔÚÑÕÉ«±íÖÐ×°ÔØÒ»¸ö16Î»ËùÒÔÊý×é£¬ËüÃÇÓëµ±Ç°Ñ¡¶¨µÄµ÷É«°åÓÐ¹Ø color table in palette indices
Public Const DIB_PAL_INDICES = 2                 '/* No color table indices into surf palette
Public Const DIB_PAL_LOGINDICES = 4              '/* No color table indices into DC palette
Public Const DIB_PAL_PHYSINDICES = 2             '/* No color table indices into surf palette
Public Const DIB_RGB_COLORS = 0                  '/* ÔÚÑÕÉ«±íÖÐ×°ÔØRGBÑÕÉ«

' ======================================================================================
' Methods
' ======================================================================================
' º¯ÊýSetBkModen²ÎÊýBkMode
Public Enum KhanBackStyles
    TRANSPARENT = 1                              '/* Í¸Ã÷´¦Àí£¬¼´²»×÷ÉÏÊöÌî³ä */
    OPAQUE = 2                                   '/* ÓÃµ±Ç°µÄ±³¾°É«Ìî³äÐéÏß»­±Ê¡¢ÒõÓ°Ë¢×ÓÒÔ¼°×Ö·ûµÄ¿ÕÏ¶ */
    NEWTRANSPARENT = 3                           '/* NT4: Uses chroma-keying upon BitBlt. Undocumented feature that is not working on Windows 2000/XP.
End Enum

' ¶à±ßÐÎµÄÌî³äÄ£Ê½
Public Enum KhanPolyFillModeFalgs
    ALTERNATE = 1                                '/* ½»ÌæÌî³ä
    WINDING = 2                                  '/* ¸ù¾Ý»æÍ¼·½ÏòÌî³ä
End Enum

' DrawIconEx
Public Enum KhanDrawIconExFlags
    DI_MASK = &H1                                '/* »æÍ¼Ê±Ê¹ÓÃÍ¼±êµÄMASK²¿·Ö£¨Èçµ¥¶ÀÊ¹ÓÃ£¬¿É»ñµÃÍ¼±êµÄÑÚÄ££©
    DI_IMAGE = &H2                               '/* »æÍ¼Ê±Ê¹ÓÃÍ¼±êµÄXOR²¿·Ö£¨¼´Í¼±êÃ»ÓÐÍ¸Ã÷ÇøÓò£©
    DI_NORMAL = &H3                              '/* ÓÃ³£¹æ·½Ê½»æÍ¼£¨ºÏ²¢ DI_IMAGE ºÍ DI_MASK£©
    DI_COMPAT = &H4                              '/* Ãè»æ±ê×¼µÄÏµÍ³Ö¸Õë£¬¶ø²»ÊÇÖ¸¶¨µÄÍ¼Ïó
    DI_DEFAULTSIZE = &H8                         '/* ºöÂÔcxWidthºÍcyWidthÉèÖÃ£¬²¢²ÉÓÃÔ­Ê¼µÄÍ¼±ê´óÐ¡
End Enum

'Ö¸¶¨±»×°ÔØÍ¼ÏñÀàÐÍ,LoadImage,CopyImage
Public Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

Public Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/* ±íÊ¾´´½¨Ò»¸öÍ¼ÏñµÄ¾«È·¸±±¾£¬¶øºöÂÔ²ÎÊýcxDesiredºÍcyDesired
    LR_COPYDELETEORG = &H8                       '/* ±íÊ¾´´½¨Ò»¸ö¸±±¾ºóÉ¾³ýÔ­Ê¼Í¼Ïñ¡£
    LR_CREATEDIBSECTION = &H2000                 '/* µ±²ÎÊýuTypeÖ¸¶¨ÎªIMAGE_BITMAPÊ±£¬Ê¹µÃº¯Êý·µ»ØÒ»¸öDIB²¿·ÖÎ»Í¼£¬¶ø²»ÊÇÒ»¸ö¼æÈÝµÄÎ»Í¼¡£Õâ¸ö±êÖ¾ÔÚ×°ÔØÒ»¸öÎ»Í¼£¬¶ø²»ÊÇÓ³ÉäËüµÄÑÕÉ«µ½ÏÔÊ¾Éè±¸Ê±·Ç³£ÓÐÓÃ¡£
    LR_DEFAULTCOLOR = &H0                        '/* ÒÔ³£¹æ·½Ê½ÔØÈëÍ¼Ïó
    LR_DEFAULTSIZE = &H40                        '/* Èô cxDesired»òcyDesiredÎ´±»ÉèÎªÁã£¬Ê¹ÓÃÏµÍ³Ö¸¶¨µÄ¹«ÖÆÖµ±êÊ¶¹â±ê»òÍ¼±êµÄ¿íºÍ¸ß¡£Èç¹ûÕâ¸ö²ÎÊý²»±»ÉèÖÃÇÒcxDesired»òcyDesired±»ÉèÎªÁã£¬º¯ÊýÊ¹ÓÃÊµ¼Ê×ÊÔ´³ß´ç¡£Èç¹û×ÊÔ´°üº¬¶à¸öÍ¼Ïñ£¬ÔòÊ¹ÓÃµÚÒ»¸öÍ¼ÏñµÄ´óÐ¡¡£
    LR_LOADFROMFILE = &H10                       '/* ¸ù¾Ý²ÎÊýlpszNameµÄÖµ×°ÔØÍ¼Ïñ¡£Èô±ê¼ÇÎ´±»¸ø¶¨£¬lpszNameµÄÖµÎª×ÊÔ´Ãû³Æ¡£
    LR_LOADMAP3DCOLORS = &H1000                  '/* ½«Í¼ÏóÖÐµÄÉî»Ò(Dk Gray RGB£¨128£¬128£¬128£©)¡¢»Ò(Gray RGB£¨192£¬192£¬192£©)¡¢ÒÔ¼°Ç³»Ò(Gray RGB£¨223£¬223£¬223£©)ÏñËØ¶¼Ìæ»»³ÉCOLOR_3DSHADOW£¬COLOR_3DFACEÒÔ¼°COLOR_3DLIGHTµÄµ±Ç°ÉèÖÃ
    LR_LOADTRANSPARENT = &H20                    '/* ÈôfuLoad°üÀ¨LR_LOADTRANSPARENTºÍLR_LOADMAP3DCOLORSÁ½¸öÖµ£¬ÔòLRLOADTRANSPARENTÓÅÏÈ¡£µ«ÊÇ£¬ÑÕÉ«±í½Ó¿ÚÓÉCOLOR_3DFACEÌæ´ú£¬¶ø²»ÊÇCOLOR_WINDOW¡£
    LR_MONOCHROME = &H1                          '/* ½«Í¼Ïó×ª»»³Éµ¥É«
    LR_SHARED = &H8000                           '/* ÈôÍ¼Ïñ½«±»¶à´Î×°ÔØÔò¹²Ïí¡£Èç¹ûLR_SHAREDÎ´±»ÉèÖÃ£¬ÔòÔÙÏòÍ¬Ò»¸ö×ÊÔ´µÚ¶þ´Îµ÷ÓÃÕâ¸öÍ¼ÏñÊÇ¾Í»áÔÙ×°ÔØÒÔ±ãÕâ¸öÍ¼ÏñÇÒ·µ»Ø²»Í¬µÄ¾ä±ú¡£
    LR_COPYFROMRESOURCE = &H4000                 '/*
End Enum

Public Enum KhanDrawTextStyles
    DT_BOTTOM = &H8&                             '/* ±ØÐëÍ¬Ê±Ö¸¶¨DT_SINGLE¡£Ö¸Ê¾ÎÄ±¾¶ÔÆë¸ñÊ½»¯¾ØÐÎµÄµ×±ß
    DT_CALCRECT = &H400&                         '/* ÏóÏÂÃæÕâÑù¼ÆËã¸ñÊ½»¯¾ØÐÎ£º¶àÐÐ»æÍ¼Ê±¾ØÐÎµÄµ×±ß¸ù¾ÝÐèÒª½øÐÐÑÓÕ¹£¬ÒÔ±ãÈÝÏÂËùÓÐÎÄ×Ö£»µ¥ÐÐ»æÍ¼Ê±£¬ÑÓÕ¹¾ØÐÎµÄÓÒ²à¡£²»Ãè»æÎÄ×Ö¡£ÓÉlpRect²ÎÊýÖ¸¶¨µÄ¾ØÐÎ»áÔØÈë¼ÆËã³öÀ´µÄÖµ
    DT_CENTER = &H1&                             '/* ÎÄ±¾´¹Ö±¾ÓÖÐ
    DT_EXPANDTABS = &H40&                        '/* Ãè»æÎÄ×ÖµÄÊ±ºò£¬¶ÔÖÆ±íÕ¾½øÐÐÀ©Õ¹¡£Ä¬ÈÏµÄÖÆ±íÕ¾¼ä¾àÊÇ8¸ö×Ö·û¡£µ«ÊÇ£¬¿ÉÓÃDT_TABSTOP±êÖ¾¸Ä±äÕâÏîÉè¶¨
    DT_EXTERNALLEADING = &H200&                  '/* ¼ÆËãÎÄ±¾ÐÐ¸ß¶ÈµÄÊ±ºò£¬Ê¹ÓÃµ±Ç°×ÖÌåµÄÍâ²¿¼ä¾àÊôÐÔ£¨the external leading attribute£©
    DT_INTERNAL = &H1000&                        '/* Uses the system font to calculate text metrics
    DT_LEFT = &H0&                               '/* ÎÄ±¾×ó¶ÔÆë
    DT_NOCLIP = &H100&                           '/* Ãè»æÎÄ×ÖÊ±²»¼ôÇÐµ½Ö¸¶¨µÄ¾ØÐÎ£¬DrawTextEx is somewhat faster when DT_NOCLIP is used.
    DT_NOPREFIX = &H800&                         '/* Í¨³££¬º¯ÊýÈÏÎª & ×Ö·û±íÊ¾Ó¦ÎªÏÂÒ»¸ö×Ö·û¼ÓÉÏÏÂ»®Ïß¡£¸Ã±êÖ¾½ûÖ¹ÕâÖÖÐÐÎª
    DT_RIGHT = &H2&                              '/* ÎÄ±¾ÓÒ¶ÔÆë
    DT_SINGLELINE = &H20&                        '/* Ö»»­µ¥ÐÐ
    DT_TABSTOP = &H80&                           '/* Ö¸¶¨ÐÂµÄÖÆ±íÕ¾¼ä¾à£¬²ÉÓÃÕâ¸öÕûÊýµÄ¸ß8Î»
    DT_TOP = &H0&                                '/* ±ØÐëÍ¬Ê±Ö¸¶¨DT_SINGLE¡£Ö¸Ê¾ÎÄ±¾¶ÔÆë¸ñÊ½»¯¾ØÐÎµÄµ×±ß
    DT_VCENTER = &H4&                            '/* ±ØÐëÍ¬Ê±Ö¸¶¨DT_SINGLE¡£Ö¸Ê¾ÎÄ±¾¶ÔÆë¸ñÊ½»¯¾ØÐÎµÄÖÐ²¿
    DT_WORDBREAK = &H10&                         '/* ½øÐÐ×Ô¶¯»»ÐÐ¡£ÈçÓÃSetTextAlignº¯ÊýÉèÖÃÁËTA_UPDATECP±êÖ¾£¬ÕâÀïµÄÉèÖÃÔòÎÞÐ§
' #if(WINVER >= =&H0400)
    DT_EDITCONTROL = &H2000&                     '/* ¶ÔÒ»¸ö¶àÐÐ±à¼­¿Ø¼þ½øÐÐÄ£Äâ¡£²»ÏÔÊ¾²¿·Ö¿É¼ûµÄÐÐ
    DT_END_ELLIPSIS = &H8000&                    '/* ÌÈÈô×Ö´®²»ÄÜÔÚ¾ØÐÎÀïÈ«²¿ÈÝÏÂ£¬¾ÍÔÚÄ©Î²ÏÔÊ¾Ê¡ÂÔºÅ
    DT_PATH_ELLIPSIS = &H4000&                   '/* Èç×Ö´®°üº¬ÁË \ ×Ö·û£¬¾ÍÓÃÊ¡ÂÔºÅÌæ»»×Ö´®ÄÚÈÝ£¬Ê¹ÆäÄÜÔÚ¾ØÐÎÖÐÈ«²¿ÈÝÏÂ¡£ÀýÈç£¬Ò»¸öºÜ³¤µÄÂ·¾¶Ãû¿ÉÄÜ»»³ÉÕâÑùÏÔÊ¾¡ª¡ªc:\windows\...\doc\readme.txt
    DT_MODIFYSTRING = &H10000                    '/* ÈçÖ¸¶¨ÁËDT_ENDELLIPSES »ò DT_PATHELLIPSES£¬¾Í»á¶Ô×Ö´®½øÐÐÐÞ¸Ä£¬Ê¹ÆäÓëÊµ¼ÊÏÔÊ¾µÄ×Ö´®Ïà·û
    DT_RTLREADING = &H20000                      '/* ÈçÑ¡ÈëÉè±¸³¡¾°µÄ×ÖÌåÊôÓÚÏ£²®À´»ò°¢À­²®ÓïÏµ£¬¾Í´ÓÓÒµ½×óÃè»æÎÄ×Ö
    DT_WORD_ELLIPSIS = &H40000                   '/* Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

Public Enum KhanDrawFrameControlType
    DFC_CAPTION = 1                              '/* Title bar.
    DFC_MENU = 2                                 '/* Menu bar.
    DFC_SCROLL = 3                               '/* Scroll bar.
    DFC_BUTTON = 4                               '/* Standard button.
    DFC_POPUPMENU = 5                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Popup menu item.
End Enum

Public Enum KhanDrawFrameControlStyle
    DFCS_BUTTONCHECK = &H0                       '/* Check box.
    DFCS_BUTTONRADIOIMAGE = &H1                  '/* Image for radio button (nonsquare needs image).
    DFCS_BUTTONRADIOMASK = &H2                   '/* Mask for radio button (nonsquare needs mask).
    DFCS_BUTTONRADIO = &H4                       '/* Radio button.
    DFCS_BUTTON3STATE = &H8                      '/* Three-state button.
    DFCS_BUTTONPUSH = &H10                       '/* Push button.
    DFCS_CAPTIONCLOSE = &H0                      '/* <b>Close</b> button.
    DFCS_CAPTIONMIN = &H1                        '/* <b>Minimize</b> button.
    DFCS_CAPTIONMAX = &H2                        '/* <b>Maximize</b> button.
    DFCS_CAPTIONRESTORE = &H3                    '/* <b>Restore</b> button.
    DFCS_CAPTIONHELP = &H4                       '/* <b>Help</b> button.
    DFCS_MENUARROW = &H0                         '/* Submenu arrow.
    DFCS_MENUCHECK = &H1                         '/* Check mark.
    DFCS_MENUBULLET = &H2                        '/* Bullet.
    DFCS_MENUARROWRIGHT = &H4                    '/* Submenu arrow pointing left. This is used for the right-to-left cascading menus used with right-to-left languages such as Arabic or Hebrew.
    DFCS_SCROLLUP = &H0                          '/* Up arrow of scroll bar.
    DFCS_SCROLLDOWN = &H1                        '/* Down arrow of scroll bar.
    DFCS_SCROLLLEFT = &H2                        '/* Left arrow of scroll bar.
    DFCS_SCROLLRIGHT = &H3                       '/* Right arrow of scroll bar.
    DFCS_SCROLLCOMBOBOX = &H5                    '/* Combo box scroll bar.
    DFCS_SCROLLSIZEGRIP = &H8                    '/* Size grip in bottom-right corner of window.
    DFCS_SCROLLSIZEGRIPRIGHT = &H10              '/* Size grip in bottom-left corner of window. This is used with right-to-left languages such as Arabic or Hebrew.
    DFCS_INACTIVE = &H100                        '/* Button is inactive (grayed).
    DFCS_PUSHED = &H200                          '/* Button is pushed.
    DFCS_CHECKED = &H400                         '/* Button is checked.
    DFCS_TRANSPARENT = &H800                     '/* <b>Windows 98/Me, Windows 2000 or later:</b> The background remains untouched.
    DFCS_HOT = &H1000                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Button is hot-tracked.
    DFCS_ADJUSTRECT = &H2000                     '/* Bounding rectangle is adjusted to exclude the surrounding edge of the push button.
    DFCS_FLAT = &H4000                           '/* Button has a flat border.
    DFCS_MONO = &H8000                           '/* Button has a monochrome border.
End Enum

' Ö¸¶¨»­±ÊÑùÊ½£¬º¯ÊýCreatePenµÄ²ÎÊýCreatePenËùÊ¹ÓÃµÄ³£Êý
Public Enum KhanPenStyles
    ' CreatePen£¬ExtCreatePen
    ' »­±ÊµÄÑùÊ½
    PS_SOLID = 0                                 '/* »­±Ê»­³öµÄÊÇÊµÏß */
    PS_DASH = 1                                  '/* »­±Ê»­³öµÄÊÇÐéÏß£¨nWidth±ØÐëÊÇ1£© */
    PS_DOT = 2                                   '/* »­±Ê»­³öµÄÊÇµãÏß£¨nWidth±ØÐëÊÇ1£© */
    PS_DASHDOT = 3                               '/* »­±Ê»­³öµÄÊÇµã»®Ïß£¨nWidth±ØÐëÊÇ1£© */
    PS_DASHDOTDOT = 4                            '/* »­±Ê»­³öµÄÊÇµã-µã-»®Ïß£¨nWidth±ØÐëÊÇ1£© */
    PS_NULL = 5                                  '/* »­±Ê²»ÄÜ»­Í¼ */
    PS_INSIDEFRAME = 6                           '/* »­±ÊÔÚÓÉÍÖÔ²¡¢¾ØÐÎ¡¢Ô²½Ç¾ØÐÎ¡¢±ýÍ¼ÒÔ¼°ÏÒµÈÉú³ÉµÄ·â±Õ¶ÔÏó¿òÖÐ»­Í¼¡£ÈçÖ¸¶¨µÄ×¼È·RGBÑÕÉ«²»´æÔÚ£¬¾Í½øÐÐ¶¶¶¯´¦Àí */
    ' ExtCreatePen
    ' »­±ÊµÄÑùÊ½
    PS_USERSTYLE = 7                             '/* <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
    PS_ALTERNATE = 8                             '/* <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
    ' »­±ÊµÄ±Ê¼â
    PS_ENDCAP_ROUND = &H0                        '/* End caps are round.
    PS_ENDCAP_SQUARE = &H100                     '/* End caps are square.
    PS_ENDCAP_FLAT = &H200                       '/* End caps are flat.
    PS_ENDCAP_MASK = &HF00                       '/* Mask for previous PS_ENDCAP_XXX values.
    ' ÔÚÍ¼ÐÎÖÐÁ¬½ÓÏß¶Î»òÔÚÂ·¾¶ÖÐÁ¬½ÓÖ±ÏßµÄ·½Ê½
    PS_JOIN_ROUND = &H0                          '/* Joins are beveled.
    PS_JOIN_BEVEL = &H1000                       '/* Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
    PS_JOIN_MITER = &H2000                       '/* Joins are round.
    PS_JOIN_MASK = &HF000                        '/* Mask for previous PS_JOIN_XXX values.
    ' »­±ÊµÄÀàÐÍ
    PS_COSMETIC = &H0                            '/* The pen is cosmetic.
    PS_GEOMETRIC = &H10000                       '/* The pen is geometric.
    '
    PS_STYLE_MASK = &HF                          '/* Mask for previous PS_XXX values.
    PS_TYPE_MASK = &HF0000                       '/* Mask for previous PS_XXX (pen type).
End Enum

Public Enum KhanBrushStyle
    BS_SOLID = 0                                 '/* Solid brush.
    BS_HOLLOW = 1                                '/* Hollow brush.
    BS_NULL = 1                                  '/* Same as BS_HOLLOW.
    BS_HATCHED = 2                               '/* Hatched brush.
    BS_PATTERN = 3                               '/* Pattern brush defined by a memory bitmap.
    BS_INDEXED = 4                               '/*
    BS_DIBPATTERN = 5                            '/* A pattern brush defined by a device-independent bitmap (DIB) specification.
    BS_DIBPATTERNPT = 6                          '/* A pattern brush defined by a device-independent bitmap (DIB) specification. If <b>lbStyle</b> is BS_DIBPATTERNPT, the <b>lbHatch</b> member contains a pointer to a packed DIB.
    BS_PATTERN8X8 = 7                            '/* Same as BS_PATTERN.
    BS_DIBPATTERN8X8 = 8                         '/* Same as BS_DIBPATTERN.
    BS_MONOPATTERN = 9                           '/* The brush is a monochrome (black & white) bitmap.
End Enum

Public Enum KhanHatchStyles
    HS_HORIZONTAL = 0                            '/* Horizontal hatch.
    HS_VERTICAL = 1                              '/* Vertical hatch.
    HS_FDIAGONAL = 2                             '/* A 45-degree downward, left-to-right hatch.
    HS_BDIAGONAL = 3                             '/* A 45-degree upward, left-to-right hatch.
    HS_CROSS = 4                                 '/* Horizontal and vertical cross-hatch.
    HS_DIAGCROSS = 5                             '/* A 45-degree crosshatch.
End Enum

' DrawEdge
Public Enum KhanBorderStyles
    BDR_RAISEDOUTER = &H1                        '/* Raised outer edge.
    BDR_SUNKENOUTER = &H2                        '/* Sunken outer edge.
    BDR_RAISEDINNER = &H4                        '/* Raised inner edge.
    BDR_SUNKENINNER = &H8                        '/* Sunken inner edge.
    BDR_OUTER = &H3                              '/* (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = &HC                              '/* (BDR_RAISEDINNER Or BDR_SUNKENINNER)
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum

Public Enum KhanBorderFlags
    BF_LEFT = &H1                                '/* Left side of border rectangle.
    BF_TOP = &H2                                 '/* Top of border rectangle.
    BF_RIGHT = &H4                               '/* Right side of border rectangle.
    BF_BOTTOM = &H8                              '/* Bottom of border rectangle.
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10                           '/* Diagonal border.
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_MIDDLE = &H800                            '/* Fill in the middle.
    BF_SOFT = &H1000                             '/* Use for softer buttons.
    BF_ADJUST = &H2000                           '/* Calculate the space left over.
    BF_FLAT = &H4000                             '/* For flat rather than 3-D borders.
    BF_MONO = &H8000&                            '/* For monochrome borders
End Enum

' ´°¿ÚÖ¸¶¨Ò»¸öÐÂÎ»ÖÃºÍ×´Ì¬£¬ÓÃÓÚSetWindowPosº¯Êý
Public Enum KhanSetWindowPosStyles
    HWND_BOTTOM = 1                              '/* ½«´°¿ÚÖÃÓÚ´°¿ÚÁÐ±íµ×²¿ */
    HWND_NOTOPMOST = -2                          '/* ½«´°¿ÚÖÃÓÚÁÐ±í¶¥²¿£¬²¢Î»ÓÚÈÎºÎ×î¶¥²¿´°¿ÚµÄºóÃæ */
    HWND_TOP = 0                                 '/* ½«´°¿ÚÖÃÓÚZÐòÁÐµÄ¶¥²¿£»ZÐòÁÐ´ú±íÔÚ·Ö¼¶½á¹¹ÖÐ£¬´°¿ÚÕë¶ÔÒ»¸ö¸ø¶¨¼¶±ðµÄ´°¿ÚÏÔÊ¾µÄË³Ðò */
    HWND_TOPMOST = -1                            '/* ½«´°¿ÚÖÃÓÚÁÐ±í¶¥²¿£¬²¢Î»ÓÚÈÎºÎ×î¶¥²¿´°¿ÚµÄÇ°Ãæ */
    SWP_SHOWWINDOW = &H40                        '/* ÏÔÊ¾´°¿Ú */
    SWP_HIDEWINDOW = &H80                        '/* Òþ²Ø´°¿Ú */
    SWP_FRAMECHANGED = &H20                      '/* Ç¿ÆÈÒ»ÌõWM_NCCALCSIZEÏûÏ¢½øÈë´°¿Ú£¬¼´Ê¹´°¿ÚµÄ´óÐ¡Ã»ÓÐ¸Ä±ä */
    SWP_NOACTIVATE = &H10                        '/* ²»¼¤»î´°¿Ú */
    SWP_NOCOPYBITS = &H100                       '
    SWP_NOMOVE = &H2                             '/* ±£³Öµ±Ç°Î»ÖÃ£¨xºÍyÉè¶¨½«±»ºöÂÔ£© */
    SWP_NOOWNERZORDER = &H200                    '/* Don't do owner Z ordering */
    SWP_NOREDRAW = &H8                           '/* ´°¿Ú²»×Ô¶¯ÖØ»­ */
    SWP_NOREPOSITION = SWP_NOOWNERZORDER         '
    SWP_NOSIZE = &H1                             '/* ±£³Öµ±Ç°´óÐ¡£¨cxºÍcy»á±»ºöÂÔ£© */
    SWP_NOZORDER = &H4                           '/* ±£³Ö´°¿ÚÔÚÁÐ±íµÄµ±Ç°Î»ÖÃ£¨hWndInsertAfter½«±»ºöÂÔ£© */
    SWP_DRAWFRAME = SWP_FRAMECHANGED             '/* Î§ÈÆ´°¿Ú»­Ò»¸ö¿ò */
'    HWND_BROADCAST = &HFFFF&
'    HWND_DESKTOP = 0
End Enum

' Ö¸¶¨´´½¨´°¿ÚµÄ·ç¸ñ
Public Enum KhanCreateWindowSytles
    ' CreateWindow
    WS_BORDER = &H800000                         '/* ´´½¨Ò»¸öµ¥±ß¿òµÄ´°¿Ú¡£
    WS_CAPTION = &HC00000                        '/* ´´½¨Ò»¸öÓÐ±êÌâ¿òµÄ´°¿Ú£¨°üÀ¨WS_BODER·ç¸ñ£©¡£
    WS_CHILD = &H40000000                        '/* ´´½¨Ò»¸ö×Ó´°¿Ú¡£Õâ¸ö·ç¸ñ²»ÄÜÓëWS_POPVP·ç¸ñºÏÓÃ¡£
    WS_CHILDWINDOW = (WS_CHILD)                  '/* ÓëWS_CHILDÏàÍ¬¡£
    WS_CLIPCHILDREN = &H2000000                  '/* µ±ÔÚ¸¸´°¿ÚÄÚ»æÍ¼Ê±£¬ÅÅ³ý×Ó´°¿ÚÇøÓò¡£ÔÚ´´½¨¸¸´°¿ÚÊ±Ê¹ÓÃÕâ¸ö·ç¸ñ¡£
    WS_CLIPSIBLINGS = &H4000000                  '/* ÅÅ³ý×Ó´°¿ÚÖ®¼äµÄÏà¶ÔÇøÓò£¬Ò²¾ÍÊÇ£¬µ±Ò»¸öÌØ¶¨µÄ´°¿Ú½ÓÊÕµ½WM_PAINTÏûÏ¢Ê±£¬WS_CLIPSIBLINGS ·ç¸ñ½«ËùÓÐ²ãµþ´°¿ÚÅÅ³ýÔÚ»æÍ¼Ö®Íâ£¬Ö»ÖØ»æÖ¸¶¨µÄ×Ó´°¿Ú¡£Èç¹ûÎ´Ö¸¶¨WS_CLIPSIBLINGS·ç¸ñ£¬²¢ÇÒ×Ó´°¿ÚÊÇ²ãµþµÄ£¬ÔòÔÚÖØ»æ×Ó´°¿ÚµÄ¿Í»§ÇøÊ±£¬¾Í»áÖØ»æÁÚ½üµÄ×Ó´°¿Ú¡£
    WS_DISABLED = &H8000000                      '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª½ûÖ¹µÄ×Ó´°¿Ú¡£Ò»¸ö½ûÖ¹×´Ì¬µÄ´°ÈÕ²»ÄÜ½ÓÊÜÀ´×ÔÓÃ»§µÄÊäÈËÐÅÏ¢¡£
    WS_DLGFRAME = &H400000                       '/* ´´½¨Ò»¸ö´ø¶Ô»°¿ò±ß¿ò·ç¸ñµÄ´°¿Ú¡£ÕâÖÖ·ç¸ñµÄ´°¿Ú²»ÄÜ´ø±êÌâÌõ¡£
    WS_GROUP = &H20000                           '/* Ö¸¶¨Ò»×é¿ØÖÆµÄµÚÒ»¸ö¿ØÖÆ¡£Õâ¸ö¿ØÖÆ×éÓÉµÚÒ»¸ö¿ØÖÆºÍËæºó¶¨ÒåµÄ¿ØÖÆ×é³É£¬×ÔµÚ¶þ¸ö¿ØÖÆ¿ªÊ¼Ã¿¸ö¿ØÖÆ£¬¾ßÓÐWS_GROUP·ç¸ñ£¬Ã¿¸ö×éµÄµÚÒ»¸ö¿ØÖÆ´øÓÐWS_TABSTOP·ç¸ñ£¬´Ó¶øÊ¹ÓÃ»§¿ÉÒÔÔÚ×é¼äÒÆ¶¯¡£ÓÃ»§Ëæºó¿ÉÒÔÊ¹ÓÃ¹â±êÔÚ×éÄÚµÄ¿ØÖÆ¼ä¸Ä±ä¼üÅÌ½¹µã¡£
    WS_HSCROLL = &H100000                        '/* ´´½¨Ò»¸öÓÐË®Æ½¹ö¶¯ÌõµÄ´°¿Ú¡£
    WS_MAXIMIZE = &H1000000                      '/* ´´½¨Ò»¸ö¾ßÓÐ×î´ó»¯°´Å¥µÄ´°¿Ú¡£¸Ã·ç¸ñ²»ÄÜÓëWS_EX_CONTEXTHELP·ç¸ñÍ¬Ê±³öÏÖ£¬Í¬Ê±±ØÐëÖ¸¶¨WS_SYSMENU·ç¸ñ¡£
    WS_MAXIMIZEBOX = &H10000                     '/*
    WS_MINIMIZE = &H20000000                     '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª×îÐ¡»¯×´Ì¬µÄ´°¿Ú¡£
    WS_ICONIC = WS_MINIMIZE                      '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª×îÐ¡»¯×´Ì¬µÄ´°¿Ú¡£ÓëWS_MINIMIZE·ç¸ñÏàÍ¬¡£
    WS_MINIMIZEBOX = &H20000                     '/*
    WS_OVERLAPPED = &H0&                         '/* ²úÉúÒ»¸ö²ãµþµÄ´°¿Ú¡£Ò»¸ö²ãµþµÄ´°¿ÚÓÐÒ»¸ö±êÌâÌõºÍÒ»¸ö±ß¿ò¡£ÓëWS_TILED·ç¸ñÏàÍ¬
    WS_POPUP = &H80000000                        '/* ´´½¨Ò»¸öµ¯³öÊ½´°¿Ú¡£¸Ã·ç¸ñ²»ÄÜÓëWS_CHLD·ç¸ñÍ¬Ê±Ê¹ÓÃ¡£
    WS_SYSMENU = &H80000                         '/* ´´½¨Ò»¸öÔÚ±êÌâÌõÉÏ´øÓÐ´°¿Ú²Ëµ¥µÄ´°¿Ú£¬±ØÐëÍ¬Ê±Éè¶¨WS_CAPTION·ç¸ñ¡£
    WS_TABSTOP = &H10000                         '/* ´´½¨Ò»¸ö¿ØÖÆ£¬Õâ¸ö¿ØÖÆÔÚÓÃ»§°´ÏÂTab¼üÊ±¿ÉÒÔ»ñµÃ¼üÅÌ½¹µã¡£°´ÏÂTab¼üºóÊ¹¼üÅÌ½¹µã×ªÒÆµ½ÏÂÒ»¾ßÓÐWS_TABSTOP·ç¸ñµÄ¿ØÖÆ¡£
    WS_THICKFRAME = &H40000                      '/* ´´½¨Ò»¸ö¾ßÓÐ¿Éµ÷±ß¿òµÄ´°¿Ú¡£
    WS_SIZEBOX = WS_THICKFRAME                   '/* ÓëWS_THICKFRAME·ç¸ñÏàÍ¬
    WS_TILED = WS_OVERLAPPED                     '/* ²úÉúÒ»¸ö²ãµþµÄ´°¿Ú¡£Ò»¸ö²ãµþµÄ´°¿ÚÓÐÒ»¸ö±êÌâºÍÒ»¸ö±ß¿ò¡£ÓëWS_OVERLAPPED·ç¸ñÏàÍ¬¡£
    WS_VISIBLE = &H10000000                      '/* ´´½¨Ò»¸ö³õÊ¼×´Ì¬Îª¿É¼ûµÄ´°¿Ú¡£
    WS_VSCROLL = &H200000                        '/* ´´½¨Ò»¸öÓÐ´¹Ö±¹ö¶¯ÌõµÄ´°¿Ú¡£
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW         '/* ´´½¨Ò»¸ö¾ßÓÐWS_OVERLAPPED£¬WS_CAPTION£¬WS_SYSMENU MS_THICKFRAME£®
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) '/* ´´½¨Ò»¸ö¾ßÓÐWS_BORDER£¬WS_POPUP,WS_SYSMENU·ç¸ñµÄ´°¿Ú£¬WS_CAPTIONºÍWS_POPUPWINDOW±ØÐëÍ¬Ê±Éè¶¨²ÅÄÜÊ¹´°¿ÚÄ³µ¥¿É¼û¡£
    ' CreateWindowEx
    WS_EX_ACCEPTFILES = &H10&                    '/* Ö¸¶¨ÒÔ¸Ã·ç¸ñ´´½¨µÄ´°¿Ú½ÓÊÜÒ»¸öÍÏ×§ÎÄ¼þ¡£
    WS_EX_APPWINDOW = &H40000                    '/* µ±´°¿Ú¿É¼ûÊ±£¬½«Ò»¸ö¶¥²ã´°¿Ú·ÅÖÃµ½ÈÎÎñÌõÉÏ¡£
    WS_EX_CLIENTEDGE = &H200                     '/* Ö¸¶¨´°¿ÚÓÐÒ»¸ö´øÒõÓ°µÄ±ß½ç¡£
    WS_EX_CONTEXTHELP = &H400                    '/* ÔÚ´°¿ÚµÄ±êÌâÌõ°üº¬Ò»¸öÎÊºÅ±êÖ¾¡£µ±ÓÃ»§µã»÷ÁËÎÊºÅÊ±£¬Êó±ê¹â±ê±äÎªÒ»¸öÎÊºÅµÄÖ¸Õë¡¢Èç¹ûµã»÷ÁËÒ»¸ö×Ó´°¿Ú£¬Ôò×Ó´°ÈÕ½ÓÊÕµ½WM_HELPÏûÏ¢¡£×Ó´°¿ÚÓ¦¸Ã½«Õâ¸öÏûÏ¢´«µÝ¸ø¸¸´°¿Ú¹ý³Ì£¬¸¸´°¿ÚÔÙÍ¨¹ýHELP_WM_HELPÃüÁîµ÷ÓÃWinHelpº¯Êý¡£Õâ¸öHelpÓ¦ÓÃ³ÌÐòÏÔÊ¾Ò»¸ö°üº¬×Ó´°¿Ú°ïÖúÐÅÏ¢µÄµ¯³öÊ½´°¿Ú¡£ WS_EX_CONTEXTHELP²»ÄÜÓëWS_MAXIMIZEBOXºÍWS_MINIMIZEBOXÍ¬Ê±Ê¹ÓÃ¡£
    WS_EX_CONTROLPARENT = &H10000                '/* ÔÊÐíÓÃ»§Ê¹ÓÃTab¼üÔÚ´°¿ÚµÄ×Ó´°¿Ú¼äËÑË÷¡£
    WS_EX_DLGMODALFRAME = &H1&                   '/* ´´½¨Ò»¸ö´øË«±ßµÄ´°¿Ú£»¸Ã´°¿Ú¿ÉÒÔÔÚdwStyleÖÐÖ¸¶¨WS_CAPTION·ç¸ñÀ´´´½¨Ò»¸ö±êÌâÀ¸¡£
    WS_EX_LEFT = &H0                             '/* ´°¿Ú¾ßÓÐ×ó¶ÔÆëÊôÐÔ£¬ÕâÊÇÈ±Ê¡ÉèÖÃµÄ¡£
    WS_EX_LEFTSCROLLBAR = &H4000                 '/* Èç¹ûÍâ¿ÇÓïÑÔÊÇÈçHebrew£¬Arabic£¬»òÆäËûÖ§³Öreading order alignmentµÄÓïÑÔ£¬Ôò±êÌâÌõ£¨Èç¹û´æÔÚ£©ÔòÔÚ¿Í»§ÇøµÄ×ó²¿·Ö¡£ÈôÊÇÆäËûÓïÑÔ£¬ÔÚ¸Ã·ç¸ñ±»ºöÂÔ²¢ÇÒ²»×÷Îª´íÎó´¦Àí¡£
    WS_EX_LTRREADING = &H0                       '/* ´°¿ÚÎÄ±¾ÒÔLEFTµ½RIGHT£¨×Ô×óÏòÓÒ£©ÊôÐÔµÄË³ÐòÏÔÊ¾¡£ÕâÊÇÈ±Ê¡ÉèÖÃµÄ¡£
    WS_EX_MDICHILD = &H40                        '/* ´´½¨Ò»¸öMDI×Ó´°¿Ú¡£
    WS_EX_NOACTIVATE = &H8000000                 '/*
    WS_EX_NOPATARENTNOTIFY = &H4&                '/* Ö¸Ã÷ÒÔÕâ¸ö·ç¸ñ´´½¨µÄ´°¿ÚÔÚ±»´´½¨ºÍÏú»ÙÊ±²»Ïò¸¸´°¿Ú·¢ËÍWM_PARENTNOTFYÏûÏ¢¡£
    WS_EX_OVERLAPPEDWINDOW = &H300               '/*
    WS_EX_PALETTEWINDOW = &H188                  '/* WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOWºÍWS_WX_TOPMOST·ç¸ñµÄ×éºÏWS_EX_RIGHT:´°¿Ú¾ßÓÐÆÕÍ¨µÄÓÒ¶ÔÆëÊôÐÔ£¬ÕâÒÀÀµÓÚ´°¿ÚÀà¡£Ö»ÓÐÔÚÍâ¿ÇÓïÑÔÊÇÈçHebrew,Arabic»òÆäËûÖ§³Ö¶ÁË³Ðò¶ÔÆë£¨reading order alignment£©µÄÓïÑÔÊ±¸Ã·ç¸ñ²ÅÓÐÐ§£¬·ñÔò£¬ºöÂÔ¸Ã±êÖ¾²¢ÇÒ²»×÷Îª´íÎó´¦Àí¡£
    WS_EX_RIGHT = &H1000                         '/*
    WS_EX_RIGHTSCROLLBAR = &H0                   '/* ´¹Ö±¹ö¶¯ÌõÔÚ´°¿ÚµÄÓÒ±ß½ç¡£ÕâÊÇÈ±Ê¡ÉèÖÃµÄ¡£
    WS_EX_RTLREADING = &H2000                    '/* Èç¹ûÍâ¿ÇÓïÑÔÊÇÈçHebrew£¬Arabic£¬»òÆäËûÖ§³Ö¶ÁË³Ðò¶ÔÆë£¨reading order alignment£©µÄÓïÑÔ£¬Ôò´°¿ÚÎÄ±¾ÊÇÒ»×Ô×óÏòÓÒ£©RIGHTµ½LEFTË³ÐòµÄ¶Á³öË³Ðò¡£ÈôÊÇÆäËûÓïÑÔ£¬ÔÚ¸Ã·ç¸ñ±»ºöÂÔ²¢ÇÒ²»×÷Îª´íÎó´¦Àí¡£
    WS_EX_STATICEDGE = &H20000                   '/* Îª²»½ÓÊÜÓÃ»§ÊäÈëµÄÏî´´½¨Ò»¸ö3Ò»Î¬±ß½ç·ç¸ñ¡£
    WS_EX_TOOLWINDOW = &H80                      '/*
    WS_EX_TOPMOST = &H8&                         '/* Ö¸Ã÷ÒÔ¸Ã·ç¸ñ´´½¨µÄ´°¿ÚÓ¦·ÅÖÃÔÚËùÓÐ·Ç×î¸ß²ã´°¿ÚµÄÉÏÃæ²¢ÇÒÍ£ÁôÔÚÆäL£¬¼´Ê¹´°¿ÚÎ´±»¼¤»î¡£Ê¹ÓÃº¯ÊýSetWindowPosÀ´ÉèÖÃºÍÒÆÈ¥Õâ¸ö·ç¸ñ¡£
    WS_EX_TRANSPARENT = &H20&                    '/* Ö¸¶¨ÒÔÕâ¸ö·ç¸ñ´´½¨µÄ´°¿ÚÔÚ´°¿ÚÏÂµÄÍ¬Êô´°¿ÚÒÑÖØ»­Ê±£¬¸Ã´°¿Ú²Å¿ÉÒÔÖØ»­¡£
    WS_EX_WINDOWEDGE = &H100
End Enum

' Windows»·¾³ÓÐ¹ØµÄÐÅÏ¢£¬ÓÃÓÚGetSystemMetricsº¯Êý
Public Enum KhanSystemMetricsFlags
    SM_CXSCREEN = 0                              '/* ÆÁÄ»´óÐ¡ */
    SM_CYSCREEN = 1                              '/* ÆÁÄ»´óÐ¡ */
    SM_CXVSCROLL = 2                             '/* ´¹Ö±¹ö¶¯ÌõÖÐµÄ¼ýÍ·°´Å¥µÄ´óÐ¡ */
    SM_CYHSCROLL = 3                             '/* Ë®Æ½¹ö¶¯ÌõÉÏµÄ¼ýÍ·´óÐ¡ */
    SM_CYCAPTION = 4                             '/* ´°¿Ú±êÌâµÄ¸ß¶È */
    SM_CXBORDER = 5                              '/* ³ß´ç²»¿É±ä±ß¿òµÄ´óÐ¡ */
    SM_CYBORDER = 6                              '/* ³ß´ç²»¿É±ä±ß¿òµÄ´óÐ¡ */
    SM_CXDLGFRAME = 7                            '/* ¶Ô»°¿ò±ß¿òµÄ´óÐ¡ */
    SM_CYDLGFRAME = 8                            '/* ¶Ô»°¿ò±ß¿òµÄ´óÐ¡ */
    SM_CYVTHUMB = 9                              '/* ¹ö¶¯¿éÔÚË®Æ½¹ö¶¯ÌõÉÏµÄ´óÐ¡ */
    SM_CXHTHUMB = 10                             '/* ¹ö¶¯¿éÔÚË®Æ½¹ö¶¯ÌõÉÏµÄ´óÐ¡ */
    SM_CXICON = 11                               '/* ±ê×¼Í¼±êµÄ´óÐ¡ */
    SM_CYICON = 12                               '/* ±ê×¼Í¼±êµÄ´óÐ¡ */
    SM_CXCURSOR = 13                             '/* ±ê×¼Ö¸Õë´óÐ¡ */
    SM_CYCURSOR = 14                             '/* ±ê×¼Ö¸Õë´óÐ¡ */
    SM_CYMENU = 15                               '/* ²Ëµ¥¸ß¶È */
    SM_CXFULLSCREEN = 16                         '/* ×î´ó»¯´°¿Ú¿Í»§ÇøµÄ´óÐ¡ */
    SM_CYFULLSCREEN = 17                         '/* ×î´ó»¯´°¿Ú¿Í»§ÇøµÄ´óÐ¡ */
    SM_CYKANJIWINDOW = 18                        '/* Kanji´°¿ÚµÄ´óÐ¡£¨Height of Kanji window£© */
    SM_MOUSEPRESENT = 19                         '/* Èç°²×°ÁËÊó±êÔòÎªTRUE */
    SM_CYVSCROLL = 20                            '/* ´¹Ö±¹ö¶¯ÌõÖÐµÄ¼ýÍ·°´Å¥µÄ´óÐ¡ */
    SM_CXHSCROLL = 21                            '/* Ë®Æ½¹ö¶¯ÌõÉÏµÄ¼ýÍ·´óÐ¡ */
    SM_DEBUG = 22                                '/* ÈçwindowsµÄµ÷ÊÔ°æÕýÔÚÔËÐÐ£¬ÔòÎªTRUE */
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28                                '/* ´°¿ÚµÄ×îÐ¡³ß´ç */
    SM_CYMIN = 29                                '/* ´°¿ÚµÄ×îÐ¡³ß´ç */
    SM_CXSIZE = 30                               '/* ±êÌâÀ¸Î»Í¼µÄ´óÐ¡ */
    SM_CYSIZE = 31                               '/* ±êÌâÀ¸Î»Í¼µÄ´óÐ¡ */
    SM_CXFRAME = 32                              '/* ³ß´ç¿É±ä±ß¿òµÄ´óÐ¡£¨ÔÚwin95ºÍnt 4.0ÖÐÊ¹ÓÃSM_C?FIXEDFRAME£© */
    SM_CYFRAME = 33                              '/* ³ß´ç¿É±ä±ß¿òµÄ´óÐ¡ */
    SM_CXMINTRACK = 34                           '/* ´°¿ÚµÄ×îÐ¡¹ì¼£¿í¶È */
    SM_CYMINTRACK = 35                           '/* ´°¿ÚµÄ×îÐ¡¹ì¼£¿í¶È */
    SM_CXDOUBLECLK = 36                          '/* Ë«»÷ÇøÓòµÄ´óÐ¡£¨Ö¸¶¨ÆÁÄ»ÉÏÒ»¸öÌØ¶¨µÄÏÔÊ¾ÇøÓò£¬Ö»ÓÐÔÚÕâ¸öÇøÓòÄÚÁ¬Ðø½øÐÐÁ½´ÎÊó±êµ¥»÷£¬²ÅÓÐ¿ÉÄÜ±»µ±×÷Ë«»÷ÊÂ¼þ´¦Àí£© */
    SM_CYDOUBLECLK = 37                          '/* Ë«»÷ÇøÓòµÄ´óÐ¡ */
    SM_CXICONSPACING = 38                        '/* ×ÀÃæÍ¼±êÖ®¼äµÄ¼ä¸ô¾àÀë¡£ÔÚwin95ºÍnt 4.0ÖÐÊÇÖ¸´óÍ¼±êµÄ¼ä¾à */
    SM_CYICONSPACING = 39                        '/* ×ÀÃæÍ¼±êÖ®¼äµÄ¼ä¸ô¾àÀë¡£ÔÚwin95ºÍnt 4.0ÖÐÊÇÖ¸´óÍ¼±êµÄ¼ä¾à */
    SM_MENUDROPALIGNMENT = 40                    '/* Èçµ¯³öÊ½²Ëµ¥¶ÔÆë²Ëµ¥À¸ÏîÄ¿µÄ×ó²à£¬ÔòÎªÁã */
    SM_PENWINDOWS = 41                           '/* Èç×°ÔØÁËÖ§³Ö±Ê´°¿ÚµÄDLL£¬Ôò±íÊ¾±Ê´°¿ÚµÄ¾ä±ú */
    SM_DBCSENABLED = 42                          '/* ÈçÖ§³ÖË«×Ö½ÚÔòÎªTRUE */
    SM_CMOUSEBUTTONS = 43                        '/* Êó±ê°´Å¥£¨°´¼ü£©µÄÊýÁ¿¡£ÈçÃ»ÓÐÊó±ê£¬¾ÍÎªÁã */
    SM_CMETRICS = 44                             '/* ¿ÉÓÃÏµÍ³»·¾³µÄÊýÁ¿ */
End Enum

' SetMapMode
Public Enum KhanMapModeStyles
    MM_ANISOTROPIC = 8                           '/* Âß¼­µ¥Î»×ª»»³É¾ßÓÐÈÎÒâ±ÈÀýÖáµÄÈÎÒâµ¥Î»£¬ÓÃSetWindowExtExºÍSetViewportExtExº¯Êý¿ÉÖ¸¶¨µ¥Î»¡¢·½ÏòºÍ±ÈÀý¡£
    MM_HIENGLISH = 5                             '/* Ã¿¸öÂß¼­µ¥Î»×ª»»Îª0.001inch(Ó¢´ç)£¬XµÄÕý·½ÃæÏòÓÒ£¬YµÄÕý·½ÏòÏòÉÏ
    MM_HIMETRIC = 3                              '/* Ã¿¸öÂß¼­µ¥Î»×ª»»Îª0.01millimeter(ºÁÃ×)£¬XÕý·½ÏòÏòÓÒ£¬YµÄÕý·½ÏòÏòÉÏ¡£
    MM_ISOTROPIC = 7                             '/* ÊÓ¿ÚºÍ´°¿Ú·¶Î§ÈÎÒâ£¬Ö»ÊÇxºÍyÂß¼­µ¥Ôª³ß´çÒªÏàÍ¬
    MM_LOENGLISH = 4                             '/* Ã¿¸öÂß¼­µ¥Î»×ª»»ÎªÓ¢´ç£¬XÕý·½ÏòÏòÓÒ£¬YÕý·½ÏòÏòÉÏ¡£
    MM_LOMETRIC = 2                              '/* Ã¿¸öÂß¼­µ¥Î»×ª»»ÎªºÁÃ×£¬XÕý·½ÏòÏòÓÒ£¬YÕý·½ÏòÏòÉÏ¡£
    MM_TEXT = 1                                  '/* Ã¿¸öÂß¼­µ¥Î»×ª»»ÎªÒ»¸öÉèÖÃ±¸ËØ£¬XÕý·½ÏòÏòÓÒ£¬YÕý·½ÏòÏòÏÂ¡£
    MM_TWIPS = 6                                 '/* Ã¿¸öÂß¼­µ¥Î»×ª»»Îª1 twip (1/1440 inch)£¬XÕý·½ÏòÏòÓÒ£¬Y·½ÏòÏòÉÏ¡£
End Enum

' GetROP2,SetROP2
Public Enum EnumDrawModeFlags
    R2_BLACK = 1                                 '/* ºÚÉ«
    R2_COPYPEN = 13                              '/* »­±ÊÑÕÉ«
    R2_LAST = 16
    R2_MASKNOTPEN = 3                            '/* »­±ÊÑÕÉ«µÄ·´É«ÓëÏÔÊ¾ÑÕÉ«½øÐÐANDÔËËã
    R2_MASKPEN = 9                               '/* ÏÔÊ¾ÑÕÉ«Óë»­±ÊÑÕÉ«½øÐÐANDÔËËã
    R2_MASKPENNOT = 5                            '/* ÏÔÊ¾ÑÕÉ«µÄ·´É«Óë»­±ÊÑÕÉ«½øÐÐANDÔËËã
    R2_MERGENOTPEN = 12                          '/* »­±ÊÑÕÉ«µÄ·´É«ÓëÏÔÊ¾ÑÕÉ«½øÐÐORÔËËã
    R2_MERGEPEN = 15                             '/* »­±ÊÑÕÉ«ÓëÏÔÊ¾ÑÕÉ«½øÐÐORÔËËã
    R2_MERGEPENNOT = 14                          '/* ÏÔÊ¾ÑÕÉ«µÄ·´É«Óë»­±ÊÑÕÉ«½øÐÐORÔËËã
    R2_NOP = 11                                  '/* ²»±ä
    R2_NOT = 6                                   '/* µ±Ç°ÏÔÊ¾ÑÕÉ«µÄ·´É«
    R2_NOTCOPYPEN = 4                            '/* R2_COPYPENµÄ·´É«
    R2_NOTMASKPEN = 8                            '/* R2_MASKPENµÄ·´É«
    R2_NOTMERGEPEN = 2                           '/* R2_MERGEPENµÄ·´É«
    R2_NOTXORPEN = 10                            '/* R2_XORPENµÄ·´É«
    R2_WHITE = 16                                '/* °×É«
    R2_XORPEN = 7                                '/* ÏÔÊ¾ÑÕÉ«Óë»­±ÊÑÕÉ«½øÐÐÒì»òÔËËã
End Enum

' ======================================================================================
' Types
' ======================================================================================

Public Type tagINITCOMMONCONTROLSEX              '/* icc
   dwSize                   As Long              '/* size of this structure
   dwICC                    As Long              '/* flags indicating which classes to be initialized.
End Type

Public Type POINTAPI
   x                        As Long
   y                        As Long
End Type

Public Type RECT
    Left                     As Long
   Top                      As Long
   Right                    As Long
   Bottom                   As Long
End Type

Public Type LOGPEN
    lopnStyle               As Long
    lopnWidth               As POINTAPI
    lopnColor               As Long
End Type

Public Type LOGBRUSH
   lbStyle                  As Long
   lbColor                  As Long
   lbHatch                  As Long
End Type

' Õâ¸ö½á¹¹°üº¬ÁË¸½¼ÓµÄ»æÍ¼²ÎÊý£¬º¯ÊýDrawTextEx
Public Type DRAWTEXTPARAMS
    cbSize                  As Long              '/* Specifies the structure size, in bytes */
    iTabLength              As Long              '/* Specifies the size of each tab stop, in units equal to the average character width */
    iLeftMargin             As Long              '/* Specifies the left margin, in units equal to the average character width */
    iRightMargin            As Long              '/* Specifies the right margin, in units equal to the average character width */
    uiLengthDrawn           As Long              '/* Receives the number of characters processed by DrawTextEx, including white-space characters. */
                                                 '/* The number can be the length of the string or the index of the first line that falls below the drawing area. */
                                                 '/* Note that DrawTextEx always processes the entire string if the DT_NOCLIP formatting flag is specified */
End Type

Private Const LF_FACESIZE   As Long = 32
Public Type LOGFONT
   lfHeight                 As Long              '/* The font size (see below) */
   lfWidth                  As Long              '/* Normally you don't set this, just let Windows create the Default */
   lfEscapement             As Long              '/* The angle, in 0.1 degrees, of the font */
   lfOrientation            As Long              '/* Leave as default */
   lfWeight                 As Long              '/* Bold, Extra Bold, Normal etc */
   lfItalic                 As Byte              '/* As it says */
   lfUnderline              As Byte              '/* As it says */
   lfStrikeOut              As Byte              '/* As it says */
   lfCharSet                As Byte              '/* As it says */
   lfOutPrecision           As Byte              '/* Leave for default */
   lfClipPrecision          As Byte              '/* Leave for defaultv
   lfQuality                As Byte              '/* Leave for default */
   lfPitchAndFamily         As Byte              '/* Leave for default */
   lfFaceName(LF_FACESIZE)  As Byte              '/* The font name converted to a byte array */
End Type

Public Type ICONINFO
   fIcon                    As Long
   xHotspot                 As Long
   yHotspot                 As Long
   hBmMask                  As Long
   hbmColor                 As Long
End Type

Public Type IMAGEINFO
    hBitmapImage            As Long
    hBitmapMask             As Long
    cPlanes                 As Long
    cBitsPerPixel           As Long
    rcImage                 As RECT
End Type


' ======================================================================================
' API declares:
' ======================================================================================

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------ÏûÏ¢º¯ÊýºÍÏûÏ¢ÁÐ¶Óº¯Êý---------------------------------©§
'©§                                                                                    ©§
'
' µ÷ÓÃÒ»¸ö´°¿ÚµÄ´°¿Úº¯Êý£¬½«Ò»ÌõÏûÏ¢·¢¸øÄÇ¸ö´°¿Ú¡£³ý·ÇÏûÏ¢´¦ÀíÍê±Ï£¬·ñÔò¸Ãº¯Êý²»»á·µ»Ø¡£
' SendMessageBynum£¬ SendMessageByStringÊÇ¸Ãº¯ÊýµÄ¡°ÀàÐÍ°²È«¡±ÉùÃ÷ÐÎÊ½
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ½«Ò»ÌõÏûÏ¢Í¶µÝµ½Ö¸¶¨´°¿ÚµÄÏûÏ¢¶ÓÁÐ¡£Í¶µÝµÄÏûÏ¢»áÔÚWindowsÊÂ¼þ´¦Àí¹ý³ÌÖÐµÃµ½´¦Àí¡£
' ÔÚÄÇ¸öÊ±ºò£¬»áËæÍ¬Í¶µÝµÄÏûÏ¢µ÷ÓÃÖ¸¶¨´°¿ÚµÄ´°¿Úº¯Êý¡£ÌØ±ðÊÊºÏÄÇÐ©²»ÐèÒªÁ¢¼´´¦ÀíµÄ´°¿ÚÏûÏ¢µÄ·¢ËÍ
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§--------------------------------´°¿Úº¯Êý(Window)------------------------------------©§
'©§                                                                                    ©§
'
' Creating new windows:
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' ×îÐ¡»¯Ö¸¶¨µÄ´°¿Ú¡£´°¿Ú²»»á´ÓÄÚ´æÖÐÇå³ý
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
' ÆÆ»µ£¨¼´Çå³ý£©Ö¸¶¨µÄ´°¿ÚÒÔ¼°ËüµÄËùÓÐ×Ó´°¿Ú
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
' ÔÚÖ¸¶¨µÄ´°¿ÚÀïÔÊÐí»ò½ûÖ¹ËùÓÐÊó±ê¼°¼üÅÌÊäÈë
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
' ÔÚ´°¿ÚÁÐ±íÖÐÑ°ÕÒÓëÖ¸¶¨Ìõ¼þÏà·ûµÄµÚÒ»¸ö×Ó´°¿Ú
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' ÅÐ¶ÏÖ¸¶¨´°¿ÚµÄ¸¸´°¿Ú
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
' Ö¸¶¨Ò»¸ö´°¿ÚµÄÐÂ¸¸£¨ÔÚvbÀïÊ¹ÓÃ£ºÀûÓÃÕâ¸öº¯Êý£¬vb¿ÉÒÔ¶àÖÖÐÎÊ½Ö§³Ö×Ó´°¿Ú¡£
' ÀýÈç£¬¿É½«¿Ø¼þ´ÓÒ»¸öÈÝÆ÷ÒÆÖÁ´°ÌåÖÐµÄÁíÒ»¸ö¡£ÓÃÕâ¸öº¯ÊýÔÚ´°Ìå¼äÒÆ¶¯¿Ø¼þÊÇÏàµ±Ã°ÏÕµÄ£¬
' µ«È´²»Ê§ÎªÒ»¸öÓÐÐ§µÄ°ì·¨¡£ÈçÕæµÄÕâÑù×ö£¬ÇëÔÚ¹Ø±ÕÈÎºÎÒ»¸ö´°ÌåÖ®Ç°£¬×¢ÒâÓÃSetParent½«¿Ø¼þµÄ¸¸Éè»ØÔ­À´µÄÄÇ¸ö£©
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' Ëø¶¨Ö¸¶¨´°¿Ú£¬½ûÖ¹Ëü¸üÐÂ¡£Í¬Ê±Ö»ÄÜÓÐÒ»¸ö´°¿Ú´¦ÓÚËø¶¨×´Ì¬
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
' Ç¿ÖÆÁ¢¼´¸üÐÂ´°¿Ú£¬´°¿ÚÖÐÒÔÇ°ÆÁ±ÎµÄËùÓÐÇøÓò¶¼»áÖØ»­
' ÔÚvbÀïÊ¹ÓÃ£ºÈçvb´°Ìå»ò¿Ø¼þµÄÈÎºÎ²¿·ÖÐèÒª¸üÐÂ£¬¿É¿¼ÂÇÖ±½ÓÊ¹ÓÃrefresh·½·¨
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
' ¿ØÖÆ´°¿ÚµÄ¿É¼ûÐÔ
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' ¸Ä±äÖ¸¶¨´°¿ÚµÄÎ»ÖÃºÍ´óÐ¡¡£¶¥¼¶´°¿Ú¿ÉÄÜÊÜ×î´ó»ò×îÐ¡³ß´çµÄÏÞÖÆ£¬ÄÇÐ©³ß´çÓÅÏÈÓÚÕâÀïÉèÖÃµÄ²ÎÊý
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' Õâ¸öº¯ÊýÄÜÎª´°¿ÚÖ¸¶¨Ò»¸öÐÂÎ»ÖÃºÍ×´Ì¬¡£ËüÒ²¿É¸Ä±ä´°¿ÚÔÚÄÚ²¿´°¿ÚÁÐ±íÖÐµÄÎ»ÖÃ¡£
' ¸Ãº¯ÊýÓëDeferWindowPosº¯ÊýÏàËÆ£¬Ö»ÊÇËüµÄ×÷ÓÃÊÇÁ¢¼´±íÏÖ³öÀ´µÄ
' ÔÚvbÀïÊ¹ÓÃ£ºÕë¶Ôvb´°Ìå£¬ÈçËüÃÇÔÚwin32ÏÂÆÁ±Î»ò×îÐ¡»¯£¬ÔòÐèÖØÉè×î¶¥²¿×´Ì¬¡£
' ÈçÓÐ±ØÒª£¬ÇëÓÃÒ»¸ö×ÓÀà´¦ÀíÄ£¿éÀ´ÖØÉè×î¶¥²¿×´Ì¬)
' ²ÎÊý
' hwnd             Óû¶¨Î»µÄ´°¿Ú
' hWndInsertAfter  ´°¿Ú¾ä±ú¡£ÔÚ´°¿ÚÁÐ±íÖÐ£¬´°¿Úhwnd»áÖÃÓÚÕâ¸ö´°¿Ú¾ä±úµÄºóÃæ£¬²Î¿´±¾Ä£¿éÃ¶¾ÙKhanSetWindowPosStyles
' x                ´°¿ÚÐÂµÄx×ø±ê¡£ÈçhwndÊÇÒ»¸ö×Ó´°¿Ú£¬ÔòxÓÃ¸¸´°¿ÚµÄ¿Í»§Çø×ø±ê±íÊ¾
' y                ´°¿ÚÐÂµÄy×ø±ê¡£ÈçhwndÊÇÒ»¸ö×Ó´°¿Ú£¬ÔòyÓÃ¸¸´°¿ÚµÄ¿Í»§Çø×ø±ê±íÊ¾
' cx               Ö¸¶¨ÐÂµÄ´°¿Ú¿í¶È
' cy               Ö¸¶¨ÐÂµÄ´°¿Ú¸ß¶È
' wFlags           °üº¬ÁËÆì±êµÄÒ»¸öÕûÊý£¬²Î¿´±¾Ä£¿éÃ¶¾ÙKhanSetWindowPosStyles
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' ´ÓÖ¸¶¨´°¿ÚµÄ½á¹¹ÖÐÈ¡µÃÐÅÏ¢£¬nIndex²ÎÊý²Î¿´±¾Ä£¿é³£Á¿ÉùÃ÷
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' ÔÚ´°¿Ú½á¹¹ÖÐÎªÖ¸¶¨µÄ´°¿ÚÉèÖÃÐÅÏ¢£¬nIndex²ÎÊý²Î¿´±¾Ä£¿é³£Á¿ÉùÃ÷
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§------------------------------´°¿ÚÀàº¯Êý(Window Class)------------------------------©§
'©§                                                                                    ©§
'
' ÎªÖ¸¶¨µÄ´°¿ÚÈ¡µÃÀàÃû
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------Êó±êÊäÈëº¯Êý(Mouse Input)------------------------------©§
'
' »ñµÃÒ»¸ö´°¿ÚµÄ¾ä±ú£¬Õâ¸ö´°¿ÚÎ»ÓÚµ±Ç°ÊäÈëÏß³Ì£¬ÇÒÓµÓÐÊó±ê²¶»ñ£¨Êó±ê»î¶¯ÓÉËü½ÓÊÕ£©
Public Declare Function GetCapture Lib "user32" () As Long
' ½«Êó±ê²¶»ñÉèÖÃµ½Ö¸¶¨µÄ´°¿Ú¡£ÔÚÊó±ê°´Å¥°´ÏÂµÄÊ±ºò£¬Õâ¸ö´°¿Ú»áÎªµ±Ç°Ó¦ÓÃ³ÌÐò»òÕû¸öÏµÍ³½ÓÊÕËùÓÐÊó±êÊäÈë
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
' Îªµ±Ç°µÄÓ¦ÓÃ³ÌÐòÊÍ·ÅÊó±ê²¶»ñ
Public Declare Function ReleaseCapture Lib "user32" () As Long
' ¿ÉÒÔÄ£ÄâÒ»´ÎÊó±êÊÂ¼þ£¬±ÈÈç×ó¼üµ¥»÷¡¢Ë«»÷ºÍÓÒ¼üµ¥»÷µÈ
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
' Õâ¸öº¯ÊýÅÐ¶ÏÖ¸¶¨µÄµãÊÇ·ñÎ»ÓÚ¾ØÐÎlpRectÄÚ²¿
'Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------¼üÅÌÊäÈëº¯Êý(Mouse Input)------------------------------©§
'
' »ñµÃÓµÓÐÊäÈë½¹µãµÄ´°¿ÚµÄ¾ä±ú
Public Declare Function GetFocus Lib "user32" () As Long
' ÊäÈë½¹µãÉèµ½Ö¸¶¨µÄ´°¿Ú
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§----------------×ø±ê¿Õ¼äÓë±ä»»º¯Êý(Coordinate Space Transtormation)-----------------©§
'
' ÅÐ¶Ï´°¿ÚÄÚÒÔ¿Í»§Çø×ø±ê±íÊ¾µÄÒ»¸öµãµÄÆÁÄ»×ø±ê
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
' ÅÐ¶ÏÆÁÄ»ÉÏÒ»¸öÖ¸¶¨µãµÄ¿Í»§Çø×ø±ê
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------Éè±¸³¡¾°º¯Êý(Device Context)-----------------------------©§
'
' ´´½¨Ò»¸öÓëÌØ¶¨Éè±¸³¡¾°Ò»ÖÂµÄÄÚ´æÉè±¸³¡¾°¡£ÔÚ»æÖÆÖ®Ç°£¬ÏÈÒªÎª¸ÃÉè±¸³¡¾°Ñ¡¶¨Ò»¸öÎ»Í¼¡£
' ²»ÔÙÐèÒªÊ±£¬¸ÃÉè±¸³¡¾°¿ÉÓÃDeleteDCº¯ÊýÉ¾³ý¡£É¾³ýÇ°£¬ÆäËùÓÐ¶ÔÏóÓ¦»Ø¸´³õÊ¼×´Ì¬
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
' Îª×¨ÃÅÉè±¸´´½¨Éè±¸³¡¾°
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
' »ñÈ¡Ö¸¶¨´°¿ÚµÄÉè±¸³¡¾°£¬ÓÃ±¾º¯Êý»ñÈ¡µÄÉè±¸³¡¾°Ò»¶¨ÒªÓÃReleaseDCº¯ÊýÊÍ·Å£¬²»ÄÜÓÃDeleteDC
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
' ÊÍ·ÅÓÉµ÷ÓÃGetDC»òGetWindowDCº¯Êý»ñÈ¡µÄÖ¸¶¨Éè±¸³¡¾°¡£Ëü¶ÔÀà»òË½ÓÐÉè±¸³¡¾°ÎÞÐ§£¨µ«ÕâÑùµÄµ÷ÓÃ²»»áÔì³ÉËðº¦£©
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
' É¾³ý×¨ÓÃÉè±¸³¡¾°»òÐÅÏ¢³¡¾°£¬ÊÍ·ÅËùÓÐÏà¹Ø´°¿Ú×ÊÔ´¡£²»Òª½«ËüÓÃÓÚGetDCº¯ÊýÈ¡»ØµÄÉè±¸³¡¾°
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
' Ã¿¸öÉè±¸³¡¾°¶¼¿ÉÄÜÓÐÑ¡ÈëÆäÖÐµÄÍ¼ÐÎ¶ÔÏó¡£ÆäÖÐ°üÀ¨Î»Í¼¡¢Ë¢×Ó¡¢×ÖÌå¡¢»­±ÊÒÔ¼°ÇøÓòµÈµÈ¡£
' Ò»´ÎÑ¡ÈëÉè±¸³¡¾°µÄÖ»ÄÜÓÐÒ»¸ö¶ÔÏó¡£Ñ¡¶¨µÄ¶ÔÏó»áÔÚÉè±¸³¡¾°µÄ»æÍ¼²Ù×÷ÖÐÊ¹ÓÃ¡£
' ÀýÈç£¬µ±Ç°Ñ¡¶¨µÄ»­±Ê¾ö¶¨ÁËÔÚÉè±¸³¡¾°ÖÐÃè»æµÄÏß¶ÎÑÕÉ«¼°ÑùÊ½
' ·µ»ØÖµÍ¨³£ÓÃÓÚ»ñµÃÑ¡ÈëDCµÄ¶ÔÏóµÄÔ­Ê¼Öµ¡£
' »æÍ¼²Ù×÷Íê³Éºó£¬Ô­Ê¼µÄ¶ÔÏóÍ¨³£Ñ¡»ØÉè±¸³¡¾°¡£ÔÚÇå³ýÒ»¸öÉè±¸³¡¾°Ç°£¬Îñ±Ø×¢Òâ»Ö¸´Ô­Ê¼µÄ¶ÔÏó
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
' ÓÃÕâ¸öº¯ÊýÉ¾³ýGDI¶ÔÏó£¬±ÈÈç»­±Ê¡¢Ë¢×Ó¡¢×ÖÌå¡¢Î»Í¼¡¢ÇøÓòÒÔ¼°µ÷É«°åµÈµÈ¡£¶ÔÏóÊ¹ÓÃµÄËùÓÐÏµÍ³×ÊÔ´¶¼»á±»ÊÍ·Å
' ²»ÒªÉ¾³ýÒ»¸öÒÑÑ¡ÈëÉè±¸³¡¾°µÄ»­±Ê¡¢Ë¢×Ó»òÎ»Í¼¡£ÈçÉ¾³ýÒÔÎ»Í¼Îª»ù´¡µÄÒõÓ°£¨Í¼°¸£©Ë¢×Ó£¬
' Î»Í¼²»»áÓÉÕâ¸öº¯ÊýÉ¾³ý¡ª¡ªÖ»ÓÐË¢×Ó±»É¾µô
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'¸ù¾ÝÖ¸¶¨Éè±¸³¡¾°´ú±íµÄÉè±¸µÄ¹¦ÄÜ·µ»ØÐÅÏ¢
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
' È¡µÃ¶ÔÖ¸¶¨¶ÔÏó½øÐÐËµÃ÷µÄÒ»¸ö½á¹¹
' lpObject ÈÎºÎÀàÐÍ£¬ÓÃÓÚÈÝÄÉ¶ÔÏóÊý¾ÝµÄ½á¹¹¡£
' Õë¶Ô»­±Ê£¬Í¨³£ÊÇÒ»¸öLOGPEN½á¹¹£»Õë¶ÔÀ©Õ¹»­±Ê£¬Í¨³£ÊÇEXTLOGPEN£»
' Õë¶Ô×ÖÌåÊÇLOGBRUSH£»Õë¶ÔÎ»Í¼ÊÇBITMAP£»Õë¶ÔDIBSectionÎ»Í¼ÊÇDIBSECTION£»
' Õë¶Ôµ÷É«°å£¬Ó¦Ö¸ÏòÒ»¸öÕûÐÍ±äÁ¿£¬´ú±íµ÷É«°åÖÐµÄÌõÄ¿ÊýÁ¿
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' ÔÚ´°¿Ú£¨ÓÉÉè±¸³¡¾°´ú±í£©ÖÐË®Æ½ºÍ£¨»ò£©´¹Ö±¹ö¶¯¾ØÐÎ
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
' ½«Á½¸öÇøÓò×éºÏÎªÒ»¸öÐÂÇøÓò
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
' ´´½¨Ò»¸öÓÉµãX1£¬Y1ºÍX2£¬Y2ÃèÊöµÄ¾ØÐÎÇøÓò£¬²»ÓÃÊ±Ò»¶¨ÒªÓÃDeleteObjectº¯ÊýÉ¾³ý¸ÃÇøÓò
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ´´½¨Ò»¸öÓÉlpRectÈ·¶¨µÄ¾ØÐÎÇøÓò
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
' ´´½¨Ò»¸öÔ²½Ç¾ØÐÎ£¬¸Ã¾ØÐÎÓÉX1£¬Y1-X2£¬Y2È·¶¨£¬²¢ÓÉX3£¬Y3È·¶¨µÄÍÖÔ²ÃèÊöÔ²½Ç»¡¶È
' ÓÃ¸Ãº¯Êý´´½¨µÄÇøÓòÓëÓÃRoundRect APIº¯Êý»­µÄÔ²½Ç¾ØÐÎ²»ÍêÈ«ÏàÍ¬£¬ÒòÎª±¾¾ØÐÎµÄÓÒ±ßºÍÏÂ±ß²»°üÀ¨ÔÚÇøÓòÖ®ÄÚ
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' ÓÃÖ¸¶¨Ë¢×ÓÌî³äÖ¸¶¨ÇøÓò
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
' ÓÃÖ¸¶¨Ë¢×ÓÎ§ÈÆÖ¸¶¨ÇøÓò»­Ò»¸öÍâ¿ò
Public Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
' ÕâÊÇÄÇÐ©ºÜÄÑÓÐÈË×¢Òâµ½µÄ¶Ô±à³ÌÕßÀ´ËµÊÇ¸ö¾Þ´óµÄ±¦²ØµÄÒþº¬µÄAPIº¯ÊýÖÐµÄÒ»¸ö¡£±¾º¯ÊýÔÊÐíÄú¸Ä±ä´°¿ÚµÄÇøÓò¡£
' Í¨³£ËùÓÐ´°¿Ú¶¼ÊÇ¾ØÐÎµÄ¡ª¡ª´°¿ÚÒ»µ©´æÔÚ¾Íº¬ÓÐÒ»¸ö¾ØÐÎÇøÓò¡£±¾º¯ÊýÔÊÐíÄú·ÅÆú¸ÃÇøÓò¡£
' ÕâÒâÎ¶×ÅÄú¿ÉÒÔ´´½¨Ô²µÄ¡¢ÐÇÐÎµÄ´°¿Ú£¬Ò²¿ÉÒÔ½«Ëü·ÖÎªÁ½¸ö»òÐí¶à²¿·Ö¡ª¡ªÊµ¼ÊÉÏ¿ÉÒÔÊÇÈÎºÎÐÎ×´
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
' ¸Ãº¯ÊýÑ¡ÔñÒ»¸öÇøÓò×÷ÎªÖ¸¶¨Éè±¸»·¾³µÄµ±Ç°¼ôÇÐÇøÓò
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------------Î»Í¼º¯Êý(Bitmap)-----------------------------------©§
'
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' ´´½¨Ò»·ùÓëÉè±¸ÓÐ¹ØÎ»Í¼
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§----------------------------------Í¼±êº¯Êý(Icon)------------------------------------©§
'
' ÖÆ×÷Ö¸¶¨Í¼±ê»òÊó±êÖ¸ÕëµÄÒ»¸ö¸±±¾¡£Õâ¸ö¸±±¾´ÓÊôÓÚ·¢³öµ÷ÓÃµÄÓ¦ÓÃ³ÌÐò
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
' ´´½¨Ò»¸öÍ¼±ê
Public Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
' ¸Ãº¯ÊýÇå³ýÍ¼±êºÍÊÍ·ÅÈÎºÎ±»Í¼±êÕ¼ÓÃµÄ´æ´¢¿Õ¼ä¡£
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
' ¸Ãº¯ÊýÔÚÏÞ¶¨µÄÉè±¸ÉÏÏÂÎÄ´°¿ÚµÄ¿Í»§ÇøÓò»æÖÆÍ¼±ê
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' ¸Ãº¯ÊýÔÚÏÞ¶¨µÄÉè±¸ÉÏÏÂÎÄ´°¿ÚµÄ¿Í»§ÇøÓò»æÖÆÍ¼±ê£¬Ö´ÐÐÏÞ¶¨µÄ¹âÕ¤²Ù×÷£¬²¢°´ÌØ¶¨ÒªÇóÉì³¤»òÑ¹ËõÍ¼±ê»ò¹â±ê¡£
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
' È¡µÃÓëÍ¼±êÓÐ¹ØµÄÐÅÏ¢
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------------¹â±êº¯Êý(Cursor)-----------------------------------©§
'
Public Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
' ´ÓÖ¸¶¨µÄÄ£¿é»òÓ¦ÓÃ³ÌÐòÊµÀýÖÐÔØÈëÒ»¸öÊó±êÖ¸Õë¡£LoadCursorBynumÊÇLoadCursorº¯ÊýµÄÀàÐÍ°²È«ÉùÃ÷
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
' ¸Ãº¯ÊýÏú»ÙÒ»¸ö¹â±ê²¢ÊÍ·ÅËüÕ¼ÓÃµÄÈÎºÎÄÚ´æ£¬²»ÒªÊ¹ÓÃ¸Ãº¯ÊýÈ¥Ïû»ÙÒ»¸ö¹²Ïí¹â±ê¡£
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
' »ñÈ¡Êó±êÖ¸ÕëµÄµ±Ç°Î»ÖÃ
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' ¸Ãº¯Êý°Ñ¹â±êÒÆµ½ÆÁÄ»µÄÖ¸¶¨Î»ÖÃ¡£Èç¹ûÐÂÎ»ÖÃ²»ÔÚÓÉ ClipCursorº¯ÊýÉèÖÃµÄÆÁÄ»¾ØÐÎÇøÓòÖ®ÄÚ£¬
' ÔòÏµÍ³×Ô¶¯µ÷Õû×ø±ê£¬Ê¹µÃ¹â±êÔÚ¾ØÐÎÖ®ÄÚ¡£
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§-----------------------------±ÊË¢º¯Êý(Pen and Brush)---------------------------------©§
'
' ÓÃÖ¸¶¨µÄÑùÊ½¡¢¿í¶ÈºÍÑÕÉ«´´½¨Ò»¸ö»­±Ê£¬ÓÃDeleteObjectº¯Êý½«ÆäÉ¾³ý
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' ¸ù¾ÝÖ¸¶¨µÄLOGPEN½á¹¹´´½¨Ò»¸ö»­±Ê
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
' ´´½¨Ò»¸öÀ©Õ¹»­±Ê£¨×°ÊÎ»ò¼¸ºÎ£©
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
' ÔÚÒ»¸öLOGBRUSHÊý¾Ý½á¹¹µÄ»ù´¡ÉÏ´´½¨Ò»¸öË¢×Ó
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
' ¸Ãº¯Êý¿ÉÒÔ´´½¨Ò»¸ö¾ßÓÐÖ¸¶¨ÒõÓ°Ä£Ê½ºÍÑÕÉ«µÄÂß¼­Ë¢×Ó¡£
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
' ¸Ãº¯Êý¿ÉÒÔ´´½¨¾ßÓÐÖ¸¶¨Î»Í¼Ä£Ê½µÄÂß¼­Ë¢×Ó£¬¸ÃÎ»Í¼²»ÄÜÊÇDIBÀàÐÍµÄÎ»Í¼£¬DIBÎ»Í¼ÊÇÓÉCreateDIBSectionº¯Êý´´½¨µÄ¡£
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
' ÓÃ´¿É«´´½¨Ò»¸öË¢×Ó£¬Ò»µ©Ë¢×Ó²»ÔÙÐèÒª£¬¾ÍÓÃDeleteObjectº¯Êý½«ÆäÉ¾³ý
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' ÎªÈÎºÎÒ»ÖÖ±ê×¼ÏµÍ³ÑÕÉ«È¡µÃÒ»¸öË¢×Ó£¬²»ÒªÓÃDeleteObjectº¯ÊýÉ¾³ýÕâÐ©Ë¢×Ó¡£
' ËüÃÇÊÇÓÉÏµÍ³ÓµÓÐµÄ¹ÌÓÐ¶ÔÏó¡£²»Òª½«ÕâÐ©Ë¢×ÓÖ¸¶¨³ÉÒ»ÖÖ´°¿ÚÀàµÄÄ¬ÈÏË¢×Ó
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§---------------------------×ÖÌåºÍÕýÎÄº¯Êý(Font and Text)-----------------------------©§
'
' ÓÃÖ¸¶¨µÄÊôÐÔ´´½¨Ò»ÖÖÂß¼­×ÖÌå£¬VBµÄ×ÖÌåÊôÐÔÔÚÑ¡Ôñ×ÖÌåµÄÊ±ºòÏÔµÃ¸üÓÐÐ§
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
' ½«ÎÄ±¾Ãè»æµ½Ö¸¶¨µÄ¾ØÐÎÖÐ£¬wFormat±êÖ¾³£Êý²Î¿´KhanDrawTextStyles
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
' ¸Ãº¯ÊýÈ¡µÃÖ¸¶¨Éè±¸»·¾³µÄµ±Ç°ÕýÎÄÑÕÉ«¡£
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
' ÉèÖÃµ±Ç°ÎÄ±¾ÑÕÉ«¡£ÕâÖÖÑÕÉ«Ò²³ÆÎª¡°Ç°¾°É«¡±£¬Èç¸Ä±äÁËÕâ¸öÉèÖÃ£¬×¢Òâ»Ö¸´VB´°Ìå»ò¿Ø¼þÔ­Ê¼µÄÎÄ±¾ÑÕÉ«
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§------------------------------------»æÍ¼º¯Êý----------------------------------------©§
'
' ¸Ãº¯Êý»­Ò»¶ÎÔ²»¡£¬Ô²»¡ÊÇÓÉÒ»¸öÍÖÔ²ºÍÒ»ÌõÏß¶Î£¨³ÆÖ®Îª¸îÏß£©Ïà½»ÏÞ¶¨µÄ±ÕºÏÇøÓò¡£
' ´Ë»¡ÓÉµ±Ç°µÄ»­±Ê»­ÂÖÀª£¬ÓÉµ±Ç°µÄ»­Ë¢Ìî³ä¡£
Public Declare Function Chord Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' ÓÃÖ¸¶¨µÄÑùÊ½Ãè»æÒ»¸ö¾ØÐÎµÄ±ß¿ò¡£ÀûÓÃÕâ¸öº¯Êý£¬ÎÒÃÇÃ»ÓÐ±ØÒªÔÙÊ¹ÓÃÐí¶à3D±ß¿òºÍÃæ°å¡£
' ËùÒÔ¾Í×ÊÔ´ºÍÄÚ´æµÄÕ¼ÓÃÂÊÀ´Ëµ£¬Õâ¸öº¯ÊýµÄÐ§ÂÊÒª¸ßµÃ¶à¡£Ëü¿ÉÔÚÒ»¶¨³Ì¶ÈÉÏÌáÉýÐÔÄÜ
' hdc      ÒªÔÚÆäÖÐ»æÍ¼µÄÉè±¸³¡¾°
' qrc      ÒªÎªÆäÃè»æ±ß¿òµÄ¾ØÐÎ
' edge     ´øÓÐÇ°×ºBDR_µÄÁ½¸ö³£ÊýµÄ×éºÏ¡£Ò»¸öÖ¸¶¨ÄÚ²¿±ß¿òÊÇÉÏÍ¹»¹ÊÇÏÂ°¼£»ÁíÒ»¸öÔòÖ¸¶¨Íâ²¿±ß¿ò¡£ÓÐÊ±ÄÜ»»ÓÃ´øEDGE_Ç°×ºµÄ³£Êý¡£
' grfFlags ´øÓÐBF_Ç°×ºµÄ³£ÊýµÄ×éºÏ
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
' »­Ò»¸ö½¹µã¾ØÐÎ¡£Õâ¸ö¾ØÐÎÊÇÔÚ±êÖ¾½¹µãµÄÑùÊ½ÖÐÍ¨¹ýÒì»òÔËËãÍê³ÉµÄ£¨½¹µãÍ¨³£ÓÃÒ»¸öµãÏß±íÊ¾£©
' ÈçÓÃÍ¬ÑùµÄ²ÎÊýÔÙ´Îµ÷ÓÃÕâ¸öº¯Êý£¬¾Í±íÊ¾É¾³ý½¹µã¾ØÐÎ
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
' Õâ¸öº¯ÊýÓÃÓÚÃè»æÒ»¸ö±ê×¼¿Ø¼þ
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
' Õâ¸öº¯Êý¿ÉÎªÒ»·ùÍ¼Ïó»ò»æÍ¼²Ù×÷Ó¦ÓÃ¸÷Ê½¸÷ÑùµÄÐ§¹û
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
' ¸Ãº¯ÊýÓÃÓÚ»­Ò»¸öÍÖÔ²£¬ÍÖÔ²µÄÖÐÐÄÊÇÏÞ¶¨¾ØÐÎµÄÖÐÐÄ£¬Ê¹ÓÃµ±Ç°»­±Ê»­ÍÖÔ²£¬ÓÃµ±Ç°µÄ»­Ë¢Ìî³äÍÖÔ²¡£
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ÓÃÖ¸¶¨µÄË¢×ÓÌî³äÒ»¸ö¾ØÐÎ£¬¾ØÐÎµÄÓÒ±ßºÍµ×±ß²»»áÃè»æ
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' ÓÃÖ¸¶¨µÄË¢×ÓÎ§ÈÆÒ»¸ö¾ØÐÎ»­Ò»¸ö±ß¿ò£¨×é³ÉÒ»¸öÖ¡£©£¬±ß¿òµÄ¿í¶ÈÊÇÒ»¸öÂß¼­µ¥Î»
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' È¡µÃÖ¸¶¨Éè±¸³¡¾°µ±Ç°µÄ±³¾°ÑÕÉ«
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
' Õë¶ÔÖ¸¶¨µÄÉè±¸³¡¾°£¬È¡µÃµ±Ç°µÄ±³¾°Ìî³äÄ£Ê½
Public Declare Function GetBkMode Lib "gdi32" (ByVal hDC As Long) As Long
' ÎªÖ¸¶¨µÄÉè±¸³¡¾°ÉèÖÃ±³¾°ÑÕÉ«¡£±³¾°ÑÕÉ«ÓÃÓÚÌî³äÒõÓ°Ë¢×Ó¡¢ÐéÏß»­±ÊÒÔ¼°×Ö·û£¨Èç±³¾°Ä£Ê½ÎªOPAQUE£©ÖÐµÄ¿ÕÏ¶¡£
' Ò²ÔÚÎ»Í¼ÑÕÉ«×ª»»ÆÚ¼äÊ¹ÓÃ¡£±³¾°Êµ¼ÊÊÇÉè±¸ÄÜ¹»ÏÔÊ¾µÄ×î½Ó½üÓÚ crColor µÄÑÕÉ«
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
' Ö¸¶¨ÒõÓ°Ë¢×Ó¡¢ÐéÏß»­±ÊÒÔ¼°×Ö·ûÖÐµÄ¿ÕÏ¶µÄÌî³ä·½Ê½£¬±³¾°Ä£Ê½²»»áÓ°ÏìÓÃÀ©Õ¹»­±ÊÃè»æµÄÏßÌõ
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
' ÔÚÖ¸¶¨µÄÉè±¸³¡¾°ÖÐÈ¡µÃÒ»¸öÏñËØµÄRGBÖµ
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
' ÔÚÖ¸¶¨µÄÉè±¸³¡¾°ÖÐÉèÖÃÒ»¸öÏñËØµÄRGBÖµ
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
' ½«À´×ÔÒ»·ùÎ»Í¼µÄ¶þ½øÖÆÎ»¸´ÖÆµ½Ò»·ùÓëÉè±¸ÎÞ¹ØµÄÎ»Í¼Àï
'Public Declare Function GetDIBits Lib "gdi32" ( ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' ½«À´×ÔÓëÉè±¸ÎÞ¹ØÎ»Í¼µÄ¶þ½øÖÆÎ»¸´ÖÆµ½Ò»·ùÓëÉè±¸ÓÐ¹ØµÄÎ»Í¼Àï
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' Õë¶ÔÖ¸¶¨µÄÉè±¸³¡¾°£¬»ñµÃ¶à±ßÐÎÌî³äÄ£Ê½¡£
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hDC As Long) As Long
' ÉèÖÃ¶à±ßÐÎµÄÌî³äÄ£Ê½
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
' Õë¶ÔÖ¸¶¨µÄÉè±¸³¡¾°£¬È¡µÃµ±Ç°µÄ»æÍ¼Ä£Ê½¡£ÕâÑù¿É¶¨Òå»æÍ¼²Ù×÷ÈçºÎÓëÕýÔÚÏÔÊ¾µÄÍ¼ÏóºÏ²¢ÆðÀ´
' Õâ¸öº¯ÊýÖ»¶Ô¹âÕ¤Éè±¸ÓÐÐ§
Public Declare Function GetROP2 Lib "gdi32" (ByVal hDC As Long) As Long
' ÉèÖÃÖ¸¶¨Éè±¸³¡¾°µÄ»æÍ¼Ä£Ê½¡£
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
' ÓÃµ±Ç°»­±Ê»­Ò»ÌõÏß£¬´Óµ±Ç°Î»ÖÃÁ¬µ½Ò»¸öÖ¸¶¨µÄµã¡£Õâ¸öº¯Êýµ÷ÓÃÍê±Ï£¬µ±Ç°Î»ÖÃ±ä³Éx,yµã
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
' ÎªÖ¸¶¨µÄÉè±¸³¡¾°Ö¸¶¨Ò»¸öÐÂµÄµ±Ç°»­±ÊÎ»ÖÃ¡£Ç°Ò»¸öÎ»ÖÃ±£´æÔÚlpPointÖÐ
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
' ¸Ãº¯Êý»­Ò»¸öÓÉÍÖÔ²ºÍÁ½Ìõ°ë¾¶Ïà½»±ÕºÏ¶ø³ÉµÄ±ý×´Ð¨ÐÎÍ¼£¬´Ë±ýÍ¼ÓÉµ±Ç°»­±Ê»­ÂÖÀª£¬ÓÉµ±Ç°»­Ë¢Ìî³ä¡£
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' ¸Ãº¯Êý»­Ò»¸öÓÉÖ±ÏßÏàÎÅµÄÁ½¸öÒÔÉÏ¶¥µã×é³ÉµÄ¶à±ßÐÎ£¬ÓÃµ±Ç°»­±Ê»­¶à±ßÐÎÂÖÀª£¬
' ÓÃµ±Ç°»­Ë¢ºÍ¶à±ßÐÎÌî³äÄ£Ê½Ìî³ä¶à±ßÐÎ¡£
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
' ÓÃµ±Ç°»­±ÊÃè»æÒ»ÏµÁÐÏß¶Î¡£Ê¹ÓÃPolylineToº¯ÊýÊ±£¬µ±Ç°Î»ÖÃ»áÉèÎª×îºóÒ»ÌõÏß¶ÎµÄÖÕµã¡£
' Ëü²»»áÓÉPolylineº¯Êý¸Ä¶¯
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PolyPolyline Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
' ¸Ãº¯Êý»­Ò»¸ö¾ØÐÎ£¬ÓÃµ±Ç°µÄ»­±Ê»­¾ØÐÎÂÖÀª£¬ÓÃµ±Ç°»­Ë¢½øÐÐÌî³ä¡£
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' º¯Êý»­Ò»¸ö´øÔ²½ÇµÄ¾ØÐÎ£¬´Ë¾ØÐÎÓÉµ±Ç°»­±Ê»­ÂÖÀÈ£¬ÓÉµ±Ç°»­Ë¢Ìî³ä¡£
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Õâ¸öº¯ÊýÓÃÓÚÔö´ó»ò¼õÐ¡Ò»¸ö¾ØÐÎµÄ´óÐ¡¡£
' x¼ÓÔÚÓÒ²àÇøÓò£¬²¢´Ó×ó²àÇøÓò¼õÈ¥£»ÈçxÎªÕý£¬ÔòÄÜÔö´ó¾ØÐÎµÄ¿í¶È£»ÈçxÎª¸º£¬ÔòÄÜ¼õÐ¡Ëü¡£
' y¶Ô¶¥²¿Óëµ×²¿ÇøÓò²úÉúµÄÓ°ÏìÊÇÊÇÀàËÆµÄ
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' ¸Ãº¯ÊýÍ¨¹ýÓ¦ÓÃÒ»¸öÖ¸¶¨µÄÆ«ÒÆ£¬´Ó¶øÈÃ¾ØÐÎÒÆ¶¯ÆðÀ´¡£
' x»áÌí¼Óµ½ÓÒ²àºÍ×ó²àÇøÓò¡£yÌí¼Óµ½¶¥²¿ºÍµ×²¿ÇøÓò¡£
' Æ«ÒÆ·½ÏòÔòÈ¡¾öÓÚ²ÎÊýÊÇÕýÊý»¹ÊÇ¸ºÊý£¬ÒÔ¼°²ÉÓÃµÄÊÇÊ²Ã´×ø±êÏµÍ³
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' ·µ»ØÓëwindows»·¾³ÓÐ¹ØµÄÐÅÏ¢£¬nIndexÖµ²Î¿´±¾Ä£¿éµÄ³£Á¿ÉùÃ÷
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
' »ñµÃÕû¸ö´°¿ÚµÄ·¶Î§¾ØÐÎ£¬´°¿ÚµÄ±ß¿ò¡¢±êÌâÀ¸¡¢¹ö¶¯Ìõ¼°²Ëµ¥µÈ¶¼ÔÚÕâ¸ö¾ØÐÎÄÚ
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' ·µ»ØÖ¸¶¨´°¿Ú¿Í»§Çø¾ØÐÎµÄ´óÐ¡
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' Õâ¸öº¯ÊýÆÁ±ÎÒ»¸ö´°¿Ú¿Í»§ÇøµÄÈ«²¿»ò²¿·ÖÇøÓò¡£Õâ»áµ¼ÖÂ´°¿ÚÔÚÊÂ¼þÆÚ¼ä²¿·ÖÖØ»­
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
' ÅÐ¶ÏÖ¸¶¨windowsÏÔÊ¾¶ÔÏóµÄÑÕÉ«£¬ÑÕÉ«¶ÔÏó¿´±¾Ä£¿éÉùÃ÷
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

'©³©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©·
'©§--------------------------------ÆäËûº¯Êý(Others)------------------------------------©§
'
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
'©§                                                                                    ©§
'©»©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¥©¿

