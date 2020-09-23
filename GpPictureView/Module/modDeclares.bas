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


' 指定窗口的结构中取得信息，用于GetWindowLong、SetWindowLong函数
Public Const GWL_EXSTYLE = (-20)                 '/* 扩展窗口样式 */
Public Const GWL_HINSTANCE = (-6)                '/* 拥有窗口的实例的句柄 */
Public Const GWL_HWNDPARENT = (-8)               '/* 该窗口之父的句柄。不要用SetWindowWord来改变这个值 */
Public Const GWL_ID = (-12)                      '/* 对话框中一个子窗口的标识符 */
Public Const GWL_STYLE = (-16)                   '/* 窗口样式 */
Public Const GWL_USERDATA = (-21)                '/* 含义由应用程序规定 */
Public Const GWL_WNDPROC = (-4)                  '/* 该窗口的窗口函数的地址 */
Public Const DWL_DLGPROC = 4                     '/* 这个窗口的对话框函数地址 */
Public Const DWL_MSGRESULT = 0                   '/* 在对话框函数中处理的一条消息返回的值 */
Public Const DWL_USER = 8                        '/* 含义由应用程序规定 */


' GetDeviceCaps索引表，用于GetDeviceCaps函数
Public Const DRIVERVERSION = 0                   '/* 备驱动程序版本
Public Const BITSPIXEL = 12                      '/*
Public Const LOGPIXELSX = 88                     '/*  Logical pixels/inch in X
Public Const LOGPIXELSY = 90                     '/*  Logical pixels/inch in Y

' Windows对象常数表，函数GetSysColor
Public Const COLOR_ACTIVEBORDER = 10             '/* 活动窗口的边框
Public Const COLOR_ACTIVECAPTION = 2             '/* 活动窗口的标题
Public Const COLOR_ADJ_MAX = 100                 '/*
Public Const COLOR_ADJ_MIN = -100                '/*
Public Const COLOR_APPWORKSPACE = 12             '/* MDI桌面的背景
Public Const COLOR_BACKGROUND = 1                '/*
Public Const COLOR_BTNDKSHADOW = 21              '/*
Public Const COLOR_BTNLIGHT = 22                 '/*
Public Const COLOR_BTNFACE = 15                  '/* 按钮
Public Const COLOR_BTNHIGHLIGHT = 20             '/* 按钮的3D加亮区
Public Const COLOR_BTNSHADOW = 16                '/* 按钮的3D阴影
Public Const COLOR_BTNTEXT = 18                  '/* 按钮文字
Public Const COLOR_CAPTIONTEXT = 9               '/* 窗口标题中的文字
Public Const COLOR_GRAYTEXT = 17                 '/* 灰色文字；如使用了抖动技术则为零
Public Const COLOR_HIGHLIGHT = 13                '/* 选定的项目背景
Public Const COLOR_HIGHLIGHTTEXT = 14            '/* 选定的项目文字
Public Const COLOR_INACTIVEBORDER = 11           '/* 不活动窗口的边框
Public Const COLOR_INACTIVECAPTION = 3           '/* 不活动窗口的标题
Public Const COLOR_INACTIVECAPTIONTEXT = 19      '/* 不活动窗口的文字
Public Const COLOR_MENU = 4                      '/* 菜单
Public Const COLOR_MENUTEXT = 7                  '/* 菜单正文
Public Const COLOR_SCROLLBAR = 0                 '/* 滚动条
Public Const COLOR_WINDOW = 5                    '/* 窗口背景
Public Const COLOR_WINDOWFRAME = 6               '/* 窗框
Public Const COLOR_WINDOWTEXT = 8                '/* 窗口正文
Public Const COLORONCOLOR = 3

' 函数CombineRgn的返回值，类型Long
Public Const COMPLEXREGION = 3                   '/* 区域有互相交叠的边界 */
Public Const SIMPLEREGION = 2                    '/* 区域边界没有互相交叠 */
Public Const NULLREGION = 1                      '/* 区域为空 */
Public Const ERRORAPI = 0                        '/* 不能创建组合区域 */

' 组合两区域的方法，函数CombineRgn的的参数nCombineMode所使用的常数
Public Const RGN_AND = 1                         '/* hDestRgn被设置为两个源区域的交集 */
Public Const RGN_COPY = 5                        '/* hDestRgn被设置为hSrcRgn1的拷贝 */
Public Const RGN_DIFF = 4                        '/* hDestRgn被设置为hSrcRgn1中与hSrcRgn2不相交的部分 */
Public Const RGN_OR = 2                          '/* hDestRgn被设置为两个区域的并集 */
Public Const RGN_XOR = 3                         '/* hDestRgn被设置为除两个源区域OR之外的部分 */

' Missing Draw State constants declarations，参看DrawState函数
'/* Image type */
Public Const DST_COMPLEX = &H0                   '/* 绘图在由lpDrawStateProc参数指定的回调函数期间执行。lParam和wParam会传递给回调事件
Public Const DST_TEXT = &H1                      '/* lParam代表文字的地址（可使用一个字串别名），wParam代表字串的长度
Public Const DST_PREFIXTEXT = &H2                '/* 与DST_TEXT类似，只是 & 字符指出为下各字符加上下划线
Public Const DST_ICON = &H3                      '/* lParam包括图标句柄
Public Const DST_BITMAP = &H4                    '/* lParam中的句柄
' /* State type */
Public Const DSS_NORMAL = &H0                    '/* 普通图象
Public Const DSS_UNION = &H10                    '/* 图象进行抖动处理
Public Const DSS_DISABLED = &H20                 '/* 图象具有浮雕效果
Public Const DSS_MONO = &H80                     '/* 用hBrush描绘图象
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

' 常见的光栅操作代码
Public Const BLACKNESS = &H42                    '/* 表示使用与物理调色板的索引0相关的色彩来填充目标矩形区域，（对缺省的物理调色板而言，该颜色为黑色）。
Public Const DSTINVERT = &H550009                '/* 表示使目标矩形区域颜色取反。
Public Const MERGECOPY = &HC000CA                '/* 表示使用布尔型的AND（与）操作符将源矩形区域的颜色与特定模式组合一起。
Public Const MERGEPAINT = &HBB0226               '/* 通过使用布尔型的OR（或）操作符将反向的源矩形区域的颜色与目标矩形区域的颜色合并。
Public Const NOTSRCCOPY = &H330008               '/* 将源矩形区域颜色取反，于拷贝到目标矩形区域。
Public Const NOTSRCERASE = &H1100A6              '/* 使用布尔类型的OR（或）操作符组合源和目标矩形区域的颜色值，然后将合成的颜色取反。
Public Const PATCOPY = &HF00021                  '/* 将特定的模式拷贝到目标位图上。
Public Const PATINVERT = &H5A0049                '/* 通过使用布尔OR（或）操作符将源矩形区域取反后的颜色值与特定模式的颜色合并。然后使用OR（或）操作符将该操作的结果与目标矩形区域内的颜色合并。
Public Const PATPAINT = &HFB0A09                 '/* 通过使用XOR（异或）操作符将源和目标矩形区域内的颜色合并。
Public Const SRCAND = &H8800C6                   '/* 通过使用AND（与）操作符来将源和目标矩形区域内的颜色合并
Public Const SRCCOPY = &HCC0020                  '/* 将源矩形区域直接拷贝到目标矩形区域。
Public Const SRCERASE = &H440328                 '/* 通过使用AND（与）操作符将目标矩形区域颜色取反后与源矩形区域的颜色值合并。
Public Const SRCINVERT = &H660046                '/* 通过使用布尔型的XOR（异或）操作符将源和目标矩形区域的颜色合并。
Public Const SRCPAINT = &HEE0086                 '/* 通过使用布尔型的OR（或）操作符将源和目标矩形区域的颜色合并。
Public Const WHITENESS = &HFF0062                '/* 使用与物理调色板中索引1有关的颜色填充目标矩形区域。（对于缺省物理调色板来说，这个颜色就是白色）。

'--- for mouse_event
Public Const MOUSE_MOVED = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000       '/*
Public Const MOUSEEVENTF_LEFTDOWN = &H2          '/* 模拟鼠标左键按下
Public Const MOUSEEVENTF_LEFTUP = &H4            '/* 模拟鼠标左键抬起
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20       '/* 模拟鼠标中键按下
Public Const MOUSEEVENTF_MIDDLEUP = &H40         '/* 模拟鼠标中键按下
Public Const MOUSEEVENTF_MOVE = &H1              '/* 移动鼠标 */
Public Const MOUSEEVENTF_RIGHTDOWN = &H8         '/* 模拟鼠标右键按下
Public Const MOUSEEVENTF_RIGHTUP = &H10          '/* 模拟鼠标右键按下
Public Const MOUSETRAILS = 39                    '/*

Public Const BMP_MAGIC_COOKIE = 19778            '/* this is equivalent to ascii string "BM" */
' constants for the biCompression field
Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const BI_BITFIELDS = 3&
'Public Const BITSPIXEL = 12                     '/* Number of bits per pixel
' DIB color table identifiers
Public Const DIB_PAL_COLORS = 1                  '/* 在颜色表中装载一个16位所以数组，它们与当前选定的调色板有关 color table in palette indices
Public Const DIB_PAL_INDICES = 2                 '/* No color table indices into surf palette
Public Const DIB_PAL_LOGINDICES = 4              '/* No color table indices into DC palette
Public Const DIB_PAL_PHYSINDICES = 2             '/* No color table indices into surf palette
Public Const DIB_RGB_COLORS = 0                  '/* 在颜色表中装载RGB颜色

' ======================================================================================
' Methods
' ======================================================================================
' 函数SetBkModen参数BkMode
Public Enum KhanBackStyles
    TRANSPARENT = 1                              '/* 透明处理，即不作上述填充 */
    OPAQUE = 2                                   '/* 用当前的背景色填充虚线画笔、阴影刷子以及字符的空隙 */
    NEWTRANSPARENT = 3                           '/* NT4: Uses chroma-keying upon BitBlt. Undocumented feature that is not working on Windows 2000/XP.
End Enum

' 多边形的填充模式
Public Enum KhanPolyFillModeFalgs
    ALTERNATE = 1                                '/* 交替填充
    WINDING = 2                                  '/* 根据绘图方向填充
End Enum

' DrawIconEx
Public Enum KhanDrawIconExFlags
    DI_MASK = &H1                                '/* 绘图时使用图标的MASK部分（如单独使用，可获得图标的掩模）
    DI_IMAGE = &H2                               '/* 绘图时使用图标的XOR部分（即图标没有透明区域）
    DI_NORMAL = &H3                              '/* 用常规方式绘图（合并 DI_IMAGE 和 DI_MASK）
    DI_COMPAT = &H4                              '/* 描绘标准的系统指针，而不是指定的图象
    DI_DEFAULTSIZE = &H8                         '/* 忽略cxWidth和cyWidth设置，并采用原始的图标大小
End Enum

'指定被装载图像类型,LoadImage,CopyImage
Public Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

Public Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/* 表示创建一个图像的精确副本，而忽略参数cxDesired和cyDesired
    LR_COPYDELETEORG = &H8                       '/* 表示创建一个副本后删除原始图像。
    LR_CREATEDIBSECTION = &H2000                 '/* 当参数uType指定为IMAGE_BITMAP时，使得函数返回一个DIB部分位图，而不是一个兼容的位图。这个标志在装载一个位图，而不是映射它的颜色到显示设备时非常有用。
    LR_DEFAULTCOLOR = &H0                        '/* 以常规方式载入图象
    LR_DEFAULTSIZE = &H40                        '/* 若 cxDesired或cyDesired未被设为零，使用系统指定的公制值标识光标或图标的宽和高。如果这个参数不被设置且cxDesired或cyDesired被设为零，函数使用实际资源尺寸。如果资源包含多个图像，则使用第一个图像的大小。
    LR_LOADFROMFILE = &H10                       '/* 根据参数lpszName的值装载图像。若标记未被给定，lpszName的值为资源名称。
    LR_LOADMAP3DCOLORS = &H1000                  '/* 将图象中的深灰(Dk Gray RGB（128，128，128）)、灰(Gray RGB（192，192，192）)、以及浅灰(Gray RGB（223，223，223）)像素都替换成COLOR_3DSHADOW，COLOR_3DFACE以及COLOR_3DLIGHT的当前设置
    LR_LOADTRANSPARENT = &H20                    '/* 若fuLoad包括LR_LOADTRANSPARENT和LR_LOADMAP3DCOLORS两个值，则LRLOADTRANSPARENT优先。但是，颜色表接口由COLOR_3DFACE替代，而不是COLOR_WINDOW。
    LR_MONOCHROME = &H1                          '/* 将图象转换成单色
    LR_SHARED = &H8000                           '/* 若图像将被多次装载则共享。如果LR_SHARED未被设置，则再向同一个资源第二次调用这个图像是就会再装载以便这个图像且返回不同的句柄。
    LR_COPYFROMRESOURCE = &H4000                 '/*
End Enum

Public Enum KhanDrawTextStyles
    DT_BOTTOM = &H8&                             '/* 必须同时指定DT_SINGLE。指示文本对齐格式化矩形的底边
    DT_CALCRECT = &H400&                         '/* 象下面这样计算格式化矩形：多行绘图时矩形的底边根据需要进行延展，以便容下所有文字；单行绘图时，延展矩形的右侧。不描绘文字。由lpRect参数指定的矩形会载入计算出来的值
    DT_CENTER = &H1&                             '/* 文本垂直居中
    DT_EXPANDTABS = &H40&                        '/* 描绘文字的时候，对制表站进行扩展。默认的制表站间距是8个字符。但是，可用DT_TABSTOP标志改变这项设定
    DT_EXTERNALLEADING = &H200&                  '/* 计算文本行高度的时候，使用当前字体的外部间距属性（the external leading attribute）
    DT_INTERNAL = &H1000&                        '/* Uses the system font to calculate text metrics
    DT_LEFT = &H0&                               '/* 文本左对齐
    DT_NOCLIP = &H100&                           '/* 描绘文字时不剪切到指定的矩形，DrawTextEx is somewhat faster when DT_NOCLIP is used.
    DT_NOPREFIX = &H800&                         '/* 通常，函数认为 & 字符表示应为下一个字符加上下划线。该标志禁止这种行为
    DT_RIGHT = &H2&                              '/* 文本右对齐
    DT_SINGLELINE = &H20&                        '/* 只画单行
    DT_TABSTOP = &H80&                           '/* 指定新的制表站间距，采用这个整数的高8位
    DT_TOP = &H0&                                '/* 必须同时指定DT_SINGLE。指示文本对齐格式化矩形的底边
    DT_VCENTER = &H4&                            '/* 必须同时指定DT_SINGLE。指示文本对齐格式化矩形的中部
    DT_WORDBREAK = &H10&                         '/* 进行自动换行。如用SetTextAlign函数设置了TA_UPDATECP标志，这里的设置则无效
' #if(WINVER >= =&H0400)
    DT_EDITCONTROL = &H2000&                     '/* 对一个多行编辑控件进行模拟。不显示部分可见的行
    DT_END_ELLIPSIS = &H8000&                    '/* 倘若字串不能在矩形里全部容下，就在末尾显示省略号
    DT_PATH_ELLIPSIS = &H4000&                   '/* 如字串包含了 \ 字符，就用省略号替换字串内容，使其能在矩形中全部容下。例如，一个很长的路径名可能换成这样显示——c:\windows\...\doc\readme.txt
    DT_MODIFYSTRING = &H10000                    '/* 如指定了DT_ENDELLIPSES 或 DT_PATHELLIPSES，就会对字串进行修改，使其与实际显示的字串相符
    DT_RTLREADING = &H20000                      '/* 如选入设备场景的字体属于希伯来或阿拉伯语系，就从右到左描绘文字
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

' 指定画笔样式，函数CreatePen的参数CreatePen所使用的常数
Public Enum KhanPenStyles
    ' CreatePen，ExtCreatePen
    ' 画笔的样式
    PS_SOLID = 0                                 '/* 画笔画出的是实线 */
    PS_DASH = 1                                  '/* 画笔画出的是虚线（nWidth必须是1） */
    PS_DOT = 2                                   '/* 画笔画出的是点线（nWidth必须是1） */
    PS_DASHDOT = 3                               '/* 画笔画出的是点划线（nWidth必须是1） */
    PS_DASHDOTDOT = 4                            '/* 画笔画出的是点-点-划线（nWidth必须是1） */
    PS_NULL = 5                                  '/* 画笔不能画图 */
    PS_INSIDEFRAME = 6                           '/* 画笔在由椭圆、矩形、圆角矩形、饼图以及弦等生成的封闭对象框中画图。如指定的准确RGB颜色不存在，就进行抖动处理 */
    ' ExtCreatePen
    ' 画笔的样式
    PS_USERSTYLE = 7                             '/* <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
    PS_ALTERNATE = 8                             '/* <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
    ' 画笔的笔尖
    PS_ENDCAP_ROUND = &H0                        '/* End caps are round.
    PS_ENDCAP_SQUARE = &H100                     '/* End caps are square.
    PS_ENDCAP_FLAT = &H200                       '/* End caps are flat.
    PS_ENDCAP_MASK = &HF00                       '/* Mask for previous PS_ENDCAP_XXX values.
    ' 在图形中连接线段或在路径中连接直线的方式
    PS_JOIN_ROUND = &H0                          '/* Joins are beveled.
    PS_JOIN_BEVEL = &H1000                       '/* Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
    PS_JOIN_MITER = &H2000                       '/* Joins are round.
    PS_JOIN_MASK = &HF000                        '/* Mask for previous PS_JOIN_XXX values.
    ' 画笔的类型
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

' 窗口指定一个新位置和状态，用于SetWindowPos函数
Public Enum KhanSetWindowPosStyles
    HWND_BOTTOM = 1                              '/* 将窗口置于窗口列表底部 */
    HWND_NOTOPMOST = -2                          '/* 将窗口置于列表顶部，并位于任何最顶部窗口的后面 */
    HWND_TOP = 0                                 '/* 将窗口置于Z序列的顶部；Z序列代表在分级结构中，窗口针对一个给定级别的窗口显示的顺序 */
    HWND_TOPMOST = -1                            '/* 将窗口置于列表顶部，并位于任何最顶部窗口的前面 */
    SWP_SHOWWINDOW = &H40                        '/* 显示窗口 */
    SWP_HIDEWINDOW = &H80                        '/* 隐藏窗口 */
    SWP_FRAMECHANGED = &H20                      '/* 强迫一条WM_NCCALCSIZE消息进入窗口，即使窗口的大小没有改变 */
    SWP_NOACTIVATE = &H10                        '/* 不激活窗口 */
    SWP_NOCOPYBITS = &H100                       '
    SWP_NOMOVE = &H2                             '/* 保持当前位置（x和y设定将被忽略） */
    SWP_NOOWNERZORDER = &H200                    '/* Don't do owner Z ordering */
    SWP_NOREDRAW = &H8                           '/* 窗口不自动重画 */
    SWP_NOREPOSITION = SWP_NOOWNERZORDER         '
    SWP_NOSIZE = &H1                             '/* 保持当前大小（cx和cy会被忽略） */
    SWP_NOZORDER = &H4                           '/* 保持窗口在列表的当前位置（hWndInsertAfter将被忽略） */
    SWP_DRAWFRAME = SWP_FRAMECHANGED             '/* 围绕窗口画一个框 */
'    HWND_BROADCAST = &HFFFF&
'    HWND_DESKTOP = 0
End Enum

' 指定创建窗口的风格
Public Enum KhanCreateWindowSytles
    ' CreateWindow
    WS_BORDER = &H800000                         '/* 创建一个单边框的窗口。
    WS_CAPTION = &HC00000                        '/* 创建一个有标题框的窗口（包括WS_BODER风格）。
    WS_CHILD = &H40000000                        '/* 创建一个子窗口。这个风格不能与WS_POPVP风格合用。
    WS_CHILDWINDOW = (WS_CHILD)                  '/* 与WS_CHILD相同。
    WS_CLIPCHILDREN = &H2000000                  '/* 当在父窗口内绘图时，排除子窗口区域。在创建父窗口时使用这个风格。
    WS_CLIPSIBLINGS = &H4000000                  '/* 排除子窗口之间的相对区域，也就是，当一个特定的窗口接收到WM_PAINT消息时，WS_CLIPSIBLINGS 风格将所有层叠窗口排除在绘图之外，只重绘指定的子窗口。如果未指定WS_CLIPSIBLINGS风格，并且子窗口是层叠的，则在重绘子窗口的客户区时，就会重绘邻近的子窗口。
    WS_DISABLED = &H8000000                      '/* 创建一个初始状态为禁止的子窗口。一个禁止状态的窗日不能接受来自用户的输人信息。
    WS_DLGFRAME = &H400000                       '/* 创建一个带对话框边框风格的窗口。这种风格的窗口不能带标题条。
    WS_GROUP = &H20000                           '/* 指定一组控制的第一个控制。这个控制组由第一个控制和随后定义的控制组成，自第二个控制开始每个控制，具有WS_GROUP风格，每个组的第一个控制带有WS_TABSTOP风格，从而使用户可以在组间移动。用户随后可以使用光标在组内的控制间改变键盘焦点。
    WS_HSCROLL = &H100000                        '/* 创建一个有水平滚动条的窗口。
    WS_MAXIMIZE = &H1000000                      '/* 创建一个具有最大化按钮的窗口。该风格不能与WS_EX_CONTEXTHELP风格同时出现，同时必须指定WS_SYSMENU风格。
    WS_MAXIMIZEBOX = &H10000                     '/*
    WS_MINIMIZE = &H20000000                     '/* 创建一个初始状态为最小化状态的窗口。
    WS_ICONIC = WS_MINIMIZE                      '/* 创建一个初始状态为最小化状态的窗口。与WS_MINIMIZE风格相同。
    WS_MINIMIZEBOX = &H20000                     '/*
    WS_OVERLAPPED = &H0&                         '/* 产生一个层叠的窗口。一个层叠的窗口有一个标题条和一个边框。与WS_TILED风格相同
    WS_POPUP = &H80000000                        '/* 创建一个弹出式窗口。该风格不能与WS_CHLD风格同时使用。
    WS_SYSMENU = &H80000                         '/* 创建一个在标题条上带有窗口菜单的窗口，必须同时设定WS_CAPTION风格。
    WS_TABSTOP = &H10000                         '/* 创建一个控制，这个控制在用户按下Tab键时可以获得键盘焦点。按下Tab键后使键盘焦点转移到下一具有WS_TABSTOP风格的控制。
    WS_THICKFRAME = &H40000                      '/* 创建一个具有可调边框的窗口。
    WS_SIZEBOX = WS_THICKFRAME                   '/* 与WS_THICKFRAME风格相同
    WS_TILED = WS_OVERLAPPED                     '/* 产生一个层叠的窗口。一个层叠的窗口有一个标题和一个边框。与WS_OVERLAPPED风格相同。
    WS_VISIBLE = &H10000000                      '/* 创建一个初始状态为可见的窗口。
    WS_VSCROLL = &H200000                        '/* 创建一个有垂直滚动条的窗口。
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW         '/* 创建一个具有WS_OVERLAPPED，WS_CAPTION，WS_SYSMENU MS_THICKFRAME．
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) '/* 创建一个具有WS_BORDER，WS_POPUP,WS_SYSMENU风格的窗口，WS_CAPTION和WS_POPUPWINDOW必须同时设定才能使窗口某单可见。
    ' CreateWindowEx
    WS_EX_ACCEPTFILES = &H10&                    '/* 指定以该风格创建的窗口接受一个拖拽文件。
    WS_EX_APPWINDOW = &H40000                    '/* 当窗口可见时，将一个顶层窗口放置到任务条上。
    WS_EX_CLIENTEDGE = &H200                     '/* 指定窗口有一个带阴影的边界。
    WS_EX_CONTEXTHELP = &H400                    '/* 在窗口的标题条包含一个问号标志。当用户点击了问号时，鼠标光标变为一个问号的指针、如果点击了一个子窗口，则子窗日接收到WM_HELP消息。子窗口应该将这个消息传递给父窗口过程，父窗口再通过HELP_WM_HELP命令调用WinHelp函数。这个Help应用程序显示一个包含子窗口帮助信息的弹出式窗口。 WS_EX_CONTEXTHELP不能与WS_MAXIMIZEBOX和WS_MINIMIZEBOX同时使用。
    WS_EX_CONTROLPARENT = &H10000                '/* 允许用户使用Tab键在窗口的子窗口间搜索。
    WS_EX_DLGMODALFRAME = &H1&                   '/* 创建一个带双边的窗口；该窗口可以在dwStyle中指定WS_CAPTION风格来创建一个标题栏。
    WS_EX_LEFT = &H0                             '/* 窗口具有左对齐属性，这是缺省设置的。
    WS_EX_LEFTSCROLLBAR = &H4000                 '/* 如果外壳语言是如Hebrew，Arabic，或其他支持reading order alignment的语言，则标题条（如果存在）则在客户区的左部分。若是其他语言，在该风格被忽略并且不作为错误处理。
    WS_EX_LTRREADING = &H0                       '/* 窗口文本以LEFT到RIGHT（自左向右）属性的顺序显示。这是缺省设置的。
    WS_EX_MDICHILD = &H40                        '/* 创建一个MDI子窗口。
    WS_EX_NOACTIVATE = &H8000000                 '/*
    WS_EX_NOPATARENTNOTIFY = &H4&                '/* 指明以这个风格创建的窗口在被创建和销毁时不向父窗口发送WM_PARENTNOTFY消息。
    WS_EX_OVERLAPPEDWINDOW = &H300               '/*
    WS_EX_PALETTEWINDOW = &H188                  '/* WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOW和WS_WX_TOPMOST风格的组合WS_EX_RIGHT:窗口具有普通的右对齐属性，这依赖于窗口类。只有在外壳语言是如Hebrew,Arabic或其他支持读顺序对齐（reading order alignment）的语言时该风格才有效，否则，忽略该标志并且不作为错误处理。
    WS_EX_RIGHT = &H1000                         '/*
    WS_EX_RIGHTSCROLLBAR = &H0                   '/* 垂直滚动条在窗口的右边界。这是缺省设置的。
    WS_EX_RTLREADING = &H2000                    '/* 如果外壳语言是如Hebrew，Arabic，或其他支持读顺序对齐（reading order alignment）的语言，则窗口文本是一自左向右）RIGHT到LEFT顺序的读出顺序。若是其他语言，在该风格被忽略并且不作为错误处理。
    WS_EX_STATICEDGE = &H20000                   '/* 为不接受用户输入的项创建一个3一维边界风格。
    WS_EX_TOOLWINDOW = &H80                      '/*
    WS_EX_TOPMOST = &H8&                         '/* 指明以该风格创建的窗口应放置在所有非最高层窗口的上面并且停留在其L，即使窗口未被激活。使用函数SetWindowPos来设置和移去这个风格。
    WS_EX_TRANSPARENT = &H20&                    '/* 指定以这个风格创建的窗口在窗口下的同属窗口已重画时，该窗口才可以重画。
    WS_EX_WINDOWEDGE = &H100
End Enum

' Windows环境有关的信息，用于GetSystemMetrics函数
Public Enum KhanSystemMetricsFlags
    SM_CXSCREEN = 0                              '/* 屏幕大小 */
    SM_CYSCREEN = 1                              '/* 屏幕大小 */
    SM_CXVSCROLL = 2                             '/* 垂直滚动条中的箭头按钮的大小 */
    SM_CYHSCROLL = 3                             '/* 水平滚动条上的箭头大小 */
    SM_CYCAPTION = 4                             '/* 窗口标题的高度 */
    SM_CXBORDER = 5                              '/* 尺寸不可变边框的大小 */
    SM_CYBORDER = 6                              '/* 尺寸不可变边框的大小 */
    SM_CXDLGFRAME = 7                            '/* 对话框边框的大小 */
    SM_CYDLGFRAME = 8                            '/* 对话框边框的大小 */
    SM_CYVTHUMB = 9                              '/* 滚动块在水平滚动条上的大小 */
    SM_CXHTHUMB = 10                             '/* 滚动块在水平滚动条上的大小 */
    SM_CXICON = 11                               '/* 标准图标的大小 */
    SM_CYICON = 12                               '/* 标准图标的大小 */
    SM_CXCURSOR = 13                             '/* 标准指针大小 */
    SM_CYCURSOR = 14                             '/* 标准指针大小 */
    SM_CYMENU = 15                               '/* 菜单高度 */
    SM_CXFULLSCREEN = 16                         '/* 最大化窗口客户区的大小 */
    SM_CYFULLSCREEN = 17                         '/* 最大化窗口客户区的大小 */
    SM_CYKANJIWINDOW = 18                        '/* Kanji窗口的大小（Height of Kanji window） */
    SM_MOUSEPRESENT = 19                         '/* 如安装了鼠标则为TRUE */
    SM_CYVSCROLL = 20                            '/* 垂直滚动条中的箭头按钮的大小 */
    SM_CXHSCROLL = 21                            '/* 水平滚动条上的箭头大小 */
    SM_DEBUG = 22                                '/* 如windows的调试版正在运行，则为TRUE */
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28                                '/* 窗口的最小尺寸 */
    SM_CYMIN = 29                                '/* 窗口的最小尺寸 */
    SM_CXSIZE = 30                               '/* 标题栏位图的大小 */
    SM_CYSIZE = 31                               '/* 标题栏位图的大小 */
    SM_CXFRAME = 32                              '/* 尺寸可变边框的大小（在win95和nt 4.0中使用SM_C?FIXEDFRAME） */
    SM_CYFRAME = 33                              '/* 尺寸可变边框的大小 */
    SM_CXMINTRACK = 34                           '/* 窗口的最小轨迹宽度 */
    SM_CYMINTRACK = 35                           '/* 窗口的最小轨迹宽度 */
    SM_CXDOUBLECLK = 36                          '/* 双击区域的大小（指定屏幕上一个特定的显示区域，只有在这个区域内连续进行两次鼠标单击，才有可能被当作双击事件处理） */
    SM_CYDOUBLECLK = 37                          '/* 双击区域的大小 */
    SM_CXICONSPACING = 38                        '/* 桌面图标之间的间隔距离。在win95和nt 4.0中是指大图标的间距 */
    SM_CYICONSPACING = 39                        '/* 桌面图标之间的间隔距离。在win95和nt 4.0中是指大图标的间距 */
    SM_MENUDROPALIGNMENT = 40                    '/* 如弹出式菜单对齐菜单栏项目的左侧，则为零 */
    SM_PENWINDOWS = 41                           '/* 如装载了支持笔窗口的DLL，则表示笔窗口的句柄 */
    SM_DBCSENABLED = 42                          '/* 如支持双字节则为TRUE */
    SM_CMOUSEBUTTONS = 43                        '/* 鼠标按钮（按键）的数量。如没有鼠标，就为零 */
    SM_CMETRICS = 44                             '/* 可用系统环境的数量 */
End Enum

' SetMapMode
Public Enum KhanMapModeStyles
    MM_ANISOTROPIC = 8                           '/* 逻辑单位转换成具有任意比例轴的任意单位，用SetWindowExtEx和SetViewportExtEx函数可指定单位、方向和比例。
    MM_HIENGLISH = 5                             '/* 每个逻辑单位转换为0.001inch(英寸)，X的正方面向右，Y的正方向向上
    MM_HIMETRIC = 3                              '/* 每个逻辑单位转换为0.01millimeter(毫米)，X正方向向右，Y的正方向向上。
    MM_ISOTROPIC = 7                             '/* 视口和窗口范围任意，只是x和y逻辑单元尺寸要相同
    MM_LOENGLISH = 4                             '/* 每个逻辑单位转换为英寸，X正方向向右，Y正方向向上。
    MM_LOMETRIC = 2                              '/* 每个逻辑单位转换为毫米，X正方向向右，Y正方向向上。
    MM_TEXT = 1                                  '/* 每个逻辑单位转换为一个设置备素，X正方向向右，Y正方向向下。
    MM_TWIPS = 6                                 '/* 每个逻辑单位转换为1 twip (1/1440 inch)，X正方向向右，Y方向向上。
End Enum

' GetROP2,SetROP2
Public Enum EnumDrawModeFlags
    R2_BLACK = 1                                 '/* 黑色
    R2_COPYPEN = 13                              '/* 画笔颜色
    R2_LAST = 16
    R2_MASKNOTPEN = 3                            '/* 画笔颜色的反色与显示颜色进行AND运算
    R2_MASKPEN = 9                               '/* 显示颜色与画笔颜色进行AND运算
    R2_MASKPENNOT = 5                            '/* 显示颜色的反色与画笔颜色进行AND运算
    R2_MERGENOTPEN = 12                          '/* 画笔颜色的反色与显示颜色进行OR运算
    R2_MERGEPEN = 15                             '/* 画笔颜色与显示颜色进行OR运算
    R2_MERGEPENNOT = 14                          '/* 显示颜色的反色与画笔颜色进行OR运算
    R2_NOP = 11                                  '/* 不变
    R2_NOT = 6                                   '/* 当前显示颜色的反色
    R2_NOTCOPYPEN = 4                            '/* R2_COPYPEN的反色
    R2_NOTMASKPEN = 8                            '/* R2_MASKPEN的反色
    R2_NOTMERGEPEN = 2                           '/* R2_MERGEPEN的反色
    R2_NOTXORPEN = 10                            '/* R2_XORPEN的反色
    R2_WHITE = 16                                '/* 白色
    R2_XORPEN = 7                                '/* 显示颜色与画笔颜色进行异或运算
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

' 这个结构包含了附加的绘图参数，函数DrawTextEx
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

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃-----------------------------消息函数和消息列队函数---------------------------------┃
'┃                                                                                    ┃
'
' 调用一个窗口的窗口函数，将一条消息发给那个窗口。除非消息处理完毕，否则该函数不会返回。
' SendMessageBynum， SendMessageByString是该函数的“类型安全”声明形式
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' 将一条消息投递到指定窗口的消息队列。投递的消息会在Windows事件处理过程中得到处理。
' 在那个时候，会随同投递的消息调用指定窗口的窗口函数。特别适合那些不需要立即处理的窗口消息的发送
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃--------------------------------窗口函数(Window)------------------------------------┃
'┃                                                                                    ┃
'
' Creating new windows:
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' 最小化指定的窗口。窗口不会从内存中清除
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
' 破坏（即清除）指定的窗口以及它的所有子窗口
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
' 在指定的窗口里允许或禁止所有鼠标及键盘输入
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
' 在窗口列表中寻找与指定条件相符的第一个子窗口
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' 判断指定窗口的父窗口
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
' 指定一个窗口的新父（在vb里使用：利用这个函数，vb可以多种形式支持子窗口。
' 例如，可将控件从一个容器移至窗体中的另一个。用这个函数在窗体间移动控件是相当冒险的，
' 但却不失为一个有效的办法。如真的这样做，请在关闭任何一个窗体之前，注意用SetParent将控件的父设回原来的那个）
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' 锁定指定窗口，禁止它更新。同时只能有一个窗口处于锁定状态
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
' 强制立即更新窗口，窗口中以前屏蔽的所有区域都会重画
' 在vb里使用：如vb窗体或控件的任何部分需要更新，可考虑直接使用refresh方法
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
' 控制窗口的可见性
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' 改变指定窗口的位置和大小。顶级窗口可能受最大或最小尺寸的限制，那些尺寸优先于这里设置的参数
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' 这个函数能为窗口指定一个新位置和状态。它也可改变窗口在内部窗口列表中的位置。
' 该函数与DeferWindowPos函数相似，只是它的作用是立即表现出来的
' 在vb里使用：针对vb窗体，如它们在win32下屏蔽或最小化，则需重设最顶部状态。
' 如有必要，请用一个子类处理模块来重设最顶部状态)
' 参数
' hwnd             欲定位的窗口
' hWndInsertAfter  窗口句柄。在窗口列表中，窗口hwnd会置于这个窗口句柄的后面，参看本模块枚举KhanSetWindowPosStyles
' x                窗口新的x坐标。如hwnd是一个子窗口，则x用父窗口的客户区坐标表示
' y                窗口新的y坐标。如hwnd是一个子窗口，则y用父窗口的客户区坐标表示
' cx               指定新的窗口宽度
' cy               指定新的窗口高度
' wFlags           包含了旗标的一个整数，参看本模块枚举KhanSetWindowPosStyles
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' 从指定窗口的结构中取得信息，nIndex参数参看本模块常量声明
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' 在窗口结构中为指定的窗口设置信息，nIndex参数参看本模块常量声明
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃------------------------------窗口类函数(Window Class)------------------------------┃
'┃                                                                                    ┃
'
' 为指定的窗口取得类名
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃-----------------------------鼠标输入函数(Mouse Input)------------------------------┃
'
' 获得一个窗口的句柄，这个窗口位于当前输入线程，且拥有鼠标捕获（鼠标活动由它接收）
Public Declare Function GetCapture Lib "user32" () As Long
' 将鼠标捕获设置到指定的窗口。在鼠标按钮按下的时候，这个窗口会为当前应用程序或整个系统接收所有鼠标输入
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
' 为当前的应用程序释放鼠标捕获
Public Declare Function ReleaseCapture Lib "user32" () As Long
' 可以模拟一次鼠标事件，比如左键单击、双击和右键单击等
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
' 这个函数判断指定的点是否位于矩形lpRect内部
'Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃-----------------------------键盘输入函数(Mouse Input)------------------------------┃
'
' 获得拥有输入焦点的窗口的句柄
Public Declare Function GetFocus Lib "user32" () As Long
' 输入焦点设到指定的窗口
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃----------------坐标空间与变换函数(Coordinate Space Transtormation)-----------------┃
'
' 判断窗口内以客户区坐标表示的一个点的屏幕坐标
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
' 判断屏幕上一个指定点的客户区坐标
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃---------------------------设备场景函数(Device Context)-----------------------------┃
'
' 创建一个与特定设备场景一致的内存设备场景。在绘制之前，先要为该设备场景选定一个位图。
' 不再需要时，该设备场景可用DeleteDC函数删除。删除前，其所有对象应回复初始状态
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
' 为专门设备创建设备场景
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
' 获取指定窗口的设备场景，用本函数获取的设备场景一定要用ReleaseDC函数释放，不能用DeleteDC
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
' 释放由调用GetDC或GetWindowDC函数获取的指定设备场景。它对类或私有设备场景无效（但这样的调用不会造成损害）
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
' 删除专用设备场景或信息场景，释放所有相关窗口资源。不要将它用于GetDC函数取回的设备场景
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
' 每个设备场景都可能有选入其中的图形对象。其中包括位图、刷子、字体、画笔以及区域等等。
' 一次选入设备场景的只能有一个对象。选定的对象会在设备场景的绘图操作中使用。
' 例如，当前选定的画笔决定了在设备场景中描绘的线段颜色及样式
' 返回值通常用于获得选入DC的对象的原始值。
' 绘图操作完成后，原始的对象通常选回设备场景。在清除一个设备场景前，务必注意恢复原始的对象
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
' 用这个函数删除GDI对象，比如画笔、刷子、字体、位图、区域以及调色板等等。对象使用的所有系统资源都会被释放
' 不要删除一个已选入设备场景的画笔、刷子或位图。如删除以位图为基础的阴影（图案）刷子，
' 位图不会由这个函数删除——只有刷子被删掉
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'根据指定设备场景代表的设备的功能返回信息
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
' 取得对指定对象进行说明的一个结构
' lpObject 任何类型，用于容纳对象数据的结构。
' 针对画笔，通常是一个LOGPEN结构；针对扩展画笔，通常是EXTLOGPEN；
' 针对字体是LOGBRUSH；针对位图是BITMAP；针对DIBSection位图是DIBSECTION；
' 针对调色板，应指向一个整型变量，代表调色板中的条目数量
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' 在窗口（由设备场景代表）中水平和（或）垂直滚动矩形
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
' 将两个区域组合为一个新区域
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
' 创建一个由点X1，Y1和X2，Y2描述的矩形区域，不用时一定要用DeleteObject函数删除该区域
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' 创建一个由lpRect确定的矩形区域
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
' 创建一个圆角矩形，该矩形由X1，Y1-X2，Y2确定，并由X3，Y3确定的椭圆描述圆角弧度
' 用该函数创建的区域与用RoundRect API函数画的圆角矩形不完全相同，因为本矩形的右边和下边不包括在区域之内
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' 用指定刷子填充指定区域
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
' 用指定刷子围绕指定区域画一个外框
Public Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
' 这是那些很难有人注意到的对编程者来说是个巨大的宝藏的隐含的API函数中的一个。本函数允许您改变窗口的区域。
' 通常所有窗口都是矩形的——窗口一旦存在就含有一个矩形区域。本函数允许您放弃该区域。
' 这意味着您可以创建圆的、星形的窗口，也可以将它分为两个或许多部分——实际上可以是任何形状
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
' 该函数选择一个区域作为指定设备环境的当前剪切区域
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃---------------------------------位图函数(Bitmap)-----------------------------------┃
'
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' 创建一幅与设备有关位图
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃----------------------------------图标函数(Icon)------------------------------------┃
'
' 制作指定图标或鼠标指针的一个副本。这个副本从属于发出调用的应用程序
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
' 创建一个图标
Public Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
' 该函数清除图标和释放任何被图标占用的存储空间。
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
' 该函数在限定的设备上下文窗口的客户区域绘制图标
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' 该函数在限定的设备上下文窗口的客户区域绘制图标，执行限定的光栅操作，并按特定要求伸长或压缩图标或光标。
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
' 取得与图标有关的信息
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃---------------------------------光标函数(Cursor)-----------------------------------┃
'
Public Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
' 从指定的模块或应用程序实例中载入一个鼠标指针。LoadCursorBynum是LoadCursor函数的类型安全声明
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
' 该函数销毁一个光标并释放它占用的任何内存，不要使用该函数去消毁一个共享光标。
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
' 获取鼠标指针的当前位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' 该函数把光标移到屏幕的指定位置。如果新位置不在由 ClipCursor函数设置的屏幕矩形区域之内，
' 则系统自动调整坐标，使得光标在矩形之内。
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃-----------------------------笔刷函数(Pen and Brush)---------------------------------┃
'
' 用指定的样式、宽度和颜色创建一个画笔，用DeleteObject函数将其删除
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' 根据指定的LOGPEN结构创建一个画笔
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
' 创建一个扩展画笔（装饰或几何）
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
' 在一个LOGBRUSH数据结构的基础上创建一个刷子
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
' 该函数可以创建一个具有指定阴影模式和颜色的逻辑刷子。
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
' 该函数可以创建具有指定位图模式的逻辑刷子，该位图不能是DIB类型的位图，DIB位图是由CreateDIBSection函数创建的。
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
' 用纯色创建一个刷子，一旦刷子不再需要，就用DeleteObject函数将其删除
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' 为任何一种标准系统颜色取得一个刷子，不要用DeleteObject函数删除这些刷子。
' 它们是由系统拥有的固有对象。不要将这些刷子指定成一种窗口类的默认刷子
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃---------------------------字体和正文函数(Font and Text)-----------------------------┃
'
' 用指定的属性创建一种逻辑字体，VB的字体属性在选择字体的时候显得更有效
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
' 将文本描绘到指定的矩形中，wFormat标志常数参看KhanDrawTextStyles
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
' 该函数取得指定设备环境的当前正文颜色。
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
' 设置当前文本颜色。这种颜色也称为“前景色”，如改变了这个设置，注意恢复VB窗体或控件原始的文本颜色
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃------------------------------------绘图函数----------------------------------------┃
'
' 该函数画一段圆弧，圆弧是由一个椭圆和一条线段（称之为割线）相交限定的闭合区域。
' 此弧由当前的画笔画轮廓，由当前的画刷填充。
Public Declare Function Chord Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' 用指定的样式描绘一个矩形的边框。利用这个函数，我们没有必要再使用许多3D边框和面板。
' 所以就资源和内存的占用率来说，这个函数的效率要高得多。它可在一定程度上提升性能
' hdc      要在其中绘图的设备场景
' qrc      要为其描绘边框的矩形
' edge     带有前缀BDR_的两个常数的组合。一个指定内部边框是上凸还是下凹；另一个则指定外部边框。有时能换用带EDGE_前缀的常数。
' grfFlags 带有BF_前缀的常数的组合
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
' 画一个焦点矩形。这个矩形是在标志焦点的样式中通过异或运算完成的（焦点通常用一个点线表示）
' 如用同样的参数再次调用这个函数，就表示删除焦点矩形
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
' 这个函数用于描绘一个标准控件
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
' 这个函数可为一幅图象或绘图操作应用各式各样的效果
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
' 该函数用于画一个椭圆，椭圆的中心是限定矩形的中心，使用当前画笔画椭圆，用当前的画刷填充椭圆。
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' 用指定的刷子填充一个矩形，矩形的右边和底边不会描绘
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' 用指定的刷子围绕一个矩形画一个边框（组成一个帧），边框的宽度是一个逻辑单位
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' 取得指定设备场景当前的背景颜色
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
' 针对指定的设备场景，取得当前的背景填充模式
Public Declare Function GetBkMode Lib "gdi32" (ByVal hDC As Long) As Long
' 为指定的设备场景设置背景颜色。背景颜色用于填充阴影刷子、虚线画笔以及字符（如背景模式为OPAQUE）中的空隙。
' 也在位图颜色转换期间使用。背景实际是设备能够显示的最接近于 crColor 的颜色
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
' 指定阴影刷子、虚线画笔以及字符中的空隙的填充方式，背景模式不会影响用扩展画笔描绘的线条
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
' 在指定的设备场景中取得一个像素的RGB值
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
' 在指定的设备场景中设置一个像素的RGB值
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
' 将来自一幅位图的二进制位复制到一幅与设备无关的位图里
'Public Declare Function GetDIBits Lib "gdi32" ( ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' 将来自与设备无关位图的二进制位复制到一幅与设备有关的位图里
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' 针对指定的设备场景，获得多边形填充模式。
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hDC As Long) As Long
' 设置多边形的填充模式
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
' 针对指定的设备场景，取得当前的绘图模式。这样可定义绘图操作如何与正在显示的图象合并起来
' 这个函数只对光栅设备有效
Public Declare Function GetROP2 Lib "gdi32" (ByVal hDC As Long) As Long
' 设置指定设备场景的绘图模式。
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
' 用当前画笔画一条线，从当前位置连到一个指定的点。这个函数调用完毕，当前位置变成x,y点
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
' 为指定的设备场景指定一个新的当前画笔位置。前一个位置保存在lpPoint中
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
' 该函数画一个由椭圆和两条半径相交闭合而成的饼状楔形图，此饼图由当前画笔画轮廓，由当前画刷填充。
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' 该函数画一个由直线相闻的两个以上顶点组成的多边形，用当前画笔画多边形轮廓，
' 用当前画刷和多边形填充模式填充多边形。
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
' 用当前画笔描绘一系列线段。使用PolylineTo函数时，当前位置会设为最后一条线段的终点。
' 它不会由Polyline函数改动
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PolyPolyline Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
' 该函数画一个矩形，用当前的画笔画矩形轮廓，用当前画刷进行填充。
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' 函数画一个带圆角的矩形，此矩形由当前画笔画轮廊，由当前画刷填充。
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' 这个函数用于增大或减小一个矩形的大小。
' x加在右侧区域，并从左侧区域减去；如x为正，则能增大矩形的宽度；如x为负，则能减小它。
' y对顶部与底部区域产生的影响是是类似的
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' 该函数通过应用一个指定的偏移，从而让矩形移动起来。
' x会添加到右侧和左侧区域。y添加到顶部和底部区域。
' 偏移方向则取决于参数是正数还是负数，以及采用的是什么坐标系统
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' 返回与windows环境有关的信息，nIndex值参看本模块的常量声明
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
' 获得整个窗口的范围矩形，窗口的边框、标题栏、滚动条及菜单等都在这个矩形内
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' 返回指定窗口客户区矩形的大小
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' 这个函数屏蔽一个窗口客户区的全部或部分区域。这会导致窗口在事件期间部分重画
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
' 判断指定windows显示对象的颜色，颜色对象看本模块声明
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
'┃--------------------------------其他函数(Others)------------------------------------┃
'
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
'┃                                                                                    ┃
'┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

