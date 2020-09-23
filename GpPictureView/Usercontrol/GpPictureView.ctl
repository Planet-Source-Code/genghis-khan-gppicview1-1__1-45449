VERSION 5.00
Begin VB.UserControl GpPictureView 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "GpPictureView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' ***************************************************************************************
' * Project  | GpPictureView                                                            *
' *----------|--------------------------------------------------------------------------*
' * Version  | V1.1                                                                     *
' *----------|--------------------------------------------------------------------------*
' * Author   | Genghis Khan(GuangJian Guo)                                              *
' *----------|--------------------------------------------------------------------------*
' * WebSite  | http://www.itkhan.com                                                    *
' *----------|--------------------------------------------------------------------------*
' * MailTo   | webmaster@itkhan.com                                                     *
' *----------|--------------------------------------------------------------------------*
' * Date     | 13 May 2003                                                              *
' *----------|--------------------------------------------------------------------------*
' * This program and source code is freeware and can be copied and/or distributed       *
' * as long as you mention the original author. I am not responsible for any harm       *
' * as the outcome of using any of this code.                                           *
' ***************************************************************************************

' ======================================================================================
' Constants
' ======================================================================================
Private Const MODULE_NAME = "GpPictureView"
Private Const DefaultPicItemWidth = 94
Private Const DefaultPicItemHeight = 106
Private Const DefaultPicItemBorder = 6
Private Const DefaultSpaceBetweenItems = 6
Private Const DefaultBackColor = &HA2A2A2 'vbApplicationWorkspace
Private Const DefaultItemBackColor = vbButtonFace
Private Const DefaultCaptionBackcolor = vbInfoBackground

' ======================================================================================
' Enums:
' ======================================================================================
Public Enum GPPVW_BORDERSTYLE_METHOD
   GpPvwBorderStyleNone = 0
   GpPvwBorderStyle3D = 1
   GpPvwBorderStyle3DThin = 2
End Enum

' ======================================================================================
' Types:
' ======================================================================================

Private Type PicViewPackFileHeader
    pfType(10)                   As Byte  ' PicViewPack
    pfAuthorInfo(11)             As Byte  ' GuangJianGuo
    pfFileCount                  As Long
End Type

Private Type PicViewPackInfoHeader
    piCaption                    As String * 50
    piInfo                       As String * 10
    PiStart                      As Double
    piSize                       As Double
End Type

' ======================================================================================
' Private variables:
' ======================================================================================

' Drawing area:
Private m_lAvailWidth             As Long
Private m_lAvailheight            As Long
Private m_lngVerBarValue          As Long
Private m_lngRowCount             As Long
Private m_lngColCount             As Long
Private m_lngFirstRow             As Long
Private m_lngLastRow              As Long
Private m_lngItemWidth            As Long
Private m_lngItemHeight           As Long
Private m_lngItemTotalWidth       As Long
Private m_lngItemTotalHeight      As Long
Private m_lngTextHeight           As Long

' Memory DC for flicker-free (1 row only) - also implements clipping
Private m_hWndCtl                 As Long
Private m_hDC                     As Long
Private m_hBmp                    As Long
Private m_hBmpOld                 As Long
Private m_hFntDC                  As Long
Private m_hFntOldDC               As Long

Private m_lngHoverIndex           As Long
Private m_lngSelFirst             As Long
Private m_colSelected             As Collection

Private m_strThumbPack            As String
Private m_blnLoadThumbPack        As Boolean
Private m_blnHotTracking          As Boolean
Private m_blnDirty                As Boolean
Private m_blnRedraw               As Boolean
Private m_blnUserMode             As Boolean
Private m_blnEnabled              As Boolean
Private m_bInFocus                As Boolean
Private m_blnHideSelection        As Boolean
Private m_blnMultiSelect          As Boolean
Private m_oleHighlightBackColor   As OLE_COLOR
Private m_oleHighlightForeColor   As OLE_COLOR
Private m_udtBorderStyle          As GPPVW_BORDERSTYLE_METHOD
Private m_udtDrawTextParams       As DRAWTEXTPARAMS
Private m_udtPackHeader           As PicViewPackFileHeader
Private m_udtPackInfo()           As PicViewPackInfoHeader
Private m_clsSelectedItem         As CPicItem
Private WithEvents m_clsPicitems  As CPicItems
Attribute m_clsPicitems.VB_VarHelpID = -1
Private WithEvents m_cScroll      As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

' ======================================================================================
' Events
' ======================================================================================
Public Event Click()
Public Event DblClick()
Public Event ItemClick(ByVal Item As CPicItem)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Get BackColor() As OLE_COLOR
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
   UserControl.BackColor = NewValue
   If (m_hDC <> 0) Then
      Call SetBkColor(m_hDC, TranslateColor(NewValue))
   End If
   PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As GPPVW_BORDERSTYLE_METHOD
    BorderStyle = m_udtBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As GPPVW_BORDERSTYLE_METHOD)
    Dim lngStyle As Long
    
    m_udtBorderStyle = NewValue
    If (NewValue = GpPvwBorderStyleNone) Then
       UserControl.BorderStyle() = vbBSNone
    Else
       UserControl.BorderStyle() = vbFixedSingle
       lngStyle = GetWindowLong(UserControl.hWnd, GWL_EXSTYLE)
       If (NewValue = GpPvwBorderStyle3DThin) Then
          lngStyle = lngStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
       Else
          lngStyle = lngStyle Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
       End If
       Call SetWindowLong(UserControl.hWnd, GWL_EXSTYLE, lngStyle)
       Call SetWindowPos(UserControl.hWnd, 0, 0, 0, 0, 0, _
                         SWP_NOACTIVATE Or SWP_NOZORDER Or _
                         SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE)
    End If
    PropertyChanged "BorderStyle"
End Property

Public Property Let CanLoadThumbPack(ByVal NewValue As Boolean)
    m_blnLoadThumbPack = NewValue
End Property

Public Property Get CanLoadThumbPack() As Boolean
    CanLoadThumbPack = m_blnLoadThumbPack
End Property

' Create Thumbnail Image Pack file
Public Sub CreateThumbImagePack(ByVal sFileName As String, ImageFiles() As String)
    Dim clsSrc                  As cDIBSection
    Dim clsThumb                As cDIBSection
    Dim lWidthMax               As Long
    Dim lheightMax              As Long
    Dim lngI                    As Long
    Dim lngY                    As Long
    Dim strFile                 As String
    Dim intFile                 As Long
    Dim blnOk                   As Boolean
    Dim byMessage()             As Byte
    Dim lPtr                    As Long
    Dim lngFileSize             As Long
    Dim lngAllSize              As Long
    Dim uHeader                 As PicViewPackFileHeader
    Dim uInfo()                 As PicViewPackInfoHeader
    Const ProcName = "CreateThumbImagePack"
    
    On Error GoTo ErrorHandle
    
    lWidthMax = Me.PictureWidthMax
    lheightMax = Me.PictureHeightMax
    With uHeader
         .pfType(0) = Asc("P")
         .pfType(1) = Asc("i")
         .pfType(2) = Asc("c")
         .pfType(3) = Asc("V")
         .pfType(4) = Asc("i")
         .pfType(5) = Asc("e")
         .pfType(6) = Asc("w")
         .pfType(7) = Asc("P")
         .pfType(8) = Asc("a")
         .pfType(9) = Asc("c")
         .pfType(10) = Asc("k")
         .pfAuthorInfo(0) = Asc("G")
         .pfAuthorInfo(1) = Asc("u")
         .pfAuthorInfo(2) = Asc("a")
         .pfAuthorInfo(3) = Asc("n")
         .pfAuthorInfo(4) = Asc("g")
         .pfAuthorInfo(5) = Asc("J")
         .pfAuthorInfo(6) = Asc("i")
         .pfAuthorInfo(7) = Asc("a")
         .pfAuthorInfo(8) = Asc("n")
         .pfAuthorInfo(9) = Asc("G")
         .pfAuthorInfo(10) = Asc("u")
         .pfAuthorInfo(11) = Asc("o")
         .pfFileCount = UBound(ImageFiles) + 1
    End With
    ReDim uInfo(LBound(ImageFiles) To UBound(ImageFiles))
    lngAllSize = Len(uHeader) + Len(uInfo(0)) * (UBound(ImageFiles) + 1)
    lngAllSize = lngAllSize + 1
    intFile = FreeFile
    Open sFileName For Output As #intFile
    Close #intFile
    intFile = FreeFile
    Open sFileName For Binary Access Write Lock Write As #intFile
    Put #intFile, 1, uHeader
    Put #intFile, , uInfo
    
    For lngI = LBound(ImageFiles) To UBound(ImageFiles)
        For lngY = Len(ImageFiles(lngI)) To 1 Step -1
           If Mid$(ImageFiles(lngI), lngY, 1) = "\" Then
              strFile = Mid$(ImageFiles(lngI), lngY + 1)
              Exit For
           End If
        Next lngY
        Set clsSrc = New cDIBSection
        blnOk = LoadJPG(clsSrc, ImageFiles(lngI))
        If Not blnOk Then blnOk = clsSrc.CreateFromFile(ImageFiles(lngI))
        If blnOk Then
           Set clsThumb = clsSrc.GetThumbnailDIB(lWidthMax, lheightMax)
           If Not (clsThumb Is Nothing) Then
              ReDim byMessage(0 To clsThumb.Height * clsThumb.BytesPerScanLine / 4)
              lPtr = VarPtr(byMessage(0))
              lngFileSize = UBound(byMessage) - 1
              blnOk = SaveJPGToPtr(clsThumb, lPtr, lngFileSize, 65)
              If blnOk Then
                 ReDim Preserve byMessage(0 To lngFileSize - 1) As Byte
                 With uInfo(lngI)
                      .piInfo = Trim$(Str$(clsSrc.Width)) & "x" & Trim$(Str$(clsSrc.Height))
                      .piCaption = strFile
                      .piSize = lngFileSize
                      .PiStart = lngAllSize
                 End With
                 Put #intFile, , byMessage()
                 lngAllSize = lngAllSize + lngFileSize
              End If
           End If
        End If
        Set clsSrc = Nothing
        Set clsThumb = Nothing
        Erase byMessage
    Next lngI
    Put #intFile, 1, uHeader
    Put #intFile, , uInfo
    Close #intFile
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Set clsSrc = Nothing
             Set clsThumb = Nothing
             Erase byMessage
    End Select
End Sub

Public Property Get Enabled() As Boolean
   Enabled = m_blnEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    Dim lngI                               As Long
    Dim lngListStart                       As Long
    Dim lngListCount                       As Long
    
    m_blnEnabled = NewValue
    If UserControl.Ambient.UserMode Then
       m_blnDirty = True
       Call UserControl_Paint
    End If
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Dim udtFont As LOGFONT
     
    Set UserControl.Font = NewValue
       
    If (m_hFntDC <> 0) Then
       If (m_hDC <> 0) Then
          If (m_hFntOldDC <> 0) Then
             Call DeleteObject(SelectObject(m_hDC, m_hFntOldDC))
          End If
          Call DeleteObject(m_hFntDC)
       End If
    End If
    
    Call OLEFontToLogFont(NewValue, UserControl.hDC, udtFont)
    m_hFntDC = CreateFontIndirect(udtFont)
    If (m_hDC <> 0) Then
       m_hFntOldDC = SelectObject(m_hDC, m_hFntDC)
    End If
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
   ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
   UserControl.ForeColor = NewValue
   If (m_hDC <> 0) Then
      Call SetTextColor(m_hDC, TranslateColor(NewValue))
   End If
   PropertyChanged "ForeColor"
End Property

Friend Sub fPaintControl(ByVal AddNew As Boolean, Optional ByVal Item As CPicItem)
    Const ProcName = "fPaintControl"
    
    On Error GoTo ErrorHandle
    If AddNew Then
       If m_clsSelectedItem Is Nothing Then
          Set m_clsSelectedItem = Item
          With m_clsSelectedItem
               .AutoReDraw = False
               .Selected = True
               .AutoReDraw = True
          End With
          m_lngSelFirst = Item.Index
          Set m_colSelected = Nothing
          Set m_colSelected = New Collection
          m_colSelected.Add Item.Index
       End If
       If m_hDC = 0 Then Call pvBuildMemDC
       If m_blnRedraw And m_blnUserMode Then
          Call pvScrollSetDirty(True)
          Call pvDraw
       End If
    Else
       If m_blnRedraw And m_blnUserMode Then Call pvDraw
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Friend Sub fPaintControlDel(ByVal IsClear As Boolean, Optional ByVal Index As Long)
    Dim lngI As Long
    Dim lngIndex As Long
    Dim tR As RECT
    Const ProcName = "fPaintControlDel"
    
    On Error GoTo ErrorHandle
    If IsClear Then
       Set m_colSelected = Nothing
       Set m_colSelected = New Collection
       Set m_clsSelectedItem = Nothing
       Call pvScrollVisible
       Call GetClientRect(UserControl.hWnd, tR)
       Call pvFillBackground(UserControl.hDC, tR)
    Else
       For lngI = 1 To SelectedCount
           lngIndex = m_colSelected(lngI)
           If lngIndex = Index Then
              m_colSelected.Remove lngI
              Exit For
           End If
       Next lngI
       Set m_colSelected = Nothing
       Set m_colSelected = New Collection
       For lngI = 1 To m_clsPicitems.Count
           With m_clsPicitems.Item(lngI)
                If .Selected Then m_colSelected.Add .Index
           End With
       Next lngI
       If m_clsSelectedItem.Index = Index Then
          If Index > m_clsPicitems.Count Then
             Set m_clsSelectedItem = m_clsPicitems.Item(m_clsPicitems.Count)
             m_lngSelFirst = m_clsPicitems.Count
          End If
       End If
       If m_blnRedraw Then
          Call pvScrollVisible
          Call pvDraw
       End If
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Friend Sub fPaintControlSelect(ByVal Index As Long)
    Dim lngI As Long
    Const ProcName = "fPaintControlSelect"
    
    On Error GoTo ErrorHandle
    If m_clsPicitems(Index).Selected Then
'       If m_blnMultiSelect Then
'          m_colSelected.Add Index
'          m_lngSelFirst = Index
'       Else
          Call pvSingleModeSelect(Index)
'       End If
    Else
       For lngI = 1 To SelectedCount
           If m_colSelected(lngI) = Index Then
              m_colSelected.Remove lngI
              Exit For
           End If
       Next lngI
    End If
    If m_blnRedraw Then Call pvDraw
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Friend Sub fPaintEnsureVisible(ByVal Index As Long)
    Const ProcName = "fPaintEnsureVisible"
    
    On Error GoTo ErrorHandle
    If m_blnRedraw Then
       If Not pvEnsureVisible(Index) Then Call pvDraw
    End If
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Public Property Get HideSelection() As Boolean
    HideSelection = m_blnHideSelection
End Property

Public Property Let HideSelection(ByVal NewValue As Boolean)
    m_blnHideSelection = NewValue
End Property

Public Property Get HighlightBackColor() As OLE_COLOR
    HighlightBackColor = m_oleHighlightBackColor
End Property

Public Property Let HighlightBackColor(NewValue As OLE_COLOR)
    m_oleHighlightBackColor = NewValue
    PropertyChanged "HighlightBackColor"
End Property

Public Property Get HighlightForeColor() As OLE_COLOR
    HighlightForeColor = m_oleHighlightForeColor
End Property

Public Property Let HighlightForeColor(NewValue As OLE_COLOR)
    m_oleHighlightForeColor = NewValue
    PropertyChanged "HighlightForeColor"
End Property

Public Function HitTest(x As Single, y As Single) As CPicItem
    Dim lngI              As Long
    Dim lngY              As Long
    Dim lngCellStart      As Long
    Dim tR                As RECT
    Const ProcName = "HitTest"
    
    On Error GoTo ErrorHandle
    
    If x < 0 Or x > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight Then
       Set HitTest = Nothing
       Exit Function
    End If
       
    For lngI = m_lngFirstRow To m_lngLastRow
        lngCellStart = (lngI - 1) * m_lngColCount
        tR.Top = (lngI - 1) * m_lngItemTotalHeight - m_lngVerBarValue
        tR.Bottom = tR.Top + m_lngItemHeight
        For lngY = 1 To m_lngColCount
            If (lngCellStart + lngY) > m_clsPicitems.Count Then Exit For
            tR.Left = (DefaultSpaceBetweenItems + m_lngItemWidth) * (lngY - 1)
            tR.Right = tR.Left + m_lngItemTotalWidth
            If PtInRect(tR, x, y) <> 0 Then
               Set HitTest = m_clsPicitems.Item(lngCellStart + lngY)
               Exit Function
            End If
        Next lngY
    Next lngI
    
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Function

Public Property Get HotTracking() As Boolean
    HotTracking = m_blnHotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    m_blnHotTracking = New_HotTracking
    PropertyChanged "HotTracking"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Sub LoadThumbPack(ByVal sFileName As String)
    Dim intFile As Integer
    Dim lngI As Long
    Const ProcName = "CreateThumbImagePack"
    
    On Error GoTo ErrorHandle
    
    m_strThumbPack = sFileName
    m_clsPicitems.Clear
    intFile = FreeFile
    Open sFileName For Binary Access Read Lock Read As #intFile
    Get #intFile, 1, m_udtPackHeader
    ReDim m_udtPackInfo(m_udtPackHeader.pfFileCount - 1)
    Get #intFile, Len(m_udtPackHeader) + 1, m_udtPackInfo
    Close #intFile
    
    For lngI = 0 To UBound(m_udtPackInfo)
        With m_udtPackInfo(lngI)
             m_clsPicitems.AddThumbItem , Trim$(.piCaption), Trim$(.piInfo)
        End With
    Next lngI
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Public Property Get MultiSelect() As Boolean
    MultiSelect = m_blnMultiSelect
End Property

Public Property Let MultiSelect(ByVal NewValue As Boolean)
    m_blnMultiSelect = NewValue
    PropertyChanged "MultiSelect"
End Property

Public Property Get PicItemHeight() As Long
    PicItemHeight = m_lngItemHeight * Screen.TwipsPerPixelX
End Property

Public Property Let PicItemHeight(ByVal NewValue As Long)
    m_lngItemHeight = NewValue \ Screen.TwipsPerPixelY
    m_lngItemTotalHeight = m_lngItemHeight + DefaultSpaceBetweenItems
    PropertyChanged "PicItemHeight"
End Property

Public Property Get PicItems() As CPicItems
    Set PicItems = m_clsPicitems
End Property

Public Property Get PicItemWidth() As Long
    PicItemWidth = m_lngItemWidth * Screen.TwipsPerPixelX
End Property

Public Property Let PicItemWidth(ByVal NewValue As Long)
    m_lngItemWidth = NewValue \ Screen.TwipsPerPixelX
    m_lngItemTotalWidth = m_lngItemWidth + DefaultSpaceBetweenItems
    PropertyChanged "PicItemWidth"
End Property

Public Property Get PictureHeightMax() As Long
    Dim lngHeight As Long
    Dim lngTextHeight As Long
    
    lngTextHeight = pvGetTextHeight
    lngHeight = m_lngItemHeight - (lngTextHeight + 8) * 2 - 2
    PictureHeightMax = lngHeight
End Property

Public Property Get PictureWidthMax() As Long
    Dim lngWidth As Long
    
    lngWidth = m_lngItemWidth - DefaultPicItemBorder * 2 - 2
    PictureWidthMax = lngWidth
End Property

Private Sub pvBuildMemDC()
    Dim tR               As RECT
    Const ProcName = "pvBuildMemDC"
    
    On Error GoTo ErrorHandle
    
    If (m_hBmp <> 0) Then
       If (m_hBmpOld <> 0) Then Call SelectObject(m_hDC, m_hBmpOld)
       If (m_hBmp <> 0) Then Call DeleteObject(m_hBmp)
       m_hBmp = 0
       m_hBmpOld = 0
    End If
    If (m_hDC = 0) Then
       m_hDC = CreateCompatibleDC(UserControl.hDC)
    Else
       Call SelectObject(m_hDC, m_hFntOldDC)
    End If
    
    If (m_hDC <> 0) Then
       m_hBmp = CreateCompatibleBitmap(UserControl.hDC, m_lngItemTotalWidth, m_lngItemTotalHeight)
       If (m_hBmp <> 0) Then
          m_hBmpOld = SelectObject(m_hDC, m_hBmp)
          If (m_hBmpOld = 0) Then
             Call DeleteObject(m_hBmp)
             Call DeleteObject(m_hDC)
             m_hBmp = 0
             m_hDC = 0
          Else
             Call SetTextColor(m_hDC, TranslateColor(UserControl.ForeColor))
             Call SetBkColor(m_hDC, TranslateColor(UserControl.BackColor))
             Call SetBkMode(m_hDC, TRANSPARENT)
             m_hFntOldDC = SelectObject(m_hDC, m_hFntDC)
             tR.Right = m_lngItemTotalWidth
             tR.Bottom = m_lngItemTotalHeight
             Call pvFillBackground(m_hDC, tR)
          End If
       Else
          Call DeleteObject(m_hDC)
          m_hDC = 0
       End If
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Call DeleteObject(m_hBmp)
             Call DeleteObject(m_hDC)
             m_hBmp = 0
             m_hDC = 0
    End Select
End Sub

Private Sub pvClearSelection()
    Dim lngI As Long
    
    On Error Resume Next
    If SelectedCount > 0 Then
       For lngI = 1 To SelectedCount
           With m_clsPicitems.Item(m_colSelected(lngI))
                .AutoReDraw = False
                .Selected = False
                .AutoReDraw = True
           End With
       Next
    Else
       With m_clsSelectedItem
            .AutoReDraw = False
            .Selected = False
            .AutoReDraw = True
       End With
       Set m_clsSelectedItem = Nothing
    End If
    Set m_colSelected = Nothing
    Set m_colSelected = New Collection
End Sub

Private Sub pvDraw()
    Dim lngI                       As Long
    Dim lngY                       As Long
    Dim lngCellStart               As Long
    Dim lngLastPos                 As Long
    Dim hBr                        As Long
    Dim hBrCaption                 As Long
    Dim lngImageFitWidth           As Long
    Dim lngImageFitHeight          As Long
    Dim strCopy                    As String
    Dim blnSel                     As Boolean
    Dim blnDrawBack                As Boolean
    Dim blnDoIt                    As Boolean
    Dim tR                         As RECT
    Dim tBR                        As RECT
    Dim tMR                        As RECT
    Dim tFR                        As RECT
    Dim tPR                        As RECT
    Dim tCR                        As RECT
    Dim tIR                        As RECT
    Dim lPtr                       As Long
    Dim intFile                    As Integer
    Dim byImage()                  As Byte
    Dim clsThumb                   As cDIBSection
    Dim lf As Integer
    
    Const ProcName = "pvDraw"
    
    On Error GoTo ErrorHandle
    If m_hDC = 0 Then Exit Sub
    If m_blnRedraw = False Or m_blnUserMode = False Then Exit Sub
    If m_clsPicitems.Count <= 0 Then Exit Sub
    
    m_lngTextHeight = pvGetTextHeight
    
    ' Find the start and end of drawing:
    If (m_lngFirstRow <= 0 Or m_lngLastRow <= 0) Then Call pvGetStartEndRow
    
    For lngI = m_lngFirstRow To m_lngLastRow
        lngCellStart = (lngI - 1) * m_lngColCount
        For lngY = 1 To m_lngColCount
            With tBR
                 .Top = 0
                 .Left = 0
                 .Right = m_lngItemTotalWidth
                 .Bottom = m_lngItemTotalHeight
            End With
            Call pvFillBackground(m_hDC, tBR)
            If (lngCellStart + lngY) <= m_clsPicitems.Count Then
               blnDrawBack = False
               LSet tR = tBR
               With tR
                    .Left = DefaultSpaceBetweenItems
                    .Right = m_lngItemTotalWidth
                    .Bottom = m_lngItemHeight
               End With
               LSet tMR = tR
               tMR.Bottom = tR.Bottom - m_lngTextHeight - 8
               Let tCR = tR
               tCR.Top = tMR.Bottom
               With m_clsPicitems(lngCellStart + lngY)
                    .Col = lngY
                    blnDoIt = m_blnDirty
                    If Not blnDoIt Then
                       If .Dirty Then
                          blnDoIt = True
                          .Dirty = False
                       End If
                    End If
                    If blnDoIt Then
                       If m_blnLoadThumbPack Then
                          intFile = FreeFile
                          Set clsThumb = New cDIBSection
                          ReDim byImage(m_udtPackInfo(.Index - 1).piSize - 1)
                          Open m_strThumbPack For Binary Access Read Lock Read As #intFile
                          Get #intFile, m_udtPackInfo(.Index - 1).PiStart, byImage
                          Close #intFile
                          
'                          lf = FreeFile
'                          Open App.Path & "\" & Trim$(Str$(.Index - 1)) & ".jpg" For Binary Access Write Lock Write As #lf
'                          Put #lf, 1, byImage
'                          Close #lf
                          
                          lPtr = VarPtr(byImage(0))
                          Call LoadJPGFromPtr(clsThumb, lPtr, m_udtPackInfo(.Index - 1).piSize)
                       End If
                       If Len(.PictureInfo) <> 0 Then
                          LSet tIR = tMR
                          tIR.Top = tMR.Bottom - m_lngTextHeight - 8
                          LSet tPR = tMR
                          tPR.Bottom = tIR.Top
                          Call InflateRect(tPR, -DefaultPicItemBorder, -DefaultPicItemBorder)
                       Else
                          LSet tPR = tMR
                          Call InflateRect(tPR, -DefaultPicItemBorder, -DefaultPicItemBorder)
                       End If
                       If m_blnLoadThumbPack Then
                          lngImageFitWidth = clsThumb.Width
                          lngImageFitHeight = clsThumb.Height
                       Else
                          Call GetFitSize(.DIBSection, Me.PictureWidthMax, Me.PictureHeightMax, _
                                          lngImageFitWidth, lngImageFitHeight)
                       End If
                       If lngImageFitHeight > lngImageFitWidth Then
                          tPR.Left = tPR.Left + (tPR.Right - tPR.Left - lngImageFitWidth) \ 2
                          tPR.Right = tPR.Left + lngImageFitWidth
                       End If
                       If lngImageFitWidth > lngImageFitHeight Then
                          tPR.Top = tPR.Top + (tPR.Bottom - tPR.Top - lngImageFitHeight) \ 2
                          tPR.Bottom = tPR.Top + lngImageFitHeight
                       End If
                       If .Selected And m_blnEnabled Then
                          If m_bInFocus Then
                             hBr = CreateSolidBrush(TranslateColor(m_oleHighlightBackColor))
                             Call SetTextColor(m_hDC, TranslateColor(m_oleHighlightForeColor))
                             LSet tFR = tR
                             Call InflateRect(tFR, -1, -1)
                             Call FillRect(m_hDC, tFR, hBr)
                             Call DeleteObject(hBr)
                             blnSel = True
                          Else
                             If Not m_blnHideSelection Then
                                LSet tFR = tMR
                                With tFR
                                     .Right = .Right - 1
                                     .Bottom = .Bottom - 1
                                End With
                                hBr = CreateSolidBrush(TranslateColor(UserControl.BackColor))
                                Call FillRect(m_hDC, tMR, hBr)
                                Call DeleteObject(hBr)
                                hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
                                Call FillRect(m_hDC, tCR, hBr)
                                Call DeleteObject(hBr)
                             Else
                                blnDrawBack = m_blnEnabled
                             End If
                          End If
                       Else
                          blnDrawBack = m_blnEnabled
                       End If
                       If blnDrawBack Then
                          If (.BackColor <> CLR_NONE) Then
                             hBr = CreateSolidBrush(TranslateColor(.BackColor))
                          Else
                             hBr = CreateSolidBrush(TranslateColor(DefaultItemBackColor))
                          End If
                          Call FillRect(m_hDC, tMR, hBr)
                          Call DeleteObject(hBr)
                          If .CaptionBackColor <> CLR_NONE Then
                             hBr = CreateSolidBrush(TranslateColor(.CaptionBackColor))
                          Else
                             hBr = CreateSolidBrush(TranslateColor(DefaultCaptionBackcolor))
                          End If
                          Call FillRect(m_hDC, tCR, hBr)
                          Call DeleteObject(hBr)
                          
                          If m_blnHotTracking And (.Index = m_lngHoverIndex) Then
                             Call SetTextColor(m_hDC, TranslateColor(vbBlue))
                             .Dirty = True
                             blnSel = True
                          Else
                             If (.ForeColor <> CLR_NONE) Then
                                Call SetTextColor(m_hDC, TranslateColor(.ForeColor))
                                blnSel = True
                             End If
                          End If
                       End If
                       
                       Call DrawEdge(m_hDC, tMR, BDR_RAISEDINNER, BF_RECT)
                       Call DrawEdge(m_hDC, tPR, BDR_SUNKENOUTER, BF_RECT)
                       If (Not (m_clsSelectedItem Is Nothing)) Then
                          If .Index = m_clsSelectedItem.Index And m_bInFocus And m_blnEnabled Then
                             LSet tFR = tCR
                             Call InflateRect(tFR, 0, 1)
                             Call DrawFocusRect(m_hDC, tFR)
                             .Dirty = True
                          Else
                             Call DrawEdge(m_hDC, tCR, BDR_SUNKENOUTER, BF_RECT)
                          End If
                       Else
                          Call DrawEdge(m_hDC, tCR, BDR_SUNKENOUTER, BF_RECT)
                       End If
                       strCopy = .PictureInfo
                       If m_blnLoadThumbPack Then
                          clsThumb.PaintPicture m_hDC, APIBitBlt, tPR.Left + 1, tPR.Top + 1, tPR.Right - tPR.Left - 2, tPR.Bottom - tPR.Top - 2
                       Else
                          If lngImageFitWidth > (tPR.Right - tPR.Left - 2) Or lngImageFitHeight > (tPR.Bottom - tPR.Top - 2) Then
                             .DIBSection.PaintPicture m_hDC, APIStretchBlt, tPR.Left + 1, tPR.Top + 1, tPR.Right - tPR.Left - 2, tPR.Bottom - tPR.Top - 2
                          Else
                             .DIBSection.PaintPicture m_hDC, APIBitBlt, tPR.Left + 1, tPR.Top + 1, tPR.Right - tPR.Left - 2, tPR.Bottom - tPR.Top - 2
                          End If
                       End If
                       
                       If Len(strCopy) <> 0 Then
                          Call DrawTextEx(m_hDC, strCopy & vbNullChar, -1, tIR, _
                                          DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CENTER Or DT_TOP, _
                                          m_udtDrawTextParams)
                       End If
                       strCopy = .Caption
                       If Len(strCopy) <> 0 Then
                          Call DrawTextEx(m_hDC, strCopy & vbNullChar, -1, tCR, _
                                          DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CENTER Or DT_VCENTER, _
                                          m_udtDrawTextParams)
                       End If
                       If blnSel Then Call SetTextColor(m_hDC, TranslateColor(UserControl.ForeColor))
                       Call BitBlt(UserControl.hDC, (lngY - 1) * m_lngItemTotalWidth, _
                                  (lngI - 1) * m_lngItemTotalHeight - m_lngVerBarValue, _
                                  m_lngItemTotalWidth, m_lngItemTotalHeight, m_hDC, 0, 0, vbSrcCopy)
                    End If
               End With
            Else
               Call BitBlt(UserControl.hDC, (lngY - 1) * m_lngItemTotalWidth, _
                          (lngI - 1) * m_lngItemTotalHeight - m_lngVerBarValue, _
                          m_lngItemTotalWidth, m_lngItemTotalHeight, m_hDC, 0, 0, vbSrcCopy)
            End If
            Set clsThumb = Nothing
            Erase byImage
        Next lngY
        ' Is there any space left over at the right?
        With tR
             .Top = (lngI - 1) * m_lngItemTotalHeight - m_lngVerBarValue
             .Left = m_lngColCount * m_lngItemTotalWidth
             .Bottom = m_lngItemTotalHeight
             .Right = .Left + m_lngItemTotalWidth
        End With
        Call pvFillBackground(UserControl.hDC, tR)
        lngLastPos = lngI * m_lngItemTotalHeight - m_lngVerBarValue
    Next lngI
    
    ' Is there any space left over at the bottom?
    tR.Bottom = UserControl.Height
    If (lngLastPos < tR.Bottom) Then
       tR.Left = 0
       tR.Top = lngLastPos
       tR.Right = m_lAvailWidth + 32
       Call pvFillBackground(UserControl.hDC, tR)
    End If
    m_blnDirty = False
    Erase byImage
    Set clsThumb = Nothing
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Erase byImage
             Set clsThumb = Nothing
    End Select
End Sub

Private Function pvEnsureVisible(ByVal Index As Long) As Boolean
    Dim lngTop                               As Long
    Dim lngEnd                               As Long
    Dim lngOffset                            As Long
    Dim lngCurrRow                           As Long
    Const ProcName = "pvEnsureVisible"
    
    On Error GoTo ErrorHandle
    
    lngCurrRow = Index \ m_lngColCount
    If (Index Mod m_lngColCount) Then lngCurrRow = lngCurrRow + 1
    
    lngTop = (lngCurrRow - 1) * m_lngItemTotalHeight
    lngEnd = lngCurrRow * m_lngItemTotalHeight
    If lngCurrRow <= m_lngFirstRow Then
       If m_cScroll.Value(GpPvwVerticalBar) <> lngTop Then
          m_cScroll.Value(GpPvwVerticalBar) = lngTop
          pvEnsureVisible = True
          Exit Function
       End If
    End If
    
'    Debug.Print vbCrLf
'    Debug.Print "m_lngFirstRow---" & m_lngFirstRow
'    Debug.Print "m_lngLastRow---" & m_lngLastRow
'    Debug.Print "lngCurrRow --" & lngCurrRow
'    Debug.Print "m_cScroll.Value---" & m_cScroll.Value(GpPvwVerticalBar)
'    Debug.Print "m_lAvailheight---" & m_lAvailheight
'    Debug.Print "lngEnd---" & lngEnd
    
'    If lngCurrRow >= m_lngLastRow Then
       If m_cScroll.Value(GpPvwVerticalBar) < (lngEnd - m_lAvailheight) Then
          m_cScroll.Value(GpPvwVerticalBar) = lngEnd - m_lAvailheight
          pvEnsureVisible = True
       End If
'    End If
    
'    If lngTop < m_lngVerBarValue Then
'       If lngEnd > m_lngVerBarValue + m_lAvailheight Then
'          If m_lngItemTotalHeight < m_lAvailheight Then
'             lngOffset = lngEnd - (m_lngVerBarValue + m_lAvailheight) + 8
'             If m_cScroll.Value(GpPvwVerticalBar) + lngOffset >= 0 Then
'                If lngOffset <> 0 Then
'                   m_cScroll.Value(GpPvwVerticalBar) = m_cScroll.Value(GpPvwVerticalBar) + lngOffset
'                   pvEnsureVisible = True
'                Else
'                   pvEnsureVisible = False
'                End If
'             Else
'                pvEnsureVisible = False
'             End If
'          End If
'       Else
'          m_cScroll.Value(GpPvwVerticalBar) = lngTop
'          pvEnsureVisible = True
'       End If
'    Else
'       lngOffset = lngTop - m_lngVerBarValue
'       If lngOffset <> 0 Then
'          m_cScroll.Value(GpPvwVerticalBar) = m_cScroll.Value(GpPvwVerticalBar) + lngOffset
'          pvEnsureVisible = True
'       Else
'          pvEnsureVisible = False
'       End If
'    End If
    
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Function

Private Sub pvFillBackground(ByVal lhDC As Long, ByRef tR As RECT)
    Dim hBr                As Long
    
    If Not (m_blnEnabled) Then
       hBr = GetSysColorBrush(vbButtonFace And &H1F&)
    Else
       If (UserControl.BackColor And &H80000000) = &H80000000 Then
          hBr = GetSysColorBrush(UserControl.BackColor And &H1F&)
       Else
          hBr = CreateSolidBrush(TranslateColor(UserControl.BackColor))
       End If
    End If
    Call FillRect(lhDC, tR, hBr)
    Call DeleteObject(hBr)
End Sub

Private Function pvGetFontHandle(ByVal vFont As StdFont) As Long
    Dim tULF As LOGFONT
    
    Call OLEFontToLogFont(vFont, UserControl.hDC, tULF)
    pvGetFontHandle = CreateFontIndirect(tULF)
End Function

Private Sub pvGetStartEndRow()
    Dim lngMod               As Long
    Dim lngMiddle            As Long
    Dim lngLastMid           As Long
    Dim lngFindStart         As Long
    Dim lngFindCount         As Long
    Dim lngPrevious          As Long
    Dim lngNext              As Long
    Const ProcName = "pvGetStartEndRow"
    
    On Error GoTo ErrorHandle
    If m_lngRowCount <= 0 Then Exit Sub
    m_lngFirstRow = 0: m_lngLastRow = 0
    lngFindCount = m_lngRowCount
    lngFindStart = 1
    lngLastMid = 0
    Do
       lngMiddle = (lngFindStart + lngFindCount) \ 2
       If lngLastMid = lngMiddle Then lngMiddle = lngMiddle + 1
       If lngMiddle <= 1 Then
          m_lngFirstRow = 1
          Exit Do
       End If
       If lngMiddle >= m_lngRowCount Then
          m_lngFirstRow = m_lngRowCount
          Exit Do
       End If
       lngLastMid = lngMiddle
       
'       Debug.Print "Start lngMiddle " & lngMiddle
       
       If m_lngItemTotalHeight * lngMiddle > m_lngVerBarValue Then
          lngFindCount = lngMiddle
       Else
          lngFindStart = lngMiddle
       End If
       
       lngPrevious = m_lngItemTotalHeight * (lngMiddle - 1)
       lngNext = m_lngItemTotalHeight * lngMiddle
       
       If lngPrevious <= m_lngVerBarValue And lngNext > m_lngVerBarValue Then
          m_lngFirstRow = lngMiddle
          Exit Do
       End If
    Loop
    
'    Debug.Print "Find End Row..............."
    m_lngLastRow = m_lAvailheight \ m_lngItemTotalHeight
    If (m_lAvailheight Mod m_lngItemTotalHeight) Then m_lngLastRow = m_lngLastRow + 1
    m_lngLastRow = m_lngLastRow + m_lngFirstRow
    If m_lngLastRow > m_lngRowCount Then m_lngLastRow = m_lngRowCount
'    If m_lngItemTotalHeight * m_lngLastRow > (m_lngVerBarValue + m_lAvailheight) Then m_lngLastRow = m_lngLastRow - 1
'    lngFindCount = m_lngRowCount
'    lngFindStart = 1
'    lngLastMid = 0
'    Do
'       lngMod = (lngFindStart + lngFindCount) Mod 2
'       lngMiddle = (lngFindStart + lngFindCount) \ 2
'       If lngMod > 0 Then lngMiddle = lngMiddle + 1
'
'       If lngLastMid = lngMiddle Then lngMiddle = lngMiddle + 1
'       If lngMiddle <= 1 Then
'          m_lngLastRow = 1
'          Exit Do
'       End If
'       If lngMiddle >= m_lngRowCount Then
'          m_lngLastRow = m_lngRowCount
'          Exit Do
'       End If
'       lngLastMid = lngMiddle
       
'       Debug.Print "End lngMiddle " & lngMiddle
'       If m_lngItemTotalHeight * (lngMiddle - 1) > (m_lngVerBarValue + m_lAvailheight) Then
'          lngFindCount = lngMiddle
'       Else
'          lngFindStart = lngMiddle
'       End If
       
'       lngPrevious = m_lngItemTotalHeight * (lngMiddle - 2)
'       lngNext = m_lngItemTotalHeight * (lngMiddle - 1)
'       If lngPrevious <= (m_lngVerBarValue + m_lAvailheight) And lngNext > (m_lngVerBarValue + m_lAvailheight) Then
'          m_lngLastRow = lngMiddle
'          Exit Do
'       End If
'    Loop
    
'    Debug.Print "Start Row is ----->  " & m_lngFirstRow
'    Debug.Print "End Row is ----->  " & m_lngLastRow
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Function pvGetTextHeight() As Long
    Dim tR As RECT
    
    tR.Right = m_lngItemWidth
    Call DrawTextEx(m_hDC, "a" & vbNullChar, -1, tR, DT_CALCRECT Or DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, m_udtDrawTextParams)
    pvGetTextHeight = tR.Bottom - tR.Top
End Function

Private Sub pvScrollSetDirty(ByVal blnNoOptimise As Boolean)
    Static lngLastStartRow                       As Long
    Static lngLastEndRow                         As Long
    Static lngLastV                              As Long
    Dim lngI                                     As Long
    Dim lngY                                     As Long
    Dim lngCellStart                             As Long
    Dim lngRowCount                              As Long
    Dim lngV                                     As Long
    Dim lngToDirtyY                              As Long
    Dim lngYStart                                As Long
    Dim lngYEnd                                  As Long
    Dim tSR                                      As RECT
    Dim tR                                       As RECT
    Dim tJunk                                    As RECT
    Const ProcName = "pvScrollSetDirty"
    
    On Error GoTo ErrorHandle
    
    If m_clsPicitems.Count <= 0 Then Exit Sub
    
    Call pvScrollVisible
    Call pvGetStartEndRow
    
    If m_lngFirstRow < 1 Then m_lngFirstRow = 1
    If m_lngLastRow > m_clsPicitems.Count Then m_lngLastRow = m_clsPicitems.Count
    
    If (m_cScroll.Visible(GpPvwVerticalBar)) Then lngV = m_cScroll.Value(GpPvwVerticalBar)
    
    lngToDirtyY = Abs(lngLastStartRow - m_lngFirstRow) + 1
    If (Abs(lngLastEndRow - m_lngLastRow) + 1) > lngToDirtyY Then lngToDirtyY = (Abs(lngLastEndRow - m_lngLastRow) + 1)
    
    
    If Not (blnNoOptimise) Then
       'GetClientRect UserControl.hwnd, tR
       tR.Top = 0: tR.Bottom = 0: tR.Right = UserControl.ScaleWidth: tR.Bottom = UserControl.ScaleHeight
       If (Abs(lngLastV - lngV) < (tR.Bottom - tR.Top) \ 2) Then
          ' We can optimise using ScrollDC:
          LSet tSR = tR
             ' scrolling in Y
             If Sgn(lngLastV - lngV) = -1 Then
                lngYStart = m_lngLastRow
                lngRowCount = 0
                Do While lngRowCount < lngToDirtyY
                   lngYStart = lngYStart - 1
                   If lngYStart < 1 Then
                      Exit Do
                   Else
                      lngRowCount = lngRowCount + 1
                   End If
                Loop
                If (lngYStart < 1) Then lngYStart = 1
                lngYEnd = m_lngLastRow
                tSR.Top = tSR.Top - (lngLastV - lngV)
             Else
                lngYStart = m_lngFirstRow
                lngYEnd = m_lngFirstRow
                lngRowCount = 0
                Do While lngRowCount < lngToDirtyY
                   lngYEnd = lngYEnd + 1
                   If lngYEnd > m_clsPicitems.Count Then
                      Exit Do
                   Else
                      lngRowCount = lngRowCount + 1
                   End If
                Loop
                tSR.Bottom = tSR.Bottom - (lngLastV - lngV)
             End If
          If (lngYStart < 1) Then lngYStart = 1
          If (lngYEnd > m_clsPicitems.Count) Then lngYEnd = m_clsPicitems.Count
          ScrollDC UserControl.hDC, 0, lngLastV - lngV, tSR, tR, 0, tJunk
'          For lngI = lngYStart To lngYEnd
'              lngCellStart = (lngI - 1) * m_lngColCount
'              For lngY = 1 To m_lngColCount
'                  If (lngCellStart + lngY) > m_clsPicitems.Count Then Exit For
'                  m_clsPicitems.Item(lngCellStart + lngY).Dirty = True
'              Next lngY
'          Next lngI
       Else
          blnNoOptimise = True
       End If
    End If
    
    Call pvScrollVisible
    
'    If (blnNoOptimise) Then
       For lngI = m_lngFirstRow To m_lngLastRow
           lngCellStart = (lngI - 1) * m_lngColCount
           For lngY = 1 To m_lngColCount
               If (lngCellStart + lngY) > m_clsPicitems.Count Then Exit For
               m_clsPicitems.Item(lngCellStart + lngY).Dirty = True
           Next lngY
       Next lngI
'    End If
    
    lngLastStartRow = m_lngFirstRow
    lngLastEndRow = m_lngLastRow
    If (m_cScroll.Visible(GpPvwVerticalBar)) Then
       lngLastV = m_cScroll.Value(GpPvwVerticalBar)
    Else
       lngLastV = 0
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub pvScrollVisible()
    Dim tR                 As RECT
    Dim lngGridHeight      As Long
    Dim lngProportion      As Long
    Dim blnVert            As Boolean
    Const ProcName = "pvScrollVisible"
    
    On Error GoTo ErrorHandle
    
    Call GetWindowRect(UserControl.hWnd, tR)
    m_lAvailWidth = tR.Right - tR.Left - (UserControl.BorderStyle * 4)
    m_lAvailheight = tR.Bottom - tR.Top - (UserControl.BorderStyle * 4)
    
    m_lngColCount = m_lAvailWidth \ (DefaultSpaceBetweenItems + m_lngItemWidth)
    m_lngRowCount = m_clsPicitems.Count \ m_lngColCount
    If (m_clsPicitems.Count Mod m_lngColCount) Then m_lngRowCount = m_lngRowCount + 1
    
    lngGridHeight = ((DefaultSpaceBetweenItems + m_lngItemHeight)) * m_lngRowCount
    If lngGridHeight > m_lAvailheight Then blnVert = True
    
    If blnVert Then
       m_lAvailWidth = m_lAvailWidth - GetSystemMetrics(SM_CXVSCROLL)
       ' reset
       m_lngColCount = m_lAvailWidth \ (DefaultSpaceBetweenItems + m_lngItemWidth)
       m_lngRowCount = m_clsPicitems.Count \ m_lngColCount
       If (m_clsPicitems.Count Mod m_lngColCount) Then m_lngRowCount = m_lngRowCount + 1
       lngGridHeight = ((DefaultSpaceBetweenItems + m_lngItemHeight)) * m_lngRowCount
    End If
    
    m_cScroll.Visible(GpPvwVerticalBar) = blnVert
    
    ' Check scaling:
    m_lngVerBarValue = 0
    With m_cScroll
         If (blnVert) Then
            .Max(GpPvwVerticalBar) = lngGridHeight - m_lAvailheight
            If (m_lAvailheight > 0) Then
               lngProportion = ((lngGridHeight - m_lAvailheight) \ m_lAvailheight) + 1
               .LargeChange(GpPvwVerticalBar) = (lngGridHeight - m_lAvailheight) \ lngProportion
               .SmallChange(GpPvwVerticalBar) = m_lngItemTotalHeight
            End If
            m_lngVerBarValue = m_cScroll.Value(GpPvwVerticalBar)
         End If
    End With
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub pvSingleModeSelect(ByVal Index As Long)
    Dim lngI                   As Long
    
    m_blnRedraw = False
    Call pvClearSelection
    m_colSelected.Add Index
    Set m_clsSelectedItem = m_clsPicitems.Item(Index)
    With m_clsSelectedItem
         .AutoReDraw = False
         .Selected = True
         .AutoReDraw = True
    End With
    m_lngSelFirst = Index
    m_blnRedraw = True
End Sub

Public Property Get ReDraw() As Boolean
    ReDraw = m_blnRedraw
End Property

Public Property Let ReDraw(ByVal NewValue As Boolean)
    m_blnRedraw = NewValue
    If (UserControl.Ambient.UserMode) Then
       m_blnDirty = True
       Call pvScrollSetDirty(True)
       Call pvDraw
    End If
    PropertyChanged "Redraw"
End Property

Public Sub Refresh()
    Const ProcName = "Refresh"
    
    On Error GoTo ErrorHandle
    If m_blnRedraw And m_blnUserMode Then
       m_blnDirty = True
       Call pvScrollSetDirty(True)
       Call pvDraw
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Property Get SelectedCount() As Long
    On Error Resume Next
    SelectedCount = m_colSelected.Count
End Property

Public Property Get SelectedItem() As CPicItem
    If m_clsSelectedItem Is Nothing Then
       Set SelectedItem = Nothing
    Else
       Set SelectedItem = m_clsSelectedItem
    End If
    
End Property

Private Sub m_clsPicitems_AddNew(ByVal Item As CPicItem)
    Const ProcName = "m_clsPicitems_AddNew"
    
    On Error GoTo ErrorHandle
    If m_clsSelectedItem Is Nothing Then
       Set m_clsSelectedItem = Item
       With m_clsSelectedItem
            .AutoReDraw = False
            .Selected = True
            .AutoReDraw = True
       End With
       m_lngSelFirst = Item.Index
       Set m_colSelected = Nothing
       Set m_colSelected = New Collection
       m_colSelected.Add Item.Index
    End If
    If m_hDC = 0 Then Call pvBuildMemDC
    If m_blnRedraw Then
       Call pvScrollVisible
       Call pvDraw
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub m_clsPicitems_PicClear()
    Dim tR As RECT
    Const ProcName = "m_clsPicitems_PicClear"
    
    On Error GoTo ErrorHandle
    Call pvClearSelection
    Call pvScrollVisible
    Call GetClientRect(UserControl.hWnd, tR)
    Call pvFillBackground(UserControl.hDC, tR)
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub m_clsPicitems_PicRemove(ByVal Index As Long)
    Dim lngI As Long
    Dim lngIndex As Long
    Const ProcName = "m_clsPicitems_PicRemove"
    
    On Error GoTo ErrorHandle
    For lngI = 1 To SelectedCount
        lngIndex = m_colSelected(lngI)
        If lngIndex = Index Then
           m_colSelected.Remove lngI
           Exit For
        End If
    Next lngI
    Set m_colSelected = Nothing
    Set m_colSelected = New Collection
    For lngI = 1 To m_clsPicitems.Count
        With m_clsPicitems.Item(lngI)
             If .Selected Then m_colSelected.Add .Index
        End With
    Next lngI
    If m_clsSelectedItem Is Nothing Then
       If m_lngSelFirst > m_clsPicitems.Count Then m_lngSelFirst = m_clsPicitems.Count
    Else
       m_lngSelFirst = m_clsSelectedItem.Index
    End If
    If m_blnRedraw Then
       Call pvScrollVisible
       Call pvDraw
    End If
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub m_clsPicitems_SelectedChanged(ByVal Index As Long)
    Dim lngI As Long
    Const ProcName = "m_clsPicitems_SelectedChanged"
    
    On Error GoTo ErrorHandle
    If m_clsPicitems(Index).Selected Then
       If m_blnMultiSelect Then
          m_colSelected.Add Index
          m_lngSelFirst = Index
       Else
          Call pvSingleModeSelect(Index)
       End If
    Else
       For lngI = 1 To SelectedCount
           If m_colSelected(lngI) = Index Then
              m_colSelected.Remove lngI
              Exit For
           End If
       Next lngI
    End If
    If m_blnRedraw Then Call pvDraw
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub m_cScroll_Change(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD)
    Dim bRedraw As Boolean
    
    m_lngVerBarValue = m_cScroll.Value(ScrollBar)
    Call pvScrollSetDirty(False)
    Call pvDraw
End Sub

Private Sub m_cScroll_Scroll(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD)
    Call m_cScroll_Change(ScrollBar)
End Sub

Private Sub UserControl_Click()
    Const ProcName = "UserControl_Click"
    
    On Error GoTo ErrorHandle
    If m_blnEnabled Then RaiseEvent Click
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_DblClick()
    Const ProcName = "UserControl_DblClick"
    
    On Error GoTo ErrorHandle
    If m_blnEnabled Then RaiseEvent DblClick
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_GotFocus()
    Const ProcName = "UserControl_GotFocus"
    
    On Error GoTo ErrorHandle
    
    m_bInFocus = True
    Call pvScrollSetDirty(True)
    Call pvDraw
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_Initialize()
    Const ProcName = "UserControl_Initialize"
    
    On Error GoTo ErrorHandle
    Set m_clsPicitems = New CPicItems
    Set m_colSelected = New Collection
    m_udtBorderStyle = GpPvwBorderStyle3D
    m_blnRedraw = True
    m_blnHotTracking = False
    m_oleHighlightBackColor = vbHighlight
    m_oleHighlightForeColor = vbHighlightText
    m_blnHideSelection = False
    m_blnEnabled = True
    m_lngItemWidth = DefaultPicItemWidth
    m_lngItemHeight = DefaultPicItemHeight
    m_blnLoadThumbPack = False
    With m_udtDrawTextParams
         .iLeftMargin = 1
         .iRightMargin = 1
         .iTabLength = 1
         .cbSize = Len(m_udtDrawTextParams)
    End With
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_InitProperties()
    Const ProcName = "UserControl_InitProperties"
    
    On Error GoTo ErrorHandle
    Me.BackColor = DefaultBackColor
    Me.ForeColor = vbWindowText
    Set Me.Font = Ambient.Font
    Me.BorderStyle = GpPvwBorderStyle3D
    Me.HideSelection = False
    Me.Enabled = True
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngI                            As Long
    Dim lngStart                        As Long
    Dim lngItemCount                    As Long
    Dim lngCurrRowCount                 As Long
    Dim blnFound                        As Boolean
    Const ProcName = "UserControl_KeyDown"
    
    On Error GoTo ErrorHandle
    If Not (m_blnEnabled) Then
       Select Case KeyCode
              Case vbKeyUp
                If (m_cScroll.Visible(GpPvwVerticalBar)) Then m_cScroll.Value(GpPvwVerticalBar) = m_cScroll.Value(GpPvwVerticalBar) - m_cScroll.SmallChange(GpPvwVerticalBar)
              Case vbKeyDown
                If (m_cScroll.Visible(GpPvwVerticalBar)) Then m_cScroll.Value(GpPvwVerticalBar) = m_cScroll.Value(GpPvwVerticalBar) + m_cScroll.SmallChange(GpPvwVerticalBar)
              Case vbKeyPageUp
                If (m_cScroll.Visible(GpPvwVerticalBar)) Then m_cScroll.Value(GpPvwVerticalBar) = m_cScroll.Value(GpPvwVerticalBar) - m_cScroll.LargeChange(GpPvwVerticalBar)
              Case vbKeyPageDown
                If (m_cScroll.Visible(GpPvwVerticalBar)) Then m_cScroll.Value(GpPvwVerticalBar) = m_cScroll.Value(GpPvwVerticalBar) + m_cScroll.LargeChange(GpPvwVerticalBar)
       End Select
       Exit Sub
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
    lngItemCount = m_clsPicitems.Count
    If lngItemCount <= 0 Then Exit Sub
    Select Case KeyCode
           Case vbKeyA
             If Not m_blnMultiSelect Then Exit Sub
             If (Shift And vbCtrlMask) = vbCtrlMask Then
                lngStart = m_clsSelectedItem.Index
                m_blnRedraw = False
                Call pvClearSelection
                Set m_clsSelectedItem = m_clsPicitems.Item(lngStart)
                For lngI = 1 To lngItemCount
                    m_colSelected.Add lngI
                    With m_clsPicitems.Item(lngI)
                         .AutoReDraw = False
                         .Selected = True
                         .AutoReDraw = True
                    End With
                Next lngI
                m_blnRedraw = True
                Call pvDraw
                Exit Sub
             End If
           Case vbKeySpace
           Case vbKeyReturn Or vbKeySeparator
           Case vbKeyLeft
             lngStart = m_clsSelectedItem.Index
             If lngStart > 1 Then lngStart = lngStart - 1
             GoTo SelectedManual
           Case vbKeyRight
             lngStart = m_clsSelectedItem.Index
             If lngStart < lngItemCount Then lngStart = lngStart + 1
             GoTo SelectedManual
           Case vbKeyUp
             lngStart = m_clsSelectedItem.Index
             If lngStart > 1 Then
                lngStart = lngStart - m_lngColCount
                If lngStart <= 0 Then Exit Sub
             End If
             GoTo SelectedManual
           Case vbKeyDown
             lngStart = m_clsSelectedItem.Index
             If lngStart < lngItemCount Then
                lngStart = lngStart + m_lngColCount
                If lngStart > lngItemCount Then Exit Sub
             End If
             GoTo SelectedManual
           Case vbKeyPageUp
             lngStart = m_clsSelectedItem.Index
             If lngStart > 1 Then
                lngCurrRowCount = m_lAvailheight \ m_lngItemTotalHeight
                blnFound = False
                Do
                  If lngStart - m_lngColCount * lngCurrRowCount >= 1 Then
                     blnFound = True
                  Else
                     lngCurrRowCount = lngCurrRowCount - 1
                  End If
                Loop While Not blnFound
                lngStart = lngStart - m_lngColCount * lngCurrRowCount
             End If
             GoTo SelectedManual
           Case vbKeyPageDown
             lngStart = m_clsSelectedItem.Index
             If lngStart < lngItemCount Then
                lngCurrRowCount = m_lAvailheight \ m_lngItemTotalHeight
                blnFound = False
                Do
                  If lngItemCount >= lngStart + m_lngColCount * lngCurrRowCount Then
                     blnFound = True
                  Else
                     lngCurrRowCount = lngCurrRowCount - 1
                  End If
                Loop While Not blnFound
                lngStart = lngStart + m_lngColCount * lngCurrRowCount
             End If
             GoTo SelectedManual
           Case vbKeyHome
             lngStart = 1
             GoTo SelectedManual
           Case vbKeyEnd
             lngStart = lngItemCount
             GoTo SelectedManual
    End Select
    Exit Sub
SelectedManual:
    If (Shift And vbShiftMask) = vbShiftMask Then
       m_blnRedraw = False
       Call pvClearSelection
       If lngStart > m_lngSelFirst Then
          For lngI = m_lngSelFirst To lngStart
              m_colSelected.Add lngI
              With m_clsPicitems.Item(lngI)
                   .AutoReDraw = False
                   .Selected = True
                   .AutoReDraw = True
              End With
          Next lngI
       Else
          For lngI = lngStart To m_lngSelFirst
              m_colSelected.Add lngI
              With m_clsPicitems.Item(lngI)
                   .AutoReDraw = False
                   .Selected = True
                   .AutoReDraw = True
              End With
          Next lngI
       End If
       Set m_clsSelectedItem = m_clsPicitems.Item(lngStart)
       With m_clsSelectedItem
            .AutoReDraw = False
            .Selected = True
            .AutoReDraw = True
       End With
       m_blnRedraw = True
       If Not pvEnsureVisible(lngStart) Then Call pvDraw
    ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
       m_lngSelFirst = lngStart
       Set m_clsSelectedItem = m_clsPicitems.Item(m_lngSelFirst)
       m_clsSelectedItem.Dirty = True
       If Not pvEnsureVisible(m_lngSelFirst) Then Call pvDraw
    Else
       m_lngSelFirst = lngStart
       Call pvSingleModeSelect(m_lngSelFirst)
       If Not pvEnsureVisible(m_lngSelFirst) Then Call pvDraw
    End If
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Const ProcName = "UserControl_KeyPress"
    
    On Error GoTo ErrorHandle
    RaiseEvent KeyPress(KeyAscii)
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Const ProcName = "UserControl_KeyUp"
    
    On Error GoTo ErrorHandle
    RaiseEvent KeyUp(KeyCode, Shift)
    If m_clsPicitems.Count <= 0 Then Exit Sub
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
       RaiseEvent ItemClick(m_clsSelectedItem)
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_LostFocus()
    Const ProcName = "UserControl_LostFocus"
    
    On Error GoTo ErrorHandle
    m_bInFocus = False
    Call pvScrollSetDirty(True)
    Call pvDraw
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngSelStart                      As Long
    Dim lngI                             As Long
    Dim lngListCount                     As Long
    Dim clsTemp                          As CPicItem
    Const ProcName = "UserControl_MouseDown"
    
    On Error GoTo ErrorHandle
    If Not m_blnEnabled Then Exit Sub
    If m_clsPicitems.Count <= 0 Then Exit Sub
    RaiseEvent MouseDown(Button, Shift, x, y)
    Set clsTemp = HitTest(x, y)
    If clsTemp Is Nothing Then Exit Sub
    m_blnRedraw = False
    If (Shift And vbShiftMask) = vbShiftMask Then
       If Button = vbLeftButton Then
          If m_blnMultiSelect Then
             Call pvClearSelection
             If m_lngSelFirst > clsTemp.Index Then
                For lngI = clsTemp.Index To m_lngSelFirst
                    m_colSelected.Add lngI
                    With m_clsPicitems(lngI)
                         .AutoReDraw = False
                         .Selected = True
                         .AutoReDraw = True
                    End With
                Next lngI
             Else
                For lngI = m_lngSelFirst To clsTemp.Index
                    m_colSelected.Add lngI
                    With m_clsPicitems(lngI)
                         .AutoReDraw = False
                         .Selected = True
                         .AutoReDraw = True
                    End With
                Next lngI
             End If
             Set m_clsSelectedItem = clsTemp
             m_blnRedraw = True
             If Not pvEnsureVisible(m_clsSelectedItem.Index) Then Call pvDraw
             RaiseEvent ItemClick(m_clsSelectedItem)
          Else
             GoTo SingleSelected
          End If
       End If
    ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
       If Button = vbLeftButton Then
          If m_blnMultiSelect Then
             Set m_clsSelectedItem = clsTemp
             With m_clsSelectedItem
                  .AutoReDraw = False
                  .Selected = Not .Selected
                  .AutoReDraw = True
                  If .Selected Then
                     m_colSelected.Add .Index
                  Else
                     For lngI = 1 To SelectedCount
                         If m_colSelected(lngI) = .Index Then
                            m_colSelected.Remove lngI
                            Exit For
                         End If
                     Next lngI
                  End If
             End With
             m_lngSelFirst = clsTemp.Index
             m_blnRedraw = True
             If Not pvEnsureVisible(m_clsSelectedItem.Index) Then Call pvDraw
             RaiseEvent ItemClick(m_clsSelectedItem)
          Else
             GoTo SingleSelected
          End If
       End If
    Else
       If Button = vbLeftButton Then
          GoTo SingleSelected
       ElseIf Button = vbRightButton Then
          If Not clsTemp.Selected Then GoTo SingleSelected
       End If
    End If
    m_blnRedraw = True
    Set clsTemp = Nothing
    Exit Sub
SingleSelected:
    m_blnRedraw = True
    Call pvSingleModeSelect(clsTemp.Index)
    If Not pvEnsureVisible(clsTemp.Index) Then Call pvDraw
    RaiseEvent ItemClick(m_clsSelectedItem)
    Set clsTemp = Nothing
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             m_blnRedraw = True
             Set clsTemp = Nothing
    End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim clsTemp As CPicItem
    Const ProcName = "UserControl_MouseMove"
    
    On Error GoTo ErrorHandle
    If m_blnEnabled Then RaiseEvent MouseMove(Button, Shift, x, y)
    If m_blnHotTracking Then
       Set clsTemp = HitTest(x, y)
       If clsTemp Is Nothing Then
          m_lngHoverIndex = 0
       Else
          If clsTemp.Index <> m_lngHoverIndex Then
             m_lngHoverIndex = clsTemp.Index
             clsTemp.Dirty = True
          End If
       End If
       Call pvDraw
    End If
    ' Capture mouse
    If x >= 0 And x < UserControl.ScaleWidth And y >= 0 And y < UserControl.ScaleHeight And Button = 0 Then
        If GetCapture() <> UserControl.hWnd Then Call SetCapture(UserControl.hWnd)
    Else
        If GetCapture() = UserControl.hWnd And Button = 0 Then Call ReleaseCapture
    End If
    Set clsTemp = Nothing
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort: Set clsTemp = Nothing
    End Select
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Const ProcName = "UserControl_MouseUp"
    
    On Error GoTo ErrorHandle
    If m_blnEnabled Then RaiseEvent MouseUp(Button, Shift, x, y)
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_Paint()
    Const ProcName = "UserControl_Paint"
    
    On Error GoTo ErrorHandle
    If m_blnRedraw And m_blnUserMode Then
       m_blnDirty = True
       Call pvScrollSetDirty(True)
       Call pvDraw
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const ProcName = "UserControl_ReadProperties"
    
    On Error GoTo ErrorHandle
    If (UserControl.Ambient.UserMode) Then
       m_blnUserMode = True
       Set m_cScroll = New cScrollBars
       With m_cScroll
            .Create UserControl.hWnd
            .Scrollbars = sbVertical
            .Visible(GpPvwHorizontalBar) = False
            .Visible(GpPvwVerticalBar) = False
       End With
       m_hWndCtl = UserControl.hWnd
       Call SetProp(m_hWndCtl, gcObjectProp, ObjPtr(Me))
       m_clsPicitems.fInit m_hWndCtl
    Else
       m_blnUserMode = False
    End If
    
    With PropBag
         Me.BackColor = .ReadProperty("BackColor", vbWindowBackground)
         Me.BorderStyle = .ReadProperty("BorderStyle", GpPvwBorderStyle3D)
         Me.Enabled = .ReadProperty("Enabled", True)
         Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
         Me.HideSelection = .ReadProperty("HideSelection", False)
         Me.HighlightBackColor = .ReadProperty("HighlightBackColor", vbHighlight)
         Me.HighlightForeColor = .ReadProperty("HighlightForeColor", vbHighlightText)
         Me.HotTracking = .ReadProperty("HotTracking", False)
         Me.MultiSelect = .ReadProperty("MultiSelect", False)
         Me.PicItemHeight = .ReadProperty("PicItemHeight", DefaultPicItemHeight * Screen.TwipsPerPixelY)
         Me.PicItemWidth = .ReadProperty("PicItemWidth", DefaultPicItemWidth * Screen.TwipsPerPixelX)
         Set Me.Font = .ReadProperty("Font", Ambient.Font)
    End With
    Call UserControl_Resize
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_Resize()
    Dim lngWidth As Long
    Const ProcName = "UserControl_Resize"
    
    On Error GoTo ErrorHandle
    If m_blnRedraw And m_blnUserMode Then
       m_blnDirty = True
       Call pvScrollSetDirty(True)
       Call pvDraw
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_Show()
    Dim lngReturn             As Long
    Static blnNotFirst        As Boolean
    Const ProcName = "UserControl_Show"
    
    On Error GoTo ErrorHandle
    If Not (blnNotFirst) Then
       lngReturn = GetWindowLong(UserControl.hWnd, GWL_STYLE)
       lngReturn = lngReturn And Not (WS_HSCROLL Or WS_VSCROLL)
       SetWindowLong UserControl.hWnd, GWL_STYLE, lngReturn
       SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
       blnNotFirst = True
    End If
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_Terminate()
    Const ProcName = "UserControl_Terminate"
    
    On Error GoTo ErrorHandle
    
    Set m_clsSelectedItem = Nothing
    Set m_clsPicitems = Nothing
    Set m_cScroll = Nothing
    Set m_colSelected = Nothing
    If (m_hDC <> 0) Then
       If (m_hBmpOld <> 0) Then Call SelectObject(m_hDC, m_hBmpOld)
       If (m_hBmp <> 0) Then Call DeleteObject(m_hBmp)
       If (m_hFntOldDC <> 0) Then Call SelectObject(m_hDC, m_hFntOldDC)
       Call DeleteDC(m_hDC)
       m_hDC = 0
    End If
    
    If (m_hFntDC <> 0) Then
       Call DeleteObject(m_hFntDC)
       m_hFntDC = 0
    End If
    Call RemoveProp(m_hWndCtl, gcObjectProp)
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const ProcName = "UserControl_WriteProperties"
    
    On Error GoTo ErrorHandle
    With PropBag
         .WriteProperty "BackColor", Me.BackColor, vbWindowBackground
         .WriteProperty "BorderStyle", Me.BorderStyle, GpPvwBorderStyle3D
         .WriteProperty "Enabled", Me.Enabled, True
         .WriteProperty "ForeColor", Me.ForeColor, vbWindowText
         .WriteProperty "HideSelection", Me.HideSelection, False
         .WriteProperty "HighlightBackColor", Me.HighlightBackColor, vbHighlight
         .WriteProperty "HighlightForeColor", Me.HighlightForeColor, vbHighlightText
         .WriteProperty "HotTracking", Me.HotTracking, False
         .WriteProperty "MultiSelect", Me.MultiSelect, False
         .WriteProperty "Font", Font, Ambient.Font
         .WriteProperty "PicItemHeight", Me.PicItemHeight, DefaultPicItemHeight * Screen.TwipsPerPixelY
         .WriteProperty "PicItemWidth", Me.PicItemWidth, DefaultPicItemWidth * Screen.TwipsPerPixelX
    End With
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub
