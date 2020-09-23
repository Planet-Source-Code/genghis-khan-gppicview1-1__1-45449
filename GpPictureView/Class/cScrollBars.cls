VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ======================================================================================
' Constants
' ======================================================================================
Private Const SB_BOTH = 3
Private Const SB_BOTTOM = 7
Private Const SB_CTL = 2
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

' Flat scroll bars:
Private Const WSB_PROP_CYVSCROLL = &H1&
Private Const WSB_PROP_CXHSCROLL = &H2&
Private Const WSB_PROP_CYHSCROLL = &H4&
Private Const WSB_PROP_CXVSCROLL = &H8&
Private Const WSB_PROP_CXHTHUMB = &H10&
Private Const WSB_PROP_CYVTHUMB = &H20&
Private Const WSB_PROP_VBKGCOLOR = &H40&
Private Const WSB_PROP_HBKGCOLOR = &H80&
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const WSB_PROP_WINSTYLE = &H400&
Private Const WSB_PROP_PALETTE = &H800&
Private Const WSB_PROP_MASK = &HFFF&

Private Const FSB_FLAT_MODE = 2&
Private Const FSB_ENCARTA_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

' ======================================================================================
' Methods
' ======================================================================================
Public Enum SCROLLBARS_ORIENTATION_METHOD
    sbHorizontal = 0
    sbVertical = 1
    sbBoth = 2
End Enum

Public Enum GPPVW_SELECTSCROLLBAR_METHOD
    GpPvwHorizontalBar = SB_HORZ
    GpPvwVerticalBar = SB_VERT
End Enum

' ======================================================================================
' Types
' ======================================================================================
Private Type SCROLLINFO
    cbSize             As Long
    fMask              As Long
    nMin               As Long
    nMax               As Long
    nPage              As Long
    nPos               As Long
    nTrackPos          As Long
End Type

' ======================================================================================
' API declares:
' ======================================================================================
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollRange Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal LPINT1 As Long, ByVal LPINT2 As Long) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_GetScrollPos Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long) As Long
Private Declare Function FlatSB_GetScrollProp Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal propIndex As Long, ByVal LPINT As Long) As Long
Private Declare Function FlatSB_SetScrollPos Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal pos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal Min As Long, ByVal Max As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function InitialiseFlatSB Lib "COMCTL32.DLL" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function InitializeFlatSB Lib "COMCTL32.DLL" (ByVal hwnd As Long) As Long
Private Declare Function UninitializeFlatSB Lib "COMCTL32.DLL" (ByVal hwnd As Long) As Long

' ======================================================================================
' Private variables:
' ======================================================================================

' Message response:
Implements ISubclass
Private m_emr                       As EMsgResponse
Private m_hWnd                      As Long
Private m_lngSmallChangeHorz        As Long
Private m_lngSmallChangeVert        As Long
Private m_blnInitialised            As Boolean
Private m_blnFlatScrollBar          As Boolean
Private m_blnCanFlatScrollBar       As Boolean
Private m_blnEnabledHorz            As Boolean
Private m_blnEnabledVert            As Boolean
Private m_blnVisibleHorz            As Boolean
Private m_blnVisibleVert            As Boolean
Private m_ScrollBars                As SCROLLBARS_ORIENTATION_METHOD
' ======================================================================================
' Events
' ======================================================================================
Public Event Scroll(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD)
Public Event Change(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD)

Private Sub AttachMessages()
   If (m_hWnd <> 0) Then
      Call AttachMessage(Me, m_hWnd, WM_HSCROLL)
      Call AttachMessage(Me, m_hWnd, WM_VSCROLL)
      m_blnInitialised = True
   End If
End Sub

Private Sub ClearUp()
    If m_hWnd <> 0 Then
       On Error Resume Next
       ' Stop flat scroll bar if we have it:
       If m_blnCanFlatScrollBar Then UninitializeFlatSB m_hWnd
       On Error GoTo 0
       ' Remove subclass:
       Call DetachMessage(Me, m_hWnd, WM_HSCROLL)
       Call DetachMessage(Me, m_hWnd, WM_VSCROLL)
    End If
    m_hWnd = 0
    m_blnInitialised = False
End Sub

Friend Sub Create(ByVal hWndA As Long)
    Call ClearUp
    m_hWnd = hWndA
    Call CreateScrollBar
    Call AttachMessages
End Sub

Private Sub CreateScrollBar()
    Dim lngReturn            As Long
    Dim lngParent            As Long

   ' Redraw the window so it will take (pinched from MFC):
   lngParent = GetParent(m_hWnd)
   lngReturn = SetParent(m_hWnd, lngParent)
   ' Redraw for added insurance ('Slimy Windows Hacks')
   lngReturn = SetWindowPos(m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, (SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOMOVE Or SWP_DRAWFRAME))
   
   On Error Resume Next
   lngReturn = InitialiseFlatSB(m_hWnd)
   If (Err.Number <> 0) Then
      ' Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
      ' Means we have version prior to 4.71
      ' We get standard scroll bars.
      m_blnCanFlatScrollBar = False
   Else
      m_blnCanFlatScrollBar = True
      Me.FlatScrollBar = False
   End If
End Sub

Friend Property Get Enabled(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Boolean
    If ScrollBar = GpPvwHorizontalBar Then
       Enabled = m_blnEnabledHorz
    Else
       Enabled = m_blnEnabledVert
    End If
End Property

Friend Property Let Enabled(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal NewValue As Boolean)
    Dim lngEnable           As Long
    Dim lngReturn           As Long
    
    If ScrollBar = GpPvwHorizontalBar Then
       m_blnEnabledHorz = NewValue
    Else
       m_blnEnabledVert = NewValue
    End If
    If (NewValue) Then
       lngEnable = ESB_ENABLE_BOTH
    Else
       lngEnable = ESB_DISABLE_BOTH
    End If
    If (m_blnCanFlatScrollBar) Then
       lngReturn = FlatSB_EnableScrollBar(m_hWnd, ScrollBar, lngEnable)
    Else
       lngReturn = EnableScrollBar(m_hWnd, ScrollBar, lngEnable)
    End If
End Property

Friend Property Get FlatScrollBar() As Boolean
    FlatScrollBar = m_blnFlatScrollBar
End Property

Friend Property Let FlatScrollBar(ByVal NewValue As Boolean)
    Dim lngReturn As Long
    
    If m_blnCanFlatScrollBar Then
       If (m_ScrollBars = sbHorizontal) Or (m_ScrollBars = sbBoth) Then
          lngReturn = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, IIf(NewValue, FSB_FLAT_MODE, FSB_REGULAR_MODE), True)
       End If
       If (m_ScrollBars = sbVertical) Or (m_ScrollBars = sbBoth) Then
          lngReturn = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, IIf(NewValue, FSB_FLAT_MODE, FSB_REGULAR_MODE), True)
       End If
       m_blnFlatScrollBar = NewValue
    End If
End Property

Private Sub GetSI(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim lO        As Long
    Dim lngReturn As Long
    
    lO = ScrollBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    If (m_blnCanFlatScrollBar) Then
        lngReturn = FlatSB_GetScrollInfo(m_hWnd, lO, tSI)
    Else
        lngReturn = GetScrollInfo(m_hWnd, lO, tSI)
    End If
End Sub

Friend Property Get LargeChange(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Long
    Dim tSI As SCROLLINFO
    
    Call GetSI(ScrollBar, tSI, SIF_PAGE)
    LargeChange = tSI.nPage
End Property

Friend Property Let LargeChange(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal NewValue As Long)
    Dim tSI                     As SCROLLINFO
    
    Call GetSI(ScrollBar, tSI, SIF_ALL)
    tSI.nMax = tSI.nMax - tSI.nPage + NewValue
    tSI.nPage = NewValue
    Call LetSI(ScrollBar, tSI, SIF_PAGE Or SIF_RANGE)
End Property

Private Sub LetSI(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim lO        As Long
    Dim lngReturn As Long
    
    lO = ScrollBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    If m_blnCanFlatScrollBar Then
        lngReturn = FlatSB_SetScrollInfo(m_hWnd, lO, tSI, True)
    Else
        lngReturn = SetScrollInfo(m_hWnd, lO, tSI, True)
    End If
End Sub

Friend Property Get Max(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Long
    Dim tSI As SCROLLINFO
    
    Call GetSI(ScrollBar, tSI, SIF_RANGE Or SIF_PAGE)
    Max = tSI.nMax - tSI.nPage
End Property

Friend Property Let Max(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal NewValue As Long)
    Dim tSI As SCROLLINFO
    tSI.nMax = NewValue + Me.LargeChange(ScrollBar)
    tSI.nMin = Me.Min(ScrollBar)
    Call LetSI(ScrollBar, tSI, SIF_RANGE)
End Property

Friend Property Get Min(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Long
    Dim tSI As SCROLLINFO
    
    Call GetSI(ScrollBar, tSI, SIF_RANGE)
    Min = tSI.nMin
End Property

Friend Property Let Min(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal NewValue As Long)
    Dim tSI As SCROLLINFO
    
    tSI.nMin = NewValue
    tSI.nMax = Me.Max(ScrollBar) + Me.LargeChange(ScrollBar)
    Call LetSI(ScrollBar, tSI, SIF_RANGE)
End Property

Private Function RaiseEvents(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal blnScroll As Boolean)
    Static s_lLastValue(0 To 1) As Long
    
    If (Value(ScrollBar) <> s_lLastValue(ScrollBar)) Then
       If (blnScroll) Then
          RaiseEvent Scroll(ScrollBar)
       Else
          RaiseEvent Change(ScrollBar)
       End If
       s_lLastValue(ScrollBar) = Value(ScrollBar)
    End If
End Function

Friend Property Get Scrollbars() As SCROLLBARS_ORIENTATION_METHOD
    Scrollbars = m_ScrollBars
End Property

Friend Property Let Scrollbars(ByVal NewValue As SCROLLBARS_ORIENTATION_METHOD)
    Dim lngReturn As Long
    
    m_ScrollBars = NewValue
    lngReturn = ShowScrollBar(m_hWnd, SB_HORZ, Abs((m_ScrollBars = sbBoth) Or (m_ScrollBars = sbHorizontal)))
    lngReturn = ShowScrollBar(m_hWnd, SB_VERT, Abs((m_ScrollBars = sbBoth) Or (m_ScrollBars = sbVertical)))
End Property

Friend Property Get SmallChange(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Long
    If (ScrollBar = GpPvwHorizontalBar) Then
       SmallChange = m_lngSmallChangeHorz
    Else
       SmallChange = m_lngSmallChangeVert
    End If
End Property

Friend Property Let SmallChange(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal lngSmallChange As Long)
    If (ScrollBar = GpPvwHorizontalBar) Then
       m_lngSmallChangeHorz = lngSmallChange
    Else
       m_lngSmallChangeVert = lngSmallChange
    End If
End Property

Friend Property Get Value(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Long
    Dim tSI As SCROLLINFO
    
    Call GetSI(ScrollBar, tSI, SIF_POS)
    Value = tSI.nPos
End Property

Friend Property Let Value(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal NewValue As Long)
    Dim tSI As SCROLLINFO
    
    If (NewValue <> Me.Value(ScrollBar)) Then
        tSI.nPos = NewValue
        Call LetSI(ScrollBar, tSI, SIF_POS)
        Call RaiseEvents(ScrollBar, False)
    End If
End Property

Friend Property Get Visible(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD) As Boolean
    If (ScrollBar = GpPvwHorizontalBar) Then
       Visible = m_blnVisibleHorz
    Else
       Visible = m_blnVisibleVert
    End If
End Property

Friend Property Let Visible(ByVal ScrollBar As GPPVW_SELECTSCROLLBAR_METHOD, ByVal NewValue As Boolean)
    Dim lngReturn As Long
    
    If (ScrollBar = GpPvwHorizontalBar) Then
       m_blnVisibleHorz = NewValue
    Else
       m_blnVisibleVert = NewValue
    End If
    If (m_blnCanFlatScrollBar) Then
       lngReturn = FlatSB_ShowScrollBar(m_hWnd, ScrollBar, Abs(NewValue))
    Else
       lngReturn = ShowScrollBar(m_hWnd, ScrollBar, Abs(NewValue))
    End If
End Property

Private Sub Class_Initialize()
    m_lngSmallChangeHorz = 1
    m_lngSmallChangeVert = 1
    m_ScrollBars = sbBoth
End Sub

Private Sub Class_Terminate()
    Call ClearUp
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lngScrollCode                 As Long
    Dim lV                            As Long
    Dim lSC                           As Long
    Dim tSI                           As SCROLLINFO
    Dim ScrollBar                     As GPPVW_SELECTSCROLLBAR_METHOD
    
    On Error GoTo ErrorHandle
    Select Case iMsg
           Case WM_VSCROLL, WM_HSCROLL
             If (iMsg = WM_HSCROLL) Then
                ScrollBar = GpPvwHorizontalBar
             Else
                ScrollBar = GpPvwVerticalBar
             End If
             lngScrollCode = (wParam And &HFFFF&)
             Select Case lngScrollCode
                    Case SB_THUMBTRACK
                      Call GetSI(ScrollBar, tSI, SIF_TRACKPOS)
                      Me.Value(ScrollBar) = tSI.nTrackPos
                      Call RaiseEvents(ScrollBar, True)
                    Case SB_LEFT, SB_BOTTOM
                      Me.Value(ScrollBar) = Me.Min(ScrollBar)
                      Call RaiseEvents(ScrollBar, False)
                    Case SB_RIGHT, SB_TOP
                      Me.Value(ScrollBar) = Me.Max(ScrollBar)
                      Call RaiseEvents(ScrollBar, False)
                    Case SB_LINELEFT, SB_LINEUP
                      'Debug.Print "Line"
                      lV = Me.Value(ScrollBar)
                      If (ScrollBar = GpPvwHorizontalBar) Then
                         lSC = m_lngSmallChangeHorz
                      Else
                         lSC = m_lngSmallChangeVert
                      End If
                      If (lV - lSC < Me.Min(ScrollBar)) Then
                         Me.Value(ScrollBar) = Me.Min(ScrollBar)
                      Else
                         Me.Value(ScrollBar) = lV - lSC
                      End If
                      Call RaiseEvents(ScrollBar, False)
                    Case SB_LINERIGHT, SB_LINEDOWN
                      'Debug.Print "Line"
                      lV = Me.Value(ScrollBar)
                      If (ScrollBar = GpPvwHorizontalBar) Then
                         lSC = m_lngSmallChangeHorz
                      Else
                         lSC = m_lngSmallChangeVert
                      End If
                      If (lV + lSC > Me.Max(ScrollBar)) Then
                         Me.Value(ScrollBar) = Me.Max(ScrollBar)
                      Else
                         Me.Value(ScrollBar) = lV + lSC
                      End If
                      Call RaiseEvents(ScrollBar, False)
                    Case SB_PAGELEFT, SB_PAGEUP
                      Me.Value(ScrollBar) = Me.Value(ScrollBar) - Me.LargeChange(ScrollBar)
                      Call RaiseEvents(ScrollBar, False)
                    Case SB_PAGERIGHT, SB_PAGEDOWN
                      Me.Value(ScrollBar) = Me.Value(ScrollBar) + Me.LargeChange(ScrollBar)
                      Call RaiseEvents(ScrollBar, False)
                    Case SB_ENDSCROLL
                     Call RaiseEvents(ScrollBar, False)
             End Select
    End Select
    Exit Function
ErrorHandle:
    '
End Function

