VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPicItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ======================================================================================
' Private variables:
' ======================================================================================
Private m_hWnd As Long

Private m_strCaption            As String
Private m_strPictureInfo        As String
Private m_strKey                As String
Private m_strTag                As String
Private m_lngRow                As Long
Private m_lngCol                As Long
Private m_lngIndex              As Long
Private m_lngWidth              As Long
Private m_blnDirty              As Boolean
Private m_blnSelected           As Boolean
Private m_blnAutoReDraw         As Boolean
Private m_oleForeColor          As OLE_COLOR
Private m_oleBackColor          As OLE_COLOR
Private m_oleCaptionBackColor   As OLE_COLOR
Private m_clsDIB                As cDIBSection

Friend Property Let AutoReDraw(ByVal NewValue As Boolean)
    m_blnAutoReDraw = NewValue
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oleBackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_oleBackColor = NewValue
    m_blnDirty = True
    If m_blnAutoReDraw Then pvPaintControl
End Property

Public Property Get Caption() As String
    Caption = m_strCaption
End Property

Public Property Let Caption(ByVal NewValue As String)
    If m_strCaption <> NewValue Then
       m_strCaption = NewValue
       m_blnDirty = True
       If m_blnAutoReDraw Then pvPaintControl
    End If
End Property

Public Property Get CaptionBackColor() As OLE_COLOR
    CaptionBackColor = m_oleCaptionBackColor
End Property

Public Property Let CaptionBackColor(ByVal NewValue As OLE_COLOR)
    m_oleCaptionBackColor = NewValue
    m_blnDirty = True
    If m_blnAutoReDraw Then pvPaintControl
End Property

Friend Property Get Col() As Long
    Col = m_lngCol
End Property

Friend Property Let Col(ByVal NewValue As Long)
    m_lngCol = NewValue
End Property

Friend Property Get DIBSection() As cDIBSection
    Set DIBSection = m_clsDIB
End Property

Friend Property Get Dirty() As Boolean
    Dirty = m_blnDirty
End Property

Friend Property Let Dirty(ByVal NewValue As Boolean)
    m_blnDirty = NewValue
End Property

Public Sub EnsureVisible()
    Dim oCtl        As GpPictureView
    On Error Resume Next
    If pbVerify(oCtl) Then oCtl.fPaintEnsureVisible m_lngIndex
End Sub

Friend Function fInit(ByVal hWnd As Long)
    m_hWnd = hWnd
End Function

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_oleForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_oleForeColor = NewValue
    m_blnDirty = True
    If m_blnAutoReDraw Then pvPaintControl
End Property

Public Property Get Index() As Long
    Index = m_lngIndex
End Property

Public Property Get Key() As String
    Key = m_strKey
End Property

Public Property Let Key(ByVal NewValue As String)
    Dim strOld As String
    
    If m_strKey <> NewValue Then
       strOld = m_strKey
       m_strKey = NewValue
       'RaiseEvent KeyChanged(strOld)
    End If
End Property

Friend Property Let PicIndex(ByVal NewValue As Long)
    m_lngIndex = NewValue
End Property

Public Property Get PictureInfo() As String
    PictureInfo = m_strPictureInfo
End Property

Public Property Let PictureInfo(ByVal NewValue As String)
    If m_strPictureInfo <> NewValue Then
       m_strPictureInfo = NewValue
       m_blnDirty = True
       If m_blnAutoReDraw Then pvPaintControl
    End If
End Property

Private Function pbVerify(ByRef ctlThis As GpPictureView) As Boolean
    Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         pbVerify = True
      Else
         gErr 1, "CPicItems"
      End If
   Else
      gErr 1, "CPicItems"
   End If
End Function

Private Sub pvPaintControl()
    Dim oCtl        As GpPictureView
    On Error Resume Next
    If pbVerify(oCtl) Then oCtl.fPaintControl False
End Sub

Private Sub pvPaintSelect(ByVal Index As Long)
    Dim oCtl        As GpPictureView
    On Error Resume Next
    If pbVerify(oCtl) Then oCtl.fPaintControlSelect Index
End Sub

Friend Property Get Row() As Long
    Row = m_lngRow
End Property

Friend Property Let Row(ByVal NewValue As Long)
    m_lngRow = NewValue
End Property

Public Property Get Selected() As Boolean
    Selected = m_blnSelected
End Property

Public Property Let Selected(ByVal NewValue As Boolean)
    If m_blnSelected <> NewValue Then
       m_blnSelected = NewValue
       m_blnDirty = True
       If m_blnAutoReDraw Then Call pvPaintSelect(m_lngIndex)
    End If
End Property

Public Property Get Tag() As String
    Tag = m_strTag
End Property

Public Property Let Tag(ByVal NewValue As String)
    m_strTag = NewValue
End Property

Private Sub Class_Initialize()
    m_blnDirty = True
    m_blnAutoReDraw = True
    Set m_clsDIB = New cDIBSection
    m_oleForeColor = CLR_NONE
    m_oleBackColor = CLR_NONE
    m_oleCaptionBackColor = CLR_NONE
End Sub

Private Sub Class_Terminate()
    Set m_clsDIB = Nothing
End Sub

