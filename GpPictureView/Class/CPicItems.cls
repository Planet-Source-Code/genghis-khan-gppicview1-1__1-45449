VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPicItems"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ======================================================================================
' Constants
' ======================================================================================
Private Const MODULE_NAME = "CPicItems"

' ======================================================================================
' Private variables:
' ======================================================================================
Private m_hWnd                    As Long
Private m_colItems                As Collection
Private m_clsPicItem              As CPicItem
Attribute m_clsPicItem.VB_VarHelpID = -1

' ======================================================================================
' Events
' ======================================================================================
Public Event AddNew(ByVal Item As CPicItem)
Public Event BackColorChanged(ByVal Index As Long)
Public Event CaptionChanged(ByVal Index As Long)
Public Event CaptionBackColorChanged(ByVal Index As Long)
Public Event ForeColorChanged(ByVal Index As Long)
Public Event PictureInfoChanged(ByVal Index As Long)
Public Event SelectedChanged(ByVal Index As Long)
Public Event PicClear()
Public Event PicRemove(ByVal Index As Long)

Public Function AddThumbItem(Optional ByVal Key As String = "", _
                             Optional ByVal Caption As String = "", _
                             Optional ByVal PictureInfo As String = "") As CPicItem
    Dim lngIndex            As Long
    Dim blnOk               As Boolean
    Const ProcName = "AddThumbItem"
    
'    On Error GoTo ErrorHandle
    If Not IsValidItemKey(Key) Then Exit Function
    Set m_clsPicItem = Nothing
    Set m_clsPicItem = New CPicItem
    m_clsPicItem.fInit m_hWnd
    lngIndex = m_colItems.Count + 1
    With m_clsPicItem
         .AutoReDraw = False
         .PicIndex = lngIndex
         .Key = Key
         .Caption = Caption
         .PictureInfo = PictureInfo
         .AutoReDraw = True
    End With
    m_colItems.Add m_clsPicItem
    Set AddThumbItem = m_clsPicItem
    Call pvPaintControl(m_clsPicItem)
    
    Exit Function
'ErrorHandle:
'    Select Case ShowError(ProcName, MODULE_NAME)
'           Case vbRetry: Resume
'           Case vbIgnore: Resume Next
'           Case vbAbort
'             Set m_clsPicItem = Nothing
'    End Select
End Function

Public Function AddFromFile(ByVal FileName As String, _
                            Optional ByVal Key As String = "", _
                            Optional ByVal Caption As String = "", _
                            Optional ByVal PictureInfo As String = "") As CPicItem
    Dim lngIndex            As Long
    Dim blnOk               As Boolean
    Const ProcName = "AddFromFile"
    
    On Error GoTo ErrorHandle
    If Not ExistFile(FileName) Then Exit Function
    If Not IsValidItemKey(Key) Then Exit Function
    Set m_clsPicItem = Nothing
    Set m_clsPicItem = New CPicItem
    m_clsPicItem.fInit m_hWnd
    lngIndex = m_colItems.Count + 1
    With m_clsPicItem
         .AutoReDraw = False
         .PicIndex = lngIndex
         .Key = Key
         .Caption = Caption
         .PictureInfo = PictureInfo
         .AutoReDraw = True
    End With
    
    blnOk = LoadJPG(m_clsPicItem.DIBSection, FileName)
    If Not blnOk Then
       blnOk = m_clsPicItem.DIBSection.CreateFromFile(FileName)
    End If
    If blnOk Then
       m_colItems.Add m_clsPicItem
       Set AddFromFile = m_clsPicItem
       Call pvPaintControl(m_clsPicItem)
    End If
    
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Set m_clsPicItem = Nothing
    End Select
End Function

Public Function AddFromHandle(ByVal hImage As Long, _
                              Optional ByVal Key As String = "", _
                              Optional ByVal Caption As String = "", _
                              Optional ByVal PictureInfo As String = "", _
                              Optional ByVal lSrcLeft As Long = 0, _
                              Optional ByVal lSrcTop As Long = 0, _
                              Optional ByVal lSrcWidth As Long = -1, _
                              Optional ByVal lSrcHeight As Long = -1) As CPicItem
    Dim lngIndex            As Long
    Dim blnOk               As Boolean
    Const ProcName = "AddFromHandle"
    
    On Error GoTo ErrorHandle
    If hImage = 0 Then Exit Function
    If Not IsValidItemKey(Key) Then Exit Function
    Set m_clsPicItem = Nothing
    Set m_clsPicItem = New CPicItem
    m_clsPicItem.fInit m_hWnd
    lngIndex = m_colItems.Count + 1
    With m_clsPicItem
         .AutoReDraw = False
         .PicIndex = lngIndex
         .Key = Key
         .Caption = Caption
         .PictureInfo = PictureInfo
         .AutoReDraw = True
    End With
    
    blnOk = m_clsPicItem.DIBSection.CreateFromHandle(hImage, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight)
    If blnOk Then
       m_colItems.Add m_clsPicItem
       Set AddFromHandle = m_clsPicItem
       Call pvPaintControl(m_clsPicItem)
    End If
    
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Set m_clsPicItem = Nothing
    End Select
End Function

Public Function AddFromPicture(ByVal Pic As StdPicture, _
                               Optional ByVal Key As String = "", _
                               Optional ByVal Caption As String = "", _
                               Optional ByVal PictureInfo As String = "", _
                               Optional ByVal lSrcLeft As Long = 0, _
                               Optional ByVal lSrcTop As Long = 0, _
                               Optional ByVal lSrcWidth As Long = -1, _
                               Optional ByVal lSrcHeight As Long = -1) As CPicItem
    Dim lngIndex            As Long
    Dim blnOk               As Boolean
    Const ProcName = "AddFromPictureBox"
    
    On Error GoTo ErrorHandle
    If Pic Is Nothing Then Exit Function
    If Not IsValidItemKey(Key) Then Exit Function
    Set m_clsPicItem = Nothing
    Set m_clsPicItem = New CPicItem
    m_clsPicItem.fInit m_hWnd
    lngIndex = m_colItems.Count + 1
    With m_clsPicItem
         .AutoReDraw = False
         .PicIndex = lngIndex
         .Key = Key
         .Caption = Caption
         .PictureInfo = PictureInfo
         .AutoReDraw = True
    End With
    
    blnOk = m_clsPicItem.DIBSection.CreateFromPicture(Pic, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight)
    If blnOk Then
       m_colItems.Add m_clsPicItem
       Set AddFromPicture = m_clsPicItem
       Call pvPaintControl(m_clsPicItem)
    End If
    
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Set m_clsPicItem = Nothing
    End Select
End Function

Public Function AddFromPtr(ByVal lPtr As Long, _
                           ByVal lSize As Long, _
                           Optional ByVal Key As String = "", _
                           Optional ByVal Caption As String = "", _
                           Optional ByVal PictureInfo As String = "") As CPicItem
    Dim lngIndex            As Long
    Dim blnOk               As Boolean
    Const ProcName = "AddFromFile"
    
    On Error GoTo ErrorHandle
    If lPtr = 0 Then Exit Function
    If Not IsValidItemKey(Key) Then Exit Function
    Set m_clsPicItem = Nothing
    Set m_clsPicItem = New CPicItem
    m_clsPicItem.fInit m_hWnd
    lngIndex = m_colItems.Count + 1
    With m_clsPicItem
         .AutoReDraw = False
         .PicIndex = lngIndex
         .Key = Key
         .Caption = Caption
         .PictureInfo = PictureInfo
         .AutoReDraw = True
    End With
    
    blnOk = LoadJPGFromPtr(m_clsPicItem.DIBSection, lPtr, lSize)
    If blnOk Then
       m_colItems.Add m_clsPicItem
       Set AddFromPtr = m_clsPicItem
       Call pvPaintControl(m_clsPicItem)
    End If
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
             Set m_clsPicItem = Nothing
    End Select
End Function

Public Sub Clear()
    Set m_clsPicItem = Nothing
    Set m_colItems = Nothing
    Set m_colItems = New Collection
    Call pvPaintControlDel(True)
End Sub

Public Function Count() As Long
    On Error Resume Next
    Count = m_colItems.Count
End Function

Friend Function fInit(ByVal hWnd As Long)
    m_hWnd = hWnd
End Function

Private Function GetItemIndex(ByVal Key As Variant) As Long
    Dim lngI                    As Long
    Dim lngIndex                As Long
    Const ProcName = "GetItemIndex"
    
    On Error GoTo ErrorHandle
    
    If IsNumeric(Key) Then
       lngIndex = CLng(Key)
       If lngIndex <= 0 Or lngIndex > m_colItems.Count Then
          Err.Raise 380, App.EXEName & ".CPicItems"
       End If
    Else
       Key = CStr(Key)
       For lngI = 1 To m_colItems.Count
           If (m_colItems(lngI).Key = Key) Then
              lngIndex = lngI
              Exit For
           End If
       Next lngI
       If lngIndex <= 0 Then
          Err.Raise 35601, App.EXEName & ".CPicItems"
       End If
    End If
    GetItemIndex = lngIndex
    
    Exit Function
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Function

Private Function IsValidItemIndex(ByVal Index As Long) As Boolean
    If Index = 0 Then
       Err.Raise 9, App.EXEName & ".CPicItems"
       Exit Function
    End If
    If m_colItems.Count > 0 Then
       If Index > m_colItems.Count + 1 Then
          Err.Raise 9, App.EXEName & ".CPicItems"
          Exit Function
       End If
    Else
       If Index >= 0 And Index <> 1 Then
          Err.Raise 9, App.EXEName & ".CPicItems"
          Exit Function
       End If
    End If
    IsValidItemIndex = True
End Function

Private Function IsValidItemKey(ByVal Key As String, Optional Index As Long = 0) As Boolean
    Dim lngI As Long
    
    If (Key <> "") Then
       For lngI = 1 To m_colItems.Count
          If lngI <> Index Then
             If (m_colItems(lngI).Key = Key) Then
                Err.Raise 457, App.EXEName & ".CPicItems"
                Exit Function
             End If
          End If
       Next
    End If
    IsValidItemKey = True
End Function

Public Property Get Item(ByVal Index As Variant) As CPicItem
Attribute Item.VB_UserMemId = 0
    Dim lngIndex As Long
    Const ProcName = "Item"
    
    On Error GoTo ErrorHandle
    
    lngIndex = GetItemIndex(Index)
    Set m_clsPicItem = m_colItems(Index)
    Set Item = m_clsPicItem
    
    Exit Property
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
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

Private Sub pvPaintControl(ByVal Item As CPicItem)
    Dim oCtl        As GpPictureView
    On Error Resume Next
    If pbVerify(oCtl) Then oCtl.fPaintControl True, Item
End Sub

Private Sub pvPaintControlDel(ByVal IsClear As Boolean, Optional ByVal Index As Long)
    Dim oCtl        As GpPictureView
    On Error Resume Next
    If IsClear Then
       If pbVerify(oCtl) Then oCtl.fPaintControlDel True
    Else
       If pbVerify(oCtl) Then oCtl.fPaintControlDel False, Index
    End If
End Sub

Public Sub Remove(ByVal Index As Variant)
    Dim lngI As Long
    Dim lngIndex As Long
    Dim blnReSort As Boolean
    Dim clsTmp As CPicItem
    Const ProcName = "Remove"
    
    On Error GoTo ErrorHandle
    
    If Me.Count <= 0 Then Exit Sub
    blnReSort = False
    lngIndex = GetItemIndex(Index)
    If lngIndex < m_colItems.Count Then blnReSort = True
    m_colItems.Remove lngIndex
    If Me.Count <= 0 Then
       Call pvPaintControlDel(True)
       Exit Sub
    End If
    If blnReSort Then
       For lngI = lngIndex To m_colItems.Count
           Set clsTmp = m_colItems(lngI)
           With clsTmp
                .PicIndex = lngI
                .Dirty = True
           End With
       Next lngI
    End If
    Call pvPaintControlDel(False, lngIndex)
    
    Exit Sub
ErrorHandle:
    Select Case ShowError(ProcName, MODULE_NAME)
           Case vbRetry: Resume
           Case vbIgnore: Resume Next
           Case vbAbort
    End Select
End Sub

Private Sub Class_Initialize()
    Set m_colItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_clsPicItem = Nothing
    Set m_colItems = Nothing
End Sub
