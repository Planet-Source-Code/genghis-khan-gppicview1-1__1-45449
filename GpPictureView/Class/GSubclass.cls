VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'
' From ssubtmr6.dll
' Copyright © 1999 Steve McMahon

Public Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    modSubclass.AttachMessage iwp, hWnd, iMsg
End Sub

Public Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    modSubclass.DetachMessage iwp, hWnd, iMsg
End Sub

Public Property Get CurrentMessage() As Long
   CurrentMessage = modSubclass.CurrentMessage
End Property

Public Function CallOldWindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
   CallOldWindowProc = modSubclass.CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function



