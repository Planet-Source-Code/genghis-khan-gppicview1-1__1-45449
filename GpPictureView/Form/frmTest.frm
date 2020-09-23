VERSION 5.00
Object = "{59776DD8-DE88-4E75-8840-C5D912C28520}#2.0#0"; "PicView11.ocx"
Begin VB.Form frmTest 
   Caption         =   "Coded by Genghis Khan(GuangJian Guo)"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin PicView.GpPictureView GpPictureView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9340
      BackColor       =   10658466
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load ThumbPack"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate ThumbPack"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normal"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   5640
      Pattern         =   "*.jpg"
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long
    Dim strPath As String
    
    strPath = App.Path & "\Image"
    File1.Path = strPath
    With GpPictureView1
         .MultiSelect = True
         .PicItems.Clear
         .CanLoadThumbPack = False
         .ReDraw = False
         For i = 0 To File1.ListCount - 1
             .PicItems.AddFromFile strPath & "\" & File1.List(i), , File1.List(i), "600 x 400"
         Next i
         .ReDraw = True
    End With
End Sub

Private Sub Command2_Click()
    Dim lngI As Long
    Dim aryFile() As String
    Dim strPath As String
        
    strPath = "e:\temp"
    File1.Path = strPath
    ReDim aryFile(0 To File1.ListCount - 1)
    For lngI = 0 To File1.ListCount - 1
        aryFile(lngI) = strPath & "\" & File1.List(lngI)
    Next lngI
    GpPictureView1.CreateThumbImagePack App.Path & "\1.pak", aryFile
End Sub

Private Sub Command3_Click()
    GpPictureView1.CanLoadThumbPack = True
    GpPictureView1.LoadThumbPack App.Path & "\1.pak"
End Sub
