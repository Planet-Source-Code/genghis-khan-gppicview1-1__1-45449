Attribute VB_Name = "modCallbacks"
Option Explicit

' ======================================================================================
' Constants
' ======================================================================================
Public Const gcObjectProp = "GpPictureView:ObjectPtr"

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0

Private Const MODULE_NAME = "modCallbacks"
' ======================================================================================
' Types:
' ======================================================================================

' ======================================================================================
' API declares:
' ======================================================================================
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public Function ExistFile(ByVal FileName As String) As Boolean
    On Error Resume Next
    Call FileLen(FileName)
    ExistFile = (Err = 0)
End Function

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oTemp As Object
    ' Turn the pointer into an illegal, uncounted interface
    Call CopyMemory(oTemp, lPtr, 4)
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oTemp
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    Call CopyMemory(oTemp, 0&, 4)
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but it will be because of the subclass, not the uncounted reference
End Property

Public Sub OLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
    Dim sFont As String
    Dim iChar As Integer
    
    ' Convert an OLE StdFont to a LOGFONT structure:
    On Error Resume Next
    With tLF
         sFont = fntThis.Name
         ' There is a quicker way involving StrConv and CopyMemory, but
         ' this is simpler!:
         For iChar = 1 To Len(sFont)
             .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
         Next iChar
         ' Based on the Win32SDK documentation:
         .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
         .lfItalic = fntThis.Italic
         If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
         Else
            .lfWeight = FW_NORMAL
         End If
         .lfUnderline = fntThis.Underline
         .lfStrikeOut = fntThis.Strikethrough
         .lfCharSet = fntThis.Charset
    End With
End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_NONE
End Function

' Returns Color as long, accepts SystemColorConstants
Public Function VerifyColor(ByVal ColorVal As Long) As Long
    VerifyColor = ColorVal
    If ColorVal > &HFFFFFF Or ColorVal < 0 Then VerifyColor = GetSysColor(ColorVal And &HFFFFFF)
End Function

Public Sub GetFitSize(ByVal dib As cDIBSection, _
                      ByVal WidthMax As Long, ByVal HeightMax As Long, _
                      ByRef RealWidth As Long, ByRef RealHeight As Long)
    Dim sngScale      As Single
    With dib
         If .Width > WidthMax Or .Height > HeightMax Then
            If .Width > .Height Then
               sngScale = .Width / WidthMax
               If .Height / sngScale > HeightMax Then sngScale = .Height / HeightMax
            Else
               sngScale = .Height / HeightMax
               If .Width / sngScale > WidthMax Then sngScale = .Width / WidthMax
            End If
            RealWidth = CLng(.Width / sngScale)
            RealHeight = CLng(.Height / sngScale)
         Else
            RealWidth = .Width
            RealHeight = .Height
         End If
    End With
End Sub

Public Sub gErr(ByVal lErrNum As Long, ByVal sSource As String)
Dim sDesc As String
'Debug.Assert False
   Select Case lErrNum
   Case 1
      ' Cannot find owner object
      lErrNum = 364
      sDesc = "Object has been unloaded."
   Case 2
      ' Bar does not exist
      lErrNum = vbObjectError + 25001
      sDesc = "PicItem does not exist."
   Case 3
      ' Invalid key: numeric
      lErrNum = 13
      sDesc = "Type Mismatch."
      
   Case 4
      ' Invalid Key: duplicate
      lErrNum = 457
      sDesc = "This key is already associated with an element of this collection."
   
   Case 5
      ' Subscript out of range
      lErrNum = 9
      sDesc = "Subscript out of range."
   
   Case Else
      Debug.Assert "Unexpected Error" = ""
   
   End Select
   
   
   Err.Raise lErrNum, App.EXEName & "." & sSource, sDesc
End Sub

