Attribute VB_Name = "modErrorHandling"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function ShowError(ByVal strFunc As String, ByVal strModule As String) As VbMsgBoxResult
    Dim lngErrNumber             As Long
    Dim strErrDescription        As String
    Dim strErrSource             As String
    
    lngErrNumber = Err.Number
    strErrDescription = Err.Description
    strErrSource = IIf(Len(strModule) > 0, _
                            "[\\" & ErrComputerName() & "] " & _
                            App.EXEName & "." & _
                            strModule & "." & _
                            strFunc & _
                            IIf(Erl <> 0, "(" & Erl & ")", ""), "") & "--" & Err.Source
    ShowError = MsgBox( _
            strErrDescription & vbCrLf & vbCrLf & _
            "Error: 0x" & Hex(lngErrNumber) & vbCrLf & vbCrLf & _
            "Call stack:" & vbCrLf & _
            strErrSource, vbCritical Or vbAbortRetryIgnore, "Error")
End Function

Private Function ErrComputerName() As String
    Static sName        As String
        
    If Len(sName) = 0 Then
        sName = String(256, 0)
        GetComputerName sName, Len(sName)
        sName = Left$(sName, InStr(sName, Chr(0)) - 1)
    End If
    ErrComputerName = sName
End Function


