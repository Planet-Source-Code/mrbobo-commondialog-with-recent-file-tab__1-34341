Attribute VB_Name = "ModGen"
Option Explicit
Public Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const DC_ACTIVE = &H1
Public Const DC_ICON = &H4
Public Const DC_TEXT = &H8
Public Const DC_GRADIENT = &H20
Public Const WM_COPY = &H301
Public Sub DoDrag(mHwnd As Long)
    ReleaseCapture
    Call SendMessage(mHwnd, &HA1, 2, 0&)
End Sub
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

