Attribute VB_Name = "AOT"
'This module isn't mine

 Option Explicit
 Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public Sub AlwaysOnTop(EnabledOrDisabled, FormID As Object)
     
     If EnabledOrDisabled = "Enabled" Then SetWindowPos FormID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
     
     If EnabledOrDisabled = "Disabled" Then SetWindowPos FormID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub



