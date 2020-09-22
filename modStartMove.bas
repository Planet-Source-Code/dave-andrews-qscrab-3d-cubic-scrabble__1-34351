Attribute VB_Name = "modStartMove"
Option Explicit

Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const HTCAPTION& = 2
Private Const WM_NCLBUTTONDOWN& = &HA1
Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function ReleaseCapture& Lib "user32" ()
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   
Public MyMousePos As POINTAPI
Private Const GWL_EXSTYLE& = (-20)
Private Const GWL_STYLE& = (-16)
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)

Public Sub StartMove(frm As Form)
ReleaseCapture
SendMessageBynum frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

