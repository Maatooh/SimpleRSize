VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MRSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
 
Private Type POINTAPI
    x As Long
    y As Long
End Type
 
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
 
Private Declare Function SetWindowPlacement Lib "user32.dll" ( _
                        ByVal hwnd As Long, _
                        ByRef lpwndpl As WINDOWPLACEMENT) As Long
 
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
                        
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
        ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Private Declare Function SetActiveWindow Lib "user32.dll" ( _
                        ByVal hwnd As Long) As Long
 
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMINIMIZED As Long = 2
Private Const SW_SHOWMAXIMIZED As Long = 3

Public Function MScaling(Ghwnd As Long, Px As Long, Py As Long, Ph As Long, Pw As Long)

On Error GoTo MyError
    Dim lRet As Long
    Dim lRet2 As Long
    Dim wPlace As WINDOWPLACEMENT
    Dim rRect As RECT
    
    lRet = Ghwnd
    wPlace.Length = Len(wPlace)
    wPlace.showCmd = SW_SHOWNORMAL
    rRect.Left = Px
    rRect.Top = Py
    rRect.Right = Ph '(Screen.Width / Screen.TwipsPerPixelX) / 2
    rRect.Bottom = Pw '(Screen.Height / Screen.TwipsPerPixelY) / 2
    wPlace.rcNormalPosition = rRect
    lRet2 = SetWindowPlacement(lRet, wPlace)
    SetActiveWindow lRet
    Exit Function
MyError:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function GetHwndMaatooh(THwnd)
Dim cWnd As Long, pWnd As Long
Dim Title As String * 255
' pWnd is the container's hWnd
Do
    cWnd = FindWindowEx(pWnd, cWnd, vbNullString, vbNullString)
    If cWnd Then
    
         'Debug.Print "Child WIndow: " & cWnd
         GetWindowText cWnd, Title, Len(Title)
         If Not InStr(Title, THwnd) = 0 Then
         'Debug.Print cWnd
         GetHwndMaatooh = cWnd
         Exit Function
         End If
    Else
         Exit Do
    End If
Loop
End Function
