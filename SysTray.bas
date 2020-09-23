Attribute VB_Name = "Module1"
Option Explicit
Dim mg As String
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal _
hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx _
As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global success As Long 'For holding the return value, Must be long in 32-Bit
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
' Some of this code is from the VB5 CD.


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

' menu declares
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lpReserved As Long) As Long

Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const GWL_WNDPROC = (-4&)
Public Const WM_COMMAND = &H111
Public Const WM_USER = &H400
Public Const TRAY_CALLBACK = (WM_USER + 101&)
Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_ICON = &H2&
Public Const NIF_TIP = &H4&

Public Const WM_MOUSEMOVE = &H200&
Public Const WM_LBUTTONDOWN = &H201&
Public Const WM_LBUTTONUP = &H202&
Public Const WM_LBUTTONDBLCLK = &H203&
Public Const WM_RBUTTONDOWN = &H204&
Public Const WM_RBUTTONUP = &H205&
Public Const WM_RBUTTONDBLCLK = &H206&

Public Const MF_STRING = &H0
Public Const WM_MENUBASE = &H2000
Public Const MF_SEPARATOR = &H800

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    X As Long
    Y As Long
End Type

Public PrevWndProc As Long
Public ObjPointer&
Private st As SysTray

Public Function SubWndProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  On Error Resume Next
 
  Select Case MSG
    Case WM_COMMAND And HiWord(wParam) = 0        ' handle menu messages ONLY
    
      If (wParam And WM_MENUBASE) = WM_MENUBASE Then ' only handle custom menus
        RtlMoveMemory st, ObjPointer, 4
          st.Clicked (wParam And 3)               ' find the state of the 2 lowest bits
        RtlMoveMemory st, 0&, 4
      End If
      
    Case TRAY_CALLBACK
        RtlMoveMemory st, ObjPointer, 4           ' Copy an unreferenced pointer to object into variable
          st.SendEvent lParam, wParam             ' Send windows message\user event to control
        RtlMoveMemory st, 0&, 4                   ' Nullify object pointer
        
    End Select
   
    ' Forward all messages to previous window procedure...(This must be done)
    SubWndProc = CallWindowProc(PrevWndProc, hWnd, MSG, wParam, lParam)
End Function

Private Function HiWord(ByVal dw As Long) As Integer
  HiWord = CLng(dw / 65536) And &HFFFF
End Function

