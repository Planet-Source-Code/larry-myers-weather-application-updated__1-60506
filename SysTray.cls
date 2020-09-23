VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This code was taken from the VB5 CD.  It originally was a control and _
' not a class.  I prefer classes so the code is compiled into the EXE.

Private gInTray As Boolean
Private gTrayId As Long
Private gTrayTip As String
Private gTrayHwnd As Long
Private gTrayIcon As StdPicture
Private gAddedToTray As Boolean
Private gFormhWnd As Long

Private Const defInTray = False
Private Const defTrayTip = "VB 5 - SysTray Control." & vbNullChar
Private Const sInTray = "InTray"
Private Const sTrayIcon = "TrayIcon"
Private Const sTrayTip = "TrayTip"

Public Event MouseMove(Id As Long)
Public Event MouseDown(Button As Integer, Id As Long)
Public Event MouseUp(Button As Integer, Id As Long)
Public Event MouseDblClick(Button As Integer, Id As Long)
Public Event MenuClick(index As Integer)

Public Property Set TrayIcon(Icon As StdPicture)
    Dim Tray As NOTIFYICONDATA                          ' Notify Icon Data structure
    Dim rc As Long                                      ' API return code

    If Not (Icon Is Nothing) Then                       ' If icon is valid...
        If (Icon.Type = vbPicTypeIcon) Then             ' Use ONLY if it is an icon
            If gAddedToTray Then                        ' Modify tray only if it is in use.
                Tray.uID = gTrayId                      ' Unique ID for each HWND and callback message.
                Tray.hwnd = gTrayHwnd                   ' HWND receiving messages.
                Tray.hIcon = Icon.Handle                ' Tray icon.
                Tray.uFlags = NIF_ICON                  ' Set flags for valid data items
                Tray.cbSize = Len(Tray)                 ' Size of struct.
                
                rc = Shell_NotifyIcon(NIM_MODIFY, Tray) ' Send data to Sys Tray.
            End If
    
            Set gTrayIcon = Icon                        ' Save Icon to global
        End If
    End If
End Property

Public Property Get TrayIcon() As StdPicture
  Set TrayIcon = gTrayIcon                              ' Return Icon value
End Property

Public Property Let FormhWnd(hwnd As Long)
  gFormhWnd = hwnd
End Property

Public Property Get FormhWnd() As Long
  FormhWnd = gFormhWnd
End Property

Public Property Let TrayTip(Tip As String)
    Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure
    Dim rc As Long                                  ' API Return code

    If gAddedToTray Then                            ' if TrayIcon is in taskbar
        Tray.uID = gTrayId                          ' Unique ID for each HWND and callback message.
        Tray.hwnd = gTrayHwnd                       ' HWND receiving messages.
        Tray.szTip = Tip & vbNullChar               ' Tray tool tip
        Tray.uFlags = NIF_TIP                       ' Set flags for valid data items
        Tray.cbSize = Len(Tray)                     ' Size of struct.
        
        rc = Shell_NotifyIcon(NIM_MODIFY, Tray)     ' Send data to Sys Tray.
    End If
    
    gTrayTip = Tip                                  ' Save Tip
 End Property


Public Property Get TrayTip() As String
    TrayTip = gTrayTip                              ' Return Global Tip...
End Property


Public Property Let InTray(Show As Boolean)
    Dim ClassAddr As Long                           ' Address pointer to Control Instance

    If (Show <> gInTray) Then                       ' Modify ONLY if state is changing!
        If Show Then                                ' If adding Icon to system tray...
                PrevWndProc = SetWindowLong(gTrayHwnd, GWL_WNDPROC, AddressOf SubWndProc)
                ObjPointer = ObjPtr(Me)
                AddIcon gTrayHwnd, gTrayId, TrayTip, TrayIcon ' Add TrayIcon to System Tray...
                gAddedToTray = True                 ' Save state of control used in teardown procedure
        Else                                        ' If removing Icon from system tray
            If gAddedToTray Then                    ' If Added to system tray then remove...
                DeleteIcon gTrayHwnd, gTrayId       ' Remove icon from system tray
                
                ' Un SubClass controls window proc.
                SetWindowLong gTrayHwnd, GWL_WNDPROC, PrevWndProc
                gAddedToTray = False                ' Maintain the state for teardown purposes
            End If
        End If
        
        gInTray = Show                              ' Update global variable
    End If
End Property

Public Property Get InTray() As Boolean
    InTray = gInTray                                ' Return global property
End Property

Private Sub AddIcon(hwnd As Long, Id As Long, Tip As String, Icon As StdPicture)
    Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure
    Dim tFlags As Long                              ' Tray action flag
    Dim rc As Long                                  ' API return code

    Tray.uID = Id                                   ' Unique ID for each HWND and callback message.
    Tray.hwnd = hwnd                                ' HWND receiving messages.
    
    If Not (Icon Is Nothing) Then                   ' Validate Icon picture
        Tray.hIcon = Icon.Handle                    ' Tray icon.
        Tray.uFlags = Tray.uFlags Or NIF_ICON       ' Set ICON flag to validate data item
        Set gTrayIcon = Icon                        ' Save icon
    End If
    
    If (Tip <> "") Then                             ' Validate Tip text
        Tray.szTip = Tip & vbNullChar               ' Tray tool tip
        Tray.uFlags = Tray.uFlags Or NIF_TIP        ' Set TIP flag to validate data item
        gTrayTip = Tip                              ' Save tool tip
    End If
    
    Tray.uCallbackMessage = TRAY_CALLBACK           ' Set user defigned message
    Tray.uFlags = Tray.uFlags Or NIF_MESSAGE        ' Set flags for valid data item
    Tray.cbSize = Len(Tray)                         ' Size of struct.
    
    rc = Shell_NotifyIcon(NIM_ADD, Tray)            ' Send data to Sys Tray.
End Sub

Private Sub DeleteIcon(hwnd As Long, Id As Long)
    Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure

    Tray.uID = Id                                   ' Unique ID for each HWND and callback message.
    Tray.hwnd = hwnd                                ' HWND receiving messages.
    Tray.uFlags = 0&                                ' Set flags for valid data items
    Tray.cbSize = Len(Tray)                         ' Size of struct.
    
    Shell_NotifyIcon NIM_DELETE, Tray               ' Send delete message.
End Sub

Friend Sub SendEvent(MouseEvent As Long, Id As Long)
    Select Case MouseEvent                          ' Dispatch mouse events to control
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove(Id)
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(vbLeftButton, Id)
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton, Id)
    Case WM_LBUTTONDBLCLK
        RaiseEvent MouseDblClick(vbLeftButton, Id)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(vbRightButton, Id)
    Case WM_RBUTTONUP
        RaiseEvent MouseUp(vbRightButton, Id)
    Case WM_RBUTTONDBLCLK
        RaiseEvent MouseDblClick(vbRightButton, Id)
    End Select
End Sub

Friend Sub Clicked(index As Integer)                ' for the menu click from our WindowProc
  RaiseEvent MenuClick(index)
End Sub

Private Sub Class_Initialize()
    gInTray = defInTray                             ' Set global InTray defalt
    gAddedToTray = False                            ' Set default state
    gTrayId = 0                                     ' Set global TrayId default
    gTrayHwnd = Form1.hwnd                          ' Set and keep HWND of user control
    InTray = defInTray                              ' Init InTray Property
    TrayTip = defTrayTip                            ' Init TrayTip Property
    Set TrayIcon = gTrayIcon                        ' Init TrayIcon property
End Sub
 
Private Sub Class_Terminate()
  If InTray Then                                    ' If TrayIcon is visible
      InTray = False                                ' Cleanup and unplug it.
  End If
End Sub
