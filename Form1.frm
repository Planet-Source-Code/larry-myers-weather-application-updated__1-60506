VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   375
   ClientTop       =   615
   ClientWidth     =   12060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12060
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6120
      Top             =   1440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Current Conditions"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WebBrowser6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "WebBrowser1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Timer2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "5 Day Forecast"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WebBrowser2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Radar"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "WebBrowser3"
      Tab(2).Control(1)=   "Command1"
      Tab(2).Control(2)=   "Command2"
      Tab(2).Control(3)=   "Command3"
      Tab(2).Control(4)=   "Command4"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Pollen Information"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "WebBrowser4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Almanac Information"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "WebBrowser5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Alerts"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "WebBrowser8"
      Tab(5).Control(1)=   "WebBrowser7"
      Tab(5).Control(2)=   "Timer3"
      Tab(5).Control(3)=   "WebBrowser9"
      Tab(5).ControlCount=   4
      Begin VB.CommandButton Command4 
         Caption         =   "Pacific Hurricane Loop"
         Height          =   315
         Left            =   -68280
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser9 
         Height          =   5475
         Left            =   -74880
         TabIndex        =   13
         Top             =   3000
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   9657
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   -69960
         Top             =   1560
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Atlantic Hurricane Loop"
         Height          =   315
         Left            =   -70680
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Regional Radar"
         Height          =   315
         Left            =   -72660
         TabIndex        =   8
         Top             =   360
         Width           =   1515
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Local Radar"
         Height          =   315
         Left            =   -74700
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser5 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   13679
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   7395
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   13044
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser4 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   13679
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   20745
         Top             =   2280
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   13679
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   7755
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   13679
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser6 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   4695
         ExtentX         =   8281
         ExtentY         =   5741
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser7 
         Height          =   2595
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   4577
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser8 
         Height          =   4275
         Left            =   -74880
         TabIndex        =   12
         Top             =   3840
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   7541
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.TextBox Text1 
      Height          =   6735
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Menu mnuset 
      Caption         =   "&Set Zip Code"
   End
   Begin VB.Menu mnutemp 
      Caption         =   "&Temperature Style"
      Begin VB.Menu mnuenglish 
         Caption         =   "English"
      End
      Begin VB.Menu mnumetric 
         Caption         =   "Metric"
      End
   End
   Begin VB.Menu mnuproxy 
      Caption         =   "&Proxy Settings"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuback 
      Caption         =   "&History (Back)"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "A&bout"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zpcd As String
Dim metric As Byte
Dim pth As String
Dim website As String
Private WithEvents IconEvent As SysTray
Attribute IconEvent.VB_VarHelpID = -1
Dim prxy As String
Dim cnt As Integer
Dim timercount As Integer
Dim lclrdr As String
Dim regrdr As String
Dim progress As Integer
Dim xmlcomplete As String
Dim timercount2 As Integer
Dim alertflag As String
Dim alertdone As String
Dim timercount3 As Integer
Dim alerterror As String
Dim founddetails As String




Dim bShown As Boolean, bHide As Boolean, hPopup&
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer

Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Sub Command1_Click()
WebBrowser3.Navigate (pth & "localrdr.htm")

End Sub

Private Sub Command2_Click()
WebBrowser3.Navigate (pth & "regionalrdr.htm")

End Sub

Private Sub Command3_Click()
WebBrowser3.Navigate ("http://sirocco.accuweather.com/sat_mosaic_640x480_public/IR/isahatl.gif")


End Sub

Private Sub Command4_Click()
WebBrowser3.Navigate ("http://sirocco.accuweather.com/sat_mosaic_640x480_public/IR/isaepac.gif")
End Sub

Private Sub Form_Load()
'The menu and systray stuff I found a while back and here is where the credit belongs
' Systray Example by Wpsjr1@syix.com - Paul aka OnErr0r @ #visualbasic EFNet

' This example uses a WindowProc to subclass the API menu created and _
' messages from the SysTray.  Make sure you do NOT exit by pressind End.
' Always exit by clicking the X on the top right of the form.  This allows _
' the the Class Terminate event to fire and which unhooks the form.

' make an API menu.  (I think this is really easier than a VB menu.)
  On Error GoTo errhandler

  
  hPopup = CreatePopupMenu()
  AppendMenu hPopup, MF_STRING, WM_MENUBASE, "&Show Weather"
  AppendMenu hPopup, MF_STRING, WM_MENUBASE + 1, "&Minimize to Tray"
  AppendMenu hPopup, MF_STRING, WM_MENUBASE + 2, "&Restore from Tray"
  AppendMenu hPopup, MF_SEPARATOR, WM_MENUBASE, ""
  AppendMenu hPopup, MF_STRING, WM_MENUBASE + 3, "&Exit"
  SetMenuDefaultItem hPopup, 0&, True    ' make the first menu item bold

  bShown = True
  Set IconEvent = New SysTray
  Set IconEvent.TrayIcon = Me.Icon
  IconEvent.FormhWnd = Me.hWnd
  IconEvent.TrayTip = "double click to show form" & vbNullChar
  IconEvent.InTray = True

SSTab1.Enabled = False 'make tab unselectable to make time for the webpage to be loaded and deciphered

Form2.Left = Screen.Width - Form2.Width - 40
Form2.Top = Form2.Height - 400

If Len(App.Path) > 3 Then pth = App.Path & "\" Else pth = App.Path
fnd = Dir(pth & "settings.ini") 'Check to see if there is a settings.ini file
If fnd = "" Then 'OOOPS no settings.ini file, make a new one with defaults
1       zpcd = InputBox("Enter Preferred Zip Code", "Enter Zip Code") 'get the zipcode from the user
    If Len(zpcd) = 0 Then GoTo 1
    If Len(zpcd) > 5 Then GoTo 1
    zp = "zip=" & zpcd
    mt = "metric=0" 'default=english
    prxytemp = "proxy=" 'default is direct connect - no proxy
    Open pth & "settings.ini" For Output As #1 'create settings.ini
        Print #1, zp
        Print #1, mt
        Print #1, prxytmp
    Close #1
End If


Open pth & "settings.ini" For Input As #1 'we are sure there is a file, now grab settings
    Input #1, settings
    Input #1, mt
    Input #1, prxytmp
Close #1

zpcd = Right$(settings, 5)

If Len(settings) < 9 Or Len(settings) > 9 Or zpcd <> Val(zpcd) Then GoTo 1 'zipcode is incorrect
If Len(mt) < 8 Or Len(mt) > 8 Then metric = 0 Else metric = Val(Right$(mt, 1)) 'verify and correct inconsistencies
If Len(prxytmp) < 7 Then prxy = "" Else prxy = Right$(prxytmp, Len(prxytmp) - 6) 'verify and correct inconsistencies

If Len(mt) < 8 Then
    zp = settings
    mt = "metric=0"
    Open pth & "settings.ini" For Output As #1
        Print #1, zp
        Print #1, mt
        Print #1, prxytmp
    Close #1
End If

If metric = 1 Then mnumetric.Checked = True: mnuenglish.Checked = False 'metric is chosen
If metric = 0 Then mnumetric.Checked = False: mnuenglish.Checked = True 'english is chosen

Timer1.Enabled = False 'timer for updating weather and current conditions

On Error GoTo errhandler
progress = 1
forecast 'this is the function that takes care of the current conditions and the forecast

  check_alerts

website = "0"

WebBrowser4.Navigate ("http://pollen.com/forecast_fourday.asp?AffiliateID=2508&zip=" & zpcd & "&ft=1&pop=1")

Timer1.Enabled = True
SSTab1.Enabled = True
progress = 0
Exit Sub
errhandler:

MsgBox "There has been an unrecoverable error....Closing"
Exit Sub

End Sub





Private Sub IconEvent_MouseUp(Button As Integer, Id As Long)
  
If Button = 1 Then Me.Show 'left click on systray icon gives you the weather screen

End Sub

Private Sub mnuabout_Click()

MsgBox "Weather Program made by Larry Myers 443-543-3107"

End Sub

Private Sub mnuAlerts_Click()

End Sub

Private Sub mnuback_Click()
If SSTab1 = 0 Then WebBrowser1.Navigate (pth & "currentweather.htm")
If SSTab1 = 1 Then WebBrowser2.Navigate (pth & "fivedayforecast.htm")
If SSTab1 = 2 Then WebBrowser3.Navigate (pth & "localrdr.htm")
If SSTab1 = 3 Then WebBrowser4.Navigate ("http://pollen.com/forecast_fourday.asp?AffiliateID=2508&zip=" & zpcd & "&ft=1&pop=1")
If SSTab1 = 4 Then WebBrowser5.Navigate (pth & "almanac.htm")

End Sub

Private Sub mnuenglish_Click()
'this is where english settings are selected

mnuenglish.Checked = True
mnumetric.Checked = False
'create settings.ini
zp = "zip=" & zpcd
mt = "metric=0"
prxytmp = "proxy=" & prxy
    Open pth & "settings.ini" For Output As #1
        Print #1, zp
        Print #1, mt
        Print #1, prxytmp
    Close #1

metric = 0 'set metric variable to match english

tb = SSTab1.Tab 'make decision as to what function to call based on the tab selected
If tb = 0 Then forecast


End Sub

Private Sub mnuhelp_Click()
'quick and dirty help

MsgBox "When main weather form is on screen the 4 tabs give different information, current conditions is updated every minute. When minimized a mini form will appear: Left click on it=show main form, Left and Shift will minimize to task bar, Ricght click kills it. The Systray icon also has some abilities."

End Sub

Private Sub mnumetric_Click()
'this is where english settings are selected

mnuenglish.Checked = False
mnumetric.Checked = True

'create settings.ini
zp = "zip=" & zpcd
mt = "metric=1"
prxytmp = "proxy=" & prxy
Open pth & "settings.ini" For Output As #1
Print #1, zp
Print #1, mt
Print #1, prxytmp
Close #1

metric = 1 'set metric variable to metric

tb = SSTab1.Tab 'make decision as to what function to call based on the tab selected
If tb = 0 Then forecast


End Sub

Private Sub mnuproxy_Click()
'this is where proxy settings are made from the menu

prxytmp = InputBox("Enter Proxy Settings - None for No Proxy, ? for Auto Setup, Type in the IP for Manual Proxy", "Proxy Settings", prxy, Me.Left, Me.Top)
If Len(prxytmp) < 1 Then prxy = "" Else prxy = prxytmp 'if proxy is set to ? set to auto, if blank set to direct connect, if it is something else set the proxy to manual with proxy as that setting

'write settings.ini file again
zp = "zip=" & zpcd
mt = "metric=" & metric
prxytmp = "proxy=" & prxy
    Open pth & "settings.ini" For Output As #1
        Print #1, zp
        Print #1, mt
        Print #1, prxytmp
    Close #1

10 tb = SSTab1.Tab 'make decision as to what function to call based on the tab selected
 forecast
check_alerts
'also need to update the pollen tab
WebBrowser4.Navigate ("http://pollen.com/forecast_fourday.asp?AffiliateID=2508&zip=" & zpcd & "&ft=1&pop=1")

End Sub

Private Sub mnuset_Click()
'this is where I set the zip code based on user input

1 zpcdtemp = InputBox("Enter Preferred Zip Code", "Enter Zip Code", zpcd, Me.Left, Me.Top)

If Len(zpcdtemp) = 0 Then GoTo 10 'if blank that means cancel the change
If Len(zpcdtemp) > 5 Or Len(zpcdtemp) < 5 Or zpcdtemp <> Val(zpcdtemp) Then GoTo 1 'bad zip code

'create settings.ini file
zp = "zip=" & zpcdtemp
zpcd = zpcdtemp
mt = "metric=" & metric
prxytmp = "proxy=" & prxy
    Open pth & "settings.ini" For Output As #1
        Print #1, zp
        Print #1, mt
        Print #1, prxytmp
    Close #1

10 tb = SSTab1.Tab 'make decision as to what function to call based on the tab selected
forecast
check_alerts

'need to update pollen count for zip code
WebBrowser4.Navigate ("http://pollen.com/forecast_fourday.asp?AffiliateID=2508&zip=" & zpcd & "&ft=1&pop=1")

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 0 Then WebBrowser1.Navigate (pth & "currentweather.htm")
If SSTab1.Tab = 1 Then WebBrowser2.Navigate (pth & "fivedayforecast.htm")
If SSTab1.Tab = 2 Then WebBrowser3.Navigate (pth & "localrdr.htm")
If SSTab1.Tab = 3 Then WebBrowser4.Navigate ("http://pollen.com/forecast_fourday.asp?AffiliateID=2508&zip=" & zpcd & "&ft=1&pop=1")
If SSTab1.Tab = 4 Then WebBrowser5.Navigate (pth & "almanac.htm")
If SSTab1.Tab = 5 Then WebBrowser7.Navigate (pth & "alerts.htm")

End Sub

Private Sub Timer1_Timer()
'refresh of current data and forecast data is set to 10 minute approx.
timercount = timercount + 1
If timercount > 9 Then
timercount = 0
forecast
End If
End Sub





Public Sub forecast()
'this is where we get the current conditions and the forecast

Open pth & "loading.htm" For Output As #1 'put it locally
        Print #1, "<html><body><h1>Loading please wait</h1></body></html>"
    Close #1
If WebBrowser8.LocationURL = "" Then WebBrowser8.Navigate2 (pth & "loading.htm")
If WebBrowser1.LocationURL = "" Then WebBrowser1.Navigate2 (pth & "loading.htm")
If WebBrowser2.LocationURL = "" Then WebBrowser2.Navigate2 (pth & "loading.htm")
If WebBrowser3.LocationURL = "" Then WebBrowser3.Navigate2 (pth & "loading.htm")
If WebBrowser5.LocationURL = "" Then WebBrowser5.Navigate2 (pth & "loading.htm")
If cnt < 1 Then cnt = 1 'trouble connecting counter

On Error GoTo errhandler

website = "1"
Timer1.Enabled = False 'do not allow overlap of timers
Timer2.Enabled = False
SSTab1.Enabled = False 'do not allow another click on tab

'this is the site I want my data from, they seem to be more accurate that some others

webtxt = OpenURL("www.wunderground.com/cgi-bin/findweather/getForecast?query=" & zpcd)
'lotsand lots of string manipulations here to grab MY data and pictures
'I grab all the data and then I remove all the ads and junk to just give me what I want
'then I store it locally

'need to strip scripts
webtxt = Replace(webtxt, "<script", "<!--<script")
webtxt = Replace(webtxt, "</script>", "</script>-->")

webtxt = Replace(webtxt, "Tropical Weather:", "<center>Tropical Weather:")
strt = InStr(1, webtxt, "<title>")
'find rss feed info (this makes finding current conditions MUCH easier
strtxml = InStr(strt, webtxt, ".xml")
ndxml = strtxml + 4
strtxml = InStrRev(webtxt, "href", strtxml) + 6
xml = Mid$(webtxt, strtxml, ndxml - strtxml)
xmlcomplete = ""
WebBrowser6.Navigate (xml)
While xmlcomplete <> "done"
DoEvents
Wend

Set ehtml = WebBrowser6.Document.All.Item(i)
xmltxt2 = ehtml.innertext
xmltxt = xmltxt2
'have xml, need to parse it
strtxmlparse = InStr(1, xmltxt, "description") + 12
ndxmlparse = InStr(1, xmltxt, "</description")
xmlcity = Mid$(xmltxt, strtxmlparse + 33, ndxmlparse - 33 - strtxmlparse)

strtxmlparse = InStr(strtxmlparse, xmltxt, "Current Conditions") + 21
ndxmlparse = InStr(strtxmlparse, xmltxt, "</")
xmlcurrenttime = Mid$(xmltxt, strtxmlparse, ndxmlparse - strtxmlparse)

strtxmlparse = InStr(strtxmlparse, xmltxt, "<link>") + 6
ndxmlparse = InStr(strtxmlparse, xmltxt, "</")
xmllink = Mid$(xmltxt, strtxmlparse, ndxmlparse - strtxmlparse)

strtxmlparse = InStr(strtxmlparse, xmltxt, "<description>") + 13
ndxmlparse = InStr(strtxmlparse, xmltxt, "</")
xmlcurrentcond = Mid$(xmltxt, strtxmlparse, ndxmlparse - strtxmlparse)

xmlcurrentsplit = Split(xmlcurrentcond, " | ")
xmltemp = Replace(xmlcurrentsplit(0), "&#176;", Chr$(176))
xmltemp = Right$(xmltemp, Len(xmltemp) - 27)
xmlhumidity = Right$(xmlcurrentsplit(1), Len(xmlcurrentsplit(1)) - 10)
xmlpressure = Right$(xmlcurrentsplit(2), Len(xmlcurrentsplit(2)) - 10)
xmlcondition = Right$(xmlcurrentsplit(3), Len(xmlcurrentsplit(3)) - 12)
xmlwinddir = Right$(xmlcurrentsplit(4), Len(xmlcurrentsplit(4)) - 16)
xmlwindspd = Left(xmlcurrentsplit(5), Len(xmlcurrentsplit(5)) - 3)
xmlwindspd = Right$(xmlwindspd, Len(xmlwindspd) - 12)
windspdend = InStr(1, xmlwindspd, "/h") + 1
xmlwindspd = Left$(xmlwindspd, windspdend)

'find eng and metric
xmltempsplit = Split(xmltemp, " / ")
temp = xmltempsplit(metric)
xmlwindspdsplit = Split(xmlwindspd, " / ")
windspd = xmlwindspdsplit(metric)
weathstrt = InStr(strt, webtxt, "<!-- Time Bar")
weathend = InStr(weathstrt, webtxt, "METAR")
If weathend = 0 Then weathend = InStr(weathstrt, webtxt, "Never show Personal Weather Stations here")
weathend = InStrRev(webtxt, "<tr", weathend)



weather = "<html><body><table><tr><td bgcolor='000099' style='width: 100%; color: rgb(255, 255, 255); font-weight: bold; font-size: 13px;'><center><b>" & xmlcity & "</b></td></tr><tr><td bgcolor='000099' style='width: 100%; color: rgb(255, 255, 255); font-weight: bold; font-size: 13px;'>" & Replace(Mid$(webtxt, weathstrt, weathend - weathstrt), "href=" & Chr$(34) & "/weatherstation", "href=" & Chr$(34) & "http://www.wunderground.com/weatherstation")
'weather = Replace(weather, "<a href=", "<a target='_blank' href=")
weather = Replace(weather, "a href=" & Chr$(34) & "/", "a target='_blank' href=" & Chr$(34) & "http://www.wunderground.com/")
weather = Replace(weather, "<h3>Current Conditions</h3>", "<center><font color=white>Current Conditions</font></center></td></style></table></table>")

'remove rapidfire
fndrapidfirestrt = InStrRev(weather, "<div", InStr(1, weather, "&raquo;")) - 1
fndrapidfireend = InStr(fndrapidfirestrt, weather, "</div>") + 7
weather = Left$(weather, fndrapidfirestrt) & Right$(weather, Len(weather) - fndrapidfireend)
fndtdstrt = InStrRev(weather, "class", InStr(1, weather, "<!--<script")) + 7
weather = Left$(weather, fndtdstrt) & "LM" & Right$(weather, Len(weather) - fndtdstrt)
weather = Replace(weather, "class=" & Chr$(34) & "vLMaT" & Chr$(34) & " style=" & Chr$(34) & "width: 375px;", "class=" & Chr$(34) & "LMVaT" & Chr$(34) & " bgcolor='000099' style='width: 100%; color: rgb(255, 255, 255); font-weight: bold; font-size: 13px;'")
fndtblstrt = InStrRev(weather, "<tr", fndtdstrt) - 1
weather = Left$(weather, fndtblstrt) & "</table><table>" & Right$(weather, Len(weather) - fndtblstrt)
fndlatlon = InStrRev(weather, "<td", InStr(1, weather, "Lat/Lon")) - 1
weather = Left$(weather, fndlatlon) & "</tr><tr>" & Right$(weather, Len(weather) - fndlatlon)
weather = Replace(weather, "100%", "700px")
weather = Replace(weather, "<td class=""taR"" style=""white-space: nowrap;"">", "<td colspan=2 class=""taR"" align='center' style=""white-space: nowrap;color: white"">")
'<td style="white-space: nowrap;">
'<td id="full" style="white-space: nowrap;">
weather = Replace(weather, "<td id=""full"" style=""white-space: nowrap;"">", "<td id=""full""  align='center' style=""white-space: nowrap; color: white"">")
weather = Replace(weather, "<td style=""white-space: nowrap;"">", "<td align='center' style=""white-space: nowrap; color: white"">")
weather = Replace(weather, "width: 375px", "width: 700px")
fndalert = InStrRev(weather, "table", InStrRev(weather, "td", InStr(1, weather, "blue_Warning.gif"))) + 4
weather = Left$(weather, fndalert) & " bgcolor=ff0000" & Right$(weather, Len(weather) - fndalert)

'need to grab 5 day weather
fivedaystrt = InStrRev(webtxt, "<table", InStr(weathend, webtxt, "5-Day Forecast"))
fivedayend = InStrRev(webtxt, "<tr>", InStr(fivedaystrt, webtxt, "p5") - 1) - 1
fivedayweb = Replace(Mid$(webtxt, fivedaystrt, fivedayend - fivedaystrt), "target='_blank' href=" & Chr$(34) & "/cgi-bin", " href=" & Chr$(34) & "http://www.wunderground.com/cgi-bin")
fivedayweb = Replace(fivedayweb, "href=/MOS", "target='_blank' href=http://www.wunderground.com/MOS")
fivedayweb = Replace(fivedayweb, "href=" & Chr$(34) & "/Display", "target='_blank' href=" & Chr$(34) & "http://www.wunderground.com/Display")
fivedayweb = Replace(fivedayweb, "<a href=" & Chr$(34) & "/cgi-bin", "<a target='blank' href=" & Chr$(34) & "http://www.wunderground.com/cgi-bin")

'need detailed five day data too
detailfivedaystrt = InStr(fivedayend, webtxt, "Forecast for ")
detailfivedayend = InStr(InStr(detailfivedaystrt, webtxt, "<h5>-") + 1, webtxt, "</table") + 9
detailfivedayweb = "<html><body><table>" & Mid$(webtxt, detailfivedaystrt, detailfivedayend - detailfivedaystrt) & "</body></html>"
detailfivedayweb = Replace(detailfivedayweb, Chr$(34) & "/cgi", Chr$(34) & "http://www.wunderground.com/cgi")
detailfivedayweb = Replace(detailfivedayweb, Chr$(34) & "/sev", Chr$(34) & "http://www.wunderground.com/sev")
detailfivedayweb = Replace(detailfivedayweb, "/Dis", "http://www.wunderground.com/Dis")
fndshowhidestart = InStrRev(detailfivedayweb, "<td", InStr(1, detailfivedayweb, ">Show")) - 1
While fndshowhidestart > 0
fndshowhideend = InStr(fndshowhidestart, detailfivedayweb, "</td>") + 6
detailfivedayweb = Left$(detailfivedayweb, fndshowhidestart) & Right$(detailfivedayweb, Len(detailfivedayweb) - fndshowhideend)
If InStr(1, detailfivedayweb, ">Show") > 0 Then fndshowhidestart = InStrRev(detailfivedayweb, "<td", InStr(1, detailfivedayweb, ">Show")) - 1 Else fndshowhidestart = 0
Wend
detailfivedayweb = Replace(detailfivedayweb, " href", " target='_blank' href")

'need to grab almanac
almanacstrt = InStr(weathstrt, webtxt, "<a name=" & Chr$(34) & "History" & Chr$(34) & "></a>")
almanacend = InStrRev(webtxt, "</table>", InStr(almanacstrt, webtxt, ">Definitions of")) + 9
almanacweb = Mid$(webtxt, almanacstrt, almanacend - almanacstrt)
almanacweb = Replace(almanacweb, "<a href", "<a target='_blank' href")
almanacweb = Replace(almanacweb, "href=/", "href=http://www.wunderground.com/")
almanacweb = Replace(almanacweb, "href=" & Chr$(34) & "/", "href=""http://www.wunderground.com/")
almanacweb = Replace(almanacweb, "action=/cgi-bin", " target='_blank' action=http://www.wunderground.com/cgi-bin")
fndshowhidestart = InStrRev(almanacweb, "<td", InStr(1, almanacweb, ">Show")) - 1
While fndshowhidestart > 0
fndshowhideend = InStr(fndshowhidestart, almanacweb, "</td>") + 6
almanacweb = Left$(almanacweb, fndshowhidestart) & Right$(almanacweb, Len(almanacweb) - fndshowhideend)
If InStr(1, almanacweb, ">Show") > 0 Then fndshowhidestart = InStrRev(almanacweb, "<td", InStr(1, almanacweb, ">Show")) - 1 Else fndshowhidestart = 0
Wend
Open pth & "almanac.htm" For Output As #1 'put it locally
        Print #1, almanacweb
    Close #1
WebBrowser5.Navigate (pth & "almanac.htm")

'now to find the local radar
radarend = InStr(weathend, webtxt, "Local Radar</a>") - 2
radarstrt = InStrRev(webtxt, "href=", radarend) + 7
localradar = "www.wunderground.com/" & Mid$(webtxt, radarstrt, radarend - radarstrt)
localradar = Left$(localradar, InStr(1, localradar, Chr$(34)) - 1)
rdrtxt = OpenURL(localradar)
rdrimg = InStr(1, rdrtxt, "<img name=" & Chr$(34) & "map") + 21
rdrimgend = InStr(rdrimg, rdrtxt, Chr$(34))
rdrimg = Mid$(rdrtxt, rdrimg, rdrimgend - rdrimg)
rdrimg = Replace(rdrimg, "&num=0&", "&num=12&")
rdrimg = Replace(rdrimg, "&num=1&", "&num=12&")
'animated image is found get new links and add to page
rdrlnksstrt = InStr(1, rdrtxt, "<h4>Base Reflectivity</h4>")
rdrlnksend = InStr(rdrlnksstrt, rdrtxt, "</table>")
rdrlnks = Mid$(rdrtxt, rdrlnksstrt, rdrlnksend - rdrlnksstrt)
rdrlnksformatstrt = InStr(1, rdrlnks, "<td") - 1
While rdrlnksformatstrt > 0
rdrlnksformatend = InStr(rdrlnksformatstrt, rdrlnks, ">")
rdrlnks = Left$(rdrlnks, rdrlnksformatstrt) & Right$(rdrlnks, Len(rdrlnks) - rdrlnksformatend)
rdrlnksformatstrt = InStr(1, rdrlnks, "<td") - 1
Wend
rdrlnksformatstrt = InStr(1, rdrlnks, "<tr") - 1
While rdrlnksformatstrt > 0
rdrlnksformatend = InStr(rdrlnksformatstrt, rdrlnks, ">")
rdrlnks = Left$(rdrlnks, rdrlnksformatstrt) & Right$(rdrlnks, Len(rdrlnks) - rdrlnksformatend)
rdrlnksformatstrt = InStr(1, rdrlnks, "<tr") - 1
Wend
rdrlnks = Replace(rdrlnks, "</tr>", "<br>")
rdrlnks = Replace(rdrlnks, "bold", "normal")
rdrlnks = Replace(rdrlnks, "</td>", "")
rdrlnks = Replace(rdrlnks, "<a href=" & Chr$(34) & "/", "<a target='_blank' href=" & Chr$(34) & "http://www.wunderground.com/")
Open pth & "localrdr.htm" For Output As #1 'put it locally
        Print #1, "<html><body><table><tr><td valign=top><img src='" & rdrimg & "'></td><td>" & rdrlnks & "</td></tr></table></body></html>"
    Close #1




'now find regional radar
radarend = InStr(radarend, webtxt, "Regional Radar</a>") - 2
radarstrt = InStrRev(webtxt, "href=", radarend) + 7
regionalradar = "www.wunderground.com/" & Mid$(webtxt, radarstrt, radarend - radarstrt)
regionalradar = Left$(regionalradar, InStr(1, regionalradar, Chr$(34)) - 1)
rdrtxt = OpenURL(regionalradar)
rdrimgstrt = InStr(InStr(1, rdrtxt, ">Nexrad Mixed Composite Radar Map</h1>"), rdrtxt, "<table")
rdrimgend = InStr(InStr(rdrimgstrt, rdrtxt, "</table>") + 1, rdrtxt, "</table>") + 13
rdrimghtml = "<html><body>" & Mid$(rdrtxt, rdrimgstrt, rdrimgend - rdrimgstrt) & "</body></html>"
rdrimghtml = Replace(Replace(rdrimghtml, Chr$(34) & "/radar", Chr$(34) & "http://www.wunderground.com/radar"), Chr$(34) & "/cgi", Chr$(34) & "http://www.wunderground.com/cgi")
rdrimghtml = Replace(rdrimghtml, "<a href=" & Chr$(34), "<a target='_blank' href=" & Chr$(34) & "http://www.wunderground.com/radar/")
Open pth & "regionalrdr.htm" For Output As #1 'put it locally
        Print #1, rdrimghtml
    Close #1


'weather=current conditions & 5 days summary
weather = weather & "</table></table>" & fivedayweb & "</table></body></html>"

'weatherfiveday=fiveday weather forecast
weatherfiveday = "<html><body>" & fivedayweb & "</table></body></html>"

iconstrt = InStr(InStr(strt, webtxt, "Current Conditions"), webtxt, "<img") + 10
iconend = InStr(iconstrt, webtxt, Chr$(34))
mg = Mid$(webtxt, iconstrt, iconend - iconstrt) 'mg = current weather image
frmcp = xmlcity & " @ " & xmlcurrenttime & "_" & temp & "_" & xmlcondition & "_" & windspd & " from the " & xmlwinddir  'create form caption
Form1.Caption = frmcp

IconEvent.TrayTip = frmcp & vbNullChar 'tooltip for tray icon

'these few lines set up the minimum weather page just temperature and current weather image
Form2.Label1 = temp 'when minimized this will become visible is just for temperature
wbtx = "<html><body background=" & Chr$(34) & mg & Chr$(34) & " scroll=no></body></html>" 'need another webpage for image

    Open pth & "test2.htm" For Output As #1 'put it locally
        Print #1, wbtx
    Close #1

tp = Form2.Shape1.Height / 2
Form2.WebBrowser1.Top = tp - Form2.WebBrowser1.Height / 2 'needed to center webbrowser on form programatically
Form2.WebBrowser1.Navigate pth & "test2.htm" 'go to local site
Form2.Label1.ToolTipText = frmcp 'put current weather in tooltips for this form
Form2.WebBrowser1.ToolTipText = frmcp
' need webpage creation here!!

Open pth & "currentweather.htm" For Output As #1 'store webpage locally
        Print #1, weather
Close #1
Open pth & "fivedayforecast.htm" For Output As #1 'store webpage locally
        Print #1, detailfivedayweb
Close #1

SSTab1.Enabled = True 'ok, you can press another tab
mnuback_Click


website = "0"

cnt = 0 'good connect
Timer1.Enabled = True
Timer2.Enabled = True
Exit Sub

errhandler:

cnt = cnt + 1
Debug.Print cnt
If progress = 1 Then Exit Sub
If cnt > 2 Then MsgBox "Not Connecting: Try Setting Your Proxy Or may be a change on the website": Exit Sub 'try connecting twice if no then tell them what may help
'2 was picked for dialup, takes quite a few seconds at dialup, LAN speeds are way different
SSTab1.Enabled = True
forecast 'try again

End Sub

Private Sub Timer2_Timer()
timercount2 = timercount2 + 1
If timercount2 > 4 Then
'use this as an alerts checker
timercount2 = 0
check_alerts

End If
End Sub

Private Sub Timer3_Timer()
timercount3 = timercount3 + 1
If timercount3 > 1 Then
timercount3 = 0
alerterror = "yes"
End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

If website = "0" And URL <> pth & "currentweather.htm" Then Cancel = True 'this stops all links that stay in this window from working
'easier than trying to remove the links themselves  all external window hops work

End Sub
Private Sub Form_Resize()
bHide = True
If bHide = True Then
  
  If WindowState = 1 And bShown = True Then
    
    Form2.Show 'bring out the mini screen
    success = SetWindowPos(Form2.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) 'set it to topmost
    
    Me.Hide 'hide big form
    bShown = False
    
  Else
    
    Form2.Hide 'hide mini screen
    
    Me.WindowState = 0 'show main screen
    bShown = True
  
  End If

End If

End Sub

Private Sub IconEvent_MenuClick(index As Integer)
'menu clicks for the popup meny on systray is handled here
  
  Select Case index
    Case 0
      Me.Show 'show main form
      Me.WindowState = 0       ' needed for when the form is in the taskbar.
    Case 1
      Form2.Hide 'minimize mini form to tray
    Case 2
    Form2.Show 'restore mini form from tray
    Case 3
      Unload Me 'exit
  End Select
  
End Sub

Private Sub IconEvent_MouseDblClick(Button As Integer, Id As Long)
  'double click and I will show main form
  
  If Button = 1 Then Me.Show

End Sub

Public Sub IconEvent_MouseDown(Button As Integer, Id As Long)
 If Button = 2 Then
    Dim pt As POINTAPI, X&
    
    GetCursorPos pt
    ' the next three lines are the trick to getting the Popup to disappear properly
    SetForegroundWindow (Me.hWnd)
    TrackPopupMenu hPopup, TPM_CENTERALIGN Or TPM_RIGHTBUTTON, pt.X, pt.Y, 0&, Me.hWnd, 0&
    PostMessage Me.hWnd, 0&, 0&, 0&
  End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
'kill everything that is necessary

Set IconEvent = Nothing
DestroyMenu hPopup
Unload Form2
Unload Me
  
End Sub

Private Function OpenURL(ByVal sUrl As String) As String
' this is to grab the html of the webpages

'****************************************************
'PURPOSE:       Returns Contents (including all HTML) from
'               a web page
'PARAMETER:     sURL (e.g., http://www.freevbcode.com)
'RETURN VALUE:  Contents of requested page, or
'               empty string if sURL is not available
'COMMENTS:  This is an alternative to using the Internet Transfer
'           Control 's OpenURL method.  That control has a bug
'           Whereby not all the contents of the page will be
'           returned in certain circumstances
'*****************************************************

    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String
If prxy = "" Then hproxyuse = 1: hproxy = ""
If prxy = "?" Then hproxtuse = 0: hproxy = ""
If prxy <> "" And prxy <> "?" Then hproxyuse = 3: hproxy = prxy
hOpen = InternetOpen(scUserAgent, hproxyuse, _
    hproxy, vbNullString, 0)

hOpenUrl = InternetOpenUrl(hOpen, "http://" & sUrl, vbNullString, 0, _
   INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, _
           Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, _
             lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
      
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer

End Function



Private Sub WebBrowser6_DocumentComplete(ByVal pDisp As Object, URL As Variant)
xmlcomplete = "done"
End Sub
Public Sub check_alerts()
'need to verify alerts are there or not
'if alerts then check text to see if new or not
Timer1.Enabled = False 'do not allow overlap of timers
Timer2.Enabled = False
SSTab1.Enabled = True
alertdone = ""
WebBrowser8.Navigate ("http://www.myforecast.com/bin/alert_summary.m?zip_code=" & zpcd & "&metric=false") 'check for alerts

While alertdone <> "done"
DoEvents
Wend
alertflag = ""
Set ehtml2 = WebBrowser8.Document.documentElement

xmltxtalert = ehtml2.innerhtml
alertsfulltxt = xmltxtalert

'should have full text, need to find just the alerts
fndalertstrt = InStr(1, alertsfulltxt, "barhead")
fndalertallstrt = InStrRev(alertsfulltxt, "<TABLE", InStr(fndalertstrt + 1, alertsfulltxt, "barhead"))
If fndalertstrt < 1 Then alertflag = "alert error"
If alertflag <> "alert error" Then
fndalertend = InStr(InStr(fndalertstrt + 1, alertsfulltxt, "</TABLE>") + 1, alertsfulltxt, "</TABLE>") + 8
If fndalertend < 1 Then
'missing some data erroring out
    Open pth & "alerts.htm" For Output As #1 'put it locally
        Print #1, "<html><body><h1>Alerts are erroring out!</h1><br>Maybe website has changed!<br><a target='_blank' href=" & alertsurl & ">" & alertsurl & "</a></body></html>"
    Close #1
Else
'good data came in from alert site
'strip out alerts for the area
fndalertallend = InStr(InStr(fndalertallstrt, alertsfulltxt, "</TABLE>") + 1, alertsfulltxt, "</TABLE>") + 8
fndalertstrt = InStrRev(alertsfulltxt, "<TABLE", fndalertstrt)
alertsareatxt = Mid$(alertsfulltxt, fndalertstrt, fndalertend - fndalertstrt)
alertsalltxt = Mid$(alertsfulltxt, fndalertallstrt, fndalertallend - fndalertallstrt)

  If InStr(1, LCase(alertsareatxt), "<b>no weather alerts</b>") Then
alertflag = "none"
Else
alertflag = "some"

End If
'have all alerts for area and nation, fix links and images
alertstxt = alertsareatxt & alertsalltxt
alertstxt = Replace(alertstxt, "background=/", "background=http://www.myforecast.com/")
alertstxt = Replace(alertstxt, "<A href=" & Chr$(34), "<A  href=" & Chr$(34) & "http://www.myforecast.com")
Open pth & "alerts.htm" For Output As #1 'put it locally
        Print #1, "<html><body>" & alertstxt & "</body></html>"
    Close #1
    
WebBrowser7.Navigate (pth & "alerts.htm")


End If
Else
    Open pth & "alerts.htm" For Output As #1 'put it locally
        Print #1, "<html><body><h1>Alerts are erroring out!</h1><br>Maybe website has changed!<br><a target='_blank' href=" & alertsurl & ">" & alertsurl & "</a></body></html>"
    Close #1

End If

If alertflag = "some" Then Form2.BackColor = vbRed Else Form2.BackColor = 32768
If alertflag = "some" Then Form2.Label1.BackColor = vbRed Else Form2.Label1.BackColor = 32768

Timer1.Enabled = True 'do not allow overlap of timers
Timer2.Enabled = True
SSTab1.Enabled = True 'do not allow another click on tab

End Sub

Private Sub WebBrowser7_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If URL = "" Then Exit Sub

If URL = "http:///" Then Exit Sub

If InStr(1, URL, "error") > 0 Then Exit Sub

If InStr(1, URL, pth) < 1 And InStr(1, URL, "about:blank") < 1 Then
founddetails = ""
WebBrowser9.Navigate (URL)

URL = ""

WebBrowser7.Navigate (pth & "alerts.htm")

End If

End Sub

Private Sub WebBrowser8_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If InStr(1, URL, pth) > 0 Then Exit Sub
If (pDisp Is WebBrowser8.object) Then

If Mid$(URL, 8, 18) = "www.myforecast.com" Then alertdone = "done"
End If
End Sub


