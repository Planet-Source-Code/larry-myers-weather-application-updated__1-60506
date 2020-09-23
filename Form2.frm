VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   870
   ClientLeft      =   11610
   ClientTop       =   7680
   ClientWidth     =   2085
   LinkTopic       =   "Form2"
   ScaleHeight     =   870
   ScaleWidth      =   2085
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   690
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   690
      ExtentX         =   1226
      ExtentY         =   1226
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
      Location        =   ""
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And Shift = 0 Then Unload Form1 'close on right click of form
If Button = 1 And Shift = 1 Then Form2.Hide 'minimize to tray if you left click with the shift key
If Button = 1 And Shift = 0 Then Form1.Show 'restore if you left click

End Sub





Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And Shift = 0 Then Unload Form1 'close on right click of label
If Button = 1 And Shift = 1 Then Form2.Hide 'minimize to tray if you left click with the shift key
If Button = 1 And Shift = 0 Then Form1.Show 'restore if you left click

End Sub

