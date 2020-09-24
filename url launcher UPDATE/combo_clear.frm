VERSION 5.00
Begin VB.Form combo_clear 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All URL's Have been Cleared."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3375
   End
End
Attribute VB_Name = "combo_clear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub StayOnTop(Theform As Form)
Dim SetWinOnTop1
SetWinOnTop1 = SetWindowPos(Theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Load()
Me.Show
StayOnTop Me
Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Timer1_Timer()

Unload Me
End Sub
