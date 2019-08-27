VERSION 5.00
Begin VB.Form frmSetTimer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Timer Setup"
   ClientHeight    =   2460
   ClientLeft      =   1125
   ClientTop       =   1470
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   5235
   Begin VB.TextBox txtDelayTime 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Text            =   "1000"
      Top             =   1980
      Width           =   1095
   End
   Begin VB.CheckBox chkDelayRestart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delay restart after stop"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1980
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Loop Type"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   4335
      Begin VB.OptionButton optLoopType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Do Loop"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optLoopType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Timer Loop"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CheckBox chkTimerStopBG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enable StopBG in Timer Event"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CheckBox chkEnableTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enable Timer"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtInterval 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "1000"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton optTimerMode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Call function count times for one sample per call"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.OptionButton optTimerMode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Call function continuously with designated count"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   4815
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Timer Interval"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmSetTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkEnableTimer_Click()

   If chkEnableTimer.value = 0 Then optTimerMode(0).value = True

End Sub

Private Sub cmdOK_Click()

   Me.Hide

End Sub

Private Sub Form_Load()

   Me.Top = Screen.Height / 6
   Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub optLoopType_Click(Index As Integer)

   If optLoopType(0).value Then
      Label1.Caption = "Timer Interval"
      chkEnableTimer.Caption = "Enable Timer"
      chkTimerStopBG.Caption = "Enable StopBG in Timer Event"
   Else
      Label1.Caption = "Delay Factor"
      chkEnableTimer.Caption = "Enable Do Loop"
      chkTimerStopBG.Caption = "Enable StopBG in Do Loop"
   End If

End Sub

Private Sub txtInterval_Change()

   If Val(txtInterval.Text) > 65535 Then txtInterval.Text = "65535"

End Sub
