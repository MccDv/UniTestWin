VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3360
   ClientLeft      =   1110
   ClientTop       =   1500
   ClientWidth     =   5805
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   5805
   Begin VB.Label lblPackage 
      BackStyle       =   0  'Transparent
      Caption         =   "DAQFlex Library"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   2940
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblPackage 
      BackStyle       =   0  'Transparent
      Caption         =   "DAQFlex Library"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   2940
      TabIndex        =   3
      Top             =   1620
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblSupports 
      BackStyle       =   0  'Transparent
      Caption         =   "Includes support for"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblPackage 
      BackStyle       =   0  'Transparent
      Caption         =   "DAQFlex Library"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   2940
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C06000&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   2100
      Width           =   1995
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()

   Me.Hide
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

   Me.Hide
   
End Sub

Private Sub Form_Load()

#If MSGOPS Then
   Me.lblSupports.Visible = True
   Me.lblPackage(PkgCount%).Visible = True
   Me.lblPackage(PkgCount%).Caption = "DAQFlex Library"
   PkgCount% = PkgCount% + 1
#End If
#If NETOPS Then
   Me.lblSupports.Visible = True
   Me.lblPackage(PkgCount%).Visible = True
   Me.lblPackage(PkgCount%).Caption = "Universal Library for .Net"
   PkgCount% = PkgCount% + 1
#End If

   Me.Top = Screen.Height / 2 - frmSplash.Height / 2
   Me.Left = Screen.Width / 2 - frmSplash.Width / 2
   AppVersion$ = App.Major & "." & App.Minor & "." & App.Revision
   lblVersion.Caption = "Version   " & AppVersion$
   
End Sub

Private Sub lblVersion_Click()

   Me.Hide
   
End Sub
