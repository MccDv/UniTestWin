VERSION 5.00
Begin VB.Form frmComposite 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Composite Read / Write"
   ClientHeight    =   2490
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5955
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
   ScaleHeight     =   2490
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSetUnselected 
      BackColor       =   &H80000009&
      Caption         =   "Set unselected channels to differential."
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   2100
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame fraChanSel 
      BackColor       =   &H80000009&
      Caption         =   "Channel Select"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5715
      Begin VB.TextBox txtChanList 
         ForeColor       =   &H00FF0000&
         Height          =   915
         Left            =   2700
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   660
         Width           =   2835
      End
      Begin VB.ListBox lstChanSelect 
         Height          =   1230
         ItemData        =   "Composit.frx":0000
         Left            =   120
         List            =   "Composit.frx":0007
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblNumChanLabel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "Num chans selected:"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lblNumChans 
         BackColor       =   &H80000009&
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5040
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Top             =   2040
      Width           =   795
   End
   Begin VB.CheckBox chkConsecutive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consecutive Registers"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CheckBox chkMaskSecond 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Apply Mask to Second Transfer"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   780
      Width           =   3015
   End
   Begin VB.CheckBox chkMaskFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Apply Mask to First Transfer"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.CheckBox chkComposite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Use Composite Read / Write"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   3015
   End
   Begin VB.TextBox txtShow 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   300
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frmComposite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkComposite_Click()

   chkMaskFirst.Enabled = (chkComposite.value = 1)
   chkMaskSecond.Enabled = (chkComposite.value = 1)
   chkConsecutive.Enabled = (chkComposite.value = 1)

End Sub

Private Sub cmdOK_Click()

   Me.Hide

End Sub

Private Sub lstChanSelect_Click()

   ChansAvailable& = Me.lstChanSelect.ListCount
   ChansSelected& = 0
   Prefix$ = ""
   Me.txtChanList.Text = ""
   For ChanItem& = 0 To ChansAvailable& - 1
      If lstChanSelect.Selected(ChanItem&) Then
         ChansSelected& = ChansSelected& + 1
         Me.txtChanList.Text = txtChanList.Text & _
            Prefix$ & Format(ChanItem&, "0")
         Prefix$ = ", "
      End If
   Next
   Me.lblNumChans.Caption = ChansSelected&
   
End Sub
