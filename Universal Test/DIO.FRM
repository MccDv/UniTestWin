VERSION 5.00
Begin VB.Form frmDigital 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Digital IO"
   ClientHeight    =   3180
   ClientLeft      =   1320
   ClientTop       =   2070
   ClientWidth     =   6930
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
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "frmDigital"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   6930
   Begin VB.CommandButton cmdClearAlarm 
      Appearance      =   0  'Flat
      Caption         =   "Clr Alm"
      Height          =   372
      Left            =   6060
      TabIndex        =   39
      Top             =   1140
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4980
      Top             =   1080
   End
   Begin VB.Frame fraScript 
      BackColor       =   &H80000005&
      Height          =   555
      Left            =   60
      TabIndex        =   37
      Top             =   1380
      Visible         =   0   'False
      Width           =   2475
      Begin VB.TextBox txtScript 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdConfigure 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   6060
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer tmrGoLoop 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4980
      Top             =   480
   End
   Begin VB.CheckBox chkUseBits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bits"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   2820
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraBits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3300
      TabIndex        =   26
      Top             =   3180
      Visible         =   0   'False
      Width           =   2595
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2220
         TabIndex        =   34
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   33
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1620
         TabIndex        =   32
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   31
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   30
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   29
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   28
         Top             =   180
         Width           =   195
      End
      Begin VB.CheckBox chkBit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   180
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdRead 
      Appearance      =   0  'Flat
      Caption         =   "&Read"
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdConfigIn 
      Appearance      =   0  'Flat
      Caption         =   "DIGITALIN"
      Height          =   375
      Left            =   1620
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdConfigOut 
      Appearance      =   0  'Flat
      Caption         =   "DIGITALOUT"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Timer tmrCheckStatus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   480
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      TabIndex        =   18
      Text            =   "1000"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      TabIndex        =   17
      Text            =   "1000"
      Top             =   420
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtOffset 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      TabIndex        =   16
      Text            =   "128"
      Top             =   1020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAmpl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3540
      TabIndex        =   15
      Text            =   "127"
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkAllPorts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "All Valid Ports"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1380
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CheckBox chkScrollUpdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Live Update"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1620
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Appearance      =   0  'Flat
      Caption         =   "&Write"
      Height          =   375
      Left            =   3300
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.HScrollBar hsbPort 
      Height          =   255
      Left            =   120
      Max             =   34
      TabIndex        =   9
      Top             =   2040
      Value           =   1
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   372
      Left            =   6060
      TabIndex        =   8
      Top             =   1140
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   372
      Left            =   6060
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      Caption         =   "&Go"
      Height          =   372
      Left            =   6060
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   860
   End
   Begin VB.PictureBox picAuxIO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   60
      ScaleHeight     =   -8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   2
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   375
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   0
         X2              =   2
         Y1              =   0.999
         Y2              =   0.999
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   0
         X2              =   2
         Y1              =   1.999
         Y2              =   1.999
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   0
         X2              =   2
         Y1              =   2.998
         Y2              =   2.998
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   0
         X2              =   2
         Y1              =   4.008
         Y2              =   4.008
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   0
         X2              =   2
         Y1              =   5.002
         Y2              =   5.002
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   2
         Y1              =   6.001
         Y2              =   6.001
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   2
         Y1              =   7.001
         Y2              =   7.001
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0.991
         X2              =   0.991
         Y1              =   8
         Y2              =   0
      End
   End
   Begin VB.Timer tmrDIn 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5520
      Top             =   1080
   End
   Begin VB.HScrollBar hsbOutVal 
      Height          =   255
      Left            =   3300
      Min             =   -32768
      TabIndex        =   2
      Top             =   2040
      Value           =   -32768
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOutVal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4500
      TabIndex        =   1
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picDIO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   540
      ScaleHeight     =   -8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   30
         X1              =   31.002
         X2              =   31.002
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   29
         X1              =   29.997
         X2              =   29.997
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   28
         X1              =   28.999
         X2              =   28.999
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   27
         X1              =   28.001
         X2              =   28.001
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   26
         X1              =   26.991
         X2              =   26.991
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   25
         X1              =   25.998
         X2              =   25.998
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   24
         X1              =   25
         X2              =   25
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   23
         X1              =   24.001
         X2              =   24.001
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   22
         X1              =   22.997
         X2              =   22.997
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   21
         X1              =   21.999
         X2              =   21.999
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   20
         X1              =   21.001
         X2              =   21.001
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   19
         X1              =   20.002
         X2              =   20.002
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   18
         X1              =   18.998
         X2              =   18.998
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   18
         X2              =   18
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   17.001
         X2              =   17.001
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   15.967
         X2              =   15.967
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   14.999
         X2              =   14.999
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   14
         X2              =   14
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   13.002
         X2              =   13.002
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   11.998
         X2              =   11.998
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   10.999
         X2              =   10.999
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   10.001
         X2              =   10.001
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   9.003
         X2              =   9.003
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   7.999
         X2              =   7.999
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   7
         X2              =   7
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   6.002
         X2              =   6.002
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   5.009
         X2              =   5.009
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   3.999
         X2              =   3.999
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   3.001
         X2              =   3.001
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   2.003
         X2              =   2.003
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0.998
         X2              =   0.998
         Y1              =   8
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   0
         X2              =   32
         Y1              =   0.999
         Y2              =   0.999
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   0
         X2              =   32
         Y1              =   1.999
         Y2              =   1.999
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   0
         X2              =   32
         Y1              =   2.998
         Y2              =   2.998
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   0
         X2              =   32
         Y1              =   4.008
         Y2              =   4.008
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   0
         X2              =   32
         Y1              =   5.002
         Y2              =   5.002
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   32
         Y1              =   6.001
         Y2              =   6.001
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   32
         Y1              =   7.001
         Y2              =   7.001
      End
   End
   Begin VB.Label lblNumPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Samples"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4560
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rate"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4560
      TabIndex        =   21
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblAmpl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Code Range"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4560
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Center Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4560
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblSelectPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Port"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " AUXPORT"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Label lblOutVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Output Value"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3300
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   8535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLibrary 
         Caption         =   "Universal Library"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "&Thread UL Calls"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "DAQFlex"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuBoardSel 
      Caption         =   "&Board"
      Begin VB.Menu mnuBoard 
         Caption         =   "None Installed"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnuFunc 
      Caption         =   "F&unction"
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDIn()"
         Index           =   0
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDIn32()"
         Index           =   1
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDInArray()"
         Index           =   2
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDBitIn()"
         Index           =   3
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDInScan()"
         Index           =   4
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDConfigPort()"
         Index           =   5
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDConfigBit()"
         Index           =   6
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetStatus 
         Caption         =   "cbGetStatus()"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStopBG 
         Caption         =   "cbStopBackground()"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "O&ptions"
      Begin VB.Menu mnuOpt 
         Caption         =   "BACKGROUND"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "CONTINUOUS"
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "EXTCLOCK"
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "WORDXFER"
         Index           =   3
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "DWORDXFER"
         Index           =   4
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "SINGLEIO"
         Index           =   5
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "DMAIO"
         Index           =   6
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "BLOCKIO"
         Index           =   7
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "DOUBLEBUF"
         Index           =   8
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "EXTTRIGGER"
         Index           =   9
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "NONSTREAMEDIO"
         Index           =   10
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "ADCCLOCKTRIG"
         Index           =   11
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "ADCCLOCK"
         Index           =   12
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "HIGHRESRATE"
         Index           =   13
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "RETRIGMODE"
         Index           =   14
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimer 
         Caption         =   "Timer Loop"
      End
      Begin VB.Menu mnuReadOnWrite 
         Caption         =   "Read On Write"
      End
      Begin VB.Menu mnuContPlot 
         Caption         =   "Continuous Plot"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDIODetect 
         Caption         =   "Use Mask for DIO Detect"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHex 
         Caption         =   "Use Hex text"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuSetTrig 
      Caption         =   "Trigg&ering"
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIGABOVE"
         Index           =   0
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIGBELOW"
         Index           =   1
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_NEG_HYS"
         Index           =   2
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_POS_HYS"
         Index           =   3
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_ABOVE"
         Index           =   4
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_BELOW"
         Index           =   5
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_IN_WINDOW"
         Index           =   6
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_OUT_WINDOW"
         Index           =   7
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_HIGH"
         Index           =   8
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "GATE_LOW"
         Index           =   9
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_HIGH"
         Index           =   10
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_LOW"
         Index           =   11
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_POS_EDGE"
         Checked         =   -1  'True
         Index           =   12
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_NEG_EDGE"
         Index           =   13
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIGRISING"
         Index           =   14
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIGFALLING"
         Index           =   15
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_PATTERN_EQ"
         Index           =   16
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_PATTERN_NQ"
         Index           =   17
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_PATTERN_ABOVE"
         Index           =   18
      End
      Begin VB.Menu mnuTrigType 
         Caption         =   "TRIG_PATTERN_BELOW"
         Index           =   19
      End
   End
   Begin VB.Menu mnuEvents 
      Caption         =   "Events"
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_SCAN_ERROR"
         Index           =   0
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_EXTERNAL_INTERRUPT..."
         Index           =   1
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_PRETRIGGER"
         Index           =   2
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_DATA_AVAILABLE..."
         Index           =   3
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_END_OF_INPUT_SCAN"
         Index           =   4
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_END_OF_OUTPUT_SCAN"
         Index           =   5
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "ON_CHANGE_DI"
         Index           =   6
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "Disable All Events"
         Index           =   7
      End
   End
   Begin VB.Menu mnuSelPlot 
      Caption         =   "Plot T&ype"
      Begin VB.Menu mnuPlotType 
         Caption         =   "Volts vs Time"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPlotType 
         Caption         =   "Histogram"
         Index           =   1
      End
      Begin VB.Menu mnuPlotType 
         Caption         =   "Text"
         Index           =   2
      End
      Begin VB.Menu mnuPlotType 
         Caption         =   "Single Points"
         Index           =   3
      End
   End
   Begin VB.Menu mnuSelData 
      Caption         =   "&Data Builder..."
   End
   Begin VB.Menu mnuUtils 
      Caption         =   "Utilitie&s"
      Begin VB.Menu mnuStore 
         Caption         =   "Store Data For Printing"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Data"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBufInfo 
         Caption         =   "Buffer Info"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEvalData 
         Caption         =   "Evaluate Data..."
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuULBuf 
         Caption         =   "UL Creates Buffer"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'temporary const dec for NONSTREAMEDIO
Const NONSTREAMEDIO = &H40000

Const CB_DIN = 0
Const CB_DIN32 = 1
Const DIN_ARRAY = 2
Const BIT_IN = 3
Const DIN_SCAN = 4

Const CB_DOUT = 0
Const CB_DOUT32 = 1
Const DOUT_ARRAY = 2
Const BIT_OUT = 3
Const DOUT_SCAN = 4

Const DCONFIG = 5
Const DBITCONF = 6

'DIGITALOUT = 1
'DIGITALIN = 2
Const NOTCONFIGURED = 3
Const INVALIDBIT = 4
Const INVALIDPORT = 5
Const MIXEDPORT = 6

Const NOCONFCOLOR& = &H8000&               '&HFF80FF
Const ONNOCONFCOLOR& = &HFFFF&           '&HFFF0FF
Const INOFFCOLOR& = &HFF0000
Const INONCOLOR& = &HFFC0C0
Const OUTOFFCOLOR& = &HC0&
Const OUTONCOLOR& = &HC0C0FF
Const INVALIDCOLOR& = &HC0C0C0

Const UPDATEDIO = 1
Const UPDATEGRID = 2

#If NETOPS Then
   'Dim WithEvents AcqThread As AcqThread.AcqThread
#End If
#If MSGOPS Then
   Dim WithEvents MsgLibrary As MBDClass.MBDComClass
Attribute MsgLibrary.VB_VarHelpID = -1
#Else
   Dim MsgLibrary As Object
#End If

Private Type DigitalGrid
   PortIndex As Long
   PortNum As Long
   PortName As String
   PortType As Long
   BitNum As Long
   ULBitNum As Long
   ULBitPort As Long
   ULBitPortName As String
   BitConf As Integer
End Type

Private DGrid(7, 34) As DigitalGrid

Private mfrmThisForm As Form
Private mnThisInstance As Integer

Private mnFormType As Integer, msTitle As String
Private mnDataType As Integer, mnTrigType As Integer
Private mnPlot As Integer, mnPlotType As Integer
Private mlEvalCount As Long
Private mbEvalEnabled As Boolean, mbEval As Boolean

Private mnBoardIndex As Integer
Private mnBoardNum As Integer, mnIsPCM As Integer
Private mnFuncType As Integer, mlStatusType As Long

Private msStringID As String, msConfig As String, msOpt As String
Private mnDIOPort() As Integer, mnDIOBit() As Integer, mnAuxBit() As Integer

Private mfRate As Single, mlCount As Long
Private mlFirstPoint As Long, mnDataChange As Integer
Private mnPortNum As Integer, mnPortIndex As Integer
Private masConfig(6) As String
Private mlStaticOpt As Long

Private mlHandle As Long, mvOptions As Variant
'Private mvOutHandle As Variant
Private mfRateReturned As Single

Private mnReadOnWrite As Integer
Private mlEventSize As Long
Private mlEventParam As Long
Private mlBlockSave As Long

Private mnLastAuxBit As Integer, mnLastDIOBit As Integer
Private mnFirstPort As Integer, mnLastPort As Integer
Private mlConfigIn(), mlConfigOut() As Long
Private mlPortType() As Long
Private mnAuxBitDisabled() As Integer
Private mnDIOBitEnable() As Integer
Private mnManualRefresh As Integer, mnDetailsNeedUpdate As Integer
Private mfrmDIOTarget As Form, mnTargetFormExists As Integer
Private mnNumSelPorts As Integer, malPortsSelected() As Long
Private malPortIndexSelected() As Long
Private mnNumSelBits As Integer, malBitsSelected() As Long
Private msBoardName As String
Private mlCurCount As Long, mlAcqPoints As Long
Private mlGenHandle As Long, mlGenPoints As Long
Private mnGenChans As Integer
Private mnAcqDataType As VbVarType, mnGenDataType As VbVarType
Private mlEventType As Long, mlEventData As Long
Private mnThreading As Integer, mnThreadState As Integer
Private msTempString As String
Private mvDataPoint As Variant

Private mnLibType As Integer, mnNumBoards As Integer
Private msStatusType As String   'added for msgDaq
Private mnUnloading As Integer, mnLoading As Integer
Private mnRefreshProps As Integer, mnMessaging As Integer
Private mnFormInitialized As Integer, msDisplayName As String

Private mnUseDGrid As Integer, mnLastColumn As Integer
Private msDIOSupport As String, mbGoClickHandled As Boolean
Private mnResolution As Integer, mbDataGenRequired As Boolean
Private mlTotalBits As Long, mlLastPortIndex As Long
Private malPortBits() As Long, malBitsInPort() As Long
Private mnMaxPortWidth As Integer, mnDelayRestart As Integer
Private mnUse32 As Integer, mnTimerStop As Integer
Private mlDelayTime As Long, mnCancel As Integer
Private mbDeviceInUse As Boolean, mbHexText As Boolean
Private msReadType As String, mbReadInvert As Boolean
Private mnCountRolls As Integer, mbHalfRoll As Integer
Private mnLowThreshold As Integer, mnHighThreshold As Integer
Private mbThresholdSet As Boolean

Private Sub chkAllPorts_Click()

   hsbPort.ENABLED = (chkAllPorts.value = 0)

End Sub

Private Sub chkBit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   
   If Not (chkBit(Index).value = 2) Then
      BitValue% = chkBit(Index).value
      If mnLibType = MSGLIB Then
         BitIndex% = Index
         Port% = Me.hsbPort.value - 1
         ValToSet% = BitValue%
         BitValSet% = SetMsgBitVal(Port%, BitIndex%, ValToSet%)
         If Not (BitValSet% = ValToSet%) Then Exit Sub
      Else
      ULStat = cbDBitOut(mnBoardNum, mnPortNum, Index, BitValue%)
      If (ULStat = BADBITNUMBER) Or (ULStat = CBUSB_INVALID_BIT_NUM And VarType(CBUSB_INVALID_BIT_NUM) > 1) Then
         chkBit(Index).value = 2
         chkBit(Index).ENABLED = False
         d% = SaveFunc(Me, DBitOut, ULStat, mnBoardNum, mnPortNum, Index, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0)
      Else
         If SaveFunc(Me, DBitOut, ULStat, mnBoardNum, mnPortNum, Index, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      End If
      End If
   End If

End Sub

Private Sub chkScrollUpdate_Click()

   cmdSend.ENABLED = Not ((chkScrollUpdate = 1) And (Me.chkUseBits.value = 0))

End Sub

Private Sub chkUseBits_Click()

   'hsbOutVal.Visible = (chkUseBits.value = 0)
   'lblOutVal.Visible = (chkUseBits.value = 0)
   'txtOutVal.Visible = (chkUseBits.value = 0)
   'cmdSend.Visible = (chkUseBits.value = 0)
   'cmdRead.Visible = (chkUseBits.Value = 0)
   'fraBits.Visible = Not (chkUseBits.value = 0)
   chkScrollUpdate.Visible = Not (chkUseBits.value = 0)
   If chkUseBits.value = 0 Then
      lblOutVal.Caption = "Select Output Value"
      chkScrollUpdate.Top = 740
      chkScrollUpdate.Left = 4560
      chkScrollUpdate.Caption = "Live Update"
   Else
      lblOutVal.Caption = "Select Bit Number"
      chkScrollUpdate.Top = 1140
      chkScrollUpdate.Left = Me.chkUseBits.Left
      chkScrollUpdate.Caption = "Bit Value"
   End If

End Sub

Private Sub cmdClearAlarm_Click()

   For DIndex& = 0 To mnLastPort - 1
      y& = cbGetConfig(DIGITALINFO, mnBoardNum, DIndex&, DIDEVTYPE, DType&)
      If DType& = mnPortNum Then
         PortIndex& = DIndex&
         Exit For
      End If
   Next
   y& = cbGetConfig(DIGITALINFO, mnBoardNum, PortIndex&, DINUMBITS, NumBits&)
   If NumBits& > 30 Then
      HalfBits& = NumBits& / 2
      DefaultVal$ = Hex(2 ^ HalfBits& - 1) & Hex(2 ^ HalfBits& - 1)
   Else
      DefaultVal$ = Hex(2 ^ NumBits& - 1)
   End If
   Maskval$ = InputBox("Alarm mask", "Clear Alarm", DefaultVal$)
   If Not (Maskval$ = "") Then
      Mask& = Val("&H" & Maskval$)
      ULStat = cbDClearAlarm(mnBoardNum, mnPortNum, Mask&)
      x% = SaveFunc(Me, DClearAlarm, ULStat, mnBoardNum, _
         mnPortNum, Mask&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      If ULStat = BADPORTNUM Then
         PortName$ = GetPortString(mnPortNum)
         MsgBox "Selected port is " & PortName$ & _
            ". Select the appropriate port using the Menu " & _
            "'Function/cbDConfigPort' for this window.", _
            vbInformation, "Select Another Port"
      End If
   End If
   
End Sub

Private Sub cmdConfigIn_Click()

   If chkAllPorts Then
      SaveIndex% = mnPortIndex
      SavePort% = mnPortNum
      For mnPortIndex = 0 To 33
         PortNum& = DGrid(0, mnPortIndex).PortNum
         If PortNum& > 0 Then
            ULStat = cbDConfigPort(mnBoardNum, PortNum&, DIGITALIN)
            x% = SaveFunc(Me, DConfigPort, ULStat, mnBoardNum, _
               PortNum&, DIGITALIN, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = NOTCONFIGURABLE Then
               mnDIOPort(mnPortIndex, mnBoardIndex) = NOTCONFIGURED
            Else
               mnDIOPort(mnPortIndex, mnBoardIndex) = DIGITALIN
            End If
         End If
      Next
      GetPortConfig
      ShowConfigStat
      RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      mnPortIndex = SaveIndex%
      mnPortNum = SavePort%
   Else
      Select Case mnLibType
         Case UNILIB
            ULStat = cbDConfigPort(mnBoardNum, mnPortNum, DIGITALIN)
            If SaveFunc(Me, DConfigPort, ULStat, mnBoardNum, mnPortNum, _
               DIGITALIN, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            mnDIOPort(mnPortIndex, mnBoardIndex) = DIGITALIN
            GetPortConfig
            ShowConfigStat
            RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
         Case NETLIB
            functionToCall$ = GetCallAsString(DConfigPort, mnBoardNum, mnPortNum, DIGITALIN)
            'AcqThread.CallULFunction functionToCall$
            Do
               DoEvents
            Loop While Not ((mnThreadState = TS_IDLE) Or _
            (mnThreadState = TS_ERROR))
            If Not (mnThreadState = TS_ERROR) Then mnDIOPort(mnPortIndex, mnBoardIndex) = DIGITALIN
         Case MSGLIB
            MsgPort$ = "{" & Format(hsbPort.value - 1, "0") & "}"
            DConfigMessage$ = msStatusType & MsgPort$ & ":DIR=IN"
            MsgResult$ = MsgLibrary.SendMessage(DConfigMessage$)
            If SaveMsg(Me, "SendMessage(" & DConfigMessage$ & ")", MsgResult$) Then Failure% = True
            RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      End Select
   End If
   If mnUseDGrid Then InitDGrid
   'RefreshDGrid
   ShowConfigStat

End Sub

Private Sub cmdConfigOut_Click()

   If chkAllPorts Then
      SaveIndex% = mnPortIndex
      SavePort% = mnPortNum
      For mnPortIndex = 0 To 33
         PortNum& = DGrid(0, mnPortIndex).PortNum
         If PortNum& > 0 Then
            ULStat = cbDConfigPort(mnBoardNum, PortNum&, DIGITALOUT)
            x% = SaveFunc(Me, DConfigPort, ULStat, mnBoardNum, PortNum&, DIGITALOUT, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = NOTCONFIGURABLE Then
               mnDIOPort(mnPortIndex, mnBoardIndex) = NOTCONFIGURED
            Else
               mnDIOPort(mnPortIndex, mnBoardIndex) = DIGITALOUT
            End If
         End If
      Next
      GetPortConfig
      ShowConfigStat
      RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      mnPortIndex = SaveIndex%
      mnPortNum = SavePort%
   Else
      Select Case mnLibType
         Case UNILIB
            ULStat = cbDConfigPort(mnBoardNum, mnPortNum, DIGITALOUT)
            If SaveFunc(Me, DConfigPort, ULStat, mnBoardNum, mnPortNum, DIGITALOUT, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            GetPortConfig
            ShowConfigStat
            RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
            mnDIOPort(mnPortIndex, mnBoardIndex) = DIGITALOUT
         Case NETLIB
            functionToCall$ = GetCallAsString(DConfigPort, mnBoardNum, mnPortNum, DIGITALOUT)
            'AcqThread.CallULFunction functionToCall$
            Do
               DoEvents
            Loop While Not ((mnThreadState = TS_IDLE) Or _
            (mnThreadState = TS_ERROR))
            If Not (mnThreadState = TS_ERROR) Then mnDIOPort(mnPortIndex, mnBoardIndex) = DIGITALOUT
         Case MSGLIB
            MsgPort$ = "{" & Format(hsbPort.value - 1, "0") & "}"
            DConfigMessage$ = msStatusType & MsgPort$ & ":DIR=OUT"
            MsgResult$ = MsgLibrary.SendMessage(DConfigMessage$)
            If SaveMsg(Me, "SendMessage(" & DConfigMessage$ & ")", MsgResult$) Then Failure% = True
            RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      End Select
   End If
   If mnUseDGrid Then InitDGrid
   ShowConfigStat

End Sub

Private Sub cmdConfigure_Click()

   'this exists to give menu access to the scripting
   'form when running scripts
   CmdStr$ = Left$(cmdConfigure.Caption, 1)
   value& = Val(Mid$(cmdConfigure.Caption, 2))
   Select Case CmdStr$
      Case "$" 'return board name
         cmdConfigure.Caption = msBoardName
      Case "0" 'return handle to data for output
         Me.cmdConfigure.Caption = Format(mlGenHandle, "0")
      Case "1" 'return handle to data
         If Not (mfrmDIOTarget Is Nothing) Then
            'if there is a target form, use the handle to data from that
            mfrmDIOTarget.cmdConfigure.Caption = "1"
            mfrmDIOTarget.cmdConfigure = True
            Me.cmdConfigure.Caption = mfrmDIOTarget.cmdConfigure.Caption
         Else
            Me.cmdConfigure.Caption = Format(mlHandle, "0")
         End If
      Case "2" 'return data resolution
         Me.cmdConfigure.Caption = Format(mnResolution, "0")
      Case "3" 'return range at which data was collected
         If mnFixedRange = -1 Then
            Range% = mnRange
         Else
            Range% = mnFixedRange
         End If
         Me.cmdConfigure.Caption = Format(Range%, "0")
      Case "4" 'return rate requested
         Me.cmdConfigure.Caption = Format(mfRate, "0")
      Case "5" 'return rate returned
         Me.cmdConfigure.Caption = Format(mfRateReturned, "0")
      Case "6" 'get reference to another form
         FormID$ = Format(value&, "0")
         If Val(FormID$) = 0 Then
            Set mfrmDIOTarget = Nothing
            mnTargetFormExists = False
         Else
            mnTargetFormExists = GetFormReference(FormID$, mfrmDIOTarget)
            Me.cmdConfigure.Caption = Format(mnTargetFormExists, "0")
         End If
      Case "8"
         CurOption$ = msOpt
         If msOpt = "Options = Default  " Then CurOption$ = ""
         Me.cmdConfigure.Caption = CurOption$
      Case "9" 'bit out FIRSTPORTA
         BitString$ = cmdConfigure.Caption
         EndFirst% = InStr(2, cmdConfigure.Caption, ":")
         EndSecond% = InStr(EndFirst% + 1, cmdConfigure.Caption, ":")
         x! = Val(Mid$(BitString$, 2, EndFirst%))
         y! = Val(Mid$(BitString$, EndFirst% + 1, EndSecond% - EndFirst%))
         BitValue% = Val(Right$(BitString$, 1))
         picDIO_MouseDown 0, BitValue%, x!, y!
      Case "="
         TotalCount& = mlCount
         If Not mlTotalCountReturn = 0 Then
            If mlTotalCountReturn < mlCount Then TotalCount& = mlTotalCountReturn
         End If
         Me.cmdConfigure.Caption = Format(TotalCount&, "0")
      Case "?"
         t$ = Format(mnNumSelPorts + 1, "0")
         cmdConfigure.Caption = t$
      Case "A" 'set continuous plot on or off
         If value& And Not mnuContPlot.Checked Then mnuContPlot_Click
         If (Not value&) And mnuContPlot.Checked Then mnuContPlot_Click
      Case "B" 'set board number
         SearchName$ = Mid$(cmdConfigure.Caption, 2)
         BoardParams = Split(SearchName$, ",")
         If UBound(BoardParams) = 1 Then DupeIndex% = Val(BoardParams(1))
         NotFound% = True
         For MenuIndex% = 0 To mnNumBoards - 1
            NameStart% = InStr(mnuBoard(MenuIndex%).Caption, ") ") + 2
            If Mid$(mnuBoard(MenuIndex%).Caption, NameStart%) = BoardParams(0) Then
               BoardFound% = True
               FirstIndex% = MenuIndex%
               If DupeFound% = DupeIndex% Then
                  mnuBoard_Click (MenuIndex%)
                  NotFound% = False
                  Exit For
               End If
               DupeFound% = DupeFound% + 1
            End If
         Next MenuIndex%
         If NotFound% Then
            If BoardFound% Then
               mnuBoard_Click (FirstIndex%)
            Else
               MsgBox SearchName$ & " not available in list of currently installed boards. Aborting script.", , "Requested Board Not Available"
               gnScriptRun = False
            End If
         End If
      Case "C" 'set total count
         txtCount.Text = value&
      Case "D" 'set data array
         'mnuData_Click (value&)
      Case "E" 'set direction and run config function
         If value& = DIGITALOUT Then cmdConfigOut = True
         If value& = DIGITALIN Then cmdConfigIn = True
      Case "F" 'set function
         'compensate for addition of DIOArray
         FuncMenuIndex& = Choose(value& + 1, 0, 2, 3, 4, 5)
         mnuFuncArray_Click (FuncMenuIndex&)
      Case "G" 'config bit AUXPORT
         BitString$ = cmdConfigure.Caption
         EndFirst% = InStr(2, cmdConfigure.Caption, ":")
         EndSecond% = InStr(EndFirst% + 1, cmdConfigure.Caption, ":")
         x! = Val(Mid$(BitString$, 2, EndFirst%))
         y! = Val(Mid$(BitString$, EndFirst% + 1, EndSecond% - EndFirst%))
         BitValue% = Val(Right$(BitString$, 1))
         Direction% = 0
         If BitValue% = DIGITALOUT Then Direction% = 1
         picAuxIO_MouseDown 0, Direction%, x!, y!
      Case "I" 'read all bits
         DoBitIn
      Case "J" 'set offset
         txtOffset.Text = value&
         txtOffset_KeyPress (13)
         ResetData
      Case "M" 'set amplitude
         txtAmpl.Text = value&
         txtAmpl_KeyPress (13)
         ResetData
      Case "O" 'set an option (or clear all)
         If value& = -1 Then
            ResetOpts
         Else
            mnuOpt_Click (value&)
         End If
      Case "P" 'set port
         'ScrollVal& = Value&
         'If ScrollVal& > 1 Then ScrollVal& = Value& - 8
         hsbPort.value = value&
      Case "Q" 'get background status
         ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
         ret% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0)
         StatReport$ = Str(ULStat) & "," & Str(Status%) & "," & Str(CurCount&) & "," & Str(CurIndex&)
         Me.cmdConfigure.Caption = StatReport$
         If (Status% = RUNNING) Then
            BGOp$ = "Background task in progress..."
         Else
            BGOp$ = "No background task running..."
         End If
         Stat$ = BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
         lblStatus.Caption = Stat$
         DoEvents
      Case "R" 'read port
         If mnFuncType = DIn Then
            DoDIn
         Else
            cmdRead = True
         End If
     Case "T" 'set rate
        txtRate.Text = value&
      Case "V" 'set output value
         txtOutVal.Text = value&
         mnDataChange = True
         ResetData
      Case "W" 'write port
         cmdSend = True
      Case "X" 'stop background
         If (cmdStop.Visible And cmdStop.ENABLED) Then
            cmdStop_Click
         Else
            ULStat = StopBackground520(mnBoardNum, mlStatusType)
         End If
      Case "Z" 'bit out AUXPORT
         BitString$ = cmdConfigure.Caption
         EndFirst% = InStr(2, cmdConfigure.Caption, ":")
         EndSecond% = InStr(EndFirst% + 1, cmdConfigure.Caption, ":")
         x! = Val(Mid$(BitString$, 2, EndFirst%))
         y! = Val(Mid$(BitString$, EndFirst% + 1, EndSecond% - EndFirst%))
         BitValue% = Val(Right$(BitString$, 1))
         picAuxIO_MouseDown 0, BitValue%, x!, y!
      Case "c" 'configure selected ports
         Port% = value&
         ConfigSelectedPorts Port%
      Case "d" 'disable event
         EventType& = value&
         If EventType& = 255 Then
            mnuEvent_Click (6)
         Else
            MenuIndex% = Log(value&) / Log(2) '- 1
            If mnuEvent(MenuIndex%).Checked Then mnuEvent_Click (MenuIndex%)
         End If
      Case "e" 'set events
         Set mfrmThisForm = Me
         EventType& = value&
         If EventType& = 255 Then
            mnuEvent_Click (6)
         Else
            MenuIndex% = Log(value&) / Log(2) '- 1
            If Not mnuEvent(MenuIndex%).Checked Then mnuEvent_Click (MenuIndex%)
         End If
      'Case "g"
      '   Result% = GenerateData(value&)
      '   Me.cmdConfigure.Caption = Format(Result%, "0")
      'Case "i" 'read selected ports
      '   ReadSelectedPorts
      Case "j" 'set offset
         txtOffset.Text = value&
      'Case "m" 'set amplitude
      '   txtAmpl.Text = value&
      Case "m" 'set block size
         'mlBlockSize = value&
         SetBlockSize value&, False
         InitBlock True
         If Not mlHandle = 0 Then TransferBlock
      'Case "o" 'write selected ports
      '   WriteSelectedPorts
      Case "s" 'select multiple ports or clear all
         PortCode$ = Mid$(cmdConfigure.Caption, 2)
         Location = InStr(1, PortCode$, ",")
         Locn& = Val(Location)
         If Locn& > 1 Then
            FirstPort% = Val(Left(PortCode$, Locn& - 1))
            LastPort% = Val(Mid(PortCode$, Locn& + 1))
         Else
            FirstPort% = Val(PortCode$)
            LastPort% = Val(PortCode$)
         End If
         SelectPorts FirstPort%, LastPort%
   End Select

End Sub

Private Sub cmdGo_Click()

#If NETOPS Then
   If gnThreading And Not mnThreading Then
      'Set AcqThread = New AcqThread.AcqThread
      'AcqThread.Init
      'x& = AcqThread.ThreadID
   Else
      msStringID = ""
   End If
#End If
      
   mnThreading = gnThreading
   mfRate = Val(txtRate.Text)
   mnCancel = False
   mnCountRolls = 1
   
   If mnuContPlot.Checked Then mnSetActive = True
   If mnuTimer.Checked And (mnLoopType = PROGLOOP) Then
      mnLoop = True
      cmdStop.ENABLED = True
      If (mnFuncType = DInScan) Or (mnFuncType = DOutScan) Then
         tmrGoLoop.ENABLED = True
      Else
         tmrDIn.ENABLED = True
      End If
   ElseIf mnuTimer.Checked And (mnLoopType = TIMERLOOP) Then
      mnLoop = False
      cmdStop.ENABLED = True
      tmrGoLoop.ENABLED = True
   End If
   If mbGoClickHandled Then
      mbGoClickHandled = False
      Exit Sub
   End If
   Select Case mnFuncType
      Case DConfigPort
      Case DBitIn, DIn, DIn32, DInArray
         If mnuTimer.Checked Then
            tmrDIn.ENABLED = True
            cmdStop.Visible = True
            cmdStop.ENABLED = True
         Else
            Select Case mnFuncType
               Case DIn, DIn32, DInArray
                  If mnLibType = UNILIB Then
                     DoDIn
                  Else
                     DoMsgDIn
                  End If
               Case DBitIn
                  If mnLibType = UNILIB Then
                     DoBitIn
                  Else
                     DoMsgBitIn
                  End If
               Case DInScan
            End Select
         End If
      Case DInScan
         If mnuTimer.Checked Then
            tmrGoLoop.ENABLED = True
            cmdStop.Visible = True
            cmdStop.ENABLED = True
         End If
         DoDInScan
         DisplayData
      Case DBitOut, DOut, DOut32
         'tmrDIn.Enabled = True
         'cmdStop.Visible = True
      Case DOutScan
         If mbDataGenRequired Then
            'Bits% = GetDIOResolution()
            'Unload frmBuildData
            mnuSelData_Click
            Exit Sub
         End If
         If mnuTimer.Checked Then
            tmrGoLoop.ENABLED = True
            cmdStop.Visible = True
            cmdStop.ENABLED = True
         End If
         DoDOutScan
   End Select

   If (Not tmrCheckStatus.ENABLED) And mnuStopBG.Checked Then
      'ULStat = cbStopBackground(mnBoardNum)
      ULStat = StopBackground520(mnBoardNum, mlStatusType)
      If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, _
         mnFormType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
         Then Exit Sub
   End If

End Sub

Private Sub cmdGo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   Dim TimerInUse As Timer
   
   If (mnFuncType = DInScan) Or (mnFuncType = DOutScan) Then
      Set TimerInUse = tmrGoLoop
   Else
      Set TimerInUse = tmrDIn
   End If
   
   If Button = 2 Then
      'if right click, change timer state
      mbGoClickHandled = True
      mnLoopType = TIMERLOOP
      If mnuTimer.Checked Then
         'if already checked, just change loop rate
         LoopInterval% = TimerInUse.Interval  ' tmrDIn.Interval
         If Shift > 0 Then
            If LoopInterval% < 16384 Then LoopInterval% = LoopInterval% * 2
         Else
            If LoopInterval% > 1 Then LoopInterval% = LoopInterval% / 2
         End If
         'tmrDIn.Interval = LoopInterval%
         TimerInUse.Interval = LoopInterval%
      End If
      mnuTimer.Checked = True
      cmdGo = True
      cmdStop.Visible = True
   End If

End Sub

Private Sub cmdRead_Click()

   Dim PortIsValid As Boolean
   
   PortIsValid = False
   If (chkAllPorts.value = 1) Then
      'use selected port if valid
      'otherwise, use first valid port
      If Not (mnPortNum = 0) Then
         For PortIndex& = 0 To 33
            PortNum& = DGrid(0, PortIndex&).PortNum
            If PortNum& = mnPortNum Then
               PortIsValid = True
               Exit For
            End If
         Next
      End If
      If Not PortIsValid Then
         PortNum& = DGrid(0, 0).PortNum
      End If
   Else
      PortNum& = mnPortNum
   End If
   If chkUseBits.value Then
      BitNum& = Me.txtOutVal.Text
      If mnLibType = MSGLIB Then
         CompMsg$ = "?DIO"
         PortBitString$ = "{" & Format(mnPortIndex, "0") & "/" & Format(i%, "0") & "}:VALUE"
         BitValMsg$ = CompMsg$ & PortBitString$
         MsgResult$ = MsgLibrary.SendMessage(BitValMsg$)
         If SaveMsg(Me, "SendMessage(" & BitValMsg$ & ")", MsgResult$) Then Exit Sub
         ValueLoc& = InStr(1, MsgResult$, "=")
         If ValueLoc& > 0 Then
            BitValString$ = Mid(MsgResult$, ValueLoc& + 1)
            BitValue% = Val(BitValString$)
         End If
      Else
         ULStat = cbDBitIn(mnBoardNum, PortNum&, BitNum&, BitValue%)
         If (ULStat = BADBITNUMBER) Or (ULStat = CBUSB_INVALID_BIT_NUM And VarType(CBUSB_INVALID_BIT_NUM) > 1) Then
            chkBit(i%).value = 2
            d% = SaveFunc(Me, DBitIn, ULStat, mnBoardNum, PortNum&, BitNum&, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0)
         Else
            If SaveFunc(Me, DBitIn, ULStat, mnBoardNum, PortNum&, BitNum&, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            'chkBit(i%).value = BitValue%
            If ULStat = 0 Then Me.chkScrollUpdate.value = BitValue%
         End If
      End If
   Else
      If mnLibType = MSGLIB Then
         CompMsg$ = "?DIO"
         PortString$ = "{" & Format(mnPortIndex, "0") & "}:VALUE"
         PortValMsg$ = CompMsg$ & PortString$
         MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
         If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Sub
         ValueLoc& = InStr(1, MsgResult$, "=")
         If ValueLoc& > 0 Then
            PortValString$ = Mid(MsgResult$, ValueLoc& + 1)
         End If
         txtOutVal.Text = PortValString$
      Else
         ULStat = cbDIn32(mnBoardNum, PortNum&, lDataValue&)
         If SaveFunc(Me, DIn32, ULStat, mnBoardNum, PortNum&, _
            lDataValue&, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         'txtOutVal.Text = IntValToULong(DataValue%)
         If mbHexText Then
            ValText$ = "0x" & Hex(lDataValue&)
         Else
            DWordVal# = LongToUDWordVal(lDataValue&)
            ValText$ = Format(DWordVal#, "0")
         End If
         txtOutVal.Text = ValText$
      End If
   End If
   
End Sub

Private Sub cmdReset_Click()

   If mnLibType = UNILIB Then
      If mnUseDGrid Then
         'ResetDGrid False
         InitDGrid
         If ((mnFuncType = DBitIn) Or (mnFuncType = DBitOut) _
            Or (mnFuncType = DConfigBit)) _
            Then RefreshDGrid
      End If
   Else
      NumProps% = GetDIOProps(msBoardName, MsgLibrary, PropsList)
      UpdateFormProps PropsList, NumProps%
      If ((mnFuncType = DBitIn) Or (mnFuncType = DBitOut) _
         Or (mnFuncType = DConfigBit)) Then
         If mnLibType = MSGLIB Then
            InitDGrid
            RefreshDGrid
         Else
            InitDGrid
            RefreshGrid
         End If
      End If
   End If
   Exit Sub
   For PortIndex% = 1 To 32
      'check if port is valid
      If mnDIOPort(PortIndex%, mnBoardIndex) <> INVALIDPORT Then
         PortNum% = PortIndex% + 9
         ULStat = cbDConfigPort(mnBoardNum, PortNum%, DIGITALIN)
         If SaveFunc(Me, DConfigPort, ULStat, mnBoardNum, PortNum%, DIGITALIN, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         If ULStat = 0 Then mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALIN
         If mnuFuncArray(BIT_IN).Checked Then
            Col% = PortIndex% - 1
            For d% = 0 To 7
               PaintDIOGrid Col%, d%, NOCONFCOLOR&
            Next d%
         End If
      End If
   Next PortIndex%
   RunOtherForms DIGITAL_IN, UPDATEDIO

End Sub

Private Sub cmdSend_Click()

   If mnFuncType = DConfigPort Then
      If mnLibType = MSGLIB Then
         DigitalValue$ = txtOutVal.Text
         CompMsg$ = "DIO"
         PortString$ = "{" & Format(mnPortIndex, "0") & "}:VALUE=" & DigitalValue$
         PortValMsg$ = CompMsg$ & PortString$
         MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
         If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Sub
      Else
         If Me.chkUseBits.value = 0 Then
            DoDOut
            'DataValue% = ULongValToInt(Val(txtOutVal.Text))
            'ULStat = cbDOut(mnBoardNum, mnPortNum, DataValue%)
            'If SaveFunc(Me, DOut, ULStat, mnBoardNum, mnPortNum, _
            'DataValue%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Else
            BitNum& = Me.txtOutVal.Text
            DataValue% = Me.chkScrollUpdate.value
            ULStat = cbDBitOut(mnBoardNum, mnPortNum, BitNum&, DataValue%)
            If SaveFunc(Me, DBitOut, ULStat, mnBoardNum, mnPortNum, _
            BitNum&, DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
      End If
   Else
      If chkAllPorts Then
         SaveIndex% = mnPortIndex
         SavePort% = mnPortNum
         For mnPortIndex = 0 To 33
            mnPortNum = DGrid(0, mnPortIndex).PortNum
            If mnPortNum > 0 Then DoDOut
         Next
         mnPortIndex = SaveIndex%
         mnPortNum = SavePort%
      Else
         If mnLibType = UNILIB Then
            DoDOut
         Else
            DoMsgDOut
         End If
      End If
   End If
   RunOtherForms DIGITAL_IN, UPDATEDIO

End Sub

Public Sub eventPlot()

   'this exists so the function defined by the
   'cbEnableEvent function can access the plot subroutine
   SavePlot% = mnPlot
   For i% = 0 To 3
      If mnuPlotType(i%).Checked Then EnablePlot% = True
   Next
   mnPlot = EnablePlot%
   mbEval = mbEvalEnabled
   DisplayData
   mnPlot = SavePlot%

End Sub

Private Sub cmdStop_Click()

   tmrDIn.ENABLED = False
   tmrGoLoop.ENABLED = False
   tmrDelay.ENABLED = False
   mnCancel = True
   cmdGo.FontBold = True
   ULStat = StopBackground520(mnBoardNum, mlStatusType)
   If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnFormType, _
   A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
   If Failure% Then
      cmdStop.ENABLED = False
      Exit Sub
   End If
   cmdStop.ENABLED = False

End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   'if right click, turn off timer
   If Button = 2 Then
      mnuTimer.Checked = False
      cmdStop.Visible = mnuTimer.Checked
      tmrGoLoop.ENABLED = False
      tmrDIn.ENABLED = False
      cmdGo.FontBold = True
   End If

End Sub

Private Function CompensateGrid(BitNum%) As Integer
   
   If BitNum% = 0 Then
      CompensateGrid = BitNum%
      Exit Function
   End If

   board$ = Right$(mnuBoard(mnBoardIndex).Caption, 7)
   If board$ = "-DISO48" Then
      CompensateGrid = (BitNum% Mod 8) + (BitNum% \ 32) * 8
      Exit Function
   End If
   If board$ = "RELAY32" Or board$ = "RELAY24" Then
      CompensateGrid = (BitNum% Mod 16) + (BitNum% \ 32) * 16
      Exit Function
   End If
   If mnIsPCM Then
      If 3 Mod BitNum% = 3 Then CompC% = CompC% + 4
   End If
   For Index% = 19 To 243 Step 32
      If Index% Mod BitNum% = Index% Then CompC% = CompC% + 4
      If (Index% + 8) Mod BitNum% = (Index% + 8) Then CompC% = CompC% + 4
   Next Index%

   CompensateGrid = BitNum% - CompC%

End Function

Private Function CompensatePort(BitNum%) As Integer
   
   'returns values 1 - 32 for port numbers based on
   'bit number (add 9 for library constants)

   CompC% = 1
   If BitNum% = 0 Then
      CompensatePort = CompC%
      Exit Function
   End If

   If mnIsPCM Then
'      If 3 Mod BitNum% = 3 Then CompC% = CompC% + 1
      CompC% = CompC% + BitNum% \ 4
      If CompC% > 32 Then CompC% = 32
   Else
      For Index% = 7 To BitNum% Step 24
         If Index% Mod BitNum% = Index% Then CompC% = CompC% + 1
         If Index% + 8 < BitNum% Then
            If Index% + 8 Mod BitNum% = Index% + 8 Then CompC% = CompC% + 1
         End If
      Next Index%
      For Index% = 19 To BitNum% Step 24
         If Index% Mod BitNum% = Index% Then CompC% = CompC% + 1
         If Index% + 4 < BitNum% Then
            If Index% + 4 Mod BitNum% = Index% + 4 Then CompC% = CompC% + 1
         End If
      Next Index%
   End If

   CompensatePort = CompC%

End Function

Private Sub ConfigureControls()
   
   'display appropriate controls and
   'hide others
   ShowInputCtls% = (mnFormType = DIGITAL_IN)
   ShowOutputCtls% = (mnFormType = DIGITAL_OUT)
   ShowDIOCtls% = mnuFuncArray(CB_DIN).Checked _
      Or mnuFuncArray(CB_DIN32).Checked _
      Or mnuFuncArray(DIN_ARRAY).Checked
   ShowBitIOCtls% = mnuFuncArray(BIT_IN).Checked
   ShowScanCtls% = mnuFuncArray(DIN_SCAN).Checked
   ShowConfigCtls% = mnuFuncArray(DCONFIG).Checked
   ShowBitConfigCtls% = mnuFuncArray(DBITCONF).Checked

   cmdClearAlarm.Visible = ShowInputCtls% And (ShowDIOCtls% Or ShowBitIOCtls%)
   If ShowConfigCtls% Then hsbPort_Change
   mnuSelData.Visible = ShowOutputCtls%
   mnuStopBG.Checked = ShowScanCtls%
   'mnuReadOnWrite.Visible = ShowOutputCtls%
   'If Not ShowOutputCtls% Then mnuReadOnWrite.Checked = False
   'If ShowScanCtls% Then mnuReadOnWrite.Checked = False
   mnuTimer.ENABLED = ShowConfigCtls% Or ShowInputCtls% Or ShowOutputCtls%

   lblStatus.Caption = msConfig & "  " & msOpt
   hsbOutVal.Visible = (ShowOutputCtls% And ShowDIOCtls%) Or ShowConfigCtls%
   txtOutVal.Visible = (ShowOutputCtls% And ShowDIOCtls%) Or ShowConfigCtls%
   lblOutVal.Visible = (ShowOutputCtls% And ShowDIOCtls%) Or ShowConfigCtls%
   If ShowConfigCtls% Then Me.chkScrollUpdate.value = 0
   chkUseBits.Visible = ShowConfigCtls%
   If Not ShowConfigCtls% Then chkUseBits.value = 0
   lblSelectPort.Visible = (ShowOutputCtls% And ShowDIOCtls%) Or ShowConfigCtls%
   hsbPort.Visible = ((ShowDIOCtls% And ShowOutputCtls%) Or ShowScanCtls%) Or ShowConfigCtls%
   lblPort.Visible = (ShowOutputCtls% Or ShowScanCtls%) Or ShowConfigCtls%
   cmdSend.Visible = (ShowOutputCtls% And ShowDIOCtls%) Or ShowConfigCtls%
   cmdRead.Visible = ShowConfigCtls%
   cmdReset.Visible = (ShowOutputCtls% Or ShowInputCtls%) And Not (ShowScanCtls% Or ShowConfigCtls%)
   chkScrollUpdate.Visible = (ShowOutputCtls% And ShowDIOCtls%) 'Or ShowConfigCtls%
   chkAllPorts.Visible = (ShowOutputCtls% And ShowDIOCtls%) Or ShowConfigCtls%
   txtRate.Visible = ShowScanCtls%
   txtCount.Visible = ShowScanCtls%
   'txtOffset.Visible = ShowOutputCtls% And ShowScanCtls%
   'txtAmpl.Visible = ShowOutputCtls% And ShowScanCtls%
   lblRate.Visible = ShowScanCtls%
   lblNumPoints.Visible = ShowScanCtls%
   'lblOffset.Visible = ShowOutputCtls% And ShowScanCtls%
   'lblAmpl.Visible = ShowOutputCtls% And ShowScanCtls%
   cmdConfigOut.Visible = ShowConfigCtls%
   cmdConfigIn.Visible = ShowConfigCtls%
   chkUseBits.Visible = ShowConfigCtls%
   If Not ShowConfigCtls% Then fraBits.Visible = False
   If ShowScanCtls% Then
      tmrGoLoop.Interval = 1200
   Else
      tmrGoLoop.Interval = 500
   End If

   cmdGo.Visible = (ShowInputCtls% Or ShowScanCtls%) _
   And Not (ShowConfigCtls% Or ShowBitConfigCtls%)
   picAuxIO.Visible = (Not mnUseDGrid) _
   And (ShowBitConfigCtls% Or ShowBitIOCtls%) _
   And Not (ShowScanCtls% Or ShowConfigCtls%)
   picDIO.Visible = (ShowInputCtls% Or ShowBitIOCtls% Or ShowBitConfigCtls%) _
   And Not (ShowScanCtls% Or ShowConfigCtls%)

   If ShowDIOCtls% Or ShowScanCtls% Then
      hsbPort_Change
      If (mnFuncType = DIn) _
         Or (mnFuncType = DIn32) _
         Or (mnFuncType = DInArray) _
         Then picDIO.Cls
   Else
      If mnFuncType = DBitIn Then ResetDGrid False
      If mnLibType = UNILIB Then RefreshDGrid
   End If
   For i% = 0 To 6
      Line1(i%).Visible = ShowBitIOCtls% Or ShowBitConfigCtls%
   Next i%
   For i% = 0 To Line2.Count - 1
      Line2(i%).Visible = ShowBitIOCtls% Or ShowBitConfigCtls%
   Next i%
   If mnFormType = DIGITAL_OUT Then
      If (Not gnScriptRun) Then
         Samples& = Val(Me.txtCount.Text)
         If (mnFuncType = DOutScan) _
            And (mlGenPoints < Samples&) _
            Then mbDataGenRequired = True
      End If
      If (mnFuncType = DOutArray) And (mlGenPoints < 1) Then
         If Not frmBuildData.Visible Then
            mbDataGenRequired = True
            'mnuSelData_Click
         End If
      End If
   End If
   
End Sub

Private Sub DisplayData()

   PlotSelected% = mnuPlotType(0).Checked _
      Or mnuPlotType(1).Checked _
      Or mnuPlotType(2).Checked
   If Not (mnPlot And PlotSelected%) Then Exit Sub   '(tmrCheckStatus.Enabled And Not mnuContPlot.Checked) Or (
   If mnUse32 Then
      SetBufferType 2
   Else
      SetBufferType 1
   End If
   SetPlotType mnPlotType + mnHardCopy, Me
   If mnDetailsNeedUpdate Then
      BoardName = mnuBoard(mnBoardIndex).Caption
      'If mnFuncType = DIN_SCAN Then FuncType% = DInScan
      'If (mnFormType = DIGITAL_OUT) Then FuncType% = DOutScan
      If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
         RateVal = mfRateReturned / 1000 & "Hz"
      Else
         If mfRateReturned < 1000 Then
            RateVal = mfRateReturned & "Hz"
         Else
            RateVal = mfRateReturned / 1000 & "kHz"
         End If
      End If
      SetDetails mnFuncType, mnPortNum, mlCount, _
         mfRateReturned, RateVal, A5, BoardName, A6
      mnDetailsNeedUpdate = False
   End If
   If mbEvalEnabled Then SetRate mfRateReturned
   Select Case mnFuncType
      Case DInScan, DOutScan
         If mnuContPlot.Checked Then
            LongHandle& = mlHandle
            PlotPartialBuf LongHandle&, (mlCount), mlFirstPoint, mnLastChan - mnFirstChan
            If mnSetActive Then
               SetContActive True
               mnSetActive = False
            End If
         Else
            LongHandle& = mlHandle
            PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
         End If
   End Select

End Sub

Private Function SetMsgBitConfig(PortIndex As Integer, BitIndex As Integer, BitConfig As Long) As Integer

   SetMsgBitConfig = False
   CompMsg$ = "DIO"
   MsgConfigDir$ = "IN"
   If BitConfig = DIGITALOUT Then MsgConfigDir$ = "OUT"
   PortBitString$ = "{" & Format(PortIndex, "0") & "/" & Format(BitIndex, "0") & "}:DIR="
   BitCfgMsg$ = CompMsg$ & PortBitString$ & MsgConfigDir$
   MsgResult$ = MsgLibrary.SendMessage(BitCfgMsg$)
   If SaveMsg(Me, "SendMessage(" & BitCfgMsg$ & ")", MsgResult$) Then Exit Function
   SetMsgBitConfig = True

End Function

Private Function SetMsgBitVal(ByVal PortNum As Integer, _
ByVal BitNum As Integer, ByVal BitValue As Integer) As Integer

   CompMsg$ = "DIO"
   PortBitString$ = "{" & Format(PortNum, "0") & "/" & Format(BitNum, "0") & "}:VALUE="
   BitValString$ = Format(BitValue, "0")
   BitValMsg$ = CompMsg$ & PortBitString$ & BitValString$
   MsgResult$ = MsgLibrary.SendMessage(BitValMsg$)
   If SaveMsg(Me, "SendMessage(" & BitValMsg$ & ")", MsgResult$) Then Exit Function
   BitValue% = GetMsgBitVal(PortNum%, BitNum%)
   SetMsgBitVal = BitValue%
   
End Function

Private Sub DoMsgBitIn()

   RefreshDGrid
   Exit Sub
   For PortIndex% = 0 To mnLastPort - 1
      LastBit% = GetNumMsgBits(PortIndex%)
      PortConfig& = GetCfgMsgBits(PortIndex%, LastBit%)
      For BitIndex% = 0 To LastBit% - 1
         BitConfiguration% = ((PortConfig& And 2 ^ BitIndex%) = 2 ^ BitIndex%)
         PortNum% = DGrid(BitIndex, PortIndex).ULBitPort
         BitNum% = DGrid(BitIndex, PortIndex).ULBitNum
         BitValue% = GetMsgBitVal(PortNum%, BitNum%)
         If BitConfiguration% Then
            ColorVal& = INOFFCOLOR&
            If BitValue% Then ColorVal& = INONCOLOR&
         Else
            ColorVal& = OUTOFFCOLOR&
            If BitValue% Then ColorVal& = OUTONCOLOR&
         End If
         'to do - this is only good for painting auxport
         Bit% = BitIndex%
         If BitNum% > 7 Then
            Col% = 1
            Bit% = BitIndex% - 8
         End If
         If mnLibType = MSGLIB Then Col% = PortIndex%
         PaintAuxGrid Col%, Bit%, ColorVal&
      Next BitIndex%
   Next PortIndex%
   
End Sub

Private Sub DoBitIn()

   RefreshDGrid
   Exit Sub
   PortNum% = AUXPORT
   LastBit% = 15
   PortIndex% = 0

   If Not WIN32APP Then
      'turn off error reporting, since it is valid to read outputs in 32 bit library but not 16
      ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
      If SaveFunc(Me, ErrHandling, ULStat, DONTPRINT, DONTSTOP, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
   If DGrid(0, PortIndex%).PortNum = PortNum% Then
   'If mlPortType(PortIndex%, mnBoardIndex) = PortNum% Then
      'added this 9/20/07, since config can change remotely using web devices
      InfoType% = DIGITALINFO
      ConfigItem% = DIINMASK
      ULStat = GetConfig520(InfoType%, mnBoardNum, PortIndex%, ConfigItem%, InConfig&)
      Res% = SaveFunc(Me, GetConfig, ULStat, InfoType%, mnBoardNum, PortIndex%, ConfigItem%, InConfig&, A6, A7, A8, A9, A10, A11, 0)
      mlConfigIn(PortIndex%, mnBoardIndex) = InConfig&
      ConfigItem% = DIOUTMASK
      ULStat = GetConfig520(InfoType%, mnBoardNum, PortIndex%, ConfigItem%, OutConfig&)
      Res% = SaveFunc(Me, GetConfig, ULStat, InfoType%, mnBoardNum, PortIndex%, ConfigItem%, InConfig&, A6, A7, A8, A9, A10, A11, 0)
      mlConfigOut(PortIndex%, mnBoardIndex) = OutConfig&
      
      For BitNum% = 0 To LastBit%
         InBit% = (mlConfigIn(PortIndex%, mnBoardIndex) And 2 ^ BitNum%) = 2 ^ BitNum%
         OutBit% = (mlConfigOut(PortIndex%, mnBoardIndex) And 2 ^ BitNum%) = 2 ^ BitNum%
         If BitNum% = 8 Then
            GridBit& = 0
            GridPortOffset& = GridPortOffset& + 1
         ElseIf BitNum% > 0 Then
            GridBit& = GridBit& + 1
         End If
         If DGrid(GridBit&, PortIndex% + GridPortOffset&).PortNum = PortNum% Then
         'If Not ((mnAuxBit(BitNum%, mnBoardIndex) = INVALIDBIT) Or _
         (mnAuxBit(BitNum%, mnBoardIndex) = 0) Or mnAuxBitDisabled(BitNum%, mnBoardIndex)) Then
            AuxExits% = True
            BitValue% = GetMsgBitVal(PortNum%, BitNum%)
            ULStat = cbDBitIn(mnBoardNum, PortNum%, BitNum%, BitValue%)
            If (Not gnScriptSave) Or (Not AlreadySaved%) Then
               If SaveFunc(Me, DBitIn, ULStat, mnBoardNum, PortNum%, BitNum%, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               AlreadySaved% = True
            End If
            'PortConf% = mnAuxBit(BitNum%, mnBoardIndex)
            If OutBit% And InBit% Then
               PortConf% = NOTCONFIGURED
            Else
               If InBit% Then
                  PortConf% = DIGITALIN
               Else
                  PortConf% = DIGITALOUT
               End If
            End If
            Select Case PortConf%
               Case 0
                  'used to fill in AUXPORT grid
                  'if no AUXPORT
                  ColorVal& = INVALIDCOLOR&
               Case DIGITALOUT
                  ColorVal& = OUTOFFCOLOR&
                  If BitValue% Then ColorVal& = OUTONCOLOR&
               Case DIGITALIN
                  ColorVal& = INOFFCOLOR&
                  If BitValue% Then ColorVal& = INONCOLOR&
               Case NOTCONFIGURED
                  ColorVal& = NOCONFCOLOR&
                  If BitValue% Then ColorVal& = ONNOCONFCOLOR&
               Case INVALIDBIT
                  ColorVal& = INVALIDCOLOR&
               Case MIXEDPORT
                  BitConf% = GetBitStatus(PortIndex%, BitNum, mnBoardIndex)
                  Select Case BitConf%
                     Case 0
                        'used to fill in AUXPORT grid
                        'if no AUXPORT
                        ColorVal& = INVALIDCOLOR&
                     Case DIGITALOUT
                        ColorVal& = OUTOFFCOLOR&
                        If BitValue% Then ColorVal& = OUTONCOLOR&
                     Case DIGITALIN
                        ColorVal& = INOFFCOLOR&
                        If BitValue% Then ColorVal& = INONCOLOR&
                     Case NOTCONFIGURED
                        ColorVal& = NOCONFCOLOR&
                        If BitValue% Then ColorVal& = ONNOCONFCOLOR&
                     Case INVALIDBIT
                        ColorVal& = INVALIDCOLOR&
                     Case Else
                        ColorVal& = INVALIDCOLOR&
                  End Select
               Case Else
                  ColorVal& = INVALIDCOLOR&
            End Select
            Bit% = BitNum%
            If BitNum% > 7 Then
               Col% = 1
               Bit% = BitNum% - 8
            End If
            'GridBit&, PortIndex% + GridPortOffset&
            PaintAuxGrid Col%, Bit%, ColorVal&
         End If
      Next BitNum%
   End If

   PortNum% = FIRSTPORTA
   LastBit% = 191
   If AuxExits% Then PortIndex% = 1
   For BitNum% = 0 To LastBit%
      If Not ((mnDIOBit(BitNum%, mnBoardIndex) = INVALIDBIT) Or (mnDIOBit(BitNum%, mnBoardIndex) = 0)) Then
         ULStat = cbDBitIn(mnBoardNum, PortNum%, BitNum%, BitValue%)
         If ULStat = WRONGDIGCONFIG Then
            mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
            'For d% = 0 To 7
            '   picDIO.Line (PortIndex% - 1, d% + .23)-(PortIndex% - .15, d% + 1), &HC0C0C0, BF
            'Next d%
            'Exit For
         Else
            If (Not gnScriptSave) Or (Not AlreadySaved%) Then
               If SaveFunc(Me, DBitIn, ULStat, mnBoardNum, PortNum%, BitNum%, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               AlreadySaved% = True
            End If
            PortConf% = mnDIOBit(BitNum%, mnBoardIndex)
            Select Case PortConf%
               Case 0
                  'used to fill in AUXPORT grid
                  'if no AUXPORT
                  ColorVal& = INVALIDCOLOR&
               Case DIGITALOUT
                  ColorVal& = OUTOFFCOLOR&
                  If BitValue% Then ColorVal& = OUTONCOLOR&
               Case DIGITALIN
                  ColorVal& = INOFFCOLOR&
                  If BitValue% Then ColorVal& = INONCOLOR&
               Case NOTCONFIGURED
                  ColorVal& = NOCONFCOLOR&
                  If BitValue% Then ColorVal& = ONNOCONFCOLOR&
               Case INVALIDBIT
                  ColorVal& = INVALIDCOLOR&
               Case INVALIDPORT
                  ColorVal& = INVALIDCOLOR&
               Case MIXEDPORT
                  'currently only AUXPORTS can be mixed
                  ColorVal& = INVALIDCOLOR&
               Case Else
                  ColorVal& = INVALIDCOLOR&
            End Select

            BitConf% = GetBitStatus(PortIndex%, BitLoc%, mnBoardIndex)
            If BitConf% = INVALIDBIT Then
               BitLoc% = 0
               PortIndex% = PortIndex% + 1
            End If
            If AuxExits% Then
               Col% = PortIndex% - 1
            Else
               'If Col% = 0 And BitNum% > 7 Then
               '   If BitNum% > 15 Then
               '      ColOffset% = PortIndex% + 2
               '   Else
               '      ColOffset% = PortIndex% + 1
               '   End If
               'End If
               Col% = PortIndex% '+ ColOffset%
            End If
            PaintDIOGrid Col%, BitLoc%, ColorVal&
            BitLoc% = BitLoc% + 1
         End If
      End If
   Next BitNum%
   If Not WIN32APP Then
      ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
      If SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, gnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If

End Sub

Private Sub ScriptDIn()

End Sub

Private Sub DoMsgDOut()

   txtOutVal.ForeColor = &H80000008
   DataValue& = Val(txtOutVal.Text)
   PortIndex% = hsbPort.value - 1
   ValueSet& = SetMsgPortVal(PortIndex%, DataValue&)
   If Not (ValueSet& = DataValue&) Then txtOutVal.ForeColor = &HFF
   
End Sub

Private Sub DoMsgDIn()

   PortIndex% = 0
   picDIO.ForeColor = INOFFCOLOR&
   picDIO.Cls

   For PortIndex% = 0 To mnLastPort - 1
      PortValue& = GetMsgPortVal(PortIndex%)
      If PicLine% > 7 Then
         Col% = Col% + 8
         PicLine% = 0
      End If
      picDIO.CurrentX = Col%
      picDIO.CurrentY = 8 - PicLine%
      picDIO.Print PortValue&
      PicLine% = PicLine% + 1
   Next PortIndex%
   
End Sub

Private Sub DoDIn()

   PortIndex% = 0
   picDIO.ForeColor = INOFFCOLOR&
   picDIO.Cls

   For PortIndex% = 0 To 33
      PortSelected% = DGrid(0, PortIndex).PortNum
      If PortSelected% > PortNum& Then
         PortNum& = PortSelected%
         If Not DGrid(0, PortIndex%).BitConf = INVALIDBIT Then
            'If Not (mnDIOPort(PortIndex%, mnBoardIndex) = INVALIDPORT) Then
            If mnFuncType = DInArray Then
               If FirstPort& = 0 Then
                  FirstPort& = PortNum&
                  ArrayIndex% = 0
               ElseIf PortNum& > FirstPort& Then
                  If Not (LastPort& = PortNum&) Then
                  LastPort& = PortNum&
                  ArrayIndex% = ArrayIndex% + 1
                  End If
               End If
            ElseIf mnFuncType = DIn Then
               ULStat = cbDIn(mnBoardNum, PortNum&, DataValue%)
               If ULStat = WRONGDIGCONFIG Then
                  mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
                  Exit For
               End If
               If SaveFunc(Me, DIn, ULStat, mnBoardNum, PortNum&, DataValue%, _
                  A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               If PicLine% > 7 Then
                  Col% = Col% + 8
                  PicLine% = 0
               End If
               picDIO.CurrentX = Col%
               picDIO.CurrentY = 8 - PicLine%
               LongVal& = IntValToULong(DataValue%)
               picDIO.Print LongVal&
               PicLine% = PicLine% + 1
            Else
               ULStat = cbDIn32(mnBoardNum, PortNum&, lDataValue&)
               If ULStat = WRONGDIGCONFIG Then
                  mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
                  Exit For
               End If
               If SaveFunc(Me, DIn32, ULStat, mnBoardNum, PortNum&, _
                  lDataValue&, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
                  Then Exit Sub
               If PicLine% > 7 Then
                  Col% = Col% + 8
                  PicLine% = 0
               End If
               picDIO.CurrentX = Col%
               picDIO.CurrentY = 8 - PicLine%
               'LongVal& = IntValToULong(DataValue%)
               If mbHexText Then
                  ValText$ = "0x" & Hex(lDataValue&)
               Else
                  DWordVal# = LongToUDWordVal(lDataValue&)
                  ValText$ = Format(DWordVal#, "0")
               End If
               picDIO.Print ValText$  'lDataValue&
               PicLine% = PicLine% + 1
            End If
         End If
      End If
   Next PortIndex%
   If mnFuncType = DInArray Then
      ReDim DataArray(ArrayIndex%) As Long
      ULStat = cbDInArray(mnBoardNum, FirstPort&, LastPort&, DataArray(0))
      If ULStat = WRONGDIGCONFIG Then
         mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
      End If
      If SaveFunc(Me, DInArray, ULStat, mnBoardNum, FirstPort&, LastPort&, _
         DataArray(0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      For ArrayItem% = 0 To ArrayIndex%
         If PicLine% > 7 Then
            Col% = Col% + 8
            PicLine% = 0
         End If
         picDIO.CurrentX = Col%
         picDIO.CurrentY = 8 - PicLine%
         'IntValToULong(
         LongVal& = DataArray(ArrayItem%)
         picDIO.Print LongVal&
         'PicLine% = PicLine% + 1
         Col% = Col% + 8
      Next
   End If
   
End Sub

Private Sub DoDInScan()

   PlotThis% = mnPlot
   mnPlot = False
   UseLibrary% = mnuULBuf.Checked
   If mlHandle Then
      If UseLibrary% Then
         ULStat = cbWinBufFree(mlHandle)
         If SaveFunc(Me, WinBufFree, ULStat, mlHandle, A2, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
            mnCancel = True
            Exit Sub
         End If
         mlHandle = 0: mnAcqDataType = vbEmpty
         LongHandle& = mlHandle: mlAcqPoints = 0
         PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
      Else
         If FreeMemory(mlHandle) Then
            mlHandle = 0: mlAcqPoints = 0
            mnAcqDataType = vbEmpty
         End If
      End If
   End If
                                              
   If UseLibrary% Then
      If mnUse32 Then
         mlHandle = BufAlloc32(Me, mlCount)
         'mlHandle = cbWinBufAlloc32(mlCount)
         If SaveFunc(Me, WinBufAlloc32, mlHandle, mlCount, A2, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mnAcqDataType = vbLong
         mlAcqPoints = mlCount
      Else
         mlHandle = cbWinBufAlloc(mlCount)
         If SaveFunc(Me, WinBufAlloc, mlHandle, mlCount, A2, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mnAcqDataType = vbInteger
         mlAcqPoints = mlCount
      End If
   Else
      mlHandle = AllocateMemory(mlCount)
      mnAcqDataType = vbInteger: mlAcqPoints = mlCount
   End If
   'bug #567, temporary line follows:
   If mlHandle = 17 Then Stop
   mlStatusType = DIFUNCTION

   CBRate& = mfRate
   ULStat = cbDInScan(mnBoardNum, mnPortNum, _
      mlCount, CBRate&, mlHandle, mvOptions)
   mfRateReturned = CBRate&
   If SaveFunc(Me, DInScan, ULStat, mnBoardNum, mnPortNum, _
      mlCount, CBRate&, mlHandle, mvOptions, A7, A8, A9, A10, A11, 0) Then
      ULStat = StopBackground520(mnBoardNum, mlStatusType)
      mnCancel = True
      If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnFormType, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
   'mfRateReturned = mfRate

   ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
   If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, _
      mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then
      mnCancel = True
      Exit Sub
   End If
   If Status% Then
      tmrCheckStatus.ENABLED = True
      cmdStop.Visible = True
      cmdStop.ENABLED = True
      mnPlot = False
   Else
      mnPlot = PlotThis%
   End If

End Sub

Private Sub DoDOut()

   If mnFuncType = DOutArray Then
      If mlGenPoints = 0 Then
         mlGenPoints = 2
         mbDataGenRequired = True
      End If
      If mbDataGenRequired Then
         mnuSelData_Click
         Exit Sub
      End If
      ArrayIndex% = mlGenPoints - 1
      ReDim DataArray(0, ArrayIndex%) As Long
      ReDim GenData(0, ArrayIndex%) As Integer
      ULStat& = LoadArrayFromWinBuf(Me, mlGenHandle, GenData(), 0, mlGenPoints)
      Samples& = IntArrayToLong(GenData(), DataArray(), ArrayIndex%, 0)
      FirstPort& = mnPortNum
      LastPort& = mnPortNum + ArrayIndex%
   ElseIf mnFuncType = DOut Then
      DigitalValue$ = txtOutVal.Text
      If mnPortNum > 0 Then
         If mnLibType = MSGLIB Then
            CompMsg$ = "DIO"
            PortString$ = "{" & Format(mnPortIndex, "0") & "}:VALUE=" & DigitalValue$
            PortValMsg$ = CompMsg$ & PortString$
            MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
            If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Sub
         Else
            If mbHexText Then
               'ValText$ = "0x" & Hex(lDataValue&)
               lDataValue& = Val("&H" & DigitalValue$)
            Else
               DWordVal# = Val(DigitalValue$)
               lDataValue& = UDWordValToLong(DWordVal#)
            End If
            ULStat = cbDOut32(mnBoardNum, mnPortNum, lDataValue&)
            x% = SaveFunc(Me, DOut32, ULStat, mnBoardNum, mnPortNum, _
               lDataValue&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         End If
      End If
   Else
      DigitalValue$ = txtOutVal.Text
      'lDataValue& = UDWordValToLong(DWordVal#)
      If (chkAllPorts.value = 1) Then
         LastIndex& = 33
      Else
         LastIndex& = 0
      End If
      If mbHexText Then
         'ValText$ = "0x" & Hex(lDataValue&)
         If Len(DigitalValue$) > 4 Then
            LowVal$ = Right(DigitalValue$, 4)
            HighVal$ = Left(DigitalValue$, Len(DigitalValue$) - 4)
            DataValLow& = IntValToULong(Val("&H" & LowVal$))
            DataValHigh& = (Val("&H" & HighVal$)) + (DataValHigh& And 32768)
            'DataValSign& = DataValHigh& And 32768
            lDataValue& = (DataValHigh& * 65536) + DataValLow&
            'lDataValue& = UDWordValToLong(DWordVal#)
         Else
            DataValue% = Val("&H" & DigitalValue$)
            lDataValue& = IntValToULong(DataValue%)
         End If
      Else
         DWordVal# = Val(DigitalValue$)
         lDataValue& = UDWordValToLong(DWordVal#)
      End If
      For PortIndex& = FirstIndex& To LastIndex&
         If (chkAllPorts.value = 1) Then
            PortNum& = DGrid(0, PortIndex&).PortNum
         Else
            PortNum& = mnPortNum
         End If
         If PortNum& > 0 Then
            ULStat = cbDOut32(mnBoardNum, PortNum&, lDataValue&)
            x% = SaveFunc(Me, DOut32, ULStat, mnBoardNum, PortNum&, _
               lDataValue, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = WRONGDIGCONFIG Then Exit For
         End If
      Next
   End If
   If mnFuncType = DOutArray Then
      'ReDim DataArray(ArrayIndex%) As Long
      ULStat = cbDOutArray(mnBoardNum, FirstPort&, LastPort&, DataArray(0, 0))
      If SaveFunc(Me, DOutArray, ULStat, mnBoardNum, FirstPort&, LastPort&, _
         DataArray(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
   
End Sub

Private Sub DoDOutScan()

   If mlGenHandle = 0 Then
      MsgBox "No data is available for output. Use the 'Generate Data' option " & _
      "before calling 'DOutScan'.", vbOKOnly, "No Output Data Available"
      Exit Sub
   End If
   mlStatusType = DOFUNCTION
   
   CBRate& = mfRate
   ULStat = cbDOutScan(mnBoardNum, mnPortNum, mlCount, CBRate&, mlGenHandle, mvOptions)
   mfRateReturned = CBRate&
   If SaveFunc(Me, DOutScan, ULStat, mnBoardNum, mnPortNum, mlCount, _
      CBRate&, mlGenHandle, mvOptions, A7, A8, A9, A10, A11, 0) Then
      ULStat = StopBackground520(mnBoardNum, mlStatusType)
      If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnFormType, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Exit Sub
   End If
   'mfRateReturned = mfRate

   ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
   If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, _
      CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   If Status% Then
      tmrCheckStatus.ENABLED = True
      cmdStop.Visible = True
      cmdStop.ENABLED = True
      mnPlot = False
   End If

End Sub

Private Sub Form_Activate()
   
   UpdateMainStatus
   If Not (mnLoading Or mnUnloading) Then mnRefreshProps = True
   mnLoading = False

End Sub

Private Sub Form_DblClick()

   UpdateStatBar True

End Sub

Private Sub Form_Initialize()
   
   mnLoading = True
   Me.mnuLibrary(MSGLIB).ENABLED = frmMain.mnuLibrary(MSGLIB).ENABLED
   Me.mnuLibrary(UNILIB).ENABLED = frmMain.mnuLibrary(UNILIB).ENABLED

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   If (Shift And 4) = 4 Then
      Select Case KeyCode
         Case Asc("I")
            mfmUniTest.cmdFormType(0) = True
         Case Asc("O")
            mfmUniTest.cmdFormType(1) = True
         Case Asc("N")
            mfmUniTest.cmdFormType(2) = True
         Case Asc("T")
            mfmUniTest.cmdFormType(3) = True
         Case Asc("C")
            mfmUniTest.cmdFormType(4) = True
         Case Asc("M")
            mfmUniTest.cmdFormType(5) = True
         Case Asc("'")
            mfmUniTest.cmdFormType(6) = True
         Case Asc("L")
            mfmUniTest.cmdUtils = True
      End Select
   Else
      Select Case KeyCode
         Case 118 'F7 - set default child size
            Me.Height = 2300
            Me.Width = 7200
         Case 120 'F9 - set height to 1/3 screen
            mfmUniTest.Height = Screen.Height / 3
         Case 122 'F11 - set to screen bottom
            mfmUniTest.Move 0, Screen.Height - mfmUniTest.Height, Screen.Width
         Case 123 'F12 - set to screen top
            mfmUniTest.Move 0, 0, Screen.Width
      End Select
   End If

End Sub

Private Sub Form_Load()
 
mnUseDGrid = True
If mnUseDGrid Then
   picDIO.Left = picAuxIO.Left + 100
   picDIO.Width = 5240
   picDIO.ScaleWidth = 34
   mnLastColumn = 35
   Load Line2(31)
   Load Line2(32)
   Line2(31).X1 = 32
   Line2(31).X2 = 32
   Line2(32).X1 = 33
   Line2(32).X2 = 33
   Line2(31).Visible = True
   Line2(32).Visible = True
   For x% = 0 To 32: Line2(x%).X1 = x% + 1: Line2(x%).X2 = x% + 1: Next
End If

#If NETOPS Then
   If gnThreading Then
      'Set AcqThread = New AcqThread.AcqThread
      'AcqThread.Init
      mnThreading = True
   End If
#End If
#If MSGOPS Then
   If gnLibType = MSGLIB Then
      Set MsgLibrary = CreateObject("MBDClass.MBDComClass")
   End If
#End If

   mnLibType = gnLibType
   mnNumBoards = gnNumBoards
   mnuLibrary(gnLibType).Checked = True
   If Not gbULLoaded Then
      Me.mnuULBuf.Checked = False
      mnuLibrary(UNILIB).Checked = False
      mnuLibrary(UNILIB).ENABLED = False
   End If
   If Not gbULLoaded Then Me.mnuULBuf.Checked = False
   If (gnLibType = UNILIB) And Not gbULLoaded Then
      mnuFunc.ENABLED = False
      cmdGo.ENABLED = False
      cmdConfigOut.ENABLED = False
      cmdConfigIn.ENABLED = False
      cmdSend.ENABLED = False
      cmdRead.ENABLED = False
      txtOutVal.ENABLED = False
      chkUseBits.ENABLED = False
      mnuEvents.ENABLED = False
   End If
   'to do - get rid of this when DGrid is fully implemented
   ReDim mnDIOPort(33, gnNumBoards)
   ReDim mnDIOBit(192, gnNumBoards)
   'ReDim mnAuxPort(1, gnNumBoards)
   ReDim mnAuxBit(16, gnNumBoards)
   ReDim mlConfigIn(33, gnNumBoards)
   ReDim mlConfigOut(33, gnNumBoards)
   ReDim mlPortType(33, gnNumBoards)
   ReDim mnAuxBitDisabled(16, gnNumBoards)
   'ReDim mnDIOBitEnable(196, gnNumBoards)
   
   masConfig(DIGITALOUT) = "DIGITALOUT"         '1
   masConfig(DIGITALIN) = "DIGITALIN"           '2
   masConfig(NOTCONFIGURED) = "NOTCONFIGURABLE"   '3
   masConfig(INVALIDBIT) = "INVALIDBIT"         '4
   masConfig(INVALIDPORT) = "INVALIDPORT"       '5
   masConfig(MIXEDPORT) = "MIXEDPORT"           '6
   
   For Row% = 0 To 7
      For Col% = 0 To 34
         DGrid(Row%, Col%).BitConf = INVALIDBIT
      Next Col%
   Next Row%

   For i% = 0 To gnNumBoards - 1
      If BoardNum% < 0 Then BoardNum% = Abs(BoardNum% + 1)
      If gnLibType = MSGLIB Then
         BoardNum% = i%
         BoardName$ = GetNameOfMsgBoard(BoardNum%)
         SplitName = Split(BoardName$, "::")
         DisplayName$ = SplitName(0)
      Else
         BoardNum% = gnBoardEnum(i%)
         BoardName$ = GetNameOfBoard(BoardNum%)
         DisplayName$ = BoardName$
      End If
      CurrentName$ = BoardNum% & ") " & DisplayName$
      mnDIOPort(33, i%) = INVALIDPORT
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      Else
         mnBoardNum = BoardNum%
         mnBoardIndex = 0
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   hsbOutVal.Top = picAuxIO.Top + 280
   txtOutVal.Top = picAuxIO.Top + 280
   lblOutVal.Top = picAuxIO.Top
   fraBits.Top = 200 'hsbOutVal.Top -
   lblSelectPort.Top = picAuxIO.Top
   hsbPort.Top = picAuxIO.Top + 280
   lblPort.Top = picAuxIO.Top + 280
   cmdSend.Top = lblPort.Top + 400
   cmdRead.Top = cmdSend.Top
   chkUseBits.Top = cmdRead.Top + 100
   chkScrollUpdate.Top = lblPort.Top + 400
   chkAllPorts.Top = lblSelectPort.Top ' + 400
   cmdConfigOut.Top = hsbPort.Top + 400
   cmdConfigIn.Top = hsbPort.Top + 400
   mnPlot = True
   mnPortNum = AUXPORT
   
   tmrGoLoop.Interval = 500
   tmrDIn.Interval = 500
   mnNumSelPorts = -1
   mnNumSelBits = -1
   If gnUniScript Then
      Me.HelpContextID = 30000
   Else
      If gnThreading Then
         Me.HelpContextID = 20000
      Else
         Me.HelpContextID = 10000
      End If
   End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   mnUnloading = True
   Samples& = 0
   UseWinAPI% = Not Me.mnuULBuf.Checked
   DataHandle& = GenerateData(DataType&, Cycles%, Samples&, Chans%, _
      mvAmplitude, mvOffset, Signal%, NewData%, Channel%, mlFirstPoint, UseWinAPI%)
   SetupData DataHandle&, Samples&, Chans%, DataType&
   If Not mlHandle = 0 Then
      UseLibrary% = mnuULBuf.Checked
      If UseLibrary% Then
         ULStat = cbWinBufFree(mlHandle)
         x% = SaveFunc(Me, WinBufFree, ULStat, mlHandle, A2, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         mlHandle = 0: mnAcqDataType = vbEmpty
         LongHandle& = mlHandle: mlAcqPoints = 0
         PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
      Else
         If FreeMemory(mlHandle) Then
            mlHandle = 0: mlAcqPoints = 0
            mnAcqDataType = vbEmpty
         End If
      End If
   End If
   If Not mlGenHandle = 0 Then
      ULStat = cbWinBufFree(mlGenHandle)
      x% = SaveFunc(Me, WinBufFree, ULStat, mlGenHandle, A2, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      mlGenHandle = 0: mnGenDataType = vbEmpty
      mlGenPoints = 0
      If x% Then Exit Sub
   End If

End Sub

Private Sub Form_Resize()

   lblStatus.Width = ScaleWidth
   lblStatus.Top = ScaleHeight - lblStatus.Height
   If gnInitializing Then
      'if the form is just loading, sets the form type
      'and sets the default function (cbDConfigPort())
      mnFormType = (Val("&H" & Tag) And &HF00&) / &H100
      mnThisInstance = Val("&H" & Tag) And &HFF
      'mnuFuncArray_Click (DCONFIG)
      If Not gnScriptRun Then
         mfRate = 1000
         mlCount = 1000
         mfRateReturned = mfRate
      End If
      msTitle = Caption
      'mnuBoard_Click (0)
      gnInitializing = False
      
      Select Case mnLibType
         Case UNILIB
            GetPortConfig
         Case NETLIB
         Case MSGLIB
            'NumProps% = GetDIOProps(msBoardName, MsgLibrary, PropsList)
            'UpdateFormProps PropsList, NumProps%
      End Select
   End If
   mlStatusType = DIFUNCTION
   If mnFormType = DIGITAL_OUT Then mlStatusType = DOFUNCTION
   msStatusType = "DIO"
   If Me.fraScript.Visible Then
      Me.fraScript.Width = Me.ScaleWidth
      Me.fraScript.Height = Me.ScaleHeight
      Me.txtScript.Width = Me.fraScript.Width - 200
      Me.txtScript.Height = Me.fraScript.Height - 300
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   mnUnloading = True
   tmrDIn.ENABLED = False
   tmrCheckStatus.ENABLED = False
   tmrDelay.ENABLED = False
   Unload frmBuildData
   
   Select Case mnLibType
      Case UNILIB
         If mnNumBoards > 0 Then ULStat = StopBackground520(mnBoardNum, mlStatusType)
         SaveFlow% = geErrFlow
         SaveLocalErrDisp% = gnLocalErrDisp
         geErrFlow = 0  'don't bother reporting or trapping error here
         gnLocalErrDisp = False
         ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
         If SaveFunc(Me, ErrHandling, ULStat, DONTPRINT, DONTSTOP, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         If Not (mlEventType = 0) Then
            ULStat = UninstallEvent(mnBoardNum, ALL_EVENT_TYPES)
            If Not (ULStat = NOTWINDOWSFUNC) Then
               'NOTWINDOWSFUNC is locally generated in this case.
               'It indicates the cbDisableEvent was not called
               'so it should not be logged in call history.
               x% = SaveFunc(Me, DisableEvent, ULStat, mnBoardNum, _
                  ALL_EVENT_TYPES, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            End If
         End If
         geErrFlow = SaveFlow%
         gnLocalErrDisp = SaveLocalErrDisp%
         ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
         If SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, gnErrHandling, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case NETLIB
         'Set AcqThread = Nothing
         mnThreading = False
      Case MSGLIB
         StopMsg$ = msStatusType & ":STOP"
         If Not MsgLibrary Is Nothing Then
            If (mvOptions And BACKGROUND) = BACKGROUND Then
               MsgResult$ = MsgLibrary.StopBackgroundScan(msStatusType)
            Else
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
            End If
            If mnMessaging Then
               If Not (MsgLibrary.DeviceID = "") Then _
               MsgLibrary.ReleaseDevice (msBoardName)
               Set MsgLibrary = Nothing
               mnMessaging = False
            End If
         End If
   End Select
   
   UnLoadChildForm Me, mnFormType, mnThisInstance
   gnDIOForms = gnDIOForms - 1

End Sub

Private Sub GetPortConfigOrg()

   'check the library version of port configuration
   'don't report errors here
   SaveLocalErrDisp% = gnLocalErrDisp
   gnLocalErrDisp = False

   Static stnBoardNum As Integer
   GridNeedsUpdate% = False
   If stnBoardNum <> mnBoardIndex Then GridNeedsUpdate% = True
   stnBoardNum = mnBoardIndex
   For PortIndex% = 0 To 32
      If mnDIOPort(PortIndex%, mnBoardIndex) <> INVALIDPORT Then
         InfoType% = DIGITALINFO
         If mnuDIODetect.Checked Then
            GridNeedsUpdate% = True
            ConfigItem% = DIINMASK
            ULStat = GetConfig520(InfoType%, mnBoardNum, _
               DevNum%, ConfigItem%, ValConfig&)
            ConfigErr = SaveFunc(Me, GetConfig, ULStat, _
               InfoType%, mnBoardNum, DevNum%, ConfigItem%, _
               ValConfig&, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = BADCONFIGITEM Then
               mnDIOPort(PortIndex%, mnBoardIndex) = INVALIDPORT
            Else
               mlConfigIn(PortIndex%, mnBoardIndex) = ValConfig&
               ConfigItem% = DIOUTMASK
               ULStat = GetConfig520(InfoType%, mnBoardNum, _
                  DevNum%, ConfigItem%, ValConfig&)
               ConfigErr = SaveFunc(Me, GetConfig, ULStat, _
                  InfoType%, mnBoardNum, DevNum%, ConfigItem%, _
                  ValConfig&, A6, A7, A8, A9, A10, A11, 0)
               mlConfigOut(PortIndex%, mnBoardIndex) = ValConfig&
               If mlConfigIn(PortIndex%, mnBoardIndex) _
                  And mlConfigOut(PortIndex%, mnBoardIndex) Then
                  'if any bits overlap, it's a non-configurable AUXPORT
                  mnDIOPort(PortIndex%, mnBoardIndex) = NOTCONFIGURED
               ElseIf (mlConfigIn(PortIndex%, mnBoardIndex) = 0) Then
                  'if either inmask or outmask is all zero,
                  'the entire port is configured for the opposite
                  mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
               ElseIf (mlConfigOut(PortIndex%, mnBoardIndex) = 0) Then
                  mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALIN
               Else
                  'otherwise, it's a mixed bit port
                  mnDIOPort(PortIndex%, mnBoardIndex) = MIXEDPORT
               End If
            End If
            DevNum% = DevNum% + 1
         Else
            ConfigItem% = DICONFIG
            ULStat = GetConfig520(InfoType%, mnBoardNum, _
               DevNum%, ConfigItem%, ValConfig&)
            ConfigVal& = ValConfig&
            If SaveFunc(Me, GetConfig, ULStat, InfoType%, _
               mnBoardNum, DevNum%, ConfigItem%, ConfigVal&, _
               A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            DevNum% = DevNum% + 1
            If Not (mnDIOPort(PortIndex%, mnBoardIndex) = ConfigVal&) Then
               GridNeedsUpdate% = True
               Select Case ConfigVal&
                  Case 1
                     mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
                  Case 2
                     mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALIN
                  Case 3
                     mnDIOPort(PortIndex%, mnBoardIndex) = NOTCONFIGURED
               End Select
            End If
         End If
      End If
   Next
   gnLocalErrDisp = SaveLocalErrDisp%
   If GridNeedsUpdate% Then RefreshGrid

End Sub

Private Sub hsbOutVal_Change()

   hsbOutVal_Scroll
   If mnuTimer.Checked Then
      tmrGoLoop.ENABLED = True
      cmdStop.ENABLED = True
   End If

End Sub

Private Sub hsbOutVal_Scroll()

   DataVal% = hsbOutVal.value Xor &H8000
   lDataValue& = IntValToULong(DataVal%)
   If mbHexText Then
      ValText$ = Hex(lDataValue&)
   Else
      ValText$ = Format(lDataValue&, "0")
   End If
   txtOutVal.Text = ValText$
   If chkScrollUpdate And ((mnFuncType = DOut) _
      Or (mnFuncType = DOutArray)) Then
      If chkAllPorts Then
         SaveIndex% = mnPortIndex
         SavePort% = mnPortNum
         For mnPortIndex = 0 To 33
            If mnDIOPort(mnPortIndex, mnBoardIndex) <> INVALIDPORT Then
               If mnPortIndex > 0 Then
                  mnPortNum = mnPortIndex + 9
               Else
                  mnPortNum = 1
               End If
               DoDOut
            End If
         Next
         mnPortIndex = SaveIndex%
         mnPortNum = SavePort%
      Else
         If mnLibType = UNILIB Then
            DoDOut
         Else
            DoMsgDOut
         End If
      End If
      RunOtherForms DIGITAL_IN, UPDATEDIO
   End If
   
End Sub

Private Sub hsbPort_Change()

   If gnNumBoards = 0 Then Exit Sub
   
   mnPortIndex = hsbPort.value - 1
   If mnPortIndex < 0 Then mnPortIndex = 0
   If mnPortIndex > 2 Then
      mnPortNum = hsbPort.value + 6
   Else
      mnPortNum = hsbPort.value
   End If
   If mnLibType = MSGLIB Then
      CurPort& = hsbPort.value - 1
      lblPort.Caption = " DIO{" & Format(CurPort&, "0\}")
   Else
      lblPort.Caption = " " & GetPortString(mnPortNum)
   End If
   ShowConfigStat

End Sub

Private Sub mnuArrange_Click(Index As Integer)

   mfmUniTest.Arrange Index

End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash

End Sub

Private Sub mnuBoard_Click(Index As Integer)

   Dim EventConflict As Boolean
   Dim BoardChange As Boolean
   Dim eventForm As Form

   If Not (gnNumBoards > 0) Then Exit Sub
   If mnLibType = UNILIB Then
      BoardNum% = gnBoardEnum(Index)
   Else
      BoardNum% = Index
   End If
   If Index <> mnBoardIndex Then BoardChange = True
   If mlEventType > 0 Then
      If Not (mnBoardNum = BoardNum%) Then
         Dim Resp As VbMsgBoxResult
         Resp = MsgBox("Events are enabled for this device. " _
            & "Disable events and change to a different board?", _
            vbOKCancel, "Disable Events?")
         If Resp = vbCancel Then
            Exit Sub
         Else
            ULStat = cbDisableEvent(mnBoardNum, ALL_EVENT_TYPES)
            If ULStat = 0 Then
               For i% = 0 To 5
                  mnuEvent(i%).Checked = False
               Next i%
            End If
         End If
      End If
   End If
   Caption = msTitle & " board " & mnuBoard(Index).Caption
   If Not gnInitializing Then mnuBoard(mnBoardIndex).Checked = False
   Caption = msTitle & " Board " & mnuBoard(Index).Caption
   
   mnBoardNum = BoardNum%
   mnuBoard(Index).Checked = True
   mnBoardIndex = Index
   If mnLibType = MSGLIB Then
      If mnMessaging Then
         If Not (MsgLibrary.DeviceID = "") Then _
         MsgLibrary.ReleaseDevice (msBoardName)
         mnMessaging = False
      End If
      BoardName$ = GetNameOfMsgBoard(mnBoardNum)
      SplitName = Split(BoardName$, "::")
      DisplayName$ = SplitName(0)
      SetPointer% = Not (mfmUniTest.MousePointer = vbHourglass)
      If SetPointer% Then mfmUniTest.MousePointer = vbHourglass
      Me.ENABLED = False
      Me.mnuBoardSel.ENABLED = False
      DoEvents
      MBDResponse$ = MsgLibrary.CreateDevice(BoardName$)
      Me.ENABLED = True
      Me.mnuBoardSel.ENABLED = True
      If SetPointer% Then mfmUniTest.MousePointer = vbDefault
      If Not SaveMsg(Me, "CreateDevice(" & BoardName$ & ")", MBDResponse$) Then
         mnMessaging = True
         Component$ = "DIO"
         msDIOSupport = MsgLibrary.GetSupportedMessages(Component$)
      End If
   Else
      BoardName$ = GetNameOfBoard(mnBoardNum)
      DisplayName$ = BoardName$
      EventConflict = CheckActiveEvents(mnBoardNum, eventForm)
      If EventConflict Then
         Dim EventType As Long, EventData As Long
         Dim EventParam As Long, formTitle As String
         formTitle = eventForm.Caption
         Resp = MsgBox("Do you want to disable events for the for titled '" _
            & formTitle & "' and transfer those events to this form?", _
            vbYesNo, "Transfer Events to This Form?")
         If Resp = vbYes Then
            DoEvents
            eventForm.GetEvent EventType, EventData, EventParam
            eventForm.ConfigureEvent ALL_EVENT_TYPES, 0, 0
            ConfigureEvent EventType, EventData, EventParam
         End If
      End If
   End If
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = BoardName$
   lpKeyName$ = "OutputReadType"
   lpDefault$ = 1
   nSize% = 70   'this will be 0, 1, or 2 (can't read, readable, unreliable)
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   msReadType = CustomParam$
   lpKeyName$ = "DigInvertWrite"
   lpDefault$ = 0
   nSize% = 2   'this will be 0 or 1 (reads inverted = 1)
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   mbReadInvert = (CustomParam$ = "1")
   msBoardName = BoardName$
   msDisplayName = DisplayName$
   mnIsPCM = False
   PCMType$ = Left$(BoardName$, 7)
   If (PCMType$ = "PCM-DAS") Or (PCMType$ = "PCM-DAC") Then mnIsPCM = True
   
   PCMType$ = Left$(BoardName$, 11)
   If (PCMType$ = "PC-CARD-DAS") Then mnIsPCM = True
   If BoardName$ = "PC-CARD-DAC08" Then mnIsPCM = True
   
   If BoardChange Or (Not mnFormInitialized) Then
      mbDeviceInUse = False
      If mnLibType = UNILIB Then
         If mnUseDGrid Then
            ResetDGrid True
            InitDGrid
            If ((mnFuncType = DBitIn) Or (mnFuncType = DBitOut) _
               Or (mnFuncType = DConfigBit)) _
               Then RefreshDGrid
         Else
            GetPortConfig
         End If
      Else
         If msDIOSupport = "" Then
            mnLastPort = 0
         Else
            NumProps% = GetDIOProps(msBoardName, MsgLibrary, PropsList)
            UpdateFormProps PropsList, NumProps%
         End If
         InitDGrid
         If ((mnFuncType = DBitIn) Or (mnFuncType = DBitOut) _
            Or (mnFuncType = DConfigPort)) _
            Then RefreshDGrid
      End If
   End If
   UpdateMainStatus
   A1 = msDisplayName
   If gnScriptSave And (Not gnInitializing) Then
      FuncID% = SSetBoardName
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetBoardName, FuncStat, _
         A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   Bits% = GetDIOResolution()
   ConfigureValidOptions
   For TrigMenu% = 0 To mnuTrigType.Count - 1
      mnuTrigType(TrigMenu%).Checked = False
   Next TrigMenu%
   mnTrigType = TRIGPOSEDGE
   mbThresholdSet = False
   For ThisForm% = 0 To Forms.Count - 1
      FName$ = Forms(ThisForm%).Name
      If FName$ = "frmAnalog" Or _
         FName$ = "frmCounter" Or _
         FName$ = "frmDigital" Then
         If (Not (Forms(ThisForm%) Is Me)) Then
            If Forms(ThisForm%).GetTrigParameters(mnBoardNum, _
               TrigType&, LowThresh%, HighThresh%) Then
               mnTrigType = TrigType
               mnLowThreshold = LowThresh
               mnHighThreshold = HighThresh
               mbThresholdSet = True
               mnuTrigType(mnTrigType).Checked = True
               Exit For
            End If
         End If
      End If
   Next
   If Not mbThresholdSet Then mnuTrigType(mnTrigType).Checked = True

End Sub

Private Sub mnuBufInfo_Click()

   ShowBufferInfo 0
   
End Sub

Private Sub mnuClose_Click()

   Unload Me

End Sub

Private Sub mnuContPlot_Click()

   mnuContPlot.Checked = Not mnuContPlot.Checked

End Sub

Private Sub mnuDIODetect_Click()

   mnuDIODetect.Checked = Not mnuDIODetect.Checked
   ReDim mnDIOPort(33, gnNumBoards)
   
   GetPortConfig

End Sub

Private Sub mnuEvalData_Click()

   NumChans% = 1
   If mnResolution = 16 Then
      frmEvalData.txtVMin.Text = "32640"
      frmEvalData.txtVMax.Text = "32896"
      frmEvalData.txtDeltaVMax.Text = "800"
      frmEvalData.txtTriggerLevel.Text = "32768"
   Else
      frmEvalData.txtVMin.Text = "2040"
      frmEvalData.txtVMax.Text = "2056"
      frmEvalData.txtDeltaVMax.Text = "50"
      frmEvalData.txtTriggerLevel.Text = "2048"
   End If

   For i& = 0 To 18
      CurParam& = 2 ^ i&
      ParamValue = GetEvalParams(CurParam&)
      Select Case CurParam&
         Case geENABLEEVAL
         Case geEVALDELTA
            DeltaEval = ParamValue
            frmEvalData.chkMaxDelta.value = 0
            If DeltaEval Then frmEvalData.chkMaxDelta.value = 1
         Case geEVALRANGE
            RangeEval = ParamValue
            frmEvalData.chkMinMaxStop.value = 0
            If RangeEval Then frmEvalData.chkMinMaxStop.value = 1
         Case geEVALMINDELTA
            DeltaEval = ParamValue
            frmEvalData.chkMinDelta.value = 0
            If DeltaEval Then frmEvalData.chkMinDelta.value = 1
         Case geSTART
            EvalStart& = ParamValue
            frmEvalData.txtFirstPoint.Text = EvalStart&
         Case geNUMPOINTS
            ECount& = ParamValue
            If ECount& = 0 Then ECount& = mlCount
            'PerChan& = (mlAcqPoints / NumChans%)
            'If (ECount& = 0) Or (ECount& > PerChan&) Then
            '   'TCount& = Me.txtCount.Text
            '   ECount& = PerChan&   'TCount& / NumChans%
            '   InitialRun = True
            'End If
            frmEvalData.txtNumSamps.Text = ECount&
         Case geEVALCHAN
            EvalChan& = ParamValue
            frmEvalData.txtChan.Text = EvalChan&
         Case geDELTAVAL
            DeltaVal& = ParamValue
            If DeltaVal& > 0 Then frmEvalData.txtDeltaVMax.Text = DeltaVal&
         Case geMINVAL
            MinVal& = ParamValue
            frmEvalData.txtVMin.Text = MinVal&
         Case geMAXVAL
            MaxVal& = ParamValue
            frmEvalData.txtVMax.Text = MaxVal&
         Case geDELTAMIN
            DvMin& = ParamValue
            If ParamValue = 0 Then DvMin& = 1
            frmEvalData.txtDeltaVMin.Text = DvMin&
         Case geSHOWPASTEMSG
            frmEvalData.chkShowMsg.value = 0
            If ParamValue Or InitialRun Then frmEvalData.chkShowMsg.value = 1
         Case geNUMMSGSAMPS
            NumMsgSamps& = ParamValue
            If ParamValue = 0 Then NumMsgSamps& = 10
            If Not InitialRun Then frmEvalData.txtNumMsgSamps.Text = NumMsgSamps&
         Case geWRITEFILEERRS
            frmEvalData.chkSaveData.value = 0
            If ParamValue Or InitialRun Then frmEvalData.chkSaveData.value = 1
      End Select
   Next i&
   If mbEventCount Then
      frmEvalData.txtCountTrig.Text = Format(mlEvalCount, "0")
      frmEvalData.chkCountStop.value = 1
   End If
      
   frmEvalData.Show 1
   mbEval = (frmEvalData.chkEnableEval.value = 1)
   mbEventCount = (frmEvalData.chkCountStop.value = 1)
   mlEvalCount = Val(frmEvalData.txtCountTrig.Text)
   If mbEval Then mlEvalCount = frmEvalData.txtNumSamps.Text
   mbEvalEnabled = mbEval
   Unload frmEvalData
   If gnScriptSave Then
      ThisFunction& = SEvalEnable
      FuncStat = 0
      A1 = mbEval
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, ThisFunction&, _
            FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   If mbEval Then mnDetailsNeedUpdate = True

End Sub

Private Sub mnuEvent_Click(Index As Integer)

   Dim EventSize As Long, EventType As Long
   Dim EventParam As Long
   Dim EnablingEvent As Boolean
   Dim MsgEvent As Boolean
   
   Set mfrmThisForm = Me
   mnuEvent(Index).Checked = Not mnuEvent(Index).Checked
   EnablingEvent = False
   lblStatus.BackColor = &HFFFFFF
   MsgEvent = (mnLibType = MSGLIB)
   FormCaption$ = Me.Caption
   If Index = 7 Then
      EventType = ALL_EVENT_TYPES
      If MsgEvent Then
         MsgResult$ = MsgLibrary.DisableEvents(0)
         If SaveMsg(Me, "DisableEvents(0)", MsgResult$) Then Exit Sub
         MsgResult$ = MsgLibrary.DisableEvents(1)
         If SaveMsg(Me, "DisableEvents(1)", MsgResult$) Then Exit Sub
         MsgResult$ = MsgLibrary.DisableEvents(2)
         If SaveMsg(Me, "DisableEvents(2)", MsgResult$) Then Exit Sub
      Else
         ULStat = UninstallEvent(mnBoardNum, EventType)
         x% = SaveFunc(Me, DisableEvent, ULStat, mnBoardNum, EventType, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      End If
      If ULStat = 0 Then
         For i% = 0 To 5
            mnuEvent(i%).Checked = False
         Next i%
      End If
      mlEventType = mlEventType Xor EventType
      If EventType = 255 Then mlEventType = 0
      Exit Sub
   End If
   
   EventType = Choose(Index + 1, ON_SCAN_ERROR, ON_EXTERNAL_INTERRUPT, _
   ON_PRETRIGGER, ON_DATA_AVAILABLE, ON_END_OF_INPUT_SCAN, ON_END_OF_OUTPUT_SCAN, _
   ON_CHANGE_DI)
   If MsgEvent Then
      MsgEventType& = Switch(EventType = ON_DATA_AVAILABLE, 1, _
      EventType = ON_END_OF_INPUT_SCAN, 2, EventType = ON_SCAN_ERROR, 4)
   End If
   EnablingEvent = mnuEvent(Index).Checked
   
   If EventType = ON_DATA_AVAILABLE Then
      If Not gnScriptRun Then
         TempVar = InputBox("Enter number of points (or less than 1 to disable)", _
         "Event Size", Format(mlEventSize, "0"))
         If Len(TempVar) Then EventParam = Val(TempVar)
      End If
      If Not (EventParam > 0) Then
         EnablingEvent = False
         mnuEvent(Index).Checked = False
      Else
         mlEventSize = EventParam
         EnablingEvent = True
         mnuEvent(Index).Checked = True
      End If
   End If

   If EventType = ON_EXTERNAL_INTERRUPT Then
      If Not gnScriptRun Then
         TempVar = InputBox("Enter event parameter (0=NONE, " _
            & "1=LATCH_DI, 2=LATCH_DO, 4=SAVE_CHIP1_CTRS, 8=SAVE_CHIP2_CTRS)", _
            "Event Parameter", Format(mlEventParam, "0"))
         If Len(TempVar) Then EventParam = TempVar
      End If
      mlEventParam = EventParam
   End If
   
   If EnablingEvent Then
      If MsgEvent Then
         MsgResult$ = MsgLibrary.EnableEvents(MsgEventType&, mlEventSize)
         If SaveMsg(Me, "EnableEvents(" & Format(MsgEventType&, "0") & _
         ", " & Format(mlEventSize, "0") & ")", MsgResult$) Then
            mnuEvent(Index).Checked = False
            mlEventSize = 0
            Exit Sub
         End If
      Else
         ULStat = EventEnable(mnThisInstance, mnBoardNum, EventType, EventParam, mfrmThisForm)
         If SaveFunc(Me, EnableEvent, ULStat, mnBoardNum, EventType, EventParam, _
            FormCaption$, A5, A6, A7, A8, A9, A10, A11, 0) Then
            mnuEvent(Index).Checked = False
            Exit Sub
         End If
         mlEventType = mlEventType Or EventType
         'mlEventSize = EventSize
      End If
   Else
      If MsgEvent Then
         MsgResult$ = MsgLibrary.DisableEvents(MsgEventType&)
         x% = SaveMsg(Me, "DisableEvents(" & _
         Format(MsgEventType&, "0") & ")", MsgResult$)
      Else
         ULStat = UninstallEvent(mnBoardNum, EventType)
         x% = SaveFunc(Me, DisableEvent, ULStat, mnBoardNum, EventType, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      End If
      mlEventType = mlEventType Xor EventType
   End If
   Me.mnuContPlot.ENABLED = ((mlEventType And ON_DATA_AVAILABLE) = ON_DATA_AVAILABLE)
 
 End Sub

Private Sub mnuExit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub mnuFuncArray_Click(Index As Integer)

   If (mnFormType = DIGITAL_OUT) Then
      NewFuncType% = Choose(Index + 1, DOut, DOut32, _
         DOutArray, DBitOut, DOutScan, DConfigPort, DConfigBit)
   Else
      NewFuncType% = Choose(Index + 1, DIn, DIn32, _
         DInArray, DBitIn, DInScan, DConfigPort, DConfigBit)
   End If
   Select Case NewFuncType%
      Case DInArray
         If Not LibSupportsFunction(DInArray) _
            Then Exit Sub
      Case DOutArray
         If Not LibSupportsFunction(DOutArray) _
            Then Exit Sub
      Case DIn32
         If Not LibSupportsFunction(DIn32) _
            Then Exit Sub
      Case DOut32
         If Not LibSupportsFunction(DOut32) _
            Then Exit Sub
      Case DClearAlarm
         If Not LibSupportsFunction(DClearAlarm) _
            Then Exit Sub
   End Select
   
   If Not mnFuncType = 0 Then
      For CurIndex% = 0 To mnuFuncArray.Count - 1
         If mnuFuncArray(CurIndex%).Checked Then
            OldMenuIndex% = CurIndex%
            If (mnFormType = DIGITAL_OUT) Then
               CurFuncType% = Choose(CurIndex% + 1, DOut, _
                  DOut32, DOutArray, DBitOut, DOutScan, _
                  DConfigPort, DConfigBit)
            Else
               CurFuncType% = Choose(CurIndex% + 1, DIn, DIn32, _
                  DInArray, DBitIn, DInScan, DConfigPort, DConfigBit)
               Exit For
            End If
         End If
      Next
      If mnFuncType = NewFuncType% Then Exit Sub
      mnuFuncArray(OldMenuIndex%).Checked = False
   End If
   mnFuncType = NewFuncType%
   msConfig = mnuFuncArray(Index).Caption
   mnuFuncArray(Index).Checked = True
   ConfigureControls
   'If (mnFuncType = DOutScan) Then mnuData_Click (3)
   SetReadOnWrite% = Not ((mnFuncType = DInScan) Or (mnFuncType = DConfigBit))
   mnuReadOnWrite.Checked = SetReadOnWrite% 'If (mnFormType = DIGITAL_OUT) Then
   Select Case mnLibType
      Case UNILIB
         If (mnFuncType = DBitOut) Or (mnFuncType = DBitIn) _
            Or (mnFuncType = DConfigBit) Then
            If mnUseDGrid Then
               RefreshDGrid
            Else
               GetPortConfig
            End If
         End If
      Case NETLIB
      Case MSGLIB
         If Not (MsgLibrary Is Nothing) Then
            If Not (MsgLibrary.DeviceID = "") Then
               NumProps% = GetDIOProps(msBoardName, MsgLibrary, PropsList)
               UpdateFormProps PropsList, NumProps%
               If ((mnFuncType = DBitIn) Or (mnFuncType = DBitOut) _
                  Or (mnFuncType = DConfigBit)) Then
                  'InitDGrid
                  RefreshDGrid  'RefreshGrid
               End If
            End If
         End If
   End Select

End Sub

Private Sub mnuGetStatus_Click()
   
   mnuGetStatus.Checked = Not mnuGetStatus.Checked

End Sub

Private Sub mnuHex_Click()

   mnuHex.Checked = Not mnuHex.Checked
   mbHexText = mnuHex.Checked
   
End Sub

Private Sub mnuLibrary_Click(Index As Integer)

   TypeOfLibrary% = mnLibType
   mnLibType = Choose(Index + 1, UNILIB, NETLIB, MSGLIB)
   If TypeOfLibrary% = mnLibType Then Exit Sub
   For LibMenuIndex% = 0 To mnuLibrary.Count - 1
      mnuLibrary(LibMenuIndex%).Checked = False
   Next
   mnuLibrary(Index).Checked = True
   Select Case mnLibType
      Case UNILIB
         mnNumBoards = GetNumInstalled()
         StopMsg$ = msStatusType & ":STOP"
         If Not MsgLibrary Is Nothing Then
            MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
            If mnMessaging Then
               If Not (MsgLibrary.DeviceID = "") Then _
               MsgLibrary.ReleaseDevice (msBoardName)
               mnMessaging = False
            End If
         End If
      Case NETLIB
         mnNumBoards = GetNumInstalled()
      Case MSGLIB
         mnNumBoards = GetNumMsgBoards()
         If (gnLibType > INVALIDLIB) Then
            If MsgLibrary Is Nothing Then
               Set MsgLibrary = CreateObject("MBDClass.MBDComClass")
               InitForm mnFuncType
               InitNewInstance% = True
            End If
         Else
            mnLibType = gnLibType
            Me.mnuLibrary(MSGLIB).Checked = False
            Me.mnuLibrary(MSGLIB).ENABLED = False
         End If
   End Select
   mnDetailsNeedUpdate = True
   ConfigureLibrary mnLibType
   
   For MenuIndex% = 1 To mnuBoard.Count - 1
      Unload mnuBoard(MenuIndex%)
   Next
   mnuBoard(0).Caption = "None Installed"
   For i% = 0 To mnNumBoards - 1
      If mnLibType = MSGLIB Then
         BoardNum% = i%
         BoardName$ = GetNameOfMsgBoard(BoardNum%)
         SplitName = Split(BoardName$, "::")
         DisplayName$ = SplitName(0)
      Else
         BoardNum% = gnBoardEnum(i%)
         BoardName$ = GetNameOfBoard(BoardNum%)
         DisplayName$ = BoardName$
      End If
      CurrentName$ = BoardNum% & ") " & DisplayName$
      msDisplayName = DisplayName$
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      Else
         mnBoardIndex = 0
         ConfigureControls
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   If mnNumBoards > 0 Then
      mnuBoard_Click (0)
      If mnLibType = MSGLIB Then   'If InitNewInstance% Then
         NumProps% = GetDIOProps(msBoardName, MsgLibrary, PropList)
         UpdateFormProps PropList, NumProps%
         Select Case mnFuncType
            Case DIn, DBitIn, DOut, DBitOut
               msStatusType = "DIO"
            Case DInScan, DOutScan
               msStatusType = "DIOSCAN"
         End Select
         mnRefreshProps = True
      End If
   Else
      Caption = msTitle & " Board " & mnuBoard(0).Caption
   End If
   gnNumBoards = mnNumBoards
   ReDim mnDIOPort(33, gnNumBoards)
   ReDim mnDIOBit(192, gnNumBoards)
   ReDim mlPortType(33, gnNumBoards)
   ReDim mnAuxBit(16, gnNumBoards)
   ReDim mlConfigIn(33, gnNumBoards)
   ReDim mlConfigOut(33, gnNumBoards)
   ReDim mnAuxBitDisabled(16, gnNumBoards)
   If mnUseDGrid Then
      InitDGrid
      RefreshDGrid
   Else
      GetPortConfig
      RefreshGrid
   End If

End Sub

Private Sub mnuOpt_Click(Index As Integer)

   mnuOpt(Index).Checked = Not mnuOpt(Index).Checked
   'mnUse32 = False
   Select Case Index
      Case 3   'WORDXFER
         If mnuOpt(Index + 1).Checked Then mnuOpt(Index + 1).Checked = False
         'Unload frmBuildData
         'UseWinAPI% = Not Me.mnuULBuf.Checked
         'mlGenHandle = GenerateData(1, 1, 0, 1, _
         '0, 0, Signal%, False, Chan%, FirstPoint&, UseWinAPI%)
         AmpVal& = ConvertDataType(2, Val(txtAmpl.Text))
         OSVal& = ConvertDataType(2, Val(txtOffset.Text))
         If mnuOpt(Index).Checked Then
            If (AmpVal& < 4096) And (OSVal& < 4096) Then
               txtAmpl.Text = 32767
               txtOffset.Text = 32768
            End If
         Else
            If Not ((AmpVal& < 4096) And (OSVal& < 4096)) Then
               txtAmpl.Text = 127
               txtOffset.Text = 128
            End If
         End If
         'ResetData
      Case 4   'DWORDXFER
         If mnuOpt(Index - 1).Checked Then mnuOpt(Index - 1).Checked = False
         'Unload frmBuildData
         'UseWinAPI% = Not Me.mnuULBuf.Checked
         'mlGenHandle = GenerateData(1, 1, 0, 1, _
         '0, 0, Signal%, False, Chan%, FirstPoint&, UseWinAPI%)
         AmpVal& = ConvertDataType(2, Val(txtAmpl.Text))
         OSVal& = ConvertDataType(2, Val(txtOffset.Text))
         If mnuOpt(Index).Checked Then
            If (AmpVal& < 32767) And (OSVal& < 32767) Then
               txtAmpl.Text = 2147483647#
               txtOffset.Text = 2147483648#
            End If
         Else
            If Not ((AmpVal& < 4096) And (OSVal& < 4096)) Then
               txtAmpl.Text = 127
               txtOffset.Text = 128
            End If
         End If
         'ResetData
      Case 5   'SINGLEIO
         mnuOpt(6).Checked = False
         mnuOpt(7).Checked = False
      Case 6   'DMAIO
         mnuOpt(7).Checked = False
         mnuOpt(5).Checked = False
      Case 7   'BLOCKIO
         mnuOpt(5).Checked = False
         mnuOpt(6).Checked = False
   End Select
   mnuContPlot.Checked = mnuOpt(1).Checked And mnuContPlot.Checked
   mnuContPlot.ENABLED = mnuOpt(1).Checked
   'Bits% = GetDIOResolution()
   SetOptions
   mnUse32 = ((mvOptions And DWORDXFER) = DWORDXFER)
   Res% = GetDIOResolution()
   mbDataGenRequired = mbDataGenRequired Or Not (Res% = mnResolution)
   mnResolution = Res%
   UpdateStatBar False
   mbDataGenRequired = mbDataGenRequired Or _
      (mnUse32 = Not (mnGenDataType = vbLong))

End Sub

Private Sub mnuPlotType_Click(Index As Integer)

   If Index = PRINT_TEXT Then
      mnuPlotType(Index).Checked = Not mnuPlotType(Index).Checked
      If mnuPlotType(Index).Checked Then mlBlockSave = GetBlockSize()
      If mnuPlotType(VOLTS_VS_TIME).Checked Then
         If Not mnuPlotType(PRINT_TEXT).Checked Then SetBlockSize mlBlockSave, False
      End If
      ShowText mnuPlotType(Index).Checked
      DisplayData
      Exit Sub
   End If

   mnPlot = False
   If Index = mnPlotType Then
      mnuPlotType(mnPlotType).Checked = Not mnuPlotType(mnPlotType).Checked
   Else
      mnuPlotType(mnPlotType).Checked = False
      mnuPlotType(Index).Checked = True
   End If
   
   If mnuPlotType(Index).Checked Then
      mnPlotType = Index
      SetPlotType mnPlotType + mnHardCopy, Me
      mnPlot = True
   End If
   If Index = SINGLE_POINT Then
      mlCount = Val(txtCount.Text)
      mnuContPlot.ENABLED = mnuPlotType(SINGLE_POINT).Checked
      SetBlockSize mlCount, False
   Else
      DisplayData
   End If

End Sub

Private Sub mnuPrint_Click()

   ExecutePrint

End Sub

Private Sub mnuReadOnWrite_Click()

   mnuReadOnWrite.Checked = Not mnuReadOnWrite.Checked

End Sub

Private Sub mnuSelData_Click()

   'new data creation form
   GenPoints& = Me.txtCount.Text
   GenDataType = 1
   mnuULBuf.Checked = True
   If mbDataGenRequired Then _
      Unload frmBuildData
   
   mnResolution = GetDIOResolution()
   
   If mnResolution > 8 Then _
         mnGenDataType = vbInteger

   If mnResolution > 16 Then _
      mnGenDataType = vbLong
   
   If mnGenDataType > 0 Then GenDataType _
      = Choose(mnGenDataType, 1, 1, 2, 4, 6)
   If mlGenHandle = 0 Then
   '   frmBuildData.txtAmplitude.Text = "50%"
   '   frmBuildData.txtOffset.Text = "50%"
   '   frmBuildData.txtSamples.Text = 1000
      If (mnFuncType = DOutArray) Then
         If Val(frmBuildData.txtSamples.Text) > mnLastPort Then
            frmBuildData.txtSamples.Text = mnLastPort
            GenPoints& = mnLastPort
         End If
      Else
         frmBuildData.txtSamples.Text = Me.txtCount.Text
      End If
   End If
   
   frmBuildData.SetFormRef Me
   frmBuildData.SetDefaults mnBoardNum, mnResolution, 0, _
   GenPoints&, 1, GenDataType, UseWinAPI%, CountIsPerChan%, GenFloats%
   
   frmBuildData.Show False, mfmUniTest
   
End Sub

Private Sub mnuStopBG_Click()

   mnuStopBG.Checked = Not mnuStopBG.Checked

End Sub

Private Sub mnuStore_Click()

   mnuStore.Checked = Not mnuStore.Checked
   mnuPrint.ENABLED = mnuStore.Checked
   If mnuStore.Checked Then
      mnHardCopy = &H10
   Else
      mnHardCopy = 0
   End If
   SetPlotType mnPlotType + mnHardCopy, Me

End Sub

Private Sub mnuTimer_Click()

   If mnuTimer.Checked Then frmSetTimer.chkEnableTimer.value = 1
   If mnTimerStop Then frmSetTimer.chkTimerStopBG = 1
   If mnDelayRestart Then
      frmSetTimer.chkDelayRestart.value = 1
      frmSetTimer.txtDelayTime.Text = Format(mlDelayTime, "0")
   End If

   frmSetTimer.txtInterval.Text = tmrGoLoop.Interval
   frmSetTimer.optTimerMode(1).value = mnTimerTillCount
   frmSetTimer.Show 1
   
   LoopRate& = Val(frmSetTimer.txtInterval.Text)
   mnTimerStop = (frmSetTimer.chkTimerStopBG = 1)
   mnDelayRestart = (frmSetTimer.chkDelayRestart.value = 1)
   mlDelayTime = Val(frmSetTimer.txtDelayTime.Text)
   If frmSetTimer.optTimerMode(1).value Then
      mlCount = 1
      mnTimerTillCount = True
   Else
      mlCount = Val(txtCount.Text)
      mnTimerTillCount = False
   End If
   mnuTimer.Checked = (frmSetTimer.chkEnableTimer.value = 1)
   Unload frmSetTimer
   If LoopRate& > 0 Then
      tmrGoLoop.Interval = LoopRate&
      tmrDIn.Interval = LoopRate&
   Else
      mnuTimer.Checked = False
   End If
   If Not mnuTimer.Checked Then
      tmrGoLoop.ENABLED = False
      tmrDIn.ENABLED = False
      cmdGo.FontBold = True
   End If
   
   cmdStop.Visible = mnuTimer.Checked

End Sub

Private Sub mnuULBuf_Click()

   mnuULBuf.Checked = Not (mnuULBuf.Checked)

End Sub

Private Sub picAuxIO_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   If mnDIOPort(0, mnBoardIndex) = NOTCONFIGURED Then
      OnColor& = ONNOCONFCOLOR&
      OffColor& = NOCONFCOLOR
   Else
      OnColor& = OUTONCOLOR&
      OffColor& = OUTOFFCOLOR&
   End If
   If Not (Button = 0) Then
      ColorVal& = picAuxIO.Point(x, y)
      BitValue% = (ColorVal& = OnColor&) + 1
   Else
      BitValue% = Shift
   End If
   BitNum% = Int(x) * 8 + Int(y)
   ColorVal& = NOCONFCOLOR&

   If mnFuncType = DConfigBit Then
      Direction& = DIGITALIN
      If Shift = 1 Then Direction& = DIGITALOUT

      'following required for scripting info
      A5 = x
      A6 = y
      Select Case mnLibType
         Case UNILIB
            ULStat = BitConfig(mnBoardNum, AUXPORT, BitNum%, Direction&)
            If SaveFunc(Me, DConfigBit, ULStat, mnBoardNum, _
               AUXPORT, BitNum%, Direction&, A5, A6, A7, _
               A8, A9, A10, A11, 0) Then Failure% = True
         Case MSGLIB
            PortIndex% = Int(x)
            BitIndex% = Int(y)
            FuncResult% = SetMsgBitConfig(PortIndex%, BitIndex%, Direction&)
            Failure% = Not FuncResult%
      End Select
      If Failure% Then
         Exit Sub
      End If
      mnAuxBit(BitNum%, mnBoardIndex) = Direction&
      If mnLibType = UNILIB Then
         GetPortConfig
         If Direction& = DIGITALIN Then
            ColorVal& = INOFFCOLOR&
         Else
            ColorVal& = OUTOFFCOLOR&
         End If
      Else
      End If
      picAuxIO.Line (Int(x) + 0.1, Int(y) + 0.2)-(Int(x) _
         + 0.8, Int(y) + 0.8), ColorVal&, BF
      RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      Exit Sub
   End If

   If mnFormType = DIGITAL_OUT Then
      'check if configured for output
      'following required for scripting info
      A5 = x
      A6 = y
      Select Case mnLibType
         Case UNILIB
            ULStat = cbDBitOut(mnBoardNum, AUXPORT, BitNum%, BitValue%)
            If SaveFunc(Me, DBitOut, ULStat, mnBoardNum, _
               AUXPORT, BitNum%, BitValue%, A5, A6, _
               A7, A8, A9, A10, A11, 0) Then Exit Sub
         Case MSGLIB
            BitIndex% = Int(y)
            PortIndex% = Int(x)
            BitValSet% = SetMsgBitVal(PortIndex%, BitIndex%, BitValue%)
            If Not (BitValSet% = BitValue%) Then Exit Sub
      End Select

      If BitValue% Then
         ColorVal& = OnColor&
      Else
         ColorVal& = OffColor&
      End If
      picAuxIO.Line (Int(x) + 0.1, Int(y) + 0.2)-(Int(x) _
         + 0.8, Int(y) + 0.8), ColorVal&, BF
      RunOtherForms DIGITAL_IN, UPDATEDIO
   ElseIf mnFormType = DIGITAL_IN Then
      If Button = 2 Then
         If mnAuxBitDisabled(BitNum%, mnBoardIndex) Then
            Dummy% = GetBitNum(BitNum%, AUXPORT, PortIndex%)
            Stat% = GetBitStatus(PortIndex%, BitNum%, mnBoardIndex)
            ColorVal& = INOFFCOLOR&
            If Stat% = DIGITALOUT Then ColorVal& = OUTOFFCOLOR&
         Else
            ColorVal& = &HFFFFFF
         End If
         picAuxIO.Line (Int(x) + 0.1, Int(y) + 0.2)-(Int(x) _
            + 0.8, Int(y) + 0.85), ColorVal&, BF
         mnAuxBitDisabled(BitNum%, mnBoardIndex) _
            = Not mnAuxBitDisabled(BitNum%, mnBoardIndex)
      End If
   End If

End Sub

Private Sub picAuxIO_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   BitNum% = Int(x) * 8 + Int(y)
   lblStatus.Caption = msConfig & "  AUXPORT, Bit " & _
      BitNum% & ",  " & masConfig(mnDIOPort(0, mnBoardIndex))

End Sub

Private Sub picDIO_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   OnColor& = OUTONCOLOR&
   OffColor& = OUTOFFCOLOR&
   If mnuTimer.Checked Then
      tmrGoLoop.ENABLED = True
      cmdStop.ENABLED = True
   End If
   If (mnFormType = DIGITAL_IN) And Not _
      (mnuFuncArray(BIT_IN).Checked Or _
      mnuFuncArray(DBITCONF).Checked) Then
      lblStatus.Caption = ""
      Exit Sub
   End If
   Bit% = Int(y)
   Port% = Int(x)
   If Bit% > 7 Then Bit% = 7
   If Port% > 34 Then Port% = 34
   If (mnLibType = MSGLIB) Or mnUseDGrid Then
      If mnLibType = MSGLIB Then
         PortNum% = DGrid(Bit%, Port%).PortNum
         BitNum% = DGrid(Bit%, Port%).BitNum
      Else
         PortNum% = DGrid(Bit%, Port%).ULBitPort
         BitNum% = DGrid(Bit%, Port%).ULBitNum
      End If
   Else
      PortNum% = Port% + 10
      For TestPort% = 0 To 32
         PortType& = mlPortType(TestPort%, mnBoardIndex)
         If (TestPort% = 0) And (PortType& > 10) And _
            Not (PortNum% < PortType&) _
            Then BitOffset% = (PortType& - 10) * 8
         If PortNum% = PortType& Then Exit For
      Next
      PortIndex% = TestPort%
      StartBit% = GetStartBit(PortIndex%)
      BitNum% = BitOffset% + StartBit% + Bit%
   
      If mnDIOPort(PortIndex%, mnBoardIndex) = NOTCONFIGURED Then
         OnColor& = ONNOCONFCOLOR&
         OffColor& = NOCONFCOLOR
      End If
   End If
   
   If Not (Button = 0) Then
      ColorVal& = picDIO.Point(x, y)
      BitValue% = (ColorVal& = OnColor&) + 1
   Else
      BitValue% = Shift
   End If
   
   If mnFuncType = DConfigBit Then
      Direction& = DIGITALIN
      ColorVal& = INOFFCOLOR&
      If Shift = 1 Then
           Direction& = DIGITALOUT
           ColorVal& = OUTOFFCOLOR&
      End If
      'following required for scripting info
      A5 = x
      A6 = y
      LibPort% = FIRSTPORTA
      If mnUseDGrid Then LibPort% = PortNum%
      If mnLibType = MSGLIB Then
         MsgPort$ = "{" & Format(PortNum%, "0") & "/" & Format(BitNum%, "0") & "}"
         DirString$ = ":DIR=IN"
         If Direction& = DIGITALOUT Then DirString$ = ":DIR=OUT"
         DConfigMessage$ = msStatusType & MsgPort$ & DirString$
         MsgResult$ = MsgLibrary.SendMessage(DConfigMessage$)
         If SaveMsg(Me, "SendMessage(" & DConfigMessage$ & ")", MsgResult$) Then Failure% = True
      Else
         ULStat = BitConfig(mnBoardNum, LibPort%, BitNum%, Direction&)
         If SaveFunc(Me, DConfigBit, ULStat, mnBoardNum, LibPort%, BitNum%, Direction&, A5, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
      End If
      If Failure% Then
         Exit Sub
      End If
      'to do - change name to generic library
      If Direction& = DIGITALIN Then BitVal% = GetMsgBitVal(PortNum%, BitNum%)
      If mnUseDGrid Then DGrid(Bit%, Port%).BitConf = Direction&
      PaintDIOGrid Int(x), Int(y), ColorVal&
      ColBit% = 0
      PortStart% = Port%
      If DGrid(ColBit%, Port%).BitNum > 0 Then PortStart% = Port% - 1
      CurPort% = PortStart%
      PortConf& = Direction&
      For GridBit% = 0 To mnResolution - 1
         If DGrid(ColBit%, CurPort%).BitConf <> Direction& Then
            PortConf& = MIXEDPORT
            Exit For
         End If
         ColBit% = ColBit% + 1
         If ColBit% = 8 Then
            ColBit% = 0
            CurPort% = CurPort% + 1
         End If
      Next
      DGrid(0, PortStart%).PortType = PortConf&
      RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      Exit Sub
   End If
      
   If Not mnFormType = DIGITAL_IN Then
      If (mnLibType = MSGLIB) Then
         ValSet% = SetMsgBitVal(PortNum%, BitNum%, BitValue%)
         If Not (ValSet% = 0) Then
            ColorVal& = OnColor&
         Else
            ColorVal& = OffColor&
         End If
         PaintDIOGrid Port%, Bit%, ColorVal&
      Else
         'check if configured for output
         'following required for scripting info
         A5 = x
         A6 = y
         LibPort% = FIRSTPORTA
         If mnUseDGrid Then LibPort% = PortNum%
         ULStat = cbDBitOut(mnBoardNum, LibPort%, BitNum%, BitValue%)
         d% = SaveFunc(Me, DBitOut, ULStat, mnBoardNum, LibPort%, BitNum%, BitValue%, A5, A6, A7, A8, A9, A10, A11, 0)
         If Not ULStat = 0 Then
         Else
            If BitValue% Then
               ColorVal& = OnColor&
            Else
               ColorVal& = OffColor&
            End If
            PaintDIOGrid Int(x), Int(y), ColorVal&
         End If
      End If
      RunOtherForms DIGITAL_IN, UPDATEDIO
   End If

End Sub

Private Sub picDIO_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   Bit% = Int(y)
   If Bit% > 7 Then Bit% = 7
   Port% = Int(x)
   If (mnLibType = MSGLIB) Or mnUseDGrid Then
      If Not DGrid(Bit%, Port%).BitConf = INVALIDBIT Then
         PortString$ = DGrid(Bit%, Port%).PortName
         BitString$ = Format(DGrid(Bit%, Port%).BitNum, "0")
         AltBitCount$ = ""
         LongPort% = True '(DGrid(Bit%, Port%).PortNum > FIRSTPORTA)
         If (Not (mnLibType = MSGLIB)) And LongPort% Then
            AltBitString$ = Format(DGrid(Bit%, Port%).ULBitNum, "0")
            AltPortString$ = DGrid(Bit%, Port%).ULBitPortName
            AltBitCount$ = "  (" & AltPortString$ & ", Bit " & AltBitString$ & ")"
         End If
         lblStatus.Caption = msConfig & "   " & PortString$ & _
         " bit " & BitString$ & AltBitCount$
      Else
         lblStatus.Caption = msConfig
      End If
   Else
      PortNum% = Port% + 10
      BitNum% = GetBitNum(Bit%, PortNum%, PortIndex%)
      lblStatus.Caption = msConfig & "   " & GetPortString(PortNum%) & ", Bit " & Int(y)
      lblStatus.Caption = lblStatus.Caption & " (FIRSTPORTA bit " & BitNum% & ")" & ",  " & masConfig(mnDIOPort(PortIndex%, mnBoardIndex))
   End If
   
End Sub

Private Sub ResetData()

   mlCount = Val(txtCount.Text)
   If mlGenPoints > 1 Then
      If mlCount > mlGenPoints _
         Then mbDataGenRequired = True
   End If
   Exit Sub
   
   If (mnFuncType = DOutScan) Then
      CodeRange& = Val(txtAmpl.Text)
      CenterCode& = Val(txtOffset.Text)
      If mnDataType > 0 Then
         Chans% = mnLastChan - mnFirstChan
         UseLibrary% = mnuULBuf.Checked
         InitOutputBuffer mlGenHandle, mnDataType, mlCount, _
         Chans%, CodeRange&, CenterCode&, UseLibrary%
         txtCount.Text = mlCount
      Else
         If Not mlGenHandle = 0 Then
            NewData% = True
            Channels% = (mnLastChan - mnFirstChan) + 1
            Amplitude = CodeRange&
            Offset = CenterCode&
            Select Case mnGenDataType
               Case vbInteger
                  DataType& = 1
               Case vbLong
                  DataType& = 2
            End Select
            Cycles% = 2: Signal% = 2
            UseWinAPI% = Not Me.mnuULBuf.Checked
            mlGenHandle = GenerateData(DataType&, Cycles%, _
            mlCount, Channels%, Amplitude, Offset, Signal%, _
            NewData%, Chan%, FirstPoint&, UseWinAPI%)
            If Not mlGenHandle = 0 Then mlGenPoints = mlCount
         End If
      End If
   End If

End Sub

Private Sub RunOtherForms(FormType As Integer, Interface As Integer)

   NumForms% = Forms.Count
   For ThisForm% = 0 To NumForms% - 1
      If Not Forms(ThisForm%) Is Me Then
         FormName$ = Forms(ThisForm%).Caption
         Select Case FormType
            Case DIGITAL_OUT, DIGITAL_IN, DIGITAL_OUT + DIGITAL_IN
               PartialName$ = Left$(FormName$, 13)
               If (PartialName$ = "Digital Outpu") Or (PartialName$ = "Digital Input") Then
                  If Interface = UPDATEGRID Then Forms(ThisForm%).cmdReset = True
                  'If mnuReadOnWrite.Checked Then Forms(ThisForm%).cmdGo = True
                  Forms(ThisForm%).GoCommand
               End If
         End Select
      End If
   Next ThisForm%

End Sub

Private Sub SetOptions()
   
   mvOptions = FOREGROUND  'init to 0
   msOpt = ""
   For i% = 0 To mnuOpt.Count - 1
      If mnuOpt(i%).Checked Then
         Options = Choose(i% + 1, BACKGROUND, CONTINUOUS, EXTCLOCK, WORDXFER, _
         DWORDXFER, SINGLEIO, DMAIO, BLOCK_IO, SIMULTANEOUS, EXTTRIGGER, NONSTREAMEDIO, _
         ADCCLOCKTRIG, ADCCLOCK, HIGHRESRATE, RETRIGMODE)
         opt$ = Choose(i% + 1, "BACKGROUND", "CONTINUOUS", "EXTCLOCK", "WORDXFER", _
         "DWORDXFER", "SINGLEIO", "DMAIO", "BLOCKIO", "SIMULTANEOUS", "EXTTRIGGER", "NONSTREAMEDIO", _
         "ADCCLOCKTRIG", "ADCCLOCK", "HIGHRESRATE", "RETRIGMODE")
         If Not IsNull(Options) Then mvOptions = mvOptions + Options
         If Not IsNull(opt$) Then msOpt = msOpt & opt$ & " "
      End If
   Next i%
   If Len(msOpt) = 0 Then msOpt = "Options = Default  "
   ConfigureControls
   mnDetailsNeedUpdate = True
   
End Sub

Private Sub ShowConfigStat()

   PortIndex% = hsbPort.value - 1
   If mnPortIndex <> PortIndex% Then
      lblStatus.Caption = Left$(msConfig, Len(msConfig) - 1) & GetPortString(mnPortNum) & ", INVALIDPORT)"
   Else
      GridCol% = 0
      Do
         ColIndex% = DGrid(0, GridCol%).PortIndex
         GridPortType& = DGrid(0, GridCol%).PortNum
         Valid% = Not (ColIndex% < 0)
         NotComplete% = Valid% And (GridPortType& <> mnPortNum)
         GridCol% = GridCol% + 1
      Loop While NotComplete%
      If Not Valid% Then
         PortStat% = 5
      Else
      GridCol% = GridCol% - 1
      PortStat% = DGrid(0, GridCol%).PortType
      'If Not (DGrid(0, GridCol%).PortNum = mnPortNum) Then PortStat% = 5
      End If
      lblStatus.Caption = msStringID & Left$(msConfig, Len(msConfig) - 1) & GetPortString(mnPortNum) & ", " & masConfig(PortStat%) & ")"
   End If

End Sub

Private Sub tmrCheckStatus_Timer()

   Static Do2ndHalf%, DoFirstHalf%
   RunningBackground% = ((mvOptions And BACKGROUND) = BACKGROUND)
   RunningContinuous% = ((mvOptions And CONTINUOUS) = CONTINUOUS)
   x$ = Me.Caption
   If mnCancel Or gnCancel Then
      tmrCheckStatus.ENABLED = False
      tmrGoLoop.ENABLED = False
      cmdGo.FontBold = True
      cmdStop = True
      gnCancel = False
      mnCancel = False
      If mnSetActive Then SetContActive False
      Element& = 0
      If RunningBackground% Then UpdateStatBar False
      If mnuStopBG.Checked Then
         Select Case mnLibType
            Case UNILIB
               ULStat = StopBackground520(mnBoardNum, mlStatusType)
            Case NETLIB
               ULStat = NetLibrary.StopBackground(mlStatusType)
            Case MSGLIB
               If mnMessaging And Not (msScanSupport = "") Then
                  MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
                  If MsgResult$ = "AISCAN:STATUS=INTERRUPTED" Then
                     ULStat = INTERRUPTED
                     MsgQualifier$ = "Error: "
                     MsgResult$ = MsgQualifier$ & MsgResult$
                  End If
               End If
         End Select
         If (Not gnScriptSave) Or (ULStat <> 0) Then
            Select Case mnLibType
               Case UNILIB, NETLIB
                  If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, _
                  A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               Case MSGLIB
                  If SaveMsg(Me, "SendMessage(" & StopMsg$ & ")", MsgResult$) Then Exit Sub
            End Select
         End If
         cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      DoEvents
      mnPlot = True
      'RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      Exit Sub
   End If
   
   If mnuGetStatus.Checked Then
      ULStat = GetStatus520(mnBoardNum, Status%, _
         CurCount&, CurIndex&, mlStatusType)
      If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, _
            Status%, CurCount&, CurIndex&, mlStatusType, _
            A6, A7, A8, A9, A10, A11, 0) Then
         tmrCheckStatus.ENABLED = False
         mnCancel = True
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
         If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, _
            mnFormType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Status% = IDLE
         Do2ndHalf% = False
         DoFirstHalf% = False
         If Not mnuTimer.Checked Then cmdStop.Visible = False
         DoEvents
         mnPlot = True
         If Not (mnFormType = DIGITAL_OUT) Then
            EvaluateCount% = (frmEvalData.chkCountStop.value = 1) _
               Or (frmEvalData.chkCountDelta.value = 1)
            If mbEvalEnabled And EvaluateCount% _
               Then EvalResult% = EvalCount(CurCount&)
         End If
         Exit Sub
      End If
   Else
      Status% = RUNNING
   End If
   If mbHalfRoll Then
      If CurCount& > 0 Then
         mnCountRolls = mnCountRolls + 1
         mbHalfRoll = False
      End If
   Else
      mbHalfRoll = (CurCount < 0)
   End If
   
   UpdateStatBar False
   If mnCancel Or gnCancel Then
      tmrCheckStatus.ENABLED = False
      tmrGoLoop.ENABLED = False
      cmdGo.FontBold = True
      cmdStop = True
      gnCancel = False
      mnCancel = False
      If mnuStopBG.Checked Then
         Select Case mnLibType
            Case UNILIB
               ULStat = StopBackground520(mnBoardNum, mlStatusType)
         End Select
      End If
      If Not mnuTimer.Checked Then cmdStop.Visible = False
   End If
   If mnuContPlot.Checked Then
      BufferSplit& = mlCount \ 2
      If (Not Do2ndHalf%) And (Not DoFirstHalf%) Then
         mlFirstPoint = 0
         If CurIndex& > BufferSplit& Then
            DoFirstHalf% = True
            mnPlot = True
         End If
      End If
      If DoFirstHalf% Then
         If CurIndex& > BufferSplit& Then
            DoFirstHalf% = False
            Do2ndHalf% = True
            DisplayData
            mlFirstPoint = (BufferSplit&) - ((BufferSplit&) _
               Mod (mnLastChan - mnFirstChan + 1))
         End If
      End If
      If Do2ndHalf% Then
         If CurIndex& < BufferSplit& Then
            DoFirstHalf% = True
            Do2ndHalf% = False
            DisplayData
            mlFirstPoint = 0
         End If
      End If
   End If
   
   If Status% = IDLE Then
      If mnuStopBG.Checked Then
         'ULStat = cbStopBackground(mnBoardNum)
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
         If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, _
            mnFormType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         If Not mnuTimer.Checked Then cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      tmrCheckStatus.ENABLED = False
      UpdateStatBar False
      DoEvents
      mnPlot = True
      If Not (mnFormType = DIGITAL_OUT) Then DisplayData
      'If Me.mnuReadOnWrite.Checked Then RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
   End If

End Sub

Private Sub tmrDelay_Timer()

   tmrDelay.ENABLED = False

End Sub

Private Sub tmrDIn_Timer()

   If mnSemaphore Then Exit Sub
   cmdGo.FontBold = Not cmdGo.FontBold
   Select Case mnFuncType
      Case DBitIn
         Select Case mnLibType
            Case UNILIB
               DoBitIn
            Case MSGLIB
               DoMsgBitIn
         End Select
      Case DIn, DIn32, DInArray
         Select Case mnLibType
            Case UNILIB
               DoDIn
            Case MSGLIB
               DoMsgDIn
         End Select
      Case DInScan
         DoDInScan
         DisplayData
   End Select

End Sub

Private Sub tmrGoLoop_Timer()

   cmdGo.FontBold = Not cmdGo.FontBold
   If gnCancel Or mnCancel Then
      If tmrGoLoop.ENABLED Then
         tmrGoLoop.ENABLED = False
         cmdGo.FontBold = True
         mbManualStop = True
         cmdStop = True
      End If
      gnCancel = False
      mnCancel = False
      lblStatus = "{Loop #" & mlTimerCount & "} " & lblStatus
      mlTimerCount = 0
      Exit Sub
   End If
   If mnTimerStop Then
      'to do - implement timer stopbg for msg daq
      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
      DevActive% = (Status% = RUNNING)
      'If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
      If DevActive% Then
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
      End If
   End If
   If mnDelayRestart Then
      OrgInterval& = tmrDelay.Interval
      Me.tmrDelay.Interval = mlDelayTime
      Me.tmrGoLoop.ENABLED = False
      tmrDelay.ENABLED = True
      Do
         DoEvents
      Loop While tmrDelay.ENABLED
      tmrDelay.Interval = OrgInterval&
      Me.tmrGoLoop.ENABLED = True
   End If
   Select Case mnFuncType
      Case DBitOut
         If mnFirstPort = 0 Then
            x! = 0.5
            y! = 0.5 + Rnd() * (mnLastAuxBit)
            If Not (y! < 8) Then
               x! = 1.5
               y! = x! - 8
            End If
            picAuxIO_MouseDown 1, 0, x!, y!
         End If
         If mnLastPort > 0 Then
            x! = 0.5 + Rnd() * (mnLastPort - 1) '3
            y! = 0.5 + Rnd() * 7
            Select Case Int(x!)
               Case 2, 3, 6, 7, 10, 11, 14, 15, 18, 19, 22, 23, 26, 27, 30, 31
                  CompVal% = 4
               Case Else
                  CompVal% = 0
            End Select
            'If ((Int(X!) + 1) Mod 3 = 0) And Y! > 4 Then Y! = Y! - 4
            'If ((Int(X!) + 1) Mod 4 = 0) And Y! > 4 Then Y! = Y! - 4
            y! = y! - CompVal%
            picDIO_MouseDown 1, 0, x!, y!
         End If
      Case DOut
         'hsbOutVal_Scroll
         If hsbOutVal.value < hsbOutVal.Max Then
            hsbOutVal.value = hsbOutVal.value + 1
         Else
            hsbOutVal.value = 0
         End If
      Case DInScan, DOutScan
         cmdGo = True
      Case Else
         cmdSend = True
   End Select

End Sub

Private Sub txtAmpl_KeyPress(KeyAscii As Integer)

   mnDataChange = True

End Sub

Private Sub txtAmpl_LostFocus()

   ResetData

End Sub

Private Sub txtCount_KeyPress(KeyAscii As Integer)

   mnDataChange = True
   mnDetailsNeedUpdate = True

End Sub

Private Sub txtCount_LostFocus()

   ResetData
   mnDetailsNeedUpdate = True

End Sub

Private Sub txtOffset_KeyPress(KeyAscii As Integer)

   mnDataChange = True

End Sub

Private Sub txtOffset_LostFocus()

   ResetData

End Sub

Private Sub txtOutVal_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Then
      DoDOut
      If mnuReadOnWrite.Checked Then RunOtherForms DIGITAL_IN, UPDATEDIO
   End If

End Sub

Private Sub txtRate_LostFocus()
   
   mfRate = Val(txtRate.Text)
   mfRateReturned = mfRate
   mnDetailsNeedUpdate = True

End Sub

Private Sub UpdateMainStatus()

   board$ = mnuBoard(mnBoardIndex).Caption
   PrintMain "Current board: " & board$
   
End Sub

Private Sub UpdateStatBar(CycleStatus As Integer)
   
   Static sStatus As Integer
   
   If CycleStatus Then
      sStatus = sStatus + 1
      If sStatus > 3 Then sStatus = 0
      StatType% = sStatus
   ElseIf tmrCheckStatus.ENABLED Then
      StatType% = 1
   Else
      StatType% = 0 'sStatus
   End If
      
   Select Case StatType%
      Case 0
         If mnFuncType = CB_AIN Then
            Stat$ = msConfig
         Else
            Stat$ = msConfig & "  " & msOpt
         End If
         lblStatus.Caption = Stat$
      Case 1
         If mnuGetStatus.Checked Then
            If (gnNumBoards > 0) Then
               If mnLibType = UNILIB Then
                  ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, _
                     CurIndex&, mlStatusType)
                  If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, _
                     CurCount&, CurIndex&, mlStatusType, A6, A7, A8, _
                     A9, A10, A11, 0) Then Exit Sub
               End If
            Else
               Status% = 3
            End If
            Select Case Status%
               Case IDLE
                  BGOp$ = "No background task running..."
               Case RUNNING
                  BGOp$ = "Background task in progress..."
               'Case Else
               '   BGOp$ = "Background task not monitored..."
               '   Stat$ = BGOp$
            End Select
            If mnCountRolls > 1 Then RollString$ = _
               " (X" & mnCountRolls & ")"
            Stat$ = BGOp$ & " Count = " & CurCount& & _
               "  Index = " & CurIndex& & RollString$
            lblStatus.Caption = Stat$
         End If
      Case 2
         Stat$ = "Trigger type = " & GetTrigString(mnTrigType)
         lblStatus.Caption = Stat$
   End Select
   
End Sub

Private Sub GetPortConfig()

   'check the library version of port configuration
   'don't report errors here
   If Not (gnNumBoards > 0) Then
      For Ctl& = 0 To Me.Controls.Count - 1
         ContType$ = Left(Controls(Ctl&).Name, 3)
         If (ContType$ = "cmd") Or _
            (ContType$ = "chk") Or _
            (ContType$ = "pic") Then _
            Controls(Ctl&).ENABLED = False
      Next
      Exit Sub
   End If
   SaveLocalErrDisp% = gnLocalErrDisp
   SaveLocalErrFlow% = geErrFlow
   gnLocalErrDisp = False
   geErrFlow = 1

   Static stnBoardNum As Integer
   GridNeedsUpdate% = False
   If stnBoardNum <> mnBoardIndex Then GridNeedsUpdate% = True
   stnBoardNum = mnBoardIndex
   PortNum% = 0
   For PortIndex% = 0 To 32
      InfoType% = DIGITALINFO
      If mnuDIODetect.Checked Then
         GridNeedsUpdate% = True
         ConfigItem% = DIDEVTYPE
         ULStat = GetConfig520(InfoType%, mnBoardNum, PortIndex%, ConfigItem%, PortType&)
         ConfigErr = SaveFunc(Me, GetConfig, ULStat, InfoType%, mnBoardNum, PortIndex%, ConfigItem%, PortType&, A6, A7, A8, A9, A10, A11, 0)
         mlPortType(PortIndex%, mnBoardIndex) = PortType&
         If PortType& > 0 Then
            Do
               PortCompare& = GetPortFromIndex(PortNum%)
               If Not (PortType& = PortCompare&) Then
                  mnDIOPort(PortNum%, mnBoardIndex) = INVALIDPORT
                  PortNum% = PortNum% + 1
                  If PortNum% > 33 Then
                     PortNum% = PortIndex%
                     Matched% = True
                  End If
               Else
                  Matched% = True
               End If
            Loop While Not Matched%
         End If
         If ULStat = BADCONFIGITEM Then
            If PortNum% > 33 Then Exit For
            mnDIOPort(PortNum%, mnBoardIndex) = INVALIDPORT
         Else
            If (EndBit% = 0) And ((PortType& <> AUXPORT) Or (PortType& <> FIRSTPORTA)) Then
               'the first bit is higher than zero, as with boards with only FIRSTPORTCL and CH
               'fill in the first bits as invalid (assume 8 bit ports)
               'If 0 Then
               For SkipPort% = PortType& To (FIRSTPORTA + 1) Step -1
                  For SkipBit% = StartBitSkip% To StartBitSkip% + 7
                     mnDIOBit(SkipBit%, mnBoardIndex) = INVALIDBIT
                  Next
                  StartBitSkip% = SkipBit% '+ 1
               Next
               EndBit% = StartBitSkip%
               LastBit% = EndBit%
               'End If
            End If
            ConfigItem% = DIINMASK
            ULStat = GetConfig520(InfoType%, mnBoardNum, DevNum%, ConfigItem%, InConfig&)
            ConfigErr = SaveFunc(Me, GetConfig, ULStat, InfoType%, mnBoardNum, DevNum%, ConfigItem%, InConfig&, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = BADCONFIGITEM Then
               mnDIOPort(PortNum%, mnBoardIndex) = INVALIDPORT
            Else
               mlConfigIn(PortIndex%, mnBoardIndex) = InConfig&
               ConfigItem% = DIOUTMASK
               ULStat = GetConfig520(InfoType%, mnBoardNum, DevNum%, ConfigItem%, OutConfig&)
               'ConfigErr = SaveFunc(Me, GetConfig, ULStat, InfoType%, mnBoardNum, DevNum%, ConfigItem%, OutConfig&, A6, A7, A8, A9, A10, A11, 0)
               mlConfigOut(PortIndex%, mnBoardIndex) = OutConfig&
               mnDIOPort(PortNum%, mnBoardIndex) = GetPortStatus(PortIndex%, mnBoardIndex)
               'check the bits
               PortBit% = 0
               For BitNum% = EndBit% To (EndBit% + 15) 'no ports have more than 16 bits
                  InBit% = (mlConfigIn(PortIndex%, mnBoardIndex) And 2 ^ PortBit%) = 2 ^ PortBit%
                  OutBit% = (mlConfigOut(PortIndex%, mnBoardIndex) And 2 ^ PortBit%) = 2 ^ PortBit%
                  BitStatus% = GetBitStatus(PortIndex%, PortBit%, mnBoardIndex)
                  If Not BitStatus% = INVALIDBIT Then
                     If PortType& = AUXPORT Then
                        If BitNum% < 16 Then
                           mnAuxBit(BitNum%, mnBoardIndex) = BitStatus%
                           LastBit% = LastBit% + 1
                        End If
                     Else
                        If BitNum% < 193 Then
                           mnDIOBit(BitNum%, mnBoardIndex) = BitStatus%
                           LastBit% = LastBit% + 1
                        End If
                     End If
                  End If
                  PortBit% = PortBit% + 1
               Next BitNum%
               EndBit% = LastBit%
            End If
         End If
         DevNum% = DevNum% + 1
         'If PortNum% = 0 Then
         '   EndBit% = 0
         '   LastBit% = 0
         'End If
         PortNum% = PortNum% + 1
      Else
         ConfigItem% = DICONFIG
         ULStat = GetConfig520(InfoType%, mnBoardNum, DevNum%, ConfigItem%, ValConfig&)
         ConfigVal& = ValConfig&
         DevNum% = DevNum% + 1
         If Not (mnDIOPort(PortNum%, mnBoardIndex) = ConfigVal&) Then
            GridNeedsUpdate% = True
            Select Case ConfigVal&
               Case 1
                  mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALOUT
               Case 2
                  mnDIOPort(PortIndex%, mnBoardIndex) = DIGITALIN
               Case 3
                  mnDIOPort(PortIndex%, mnBoardIndex) = NOTCONFIGURED
            End Select
         End If
      End If
      If PortType& = AUXPORT Then
         If Not (mnDIOPort(PortNum%, mnBoardIndex) = INVALIDPORT) Then EndBit% = 0
      End If
   Next
   gnLocalErrDisp = SaveLocalErrDisp%
   geErrFlow = SaveLocalErrFlow%
   If GridNeedsUpdate% And Not ((mnFuncType = DIn) Or (mnFuncType = DOut)) Then RefreshGrid     'And (mnFuncType = DBitIn)
   mnManualRefresh = False

End Sub

Private Sub RefreshGrid()

   'to do - this will be obsoleted by RefreshDGrid when fully implemented
   If (mnLibType = MSGLIB) Or mnUseDGrid Then Exit Sub
   picAuxIO.Cls
   picDIO.Cls
   DoEvents

   If mnLibType = MSGLIB Then
      'get the number of ports on board
      NumPorts% = GetNumMsgPorts()
   End If
   For PortNum% = 0 To 32
      Select Case mnLibType
         Case UNILIB
            ULStat = GetConfig520(DIGITALINFO, mnBoardNum, PortIndex%, DICURVAL, ValConfig&)
            ULStat = GetConfig520(DIGITALINFO, mnBoardNum, PortIndex%, DIDEVTYPE, PortType&)
            SkipPort% = False
            LongPort& = GetPortFromIndex(PortNum%)
            TypeOfPort% = LongPort&
            EndBit% = 7
            If TypeOfPort% = AUXPORT Then EndBit% = 15
            If Not (PortType& = TypeOfPort%) Then
               ColorVal& = INVALIDCOLOR&
               SkipPort% = True
               BitConf% = 0
            End If
         Case MSGLIB
            If Not (PortNum% < NumPorts%) Then
               ColorVal& = INVALIDCOLOR&
               'SkipPort% = True
               BitConf% = INVALIDBIT
            Else
               NumBits% = GetNumMsgBits(PortNum%)
               EndBit% = NumBits% - 1
               ValConfig& = GetCfgMsgBits(PortNum%, NumBits%)
               If mnDIOPort(CurPort%, mnBoardNum) = MIXEDPORT Then
                  TypeOfPort% = AUXPORT
                  EndBit% = 15
                  LastBit% = NumBits% - 1
               Else
                  TypeOfPort% = FIRSTPORTA + PortNum%
                  EndBit% = 7
                  LastBit% = NumBits% - 1
               End If
            End If
      End Select
      
      For PortBit% = 0 To EndBit%
         If mnLibType = UNILIB Then
            If Not SkipPort% Then BitConf% = GetBitStatus(PortIndex%, PortBit%, mnBoardIndex)
            BitVal% = ((ValConfig& And 2 ^ PortBit%) = 2 ^ PortBit%)
         Else
            If PortBit% > LastBit% Then BitConf% = INVALIDBIT
            If Not BitConf% = INVALIDBIT Then
               BitConfiguration% = ((ValConfig& And 2 ^ PortBit%) = 2 ^ PortBit%)
               BitConf% = DIGITALOUT
               If BitConfiguration% Then BitConf% = DIGITALIN
               'to do - may be better to use port when possible here
               PortNum% = DGrid(PortBit%, PortIndex%).ULBitPort
               BitVal% = GetMsgBitVal(PortNum%, PortBit%)
            End If
         End If
         Select Case BitConf%
            Case 0
               'used to fill in AUXPORT grid
               'if no AUXPORT
               ColorVal& = INVALIDCOLOR&
            Case DIGITALOUT
               ColorVal& = OUTOFFCOLOR&
               If BitVal% Then ColorVal& = OUTONCOLOR&
            Case DIGITALIN
               ColorVal& = INOFFCOLOR&
               If BitVal% Then ColorVal& = INONCOLOR&
            Case NOTCONFIGURED
               ColorVal& = NOCONFCOLOR&
               If BitVal% Then ColorVal& = ONNOCONFCOLOR&
            Case INVALIDBIT
               ColorVal& = INVALIDCOLOR&
            Case Else
               ColorVal& = INVALIDCOLOR&
         End Select
         Col% = PortNum%
         Bit% = PortBit%
         If (PortNum% = 0) Or ((mnLibType = MSGLIB) And PortNum% = 1) Then
            If mnManualRefresh Or Not (mnAuxBitDisabled(Bit%, mnBoardIndex)) Then
               If Bit% > 7 Then
                  Bit% = Bit% - 8
                  Col% = Col% + 1
               End If
               PaintAuxGrid Col%, Bit%, ColorVal&
            End If
         Else
            Col% = PortNum% - 1
            If Bit% < 8 Then PaintDIOGrid Col%, Bit%, ColorVal&
         End If
      Next PortBit%
      If Not SkipPort% Then PortIndex% = PortIndex% + 1
   Next PortNum%

End Sub

Private Function GetPortStatus(ByVal PortNum As Integer, ByVal boardindex As Integer) As Integer

   If mlConfigIn(PortNum, boardindex) And mlConfigOut(PortNum, boardindex) Then
      'if any bits overlap, it's a non-configurable AUXPORT
      GetPortStatus = NOTCONFIGURED
   ElseIf (mlConfigIn(PortNum, boardindex) = 0) Then
      'if either inmask or outmask is all zero,
      'the entire port is configured for the opposite
      GetPortStatus = DIGITALOUT
   ElseIf (mlConfigOut(PortNum, boardindex) = 0) Then
      GetPortStatus = DIGITALIN
   Else
      'otherwise, it's a mixed bit port
      GetPortStatus = MIXEDPORT
   End If

End Function

Private Function GetBitStatus(PortNum As Integer, PortBit As Integer, boardindex As Integer) As Integer

   InBit% = (mlConfigIn(PortNum, boardindex) And 2 ^ PortBit) = 2 ^ PortBit
   OutBit% = (mlConfigOut(PortNum, boardindex) And 2 ^ PortBit) = 2 ^ PortBit
   If OutBit% Or InBit% Then
      'if either is true, bit is valid
      If OutBit% And InBit% Then
         GetBitStatus = NOTCONFIGURED
      Else
         'AUXPORT may have different no of input vs output
         'If ((mnFormType = DIGITAL_OUT) And _
         'Not OutBit%) Or ((mnFormType = DIGITAL_IN) And Not InBit%) Then
         '   GetBitStatus = INVALIDBIT
         'Else
            If InBit% = 0 Then GetBitStatus = DIGITALOUT
            If OutBit% = 0 Then GetBitStatus = DIGITALIN
         'End If
      End If
   Else
      GetBitStatus = INVALIDBIT
   End If

End Function

Private Function GetStartBit(Port As Integer) As Integer

   For CurPort% = 0 To Port - 1
      If mlPortType(CurPort%, mnBoardIndex) <> AUXPORT Then
         Mask% = (mlConfigIn(CurPort%, mnBoardIndex) Or mlConfigOut(CurPort%, mnBoardIndex))
         For TestBit% = 16 To 0 Step -1
            If 2 ^ TestBit% = (Mask% And 2 ^ TestBit%) Then
               StartBit% = StartBit% + TestBit% + 1
               Exit For
            End If
         Next
      End If
   Next CurPort%
   GetStartBit = StartBit%

End Function

Private Sub ResetOpts()

   'reset options to default
   For i% = 0 To mnuOpt.Count - 1
      mnuOpt(i%).Checked = False
   Next i%
   'mvOptions = 0
   mnDetailsNeedUpdate = True
   SetOptions
   UpdateStatBar False


End Sub

Private Function GetBitNum(ByVal Bit As Integer, ByVal Port As Integer, ByRef PortIndex As Integer) As Integer

   For TestPort% = 0 To 32
      PortType& = mlPortType(TestPort%, mnBoardIndex)
      If (TestPort% = 0) And (PortType& > 10) And Not (Port < PortType&) Then BitOffset% = (PortType& - 10) * 8
      If Port = PortType& Then Exit For
      If Not (PortType& = 0) Then
         LastValidPort& = PortType&
         LastValidIndex% = TestPort%
      End If
   Next
   If TestPort% = 33 Then
      If Port < LastValidPort& Then
         BitOffset% = 0
         LastStartBit% = 0
         Bit% = (Port - 10) * 8
      Else
         LastStartBit% = GetStartBit(LastValidIndex%)
         BitOffset% = BitOffset% + LastStartBit% + (Port - LastValidPort&) * 8
      End If
   Else
      PortIndex% = TestPort%
      StartBit% = GetStartBit(PortIndex%)
   End If
   If (mnDIOPort(0, mnBoardIndex) = INVALIDPORT) And (TestPort% < 33) Then
      PortIndex% = PortIndex% + 1
   End If
   BitNum% = BitOffset% + StartBit% + Bit%
   GetBitNum = BitNum%

End Function

Private Sub PaintAuxGrid(XStart As Integer, YStart As Integer, ColorVal As Long)

   picAuxIO.Line (XStart + 0.1, YStart + 0.2)-(XStart + 0.8, YStart + 0.85), ColorVal, BF

End Sub

Private Sub PaintDIOGrid(XStart As Integer, YStart As Integer, ColorVal As Long)

   picDIO.Line (XStart + 0.15, YStart + 0.2)-(XStart + 0.85, YStart + 0.85), ColorVal, BF

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   mnManualRefresh = True

End Sub

Private Sub mnuTrigType_Click(Index As Integer)
   
   For TrigMenu% = 0 To mnuTrigType.Count - 1
      If mnuTrigType(TrigMenu%).Checked Then TempTrigIndex% = TrigMenu%
      mnuTrigType(TrigMenu%).Checked = False
   Next TrigMenu%
   
   If Index > 15 Then
      ULStat = cbGetConfig(BOARDINFO, mnBoardNum, 0, BIDINUMDEVS, ConfigVal&)
      If ConfigVal& = 1 Then
         ULStat = cbGetConfig(DIGITALINFO, mnBoardNum, 0, DINUMBITS, ConfigVal&)
         DefaultMask& = (2 ^ ConfigVal&)
         sDefault$ = "0x" & Hex(DefaultMask& - 1)
         sDefaultPtn$ = "0x" & Hex(DefaultMask& / 2)
         If mbThresholdSet Then
            If (Not (mnLowThreshold < DefaultMask&)) Or (DefaultMask& <> mnLowThreshold) Then
               sDefault$ = "0x" & Hex(mnLowThreshold)
               sDefaultPtn$ = "0x" & Hex(mnHighThreshold)
            End If
         End If
      Else
         'handle no ports or multiple ports
         If ConfigVal& = 0 Then
            MsgBox "This device has no digital ports.", vbCritical, "Invalid Trigger Type"
            Exit Sub
         End If
      End If
      LowTH$ = InputBox("Enter mask value", "Set Digital Mask", sDefault$)
      HighTH$ = InputBox("Enter digital trigger value", "Set Digital Trigger", sDefaultPtn$)
      mnLowThreshold = ConvStringToInt(LowTH$)
      mnHighThreshold = ConvStringToInt(HighTH$)
   End If
   
   ULStat = cbSetTrigger(mnBoardNum, Index, mnLowThreshold, mnHighThreshold)
   x% = SaveFunc(Me, SetTrigger, ULStat, mnBoardNum, Index, _
      mnLowThreshold, mnHighThreshold, A5, A6, A7, A8, A9, A10, A11, 0)
   mbThresholdSet = True

   If (ULStat = BADTRIGTYPE) Or (ULStat = BADBOARDTYPE) Then
      Trig$ = mnuTrigType(mnTrigType).Caption & "."
      NewTrig$ = mnuTrigType(Index).Caption & " "
      mnuTrigType(mnTrigType).Checked = True
      MsgBox NewTrig$ & "is an invalid trigger type for this board.  Trigger reset to " & Trig$
      Exit Sub
   ElseIf ULStat <> 0 And geErrFlow <> 0 Then
      If geErrFlow = 1 Then
         mnPlot = False
         Exit Sub
      End If
   Else
      mnTrigType = Index
      mnuTrigType(mnTrigType).Checked = Not mnuTrigType(mnTrigType).Checked
      For FormNum% = 0 To Forms.Count - 1
         FName$ = Forms(FormNum%).Name
         If FName$ = "frmAnalog" Or _
            FName$ = "frmCounter" Or _
            FName$ = "frmDigital" Then
            If Not (Forms(FormNum%) Is Me) Then
               Forms(FormNum%).SetTrigParameters mnBoardNum, _
                  mnTrigType, mnLowThreshold, mnHighThreshold
            End If
         End If
      Next
   End If
   
   UpdateStatBar False

End Sub

Public Sub SelectBits(FirstBit As Long, LastBit As Long)

   'used in scripted tests
   'subsequent bit functions use the bits selected here
   
   If FirstBit = -1 Then
      'release ports
      mnNumSelBits = -1
      ReDim malBitsSelected(0)
   Else
      For SelectedBit& = FirstBit To LastBit
         'ActualPort& = GetPortFromIndex(SelectedBit&)
         For CurBit% = 0 To mnNumSelBits
            If malBitsSelected(CurBit%) = SelectedBit& Then
               AlreadySelected% = True
               Exit For
            End If
         Next CurBit%
         If Not AlreadySelected% Then
            mnNumSelBits = mnNumSelBits + 1
            ReDim Preserve malBitsSelected(mnNumSelBits)
            malBitsSelected(mnNumSelBits) = SelectedBit&
         End If
         AlreadySelected% = False
      Next SelectedBit&
   End If

End Sub

Sub SelectPorts(FirstPort As Integer, LastPort As Integer)

   'used in scripted tests
   'subsequent functions use the ports selected here
   If FirstPort = -1 Then
      'release ports
      mnNumSelPorts = -1
      ReDim malPortsSelected(0)
      ReDim malPortIndexSelected(0)
      mnMaxPortWidth = 0
   Else
      For SelectedPort% = FirstPort To LastPort
         ActualPort& = SelectedPort%
         If Not (mnLibType = MSGLIB) Then
            'ActualPort& = GetPortFromIndex(SelectedPort%)
            ActualPort& = DGrid(0, SelectedPort%).PortNum
         End If
         For CurPorts% = 0 To mnNumSelPorts
            If malPortsSelected(CurPorts%) = ActualPort& Then
               AlreadySelected% = True
               Exit For
            End If
         Next CurPorts%
         If Not AlreadySelected% Then
            mnNumSelPorts = mnNumSelPorts + 1
            ReDim Preserve malPortsSelected(mnNumSelPorts)
            ReDim Preserve malPortIndexSelected(mnNumSelPorts)
            malPortsSelected(mnNumSelPorts) = ActualPort&
            malPortIndexSelected(mnNumSelPorts) = SelectedPort%
            If malBitsInPort(SelectedPort%) > mnMaxPortWidth _
            Then mnMaxPortWidth = malBitsInPort(SelectedPort%)
         End If
         AlreadySelected% = False
      Next SelectedPort%
   End If
   
End Sub

Public Sub ConfigSelectedBits(Direction As Long)

   For Port% = 0 To mnNumSelPorts
      PortType& = malPortsSelected(Port%)
      For BitNum& = 0 To mnNumSelBits
         SelBit& = malBitsSelected(BitNum&)
         If mnLibType = MSGLIB Then
            PortIndex% = PortType&
            BitIndex% = SelBit&
            result% = SetMsgBitConfig(PortIndex%, BitIndex%, Direction&)
         Else
            ULStat = cbDConfigBit(mnBoardNum, PortType&, SelBit&, Direction)
            If SaveFunc(Me, DConfigBit, ULStat, mnBoardNum, PortType&, SelBit&, Direction, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
      Next BitNum&
   Next Port%

End Sub

Sub ConfigSelectedPorts(Direction As Integer)

   For Port% = 0 To mnNumSelPorts
      PortType& = malPortsSelected(Port%)
      If mnLibType = MSGLIB Then
         MsgPort$ = "{" & Format(PortType&, "0") & "}"
         DirMessage$ = ":DIR=IN"
         If Direction = DIGITALOUT Then DirMessage$ = ":DIR=OUT"
         DConfigMessage$ = msStatusType & MsgPort$ & DirMessage$
         MsgResult$ = MsgLibrary.SendMessage(DConfigMessage$)
         If SaveMsg(Me, "SendMessage(" & DConfigMessage$ & ")", MsgResult$) Then Exit Sub
      Else
         ULStat = cbDConfigPort(mnBoardNum, PortType&, Direction)
         If SaveFunc(Me, DConfigPort, ULStat, mnBoardNum, PortType&, Direction, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      End If
   Next
   
End Sub

Public Sub WriteSelectedBits(NumberOfBlocks As Long)

   Static stlSampleIndex As Long
   
   If NumberOfBlocks < 1 Then
      'just reset the sample index
      stlSampleIndex = 0
      Exit Sub
   End If
   Sample& = stlSampleIndex
   
   NumBits% = mnNumSelBits + 1: NumPorts% = mnNumSelPorts + 1
   TotalCount& = mlGenPoints * NumBits%
   
   If NumBits% < 1 Or NumPorts% < 1 Then
      If NumBits% < 1 Then TheProblemIs$ = "No bit range has been selected on " & msBoardName & "."
      If NumPorts% < 1 Then TheProblemIs$ = "No port range has been selected on " & msBoardName & "."
      MsgBox "Cannot write bit range. " & TheProblemIs$
      Exit Sub
   End If

   If Not (mlGenHandle = 0) Then
      'get bit values up to mnNumSelBits
      ConvResult% = GetBytesFromWinBuf(mlGenHandle, _
      mnGenDataType, NumPorts%, NumPorts%, vDataArray, Sample&)
      ConvResult% = GetBitsByPortFromArray(vDataArray, _
      mnGenDataType, NumPorts%, mnMaxPortWidth, Bits)
   Else
      MsgBox "No data has been generated for output. Exiting " & _
      "'WriteSelectedBits'.", vbOKOnly, "Error Writing Bit Values"
      Exit Sub
   End If
   If Not Me.fraScript.Visible Then
      Me.fraScript.Visible = True
      Me.fraScript.Width = Me.ScaleWidth
      Me.fraScript.Height = Me.ScaleHeight
      Me.fraScript.Top = Me.ScaleTop
      Me.fraScript.Left = Me.ScaleLeft
      Me.txtScript.Width = Me.fraScript.Width - 200
      Me.txtScript.Left = Me.fraScript.Left + 100
      Me.txtScript.Height = Me.fraScript.Height - 300
   End If
   NotDone% = True
   PrintBit& = malBitsSelected(0)
   txtScript.Text = "Output bits  " & Format(PrintBit&, "0") & " - "
   For PrintBits% = 1 To mnNumSelBits
      NextBit& = malBitsSelected(PrintBits%)
      If (NextBit& - PrintBit&) > 1 Then
         txtScript.Text = txtScript.Text & Format(PrintBit&, "0") & _
         ", " & Format(NextBit&, "0") & " - "
      End If
      PrintBit& = NextBit&
   Next
   txtScript.Text = txtScript.Text & Format(NextBit&, "0")
   
   Do
      For Port% = 0 To mnNumSelPorts
         PortType& = malPortsSelected(Port%)
         If mnLibType = MSGLIB Then
            PortName$ = "PORTINDEX" & Format(PortType&, "0")
         Else
            ULBitPort& = AUXPORT
            If PortType& > 1 Then ULBitPort& = FIRSTPORTA
            PortName$ = GetPortString(PortType&)
         End If
         If (Port% Mod 4 = 0) Then
            txtScript.Text = txtScript.Text & vbCrLf & _
            vbCrLf & PortName$ & vbCrLf & vbTab
         Else
            'txtScript.Text = txtScript.Text & "  "
         End If
         CurPortIndex& = malPortIndexSelected(Port%)
         LastInPort& = LastInPort& + malBitsInPort(CurPortIndex&)
         Bit% = 0
         Do While (Iteration& < LastInPort&) And Not (Iteration& > mnNumSelBits)
            BitNum& = malBitsSelected(Iteration&)
            BitValue& = Bits(Port%, Bit%)
            If mnLibType = MSGLIB Then
               BitComp = 0
               For PortIndex% = 0 To mlLastPortIndex
                  If BitNum& < malPortBits(PortIndex%) Then
                     Exit For
                  End If
                  BitComp = BitComp - malPortBits(PortIndex%)
               Next
               BitIndex% = (BitNum& + BitComp)
               ValToSet% = BitValue&
               BitValSet% = SetMsgBitVal(CurPortIndex&, BitIndex%, ValToSet%)
               If Not (BitValSet% = ValToSet%) Then Exit Sub
            Else
               ULStat = cbDBitOut(mnBoardNum, ULBitPort&, BitNum&, BitValue&)
               If SaveFunc(Me, DBitOut, ULStat, mnBoardNum, ULBitPort&, BitNum&, _
               BitValue&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            End If
            Bit% = Bit% + 1
            Iteration& = Iteration& + 1
            txtScript.Text = txtScript.Text & Format(BitValue&, "0")
         Loop
         txtScript.Text = txtScript.Text & "  "
      Next Port%
      BlockNum& = BlockNum& + 1
      If BlockNum& > (NumberOfBlocks - 1) Then NotDone% = False
   Loop While NotDone%
   Sample& = Sample& + NumPorts%
   stlSampleIndex = Sample&
   mlCount = stlSampleIndex
   
End Sub

Public Sub WriteSelectedPorts(NumberOfBlocks As Long)

   ForceWinAPI% = Not (mnuULBuf.Checked)
   If mnNumSelPorts < 0 Then
      MsgBox "No digital ports have been selected on " & msBoardName & ".", vbCritical, "Script Error"
      Exit Sub
   End If
   Static stlSampleIndex As Long
   Dim OutArray() As Integer
   NumPorts% = mnNumSelPorts + 1
   
   If (NumberOfBlocks < 1) Then
      'just reset the sample index
      stlSampleIndex = 0
      Exit Sub
   End If
   If Not (mlGenHandle = 0) Then
      PerChan& = mlGenPoints \ mnGenChans
      ReDim OutArray(mnGenChans - 1, PerChan& - 1)
      'ULStat = cbWinBufToArray(mlGenHandle, OutArray(0, 0), 0, mlGenPoints)
      Sr& = LoadArrayFromWinBuf(Me, mlGenHandle, OutArray(), 0, mlGenPoints, ForceWinAPI%)
   Else
      MsgBox "No data has been generated for output. Exiting " & _
      "'WriteSelectedPorts'.", vbOKOnly, "Error Writing Port Values"
      Exit Sub
   End If
   If (stlSampleIndex = PerChan&) Then stlSampleIndex = 0
   Sample& = stlSampleIndex
   'If Sample& > PerChan& - 1 Then Sample& = 0
   If Not Me.fraScript.Visible Then
      Me.fraScript.Visible = True
      Me.fraScript.Width = Me.ScaleWidth
      Me.fraScript.Height = Me.ScaleHeight
      Me.fraScript.Top = Me.ScaleTop
      Me.fraScript.Left = Me.ScaleLeft
      Me.txtScript.Width = Me.fraScript.Width - 200
      Me.txtScript.Left = Me.fraScript.Left + 100
      Me.txtScript.Height = Me.fraScript.Height - 300
   End If
   NotDone% = True
   txtScript.Text = "Port Output" & vbCrLf
   
   CompMsg$ = "DIO"
   Do
      For Port% = 0 To mnGenChans - 1
         CurPortIndex& = malPortIndexSelected(Port%)
         PortWidth& = malBitsInPort(CurPortIndex&)
         Maskval& = (2 ^ PortWidth&) - 1
         PortType& = malPortsSelected(Port%)
         DataValue% = OutArray(Port%, Sample&)
         DigitalValue$ = Format(DataValue%, "0")
         If mnLibType = MSGLIB Then
            PortString$ = "{" & Format(PortType&, "0") & "}:VALUE=" & DigitalValue$
            PortValMsg$ = CompMsg$ & PortString$
            MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
            If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Sub
            PortName$ = "PORTINDEX" & Format(PortType&, "0")
         Else
            ULStat = cbDOut(mnBoardNum, PortType&, DataValue%)
            If SaveFunc(Me, DOut, ULStat, mnBoardNum, PortType&, DataValue%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            PortName$ = GetPortString(PortType&)
         End If
         If Port% Mod 4 = 0 Then txtScript.Text = txtScript.Text & vbCrLf
         txtScript.Text = txtScript.Text & PortName$ & "  " & _
         Format(DataValue% And Maskval&, "0") & "     "
      Next
      'Sample& = Sample& + 1
      BlockNum& = BlockNum& + 1
      If BlockNum& > (NumberOfBlocks - 1) Then NotDone% = False
      txtScript.Text = txtScript.Text & vbCrLf
   Loop While NotDone%
   Sample& = Sample& + 1 'NumPorts%
   stlSampleIndex = Sample&
   mlCount = stlSampleIndex
   
End Sub

Public Sub ReadSelectedBits(NumberOfBlocks As Long)

   Dim BitData() As Integer
   
   NumBits% = mnNumSelBits + 1: NumPorts% = mnNumSelPorts + 1
   'TotalCount& = NumberOfBlocks * NumBits% '* NumPorts%
   'BitsPerPort& = NumBits% / NumPorts%
   
   If NumBits% < 1 Or NumPorts% < 1 Then
      If NumBits% < 1 Then TheProblemIs$ = "No bit range has been selected on " & msBoardName & "."
      If NumPorts% < 1 Then TheProblemIs$ = "No port range has been selected on " & msBoardName & "."
      MsgBox "Cannot read bit range. " & TheProblemIs$
      Exit Sub
   End If

   TotalCount& = NumPorts% * mnMaxPortWidth
   ReDim BitData(NumPorts% - 1, mnMaxPortWidth - 1)
   If Not Me.fraScript.Visible Then
      Me.fraScript.Visible = True
      Me.fraScript.Width = Me.ScaleWidth
      Me.fraScript.Height = Me.ScaleHeight
      Me.fraScript.Top = Me.ScaleTop
      Me.fraScript.Left = Me.ScaleLeft
      Me.txtScript.Width = Me.fraScript.Width - 200
      Me.txtScript.Left = Me.fraScript.Left + 100
      Me.txtScript.Height = Me.fraScript.Height - 300
   End If
   NotDone% = True
   PrintBit& = malBitsSelected(0)
   txtScript.Text = "Input bits  " & Format(PrintBit&, "0") & " - "
   For PrintBits% = 1 To mnNumSelBits
      NextBit& = malBitsSelected(PrintBits%)
      If (NextBit& - PrintBit&) > 1 Then
         txtScript.Text = txtScript.Text & Format(PrintBit&, "0") & _
         ", " & Format(NextBit&, "0") & " - "
      End If
      PrintBit& = NextBit&
   Next
   txtScript.Text = txtScript.Text & Format(NextBit&, "0")

   If mlHandle Then
      ULStat = cbWinBufFree(mlHandle)
      If SaveFunc(Me, WinBufFree, ULStat, mlHandle, A2, A3, _
      A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      mlHandle = 0: mnAcqDataType = vbEmpty
      mlAcqPoints = 0
   End If

   Do
      For Port% = 0 To mnNumSelPorts
         PortType& = malPortsSelected(Port%)
         If mnLibType = MSGLIB Then
            PortName$ = "PORTINDEX" & Format(PortType&, "0")
         Else
            PortName$ = GetPortString(PortType&)
            ULBitPort& = AUXPORT
            If PortType& > 1 Then ULBitPort& = FIRSTPORTA
         End If
         'Separator$ = ", "
         If (Port% Mod 4 = 0) Then
            txtScript.Text = txtScript.Text & vbCrLf & _
            vbCrLf & PortName$ & vbCrLf & vbTab
         Else
            'txtScript.Text = txtScript.Text & "  "
         End If
         CurPortIndex& = malPortIndexSelected(Port%)
         LastInPort& = LastInPort& + malBitsInPort(CurPortIndex&)
         'LastInPort& = malPortBits(Port%)
         Bit% = 0
         'For Bit% = 0 To mnNumSelBits
         Do While (Iteration& < LastInPort&) And Not (Iteration& > mnNumSelBits)
            BitNum& = malBitsSelected(Iteration&)
            If mnLibType = MSGLIB Then
               CompMsg$ = "?DIO"
               BitComp = 0
               For PortIndex% = 0 To mlLastPortIndex
                  If BitNum& < malPortBits(PortIndex%) Then
                     Exit For
                  End If
                  BitComp = BitComp - malPortBits(PortIndex%)
               Next
               BitIndex% = (BitNum& + BitComp)
               PortBitString$ = "{" & Format(PortType&, "0") & "/" & Format(BitIndex%, "0") & "}:VALUE"
               BitValMsg$ = CompMsg$ & PortBitString$
               MsgResult$ = MsgLibrary.SendMessage(BitValMsg$)
               If SaveMsg(Me, "SendMessage(" & BitValMsg$ & ")", MsgResult$) Then Exit Sub
               ValueLoc& = InStr(1, MsgResult$, "=")
               If ValueLoc& > 0 Then
                  BitValString$ = Mid(MsgResult$, ValueLoc& + 1)
                  DataValue% = Val(BitValString$)
               End If
            Else
               ULStat = cbDBitIn(mnBoardNum, ULBitPort&, BitNum&, DataValue%)
               If SaveFunc(Me, DBitIn, ULStat, mnBoardNum, ULBitPort&, BitNum&, _
               DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            End If
            If Sample& = 0 Then Sample& = (Port% * NumBits%)
            BitData(Port%, Bit%) = DataValue%
            Bit% = Bit% + 1
            Iteration& = Iteration& + 1
            txtScript.Text = txtScript.Text & Format(DataValue%, "0")
         Loop
         txtScript.Text = txtScript.Text & "  "
      Next Port%
      BlockNum& = BlockNum& + 1
      If BlockNum& > (NumberOfBlocks - 1) Then NotDone% = False
   Loop While NotDone%

   FirstPoint& = 0
   mlHandle = cbWinBufAlloc(TotalCount&)
   mnAcqDataType = vbInteger
   mlAcqPoints = TotalCount&
   ULStat = cbWinArrayToBuf(BitData(0, 0), mlHandle, FirstPoint&, TotalCount&)

End Sub

Public Sub ReadSelectedPorts(NumberOfBlocks As Long)

   ForceWinAPI% = Not (mnuULBuf.Checked)
   If mnNumSelPorts < 0 Then
      MsgBox "No digital ports have been selected on " & msBoardName & ".", vbCritical, "Script Error"
      Exit Sub
   End If
   Dim PortData() As Long
   
   ReDim PortData(mnNumSelPorts, NumberOfBlocks - 1)
   TotalCount& = NumberOfBlocks * (mnNumSelPorts + 1)
   If Not Me.fraScript.Visible Then
      Me.fraScript.Visible = True
      Me.fraScript.Width = Me.ScaleWidth
      Me.fraScript.Height = Me.ScaleHeight
      Me.fraScript.Top = Me.ScaleTop
      Me.fraScript.Left = Me.ScaleLeft
      Me.txtScript.Width = Me.fraScript.Width - 200
      Me.txtScript.Left = Me.fraScript.Left + 100
      Me.txtScript.Height = Me.fraScript.Height - 300
   End If
   NotDone% = True
   txtScript.Text = "Port Input" & vbCrLf
   
   If mlHandle Then
      'ULStat = cbWinBufFree(mlHandle)
      If BufFree(Me, mlHandle, ForceWinAPI%) Then
         mlHandle = 0: mnAcqDataType = vbEmpty
         mlAcqPoints = 0
      End If
   End If

   Do
      For Port% = 0 To mnNumSelPorts
         CurPortIndex& = malPortIndexSelected(Port%)
         PortWidth& = malBitsInPort(CurPortIndex&)
         Maskval& = (2 ^ PortWidth&) - 1
         PortType& = malPortsSelected(Port%)
         If mnLibType = MSGLIB Then
            PortName$ = "PORTINDEX" & Format(PortType&, "0")
            CompMsg$ = "?DIO"
            PortString$ = "{" & Format(PortType&, "0") & "}:VALUE"
            PortValMsg$ = CompMsg$ & PortString$
            MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
            If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Sub
            ValueLoc& = InStr(1, MsgResult$, "=")
            If ValueLoc& > 0 Then
               PortValString$ = Mid(MsgResult$, ValueLoc& + 1)
               If IsNumeric(PortValString$) Then _
               PortData(Port%, Sample&) = Val(PortValString$)
            End If
         Else
            ULStat = cbDIn(mnBoardNum, PortType&, DataValue%)
            If SaveFunc(Me, DIn, ULStat, mnBoardNum, PortType&, DataValue%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            PortData(Port%, Sample&) = IntValToULong(DataValue%)
            PortName$ = GetPortString(PortType&)
            PortValString$ = Format(DataValue% And Maskval&, "0")
         End If
         If Port% Mod 4 = 0 Then txtScript.Text = txtScript.Text & vbCrLf
         txtScript.Text = txtScript.Text & PortName$ & "  " & PortValString$ & "     "
      Next
      Sample& = Sample& + 1
      BlockNum& = BlockNum& + 1
      If BlockNum& > (NumberOfBlocks - 1) Then NotDone% = False
      txtScript.Text = txtScript.Text & vbCrLf
   Loop While NotDone%

   FirstPoint& = 0
   'mlHandle = cbWinBufAlloc32(TotalCount&)
   mlHandle = BufAlloc32(Me, TotalCount&, ForceWinAPI%)
   If Not mlHandle = 0 Then
      mnAcqDataType = vbLong
      mlAcqPoints = TotalCount&
      'CopyMemory ByVal mlHandle, PortData(0, 0), TotalCount& * 4
      ForceWinAPI% = True
      Sr& = WArrayToBuf32(Me, mlHandle, PortData(), mlAcqPoints, ForceWinAPI%)
      mlCount = TotalCount&
   End If

End Sub

Public Sub SetupData(DataHandle&, NumPoints&, Channels%, DataType&)

   mlGenHandle = DataHandle&
   mlGenPoints = NumPoints&: mnGenChans = Channels%
   WriteSelectedBits 0
   'WriteSelectedPorts 0

   If Not DataHandle& = 0 Then
      Select Case DataType&
         Case 1
            mnGenDataType = vbInteger
         Case 2
            mnGenDataType = vbLong
      End Select
   Else
      mnGenDataType = vbEmpty
   End If
   
End Sub

Public Function GetDataHandle(AcqOrGen As Integer, DataType As VbVarType, NumSamples As Long) As Long

   Select Case AcqOrGen
      Case ACQUIREDDATA
         Handle& = mlHandle
         GetDataHandle = Handle&
         DataType = mnAcqDataType
         NumSamples = mlAcqPoints
      Case GENERATEDDATA
         GetDataHandle = mlGenHandle
         DataType = mnGenDataType
         NumSamples = mlGenPoints
   End Select
   
End Function

Public Sub PlotGenData()

   If (mlGenHandle = 0) Or (mlGenPoints = 0) Then Exit Sub
   If mlGenPoints < 6 Then mnuPlotType_Click (PRINT_TEXT)
   SetPlotType mnPlotType, Me
   SetBufferType 1
   If mnGenDataType = vbLong Then SetBufferType 2
   PlotBuffer mlGenHandle, mlGenPoints, mnGenChans - 1
   mbDataGenRequired = False
   
End Sub

Public Sub PlotAcquiredData()

   If (mlHandle = 0) Or (mlCount = 0) Then Exit Sub
   DataHandle& = mlHandle
   SetBufferType 2
   PlotBuffer DataHandle&, mlCount, mnNumSelPorts
   SetBufferType 1

End Sub

Public Sub ScriptDInScan(NumSamples As Long, ScanRate As Long)

   mlCount = NumSamples
   Me.txtCount.Text = Format(NumSamples, "0")
   mfRate = ScanRate
   txtRate.Text = Format(ScanRate, "0")
   'DoDInScan
   ScanSelectedInPorts (0)
   
End Sub

Public Sub ScriptDOutScan(NumSamples As Long, ScanRate As Long)

   mlCount = NumSamples
   Me.txtCount.Text = Format(NumSamples, "0")
   mfRate = ScanRate
   txtRate.Text = Format(ScanRate, "0")
   'DoDOutScan
   ScanSelectedOutPorts (0)

End Sub

Public Sub SetEvent(ByVal EventType As Long, ByVal EventData As Long)

   mlEventType = EventType
   mlEventData = EventData
   ContinuousScan% = (InStr(1, msOpt, "CONTINUOUS") > 0)
   Select Case EventType
      Case ON_DATA_AVAILABLE
         Me.cmdStop.Visible = True
         Me.cmdStop.ENABLED = True
         If Not ContinuousScan% Then
            PlotRealChunk EventData - 1
            'mnPlotType = PLOT_CHUNKS
         End If
      Case ON_END_OF_INPUT_SCAN
         If Not ContinuousScan% Then
            PlotRealChunk EventData - 1
         End If
         Me.cmdStop = True
   End Select
   DoEvents
   
End Sub

Public Sub ConfigureEvent(ByVal EventType As Long, ByVal EventData As Long, ByVal EventParam As Long)
   
   Dim Index As Integer
   Dim CurEventType As Long
   Dim EventRequested As Boolean
   
   Set mfrmThisForm = Me
   lblStatus.BackColor = &HFFFFFF
   FormCaption$ = Me.Caption
   For i% = 0 To 6
      mnuEvent(i%).Checked = False
   Next i%
   If EventType = ALL_EVENT_TYPES Then
      ULStat = UninstallEvent(mnBoardNum, EventType)
      x% = SaveFunc(Me, DisableEvent, ULStat, mnBoardNum, EventType, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      If ULStat = 0 Then
         mlEventType = 0
         Exit Sub
      End If
   End If
   
   For Index = 0 To 6
      CurEventType = 2 ^ Index
      EventRequested = ((CurEventType And EventType) > 0)
      If EventRequested Then
         mlEventSize = EventData
         ULStat = EventEnable(mnThisInstance, mnBoardNum, _
            CurEventType, mlEventSize, mfrmThisForm)
         If SaveFunc(Me, EnableEvent, ULStat, mnBoardNum, _
            EventType, mlEventSize, FormCaption$, _
            A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mnuEvent(Index).Checked = True
         mlEventType = mlEventType Or CurEventType
         mlEventParam = EventParam
      End If
   Next
   Me.mnuContPlot.ENABLED = ((mlEventType And _
      ON_DATA_AVAILABLE) = ON_DATA_AVAILABLE)
   
End Sub

Public Sub GetEvent(EventType As Long, EventData As Long, EventParam As Long)

   EventType = mlEventType
   EventData = mlEventSize
   EventParam = mlEventParam
   DoEvents
   
End Sub

Public Function GetInstance() As Integer
   
   GetInstance = mnThisInstance
   
End Function

Private Sub AcqThread_FunctionResultsAvailable(ByVal contents As String)

   Dim Elements As Variant, A(9) As Variant, Element As Long
   mnDetailsNeedUpdate = True
   ThreadStop& = InStr(1, contents, ";")
   If ThreadStop& > 0 Then
      msStringID = "[" & Left(contents, ThreadStop& - 1) & "] "
      contents = Mid(contents, ThreadStop& + 1)
      Me.lblStatus.Caption = msStringID & msConfig & " " & msOpt
   End If
   If contents = "" Then
      MsgBox "Function not yet implemented in .Net version of Universal Test", _
      vbOKOnly, "Function Not Implemented"
      mnThreadState = TS_ERROR
      Exit Sub
   End If
   Elements = Split(contents, ",")
   FuncID% = Elements(0)
   For i% = 2 To UBound(Elements)
      A(i% - 2) = Elements(i%)
   Next
   Stat% = SaveFunc(Me, FuncID%, Elements(1), mnBoardNum, A(0), _
   A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), 0)
      
   Select Case FuncID%
      Case DConfigPort
         GetPortConfig
         ShowConfigStat
         RunOtherForms DIGITAL_IN + DIGITAL_OUT, UPDATEGRID
      Case DIn
         mvDataPoint = A(1)
      Case DInScan
         mfRateReturned = A(3)
      Case DOutScan
         mfRateReturned = A(3)
      Case GetStatus
         Select Case mnThreadState
            Case TS_PRECALL
               If Elements(1) = OVERRUN Then
                  'should not occur here - this is bug #920 - clear it with stop background
                  functionToCall$ = GetCallAsString(StopBackground, mnBoardNum, mlStatusType)
                  'AcqThread.CallULFunction functionToCall$
               End If
            Case TS_FIRSTCHECK
               If A(0) = IDLE Then
                  Me.cmdStop.Visible = False
                  ULStat& = Elements(1)
                  Status% = A(0)
                  CurCount& = A(1)
                  CurIndex& = A(2)
                  x% = CheckStatus(ULStat&, Status%, CurCount&, CurIndex&)
               Else
                  mnPlot = False
                  mnThreadState = TS_TMRCHK
                  ULStat& = Elements(1)
                  Status% = A(0)
                  CurCount& = A(1)
                  CurIndex& = A(2)
                  x% = CheckStatus(ULStat&, Status%, CurCount&, CurIndex&)
               End If
            Case TS_TMRCHK
               ULStat& = Elements(1)
               Status% = A(0)
               CurCount& = A(1)
               CurIndex& = A(2)
               x% = CheckStatus(ULStat&, Status%, CurCount&, CurIndex&)
            Case Else
               ULStat& = Elements(1)
               Status% = A(0)
               CurCount& = A(1)
               CurIndex& = A(2)
               x% = CheckStatus(ULStat&, Status%, CurCount&, CurIndex&)
         End Select
   End Select
   mnThreadState = TS_IDLE

End Sub

Private Sub AcqThread_FunctionCalled(ByVal FuncName As String)

   Dim Elements As Variant, A(9) As Variant, Element As Long
   ThreadStop& = InStr(1, FuncName, ";")
   If ThreadStop& > 0 Then
      msStringID = "[" & Left(FuncName, ThreadStop& - 1) & "] "
      FuncName = Mid(FuncName, ThreadStop& + 1)
   End If
   
   Elements = Split(FuncName, ",")
   If UBound(Elements) < 2 Then Exit Sub
   FuncID% = Elements(0)
   Select Case FuncID%
      Case DInScan
      Case GetStatus
         For i% = 2 To UBound(Elements)
            A(i% - 2) = Elements(i%)
         Next
         Stat% = SaveFunc(Me, FuncID%, Elements(1), mnBoardNum, A(0), _
         A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), 0)
         If A(0) = RUNNING Then
            ULStat& = Elements(1)
            Status% = A(0)
            CurCount& = A(1)
            CurIndex& = A(2)
            x% = CheckStatus(ULStat&, Status%, CurCount&, CurIndex&)
         Else
            Me.cmdStop.Visible = False
         End If
   End Select
   If Not mnThreadState = TS_FIRSTCHECK Then mnThreadState = TS_IDLE

End Sub

Private Function CheckStatus(ULStat As Long, Status%, CurCount&, CurIndex&) As Integer
    
   Static Do2ndHalf%, DoFirstHalf%
   
   If mnCancel Or gnCancel Then
      tmrGoLoop.ENABLED = False
      cmdGo.FontBold = True
      cmdStop = True
      gnCancel = False
      mnCancel = False
      If mnSetActive Then SetContActive False
      Element& = 0
      'If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
      'If DevActive% Then
      '   BGOp$ = "Background task in progress..."
      'Else
      '   BGOp$ = "No background task running..."
      'End If
      Stat$ = msStringID & BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
      lblStatus.Caption = Stat$
      If mnuStopBG.Checked Then
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
         If (Not gnScriptSave) Or (ULStat <> 0) Then
            If SaveFunc(Me, StopBackground, ULStat, _
               mnBoardNum, mlStatusType, A3, A4, A5, _
               A6, A7, A8, A9, A10, A11, 0) Then Exit Function
         End If
         cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      DoEvents
      mnPlot = True
      Exit Function
   End If
   
   If mnuContPlot.Checked Then
      BufferSplit& = mlCount \ 2
      If (Not Do2ndHalf%) And (Not DoFirstHalf%) Then
         mlFirstPoint = 0
         If CurIndex& > BufferSplit& Then
            DoFirstHalf% = True
            mnPlot = True
         End If
      End If
      If DoFirstHalf% Then
         If CurIndex& > BufferSplit& Then
            DoFirstHalf% = False
            Do2ndHalf% = True
            DisplayData
            mlFirstPoint = (BufferSplit&) - ((BufferSplit&) Mod (mnLastChan - mnFirstChan + 1))
         End If
      End If
      If Do2ndHalf% Then
         If CurIndex& < BufferSplit& Then
            DoFirstHalf% = True
            Do2ndHalf% = False
            DisplayData
            mlFirstPoint = 0
         End If
      End If
   End If
   
   If Status% = 0 Then
      If mnuStopBG.Checked Then
         If Not mnThreading Then
            If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, A2, A3, A4, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
         End If
         If mnuStopBG.Checked Then
            If mnThreading Then
               functionToCall$ = GetCallAsString(StopBackground, mnBoardNum, mlStatusType)
               'AcqThread.CallULFunction functionToCall$
            Else
               ULStat = StopBackground520(mnBoardNum, mlStatusType)
               If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
            End If
            If Not mnuTimer.Checked Then cmdStop.Visible = False
            Status% = 0
            Do2ndHalf% = False
            DoFirstHalf% = False
            tmrCheckStatus.ENABLED = False
            DoEvents
            mnPlot = True
            If Not (mnFormType = DIGITAL_OUT) Then DisplayData
         End If
      End If
   End If

End Function

Public Function GetStaticOption() As Long

   GetStaticOption = mlStaticOpt
   
End Function

Public Sub SetStaticOption(NewOption As Long)

   If NewOption = 0 Then
      mlStaticOpt = NewOption
   Else
      mlStaticOpt = mlStaticOpt Or NewOption
   End If
   
End Sub

Public Sub ShowBufferInfo(ByVal Source As Integer)

   If IsEmpty(mlHandle) Or (mlHandle = 0) Then
      'check if there's an output buffer
      If Not ((mlGenHandle = 0) Or (mlGenPoints = 0)) Then
         SetPlotType mnPlotType, Me
         PlotGenData
         If ((mvOptions And HIGHRESRATE) = HIGHRESRATE) Then
            RateVal = mfRateReturned / 1000 & "Hz"
         Else
            RateVal = mfRateReturned / 1000 & "kHz"
         End If
         BoardName = mnuBoard(mnBoardIndex).Caption
         lblStatus.Caption = msStringID & msConfig & " " & msOpt
         SetDetails mnFuncType, mnPortNum, mlCount, mfRateReturned, _
            RateVal, Range%, BoardName, mlPTCountReturn
         PrintMain "Buffer handle = 0x" & Hex$(mlGenHandle) & _
            ",  buffer size = " & mlGenPoints
         Exit Sub
      End If
   End If
   If mnDataChange Then ResetData
   mnDetailsNeedUpdate = True
   DoEvents
   
   BufSize& = mlCount
   BufferStat% = CheckBuffer(mlHandle, BufSize&)
   Select Case BufferStat%
      Case INVALIDBUFFER
         MsgBox "Windows buffer no longer valid", , "Buffer Status"
      Case BUFFERTOOSMALL
         MsgBox "Windows buffer too small (size = " & BufSize& & ", data requested = " & mlCount * 2 & ")", , "Buffer Status"
      Case NOBUFFER
         MsgBox "No Windows buffer has been allocated", , "Buffer Status"
      Case Else
         PrintMain "Buffer handle = 0x" & Hex$(mlHandle) & ",  buffer size = " & BufSize&
         mnPlot = True
   End Select
   DisplayData

End Sub

Public Sub InitForm(FunctionInit As Integer)

   If mnFormType = DIGITAL_IN Then
      mnuFuncArray(CB_DIN).Caption = "cbDIn()"
      mnuFuncArray(CB_DIN32).Caption = "cbDIn32()"
      mnuFuncArray(DIN_ARRAY).Caption = "cbDInArray()"
      mnuFuncArray(BIT_IN).Caption = "cbDBitIn()"
      mnuFuncArray(DIN_SCAN).Caption = "cbDInScan()"
   Else
      mnuFuncArray(CB_DOUT).Caption = "cbDOut()"
      mnuFuncArray(CB_DOUT32).Caption = "cbDOut32()"
      mnuFuncArray(DOUT_ARRAY).Caption = "cbDOutArray()"
      mnuFuncArray(BIT_OUT).Caption = "cbDBitOut()"
      mnuFuncArray(DOUT_SCAN).Caption = "cbDOutScan()"
   End If
   Select Case mnFuncType
      Case DIn, DIn32, DBitIn, DOut, DOut32, _
         DBitOut, DInArray, DOutArray
            msStatusType = "DIO"
      Case DInScan, DOutScan
         msStatusType = "DIOSCAN"
   End Select
   If Not mnFormInitialized Then
      mnuBoard_Click (0)
      mnuFuncArray_Click (FunctionInit)
   End If
   If mnLibType = MSGLIB Then
      If Not mnMessaging Then Exit Sub
      NumProps% = GetDIOProps(msBoardName, MsgLibrary, PropList)
      UpdateFormProps PropList, NumProps%
   End If
   mnFormInitialized = True
   
End Sub

Private Function GetCfgMsgBits(PortIndex As Integer, NumBits As Integer) As Long
      
   CompMsg$ = "?DIO"
   PortString$ = "{" & Format(PortIndex, "0") & "}"
   BitCfgMsg$ = CompMsg$ & PortString$ & ":DIR"
   MsgResult$ = MsgLibrary.SendMessage(BitCfgMsg$)
   If SaveMsg(Me, "SendMessage(" & BitCfgMsg$ & ")", MsgResult$) Then Exit Function
   ValueLoc& = InStr(1, MsgResult$, ":")
   If ValueLoc& > 0 Then
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         MsgBitConfig$ = Mid(MsgResult$, ValueLoc& + 1)
         If IsNumeric(MsgBitConfig$) Then
            PortConfig& = Val(MsgBitConfig$)
            mnDIOPort(CurPort%, mnBoardNum) = MIXEDPORT
         Else
            PortConfig& = (2 ^ NumBits) - 1
            If MsgBitConfig$ = "OUT" Then PortConfig& = 0
         End If
      End If
   End If
   GetCfgMsgBits = PortConfig&

End Function

Private Function SetMsgPortVal(PortIndex As Integer, PortValue As Long) As Long

   CompMsg$ = "DIO"
   PortString$ = "{" & Format(PortIndex, "0") & "}:VALUE=" & Format(PortValue, "0")
   PortValMsg$ = CompMsg$ & PortString$
   MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
   x% = SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$)
   PortVal& = GetMsgPortVal(PortIndex)
   SetMsgPortVal = PortVal&

End Function

Private Function GetMsgPortVal(PortIndex As Integer) As Long

   CompMsg$ = "?DIO"
   PortString$ = "{" & Format(PortIndex, "0") & "}:VALUE"
   PortValMsg$ = CompMsg$ & PortString$
   MsgResult$ = MsgLibrary.SendMessage(PortValMsg$)
   If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Function
   ValueLoc& = InStr(1, MsgResult$, "=")
   If ValueLoc& > 0 Then
      PortValString$ = Mid(MsgResult$, ValueLoc& + 1)
      PortVal& = Val(PortValString$)
   End If
   GetMsgPortVal = PortVal&

End Function

Private Function GetMsgBitVal(ByVal PortNum As Integer, ByVal BitNum As Integer) As Integer

   If mbDeviceInUse Then Exit Function
   If mnLibType = MSGLIB Then
      CompMsg$ = "?DIO"
      'Direction% = DGrid(BitIndex, PortIndex).BitConf
      'PortNum% = DGrid(BitIndex, PortIndex).PortNum
      'BitNum% = DGrid(BitIndex, PortIndex).ULBitNum
      PortBitString$ = "{" & Format(PortNum%, "0") & "/" & Format(BitNum%, "0") & "}:VALUE"
      BitValMsg$ = CompMsg$ & PortBitString$
      MsgResult$ = MsgLibrary.SendMessage(BitValMsg$)
      If SaveMsg(Me, "SendMessage(" & BitValMsg$ & ")", MsgResult$) Then Exit Function
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         BitValString$ = Mid(MsgResult$, ValueLoc& + 1)
         BitVal% = Val(BitValString$)
      End If
   Else
      'turn off error reporting, since some devices don't allow readback of outputs
      ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
      SaveError% = geErrFlow
      SaveDisp% = gnLocalErrDisp
      geErrFlow = 0
      gnLocalErrDisp = False
      'If SaveFunc(Me, ErrHandling, ULStat, DONTPRINT, DONTSTOP, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
      ULStat = cbDBitIn(mnBoardNum, PortNum%, BitNum%, BitVal%)
      If (ULStat = DEADDIGITALDEV) Or (ULStat = SOCKETDISCONNECTED) Then
         ResetDGrid False
         InitDGrid
         ULStat = cbDBitIn(mnBoardNum, PortNum%, BitNum%, BitVal%)
      End If
      If SaveFunc(Me, DBitIn, ULStat, mnBoardNum, PortNum%, _
      BitNum%, BitVal%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
      Dim MResult As VbMsgBoxResult
      If ULStat = NETDEVINUSEBYANOTHERPROC Then
         MResult = MsgBox("Device is in use by another process. Abort initialization?", _
            vbYesNo, "Device In Use")
         If MResult = vbYes Then
            mbDeviceInUse = True
            ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
            geErrFlow = SaveError%
            gnLocalErrDisp = SaveDisp%
            Exit Function
         End If
      End If
      If Not (ULStat = 0) Then
         BitVal% = -1
      End If
      ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
      geErrFlow = SaveError%
      gnLocalErrDisp = SaveDisp%
      'If SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, gnErrHandling, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   End If
   GetMsgBitVal = BitVal%

End Function

Private Function GetNumMsgBits(PortIndex As Integer) As Integer

   CompMsg$ = "?DIO"
   PortString$ = "{" & Format(PortIndex, "0") & "}"
   NumBitsMsg$ = CompMsg$ & PortString$
   MsgResult$ = MsgLibrary.SendMessage(NumBitsMsg$)
   If SaveMsg(Me, "SendMessage(" & NumBitsMsg$ & ")", MsgResult$) Then Exit Function
   ValueLoc& = InStr(1, MsgResult$, "=")
   If ValueLoc& > 0 Then
      NumBitsString$ = Mid(MsgResult$, ValueLoc& + 1)
      NumBits% = Val(NumBitsString$)
   End If
   GetNumMsgBits = NumBits%

End Function

Private Function GetNumMsgPorts() As Integer

   CompMsg$ = "?DIO"
   MsgResult$ = MsgLibrary.SendMessage(CompMsg$)
   If SaveMsg(Me, "SendMessage(" & CompMsg$ & ")", MsgResult$) Then Exit Function
   ValueLoc& = InStr(1, MsgResult$, "=")
   If ValueLoc& > 0 Then
      NumPortsString$ = Mid(MsgResult$, ValueLoc& + 1)
      NumPorts% = Val(NumPortsString$)
   End If
   GetNumMsgPorts = NumPorts%

End Function

Private Sub UpdateFormProps(PropList As Variant, ListSize As Integer)

   'ResetOpts
   For i% = 0 To ListSize
      ListedProp$ = PropList(i%)
      ParseProp = Split(ListedProp$, "=")
      If UBound(ParseProp) > 0 Then
         Prop$ = ParseProp(0)
         PropVal$ = ParseProp(1)
         CurProp$ = Prop$
         BitSpec& = InStr(1, Prop$, "/")
         If BitSpec& = 0 Then
            PortSpecStart& = InStr(1, Prop$, "{")
            PortSpecEnd& = InStr(1, Prop$, "}")
            PSLength& = (PortSpecEnd& - PortSpecStart&)
            If PSLength& > 1 Then
               PortString$ = Mid(Prop$, PortSpecStart& + 1, PSLength& - 1)
               If Len(Prop$) > PortSpecEnd& Then PropSpec$ = Mid(Prop$, PortSpecEnd& + 1)
               CurProp$ = Left(Prop$, PortSpecStart& - 1) & "{port}" & PropSpec$
            End If
         End If
         Select Case CurProp$
            Case "DIO"  'number of ports
               mnLastPort = Val(PropVal$)
            Case "DIO{port}"  'number of bits in port
               BitsInPort% = Val(PropVal$)
            Case "DIO{port}:DIR"  'configuration of port (IN, OUT, or bitfield)
               If IsNumeric(PropVal$) Then
                  'value is a bitfield specifying input or output
                  CurPort% = Val(PortString$)
                  BitMask& = Val(PropVal$)
                  mnDIOPort(CurPort%, mnBoardNum) = MIXEDPORT
                  For Bit% = 0 To BitsInPort% - 1
                     If (2 ^ Bit% And BitMask&) = 2 ^ Bit% Then
                        mnAuxBit(Bit%, mnBoardNum) = DIGITALIN
                     Else
                        mnAuxBit(Bit%, mnBoardNum) = DIGITALOUT
                     End If
                  Next
               Else
                  If PropVal$ = "OUT" Then
                     Direction% = DIGITALOUT
                  Else
                     Direction% = DIGITALIN
                  End If
               End If
            Case "TRIG"
               If PropVal$ = "ENABLE" Then
                  OptionVal& = GetOptionCodeFromMsg(Prop$, OptionIndex%)
                  If Not OptionIndex% = -1 Then mnuOpt_Click (OptionIndex%)
               End If
            Case "EXTPACER"   'EXTSYNC
               If PropVal$ = "ENABLE" Then
                  OptionVal& = GetOptionCodeFromMsg(Prop$, OptionIndex%)
                  If Not OptionIndex% = -1 Then mnuOpt_Click (OptionIndex%)
               End If
            Case "RATE"
               If Not (txtRate.Text = PropVal$) Then
                  mnDetailsNeedUpdate = True
                  mfRate = Val(PropVal$)
                  Me.txtRate.Text = Format(mfRate, "0.0####")
               End If
            Case "SAMPLES"
               If Not (txtCount.Text = PropVal$) Then
                  Me.txtCount.Text = PropVal$
                  mlCount = PropVal$
                  mnDetailsNeedUpdate = True
               End If
            Case "LOWCHAN"
               If IsNumeric(PropVal$) Then
                  'Me.hsbLowChan.value = Val(PropVal$)
               End If
            Case "HIGHCHAN"
               If IsNumeric(PropVal$) Then
                  'Me.hsbHighChan.value = Val(PropVal$)
               End If
            Case "RANGE{0}"
               mnRange = GetRangeCodeFromMsg(PropVal$)
         End Select
      End If
   Next
   UpdateStatBar False
   
End Sub

Sub WritePropsToDevice()

   SamplesString$ = txtCount.Text
   MessageUpdate$ = msStatusType & ":SAMPLES=" & SamplesString$
   GoSub DeliverMessage
   'msMsgRange = GetMsgRange(mnRange)
   'RangeString$ = "RANGE=" & msMsgRange
   'For Chan% = mnFirstChan To mnLastChan
   '   ChanString$ = Format(Chan%, "0")
   '   MessageUpdate$ = "AI{" & ChanString$ & "}:" & RangeString$
   '   GoSub DeliverMessage
   'Next
   
   Select Case mnFuncType
      Case AIn, ATrig, TIn, TInScan, VIn, AIn32
      Case AInScan
         SamplesString$ = txtCount.Text
         MessageUpdate$ = msStatusType & ":SAMPLES=" & SamplesString$
         GoSub DeliverMessage
         MessageUpdate$ = msStatusType & ":LOWCHAN=" & Format(mnFirstChan, "0")
         GoSub DeliverMessage
         MessageUpdate$ = msStatusType & ":HIGHCHAN=" & Format(mnLastChan, "0")
         GoSub DeliverMessage
         RateString$ = Left(txtRate.Text, 9)
         MessageUpdate$ = msStatusType & ":RATE=" & RateString$
         GoSub DeliverMessage
      Case APretrig
      Case FileAInScan, FilePretrig
      Case FileRead
   End Select
   mnRefreshProps = False
   Exit Sub

DeliverMessage:
   MsgResult$ = MsgLibrary.SendMessage(MessageUpdate$)
   If SaveMsg(Me, "SendMessage(" & MessageUpdate$ & ")", MsgResult$) Then Exit Sub
   Return
   
End Sub

Public Sub SetLibType(ByVal LibType As Integer)

   mnuLibrary_Click (LibType)
   
End Sub

Public Sub InitDGrid()

   GoSub ClearDGrid
   ResetDGrid True
   mlTotalBits = 0
   If mnLibType = MSGLIB Then
      For Port& = 0 To mnLastPort - 1
         'get # bits in port
         PropMsg$ = "?DIO{" & Format(Port&, "0") & "}" 'number of bits in port
         MsgResult$ = MsgLibrary.SendMessage(PropMsg$)
         x% = SaveMsg(Me, "SendMessage(" & PropMsg$ & ")", MsgResult$) 'Then
         BitReturn = Split(MsgResult$, "=")
         If UBound(BitReturn) > 0 Then NumBits& = Val(BitReturn(1))
         PropMsg$ = "?DIO{" & Format(Port&, "0") & "}:DIR" 'port direction
         MsgResult$ = MsgLibrary.SendMessage(PropMsg$)
         x% = SaveMsg(Me, "SendMessage(" & PropMsg$ & ")", MsgResult$) 'Then
         DirReturn = Split(MsgResult$, "=")
         If UBound(DirReturn) > 0 Then PortDir$ = DirReturn(1)
         If IsNumeric(PortDir$) Then
            ParseBitField% = True
         Else
            ParseBitField% = False
            Select Case PortDir$
               Case "IN"
                  CurDir& = DIGITALIN
               Case "OUT"
                  CurDir& = DIGITALOUT
               Case Else
                  CurDir& = NOTCONFIGURED
            End Select
         End If
         For Bit& = 0 To NumBits& - 1
            If Bit& = 8 Then
               GridBit& = 0
               GridPortOffset& = GridPortOffset& + 1
            ElseIf Bit& > 0 Then
               GridBit& = GridBit& + 1
            Else
               GridBit& = 0
            End If
            DGrid(GridBit&, Port& + GridPortOffset&).PortNum = Port&
            DGrid(GridBit&, Port& + GridPortOffset&).BitNum = Bit&
            DGrid(GridBit&, Port& + GridPortOffset&).PortName = "DIO{" & Format(Port&, "0") & "}"
            If ParseBitField% Then
               PortField& = Val(PortDir$)
               CurDir& = DIGITALOUT
               If ((2 ^ Bit&) And PortField&) = (2 ^ Bit&) Then CurDir& = DIGITALIN
               DGrid(GridBit&, Port& + GridPortOffset&).PortType = MIXEDPORT
            Else
               DGrid(GridBit&, Port& + GridPortOffset&).PortType = CurDir&
            End If
            DGrid(GridBit&, Port& + GridPortOffset&).BitConf = CurDir&
         Next Bit&
      Next
   Else
      InfoType% = BOARDINFO
      ConfigItem% = BIDINUMDEVS
      ULStat = GetConfig520(InfoType%, mnBoardNum, Port&, ConfigItem%, LastPort&)
      mnLastPort = LastPort&
      For Port& = 0 To mnLastPort - 1
         InfoType% = DIGITALINFO
         ConfigItem% = DIDEVTYPE
         ULStat = GetConfig520(InfoType%, mnBoardNum, Port&, ConfigItem%, PortType&)
         ConfigItem% = DIINMASK
         ULStat = GetConfig520(InfoType%, mnBoardNum, Port&, ConfigItem%, ValInConfig&)
         If ValInConfig& < 0 Then
            IConfig# = ValInConfig& + 2 ^ 31
            InSignBit& = 1
         Else
            IConfig# = ValInConfig&
            InSignBit& = 0
         End If
         ConfigItem% = DIOUTMASK
         ULStat = GetConfig520(InfoType%, mnBoardNum, Port&, ConfigItem%, ValOutConfig&)
         If ValOutConfig& < 0 Then
            OConfig# = ValOutConfig& + 2 ^ 31
            OutSignBit& = 1
         Else
            OConfig# = ValOutConfig&
            OutSignBit& = 0
         End If
         ConfigItem% = DINUMBITS
         ULStat = GetConfig520(InfoType%, mnBoardNum, _
            Port&, ConfigItem%, ConfigBits&)
         MisMatch% = False
         If Not (IConfig# And OConfig#) = 0 Then
            'if any bits overlap, it's a non-configurable AUXPORT
            PortStatus& = NOTCONFIGURED
            BitsIn& = Log(IConfig# + 1) / Log(2)
            BitsOut& = Log(OConfig# + 1) / Log(2)
            BitsIO& = BitsIn&
            If mnFormType = DIGITAL_OUT Then BitsIO& = BitsOut&
         ElseIf IConfig# = 0 Then
            'if either inmask or outmask is all zero,
            'the entire port is configured for the opposite
            PortStatus& = DIGITALOUT
            BitsIO& = (Log(OConfig# + 1) / Log(2)) + OutSignBit&
            If Not (BitsIO& = ConfigBits&) Then MisMatch% = True
         ElseIf OConfig# = 0 Then
            PortStatus& = DIGITALIN
            BitsIO& = Log(IConfig# + 1) / Log(2)
            If Not (BitsIO& = ConfigBits&) Then MisMatch% = True
         Else
            'otherwise, it's a mixed bit port
            PortStatus& = MIXEDPORT
            AllBits& = (IConfig# Or OConfig#)
            BitsIO& = Log(AllBits& + 1) / Log(2)
            If Not (BitsIO& = ConfigBits&) Then MisMatch% = True
            ParseBitField% = True
         End If
         mlTotalBits = mlTotalBits + BitsIO&
         'special cases
         Select Case msBoardName
            Case "USB-ERB08", "USB-ERB24"
                PortStatus& = NOTCONFIGURED
         End Select
         If MisMatch% Then MsgBox "GetConfig returned a different number of bits (" & _
         Format(ConfigBits&, "0") & ") than DIOUTMASK/DIINMASK (" & _
         ")", vbInformation, "Number of Bits Conflict"
         If Port& = 0 And PortType& = 12 Then BitOffset& = 16
         For Bit& = 0 To BitsIO& - 1
            If Bit& > 0 Then
               If Bit& Mod 8 = 0 Then
                  GridBit& = 0
                  GridPortOffset& = GridPortOffset& + 1
               Else
                  GridBit& = GridBit& + 1
               End If
            Else
               GridBit& = 0
            End If
            BitDir& = PortStatus&
            If PortStatus& = NOTCONFIGURED Then
               'see if this bit is not Input/Output
               InVal& = (2 ^ Bit& And IConfig#)
               OutVal& = (2 ^ Bit& And OConfig#)
               If Not (InVal& = OutVal&) Then
                  BitDir& = DIGITALOUT
                  If OutVal& = 0 Then BitDir& = DIGITALIN
               End If
            End If
            'special cases
            Select Case msBoardName
               Case "USB-ERB08", "USB-ERB24"
                   BitDir& = NOTCONFIGURED
            End Select
            If PortStatus& = MIXEDPORT Then
               BitDir& = DIGITALIN
               OutVal& = (2 ^ Bit& And OConfig#)
               If Not (OutVal& = 0) Then BitDir& = DIGITALOUT
            End If
            DGrid(GridBit&, Port& + GridPortOffset&).PortType = PortStatus& 'input, output or other
            DGrid(GridBit&, Port& + GridPortOffset&).PortNum = PortType&    'AUXPORT, FIRSTPORTA, etc
            DGrid(GridBit&, Port& + GridPortOffset&).BitNum = Bit& + BitOffset& 'bit within PortType
            DGrid(GridBit&, Port& + GridPortOffset&).PortName = GetPortString(PortType&)
            CurDir& = PortStatus&
            If ParseBitField% Then
               CurDir& = DIGITALOUT
               If ((2 ^ Bit&) And IConfig#) = (2 ^ Bit&) Then CurDir& = DIGITALIN
            End If
            If (Port& > 0) And (Bit& = 0) Then
               'check if transitioning from AUXPORT to FIRSTPORTA or vice-versa
               'if so, reset bitport to zero
               If DGrid(Bit&, Port&).PortNum > 9 Then
                  If DGrid(Bit&, Port& - 1).PortNum = 1 Then PortBit& = 0
               Else
                  If DGrid(Bit&, Port& - 1).PortNum > 9 Then PortBit& = 0
               End If
            End If
            BitPort& = AUXPORT
            BitPortName$ = "AUXPORT"
            If DGrid(GridBit&, Port& + GridPortOffset&).PortNum > 9 Then
               BitPort& = FIRSTPORTA
               BitPortName$ = "FIRSTPORTA"
            'ElseIf DGrid(GridBit&, Port& + GridPortOffset&).PortNum > 1 Then
            End If
            DGrid(GridBit&, Port& + GridPortOffset&).PortIndex = Port&
            DGrid(GridBit&, Port& + GridPortOffset&).BitConf = BitDir&
            DGrid(GridBit&, Port& + GridPortOffset&).ULBitNum = PortBit& + BitOffset&
            DGrid(GridBit&, Port& + GridPortOffset&).ULBitPort = BitPort&
            DGrid(GridBit&, Port& + GridPortOffset&).ULBitPortName = BitPortName$
            PortBit& = PortBit& + 1
         Next Bit&
      Next Port&
   End If
   'RefreshDGrid
   Exit Sub
   
ClearDGrid:
   For Row% = 0 To 7
      For Col% = 0 To mnLastColumn - 1
         DGrid(Row%, Col%).BitConf = INVALIDBIT
         DGrid(Row%, Col%).PortIndex = -1
         DGrid(Row%, Col%).PortType = -1
         DGrid(Row%, Col%).PortNum = -1
      Next Col%
   Next Row%
   Return
   
End Sub

Private Sub RefreshDGrid()

   Dim ReadOutput As Boolean
   
   ReadOutput = False
   Select Case mnReadType
      Case 1
         ReadOutput = True
      Case 2
         ULStat = SetConfig520(DIGITALINFO, mnBoardNum, DevNum%, DIDISABLEDIRCHECK, ValConfig&)
         If Not SaveFunc(Me, SetConfig, ULStat, DIGITALINFO, mnBoardNum, DevNum%, _
            DIDISABLEDIRCHECK, ValConfig&, A6, A7, A8, A9, A10, A11, 0) _
            Then ReadOutput = True
      Case Else
         ReadOutput = False
   End Select
   
   StartPos& = 1
   'Do
   '
   '   CPos& = InStr(StartPos&, msReadType, ",")
   '   StartPos& = CPos& + 1
   'Loop While CPos& > 0
   PortsListed& = Len(msReadType)
   
   For Col% = 0 To mnLastColumn - 1
      If Col% < PortsListed& Then
         Characteristic& = Val(Mid(msReadType, Col% + 1, 1))
      Else
         Characteristic& = 0
      End If
      ReadOutput = Not (Characteristic& = 0)
      For Row% = 0 To 7
         If Not DGrid(Row%, Col%).BitConf = INVALIDBIT Then
            SpecialCase% = (msBoardName = "USB-2416-4AO")
            SpecialCase% = SpecialCase% Or (msBoardName = "USB-2416")
            PortConf& = DGrid(Row%, Col%).PortType
            'to do - make sure this works with following commented 1/4/2010
            If (Not ReadOutput) And ((PortConf& = NOTCONFIGURED) And (mnFormType = DIGITAL_OUT)) Then
               'can't get value on in/out ports
               BitVal% = 0
            Else
               If mnLibType = MSGLIB Then
                  BitNum% = DGrid(Row%, Col%).BitNum
                  PortNum% = DGrid(Row%, Col%).PortNum
               Else
                  BitNum% = DGrid(Row%, Col%).ULBitNum
                  PortNum% = DGrid(Row%, Col%).ULBitPort
               End If
               BitVal% = GetMsgBitVal(PortNum%, BitNum%)
            End If
            Direction& = DGrid(Row%, Col%).BitConf
            If mbDeviceInUse Then Exit For
            'BitsRead% = BitsRead% + 1
            If BitVal% > -1 Then 'Exit For
               Select Case Direction&
                  Case DIGITALOUT
                     If Characteristic& = 2 Then
                        'unreliable values read from output
                        ColorVal& = NOCONFCOLOR&
                        If Not BitVal% = 0 Then ColorVal& = ONNOCONFCOLOR&
                     Else
                        ColorVal& = OUTOFFCOLOR&
                        If Not BitVal% = 0 Then ColorVal& = OUTONCOLOR&
                     End If
                  Case DIGITALIN
                     ColorVal& = INOFFCOLOR&
                     If Not BitVal% = 0 Then ColorVal& = INONCOLOR&
                  Case NOTCONFIGURED
                     'both input and output
                     If mnFormType = DIGITAL_IN Then
                        ColorVal& = INOFFCOLOR&
                        If Not BitVal% = 0 Then ColorVal& = INONCOLOR&
                     Else
                        If mbReadInvert Then
                           If BitVal% = 0 Then
                              BitVal% = 1
                           Else
                              BitVal% = 0
                           End If
                        End If
                        ColorVal& = OUTOFFCOLOR&
                        If Not BitVal% = 0 Then ColorVal& = OUTONCOLOR&
                     End If
                  Case Else
               End Select
            Else
               ColorVal& = NOCONFCOLOR&
            End If
            PaintDIOGrid Col%, Row%, ColorVal&
         End If
      Next Row%
      If mbDeviceInUse Then Exit For
   Next Col%
   
End Sub

Private Sub ResetDGrid(ByVal BoardChange As Integer)

   For Col% = 0 To mnLastColumn - 1
      For Row% = 0 To 7
         PaintDIOGrid Col%, Row%, INVALIDCOLOR&
         If BoardChange Then DGrid(Row%, Col%).BitConf = INVALIDBIT
      Next Row%
   Next Col%
   
End Sub

Private Sub ScanSelectedInPorts(NumberOfBlocks As Long)

   ForceWinAPI% = Not (mnuULBuf.Checked)
   UseLibrary% = mnuULBuf.Checked
   If mnNumSelPorts < 0 Then
      MsgBox "No digital ports have been selected.", vbCritical, "Script Error"
      Exit Sub
   End If
   
   PlotThis% = mnPlot
   mnPlot = False
   If mlHandle Then
      If UseLibrary% Then
         ULStat = cbWinBufFree(mlHandle)
         If SaveFunc(Me, WinBufFree, ULStat, mlHandle, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mlHandle = 0: mnAcqDataType = vbEmpty
         LongHandle& = mlHandle: mlAcqPoints = 0
         PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
      Else
         If FreeMemory(mlHandle) Then
            mlHandle = 0: mlAcqPoints = 0
            mnAcqDataType = vbEmpty
         End If
      End If
   End If
                                              
   If UseLibrary% Then
      mlHandle = cbWinBufAlloc(mlCount)
      If SaveFunc(Me, WinBufAlloc, mlHandle, mlCount, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      mnAcqDataType = vbInteger
      mlAcqPoints = mlCount
   Else
      mlHandle = AllocateMemory(mlCount)
      mnAcqDataType = vbInteger: mlAcqPoints = mlCount
   End If
   'bug #567, temporary line follows:
   If mlHandle = 17 Then Stop
   mlStatusType = DIFUNCTION

   CBRate& = mfRate
   PortType& = malPortsSelected(0)
   ULStat = cbDInScan(mnBoardNum, PortType&, mlCount, CBRate&, mlHandle, mvOptions)
   mfRateReturned = CBRate&
   If SaveFunc(Me, DInScan, ULStat, mnBoardNum, mnPortNum, mlCount, CBRate&, mlHandle, mvOptions, A7, A8, A9, A10, A11, 0) Then
      'ULStat = cbStopBackground(mnBoardNum)
      ULStat = StopBackground520(mnBoardNum, mlStatusType)
      If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnFormType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Exit Sub
   End If
   mfRateReturned = mfRate

   'ULStat = cbGetStatus(mnBoardNum, Status%, CurCount&, CurIndex&)
   ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
   If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   If Status% Then
      tmrCheckStatus.ENABLED = True
      cmdStop.Visible = True
      cmdStop.ENABLED = True
      mnPlot = False
   Else
      mnPlot = PlotThis%
   End If

End Sub

Private Sub ScanSelectedOutPorts(NumberOfBlocks As Long)

   ForceWinAPI% = Not (mnuULBuf.Checked)
   If mnNumSelPorts < 0 Then
      MsgBox "No digital ports have been selected.", vbCritical, "Script Error"
      Exit Sub
   End If
   
   If 0 Then
   'If Not Me.fraScript.Visible Then
      Me.fraScript.Visible = True
      Me.fraScript.Width = Me.ScaleWidth
      Me.fraScript.Height = Me.ScaleHeight
      Me.fraScript.Top = Me.ScaleTop
      Me.fraScript.Left = Me.ScaleLeft
      Me.txtScript.Width = Me.fraScript.Width - 200
      Me.txtScript.Left = Me.fraScript.Left + 100
      Me.txtScript.Height = Me.fraScript.Height - 300
   End If
   NotDone% = True
   'txtScript.Text = "Port Output Scan" & vbCrLf
   
   If Not (mlGenHandle = 0) Then
      PortType& = malPortsSelected(0)
      CBRate& = mfRate
      ULStat = cbDOutScan(mnBoardNum, PortType&, mlCount, CBRate&, mlGenHandle, mvOptions)
      mfRateReturned = CBRate&
      If SaveFunc(Me, DOutScan, ULStat, mnBoardNum, mnPortNum, mlCount, CBRate&, mlGenHandle, mvOptions, A7, A8, A9, A10, A11, 0) Then
         'ULStat = cbStopBackground(mnBoardNum)
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
         If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnFormType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Exit Sub
      End If
      mfRateReturned = mfRate
   
      'ULStat = cbGetStatus(mnBoardNum, Status%, CurCount&, CurIndex&)
      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
      If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      If Status% Then
         tmrCheckStatus.ENABLED = True
         cmdStop.Visible = True
         cmdStop.ENABLED = True
         mnPlot = False
      End If
   Else
      MsgBox "No data has been generated for output. Exiting " & _
      "'ScanSelectedPorts'.", vbOKOnly, "Error Scanning Output Ports"
      Exit Sub
   End If
   Sample& = stlSampleIndex
   If Sample& > PerChan& - 1 Then Sample& = 0

   PortType& = malPortsSelected(Port%)
   'DataValue% = OutArray(Port%, Sample&)
   'DigitalValue$ = Format(DataValue%, "0")
   If mnLibType = MSGLIB Then
      'PortString$ = "{" & Format(PortType&, "0") & "}:VALUE=" & DigitalValue$
      'PortValMsg$ = CompMsg$ & PortString$
      'MsgResult$ = MsgLibrary.SendMessage(msBoardName, PortValMsg$)
      'If SaveMsg(Me, "SendMessage(" & PortValMsg$ & ")", MsgResult$) Then Exit Sub
      'PortName$ = "PORTINDEX" & Format(PortType&, "0")
   Else
      ULStat = cbDOut(mnBoardNum, PortType&, DataValue%)
      If SaveFunc(Me, DOut, ULStat, mnBoardNum, PortType&, DataValue%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      PortName$ = GetPortString(PortType&)
   End If
   'If Port% Mod 4 = 0 Then txtScript.Text = txtScript.Text & vbCrLf
   'txtScript.Text = txtScript.Text & PortName$ & "  " & Format(DataValue%, "0") & "     "


End Sub

Public Function GetMsgDevice() As String

   BoardName$ = ""
   If Not MsgLibrary Is Nothing Then
      BoardName$ = msBoardName
   End If
   GetMsgDevice = BoardName$
   
End Function

Public Function GetCountsFromUnits(ByVal DataString As String, _
ByVal AsOffset As Integer) As Long

   If Not InStr(DataString, "%") = 0 Then
      Perc$ = Left(DataString, Len(DataString) - 1)
      FSFactor! = Val(Perc$) / 100
      AmpMax! = malPortBits(0)
      CalcAmplitude! = AmpMax! * FSFactor!
   Else
      CalcAmplitude! = Val(DataString)
   End If
   LongCounts& = Val(CalcAmplitude!)
   
   GetCountsFromUnits = LongCounts&
   
End Function

Public Function ScriptGenData(ByVal DataType&, ByVal Cycles%, ByVal TotalCount&, _
ByVal NumChans%, ByVal Amplitude, ByVal Offset, ByVal SigType%, ByVal NewData%, _
ByVal Channel%, ByVal FirstPoint&) As Long

   SampleMultiplier% = 1
   'If mnLibType = MSGLIB Then SampleMultiplier% = NumChans%
   If NewData% Then
      FirstPoint& = 0
      Samples& = TotalCount&
      Channel% = -1
      TotalSamples& = Samples& * SampleMultiplier%
   Else
      FirstPoint& = FirstPoint& * SampleMultiplier%
      TotalSamples& = TotalCount& * SampleMultiplier%
      If ((FirstPoint& + TotalSamples&) > (Samples& * SampleMultiplier%)) Then
         MsgBox "Cannot increase the number of samples when modifying existing data.", _
         vbOKOnly, "Too Many Samples"
         Exit Function
      End If
   End If
   UseWinAPI% = Not Me.mnuULBuf.Checked
   DataHandle& = GenerateData(DataType&, Cycles%, TotalSamples&, NumChans%, _
   Amplitude, Offset, SigType%, NewData%, Channel%, FirstPoint&, UseWinAPI%)
   
   SetupData DataHandle&, Samples&, NumChans%, DataType&
   PlotGenData
   ScriptGenData = DataHandle&
   
End Function

Public Sub SetBitsPerPort(ByVal PortNum As Long, _
ByVal NumBits As Long, ByVal BitsInPort As Long)

   If NumBits = -1 Then
      mlLastPortIndex = PortNum
      ReDim malPortBits(mlLastPortIndex)
      ReDim malBitsInPort(mlLastPortIndex)
   Else
      malPortBits(PortNum) = NumBits
      malBitsInPort(PortNum) = BitsInPort
   End If
   
End Sub

Function GetBitsByPortFromArray(ByVal FullDataSet As Variant, ByVal DataType As Long, _
ByVal NumPorts As Long, ByVal NumBits As Integer, Bits As Variant) As Integer

   'returns array of binary values for each port,
   'one per bit, for each value in array
   If NumBits < 1 Then Exit Function
   
   BitsPerPort& = NumBits / NumPorts
   ReDim LngBits(NumPorts - 1, NumBits - 1) As Long

   If DataType = vbLong Then
      BitRes% = 32
   Else
      BitRes% = 16
   End If
   
   If BitsPerPort& > BitRes% Then
      Residual% = NumBits Mod (BitRes%)
      BitsInRes% = ((NumBits + Residual%) - NumBits) \ BitRes%
      BitsToStuff% = BitRes%
   Else
      BitsToStuff% = NumBits
   End If
   
   For Port& = 0 To NumPorts - 1
      CurVal& = FullDataSet(Port&, Samp&)
      BitNum& = 0
      For ResBit% = 0 To Residual% - 1
         'repeat read of sample to fill all bits
         BitValue& = Abs((CurVal& And (2 ^ ResBit%)) = (2 ^ ResBit%))
         LngBits(Port&, BitNum&) = BitValue&
         BitNum& = BitNum& + 1
      Next ResBit%
      For ResDupe% = 0 To BitsInRes% '- 1
         For BitLoc% = 0 To BitsToStuff% - 1
            vCurVal = CurVal&: vLoc = BitLoc%
            vResult = Abs((vCurVal And (2 ^ vLoc)) = (2 ^ vLoc))
            BitValue& = vResult
            LngBits(Port&, BitNum&) = BitValue&
            BitNum& = BitNum& + 1
         Next BitLoc%
      Next ResDupe%
   Next Port&
   GetBitsByPortFromArray = True
   Bits = LngBits()
   
End Function

Public Function GetFormProperty(PropName As String) As Variant

   Select Case PropName
      Case "amplitude"
         GetFormProperty = msAmplitude
      Case "genoffset"
         GetFormProperty = msGenOffset
      Case "genvolts"
         GetFormProperty = msGenVolts
      Case "plottype"
         GetFormProperty = mnPlotType
      Case "staticrange"
         GetFormProperty = mlStaticRange
      Case "rate"
         GetFormProperty = mfRateReturned
      Case "resolution"
         GetFormProperty = mnResolution
      Case "siminput"
         GetFormProperty = mnSimIn
      Case "threshold"
         GetFormProperty = msThreshold
      Case Else
         GetFormProperty = "Invalid"
   End Select

End Function

Public Function SetFormProperty(PropName As String, PropVal As Variant) As Integer

   SetFormProperty = True
   Select Case PropName
      Case "cancel"
         mnCancel = PropVal
      Case Else
         MsgBox "Attempt was made to set a property (" & PropName & _
         ") that cannot be set by a script.", vbOKOnly, "Read Only Property"
         SetFormProperty = False
   End Select

End Function

Public Sub SetTrigParameters(ByVal BoardNum As Long, ByVal TrigType As Long, _
   ByVal LowThresh As Integer, ByVal HighThresh As Integer)
   
   If BoardNum = mnBoardNum Then
      mnTrigType = TrigType
      mnLowThreshold = LowThresh
      mnHighThreshold = HighThresh
      mbThresholdSet = True
      For TrigMenu% = 0 To mnuTrigType.Count - 1
         mnuTrigType(TrigMenu%).Checked = False
      Next TrigMenu%
      mnuTrigType(mnTrigType).Checked = True
   End If

End Sub

Public Function GetTrigParameters(ByVal BoardNum As Long, ByRef TrigType As Long, ByRef LowThresh As Integer, ByRef HighThresh As Integer) As Boolean

   GetTrigParameters = False
   If mbThresholdSet And (mnBoardNum = BoardNum) Then
      TrigType = mnTrigType
      LowThresh = mnLowThreshold
      HighThresh = mnHighThreshold
      GetTrigParameters = True
   End If
   
End Function

Public Sub GoCommand()

   If mnuReadOnWrite.Checked Then
      Select Case mnFuncType
         Case DIn, DIn32, DInArray, DBitIn, DInScan
            cmdGo = True
         Case DConfigPort, DConfigBit
            'Me.cmdRead = True
      End Select
   End If
   
End Sub

Private Function GetDIOResolution() As Integer
   
   Resolution% = 8
   ULStat& = cbGetConfig(DIGITALINFO, mnBoardNum, _
      DevNum&, DINUMBITS, ConfigVal&)
   If (ULStat& = 0) And (ConfigVal& > 0) Then
      BitWidth& = ConfigVal&
      If mnuOpt(3).Checked Or (ConfigVal& > 8) Then
         'word transfer
         BitWidth& = 16 'ConfigVal& * 2
         mnGenDataType = vbInteger
      End If
      If ((mvOptions And DWORDXFER) = DWORDXFER) Or (ConfigVal& > 16) Then
        BitWidth& = 32
        mnGenDataType = vbLong
        'UseWinAPI% = True
        'mnuULBuf.Checked = False
        'to do - temporary until 32-bit winArrayToBuf
      End If
      Resolution% = BitWidth&
   End If
   GetDIOResolution = Resolution%

End Function

Private Sub ConfigureValidOptions()

   Dim OptionFound As Boolean
   
   If mnFormType = DIGITAL_OUT Then
      ConfigItem& = BIDOSCANOPTIONS
   Else
      ConfigItem& = BIDISCANOPTIONS
   End If
   ULStat = cbGetConfig(BOARDINFO, mnBoardNum, 0, ConfigItem&, ConfigVal&)
   If ConfigVal& = 0 Then
      For mnuIndex = 0 To Me.mnuOpt.Count - 1
         mnuName$ = mnuOpt(mnuIndex).Caption
         NewName$ = LCase(mnuName$)
         mnuOpt(mnuIndex).Caption = NewName$
      Next
   Else
      For mnuIndex = 0 To Me.mnuOpt.Count - 1
         mnuName$ = mnuOpt(mnuIndex).Caption
         NewName$ = LCase(mnuName$)
         mnuOpt(mnuIndex).Caption = NewName$
      Next
      OptString$ = GetOptionsString(ConfigVal&, mnFormType)
      PrevLoc& = 0
      StringLength& = Len(OptString$)
      Do
         CurLoc& = InStr(PrevLoc& + 1, OptString$, " ")
         If CurLoc& > 0 Then
            CurOpt$ = Mid(OptString$, PrevLoc& + 1, CurLoc& - PrevLoc& - 1)
         ElseIf (StringLength& - PrevLoc&) > 1 Then
            CurOpt$ = Mid(OptString$, PrevLoc& + 1, StringLength& - (PrevLoc& - 1))
            CurLoc& = StringLength&
         Else
            CurLoc& = StringLength&
         End If
         PrevLoc& = CurLoc&
         OptionFound = False
         For mnuIndex = 0 To Me.mnuOpt.Count - 1
            mnuName$ = mnuOpt(mnuIndex).Caption
            If UCase(mnuName$) = CurOpt$ Then
               mnuOpt(mnuIndex).Caption = CurOpt$
               OptionFound = True
               Exit For
            End If
         Next
         If Not OptionFound Then _
            MsgBox "Option " & CurOpt$ & " not found in menus.", _
            vbInformation, "Missing Menu Item"
      Loop While CurLoc& < StringLength&
   End If
   
End Sub
