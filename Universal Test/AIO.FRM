VERSION 5.00
Begin VB.Form frmAnalog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Analog IO"
   ClientHeight    =   3675
   ClientLeft      =   225
   ClientTop       =   4125
   ClientWidth     =   6315
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
   LinkTopic       =   "frmAnalog"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   6315
   Tag             =   "AIn"
   Visible         =   0   'False
   Begin VB.CommandButton cmdConfigure 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5340
      TabIndex        =   34
      Top             =   1380
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdPlot 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2160
      TabIndex        =   33
      Top             =   1500
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdOS 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.OptionButton optFSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "to -FS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optFSP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "to +FS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraSetQueue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gain Queue Setup"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      TabIndex        =   28
      Top             =   1980
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkRandRange 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Random Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkRandChan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Random Chan"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   45
         Top             =   1020
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3060
         TabIndex        =   44
         Top             =   900
         Width           =   795
      End
      Begin VB.ComboBox cmbDataRate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cmbType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "AIO.frx":0000
         Left            =   900
         List            =   "AIO.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1140
         Width           =   1995
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "Cancel"
         Height          =   336
         Left            =   4020
         TabIndex        =   5
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox txtChan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   7
         Text            =   "0"
         Top             =   660
         Width           =   495
      End
      Begin VB.VScrollBar vsbChan 
         Height          =   420
         Left            =   900
         Max             =   72
         TabIndex        =   2
         Top             =   648
         Value           =   72
         Width           =   195
      End
      Begin VB.CommandButton cmdLoadArray 
         Appearance      =   0  'Flat
         Caption         =   "Load"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   1140
         Width           =   915
      End
      Begin VB.CommandButton cmdDone 
         Appearance      =   0  'Flat
         Caption         =   "Done"
         Default         =   -1  'True
         Height          =   336
         Left            =   5040
         TabIndex        =   6
         Top             =   1140
         Width           =   915
      End
      Begin VB.ListBox lstElement 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   180
         TabIndex        =   1
         Top             =   648
         Width           =   675
      End
      Begin VB.HScrollBar hsbQCount 
         Height          =   255
         LargeChange     =   32
         Left            =   180
         Max             =   255
         TabIndex        =   0
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtQCount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   29
         Text            =   "0"
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox chkEnableSPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enable Setpoint"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4020
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Channel to load"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "# of queue elements"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label lblGain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4020
         TabIndex        =   31
         ToolTipText     =   "Gain selected in menu applied on Load."
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label lblChannel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4020
         TabIndex        =   30
         ToolTipText     =   "Channel selected in list box (right) applied on Load."
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   1200
   End
   Begin VB.Timer tmrGoLoop 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   1200
   End
   Begin VB.CheckBox chkBuildArray 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check to load scroll data"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   21
      Text            =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox txtAmpl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   23
      Text            =   "2047"
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtOffset 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   22
      Text            =   "2048"
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.HScrollBar hsbData 
      Height          =   252
      LargeChange     =   512
      Left            =   120
      Max             =   4095
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      ToolTipText     =   "Escape"
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer tmrCheckStatus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5220
      Top             =   1200
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      Caption         =   "Go"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      ToolTipText     =   "Ctl-Enter"
      Top             =   60
      Width           =   795
   End
   Begin VB.TextBox txtCount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Text            =   "1000"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Text            =   "1000"
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtHighChan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   10
      Text            =   "0"
      Top             =   360
      Width           =   315
   End
   Begin VB.HScrollBar hsbHighChan 
      Height          =   255
      Left            =   120
      Max             =   50
      TabIndex        =   9
      Top             =   360
      Width           =   1000
   End
   Begin VB.TextBox txtLowChan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   8
      Text            =   "0"
      Top             =   60
      Width           =   315
   End
   Begin VB.HScrollBar hsbLowChan 
      Height          =   255
      Left            =   120
      Max             =   50
      TabIndex        =   4
      Top             =   60
      Width           =   1000
   End
   Begin VB.Label lblTotalCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPTCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   35
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDDEStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3300
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Offset"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3780
      TabIndex        =   25
      Top             =   1020
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.Label lblAmpl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Amplitude"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3780
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   18
      Top             =   1380
      Width           =   4635
   End
   Begin VB.Label lblRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rate"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3780
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblNumPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Samples"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3780
      TabIndex        =   16
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label lblHighChan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "High Chan"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1500
      TabIndex        =   15
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label lblLowChan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Low Chan"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1500
      TabIndex        =   14
      Top             =   120
      Width           =   1095
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
      Begin VB.Menu mnuQuit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuBoardSel 
      Caption         =   "&Board"
      Begin VB.Menu mnuBoard 
         Caption         =   "None Installed"
         Index           =   0
      End
   End
   Begin VB.Menu mnuFunc 
      Caption         =   "F&unction"
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbAIn()"
         Index           =   0
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbAInScan()"
         Index           =   1
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbVIn()"
         Index           =   8
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetSEMode 
         Caption         =   "cbAInputMode(SE)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSetSEChanMode 
         Caption         =   "cbAChanInputMode()"
      End
      Begin VB.Menu mnuAConv 
         Caption         =   "cbAConvertData()"
      End
      Begin VB.Menu mnuConvPT 
         Caption         =   "cbAConvertPretrigData()"
      End
      Begin VB.Menu mnuCalData 
         Caption         =   "cbACalibrateData()"
      End
      Begin VB.Menu mnuToEng 
         Caption         =   "cbToEngUnits()"
      End
      Begin VB.Menu mnuGetTC 
         Caption         =   "cbGetTCValues()"
      End
      Begin VB.Menu mnuFileRead 
         Caption         =   "cbFileRead()"
      End
      Begin VB.Menu mnuGetStatus 
         Caption         =   "cbGetStatus()"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGetSubSystemStatus 
         Caption         =   "cbGetSubSystemStatus()"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStopBG 
         Caption         =   "cbStopBackground()"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStopSubSystemBackground 
         Caption         =   "cbStopSubSystemBackground()"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "O&ptions"
      Begin VB.Menu mnuClearOpts 
         Caption         =   "Reset Options"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
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
         Caption         =   "CONVERTDATA"
         Index           =   3
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "SCALEDATA"
         Index           =   4
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "SINGLEIO"
         Index           =   5
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "DMAIO"
         Index           =   6
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "BLOCKIO"
         Index           =   7
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "WORDXFER"
         Index           =   8
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "SIMULTANEOUS"
         Index           =   9
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "NOFILTER"
         Index           =   10
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "NOCLEAR"
         Index           =   11
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "BURSTMODE"
         Index           =   12
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "WAITFORNEWDATA"
         Index           =   13
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "EXTTRIGGER"
         Index           =   14
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "NOCALIBRATEDATA"
         Index           =   15
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "BURSTIO"
         Index           =   16
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "RETRIGMODE"
         Index           =   17
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "NONSTREAMEDIO"
         Index           =   18
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "ADCCLOCKTRIG"
         Index           =   19
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "ADCCLOCK"
         Index           =   20
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "HIGHRESRATE"
         Index           =   21
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "SHUNTCAL"
         Index           =   22
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Custom..."
         Index           =   23
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoop 
         Caption         =   "Program or Timer Loop"
      End
      Begin VB.Menu mnuContPlot 
         Caption         =   "Continuous Plot"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuForceRes 
         Caption         =   "Resolution"
         Begin VB.Menu mnuRes 
            Caption         =   "Autodetect"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 4 bit"
            Index           =   1
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 8 bit"
            Index           =   2
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 10 bit"
            Index           =   3
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 12 bit"
            Index           =   4
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 13 bit"
            Index           =   5
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 14 bit"
            Index           =   6
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 16 bit"
            Index           =   7
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 24 bit"
            Index           =   8
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 32 bit"
            Index           =   9
         End
         Begin VB.Menu mnuRes 
            Caption         =   "Force 48 bit"
            Index           =   10
         End
         Begin VB.Menu mnuSep20 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseList 
            Caption         =   "Use List"
         End
      End
      Begin VB.Menu mnuBuff32 
         Caption         =   "Use 32-bit Buffer"
      End
      Begin VB.Menu mnuBuff64 
         Caption         =   "Use 64-bit Buffer"
      End
      Begin VB.Menu mnuAddPT 
         Caption         =   "Pretrig +512"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSaveStatus 
         Caption         =   "Save Status"
      End
      Begin VB.Menu mnuStatusRate 
         Caption         =   "Status Rate..."
      End
      Begin VB.Menu mnuUseIdle 
         Caption         =   "Use Idle"
      End
   End
   Begin VB.Menu mnuRange 
      Caption         =   "&Range (±5V)"
      Begin VB.Menu mnuNoRange 
         Caption         =   "NOTUSED"
      End
      Begin VB.Menu mnuBip 
         Caption         =   "Bipolar"
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP30VOLTS"
            Index           =   0
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP20VOLTS"
            Index           =   1
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP10VOLTS"
            Index           =   2
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP5VOLTS"
            Checked         =   -1  'True
            Index           =   3
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP4VOLTS"
            Index           =   4
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP2PT5VOLTS"
            Index           =   5
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP2VOLTS"
            Index           =   6
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP1PT25VOLTS"
            Index           =   7
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP1VOLTS"
            Index           =   8
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT625VOLTS"
            Index           =   9
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT5VOLTS"
            Index           =   10
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT25VOLTS"
            Index           =   11
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT2VOLTS"
            Index           =   12
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT1VOLTS"
            Index           =   13
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT05VOLTS"
            Index           =   14
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT01VOLTS"
            Index           =   15
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT005VOLTS"
            Index           =   16
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP1PT67VOLTS"
            Index           =   17
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT312VOLTS"
            Index           =   18
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT156VOLTS"
            Index           =   19
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT078VOLTS"
            Index           =   20
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP60VOLTS"
            Index           =   21
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP15VOLTS"
            Index           =   22
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT125VOLTS"
            Index           =   23
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT025VOLTSPERVOLT"
            Index           =   24
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT073125VOLTS"
            Index           =   25
         End
      End
      Begin VB.Menu mnuUni 
         Caption         =   "Unipolar"
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI10VOLTS"
            Index           =   0
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI5VOLTS"
            Index           =   1
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI4VOLTS"
            Index           =   2
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI2PT5VOLTS"
            Index           =   3
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI2VOLTS"
            Index           =   4
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI1PT25VOLTS"
            Index           =   5
            Shortcut        =   +{F4}
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI1VOLTS"
            Index           =   6
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT5VOLTS"
            Index           =   7
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT25VOLTS"
            Index           =   8
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT2VOLTS"
            Index           =   9
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT1VOLTS"
            Index           =   10
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT05VOLTS"
            Index           =   11
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT01VOLTS"
            Index           =   12
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNIPT02VOLTS"
            Index           =   13
         End
         Begin VB.Menu mnuUniRange 
            Caption         =   "UNI1PT67VOLTS"
            Index           =   14
         End
      End
      Begin VB.Menu mnuCur 
         Caption         =   "Current"
         Begin VB.Menu mnuCurRange 
            Caption         =   "MA4TO20"
            Index           =   0
         End
         Begin VB.Menu mnuCurRange 
            Caption         =   "MA2to10"
            Index           =   1
         End
         Begin VB.Menu mnuCurRange 
            Caption         =   "MA1TO5"
            Index           =   2
         End
         Begin VB.Menu mnuCurRange 
            Caption         =   "MAPT5TO2PT5"
            Index           =   3
         End
         Begin VB.Menu mnuCurRange 
            Caption         =   "MA0TO20"
            Index           =   4
         End
         Begin VB.Menu mnuCurRange 
            Caption         =   "BIPPT025AMPS"
            Index           =   5
         End
      End
      Begin VB.Menu mnuCustRange 
         Caption         =   "Custom..."
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQueue 
         Caption         =   "Gain Queue"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScale 
         Caption         =   "CELSIUS"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuScale 
         Caption         =   "FAHRENHEIT"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuScale 
         Caption         =   "KELVIN"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuScale 
         Caption         =   "VOLTS"
         Index           =   3
      End
      Begin VB.Menu mnuScale 
         Caption         =   "NOSCALE"
         Index           =   4
      End
      Begin VB.Menu mnuScale 
         Caption         =   "RAW"
         Index           =   5
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
   Begin VB.Menu mnuPlotSel 
      Caption         =   "Plot T&ype"
      Begin VB.Menu mnuPlotType 
         Caption         =   "Volts vs Time"
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
      Begin VB.Menu mnuPlotType 
         Caption         =   "Chunks"
         Index           =   4
      End
      Begin VB.Menu mnuPlotType 
         Caption         =   "Print List"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuPlotType 
         Caption         =   "Derivative"
         Index           =   6
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoise 
         Caption         =   "Calc RMS Noise"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBoardTitle 
         Caption         =   "Show Board in Title"
      End
      Begin VB.Menu mnuRetainPlot 
         Caption         =   "Retain Plot"
      End
      Begin VB.Menu mnuPlotChan 
         Caption         =   "Set Plot Chan"
      End
      Begin VB.Menu mnuSetBlockSize 
         Caption         =   "Set Block Size"
      End
      Begin VB.Menu mnuNextBlock 
         Caption         =   "Next Block"
      End
      Begin VB.Menu mnuSetPlotRes 
         Caption         =   "Set Plot Res"
         Begin VB.Menu mnuResolution 
            Caption         =   "Normal"
            Index           =   0
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "Autodetect"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "4 Bit"
            Index           =   2
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "8 Bit"
            Index           =   3
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "12 Bit"
            Index           =   4
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "13 Bit"
            Index           =   5
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "14 Bit"
            Index           =   6
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "16 Bit"
            Index           =   7
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "24 Bit"
            Index           =   8
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "32 Bit"
            Index           =   9
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "48 Bit"
            Index           =   10
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "64 Bit"
            Index           =   11
         End
      End
   End
   Begin VB.Menu mnuBuildData 
      Caption         =   "&Data Builder..."
   End
   Begin VB.Menu mnuUtils 
      Caption         =   "Utilitie&s"
      Begin VB.Menu mnuStore 
         Caption         =   "Store Data for Printing"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCal 
         Caption         =   "Calibration"
      End
      Begin VB.Menu mnuCalInvert 
         Caption         =   "Inverted Input"
      End
      Begin VB.Menu mnSepCal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAvg 
         Caption         =   "Set # to Average"
      End
      Begin VB.Menu mnuBufInfo 
         Caption         =   "Buffer Info"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEvalData 
         Caption         =   "Evaluate Data..."
      End
      Begin VB.Menu mnuSettle 
         Caption         =   "Settling Time"
      End
      Begin VB.Menu mnuDrift 
         Caption         =   "Drift"
      End
      Begin VB.Menu mnuErrCount 
         Caption         =   "Error vs Count"
      End
      Begin VB.Menu mnuClearPlot 
         Caption         =   "Clear Plot"
      End
      Begin VB.Menu mnuClearOnGo 
         Caption         =   "Clear on Go"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuULBuf 
         Caption         =   "UL Creates Buffer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCalLoopback 
         Caption         =   "Cal Loopback (HP34401)"
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
Attribute VB_Name = "frmAnalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'following constants are defined here temporarily
'until the version of the library supporting them is released
Const RECYCLE_EVENT = 1
'Const RECYCLED = &H80 (replaced by retrigger)
'Const RETRIGGER = &H20000
'Const NONSTREAMEDIO = &H40000

Const PROGLOOP = 1
Const TIMERLOOP = 2

'Const AUX_FILEREAD = &H10 (replaced by FileRead)

'Const CB_AOUT = 0
'Const AOUT_SCAN = 1
#If NETOPS Then
   Dim WithEvents NetLibrary As ULNet.ULNetClass
Attribute NetLibrary.VB_VarHelpID = -1
#Else
   Dim NetLibrary As Object
#End If
#If MSGOPS Then
   Dim WithEvents MsgLibrary As MBDClass.MBDComClass
Attribute MsgLibrary.VB_VarHelpID = -1
#Else
   Dim MsgLibrary As Object
#End If

Private mfrmThisForm As Form
Private mnThisInstance As Integer
Private mnFormType As Integer, msTitle As String
'Private mnDDEControl As Integer, mnDDEStat As Integer

Private mnBoardNum As Integer, mnBoardIndex As Integer
Private mnCancel As Integer

Private mnFuncType As Integer
Private mnDataType As Integer

Private mnPlot As Integer, mnLoaded As Integer
Private mlEvalCount As Long, mlTrapCount As Long
Private mbEvalEnabled As Boolean, mbEval As Boolean
Private mbEventCount As Boolean, mbStartTrap As Boolean
Private mnCountFirst As Integer, mnCountLast As Integer
Private mnCountIterations As Integer, mnCountInstances As Integer
Private mnPlotType As Integer, mnHardCopy As Integer

Private msStringID As String, msConfig As String, msOpt As String
Private mnLoop As Integer, mnExtMem As Integer
Private mnTimerTillCount As Integer

Private mDAData() As Integer, mnChanLoaded As Integer
Private mnDataChange As Integer

Private mnFirstChan As Integer, mnLastChan As Integer
Private mfRate As Double, mlCount As Long, mnRange As Integer
Private mnFixedRange As Integer
Private mlFixedOutRange As Long
Private mlFirstPoint As Long
Private mfRateReturned As Single, mnScale As Integer
Private mlPTCountReturn As Long
Private mlTotalCountReturn As Long
Private mnResolution As Integer, mlTotalCount As Long
Private mlPreTrigCount As Long, mlBufferSize As Long
Private mnTrigType As Integer, mnLowThreshold As Integer, mnHighThreshold As Integer
Private mbThresholdSet As Boolean
Private mnTrigSet As Integer, msRange As String

Private mnMemBoard As Integer

Private mlHandle As Long, mvOptions As Variant
Private mlAvgVal As Long
Private malStat() As Long

Private mnSaveAvg As Integer

Private manGainArray() As Integer, manChanArray() As Integer, mlQCount As Long
Private manTypeArray() As Integer, manModeArray() As Integer, manDRateArray() As Integer
Private manForcedRes() As Integer, mnDetailsNeedUpdate As Integer
Private sngRealArray() As Single
Private nDatArray() As Integer

Private mnIOTFunc As Integer, mlStatusType As Long
Private mnFirstSourceSet As Integer

Private mlValSet As Long, mnSampleNum As Integer
Private mlCalInterval As Long
Private mnCalConst As Integer

Private msDevName As String
Private mlEventSize As Long
Private mlEventParam As Long
Private mlSubSystem As Long
Private mnSaveAmpl As Integer, mnSaveOffset As Integer

Private mlBlockSave As Long, mlBlockSize As Long
Private mnCalcRMS As Integer

Private mnCalWarn As Integer

Private msSettleVal As Single
Private mnSettling, mnSettlePos, mnUniSettle As Integer
Private mlSettleStep As Long
Private mnErrorCounting As Integer
Private mnTimerStop As Integer
Private mnLoopType As Integer
Private mnCustomInfo As Integer, mnCustomRes As Integer
Private mnSetActive As Integer
Private mnIniChecked As Integer
Private mlChunkNum As Long

Private mfValSet As Single
Private msBoardName As String
Private mnUseTwoSources As Integer
Private mnUse32 As Integer, mnUse64 As Integer
Private mlGenHandle As Long, mlGenPoints As Long
Private mnGenChans As Integer  'mlGenDataType As Long,
Private mnAcqDataType As VbVarType, mnGenDataType As VbVarType
Private mlAcqPoints As Long
Private mlTimerCount As Long, mlCalHandle As Long

Private mlCustomRange As Long
Private mvCustomRange As Variant
Private mnScriptCommand As Integer
Private mnThreading As Integer, mnThreadState As Integer
Private mnMessaging As Integer
Private msTempString As String
Private mvDataPoint As Variant
Private mlEventType As Long, mlEventData As Long
'form properties
Private mnSimIn As Integer, mlStaticOpt As Long
Private mlStaticRange As Long, msAmplitude As String
Private msGenOffset As String, msGenVolts As String
Private msThreshold As String, mnSemaphore As Integer
Private mnDelayRestart As Integer, mlDelayTime As Long
Private mlCustomOption As Long

Private mnLibType As Integer, msMsgRange As String   'added for msgDaq
Private msStatusType As String, mnNumBoards As Integer
Private mnUnloading As Integer, mnLoading As Integer, mlQEnabled As Integer
Private mnQueueConfigurable As Integer, mnQueueSupported As Integer
Private mnQueueDataRate As Integer, mnQueueChanMode As Integer
Private mnRefreshProps As Integer, mnNumAIChans As Integer, mnNumAOChans As Integer
Private mnFormInitialized As Integer, msDisplayName As String
Private msClockOff As String, msClockOn As String
Private msScanSupport As String, msAiSupport As String
Private mnTempSupport As Integer
Private mnCalledLocally As Integer
Private mnBufferEmpty As Integer
Private mlReadTimeout As Long, mlWriteTimeout As Long
Private mnEventTimeout As Integer ', mnAIScanDataRead As Integer
Private masDefaultScanProps() As String
Private mbCatastrophicError As Boolean
Private mbGSMConflictReported As Boolean
Private mbManualStop As Boolean, mbGoClickHandled As Boolean
Private mbJustReadBuffer As Boolean, mbDataGenRequired As Boolean
Private mnQTypeOffset As Integer
Private mbIgnore578045 As Boolean
Private mnCountRolls As Integer, mbHalfRoll As Integer
Private mnInvertCal As Integer

Private Sub A_In()

   Dim PointValue As Variant
   Static stnInProgress As Integer
   mnSemaphore = True

   If stnInProgress Then Exit Sub
   stnInProgress = True
   'If mnLibType = MSGLIB Then mnUse32 = True
   UseLibrary% = mnuULBuf.Checked
   UseWinAPI% = Not mnuULBuf.Checked
   If mnuScale(5).Checked Then VOpt& = 1
   
   If mlQCount > 0 Then
      NumChans% = mlQCount - 1
      NumberOfChannels% = mlQCount
   Else
      NumChans% = Abs(mnLastChan - mnFirstChan)
      NumberOfChannels% = NumChans% + 1
      If NumberOfChannels% = 0 Then NumberOfChannels% = 1
   End If
   PerChan& = mlCount \ NumberOfChannels%
   If PerChan& = 0 Then PerChan& = 1
   mlTotalCountReturn = PerChan& * NumberOfChannels%
   DimSize& = PerChan& - 1
   If mnuCal.Checked Then DimSize& = mlCount - 1
   PlotAIn% = mnPlot
   mnPlot = False
   If mlHandle Then
      If Not BufFree(Me, mlHandle, UseWinAPI%) Then
         stnInProgress = False
         Exit Sub
      End If
      mlHandle = 0: mnAcqDataType = vbEmpty
      mlAcqPoints = 0
   End If
   
   If mnuToEng.Checked Then ConvertToReal% = True
   If (mnuPlotType(SINGLE_POINT).Checked Or Me.mnuCal.Checked) Then PlotSinglePoints% = True
   If mnuPlotType(PLOT_CHUNKS).Checked Then PlotChunks% = True
   StoreReal% = (mnFuncType = TIn) Or (mnFuncType = TInScan) Or (mnFuncType = VIn) 'Or ConvertToReal%
   
   'allocate a buffer
   If StoreReal% Or mnuCal.Checked Then
      If (mlChunkNum = 0) And (mnSampleNum = 0) Then _
        ReDim sngRealArray(NumChans%, DimSize&) As Single
      mlHandle = BufAlloc32(Me, mlCount, UseWinAPI%)
      mlAcqPoints = mlCount
      mnAcqDataType = vbSingle
   End If
   If mnLibType = MSGLIB Then
      'If 0 Then
         'to do - find another way to configure temp
               If mnTempSupport Then
                  Scales% = mnuScale.Count - 1
                  For ScaleMenu% = 0 To Scales%
                     If mnuScale(ScaleMenu%).Checked Then
                        UnitString$ = Choose(ScaleMenu% + 1, _
                        "/DEGC", "/DEGF", "/KELVIN", _
                        "/VOLTS", "", "/RAW")
                        msOpt = mnuScale(ScaleMenu%).Caption
                        Exit For
                     End If
                  Next
               End If
      'End If
      If Not Me.mnuCal.Checked Then
         mlHandle = ScaledBufAlloc(Me, mlCount, UseWinAPI%)
         mnAcqDataType = vbDouble: mlAcqPoints = BufferSize&
      End If
   Else
      If StoreReal% Or mnuCal.Checked Then
         'moved above
         'If (mlChunkNum = 0) And (mnSampleNum = 0) Then _
         'ReDim sngRealArray(NumChans%, DimSize&) As Single
         'mlHandle = BufAlloc32(Me, mlCount, UseWinAPI%)
         'mlAcqPoints = mlCount
         'mnAcqDataType = vbSingle
      Else
         If mnUse32 Then
            mlHandle = BufAlloc32(Me, mlTotalCountReturn, UseWinAPI%)
            If mlHandle = 0 Then
               stnInProgress = False
               Exit Sub
            End If
            mnAcqDataType = vbLong
            mlAcqPoints = mlTotalCountReturn
         Else
            mlHandle = BufAlloc16(Me, mlTotalCountReturn, UseWinAPI%)
            If mlHandle = 0 Then
               stnInProgress = False
               Exit Sub
            End If
            mnAcqDataType = vbInteger
            mlAcqPoints = mlTotalCountReturn
         End If
      End If
   End If
   
   If PlotSinglePoints% And Not mnuCal.Checked Then
      'to do - re-evaluate this section
      '(changed arrays to local 6/2009)
      'single point plots only one channel at a time
      ReDim DatArray%(NumChans%, 0)
      SetBlockSize mlTotalCountReturn + mnCalConst, False
   ElseIf PlotChunks% Then
      If StoreReal% Then
         If mlChunkNum = 0 Then ReDim sngRealArray(NumChans%, DimSize&)
      Else
         If mlChunkNum = 0 Then
            If mnUse32 Then
               ReDim lDatArray(NumChans%, DimSize&) As Long
            Else
               ReDim nDatArray(NumChans%, DimSize&) As Integer
            End If
         End If
         'LoopThruCount% = True
      End If
      DimSize& = 0   'so only one iteration through the Do loop occurs
   Else
      If mnLibType = MSGLIB Then
         ReDim DblArray(NumChans%, DimSize&) As Double
         mnUse32 = False
         mnUse64 = True
         If mnuCal.Checked Then
            'calibration, so average, but loop once
            DimSize& = 0   'so only one iteration through the Do loop occurs
         End If
      ElseIf Not mnuCal.Checked Then
         If StoreReal% Then
            ReDim RealArray!(NumChans%, DimSize&)
            If (mnFuncType = TInScan) Then
               NumChans% = 0
               ScanChans% = mnLastChan - mnFirstChan
               ReDim TempArray!(ScanChans%, 0)
            End If
         Else
            If mnUse32 Then
               ReDim lDatArray&(NumChans%, DimSize&)
            Else
               ReDim DatArray%(NumChans%, DimSize&)
            End If
            If ConvertToReal% Then ReDim RealArray!(NumChans%, DimSize&)
         End If
      Else
         'calibration, so average, but loop once
         DimSize& = 0   'so only one iteration through the Do loop occurs
      End If
   End If
   
   If mnDetailsNeedUpdate Then
      If PlotSinglePoints% Then
         If mnTimerTillCount And (mnLoopType = TIMERLOOP) Then
            NumPoints& = mlTotalCountReturn
            mfRateReturned = tmrGoLoop.Interval
         Else
            NumPoints& = mlAvgVal
            mfRateReturned = 0
         End If
      Else
         NumPoints& = mlTotalCountReturn
         mfRateReturned = 0
      End If
      BoardName = mnuBoard(mnBoardIndex).Caption
      If mnuPlotType(PLOT_CHUNKS).Checked Then mfRateReturned = Me.tmrGoLoop.Interval
      CurRange% = mnRange
      If mnFuncType = TIn Then
         ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
         CurRange% = mnScale
         StLen& = ERRSTRLEN   '64
         ReturnString$ = Space$(StLen&)
         ConfigLen& = StLen&
         ULStat = GetConfigString573(BOARDINFO, mnBoardNum, DevNum%, BINODEID, ReturnString$, ConfigLen&)
         boardID$ = Left$(ReturnString$, ConfigLen&)
         ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
      End If
      'RateVal = mfRateReturned / 1000 & "kHz"
      'SetDetails mnFuncType, mnFirstChan, mnLastChan - mnFirstChan, _
      NumPoints&, RateVal, CurRange%, BoardName, boardID$
      'mnDetailsNeedUpdate = False
   End If
   If mnFuncType = ATrig Then
      TrigSaveScript% = gnScriptSave
      If (mnLowThreshold = 0) Or (mnHighThreshold = 0) Then
         If Resp% = 0 Then Resp% = MsgBox("Thresholds do not appear to be set. " & _
         "Call cbATrig() anyway?", 4, "cbATrig Warning")
         If Resp% = 7 Then
            stnInProgress = False
            Exit Sub
         End If
      End If
      PreTrigType% = 1 Xor mnTrigType
      If PreTrigType% = TRIGABOVE Then
         Threshold1 = mnHighThreshold
         Threshold2 = mnLowThreshold
      Else
         Threshold1 = mnLowThreshold
         Threshold2 = mnHighThreshold
      End If
      'following required for scripting info
      A7 = mnFirstChan
      A8 = mnLastChan
      A9 = mlTotalCountReturn
      A10 = mnLowThreshold
      A11 = mnHighThreshold
      gnScriptSave = False
      ULStat = cbATrig(mnBoardNum, mnFirstChan, PreTrigType%, Threshold1, mnRange, DataValue%)
      If SaveFunc(Me, ATrig, ULStat, mnBoardNum, mnFirstChan, PreTrigType%, Threshold1, mnRange, DataValue%, A7, A8, A9, A10, A11, 0) Then
         stnInProgress = False
         Exit Sub
      End If
      gnScriptSave = TrigSaveScript%
      ULStat = cbATrig(mnBoardNum, mnFirstChan, mnTrigType, Threshold2, mnRange, DataValue%)
      If SaveFunc(Me, ATrig, ULStat, mnBoardNum, mnFirstChan, mnTrigType, Threshold2, mnRange, DataValue%, A7, A8, A9, A10, A11, 0) Then
         stnInProgress = False
         Exit Sub
      End If
      gnScriptSave = False
   End If

   If mnuContPlot.Checked Then
      cmdStop.Visible = True
      cmdStop.ENABLED = True
      mnLoop = True
   End If
   If (mnLoopType = TIMERLOOP) And mnuLoop.Checked And _
   Not (mnuCal.Checked) Then tmrGoLoop.ENABLED = True

   If mnSampleNum = 0 Then
      mnFirstSourceSet = True
   End If
   If Not gnScriptRun Then SetSource False
   SaveScript% = gnScriptSave
   Do
      For Sample& = 0 To DimSize&
         'If PlotSinglePoints% Then SetChannel Chan%
         For ChanIndex% = 0 To NumChans%
            Chan% = ChanIndex%
            If mlQCount > 0 Then Chan% = manChanArray(ChanIndex%)
            If (mnFuncType = TIn) Or (mnFuncType = VIn) Then
               If (mnFuncType = TIn) Then
                  ULStat = cbTIn(mnBoardNum, Chan% + mnFirstChan, mnScale, TempValue!, mvOptions)
                  'following required for scripting info
                  A4 = TempValue!
                  A6 = mnFirstChan
                  A7 = mnLastChan
                  A8 = mlTotalCountReturn
                  If SaveFunc(Me, TIn, ULStat, mnBoardNum, Chan% + mnFirstChan, mnScale, A4, mvOptions, A6, A7, A8, A9, A10, A11, 0) Then
                     stnInProgress = False
                     mnCancel = True
                     Exit Sub
                  End If
               Else
                  'requires time after source is set to acquire measurement
                  If mnuCal.Checked Then
                     OrgInterval& = tmrDelay.Interval
                     tmrDelay.Interval = 1000
                     tmrDelay.ENABLED = True
                     Do
                        'wait to allow reading to be taken from source
                        DoEvents
                     Loop While tmrDelay.ENABLED
                     tmrDelay.Interval = OrgInterval&
                  End If
                  If mnUse32 Then
                     CurFunc% = VIn32
                     ULStat = cbVIn32(mnBoardNum, Chan% + mnFirstChan, mnRange, _
                     dTempValue#, VOpt&)
                     A4 = dTempValue#
                  Else
                     CurFunc% = VIn
                     ULStat = cbVIn(mnBoardNum, Chan% + mnFirstChan, mnRange, _
                     TempValue!, VOpt&)
                     A4 = TempValue!
                  End If
                  'following required for scripting info
                  A6 = mnFirstChan
                  A7 = mnLastChan
                  A8 = mlTotalCountReturn
                  If SaveFunc(Me, CurFunc%, ULStat, mnBoardNum, Chan% + mnFirstChan, mnRange, A4, VOpt&, A6, A7, A8, A9, A10, A11, 0) Then
                     stnInProgress = False
                     Exit Sub
                  End If
               End If
               If PlotSinglePoints% Then
                  If Not (tmrGoLoop.ENABLED) And mnuCal.Checked Then
                     mnSampleNum = 0
                     stnInProgress = False
                     Exit Sub
                  End If
                  PointValue = A4 - mfValSet
                  Sample& = mlChunkNum
                  sngRealArray(ChanIndex%, Sample&) = PointValue
                  x% = PlotPoint(PointValue, ChanIndex%)
               ElseIf PlotChunks% Then
                  PointValue = A4 - mfValSet
                  sngRealArray(ChanIndex%, mlChunkNum) = PointValue
                  'Result = PlotRealChunk(sngRealArray(), mlChunkNum)
                  'nSampleNum = nSampleNum + 1
               Else
                  If mnUse64 Then
                     'to do - create double array
                     DblArray(ChanIndex%, Sample&) = A4
                  Else
                     RealArray!(ChanIndex%, Sample&) = A4
                  End If
               End If
               mnSampleNum = mnSampleNum + 1
               If Not gnScriptRun Then
                  If Not (mnTimerTillCount) And (mnSampleNum = (mlTotalCountReturn + mnCalConst)) And mnuCal.Checked Then
                     mnSampleNum = 0
                     LoopThruCount% = False
                     tmrGoLoop.ENABLED = False
                     cmdGo.FontBold = True
                     ReleaseSource
                  End If
               End If
            ElseIf mnFuncType = TInScan Then
               ULStat = cbTInScan(mnBoardNum, mnFirstChan, _
                  mnLastChan, mnScale, TempArray!(0, 0), mvOptions)
               A7 = mlTotalCountReturn
               If SaveFunc(Me, TInScan, ULStat, mnBoardNum, _
                  mnFirstChan, mnLastChan, mnScale, TempArray!(0, 0), _
                  mvOptions, A7, A8, A9, A10, A11, 0) Then
                  stnInProgress = False
                  Exit Sub
               End If
               For i% = 0 To ScanChans%
                  RealArray!(i%, Sample&) = TempArray!(i%, 0)
               Next i%
            Else
               If PlotSinglePoints% Then
                  If Not gnScriptRun And Not (tmrGoLoop.ENABLED) And mnuCal.Checked Then
                     mnSampleNum = 0
                     SetSource True
                     ReleaseSource
                     stnInProgress = False
                     Exit Sub
                  End If
               End If
               Iteration! = 0
               If mnUse32 Then
                  Do
                     Select Case mnLibType
                        Case UNILIB
                           ULStat = cbAIn32(mnBoardNum, Chan% + mnFirstChan, mnRange, lDataValue&, Options&)
                           If ULStat <> 0 Then Exit Do
                        Case MSGLIB
                           ChanString$ = Format(Chan% + mnFirstChan, "0")
                           AInMsg$ = "?AI{" & ChanString$ & "}:VALUE" '& UnitString$
                           MsgResult$ = MsgLibrary.SendMessage(AInMsg$)
                           MeasLoc& = InStr(1, MsgResult$, "=")
                           If MeasLoc& > 0 Then
                              ValString$ = Mid(MsgResult$, MeasLoc& + 1)
                              If IsNumeric(ValString$) Then
                                 PointValue = Val(ValString$)
                              Else
                                 Select Case ValString$
                                    Case "OTD"
                                       PointValue = -9999
                                    Case "MINRNG"
                                       PointValue = -8888
                                    Case "MAXRNG"
                                       PointValue = 8888
                                 End Select
                              End If
                              lDataValue& = PointValue
                              Do
                                 StatResult$ = MsgLibrary.SendMessage("?AI{0}:STATUS")
                              Loop While StatResult$ = "BUSY"
                           End If
                     End Select
                     If mlAvgVal > 0 Then
                        DataCum& = DataCum& + lDataValue&
                        Iteration! = Iteration! + 1
                     End If
                  Loop While Iteration! < mlAvgVal
               Else
                  Do
                     Select Case mnLibType
                        Case UNILIB
                           ULStat = cbAIn(mnBoardNum, Chan% + mnFirstChan, mnRange, DataValue%)
                           If ULStat <> 0 Then Exit Do
                        Case MSGLIB
                           If Not mnMessaging Then Exit Sub
                           ChanString$ = Format(Chan% + mnFirstChan, "0")
                           AIValMsg$ = "?AI{" & ChanString$ & "}:VALUE" & UnitString$
                           StatMsg$ = "?AI{" & ChanString$ & "}:STATUS" '& UnitString$
                           MsgResult$ = MsgLibrary.SendMessage(AIValMsg$)
                           If SaveMsg(Me, "SendMessage(" & AIValMsg$ & ")", MsgResult$) Then
                              stnInProgress = False
                              Exit Sub
                           End If
                           StatCount& = 0
                           Do
                              StatResult$ = MsgLibrary.SendMessage(StatMsg$)
                              LoopAgain% = (InStr(1, StatResult$, "BUSY") > 0)
                              LoopAgain% = LoopAgain% Or (InStr(1, StatResult$, "ERROR") > 0)
                              StatCount& = StatCount& + 1
                              If StatCount& = 1000 Then
                                 LoopAgain% = False
                                 ErrResult$ = "Error-" & StatResult$
                                 If SaveMsg(Me, "SendMessage(" & StatMsg$ & ")", ErrResult$) Then
                                    stnInProgress = False
                                    Exit Sub
                                 End If
                              End If
                           Loop While LoopAgain%
                           MeasLoc& = InStr(1, MsgResult$, "=")
                           If MeasLoc& > 0 Then
                              ValString$ = Mid(MsgResult$, MeasLoc& + 1)
                              If IsNumeric(ValString$) Then
                                 If InStr(1, ValString$, ",") Then
                                     TempVal = Split(ValString$, ",")
                                     If UBound(TempVal) > 0 Then
                                         LeftSide$ = TempVal(0)
                                         RightSide$ = TempVal(1)
                                         ValString$ = LeftSide$ & "." & RightSide$
                                     End If
                                 End If
                                 PointValue = Val(ValString$)
                              Else
                                 Select Case ValString$
                                    Case "OTD"
                                       PointValue = -9999
                                    Case "MINRNG"
                                       PointValue = -8888
                                    Case "MAXRNG"
                                       PointValue = 8888
                                    Case "NaN"
                                       ErrResult$ = "Error-" & MsgResult$
                                       If SaveMsg(Me, "SendMessage(" & AIValMsg$ & ")", ErrResult$) Then
                                          stnInProgress = False
                                          Exit Sub
                                       End If
                                 End Select
                              End If
                           End If
                     End Select
                     If mlAvgVal > 0 Then
                        If mnLibType = UNILIB Then
                           LongDataVal& = IntValToULong(DataValue%)
                           DataCum& = DataCum& + LongDataVal&
                        Else
                           FloatCum# = FloatCum# + PointValue
                        End If
                        Iteration! = Iteration! + 1
                     End If
                  Loop While Iteration! < mlAvgVal
               End If
               If mnUse32 Then
                  'following required for scripting info
                  A6 = mnFirstChan
                  A7 = mnLastChan
                  A8 = mlTotalCountReturn
                  Select Case mnLibType
                     Case UNILIB
                        If SaveFunc(Me, AIn32, ULStat, mnBoardNum, Chan% + mnFirstChan, mnRange, lDataValue&, Options&, A6, A7, A8, A9, A10, A11, 0) Then
                           tmrGoLoop.ENABLED = False
                           cmdGo.FontBold = True
                           stnInProgress = False
                           Exit Sub
                        End If
                     Case MSGLIB
                        If SaveMsg(Me, "SendMessage(" & AInMsg$ & ")", MsgResult$) Then
                           tmrGoLoop.ENABLED = False
                           cmdGo.FontBold = True
                           stnInProgress = False
                           Exit Sub
                        End If
                  End Select
               Else
                  'following required for scripting info
                  A5 = mnFirstChan
                  A6 = mnLastChan
                  A7 = mlTotalCountReturn
                  'to do - move error check here
                  If mnLibType = UNILIB Then
                     If SaveFunc(Me, AIn, ULStat, mnBoardNum, Chan% + mnFirstChan, mnRange, DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then
                        tmrGoLoop.ENABLED = False
                        cmdGo.FontBold = True
                        stnInProgress = False
                        Exit Sub
                     End If
                  End If
               End If
               'uncommented first if block below for AI Cal 4/20/2011
               If PlotSinglePoints% Then
                  If ConvertToReal% Then
                     If mnLibType = UNILIB Then
                     ULStat = cbToEngUnits(mnBoardNum, mnRange, DataValue%, EngUnits!)
                     If SaveFunc(Me, ToEngUnits, ULStat, mnBoardNum, mnRange, DataValue%, EngUnits!, A5, A6, A7, A8, A9, A10, A11, 0) Then
                        stnInProgress = False
                        Exit Sub
                     End If
                     PointValue = EngUnits!
                     End If
                  Else
                     If mnuCal.Checked Then
                        If gnScriptRun Then
                           Voltage! = Val(cmdPlot.Caption)
                           mlValSet = GetCounts(mnResolution, mnRange, Voltage!)
                        End If
                        If mnLibType = UNILIB Then
                           AveragedData! = DataCum& / Iteration!
                           DataCum& = 0
                        Else
                           AveragedData! = FloatCum# / Iteration!
                           FloatCum# = 0
                        End If
                        PointValue = AveragedData! - mlValSet
                        sngRealArray(ChanIndex%, mnSampleNum) = PointValue
                     Else
                        If mnUse32 Then
                           PointValue = lDataValue&
                        Else
                           PointValue = DataValue%
                        End If
                     End If
                  End If
                  'mnSampleNum = mnSampleNum + 1
                  'x% = PlotPoint(PointValue, Chan%)
                  If Not (tmrGoLoop.ENABLED Or mnTimerTillCount Or mnuCal.Checked) Then LoopThruCount% = True
                  If Not gnScriptRun Then
                     If 0 Then
                     If Not (mnTimerTillCount) And (mnSampleNum = (mlTotalCountReturn)) Then ' + mnCalConst
                        mnSampleNum = 0
                        mlChunkNum = 0
                        LoopThruCount% = False
                        tmrGoLoop.ENABLED = False
                        cmdGo.FontBold = True
                        ReleaseSource
                     End If
                     End If
                  End If
               Else
                  If ConvertToReal% And 0 Then
                     'If 0 Then
                     'this is done in Plot interface now
                     If mnUse32 Then
                        If mnLibType = UNILIB Then
                           ULStat = cbToEngUnits32(mnBoardNum, mnRange, lDataValue&, dEngUnits#)
                           If SaveFunc(Me, ToEngUnits32, ULStat, mnBoardNum, mnRange, lDataValue&, dEngUnits#, A5, A6, A7, A8, A9, A10, A11, 0) Then
                              stnInProgress = False
                              Exit Sub
                           End If
                           RealArray!(ChanIndex%, Sample&) = dEngUnits#
                        Else
                           RealArray!(ChanIndex%, Sample&) = PointValue
                        End If
                     Else
                        ULStat = cbToEngUnits(mnBoardNum, mnRange, DataValue%, EngUnits!)
                        If SaveFunc(Me, ToEngUnits, ULStat, mnBoardNum, mnRange, DataValue%, EngUnits!, A5, A6, A7, A8, A9, A10, A11, 0) Then
                           stnInProgress = False
                           Exit Sub
                        End If
                        RealArray!(ChanIndex%, Sample&) = EngUnits!
                     End If
                     'End If
                  Else
                     If mnUse32 Then
                        If PlotChunks% Then
                           lDatArray(ChanIndex%, mlChunkNum) = lDataValue&
                        Else
                           lDatArray&(ChanIndex%, Sample&) = lDataValue&
                        End If
                     Else
                        If PlotChunks% Then
                           nDatArray(ChanIndex%, mlChunkNum) = DataValue%
                        Else
                           If mnLibType = MSGLIB Then
                              DblArray(ChanIndex%, Sample&) = PointValue
                           Else
                              DatArray%(ChanIndex%, Sample&) = DataValue%
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            gnScriptSave = False
         Next ChanIndex%
         mnSampleNum = mnSampleNum + 1
         DoEvents
      Next Sample&
      If mnSampleNum > (mlTotalCountReturn + mnCalConst) Then mnSampleNum = 0
      'add mnuCal.Checked below on 1/21/2014 to allow timer TIn loop
      If Not (mnSampleNum < mlCount) And mnuCal.Checked Then
         mnSampleNum = 0
         tmrGoLoop.ENABLED = False
         cmdGo.FontBold = True
         If mnuCal.Checked Then ReleaseSource
      End If
      'If 0 Then   'for cal changes 4/20/2011
      If ((Not PlotSinglePoints%) And (Not PlotChunks%)) And Not mnuCal.Checked Then
         If mnLibType = MSGLIB Then
            ULStat = WDblArrayToBuf(Me, mlHandle, DblArray(), mlTotalCountReturn, UseWinAPI%)
         Else
            If StoreReal% Then
               If mbEval Then x% = EvalRealArray(RealArray!(), 0, 0)
               If mlHandle = 0 Then Exit Sub
               ULStat = WSngArrayToBuf(Me, mlHandle, RealArray!(), mlTotalCountReturn, UseWinAPI%)
            Else
               If mlHandle = 0 Then Exit Sub
               If mnUse32 Then
                  'UseWinAPI% = True
                  ULStat = WArrayToBuf32(Me, mlHandle, lDatArray&(), mlTotalCountReturn, UseWinAPI%)
               Else
                  ULStat = WArrayToBuf(Me, mlHandle, DatArray%(), mlTotalCountReturn, UseWinAPI%)
                  If Not ULStat = 0 Then
                     stnInProgress = False
                     Exit Sub
                  End If
               End If
            End If
         End If
         If ULStat = 0 Then
            mnPlot = PlotAIn%
            'PlotAcquiredData
         End If
      ElseIf PlotChunks% Or mnuCal.Checked Then
         If mlHandle = 0 Then Exit Sub
         mnPlot = True
         Select Case mnAcqDataType
            Case vbInteger
               ULStat = WArrayToBuf(Me, mlHandle, nDatArray(), mlCount, UseWinAPI%)
               'CopyMemory ByVal MemHandle&, nDatArray(0, 0), mlCount * 2
            Case vbLong
               ULStat = WArrayToBuf32(Me, mlHandle, lDatArray(), mlCount, UseWinAPI%)
               'CopyMemory ByVal MemHandle&, lDatArray(0, 0), mlCount * 4
            Case vbSingle
               BuffSamples& = mlCount '* NumberOfChannels%
               ULStat = WSngArrayToBuf(Me, mlHandle, sngRealArray(), BuffSamples&, UseWinAPI%)
               'CopyMemory ByVal MemHandle&, sngRealArray(0, 0), mlCount * 4
            Case vbDouble
               ULStat = WDblArrayToBuf(Me, mlHandle, DblArray(), mlCount, UseWinAPI%)
         End Select
      End If
      'to do - re-evaluate the chunk plotting
      '(changed arrays to local 6/2009)
      If Not Me.mnuCal.Checked Then
         mlChunkNum = mlChunkNum + 1
         PlotRealChunk mlChunkNum - 1
         If (Not (mlChunkNum < (mlCount / NumberOfChannels%))) Or (Not PlotChunks%) Then
            mlChunkNum = 0
            ReDim nDatArray(0, 0)
            ReDim lDatArray(0, 0)
            ReDim sngRealArray(0, 0)
         End If
      End If
      If mnLoop Then
         DelayVal& = tmrGoLoop.Interval * gnTimeFactor
         For LoopDelay& = 0 To DelayVal&
            DoEvents
         Next LoopDelay&
      End If
   Loop While mnLoop Or LoopThruCount%
   gnScriptSave = SaveScript%
   If mnFuncType = ATrig Then gnScriptSave = TrigSaveScript%
   stnInProgress = False
   mnSemaphore = False
   
End Sub

Private Sub A_Out()

   mnSemaphore = True
   NumChans% = (mnLastChan - mnFirstChan) + 1
   TotalCount& = Val(Me.txtCount.Text)
   If TotalCount& = 0 Then TotalCount& = 1
   If TotalCount& > mlGenPoints Then TotalCount& = mlGenPoints
   
   If TotalCount& < 32768 Then
      PerChan& = TotalCount& \ NumChans%
      FirstPoint& = 0
      If mnuCal.Checked Then
         FirstPoint& = mnSampleNum
         PerChan& = 1
         TotalCount& = 1
         If mnLibType = MSGLIB Then
            ReDim DblDAData(0, 0) As Double
         Else
            ReDim DAData(0, 0) As Integer
         End If
      Else
         If PerChan& = 0 Then PerChan& = 1
         If mnLibType = MSGLIB Then
            ReDim DblDAData(NumChans% - 1, PerChan& - 1) As Double
         Else
            ReDim DAData(NumChans% - 1, PerChan& - 1) As Integer
         End If
      End If
   Else
      MsgBox "Not yet configured for large arrays"
      mnSemaphore = False
      Exit Sub
   End If
   If mnLibType = MSGLIB Then
      ULStat = LoadDblArrayFromWinBuf64(Me, mlGenHandle, _
         DblDAData(), FirstPoint&, TotalCount&)
   Else
      ULStat = LoadArrayFromWinBuf(Me, mlGenHandle, _
         DAData(), FirstPoint&, TotalCount&)
   End If
   If Not ULStat = 0 Then
      mnSemaphore = False
      Exit Sub
   End If

   NumElements& = PerChan& - 1
   If mnTimerTillCount And (mnLoopType = TIMERLOOP) Then
      NumElements& = 0  'once per timer tick
      mfRateReturned = tmrGoLoop.Interval
   End If
   
   StoreReal% = mnuCal.Checked
   'allocate a buffer
   If StoreReal% Then
      If mnSampleNum = 0 Then ReDim _
         sngRealArray(NumChans% - 1, _
         mlCount - 1) As Single
   End If
   
   Me.cmdGo.ENABLED = False
   Do
      For Element& = 0 To NumElements&
         For Chan% = 0 To NumChans% - 1
            ElementNum& = Element&
            If mnTimerTillCount Then ElementNum& = mlTimerCount
            Select Case mnLibType
               Case UNILIB
                  DataValue% = DAData(Chan%, ElementNum&)
                  ULStat = cbAOut(mnBoardNum, Chan% + mnFirstChan, mnRange, DataValue%)
               Case MSGLIB
                  DblDataValue# = DblDAData(Chan%, ElementNum&)
                  'to do - data type work needed here (float, long)
                  ChanString$ = "{" & Format(Chan% + mnFirstChan, "0") & "}"
                  ValueString$ = ":VALUE=" & Format(DblDataValue#, "0.00000")
                  AOValMsg$ = msStatusType & ChanString$ & ValueString$
                  MsgResult$ = MsgLibrary.SendMessage(AOValMsg$)
            End Select
            'following required for scripting info
            A5 = mnFirstChan
            A6 = mnLastChan
            A7 = mlCount
            Select Case mnLibType
               Case UNILIB
                  If SaveFunc(Me, AOut, ULStat, mnBoardNum, _
                     Chan% + mnFirstChan, mnRange, DataValue%, _
                     A5, A6, A7, A8, A9, A10, A11, 0) Then
                     Me.cmdGo.ENABLED = True
                     mnSemaphore = False
                     Exit Sub
                  End If
               Case MSGLIB
                  If SaveMsg(Me, "SendMessage(" & AOValMsg$ & _
                     ")", MsgResult$) Then
                     Me.cmdGo.ENABLED = True
                     mnSemaphore = False
                     Exit Sub
                  End If
            End Select
            If mnuCal.Checked Then
               cmdStop.Visible = True
               cmdStop.ENABLED = True
               tmrDelay.ENABLED = True
               Do
                  DoEvents
                  'If Not cmdStop.Enabled Then
                  '   cmdStop.Visible = False
                  '   Exit Sub
                  'End If
               Loop While tmrDelay.ENABLED
               DataCum! = 0
               Instance% = 0
               For Iteration! = 1 To mlAvgVal
                  If InStr(1, msDevName, "34401") Then
                     LinkCommand Instance%, "Read?"
                     LinkStart Instance%
                     DoEvents
                  End If
                  'next line required due to 32 bit GPIB bug
                  'need to read 8840 twice to get updated reading
                  If InStr(1, msDevName, "8840") Then LinkCommand Instance%, "Read"
                  LinkCommand Instance%, "Read"
                  'LinkStart
                  DoEvents
                  Voltage! = LinkRetrieve(Instance%)
                  NewRange% = mnRange
                  ValRead! = GetHiResCounts(mnResolution, NewRange%, Voltage!)
                  'If mnRange = MA0TO20 Then NewRange% = UNI10VOLTS
                  'ValRead! = GetFloatCounts(mnResolution, NewRange%, Voltage!)
                  DataCum! = DataCum! + ValRead!
               Next Iteration!
               If mnDetailsNeedUpdate Then
                  BoardName = mnuBoard(mnBoardIndex).Caption
                  If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
                     RateVal = mfRateReturned / 1000 & "Hz"
                  Else
                     RateVal = mfRateReturned / 1000 & "kHz"
                  End If
                  SetDetails mnFuncType, mnFirstChan, mnLastChan - _
                    mnFirstChan, mlAvgVal, RateVal, mnRange, _
                    BoardName, mlPTCountReturn
                  mnDetailsNeedUpdate = False
               End If
               AveragedData! = DataCum! / CSng(mlAvgVal)
               If gnScriptRun Then
                  Voltage! = Val(cmdPlot.Caption)
                  mlValSet = GetCounts(mnResolution, mnRange, Voltage!)
               Else
                  If mnLibType = MSGLIB Then
                     mlValSet = DblDataValue#
                  Else
                     mlValSet = IntValToULong(DataValue%)
                  End If
               End If
               PointValue = mlValSet - AveragedData!
               If Abs(PointValue) > 4000 Then
                  mnCancel = True
                  MsgBox "Error value (" & PointValue & _
                    ") is too large to plot. Check Source.", _
                    vbOKOnly, "Data Cannot Be Plotted"
                  mnPlot = False
                  Exit Sub
               End If
               sngRealArray(0, mnSampleNum) = PointValue
            End If
         Next Chan%
         DoEvents
      Next Element&
      'If mnuCal.Checked Then
      '   cmdStop.Visible = False
      '   cmdStop.Enabled = False
      'End If
      DoEvents
   Loop While mnLoop Or LoopThruCount%
   If mnuCal.Checked Then
      'copy data to buffer
      CopyMemory ByVal mlCalHandle, sngRealArray(0, 0), mlCount * 4
   End If
   Me.cmdGo.ENABLED = True
   mnSemaphore = False

End Sub

Private Sub AnalogIn()

   If (mnLibType = MSGLIB) And mnRefreshProps Then
      WritePropsToDevice
      QChanged = ConfigureMsgQueue()
   End If
   If Me.mnuCalLoopback.Checked Then
      GetMeterValue
      DisplayData
   Else
      Select Case mnFuncType
         Case AIn, ATrig, TIn, TInScan, VIn, AIn32
            A_In
            If gnScriptRun Then
               'script closed forms while running AIn
               'three forms would be main, scripter, and the MDI form
               If Not Forms.Count > 3 Then Exit Sub
            End If
            If Not (mnuPlotType(SINGLE_POINT).Checked) Then DisplayData
         Case AInScan
            mlStatusType = AIFUNCTION
            mlPTCountReturn = 0
            If mnIOTFunc Then mlStatusType = DAQIFUNCTION
            Do
               InScan
               DoEvents
               Select Case mnLibType
                  Case UNILIB
                     If Not (mnuEvent(3).Checked Or mnuEvent(4).Checked) Then
                        If Not mbCatastrophicError Then DisplayData
                     End If
                  Case NETLIB
                  Case MSGLIB
                     If Not Me.tmrCheckStatus.ENABLED Then
                        If Not mnuGetStatus.Checked Then
                           If Not mlEventType = 0 Then
                              cmdStop.Visible = Not mnBufferEmpty
                              cmdStop.ENABLED = Not mnBufferEmpty
                           End If
                        Else
                           DisplayData
                        End If
                     End If
               End Select
            Loop While mnLoop
         Case APretrig
            mlStatusType = AIFUNCTION
            Do
               InScan
               DoEvents
               Select Case mnLibType
                  Case UNILIB
                     If Not (mnuEvent(3).Checked Or mnuEvent(4).Checked) Then DisplayData
                  Case NETLIB
                     If Not (mnuEvent(3).Checked Or mnuEvent(4).Checked) Then DisplayData
                  Case MSGLIB
               End Select
            Loop While mnLoop
         Case FileAInScan, FilePretrig
            FileScan
            DisplayData
         Case FileRead
            DisplayData
      End Select
      If (mnuGetStatus.Checked) And (Not tmrCheckStatus.ENABLED) Then
         If mnuStopBG.Checked Then
            If Not mbCatastrophicError Then
               Select Case mnLibType
                  Case UNILIB
                     ULStat = StopBackground520(mnBoardNum, mlStatusType)
                     If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                  Case NETLIB
                     ULStat = NetLibrary.StopBackground(mlStatusType)
                     If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                     If 0 Then
                        functionToCall$ = GetCallAsString(StopBackground, mnBoardNum, mlStatusType)
                        'AcqThread.CallULFunction functionToCall$
                     End If
                  Case Else
               End Select
            End If
         End If
      End If
   End If

End Sub

Private Sub AnalogOut()

   If (mnLibType = MSGLIB) And mnRefreshProps Then
      WritePropsToDevice
   End If
   Select Case mnFuncType
      Case AOut
         If mlGenHandle = 0 Then
            mnDataChange = True
            ResetData
         End If
         A_Out
      Case AOutScan
         If mlGenHandle = 0 Then
            mnDataChange = True
            ResetData
         End If
         mlStatusType = AOFUNCTION
         If mnIOTFunc Then mlStatusType = DAQOFUNCTION
         OutScan
         If (Not cmdStop.ENABLED) Then
            If mnuStopBG.Checked Then
               Select Case mnLibType
                  Case UNILIB
                     ULStat = StopBackground520(mnBoardNum, mlStatusType)
                     If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                  Case MSGLIB
                     MsgResult$ = MsgLibrary.SendMessage("AOSCAN:STOP")
                     If SaveMsg(Me, "SendMessage(AOSCAN:STOP)", MsgResult$) Then
                        cmdStop = True
                        mnBufferEmpty = True
                        Exit Sub
                     End If
                     mnBufferEmpty = False
               End Select
            End If
         End If
      Case VOut
         V_Out
   End Select

End Sub

Private Sub CheckQueue()
   
   CurElement& = lstElement.ListIndex
   If CurElement& < 0 Then
      If mnLibType = UNILIB Then Exit Sub
      'default range and chan values to queue arrays
      LowChan% = Val(txtLowChan.Text)
      HighChan% = Val(txtHighChan.Text)
      ArrayConfig% = UBound(manGainArray)
      If (HighChan% - LowChan%) > ArrayConfig% Then
         MsgBox "Attempt to set gain array with channels " & _
         "beyond array configuration (channel " & _
         Format(LowChan%, "0") & ").", vbInformation, _
         "Check Channel Configuration"
         Exit Sub
      End If
      For CurChan% = LowChan% To HighChan%
         manGainArray(Element%) = mnRange
         manChanArray(Element%) = CurChan%
         Element% = Element% + 1
      Next
      Exit Sub
   End If
   DoEvents
   If mnQueueConfigurable Then
      mnQTypeOffset = 0
      CurType% = CurElement&
      StoredType% = manTypeArray(CurElement&)
   Else
      mnQTypeOffset = 1
      CurType% = Choose(cmbType.ListIndex + 1, -1, ANALOG, DIGITAL, DIGITAL8, DIGITAL16, _
      CTR, CTR16, CTR32LOW, CTR32HIGH, CJC, TC, ANALOG_SE, ANALOG_DIFF, SETPOINTSTATUS, _
      CTRBANK0, CTRBANK1, CTRBANK2, CTRBANK3, PADZERO)
      If ((manTypeArray(CurElement&) And SETPOINT_ENABLE) = SETPOINT_ENABLE) Then SPEnabled% = 1
      StoredType% = manTypeArray(CurElement&) And &HFF
   End If
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or (Not (manGainArray(CurElement&) = mnRange))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or (Not (manChanArray(CurElement&) = Val(txtChan.Text)) Or (Not (StoredType% = CurType%)))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (chkEnableSPoint.value = SPEnabled%)
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or (Not (StoredType% = CurType%))

End Sub

Private Sub chkAuto_Click()
      
   If chkAuto.value = 0 Then
      'chkRandChan.Visible = False
      'chkRandRange.Visible = False
      'cmbType.Visible = True
   Else
      'chkRandChan.Visible = True
      'chkRandRange.Visible = True
      'cmbType.Visible = False
   End If

End Sub

Private Sub chkEnableSPoint_Click()

   CheckQueue

End Sub

Private Sub chkBuildArray_Click()

   If chkBuildArray.value = 0 Then
      hsbData.ENABLED = True
      hsbData_Scroll
   End If

End Sub

Private Sub cmdCancel_Click()

   fraSetQueue.Visible = False

End Sub

Private Sub cmdConfigure_Click()

   'this exists to give menu access to the scripting
   'form when running scripts
   CmdStr$ = Left$(cmdConfigure.Caption, 1)
   If Mid$(cmdConfigure.Caption, 2) = "True" Then
      value& = -1
   Else
      value& = Val(Mid$(cmdConfigure.Caption, 2))
   End If
   mnScriptCommand = True
   QueueCapable% = (mnQueueConfigurable Or mnQueueSupported)
   Select Case CmdStr$
      Case "#" 'return lowchan
         If mlQCount > 0 Then
            Me.cmdConfigure.Caption = "-1"
         Else
            LowChan% = mnFirstChan
            Me.cmdConfigure.Caption = Format(LowChan%, "0")
         End If
      Case "$" 'return board name
         cmdConfigure.Caption = msDisplayName
      Case "1" 'return handle to data
         Me.cmdConfigure.Caption = Format(mlHandle, "0")
      Case "2" 'return data resolution
         Me.cmdConfigure.Caption = Format(mnResolution, "0")
      Case "3" 'return range at which data was collected
         If Not (mnLibType = MSGLIB) Then
            CurRange& = GetManualRange(msBoardName, mnBoardNum, FixedRange&)
            mnFixedRange = FixedRange&
         End If
         If Not IsNull(mvCustomRange) Then
            Me.cmdConfigure.Caption = "C," & _
            Format(mnRange, "0") & "," & mvCustomRange
            Exit Sub
         End If
         Select Case mnFixedRange
            Case -2
               'undefined range (use cbToEngUnits to determine FS)
            Case -1, BIPOLAR, UNIPOLAR
               Range% = mnRange
               If mlQCount > 0 Then
                  Me.cmdConfigure.Caption = "Q"
               Else
                  Me.cmdConfigure.Caption = Format(Range%, "0")
               End If
            Case Else
               If mnNumAIChans > 0 Then
                  'MSG boards may have queue or individual chans
                  If QueueCapable% And (mlQCount > 0) Then
                     Me.cmdConfigure.Caption = "Q"
                  Else
                     Me.cmdConfigure.Caption = "I"
                  End If
               Else
                  Range% = mnFixedRange
                  Me.cmdConfigure.Caption = Format(Range%, "0")
               End If
            End Select
      Case "4" 'return rate requested
         RequestedRate! = mfRate
         If (mvOptions And HIGHRESRATE) = HIGHRESRATE Then RequestedRate! = mfRate / 1000
         Me.cmdConfigure.Caption = Format(RequestedRate!, "0.0####")
      Case "5" 'return rate returned
         ReturnedRate! = mfRateReturned
         If Not ((mvOptions And EXTCLOCK) = EXTCLOCK) Then
            If (mvOptions And HIGHRESRATE) = HIGHRESRATE Then ReturnedRate! = mfRateReturned / 1000
         End If
         Me.cmdConfigure.Caption = Format(ReturnedRate!, "0.0####")
      Case "7"
         Me.cmdConfigure.Caption = Format(mlPTCountReturn, "0")
      Case "8"
         CurOption$ = msOpt
         If msOpt = "Options = Default  " Then CurOption$ = ""
         Me.cmdConfigure.Caption = CurOption$
      Case "="
         TotalCount& = mlCount
         If mlAcqPoints > mlCount Then TotalCount& = mlAcqPoints
         If Not mlTotalCountReturn = 0 Then
            If mlTotalCountReturn < mlCount Then TotalCount& = mlTotalCountReturn
         End If
         Me.cmdConfigure.Caption = Format(TotalCount&, "0")
      Case "?"
         'number of channels
         If (mnLibType = MSGLIB) Then
            'either there's a real queue or individual
            'ranges per channel are possible (I)
            If QueueCapable% And (mlQCount > 0) Then
               QueueSet% = True
            Else
               IndividualRange% = True
            End If
         Else
            If (mlQCount > 0) Then QueueSet% = True
         End If
         NumberOfChans% = (mnLastChan - mnFirstChan) + 1
         If QueueSet% Then
            Me.cmdConfigure.Caption = "Q"
         ElseIf IndividualRange% Then
            ChansScanned$ = Format(NumberOfChans%, "0")
            NumChans% = mnNumAIChans
            TotalChans$ = Format(NumChans%, "0")
            Me.cmdConfigure.Caption = "I" & ChansScanned$ & "/" & TotalChans$
         Else
            Me.cmdConfigure.Caption = Format(NumberOfChans%, "0")
         End If
      Case "@" 'set up for calibration mode
         If Not (mnuCal.Checked = value&) Then mnuCal_Click
      Case "A" 'set continuous plot on or off
         If value& And Not mnuContPlot.Checked Then mnuContPlot_Click
         If (Not value&) And mnuContPlot.Checked Then mnuContPlot_Click
      Case "B" 'set board number
         SearchName$ = Mid$(cmdConfigure.Caption, 2)
         BoardParams = Split(SearchName$, ",")
         If UBound(BoardParams) = 1 Then DupeIndex% = Val(BoardParams(1))
         For MenuIndex% = 0 To gnNumBoards - 1
            NotFound% = True
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
         txtCount_KeyPress (13)
         ResetData
      Case "D" 'set conversion of pretrig data
         mnuConvPT.Checked = Not (value& = 0)
         SetVal% = mnuConvPT.Checked
         SetConvertData mnBoardNum, SetVal%
      Case "E" 'cbAConvertData or not
         If value& And Not mnuAConv.Checked Then mnuAConv_Click
         If Not value& And mnuAConv.Checked Then mnuAConv_Click
      Case "F" 'set function
         mnuFuncArray_Click (value&)
         If (mnLibType = MSGLIB) Then
            'options need to be set here after
            'previous configuration with static options
            Select Case mnFuncType
               Case AIn, ATrig, TIn, TInScan, VIn, AIn32
                  SetOptionsMenuByValue mlStaticOpt
            End Select
         End If
      Case "G" 'set trigger type
         mnTrigSet = True
         mnTrigType = value&
      Case "H" 'set channel high
         txtHighChan.Text = value&
      Case "I" 'set offset
         txtOffset.Text = value&
         txtOffset_KeyPress (13)
         ResetData
      Case "J" 'set convertpretrigdata on or off
         If value& And Not mnuConvPT.Checked Then mnuConvPT_Click
         If (Not value&) And mnuConvPT.Checked Then mnuConvPT_Click
      Case "K" 'set convertdata on or off
         If value& And Not mnuAConv.Checked Then mnuAConv_Click
         If (Not value&) And mnuAConv.Checked Then mnuAConv_Click
      Case "L" 'set channel low
         txtLowChan.Text = value&
      Case "M" 'set amplitude
         txtAmpl.Text = value&
         txtAmpl_KeyPress (13)
         ResetData
      Case "N" 'set file name
         txtOffset.Text = Mid$(cmdConfigure.Caption, 2)
      Case "O" 'set an option (or clear all)
         If value& = -1 Then
            ResetOpts
         Else
            mnuOpt_Click (value&)
         End If
      Case "P" 'set pretrig count
         txtAmpl.Text = value&
      Case "Q" 'get background status
         If mnLibType = MSGLIB Then
            Status% = MsgCheckStatus(Me, msBoardName, _
            MsgLibrary, msStatusType, CurCount&, CurIndex&)
            If Status% = OVERRUN Then ULStat = OVERRUN
         Else
            ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
            x% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, _
            mlStatusType, A6, A7, A8, A9, A10, A11, 0)
         End If
         StatReport$ = Str(ULStat) & "," & Str(Status%) & "," & Str(CurCount&) & "," & _
         Str(CurIndex&)
         Me.cmdConfigure.Caption = StatReport$
         If (Status% = RUNNING) Then
            BGOp$ = "Background task in progress..."
         Else
            BGOp$ = "No background task running..."
         End If
         Stat$ = BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
         lblStatus.Caption = Stat$
         DoEvents
      Case "R" 'set range
         SelectRangeUsingCode value&
      Case "S" 'set temperature scale
         mnuScale_Click (value&)
      Case "T" 'set rate
         RateVal$ = Mid$(cmdConfigure.Caption, 2)
         txtRate.Text = RateVal$
      Case "U" 'set device for GPIB
         msDevName = Mid$(cmdConfigure.Caption, 2)
      Case "V" 'set output value
         value& = IntValToULong(CInt(value&))
         txtData.Text = value&
         mnDataChange = True
         ResetData
      Case "W" 'set trigger
         'mnLowThreshold is used for trigger channel
         If mnLibType = MSGLIB Then
            Prop$ = "AITRIG:TYPE="
            TrigString$ = GetMsgTrigTypeString(mnTrigType)
            TrigTypeMsg$ = Prop$ & TrigString$
            MsgResult$ = MsgLibrary.SendMessage(TrigTypeMsg$)
            If SaveMsg(Me, "SendMessage(" & TrigTypeMsg$ & ")", MsgResult$) Then Exit Sub
         Else
            ULStat = cbSetTrigger(mnBoardNum, mnTrigType, mnLowThreshold, mnHighThreshold)
            x% = SaveFunc(Me, SetTrigger, ULStat, mnBoardNum, mnTrigType, mnLowThreshold, mnHighThreshold, A5, A6, A7, A8, A9, A10, A11, 0)
         End If
      Case "X" 'stop background
         If (cmdStop.Visible And cmdStop.ENABLED) Then
            cmdStop_Click
         Else
            If mnLibType = MSGLIB Then
               StopMsg$ = msStatusType & ":STOP"
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
            Else
                ULStat = StopBackground520(mnBoardNum, mlStatusType)
            End If
         End If
      Case "Y" 'set low threshold
         mnLowThreshold = value&
      Case "Z" 'set high threshold
         mnHighThreshold = value&
      Case "a" 'add 512 to PT buffer
         If value& And Not mnuAddPT.Checked Then mnuAddPT_Click
         If Not value& And mnuAddPT.Checked Then mnuAddPT_Click
      Case "b" 'update buffer
         mnuBufInfo_Click
      Case "c" 'set plot channel
         Chan% = value&
         SetChannel Chan%
         RePlot False
      Case "d" 'disable event
         EventType& = value&
         If EventType& = 255 Then
            mnuEvent_Click (7)
         Else
            MenuIndex% = Log(value&) / Log(2) '- 1
            If mnuEvent(MenuIndex%).Checked Then mnuEvent_Click (MenuIndex%)
         End If
         mlEventType = 0
         mlEventData = 0
      Case "e" 'set events
         Set mfrmThisForm = Me
         EventType& = value&
         If EventType& = 255 Then
            mnuEvent_Click (6)
         Else
            MenuIndex% = Log(value&) / Log(2) '- 1
            If (Not mnuEvent(MenuIndex%).Checked) Or (MenuIndex% = 3) Then mnuEvent_Click (MenuIndex%)
         End If
      Case "f" 'turn GetTC on or off
         If value& And Not mnuGetTC.Checked Then mnuGetTC_Click
         If Not value& And mnuGetTC.Checked Then mnuGetTC_Click
      Case "g" 'turn GetStatus on or off
         If value& And Not mnuGetStatus.Checked Then mnuGetStatus_Click
         If Not value& And mnuGetStatus.Checked Then mnuGetStatus_Click
      Case "h" 'set high threshold
         'mnHighThreshold = value&
      Case "i" 'load the queue element
         cmdLoadArray = True
      Case "j" 'finish the queue setup
         cmdDone = True
      Case "k" 'enable or disable data evaluation
         mbEval = value&
         mbEvalEnabled = value&
      Case "l" 'set low threshold
         mnLowThreshold = value&
      Case "m" 'set block size
         mlBlockSize = value&
         SetBlockSize mlBlockSize, False
         InitBlock True
         If Not mlHandle = 0 Then TransferBlock
      Case "n" 'plot next block of data
         NextBlock (0)
      Case "o" 'set plot type to histogram, volts vs time or none
         
      Case "p" 'set error printing
         'this is global so is set in frmScript
         geErrFlow = value&
      Case "q" 'quit - stop background task
         mnCancel = True
      Case "r" 'set resolution on plot
         ResVal% = value&
         SetPlotResolution ResVal%
      Case "s" 'set event size
         mlEventSize = value&
      Case "t" 'set trigger type
         mnTrigType = value&
      Case "u" 'set number of queue elements
         txtQCount.Text = value&
      Case "v" 'select a queue element
         If Not value& < lstElement.ListCount Then
            cmdConfigure.Caption = "Error"
            Exit Sub
         End If
         lstElement.ListIndex = value&
      Case "w" 'select a queue channel type and indicate if setpoint
         TypeValue& = value& 'And &HFF
         'If (value& And &H100) = &H100 Then CPEnabled& = 1
         If (value& = 10) Then CPEnabled& = 1
         If (value& < 0) Then TypeValue& = 0
         chkEnableSPoint.value = CPEnabled&
         If Not TypeValue& > cmbType.ListCount Then cmbType.ListIndex = TypeValue& '+ mnQTypeOffset
         CheckQueue
      Case "x" 'set plot options (retain plot)
         If mnuBoardTitle.Checked Then TitleType% = 1
         PlotOptions% = value&
         mnuRetainPlot.Checked = PlotOptions%
         SetPlotOptions PlotOptions%, TitleType%
      Case "y" 'set plot options (plot title)
         TitleType% = value&
         If TitleType% = 1 Then mnuBoardTitle.Checked = True
         PlotOptions% = mnuRetainPlot.Checked
         SetPlotOptions PlotOptions%, TitleType%
      Case "z" 'select a queue channel number
         txtChan.Text = value&
      Case "[" 'set queue channel mode
         If mnLibType = MSGLIB Then
            If value& < cmbType.ListCount Then cmbType.ListIndex = value& + mnQTypeOffset
         End If
      Case "]" 'set queue data rate
         dr! = Val(Mid$(cmdConfigure.Caption, 2))
         For ComboIndex% = 0 To Me.cmbDataRate.ListCount - 1
            ListText$ = cmbDataRate.List(ComboIndex%)
            If Val(ListText$) = dr! Then
               cmbDataRate.ListIndex = ComboIndex%
               Exit For
            End If
         Next
      Case "{" 'set plot type
         PlotType% = value&
         If Not mnuPlotType(PlotType%).Checked Then mnuPlotType_Click (PlotType%)
      Case "|"
         
      Case "}" 'convert to eng units
         mnuToEng.Checked = value&
      Case "!" 'set enable setpoint
         chkEnableSPoint.value = value&
   End Select
   mnScriptCommand = False
   DoEvents

End Sub

Private Sub cmdDone_Click()

   Dim Resp As VbMsgBoxResult
   
   QueueCapable% = (mnQueueConfigurable Or mnQueueSupported)
   If mnLibType = MSGLIB Then
      If QueueCapable% Then
         If mlQCount = 0 Then
            QueueString$ = ":QUEUE=RESET"  '":QUEUE=DISABLE" 'disable leaves count unchanged
            QueueMsg$ = msStatusType & QueueString$
            MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            If Not gnScriptRun Then
                'condition related to CAR 337583
                'this CAR won't be fixed - missing by design
                If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", _
                MsgResult$) Then Failure% = True
            End If
            QueueMsg$ = "?AIQUEUE:COUNT"
            MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            If Not (InStr(1, MsgResult$, "=0") > 1) Then
               QueueMsg$ = "AIQUEUE:CLEAR"
               MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
               If Not gnScriptRun Then MsgBox _
                  "RESET did not set queue count to zero.", _
                  vbInformation, "Queue Error"
               QueueMsg$ = "AIQUEUE:DISABLE"
               MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            End If
            If Not Failure% Then mlQEnabled = False
         Else
            QChanged% = ConfigureMsgQueue(True)
         End If
      Else
         QChanged% = ConfigureMsgQueue(True)
      End If
   Else
      EmulateQueue% = (mnFuncType = TIn) Or (mnFuncType = AIn) Or (mnFuncType = VIn)
      If EmulateQueue% Then
         fraSetQueue.Visible = False
         txtLowChan.Text = "0"
         txtHighChan.Text = "0"
         If mlQCount > 0 Then txtHighChan.Text = mlQCount - 1
         Exit Sub
      Else
         If Not mnIOTFunc Then
            ULStat = cbALoadQueue(mnBoardNum, manChanArray(0), manGainArray(0), mlQCount)
            If SaveFunc(Me, ALoadQueue, ULStat, mnBoardNum, manChanArray(0), _
            manGainArray(0), mlQCount, A5, A6, A7, A8, A9, A10, A11, 0) _
            Then Failure% = True
         End If
      End If
   End If

   If Failure% Then
      Resp = MsgBox("Channel gain queue not set. Exit Gain Queue Setup anyway?", _
      4, "Exit Queue Setup?")
      If Not (Resp = vbYes) Then Exit Sub
   End If
   fraSetQueue.Visible = False
   txtLowChan.Text = "0"
   txtHighChan.Text = "0"
   If mlQCount > 0 Then
      For QEl% = 0 To mlQCount - 1
         If manTypeArray(QEl%) > 10 Then
            'counter banks don't add channels, only resolution
            If CtrChan% = manChanArray(QEl%) Then
               ResMult% = ResMult% + 1
               If QCount% = 0 Then QCount% = QCount% + 1
            Else
               QCount% = QCount% + 1
            End If
            CtrChan% = manChanArray(QEl%)
         Else
            QCount% = QCount% + 1
         End If
      Next
      If ResMult% > 0 Then mnResolution = 16 '* ResMult%
      If (mvOptions And SCALEDATA) = SCALEDATA Then
         mnUse64 = True
      Else
         mnUse32 = (mnResolution > 16)
         mnUse64 = (mnResolution = 64)
      End If
      txtHighChan.Text = mlQCount - 1 'Format$(QCount% - 1, "0")
   End If
   ConfigureValidOptions

End Sub

Private Sub cmdGo_Click()

#If NETOPS Then
   If gnThreading And Not mnThreading Then
      Set NetLibrary = New ULNet.ULNetClass
      'AcqThread.Init
      'x& = AcqThread.ThreadID
   Else
      msStringID = ""
   End If
#End If
#If MSGOPS Then
   If mnLibType = MSGLIB Then
      If MsgLibrary Is Nothing Then
         If gnIDERunning Then
            Stop
         Else
            Dim Resp As VbMsgBoxResult
            Resp = MsgBox("This path is a Stop statement " & _
            "in the IDE. Check Local Error Handling options. " _
            & vbCrLf & vbCrLf & "          Click Yes to attempt " & _
            "to continue, No to exit application.", _
            vbYesNo, "Attempt To Continue?")
            If Resp = vbNo Then End
         End If
         'to do - remove this if it never gets hit
         Set MsgLibrary = New MBDClass.MBDComClass
      End If
   End If
#End If

   If mnuLoop.Checked And (mnLoopType = PROGLOOP) Then
      mnLoop = True
      cmdStop.ENABLED = True
   ElseIf mnuLoop.Checked And (mnLoopType = TIMERLOOP) Then
      mnLoop = False
      cmdStop.ENABLED = True
      tmrGoLoop.ENABLED = True
   End If
   If mbGoClickHandled Then
      mbGoClickHandled = False
      Exit Sub
   End If
   NumChans% = Val(txtHighChan.Text) - Val(txtLowChan.Text) + 1
   lblStatus.BackColor = &HFFFFFF
   UpdateMainStatus ""
   gnCancel = False
   mnCancel = False
   mbStartTrap = False
   mnCountRolls = 1
   gnScriptPaused = False
   cmdConfigure.Caption = ""
   If mnuClearOnGo.Checked Then ClearPlot
   DoEvents
   
   If mnuPlotType(SINGLE_POINT).Checked Then
      If mlCount <> Val(txtCount.Text) Then
         mlCount = Val(txtCount.Text)
         If mnDataChange Then SetPlotType mnPlotType + mnHardCopy, Me
         SetBlockSize mlCount + mnCalConst, False
      End If
   ElseIf mnuContPlot.Checked Then
      TempCount& = Val(txtCount.Text)
      mlCount = TempCount& - (TempCount& Mod NumChans%)
      mnSetActive = True
   Else
      '9 chars prevents overflow error
      TCount& = Val(Left(txtCount.Text, 9))
      mlCount = TCount&
      If (mnFuncType = TInScan) Or _
         (mnFuncType = TIn) Then _
         mlCount = TCount& * NumChans%
   End If
   
   For i% = 0 To 3
      If mnuPlotType(i%).Checked Then
         mnPlot = True
         Exit For
      End If
   Next i%
   If mnuPlotType(DERIVATIVE).Checked Then mnPlot = True
   DoEvents
   
   lblPTCount.Caption = ""
   lblTotalCount.Caption = ""
   Select Case mnFormType
      Case ANALOG_IN
         If mnuSettle.Checked Then
            If mnLastChan = 0 Then txtHighChan.Text = 1
            Settling
            Exit Sub
         ElseIf mnuErrCount.Checked Then
            ErrorCount
         Else
            If mnuGetTC.Checked Then
               If Not (mlCount Mod NumChans%) = 0 Then
                  txtCount.Text = Format((mlCount / NumChans%) * NumChans%, "0")
               End If
            End If
            AnalogIn
         End If
      Case ANALOG_OUT
         If (mnFuncType = AOutScan) And mbDataGenRequired Then
            mnuBuildData_Click
            Exit Sub
         End If
         AnalogOut
         If mnuCal.Checked Then DisplayData
   End Select
   If Not gnAIOForms = 0 Then
      RunningBackground% = ((mvOptions And BACKGROUND) = BACKGROUND)
      If RunningBackground% Then UpdateStatBar False
   End If
   If Not gnAIOForms = 0 Then UpdateMainStatus msStringID & msConfig & "  " & msOpt
   mnDataChange = False

End Sub

Private Sub cmdGo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
      'if right click, change timer state
      mbGoClickHandled = True
      mnLoopType = TIMERLOOP
      If mnuLoop.Checked Then
         'if already checked, just change loop rate
         LoopInterval% = tmrGoLoop.Interval
         If Shift > 0 Then
            If LoopInterval% < 16384 Then LoopInterval% = LoopInterval% * 2
         Else
            If LoopInterval% > 1 Then LoopInterval% = LoopInterval% / 2
         End If
         tmrGoLoop.Interval = LoopInterval%
      End If
      mnuLoop.Checked = True
      cmdGo = True
      cmdStop.Visible = True
   End If

End Sub

Private Sub cmdLoadArray_Click()

   AutoFill% = (chkAuto.value = 1)
   If AutoFill% Then NumElements& = lstElement.ListCount - 1
   For Element& = 0 To NumElements&
      If AutoFill% Then
         CurElement& = Element&
      Else
         CurElement& = lstElement.ListIndex
      End If
      If mnLibType = UNILIB Then
         CurType% = Choose(cmbType.ListIndex + 1, -1, ANALOG, DIGITAL, DIGITAL8, DIGITAL16, CTR, CTR16, _
         CTR32LOW, CTR32HIGH, CJC, TC, ANALOG_SE, ANALOG_DIFF, SETPOINTSTATUS, CTRBANK0, _
         CTRBANK1, CTRBANK2, CTRBANK3, PADZERO)
         If chkEnableSPoint.value = 1 Then CurType% = CurType% Or SETPOINT_ENABLE
         If Not mnIOTFunc Then
            If (CurElement& = 0) Then
               If Not (CurType% < 0) Then
                  'set first element to a chan type and require further elements define types
                  mnIOTFunc = True
                  mnuScale(0).ENABLED = mnIOTFunc
                  mnuScale(1).ENABLED = mnIOTFunc
                  mnuScale(2).ENABLED = mnIOTFunc
                  mlStatusType = DAQIFUNCTION
                  If msConfig = "cbAInScan()" Then UpdateMenu% = True
                  TempConfig$ = "cbDaqInScan()"
                  If mnFormType = ANALOG_OUT Then
                     mlStatusType = DAQOFUNCTION
                     If msConfig = "cbAOutScan()" Then UpdateMenu% = True
                     TempConfig$ = "cbDaqOutScan()"
                  End If
                  mnuFuncArray(1).Caption = TempConfig$
                  If UpdateMenu% Then msConfig = TempConfig$
                  UpdateStatBar 0
               End If
            Else
               If Not (CurType% < 0) Then
                  MsgBox "Elements already defined without chan types. Select QONLY.", , "Chan Type Required"
                  Exit Sub
               End If
            End If
         Else
            If (CurType% < 0) Then
               MsgBox "Elements already defined with chan types. Select a chan type.", , "Chan Type Required"
               Exit Sub
            End If
         End If
      Else
         If mnQueueConfigurable Then
            DoEvents
            CurType% = Me.cmbType.ListIndex + mnQTypeOffset
         End If
      End If
      
      If Not (CurElement& < 0) Then
         DoEvents
         If AutoFill% Then
            manChanArray(CurElement&) = Element&
         Else
            manChanArray(CurElement&) = Val(txtChan.Text)
         End If
         manGainArray(CurElement&) = mnRange
         manTypeArray(CurElement&) = CurType%
         lblChannel.Caption = "Chan " & manChanArray(CurElement&)
         GainLabel$ = GetRangeString(manGainArray(CurElement&))
         manModeArray(CurElement&) = Me.cmbType.ListIndex + mnQTypeOffset
         DRateKey% = Me.cmbDataRate.ListIndex + 1
         manDRateArray(CurElement&) = Choose(DRateKey%, 3750, 2000, _
            1000, 500, 100, 60, 50, 30, 25, 15, 10, 5, 2.5)
      End If
      lblGain.Caption = GainLabel$
      ConfigureControls
      If gnScriptSave Then
         ThisFunction& = ALoadQueue
         FuncStat = 0
         A1 = 0
         A2 = Val(txtQCount.Text)
         A3 = CurElement&
         A4 = CurType%
         A5 = Val(txtChan.Text)
         A6 = mnRange
         A7 = 0
         A8 = 0
         A9 = 0
         A10 = 0
         A11 = 0
         For ArgNum% = 1 To 14
            ArgVar = Choose(ArgNum%, Me.Tag, ThisFunction&, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
            If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
               PrintString$ = PrintString$ & ", "
            Else
               PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
            End If
         Next
         Print #2, PrintString$; Format$(AuxHandle, "0")
      End If
   Next
   If AutoFill% Then cmdLoadArray.ENABLED = False
   
End Sub

Private Sub cmdOS_Click()

   MidScale& = (2 ^ mnResolution / 2)
   If Me.optFSN.value = True Then
      Offset& = (-1 * MidScale&) + Val(Me.txtAmpl.Text) / 2
   Else
      Offset& = MidScale& - Val(Me.txtAmpl.Text) / 2
   End If
   Me.txtOffset.Text = Format(Offset&, "0")
   
End Sub

Public Sub eventPlot()

   'this exists so the function defined by the
   'cbEnableEvent function can access the plot subroutine
   SavePlot% = mnPlot
   For i% = 0 To 4
      If mnuPlotType(i%).Checked Then EnablePlot% = True
   Next
   mnPlot = EnablePlot%
   mbEval = mbEvalEnabled
   DisplayData
   mnPlot = SavePlot%

End Sub

Private Sub cmdStop_Click()

   LoopState% = mnLoop
   If (Not (mlEventType = ON_END_OF_INPUT_SCAN)) Or mbManualStop Then
      mbManualStop = False
      mnLoop = False
      tmrGoLoop.ENABLED = False
      tmrDelay.ENABLED = False
      mnCancel = True
      tmrGoLoop_Timer
   End If
   cmdGo.FontBold = True
   cmdStop.ENABLED = False
   lblStatus.BackColor = &HFFFFFF
   DoEvents
   If mnuStopBG.Checked Then
      Select Case mnLibType
         Case UNILIB
            ULStat = StopBackground520(mnBoardNum, mlStatusType)
            If SaveFunc(Me, StopBackground, ULStat, _
               mnBoardNum, mnFormType, A3, A4, A5, A6, _
               A7, A8, A9, A10, A11, 0) Then Failure% = True
         Case MSGLIB
            StopValid% = Not (msScanSupport = "") And Not (msStatusType = "AI")
            StopValid% = StopValid% And Not (msStatusType = "AO")
            If StopValid% Then
               StopMsg$ = msStatusType & ":STOP"
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
               If MsgResult$ = "AISCAN:STATUS=INTERRUPTED" Then
                  ULStat = INTERRUPTED
                  MsgQualifier$ = "Error: "
               End If
               If SaveMsg(Me, "SendMessage(" & StopMsg$ & ")", _
               MsgQualifier$ & MsgResult$) Then Failure% = True
               mnEventTimeout = False
            End If
      End Select
      If Failure% Then
         cmdStop.ENABLED = False
         Exit Sub
      End If
      If gnScriptRun Then cmdStop.Visible = False
   End If
   
   If Not (mlEventType = ON_END_OF_INPUT_SCAN) Then cmdStop.ENABLED = False
   If mnuContPlot.Checked Then SetContActive False
   mnPlot = mnuPlotType(0).Checked _
      Or mnuPlotType(1).Checked _
      Or mnuPlotType(2).Checked _
      Or mnuPlotType(3).Checked
   If Not mnuGetStatus.Checked Then
      If (Not (mnFormType = ANALOG_OUT)) And (Not mbEval) Then
         ScanErrorHandled% = ((mlEventType And ON_SCAN_ERROR) = ON_SCAN_ERROR)
         ScanEndHandled% = ((mlEventType And ON_END_OF_INPUT_SCAN) = ON_END_OF_INPUT_SCAN)
         If Not (ScanErrorHandled% Or ScanEndHandled%) Then DisplayData
      End If
   End If
   DoEvents
   If mnuCal.Checked Then
      If mnFormType = ANALOG_IN Then SetSource True
      mnDataChange = True
      ResetData
   End If
   If mnPlotType = SINGLE_POINT Then SetPlotType mnPlotType + mnHardCopy, Me
   mnSampleNum = 0
   mlChunkNum = 0
   Me.cmdGo.ENABLED = True

End Sub

Private Sub ConfigureControls()
   
   If mnLibType = MSGLIB Then
      MsgControls% = True
      For MenuIndex% = 2 To mnuFuncArray.Count - 1
         mnuFuncArray(MenuIndex%).Checked = False
         mnuFuncArray(MenuIndex%).ENABLED = False
      Next
      mnuAConv.Checked = False: mnuAConv.ENABLED = False
      mnuConvPT.Checked = False: mnuConvPT.ENABLED = False
      mnuCalData.Checked = False: mnuCalData.ENABLED = False
      mnuGetTC.Checked = False: mnuGetTC.ENABLED = False
      mnuFileRead.Checked = False: mnuFileRead.ENABLED = False
      ScanSupport% = Not (msScanSupport = "")
      For MenuIndex% = 0 To 2
         mnuOpt(MenuIndex%).ENABLED = ScanSupport%
      Next
      For MenuIndex% = 3 To mnuOpt.Count - 1
         mnuOpt(MenuIndex%).ENABLED = False
      Next
      mnuOpt(5).ENABLED = ScanSupport%
      mnuOpt(7).ENABLED = ScanSupport%
      mnuOpt(12).ENABLED = ScanSupport%
      mnuOpt(14).ENABLED = ScanSupport%
      mnuOpt(15).ENABLED = ScanSupport%
      mnuOpt(16).ENABLED = ScanSupport%
      mnuOpt(17).ENABLED = ScanSupport%
      'For MenuIndex% = 0 To mnuEvent.Count - 1
         'mnuEvent(MenuIndex%).Checked = False
      '   mnuEvent(MenuIndex%).ENABLED = False
      'Next
      'mnuEvent(0).ENABLED = ScanSupport%
      'mnuEvent(3).ENABLED = ScanSupport%
      'mnuEvent(4).ENABLED = ScanSupport%
      'mnuEvent(6).ENABLED = ScanSupport%
      For MenuIndex% = 2 To 9
         mnuTrigType(MenuIndex%).Checked = False
         mnuTrigType(MenuIndex%).ENABLED = False
      Next
   Else
      For MenuIndex% = 2 To mnuFuncArray.Count - 1
         mnuFuncArray(MenuIndex%).ENABLED = True
      Next
      mnuAConv.ENABLED = True
      mnuConvPT.ENABLED = True
      mnuCalData.ENABLED = True
      mnuGetTC.ENABLED = True
      mnuFileRead.ENABLED = True
      For MenuIndex% = 3 To mnuOpt.Count - 1
         mnuOpt(MenuIndex%).ENABLED = True
      Next
      'For MenuIndex% = 0 To mnuEvent.Count - 1
      '   mnuEvent(0).ENABLED = True
      'Next
      For MenuIndex% = 2 To 9
         mnuTrigType(MenuIndex%).ENABLED = True
      Next
   End If
   'display appropriate controls and hide others
   lblNumPoints.Caption = "Total Samples"
   If mnLibType = MSGLIB Then lblNumPoints.Caption = "Samples / Chan"
   If (mnFuncType = TInScan) Then lblNumPoints.Caption = "Total Scans"
   chkBuildArray.Visible = False
   ShowSingleCtls% = ((mnFuncType = AOut) Or (mnFuncType = VOut)) _
      And mnFormType = ANALOG_OUT
   ShowDataCtls% = (mnFuncType = AOutScan) Or Me.mnuCal.Checked   'mlGenPoints > 1 And
   ShowLevelCtls% = mnuCal.Checked
   hsbData.Visible = ShowSingleCtls%
   txtData.Visible = ShowSingleCtls%
   cmdGo.Visible = Not ShowSingleCtls%
   txtAmpl.Visible = ShowLevelCtls%
   lblAmpl.Visible = ShowLevelCtls%
   txtOffset.Visible = ShowLevelCtls% Or (MsgControls% And (mnFuncType = AInScan))
   lblOffset.Visible = ShowLevelCtls% Or (MsgControls% And (mnFuncType = AInScan))
   If MsgControls% And Not mnuCal.Checked Then
      txtOffset.Text = Format(mlReadTimeout, "0")
      lblOffset.Caption = "Timeout"
   Else
      ResVal@ = 2 'mnResolution
      BoardMax@ = ResVal@ ^ mnResolution
      txtOffset.Text = Format(BoardMax@ / 2, "0")
      If mnRange < 100 Then txtOffset.Text = Format(0, "0")
      lblOffset.Caption = "Offset"
   End If
   
   Me.txtCount.Visible = (ShowDataCtls% Or Me.mnuCal.Checked) Or (mnFormType = ANALOG_IN)
   Me.txtRate.Visible = (ShowDataCtls% Or Me.mnuCal.Checked) Or mnFormType = ANALOG_IN
   Me.lblNumPoints.Visible = (ShowDataCtls% Or Me.mnuCal.Checked) Or mnFormType = ANALOG_IN
   Me.lblRate.Visible = (ShowDataCtls% Or Me.mnuCal.Checked) Or mnFormType = ANALOG_IN
   If Not Me.mnuCal.Checked Then
      optFSP.Visible = False
      optFSN.Visible = False
      cmdOS.Visible = False
   End If
   mnuPlotType(SINGLE_POINT).ENABLED = (mnFuncType = AIn) Or (mnFuncType = VIn) Or (mnFuncType = TIn)
   txtData.Width = 852
   If mnFuncType = VOut Then txtData.Width = 1500
   'mnuEvent(2).Visible = (mnFormType = ANALOG_IN)    'on_pretrigger
   'mnuEvent(4).Visible = (mnFormType = ANALOG_IN)    'on_end_of_ai_scan
   'mnuEvent(5).Visible = (mnFormType = ANALOG_OUT)   'on_end_of_ao_scan
   'mnuScale(5).ENABLED = (mnFuncType = VIn) Or mnTempSupport
   chkBuildArray.Visible = False
   chkBuildArray.Caption = "Check to load scroll data"

   If mnuPlotType(SINGLE_POINT).Checked Then
      If (Not mnuCal.Checked) And (Not (mnFuncType = AIn)) And (Not (mnFuncType = TIn)) _
      And (Not (mnFuncType = VIn)) Then mnuPlotType_Click (VOLTS_VS_TIME)
   End If
   If mnFormType = ANALOG_IN Then
      Select Case mnFuncType
         Case FileAInScan, FilePretrig, FileRead
            lpFileName$ = "UniTest.ini"
            lpApplicationName$ = "DataFile"
            lpKeyName$ = "FileName"
            lpDefault$ = "C:\Demo.dat"
            nSize% = 80
            DataFile$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, DataFile$, nSize%, lpFileName$)
            DataFile$ = Left$(DataFile$, StringSize%)
            txtOffset.Width = 3500
            txtOffset.Text = DataFile$
            txtOffset.Visible = True
         Case Else
            txtOffset.Width = 975
            If Not MsgControls% Then txtOffset.Text = "0"
      End Select
      Select Case mnFuncType
         Case AInScan
            If mnIOTFunc Then
               lblAmpl.Caption = "Pretrig Count"
               txtAmpl.Text = "0"
               txtAmpl.Visible = True
               lblAmpl.Visible = True
            End If
         Case APretrig, FilePretrig
            lblAmpl.Caption = "Pretrig Count"
            If txtAmpl.Text = "" Then txtAmpl.Text = "300"
            txtAmpl.Visible = True
            lblAmpl.Visible = True
            chkBuildArray.Visible = True
            chkBuildArray.value = 0
            chkBuildArray.Caption = "Ignore TRIGSTATE error"
         Case VIn
            lblAmpl.Caption = "Amplitude"
            txtAmpl.Text = 2 ^ Abs(mnResolution)
         Case Else
            lblAmpl.Caption = "Amplitude"
            If mnuCal.Checked Then
               txtAmpl.Text = 2 ^ Abs(mnResolution)
               Me.txtOffset.Text = ""
            Else
               txtAmpl.Text = 2 ^ Abs(mnResolution) / 2
            End If
      End Select
      If (Not gnScriptRun) Then
         If (mnResolution > 16) Then
            If Not Me.mnuBuff32.Checked Then mnuBuff32_Click
         Else
            If Me.mnuBuff32.Checked Then mnuBuff32_Click
         End If
      End If
   Else
      If (Not (mlGenHandle = 0)) And (Not gnScriptRun) Then
         If (mnFuncType = AOut) Then
            mlGenPoints = 1
            msOpt = ""
         End If
         Samples& = Val(Me.txtCount.Text)
         If (mnFuncType = AOutScan) And (mlGenPoints < Samples&) _
            Then mbDataGenRequired = True 'mnuBuildData_Click
      End If
   End If
   
End Sub

Private Sub DisplayData()

   Dim PlottingForm As Form
   
   If mnLibType = NETLIB Then SetBoardReference NetLibrary
   
   For FormIndex% = 0 To Forms.Count - 1
      'error form is modal - can't show plot
      If Forms(FormIndex%).Name = "frmErrors" Then Exit Sub
   Next
   If mnuPlotType(SINGLE_POINT).Checked Then PlotSinglePoints% = True
   Chans% = mnLastChan - mnFirstChan
   'If mlQCount > 0 Then Chans% = mlQCount - 1
   If Not ((mnFuncType = TIn) Or (mnuCal.Checked)) Then
      If mnLibType = MSGLIB Then
         ScalingData% = mnuToEng.Checked
         SpecifyMsgLib& = &H100
         '0x40 = FROMMESSAGE, 0x20 = float, 3 = double
         '0x10 = scaled, 0x80 = BUFFER, 4 = single
         If ScalingData% Then
            Select Case mnFuncType
               Case AIn, ATrig, TIn, TInScan, VIn, AIn32
                  BufType% = 3 Or &H30 'Or SpecifyMsgLib&
               Case Else
                  'should not be required - done in TransferDataFromMsg
                  'BufType% = 2 Or &H30 'Or SpecifyMsgLib&
            End Select
         Else
            Select Case mnFuncType
               Case AIn, ATrig, TIn, TInScan, VIn, AIn32
                  '0x80 = BUFFER
                  BufType% = 3 Or &H20 'Or SpecifyMsgLib&
               Case Else
                  BufType% = 2 'Or SpecifyMsgLib&
            End Select
         End If
         SetBufferType BufType%
      Else
         If mnUse64 Then
            If (mvOptions And SCALEDATA) = SCALEDATA Then
               '3 sets resolution, 0x10 sets use of cbScaleWinBufToArray
               '0x20 sets float type
               BufRes% = 3
               BufType% = BufRes% Or &H30
            Else
               BufType% = 5
               If mnUse32 Then BufType% = 3
            End If
            SetBufferType BufType%
         ElseIf mnUse32 Then
            SetBufferType 2
         ElseIf ScalingData% Then
            BufType% = &H10 Or 1
            SetBufferType BufType%
         Else
            SetBufferType 1
         End If
      End If
   Else
      '0x40 = FROMMESSAGE, 0x20 = float, 3 = double
      '0x10 = scaled, 0x80 = BUFFER, 4 = single
      If Me.mnuCal.Checked Then
         BufType% = 4 Or &H20
         SetBufferType BufType%
      End If
      Select Case mnFuncType
         Case TIn, TInScan, VIn
            BufType% = 4 Or &H20
            SetBufferType BufType%
      End Select
   End If
   If (mnuCal.Checked Or mnuDrift.Checked) And Not _
      ((mnFuncType = AIn) Or (mnFuncType = AIn32)) Then
      'set up local array and copy existing data into it
      NumPoints& = mlCount + mnCalConst
      CalSamples& = (Chans% + 1) * mlAvgVal
      CalPoints& = (Chans% + 1) * mlCount
      ReDim CalBuffer(Chans%, CalPoints& - 1) As Single   'CalPoints&
      If Not mlCalHandle = 0 Then
         CopyMemory CalBuffer(0, 0), ByVal mlCalHandle, CalPoints& * 4
      End If
      If mnuCal.Checked Then  'PlotSinglePoints%
         If (Not (tmrGoLoop.ENABLED)) And (Not gnScriptRun) Then
            mnSampleNum = 0
            If mnuCal.Checked And (mnFormType = ANALOG_IN) Then
               SetSource True
               ReleaseSource
            End If
            RePlot False
            Exit Sub
         End If
      End If
      
      If mnFormType = ANALOG_IN Then
         If mnAcqDataType = vbDouble Then
            ReDim DblArray(Chans%, mlAvgVal - 1) As Double
         ElseIf mnUse32 Then
            ReDim AvgLongData&(Chans%, mlAvgVal - 1)
         Else
            ReDim AvgData%(Chans%, mlAvgVal - 1)
         End If
         If Not IsEmpty(mlHandle) Then
            If mnAcqDataType = vbDouble Then
               Success% = TransferDataFromMsg(mlAvgVal, False)
               ULStat = LoadDblArrayFromWinBuf64(Me, mlHandle, DblArray(), 0, CalSamples&)
            ElseIf mnUse32 Then
               ULStat = WBufToArray32(mlHandle, AvgLongData&(0, 0), 0, CalSamples&)
            Else
               ULStat = cbWinBufToArray(mlHandle, AvgData%(0, 0), 0, CalSamples&)
            End If
            If mnuCalData.Checked Then
               ULStat = cbACalibrateData(mnBoardNum, CalSamples&, mnRange, AvgData%(0, 0))
               If SaveFunc(Me, ACalibrateData, ULStat, mnBoardNum, CalSamples&, _
               mnRange, AvgData%(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            End If
            If ULStat <> 0 Then
               tmrGoLoop.ENABLED = False
               cmdGo.FontBold = True
               Exit Sub
            End If
         End If
         
         For CurChan% = 0 To Chans%
            DataCum = 0
            For SampleNum& = 0 To mlAvgVal - 1
               If mnAcqDataType = vbDouble Then
                  UnsData = DblArray(CurChan%, SampleNum&)
                  DataCum = DataCum + UnsData
               ElseIf mnUse32 Then
                  If AvgLongData&(CurChan%, SampleNum&) < 0 Then
                     UnsData = 2 ^ 32 + AvgLongData&(0, SampleNum&)
                  Else
                     UnsData = AvgLongData&(CurChan%, SampleNum&)
                  End If
                  DataCum = DataCum + UnsData
               Else
                  DataCum = DataCum + IntValToULong(AvgData%(CurChan%, SampleNum&))
               End If
            Next SampleNum&
            
            If mnuCal.Checked Then  'PlotSinglePoints%
               If gnScriptRun Then
                  Voltage! = Val(cmdPlot.Caption)
                  mlValSet = GetCounts(mnResolution, mnRange, Voltage!)
               End If
               AveragedData! = DataCum / mlAvgVal
               PointValue = AveragedData! - mlValSet
            End If
            If Abs(PointValue) > 28000 Then
               mnCancel = True
               MsgBox "Error value is too large to plot. Check Source.", _
               vbOKOnly, "Data Cannot Be Plotted"
               Exit Sub
            End If
            CalBuffer(CurChan%, mnSampleNum) = PointValue
         Next
         CopyMemory ByVal mlCalHandle, CalBuffer(0, 0), (CalPoints& * 4)
      End If
      
      PlotSamples& = mnSampleNum
      PlotRealChunk PlotSamples&
      mnSampleNum = mnSampleNum + 1
      If mnPlot Then PlotBuffer mlCalHandle, CalPoints& + mnCalConst, Chans%
      If mnSampleNum = (mlCount + mnCalConst) Then
         mnSampleNum = 0
         tmrGoLoop.ENABLED = False
         cmdGo.FontBold = True
         If mnuCal.Checked And (mnFormType = ANALOG_IN) Then
            SetSource True
            Success% = GPIBSetDevice(0, "DP8200")
            If Success% Then ReleaseSource
         End If
      End If
      If mnDetailsNeedUpdate Then
         If Me.mnuCalLoopback.Checked Then
            BoardName = "HP34401 loopback"
         Else
            BoardName = mnuBoard(mnBoardIndex).Caption
         End If
         lblStatus.Caption = msStringID & msConfig & " " & msOpt
         If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
            RateVal = mfRateReturned / 1000 & "Hz"
         Else
            If mfRateReturned < 1000 Then
               RateVal = mfRateReturned & "Hz"
            Else
               RateVal = mfRateReturned / 1000 & "kHz"
            End If
         End If
         SetDetails mnFuncType, mnFirstChan, Chans%, mlAvgVal, _
         RateVal, mnRange, BoardName, mlPTCountReturn
         mnDetailsNeedUpdate = False
         If mbEvalEnabled Then SetRate mfRateReturned
      End If
   
   ElseIf mnSettling Then
      If Not Chans% < 0 Then
         ReDim AvgData%(Chans%, mlAvgVal)
         If Not IsEmpty(mlHandle) Then
            ULStat = cbWinBufToArray(mlHandle, AvgData%(0, 0), 0, mlCount - 1)
         End If
      End If
      
      For SampleNum& = 0 To mlAvgVal - 1
         DataCum = DataCum + IntValToULong(AvgData%(mnLastChan, SampleNum&))
      Next SampleNum&
      
      AveragedData! = DataCum / mlAvgVal
      PointValue = AveragedData! - msSettleVal
      
      x% = PlotPoint(PointValue, 0)

   ElseIf mnErrorCounting Then
      If Not Chans% < 0 Then
         ReDim AvgData%(mnLastChan - mnFirstChan, mlCount)
         If Not IsEmpty(mlHandle) Then
            ULStat = cbWinBufToArray(mlHandle, AvgData%(0, 0), 0, mlCount)
         End If
      End If
      
      For SampleNum& = 0 To mlCount - 1
         DataCum = DataCum + IntValToULong(AvgData%(mnFirstChan, SampleNum&))
      Next SampleNum&
      
      AveragedData! = DataCum / mlCount
      PointValue = Abs(AveragedData! - msSettleVal)
      
      If mnPlot Then x% = PlotPoint(PointValue, 0)

   ElseIf mnPlot Then
      Set PlottingForm = Me
      SetPlotType mnPlotType + mnHardCopy, PlottingForm
      If PlottingForm Is Nothing Then mnDetailsNeedUpdate = True
      DoEvents
      Convert% = mnuToEng.Checked
      'ConvertToEng mnBoardNum, Convert%
      'to do - remove following line when queue bug is fixed
      mnDetailsNeedUpdate = mnDetailsNeedUpdate Or (mnLibType = MSGLIB)
      If mnDetailsNeedUpdate Then
         BoardName = mnuBoard(mnBoardIndex).Caption
         lblStatus.Caption = msStringID & msConfig & " " & msOpt
         If mlQCount > 0 Then
            ArrSize& = UBound(manTypeArray)
            For QElement& = 0 To mlQCount - 1
               If mnQueueConfigurable Then
                  If mnLibType = MSGLIB Then
                     QueueString$ = "?AIQUEUE{" & Format(QElement&, "0") & "}:"
                     For QPart% = 1 To 4
                        ItemValid% = Choose(QPart%, -1, _
                        -1, mnQueueChanMode, mnQueueDataRate)
                        If ItemValid% Then
                           QueueItem$ = Choose(QPart%, "CHAN", _
                           "RANGE", "CHMODE", "DATARATE")
                           QueueMsg$ = QueueString$ & QueueItem$
                           MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
                           If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then
                              Failure% = True
                              Exit For
                           End If
                           SplitMsg = Split(MsgResult$, "=")
                           If UBound(SplitMsg) > 0 Then
                              QueueVals$ = SplitMsg(1)
                              Select Case QPart%
                                 Case 1
                                    Chan$ = QueueVals$
                                 Case 2
                                    RangeString$ = QueueVals$
                                    RangeElement% = GetRangeCodeFromMsg(RangeString$)
                                 Case 3
                                    TypeStr$ = QueueVals$
                                    If QueueVals$ = "DIFF" Then TypeStr$ = "DF"
                                    'TypeStr$ = Choose(ModeElement% + 1, "DF", "SE")
                                 Case 4
                                    DRateElement% = Val(QueueVals$)
                              End Select
                           End If
                        End If
                        'ChanVar = ChanVar & TypeStr$ & Chan$ & ", "
                        'RangeVar = RangeVar & Format$(RangeElement%, "0") & ", "
                     Next
                  End If
               Else
                  If (ArrSize& + 1) < mlQCount Then
                     TypeStr$ = "A"
                  Else
                     If (Not manTypeArray(QElement&) < 0) Then
                        TypeVal% = (manTypeArray(QElement&) And &HFF) + 1
                        TypeStr$ = Choose(TypeVal%, "A", "d", "D", "C", _
                        "CL", "CH", "J", "T", "SE", "DF", "SP", "C0", "C1", _
                        "C2", "C3", "P0", "DD", "CC")
                     Else
                        TypeStr$ = "A"
                     End If
                  End If
               End If
               ChanVar = ChanVar & TypeStr$ & Format$(manChanArray(QElement&), "0") & ", "
               RangeVar = RangeVar & Format$(manGainArray(QElement&), "0") & ", "
            Next QElement&
            ChanVar = Left$(ChanVar, Len(ChanVar) - 2)
            RangeVar = Left$(RangeVar, Len(RangeVar) - 2)
            QVar = "Q"
         ElseIf mnNumAIChans > 0 Then
            NumChans% = (mnLastChan - mnFirstChan)
            For ChanIndex% = 0 To NumChans%
               If ChanIndex% > UBound(manGainArray) Then
                  MsgBox "Gain array not set up properly. " & _
                  "Check device properties.", vbCritical, "Invalid Gain Array"
                  Exit Sub
               End If
               CurChan% = ChanIndex% + mnFirstChan
               ChanVar = ChanVar & Format(CurChan%, "0") & ", "
               If mnLibType = MSGLIB Then
                  ChanString$ = "{" & Format(CurChan%, "0") & "}"
                  RangeString$ = ":RANGE"
                  If mnFuncType = AInScan Then
                     RangeMsg$ = "?" & msStatusType & RangeString$ & ChanString$
                  Else
                     StatType$ = "?" & msStatusType
                     If msStatusType = "" Then StatType$ = "?AI"
                     RangeMsg$ = StatType$ & ChanString$ & RangeString$
                  End If
                  MsgResult$ = MsgLibrary.SendMessage(RangeMsg$)
                  ValLoc& = InStr(1, MsgResult$, "=")
                  If ValLoc& > 0 Then CurRange$ = Mid(MsgResult$, ValLoc& + 1)
                  RangeCode% = GetRangeCodeFromMsg(CurRange$)
                  RangeVar = RangeVar & Format$(RangeCode%, "0") & ", "
               Else
                  If UBound(manGainArray) >= (CurChan%) Then
                     RangeVar = RangeVar & Format$(manGainArray(CurChan%), "0") & ", "
                  End If
               End If
            Next
            ChanVar = Left$(ChanVar, Len(ChanVar) - 2)
            If Len(RangeVar) > 0 Then RangeVar = Left$(RangeVar, Len(RangeVar) - 2)
            QVar = "Q"
         Else
            ChanVar = Format$(mnFirstChan, "0")
            RangeVar = Format$(mnRange, "0")
            QVar = Chans%
         End If
         If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
            RateVal = mfRateReturned / 1000 & "Hz"
         Else
            If mfRateReturned < 1000 Then
               RateVal = mfRateReturned & "Hz"
            Else
               RateVal = mfRateReturned / 1000 & "kHz"
            End If
         End If
         If (mnFuncType = TIn) Or (mnFuncType = TInScan) Then
            Scales% = mnuScale.Count - 1
            For ScaleMenu% = 0 To Scales%
               If mnuScale(ScaleMenu%).Checked Then
                  RangeVar = mnuScale(ScaleMenu%).Caption
                  Exit For
               End If
            Next
         End If
         SetDetails mnFuncType, ChanVar, QVar, mlTotalCountReturn, _
         RateVal, RangeVar, BoardName, mlPTCountReturn
         mnDetailsNeedUpdate = False
         If mbEvalEnabled Then SetRate mfRateReturned
      End If
      If mnPlot Then DrawLine 0
      If Chans% < 0 Then Exit Sub

      Select Case mnFuncType
         Case AIn, ATrig, AIn32, TIn
            If mlHandle = 0 Then Exit Sub
            Samples& = mlCount
            If mnuCal.Checked Then Samples& = mlCount * (Chans% + 1)
            If mbEval Then ErrorDetected% = EvalBuffer(mlHandle, Samples&, Chans%)
            PlotSamples& = mnSampleNum
            If PlotSamples& = 0 Then PlotSamples& = mlCount
            If mnuCal.Checked Then PlotRealChunk PlotSamples& - 1
            DoEvents
            If mnPlot Or ErrorDetected% Then PlotBuffer mlHandle, Samples&, Chans%
         Case AInScan, AOutScan, AOut
            If mnuEvent(4).Checked Or mnuEvent(5).Checked Then mlTotalCount = Val(lblTotalCount.Caption)
            If mnuContPlot.Checked Then
               If mnLibType = MSGLIB Then
                  UsingCallback% = (mlEventType = ON_DATA_AVAILABLE)
                  If Not UsingCallback% Then
                     If mnMessaging Then
                         Status% = MsgCheckStatus(Me, msBoardName, _
                         MsgLibrary, "AISCAN", CurCount&, CurIndex&)
                         NumSamples& = mlCount / 2
                         Success% = TransferDataFromMsg(NumSamples&)
                     End If
                  Else
                     Exit Sub
                  End If
                  'If mnMessaging Then TransferDataFromMsg 'And (Not mnBufferEmpty)
               Else
                  If mlHandle = 0 Then Exit Sub
                  If mbEval Then ErrorDetected% = EvalPartialBuf(mlHandle, _
                  (mlCount) \ 2, mlFirstPoint, Chans%)
                  If mnPlot Or ErrorDetected% Then PlotPartialBuf mlHandle, _
                  (mlCount), mlFirstPoint, Chans%
               End If
               If mnSetActive Then
                  SetContActive True
                  mnSetActive = False
               End If
            ElseIf mnuGetTC.Checked Then
               FirstPoint& = 0
               NumChans% = (Chans%) + 1
               If mlCount Mod NumChans% Then
                  txtCount.Text = Format((mlCount / NumChans%) * NumChans%, "0")
               End If
               If mlQCount > 0 Then NumPerChan& = mlCount / mlQCount
               PlotTCValues mnBoardNum, manChanArray(), manTypeArray(), _
               mlQCount, mlHandle, FirstPoint&, NumPerChan&, mnScale, manGainArray()
            Else
               If mnIOTFunc Then
                  If mlPTCountReturn > 0 Then
                     PTCount& = mlPTCountReturn
                     If mnPlot Then DrawLine PTCount&
                  End If
               End If
               If mnExtMem Then
                  If mbEval Then ErrorDetected% = EvalBuffer(mlHandle, mlCount, Chans%)
                  If mnPlot Or ErrorDetected% Then PlotMemory mnBoardNum, mnMemBoard, mlCount, 0, Chans%
               Else
                  If (mlHandle = 0) Or (mnLibType = MSGLIB) Then
                     UsingCallback% = (mlEventType = ON_DATA_AVAILABLE)
                     If (mnLibType = MSGLIB) And (Not UsingCallback%) Then
                        EndScanHandled% = ((mlEventType And ON_END_OF_INPUT_SCAN) _
                        = ON_END_OF_INPUT_SCAN)
                        'If mnAIScanDataRead Then Exit Sub
                        If mnMessaging And Not (EndScanHandled%) Then
                            Status% = MsgCheckStatus(Me, msBoardName, MsgLibrary, _
                            "AISCAN", CurCount&, CurIndex&)
                            Success% = TransferDataFromMsg(CurCount&)
                            If Not Success% Then Exit Sub
                        End If
                     Else
                        Exit Sub
                     End If
                  Else
                     PlotCount& = mlCount
                     If (mlQCount > 1) And ((msDisplayName = "USB-CTR08") Or _
                        (msDisplayName = "USB-CTR04")) Then
                        If mnUse32 Or mnUse64 Then
                           PlotCount& = mlCount / (Chans% + 1)
                           Chans% = 0
                        End If
                     End If
                     If mbEval Then ErrorDetected% = EvalBuffer(mlHandle, PlotCount&, Chans%)
                     If mnPlot Or ErrorDetected% Then PlotBuffer mlHandle, PlotCount&, Chans%
                  End If
               End If
            End If
         Case APretrig
            If mlPreTrigCount < 0 Then
               mlPreTrigCount = Val(txtAmpl.Text)
            Else
               PTCount& = mlPTCountReturn
               If mnPlot Then DrawLine PTCount&
            End If
            If mnExtMem Then
               If mnPlot Then PlotMemPretrig mnBoardNum, mnMemBoard, mlCount, Chans%
            Else
               If mlHandle = 0 Then Exit Sub
               If mnuConvPT.Checked Then
                  ConvertPTBuffer mnBoardNum, mlHandle, mlPTCountReturn, mlTotalCountReturn, mlBufferSize, Chans%
               Else
                  If mnPlot Then PlotBuffer mlHandle, mlTotalCountReturn, Chans%
               End If
            End If
         Case FileAInScan, FilePretrig
            'final counts are determined using cbFileGetInfo() in the PlotFile proc.
            Filename$ = txtOffset.Text
            If mnPlot Then PlotFile Filename$
         Case TIn, TInScan, VIn, VIn32
            PlotAcquiredData
         Case FileRead
            Filename$ = txtOffset.Text
            If mnPlot Then PlotFile Filename$
      End Select
   End If
   If ErrorDetected% = 999 Then
      'bad setup parameters
      mnuEvalData_Click
   End If
   DoEvents

End Sub

Private Sub ErrorCount()
   
   PlotThis% = mnPlot
   mnPlot = False
   mlAvgVal = Val(txtCount.Text)
   ResetData

   'set number of samples to at least 1k to establish
   'avg value (msSettleVal)
   mlCount = mlAvgVal
   If mlAvgVal < 1000 Then mlCount = 1000
   AnalogIn
   If gnCancel Or mnCancel Then
        gnCancel = False
        mnCancel = False
        cmdStop_Click
        Exit Sub
   End If
   ReDim AvgData%(0, mlCount)
   If Not IsEmpty(mlHandle) Then
      ULStat = cbWinBufToArray(mlHandle, AvgData%(0, 0), 0, mlCount)
   End If
   
   For SampleNum& = 0 To mlCount - 1
      DataCum& = DataCum& + IntValToULong(AvgData%(0, SampleNum&))
   Next SampleNum&
   
   msSettleVal = DataCum& / mlCount
   mnPlot = PlotThis%
   'check results from 1 sample up to txtCount.txt
   
   NumPoints& = mlAvgVal ' / mlSettleStep
   SetBlockSize NumPoints&, False
   SetSettleStep -1, mfRate
   BoardName = mnuBoard(mnBoardIndex).Caption
   mnErrorCounting = True
   
   SampleSpan$ = "From 1 to " & Format$(mlAvgVal, "0") & " samples"
   BoardName = mnuBoard(mnBoardIndex).Caption
   'SettlingTo$ = "Settling to +FS"
   If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
      RateVal = mfRateReturned / 1000 & "Hz"
   Else
      RateVal = mfRateReturned / 1000 & "kHz"
   End If
   SetDetails mnFuncType, mnFirstChan, mnLastChan - mnFirstChan, RateVal, SampleSpan$, mnRange, BoardName, SettlingTo$
   
   For i& = 1 To mlAvgVal
      mlCount = i&
      AnalogIn
      If gnCancel Or mnCancel Then
           cmdStop_Click
           gnCancel = False
           mnCancel = False
           Exit For
      End If
   Next i&
   mlCount = Val(txtCount.Text)
   mnErrorCounting = False

End Sub

Private Sub FileScan()

   Filename$ = txtOffset.Text
   CBRate& = mfRate
   If mnFuncType = FilePretrig Then
      mlPreTrigCount = Val(txtAmpl.Text)
      mlTotalCount = mlCount
      Do
         ULStat = cbFilePretrig(mnBoardNum, mnFirstChan, mnLastChan, mlPreTrigCount, mlTotalCount, CBRate&, mnRange, Filename$, mvOptions)
      Loop While ULStat = TRIGSTATE
      If SaveFunc(Me, FilePretrig, ULStat, mnBoardNum, mnFirstChan, mnLastChan, mlPreTrigCount, mlTotalCount, CBRate&, mnRange, Filename$, mvOptions, A10, A11, 0) Then Exit Sub
   Else
      ULStat = cbFileAInScan(mnBoardNum, mnFirstChan, mnLastChan, mlCount, CBRate&, mnRange, Filename$, mvOptions)
      If SaveFunc(Me, FileAInScan, ULStat, mnBoardNum, mnFirstChan, mnLastChan, mlCount, CBRate&, mnRange, Filename$, mvOptions, A9, A10, A11, 0) Then Exit Sub
   End If
   mfRateReturned = CBRate&
   'ULStat = cbGetStatus(mnBoardNum, Status%, CurCount&, CurIndex&)
   If mnuGetStatus.Checked Then
      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
      If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      If Not (ULStat = 0) Then Exit Sub
   End If

   DevActive% = (Status% = RUNNING)
   If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
   If DevActive% Then
      If mnuGetStatus.Checked Then tmrCheckStatus.ENABLED = True
      cmdStop.Visible = True
   End If

End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   mbManualStop = True
   If Button = 2 Then
      'if right click, disable timer
      cmdStop_Click
      mnuLoop.Checked = False
      cmdStop.Visible = False
   End If

End Sub

Private Sub Form_Activate()
   
   UpdateMainStatus ""
   If Not (mnLoading Or mnUnloading) Then mnRefreshProps = True
   mnLoading = False
   mnCalledLocally = False

End Sub

Private Sub Form_DblClick()

   UpdateStatBar True

End Sub

Private Sub Form_Initialize()
   
   mnLoading = True
   Me.mnuLibrary(MSGLIB).ENABLED = frmMain.mnuLibrary(MSGLIB).ENABLED
   Me.mnuLibrary(UNILIB).ENABLED = frmMain.mnuLibrary(UNILIB).ENABLED
   Me.mnuLibrary(NETLIB).ENABLED = frmMain.mnuLibrary(NETLIB).ENABLED
   If mnLibType = MSGLIB Then
      mlReadTimeout = 5000
      mlWriteTimeout = 1000
      If Not mnMessaging Then Exit Sub
      If mnFormInitialized Then
         Select Case mnFormType
            Case ANALOG_IN
               NumProps% = GetAIProps(msBoardName, MsgLibrary, PropList)
               UpdateFormProps PropList, NumProps%
               QChanged% = ConfigureMsgQueue()
         End Select
      End If
   End If
   mlTrapCount = -1
   mnCountIterations = 1
   mnInvertCal = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If (Shift And 4) = 4 Then
      Select Case KeyCode
         Case 49  'Alt 1
            txtLowChan.SetFocus
         Case 50  'Alt 2
            txtHighChan.SetFocus
         Case 51  'Alt 3
            txtRate.SetFocus
         Case 52  'Alt 4
            txtCount.SetFocus
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
         Case Asc("G")
            cmdGo = True
         Case Asc("Q")
            cmdStop = True
         Case Else
            NoActiveShortcut% = True
      End Select
      If TypeOf ActiveControl Is TextBox Then
         ActiveControl.SelStart = 0
         ActiveControl.SelLength = Len(ActiveControl.Text)
      End If
   ElseIf (Shift And 2) = 2 Then
      Select Case KeyCode
         Case 13  'Enter
            cmdGo = True
         Case Asc("R")  'Ctl R
            'resets all option menus to default
            ResetOpts
         Case Asc("X")  'Ctl X
            'program run by SendKeys
            gnXternalCtl = True
         Case Else
            NoActiveShortcut% = True
      End Select
   Else
      Select Case KeyCode
         Case 118 'F7 - set default child size
            Me.Height = 2350
            Me.Width = 6435
         Case 120 'F9 - set height to 1/3 screen
            mfmUniTest.Height = Screen.Height / 3
         Case 122 'F11 - set to screen bottom
            mfmUniTest.Move 0, Screen.Height - mfmUniTest.Height, Screen.Width
         Case 123 'F12 - set to screen top
            mfmUniTest.Move 0, 0, Screen.Width
         Case Else
            NoActiveShortcut% = True
      End Select
   End If
   'If NoActiveShortcut% Then KeyCode = 0

End Sub

Private Sub Form_Load()

#If NETOPS Then
   Set NetLibrary = New ULNet.ULNetClass
   For board% = 0 To gnNumBoards - 1
      BoardNumber& = gnBoardEnum(board%)
      BoardName$ = NetLibrary.CreateBoard(BoardNumber&)
   Next
   mnThreading = True
#End If
#If MSGOPS Then
   If gnLibType = MSGLIB Then
      Set MsgLibrary = New MBDClass.MBDComClass
   End If
#Else
   Dim MsgLibrary As Object
#End If

   mnLibType = gnLibType
   mnNumBoards = gnNumBoards
   'NumDevs% = NetLibrary.NumberOfDevices
   mnuLibrary(gnLibType).Checked = True
   If Not gbULLoaded Then
      Me.mnuULBuf.Checked = False
      mnuLibrary(UNILIB).Checked = False
      mnuLibrary(UNILIB).ENABLED = False
   End If
   If (gnLibType = UNILIB) And Not gbULLoaded Then
      mnuFunc.ENABLED = False
      cmdGo.ENABLED = False
      mnuEvents.ENABLED = False
      hsbData.ENABLED = False
   End If
   ReDim manChanArray(0), manGainArray(0), manTypeArray(0)
   ReDim manModeArray(0), manDRateArray(0)
   If gnNumBoards > 0 Then
      ReDim manForcedRes(gnNumBoards - 1)
      ReDim masDefaultScanProps(gnNumBoards - 1)
   End If
   For i% = 0 To gnNumBoards - 1
      'If BoardNum% < 0 Then BoardNum% = Abs(BoardNum% + 1)
      If gnLibType = MSGLIB Then
         BoardNum% = i%
         BoardName$ = GetNameOfMsgBoard(BoardNum%)
         SplitName = Split(BoardName$, "::")
         DisplayName$ = SplitName(0)
      Else
         BoardNum% = gnBoardEnum(i%)
         BoardName$ = GetNameOfBoard(BoardNum%)
         DisplayName$ = BoardName$
         Resolution% = GetADResolution(BoardName$, BoardNum%)
      End If
      CurrentName$ = BoardNum% & ") " & DisplayName$
      msDisplayName = DisplayName$
      UseList% = mnuUseList.Checked
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      Else
         mnBoardIndex = 0
         msBoardName = BoardName$
         ConfigureControls
         mnuBipRange_Click (2)
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   Me.cmbDataRate.AddItem "3750"
   Me.cmbDataRate.AddItem "2000"
   Me.cmbDataRate.AddItem "1000"
   Me.cmbDataRate.AddItem "500"
   Me.cmbDataRate.AddItem "100"
   Me.cmbDataRate.AddItem "60"
   Me.cmbDataRate.AddItem "50"
   Me.cmbDataRate.AddItem "30"
   Me.cmbDataRate.AddItem "25"
   Me.cmbDataRate.AddItem "15"
   Me.cmbDataRate.AddItem "10"
   Me.cmbDataRate.AddItem "5"
   Me.cmbDataRate.AddItem "2.5"
   Me.cmbDataRate.ListIndex = 0
   If gnLibType = MSGLIB Then
      mnQTypeOffset = 0
      cmbType.AddItem "DIFF"
      cmbType.AddItem "SE"
      cmbType.ListIndex = 0
   Else
      mnQTypeOffset = 1
      cmbType.AddItem "QONLY"
      cmbType.AddItem "ANALOG"
      cmbType.AddItem "DIGITAL"
      cmbType.AddItem "DIGITAL8"
      cmbType.AddItem "DIGITAL16"
      cmbType.AddItem "CTR"
      cmbType.AddItem "CTR16"
      cmbType.AddItem "CTR32LOW"
      cmbType.AddItem "CTR32HIGH"
      cmbType.AddItem "CJC"
      cmbType.AddItem "TC"
      cmbType.AddItem "ANALOG_SE"
      cmbType.AddItem "ANALOG_DIFF"
      cmbType.AddItem "SETPOINTSTATUS"
      cmbType.AddItem "CTRBANK0"
      cmbType.AddItem "CTRBANK1"
      cmbType.AddItem "CTRBANK2"
      cmbType.AddItem "CTRBANK3"
      cmbType.AddItem "PADZERO"
      cmbType.ListIndex = 0
   End If
   
   mnPlot = True
   mlEventSize = 0
   mlSettleStep = 1000
   mnSettlePos = True
   mvCustomRange = Null
   mnTrigType = TRIGPOSEDGE
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

End Sub

Private Sub Form_Resize()

   lblStatus.Width = ScaleWidth
   lblStatus.Top = ScaleHeight - lblStatus.Height
   If gnInitializing Then
      'if the form is just loading, sets the form type
      'and sets the default function (cbAInScan() for
      'AIn FormType, cbAOut() for AOut FormType)
      mnFormType = (Val("&H" & Tag) And &HF00&) / &H100 ' - 1
      mnThisInstance = Val("&H" & Tag) And &HFF
      mfRate = 1000
      mnuPlotType_Click (0)
      Select Case mnFormType
         Case ANALOG_IN
            'mnFormType = ANALOG_IN
            'mnuFuncArray_Click (1)  'AInScan
            mnuToEng.Caption = "cbToEngUnits()"
            'mnuOpt_Click (3)
            mlCount = 1000
            mlStatusType = AIFUNCTION
         Case ANALOG_OUT
            'mnFormType = ANALOG_OUT
            'mnuFuncArray_Click (0)  'AOut
            mnuToEng.Caption = "cbFromEngUnits()"
            SetOptions
            mlCount = 100
            mlStatusType = AOFUNCTION
      End Select
      mfRateReturned = mfRate
      'mlTotalCountReturn = mlCount
      msTitle = Caption
      'mnuBoard_Click (0)
      gnInitializing = False
   End If

   If fraSetQueue.Visible Then
      fraSetQueue.Top = 0
      fraSetQueue.Left = frmNewAnalog(mnThisInstance).ScaleHeight * 0.04
      fraSetQueue.Width = frmNewAnalog(mnThisInstance).Width * 0.95
      fraSetQueue.Height = frmNewAnalog(mnThisInstance).ScaleHeight - lblStatus.Height
   End If
   'Me.Visible = True
   'mnLoading = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

   If mnuStopBG.Checked Then
      Select Case mnLibType
         Case UNILIB
            If (mnNumBoards > 0) Then
               If Not mbCatastrophicError Then _
                  ULStat = StopBackground520(mnBoardNum, mlStatusType)
            End If
            If WIN32APP And gbULLoaded Then
               'If Not gnScriptRun Then
               SaveFlow% = geErrFlow
               geErrFlow = 0  'don't bother reporting or trapping error here
               DispSave% = gnLocalErrDisp
               gnLocalErrDisp = False
               ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
               If SaveFunc(Me, ErrHandling, ULStat, DONTPRINT, _
                  DONTSTOP, A3, A4, A5, A6, A7, A8, A9, _
                  A10, A11, 0) Then Exit Sub
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
               gnLocalErrDisp = DispSave%
               ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
               If SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, _
                  gnErrHandling, A3, A4, A5, A6, A7, A8, A9, _
                  A10, A11, 0) Then Exit Sub
            End If
         Case NETLIB
            Set NetLibrary = Nothing
            mnThreading = False
         Case MSGLIB
            If Not MsgLibrary Is Nothing Then
               If Not (msScanSupport = "") Then
                  StopMsg$ = msStatusType & ":STOP"
                  MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
                  If Not (mlEventType = 0) Then
                     MsgResult$ = MsgLibrary.DisableEvents(0)
                     x% = SaveMsg(Me, "DisableEvents(0)", MsgResult$)
                  End If
               End If
               If mnMessaging Then
                  If Not (MsgLibrary.DeviceID = "") Then _
                  MsgLibrary.ReleaseDevice (msBoardName)
                  Set MsgLibrary = Nothing
                  mnMessaging = False
               End If
            End If
      End Select
   End If
   
   mnLoop = False
   tmrCheckStatus.ENABLED = False
   tmrDelay.ENABLED = False
   tmrGoLoop.ENABLED = False
   DoEvents
   cmdGo.FontBold = True
   
   If mnuStopBG.Checked Then
      Samples& = 0
      UseWinAPI% = Not mnuULBuf.Checked
      DataHandle& = GenerateData(DataType&, Cycles%, Samples&, Chans%, mvAmplitude, _
      mvOffset, Signal%, NewData%, Channel%, mlFirstPoint, UseWinAPI%)
      SetupData DataHandle&, Samples&, Chans%, DataType&
   
      If mlHandle Then
         If Not UseWinAPI% Then
            ULStat = cbWinBufFree(mlHandle)
            If SaveFunc(Me, WinBufFree, ULStat, mlHandle, _
               A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            mlHandle = 0: mnAcqDataType = vbEmpty
            mlAcqPoints = 0
            PlotBuffer mlHandle, mlCount, mnLastChan - mnFirstChan
         Else
            If FreeMemory(mlHandle) Then
               mlHandle = 0: mlAcqPoints = 0
               mnAcqDataType = vbEmpty
            End If
         End If
      End If
   End If
   
   UnLoadChildForm Me, mnFormType, mnThisInstance
   gnAIOForms = gnAIOForms - 1

End Sub

Private Sub hsbData_Change()

   hsbData_Scroll

End Sub

Private Sub hsbData_Scroll()

   If hsbData.Max = 32767 Then
      DataVal = hsbData.value Xor &H8000
   Else
      DataVal = hsbData.value
   End If
   
   If mnuToEng.Checked Or mnFuncType = VOut Then
      txtData.Text = GetVolts(16, mnRange, CInt(DataVal))
   Else
      txtData.Text = IntValToULong(CInt(DataVal))
   End If
   
   Select Case mnDataType
      Case 0
         mnDataChange = True
         If Not mnFuncType = VOut Then ResetData
         AnalogOut
      Case 4
         If chkBuildArray.value Then
            If Element% = 0 Then
               ReDim Preserve mDAData(mlCount, Chan%)
               txtAmpl.Text = 0
               txtAmpl.Visible = True
            End If
            mDAData(Element%, Chan%) = hsbData.value
            Element% = Element% + 1
            txtAmpl.Text = Element%
            If Element% > mlCount Then
               hsbData.ENABLED = False
               cmdGo.Visible = True
            End If
            DoEvents
         Else
            Element% = 0
         End If
   End Select

End Sub

Private Sub hsbHighChan_Change()

   txtHighChan.Text = hsbHighChan.value
   ResetData

End Sub

Private Sub hsbLowChan_Change()

   txtLowChan.Text = hsbLowChan.value
   ResetData

End Sub

Private Sub SetQueueCount()
   
   EnableQueue% = hsbQCount.value > 0
   QCount& = Val(txtQCount.Text) - 1
   QueueCapable% = (mnQueueConfigurable Or mnQueueSupported)
   
   If Not (QCount& < 0) Then
      If (mnLibType = MSGLIB) And QueueCapable% Then
         QMessage$ = "AISCAN:QUEUE=ENABLE"
         MsgResult$ = MsgLibrary.SendMessage(QMessage$)
         QFail% = SaveMsg(Me, "SendMessage(" & QMessage$ & ")", MsgResult$)
      End If
      ReDim Preserve manChanArray(QCount&)
      ReDim Preserve manGainArray(QCount&)
      ReDim Preserve manTypeArray(QCount&)
      ReDim Preserve manModeArray(QCount&)
      ReDim Preserve manDRateArray(QCount&)
      If QCount& >= mlQCount Then
         For QElemnt& = mlQCount To QCount&
            manTypeArray(QElemnt&) = -1
         Next
      End If
      chkAuto.ENABLED = True
      cmdLoadArray.ENABLED = True
   Else
      chkAuto.value = 0
      chkAuto.ENABLED = False
      If (mnLibType = MSGLIB) And QueueCapable% Then
         QMessage$ = "AISCAN:QUEUE=DISABLE"
         MsgResult$ = MsgLibrary.SendMessage(QMessage$)
         QFail% = SaveMsg(Me, "SendMessage(" & QMessage$ & ")", MsgResult$)
      End If
      mnIOTFunc = False
      mnuScale(0).ENABLED = mnIOTFunc
      mnuScale(1).ENABLED = mnIOTFunc
      mnuScale(2).ENABLED = mnIOTFunc
      mlStatusType = AIFUNCTION
      If (msConfig = "cbDaqInScan()") Or (msConfig = "cbDaqOutScan()") Then UpdateMenu% = True
      TempConfig$ = "cbAInScan()"
      If mnFormType = ANALOG_OUT Then
         mlStatusType = AOFUNCTION
         If msConfig = "cbDaqOutScan()" Then UpdateMenu% = True
         TempConfig$ = "cbAOutScan()"
      End If
      If UpdateMenu% Then msConfig = TempConfig$
      mnuFuncArray(1).Caption = TempConfig$
      UpdateStatBar 0
   End If
   mlQCount = QCount& + 1
   lstElement.Clear
   For Element% = 0 To mlQCount - 1
      lstElement.AddItem Format$(Element%, "0"), Element%
      'manChanArray(CurElement&)
   Next Element%
   cmdLoadArray.ENABLED = True
   'CheckQueue

End Sub

Private Sub hsbQCount_Change()

   txtQCount.Text = hsbQCount.value

End Sub

Private Sub InScan()
   
   PlotThis% = mnPlot
   mnPlot = False
   Looping% = mnuLoop.Checked
   UseLibrary% = mnuULBuf.Checked
   UseWinAPI% = Not mnuULBuf.Checked
   NumChans& = (mnLastChan - mnFirstChan) + 1
   If (Me.mnuGetStatus.Checked) And Not mbCatastrophicError Then
      Select Case mnLibType
         Case UNILIB
            ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
            'added local check 4/6/09
            CurStat% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, _
            CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0)
         Case NETLIB
            ULStat = NetLibrary.GetStatus(Status%, CurCount&, CurIndex&, mlStatusType)
            CurStat% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, _
            CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0)
         Case MSGLIB
            Status% = MsgCheckStatus(Me, msBoardName, _
            MsgLibrary, msStatusType, CurCount&, CurIndex&)
            If Status% = OVERRUN Then ULStat = OVERRUN
            RateMsg$ = "?" & msStatusType & ":RATE"
            MsgResult$ = MsgLibrary.SendMessage(RateMsg$)
            If Not SaveMsg(Me, "SendMessage(" & RateMsg$ & ")", MsgResult$) Then
               ValueLoc& = InStr(1, MsgResult$, "=")
               If ValueLoc& > 0 Then
                  RateString$ = Mid(MsgResult$, ValueLoc& + 1)
                  If IsNumeric(RateString$) Then
                     mfRateReturned = Val(RateString$)
                  End If
               End If
            End If
      End Select
   
   End If
   
   Dim result As VbMsgBoxResult
   If mbEventCount Then
      If mlTrapCount = -1 Then
         If (Not (CurCount& = mlEvalCount)) And (CurCount& > 0) Then
            result = MsgBox("Expected count of " & mlEvalCount & _
               " but a count of " & CurCount& & " returned " & _
               "with a status of " & Status & ".", vbOKCancel, _
               "Count Discrepancy")
            If result = vbCancel Then
               mnCancel = True
               tmrGoLoop_Timer
               Exit Sub
            End If
         End If
      Else
         If (Status% = RUNNING) And (CurCount& > mlEvalCount) Then mbStartTrap = True
         If mbStartTrap Then
            SetCountThenValue mlTrapCount
            If (Not (CurCount& < mlTrapCount)) And (CurCount& > 0) Then
               result = MsgBox("Trapped count of " & mlTrapCount & _
                  " detected: " & CurCount& & " returned " & _
                  "with a status of " & Status & ".", vbOKCancel, _
                  "Count Trapped")
               If result = vbCancel Then
                  mnCancel = True
                  tmrGoLoop_Timer
                  mbStartTrap = False
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
   
   If ULStat = OVERRUN Then
      'should not occur here - this is bug #920 (CAR 254437) - clear it with stop background
      If mnLibType = UNILIB Then
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
         Stat = SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         MsgBox "This code path should not occur. Check CAR 254437.", vbInformation, "CAR 254437"
      End If
   End If
   If Not ULStat = 0 Then
      'should not occur here - this is bug #920 (CAR 254437) - clear it with stop background if minor error
      If Not (mbCatastrophicError Or (ULStat = BADOPTION)) Then
         ReportStatus% = True
         If mnLibType = UNILIB Then
            If (ULStat = NOCBUL32SYS) Then
               Car$ = "581523 or 518488"
            Else
               Car$ = "254437"
            End If
            ULStat = StopBackground520(mnBoardNum, mlStatusType)
            Stat = SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            MsgBox "This code path should not occur. Check CAR " & Car$ & ".", vbInformation, "CAR " & Car$
         End If
      Else
         Exit Sub
      End If
   End If
   If (Not gnScriptSave) And (ReportStatus%) Then
      'reinstated error checking here because overruns early were not caught
         'errors such as TOOFEW are not cleared for cbGetStatus so don't abort on error here
         'If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, _
         'CurCount&, CurIndex&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         'Also, due to bug #920, don't report error here, just log it
         'ErrSave% = gnLocalErrDisp
      'to do -
      Select Case mnLibType
         Case UNILIB
            x% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, _
            CurCount&, CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0)
         Case NETLIB
         Case MSGLIB
      End Select
   End If
   ActiveBuffer% = (Status% = RUNNING)
   If mnuUseIdle.Checked Then ActiveBuffer% = Not (Status% = IDLE)
   If (Not (mlHandle = 0)) And ((Not Looping%) Or (mnDataChange)) And (Not ActiveBuffer%) Then
      If BufFree(Me, mlHandle, UseWinAPI%) Then
         mlHandle = 0: mnAcqDataType = vbEmpty
         mlAcqPoints = 0
      End If
   End If

   TotalCount& = mlCount
   If mnuCal.Checked Or mnTimerTillCount Then
      TotalCount& = mlAvgVal
      If mnLibType = UNILIB Then TotalCount& = mlAvgVal * NumChans&
   End If
   
   If 0 Then   '(mvOptions And EXTMEMORY) = EXTMEMORY (replaced with NOCLEAR)
      mnExtMem = True
      InfoType% = BOARDINFO
      DevNum% = 0
      ConfigItem% = BIDTBOARD
      ULStat = GetConfig520(InfoType%, mnBoardNum, DevNum%, ConfigItem%, ValConfig&)
      ConfigVal% = ValConfig&
      'ULStat = cbGetConfig(InfoType%, mnBoardNum, DevNum%, ConfigItem%, ConfigVal%)
      If SaveFunc(Me, GetConfig, ULStat, InfoType%, mnBoardNum, DevNum%, _
      ConfigItem%, ConfigVal%, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      If ConfigVal% <> -1 Then
         mnMemBoard = ConfigVal%
         MemBoardFound% = True
      End If
      If MemBoardFound% Then
         ULStat = cbMemReset(mnMemBoard)
         If SaveFunc(Me, MemReset, ULStat, mnMemBoard, A2, A3, _
         A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Else
         MsgBox "No memory board found."
      End If
   Else
      If (mlHandle = 0) Then
         mnExtMem = False
         AddBuff% = 0
         If (mnFuncType = APretrig) And (mnuAddPT.Checked) Then AddBuff% = 512
         mlBufferSize = TotalCount& + AddBuff%
         BufferSize& = mlBufferSize
         If BufferSize& = 0 Then BufferSize& = 1
         If mnLibType = NETLIB Then
            mlHandle = NetLibrary.AllocBuffer32(BufferSize&)
            'Stop
         Else
            If mnLibType = MSGLIB Then
               NumChans& = (mnLastChan - mnFirstChan) + 1
               BufferSize& = BufferSize& * NumChans&
               If mnuCal.Checked Then BufferSize& = mlAvgVal * NumChans&
               'count for MSGLIB is per channel
               mlHandle = ScaledBufAlloc(Me, BufferSize&, UseWinAPI%)
               mnAcqDataType = vbDouble: mlAcqPoints = BufferSize&
            Else
               If mnUse64 Then
                  mlHandle = BufAlloc64(Me, BufferSize&, UseWinAPI%)
                  mnAcqDataType = vbCurrency: mlAcqPoints = BufferSize&
               ElseIf mnUse32 Then
                  mlHandle = BufAlloc32(Me, BufferSize&, UseWinAPI%)
                  mnAcqDataType = vbLong: mlAcqPoints = BufferSize&
               Else
                  mlHandle = BufAlloc16(Me, BufferSize&, UseWinAPI%)
                  mnAcqDataType = vbInteger: mlAcqPoints = BufferSize&
               End If
            End If
         End If
         If mlHandle = 0 Then Exit Sub
      End If
   End If
   
   If mnSampleNum = 0 Then
      mnFirstSourceSet = True
   End If
   If Not gnScriptRun Then SetSource False
   
   CBRate& = mfRate
   Do
      If mnLoop Then
         DelayVal& = tmrGoLoop.Interval * gnTimeFactor
         For LoopDelay& = 0 To DelayVal&
            DoEvents
            If Not cmdStop.ENABLED Then Exit Do
         Next LoopDelay&
         If mnTimerStop Then
            Select Case mnLibType
               Case UNILIB
                  ULStat = StopBackground520(mnBoardNum, mlStatusType)
                  Stat = SaveFunc(Me, StopBackground, ULStat, mnBoardNum, _
                  mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
               Case MSGLIB
                  StopMsg$ = msStatusType & ":STOP"
                  MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
                  Stat = SaveMsg(Me, "SendMessage(" & StopMsg$ & ")", MsgResult$)
            End Select
            cmdGo.FontBold = Not cmdGo.FontBold
            DoEvents
         End If
      End If
      If mnFuncType = APretrig Then
         If chkBuildArray.value = 1 Then
            ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
            If SaveFunc(Me, ErrHandling, ULStat, DONTPRINT, DONTSTOP, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
         mlPreTrigCount = Val(txtAmpl.Text)
         mlPTCountReturn = mlPreTrigCount
         mlTotalCount = TotalCount&
         mlTotalCountReturn = mlTotalCount
         Do
            ULStat = cbAPretrig(mnBoardNum, mnFirstChan, mnLastChan, mlPTCountReturn, _
            mlTotalCountReturn, CBRate&, mnRange, mlHandle, mvOptions)
         Loop While ULStat = TRIGSTATE
         If SaveFunc(Me, APretrig, ULStat, mnBoardNum, mnFirstChan, mnLastChan, _
         mlPTCountReturn, mlTotalCountReturn, CBRate&, mnRange, mlHandle, mvOptions, _
         A10, A11, 0) Then
            TrigError = ULStat
            ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
            If Not (TrigError = TOOFEW) Then Exit Sub
         End If
         mnPlot = PlotThis%
         ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
         If SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, gnErrHandling, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Else
         mlTotalCountReturn = TotalCount&
         If mnIOTFunc Then
            mlPTCountReturn = Val(txtAmpl.Text)
            A11 = mlPTCountReturn    'required to preserve scripting value
            ULStat = IOTDaqInScan(mnBoardNum, manChanArray(0), manTypeArray(0), _
            manGainArray(0), mlQCount, CBRate&, mlPTCountReturn, _
            mlTotalCountReturn, mlHandle, mvOptions)
            If SaveFunc(Me, DaqInScan, ULStat, mnBoardNum, manChanArray(0), _
            manTypeArray(0), manGainArray(0), mlQCount, CBRate&, mlPTCountReturn, _
            mlTotalCountReturn, mlHandle, mvOptions, A11, 0) Then Exit Sub
         Else
            Select Case mnLibType
               Case UNILIB
                  ULStat = cbAInScan(mnBoardNum, mnFirstChan, mnLastChan, _
                  mlTotalCountReturn, CBRate&, mnRange, mlHandle, mvOptions)
                  If SaveFunc(Me, AInScan, ULStat, mnBoardNum, mnFirstChan, _
                  mnLastChan, mlTotalCountReturn, CBRate&, mnRange, mlHandle, _
                  mvOptions, A9, A10, A11, 0) Then Exit Sub
               Case NETLIB
                  ULStat = NetLibrary.AnalogInScan(mnFirstChan, mnLastChan, _
                  mlTotalCountReturn, CBRate&, mnRange, mlHandle, mvOptions)
                  If SaveFunc(Me, AInScan, ULStat, mnBoardNum, mnFirstChan, _
                  mnLastChan, mlTotalCountReturn, CBRate&, mnRange, mlHandle, _
                  mvOptions, A9, A10, A11, 0) Then Exit Sub
               Case MSGLIB
                  MsgResult$ = MsgLibrary.SendMessage("AISCAN:START")
                  If SaveMsg(Me, "SendMessage(AISCAN:START)", MsgResult$) Then
                     cmdStop = True
                     mnBufferEmpty = True
                     Exit Sub
                  End If
                  mnBufferEmpty = False
                  'mnAIScanDataRead = False
            End Select
         End If
         mnPlot = PlotThis%
         DoEvents
      End If
      mbCatastrophicError = ((ULStat = BADBOARD) Or _
         (ULStat = DEADADDEV) Or (ULStat = NO_USB_BOARD))
      If Me.mnuGetStatus.Checked Then
         If Not mbCatastrophicError Then
            Select Case mnLibType
               Case UNILIB
                  ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
                  If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, _
                  CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               Case NETLIB
                  mnThreadState = TS_FIRSTCHECK
                  ULStat = NetLibrary.GetStatus(Status%, CurCount&, CurIndex&, mlStatusType)
                  If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, _
                  CurIndex&, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                  If 0 Then
                     functionToCall$ = GetCallAsString(GetStatus, mnBoardNum, _
                     Status%, CurCount&, CurIndex&, mlStatusType)
                     msTempString = functionToCall$
                  End If
               Case MSGLIB
                  Status% = MsgCheckStatus(Me, msBoardName, _
                  MsgLibrary, msStatusType, CurCount&, CurIndex&)
            End Select
         End If
      End If

      Select Case mnLibType
         Case UNILIB
            If ULStat = FREERUNNING Then
               mfRateReturned = 0
            Else
               If CBRate& > 0 Then
                  mfRateReturned = CBRate&
               Else
                  mfRateReturned = mfRate
               End If
            End If
         Case MSGLIB
            If 0 Then
               RateMsg$ = "?" & msStatusType & ":RATE"
               MsgResult$ = MsgLibrary.SendMessage(RateMsg$)
               If Not SaveMsg(Me, "SendMessage(" & RateMsg$ & ")", MsgResult$) Then
                  ValueLoc& = InStr(1, MsgResult$, "=")
                  If ValueLoc& > 0 Then
                     RateString$ = Mid(MsgResult$, ValueLoc& + 1)
                     If IsNumeric(RateString$) Then
                        mfRateReturned = Val(RateString$)
                     End If
                  End If
               End If
            End If
      End Select
      
      If Not mnThreading Then
         DevActive% = (Status% = RUNNING)
         If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
         If DevActive% Then
            If mnuGetStatus.Checked Or mnuGetSubSystemStatus.Checked _
               Then tmrCheckStatus.ENABLED = True
            cmdStop.Visible = True
            cmdStop.ENABLED = True
            mnPlot = False
            If Not gnScriptRun Then mbEval = False
            If gnScriptSave Then
               ScrDelayVal% = CInt((mlCount / (mfRate * ((mnLastChan - mnFirstChan) + 1)))) + 0.6
               Delay$ = Format$(ScrDelayVal%, "0")
               Print #2, "0, 3000, 0, " & Delay$ & ",,,,,,,,,,,"
            End If
         End If
      End If
   Loop While mnLoop
   If mnLoopType = PROGLOOP Then cmdGo.FontBold = True

End Sub

Private Sub lblStatus_DblClick()

   Clipboard.SetText lblStatus.Caption
   MsgBox "Status copied to clipboard"

End Sub

Private Sub lstElement_Click()

   CurElement& = lstElement.ListIndex
   lblChannel.Caption = "Chan " & manChanArray(CurElement&)
   GainLabel$ = GetRangeString(manGainArray(CurElement&))
   chkEnableSPoint.value = 0
   If Not (manTypeArray(CurElement&) < 0) Then
      CurType% = manTypeArray(CurElement&) And &HFF
      SPointStat% = ((manTypeArray(CurElement&) And &H100) = SETPOINT_ENABLE)
      If SPointStat% Then chkEnableSPoint.value = 1
      'cmbType.ListIndex = CurType% + mnQTypeOffset
      cmbType.ListIndex = Choose(CurType% + mnQTypeOffset + 1, 0, 1, 3, 3, _
      6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 2, 5)
      '-1, ANALOG, DIGITAL, DIGITAL8, DIGITAL16, _
      'CTR, CTR16, CTR32LOW, CTR32HIGH, CJC, TC, ANALOG_SE, ANALOG_DIFF, SETPOINTSTATUS, _
      'CTRBANK0, CTRBANK1, CTRBANK2, CTRBANK3, PADZERO)
      If mnQueueConfigurable Then
         If mnQueueChanMode Then
            ChMode% = manModeArray(CurElement&)
            Me.cmbType.ListIndex = ChMode% + mnQTypeOffset
         End If
         If mnQueueDataRate Then
            dr! = manDRateArray(CurElement&)
            For ComboIndex% = 0 To Me.cmbDataRate.ListCount - 1
               ListText$ = cmbDataRate.List(ComboIndex%)
               If Val(ListText$) = dr! Then
                  cmbDataRate.ListIndex = ComboIndex%
                  Exit For
               End If
            Next
         End If
      End If
   Else
      'If mnQueueConfigurable Then
         If cmbType.ListIndex < 0 Then cmbType.ListIndex = 0
      'Else
      '   cmbType.ListIndex = cmbType.ListCount - 1
      'End If
   End If
   lblGain.Caption = GainLabel$
   CheckQueue

End Sub

Private Sub lstElement_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   'lblChannel.Caption = manChanArray(lstElement.ListIndex)
   'lblGain.Caption = GetRangeString(manGainArray(lstElement.ListIndex))

End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash

End Sub

Private Sub mnuAConv_Click()

   mnuAConv.Checked = Not mnuAConv.Checked
   SetVal% = mnuAConv.Checked
   SetConvertData mnBoardNum, SetVal%
   FuncStat = 0
   A1 = mnuAConv.Checked
   If gnScriptSave Then
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SConvData, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuAddPT_Click()

   mnuAddPT.Checked = Not mnuAddPT.Checked
   FuncStat = 0
   A1 = mnuAddPT.Checked
   If gnScriptSave Then
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SAddPTBuf, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuAvg_Click()

   If mnuSettle.Checked Then
      If mlSettleStep = 0 Then mlSettleStep = 1000
      AvgVal = InputBox("Enter value in Hz to step sample rate:", "Set Sample Rate Step", CStr(mlSettleStep))
      If Not (AvgVal = "") Then mlSettleStep = CLng(AvgVal)
   Else
      If mlAvgVal = 0 Then
         mlAvgVal = 100
         mnDataChange = True
         ResetData
      End If
      AvgVal = InputBox("Enter number of points to average:", "Set Average", CStr(mlAvgVal))
      If Not (AvgVal = "") Then
         mlAvgVal = CInt(AvgVal)
         mnDataChange = True
         ResetData
         If mnLibType = MSGLIB Then SetMsgCount
      End If
   End If

End Sub

Private Sub mnuBuff32_Click()

   mnuBuff32.Checked = Not mnuBuff32.Checked
   mnuBuff64.Checked = False
   mnUse32 = mnuBuff32.Checked
   mnUse64 = mnuBuff64.Checked
   If mnUse32 Then
      SetBufferType 2
      Me.mnuFuncArray(0).Caption = "cbAIn32()"
      Me.mnuFuncArray(8).Caption = "cbVIn32()"
   Else
      SetBufferType 1
      Me.mnuFuncArray(0).Caption = "cbAIn()"
      Me.mnuFuncArray(8).Caption = "cbVIn()"
   End If
   If mnuFuncArray(0).Checked Then msConfig = mnuFuncArray(0).Caption
   If mnuFuncArray(8).Checked Then msConfig = mnuFuncArray(8).Caption
   Me.lblStatus.Caption = msStringID & msConfig & " " & msOpt
   
End Sub

Private Sub mnuBuff64_Click()

   mnuBuff64.Checked = Not mnuBuff64.Checked
   mnuBuff32.Checked = Not mnuBuff64.Checked
   mnUse64 = mnuBuff64.Checked
   mnUse32 = False
   If mnUse64 Then
      SetBufferType 5
      Me.mnuFuncArray(0).Caption = "cbAIn64()"
      Me.mnuFuncArray(8).Caption = "cbVIn64()"
   Else
      SetBufferType 1
      Me.mnuFuncArray(0).Caption = "cbAIn()"
      Me.mnuFuncArray(8).Caption = "cbVIn()"
   End If
   If mnuFuncArray(0).Checked Then msConfig = mnuFuncArray(0).Caption
   If mnuFuncArray(8).Checked Then msConfig = mnuFuncArray(8).Caption
   Me.lblStatus.Caption = msStringID & msConfig & " " & msOpt

End Sub

Private Sub mnuBuildData_Click()

   'new data creation form
   If Not (gnNumBoards > 0) Then Exit Sub
   DaqFlexLib% = (mnLibType = MSGLIB)
   UseWinAPIMemFuncs% = Not mnuULBuf.Checked
   BoardMax& = 2 ^ mnResolution
   NumChans = (mnLastChan - mnFirstChan) + 1
   GenDataType = 1
   If mnLibType = MSGLIB Then
      GenDataType = 6
      CountIsPerChan% = True
   End If
   'If (mvOptions And SCALEDATA) = SCALEDATA _
   '   Then mnGenDataType = vbDouble
   If mbDataGenRequired Then _
      Unload frmBuildData
   If mlGenHandle = 0 Then
      frmBuildData.txtSamples.Text = 1000
   Else
      If (mlGenPoints = NumChans) And (mnFuncType = AOutScan) Then
         GenPoints& = Me.txtCount.Text  'mlGenPoints
      ElseIf mlGenPoints = 1 Then
         If mnGenDataType = vbInteger Then
            ReDim OutArray(0, 0) As Integer
            ULStat = LoadArrayFromWinBuf(Me, mlGenHandle, OutArray(), 0, mlGenPoints)
            Amplitude& = IntValToULong(OutArray(0, 0))
            If Amplitude& > BoardMax& Then Amplitude& = BoardMax& / 2
            frmBuildData.txtAmplitude = Amplitude&
            frmBuildData.txtOffset.Text = BoardMax& / 2
         End If
      Else
         GenPoints& = Val(txtCount.Text) 'mlGenPoints
         If CountIsPerChan% Then GenPoints& = mlGenPoints / NumChans
         frmBuildData.txtSamples.Text = GenPoints&
         NumChans = (mnLastChan - mnFirstChan) + 1
         frmBuildData.txtNumChans.Text = Format(NumChans, "0")
         DefaultsSet% = True
      End If
   End If
   UseWinAPI% = Not mnuULBuf.Checked
   GenDataType = Choose(mnGenDataType, 1, 1, 2, 4, 6)
   If IsNull(GenDataType) Then GenDataType = 1
   GenFloats% = ((mvOptions And SCALEDATA) = SCALEDATA) 'Me.mnuToEng.Checked
   
   frmBuildData.SetFormRef Me
   frmBuildData.SetDefaults mnBoardNum, mnResolution, mnRange, _
   GenPoints&, NumChans, GenDataType, UseWinAPI%, _
   CountIsPerChan%, GenFloats%, DaqFlexLib%
   
   frmBuildData.Show False, mfmUniTest
   If mnLibType = MSGLIB Then
      frmBuildData.txtTimeOut.Text = Format(mlWriteTimeout, "0")
      frmBuildData.txtTimeOut.Visible = True
      frmBuildData.lblTimeOut.Visible = True
   End If
   
   Exit Sub
   Unload frmBuildData
   
   If mnFormType = ANALOG_OUT Then
      If mlGenPoints > 0 Then
         TextCount& = Val(txtCount.Text)
         If (TextCount& = 1) Or (TextCount& > mlGenPoints) Then Me.txtCount.Text = Format(mlGenPoints, "0")
         ConfigureControls
      Else
         If Not NumSamples& = -1 Then
            Me.txtCount.Text = "1"
            ConfigureControls
         End If
      End If
   End If
   
End Sub

Private Sub mnuCalInvert_Click()

   mnuCalInvert.Checked = Not mnuCalInvert.Checked
   mnInvertCal = 1
   If mnuCalInvert.Checked Then mnInvertCal = -1
   
End Sub

Private Sub mnuCalLoopback_Click()

   mnuCalLoopback.Checked = Not mnuCalLoopback.Checked
   mnDetailsNeedUpdate = True
   Dev$ = "HP34401"
   If Not GPIBSetDevice(0, Dev$) Then
      mnuCalLoopback.Checked = False
      MsgBox "Couldn't find " & Dev$
      Instance% = Val(Right$(Me.Tag, 2))
   End If
   
   If mnuCalLoopback.Checked Then
      mnSaveAvg = mlAvgVal
      mlAvgVal = 1
      MeterCmd$ = "VOLT:DC:RES MIN"
      GPIBWrite 0, MeterCmd$
   Else
      mlAvgVal = mnSaveAvg
      MeterCmd$ = "*RST"
      GPIBWrite 0, MeterCmd$
      Dev$ = "DP8200"
      If Not GPIBSetDevice(0, Dev$) Then
         mnuCalLoopback.Checked = False
         MsgBox "Couldn't find " & Dev$
         Instance% = Val(Right$(Me.Tag, 2))
      End If
   End If
   
End Sub

Private Sub mnuCustRange_Click()

   'use value only
   SetRange 3, 0
   
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
         mnUse64 = 0
         mnNumBoards = GetNumInstalled()
         If Not MsgLibrary Is Nothing Then
            If Not (msScanSupport = "") Then
               StopMsg$ = msStatusType & ":STOP"
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
            End If
            If mnMessaging Then
               If Not (MsgLibrary.DeviceID = "") Then _
               MsgLibrary.ReleaseDevice (msBoardName)
               mnMessaging = False
            End If
         End If
         mnNumAIChans = 0: mnNumAOChans = 0
      Case NETLIB
         mnNumBoards = GetNumInstalled()
      Case MSGLIB
         mnNumBoards = GetNumMsgBoards()
         If (gnLibType > INVALIDLIB) Then
            If (MsgLibrary Is Nothing) Then
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
         Resolution% = GetADResolution(BoardName$, BoardNum%)
      End If
      CurrentName$ = BoardNum% & ") " & DisplayName$
      msDisplayName = DisplayName$
      UseList% = mnuUseList.Checked
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      Else
         mnBoardIndex = 0
         ConfigureControls
         mnuBipRange_Click (2)
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   If mnNumBoards > 0 Then
      ReDim manForcedRes(mnNumBoards - 1)
      mnuBoard_Click (0)
      If InitNewInstance% Then
         Select Case mnFormType
            Case ANALOG_IN
               NumProps% = GetAIProps(msBoardName, MsgLibrary, PropList)
               UpdateFormProps PropList, NumProps%
               QChanged% = ConfigureMsgQueue()
         End Select
         Select Case mnFuncType
            Case AIn, AIn32
               msStatusType = "AI"
            Case AInScan
               msStatusType = "AISCAN"
               NumProps% = GetAIScanProps(msBoardName, MsgLibrary, PropsList)
               ResetOpts
               UpdateFormProps PropsList, NumProps%
            Case AOut
               msStatusType = "AO"
            Case AOutScan
               msStatusType = "AOSCAN"
         End Select
         mnRefreshProps = True
         'ConfigureMsgQueue
      End If
   Else
      Caption = msTitle & " Board " & mnuBoard(0).Caption
   End If
   gnNumBoards = mnNumBoards
   
End Sub

Private Sub mnuNoRange_Click()

   SetRange -1, 0

End Sub

Private Sub mnuBipRange_Click(Index As Integer)

   SetRange 0, Index
   If mnUseTwoSources Then
      If (mnRange = BIP20VOLTS) And (mnuCal.Checked) Then
         If Not mnCalWarn Then
            MsgBox "Use two DP8200s at the same GPIB address connected with opposite polarities, one to CH IN LO, the other to CH IN HI", , "Two Sources Required"
            mnCalWarn = True
         End If
      ElseIf (Not mnRange = BIP20VOLTS) And (mnuCal.Checked) Then
         If mnCalWarn Then
            MsgBox "Remove DP8200 from CH IN LO and connect CH IN LO to GND", , "Remove One Source"
            mnCalWarn = False
         End If
      End If
   End If

End Sub

Private Sub mnuBoard_Click(Index As Integer)

   Dim OtherTrigForms As Form
   Dim EventConflict As Boolean
   Dim eventForm As Form
   
   mbCatastrophicError = False
   If Not (gnNumBoards > 0) Then Exit Sub
   If Not (mnLibType = MSGLIB) Then
      BoardNum% = gnBoardEnum(Index)
   Else
      BoardNum% = Index
   End If
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
   If BoardNum% < 0 Then
      BoardNum% = Abs(BoardNum% + 1)
      ResetOpts
      mnuContPlot.ENABLED = True
   Else
      If Not (mnuContPlot.Checked = (mnuOpt(1).Checked And mnuContPlot.Checked)) Then
         mnuContPlot_Click
         mnuContPlot.ENABLED = Not ((mlEventType And ON_DATA_AVAILABLE) = ON_DATA_AVAILABLE)
      End If
   End If
   If Not gnInitializing Then mnuBoard(mnBoardIndex).Checked = False
   mnBoardNum = BoardNum%
   mnBoardIndex = Index
   mnuBoard(mnBoardIndex).Checked = True
   UseList% = mnuUseList.Checked
   Select Case mnLibType
      Case UNILIB
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
      Case MSGLIB
         If mnMessaging Then
            If Not (MsgLibrary.DeviceID = "") Then _
            MsgLibrary.ReleaseDevice (msBoardName)
            mnMessaging = False
            mbGSMConflictReported = False
         End If
         BoardName$ = GetNameOfMsgBoard(mnBoardNum)
         If Not BoardName$ = "None Installed" Then
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
               mnRefreshProps = True
               If mnFormType = ANALOG_IN Then
                  Component$ = "AISCAN"
                  msScanSupport = MsgLibrary.GetSupportedMessages(Component$)
                  x$ = MsgLibrary.SendMessage("@AISCAN:QUEUECONFIG")
                  mnQueueConfigurable = Not (InStr(1, x$, "NOT_SUPPORTED") > 0)
                  mnQueueChanMode = (InStr(1, x$, "CHMODE") > 0)
                  mnQueueDataRate = (InStr(1, x$, "DATARATE") > 0)
                  If Not mnQueueConfigurable Then
                     QueueResult$ = MsgLibrary.SendMessage("@AISCAN:QUEUELEN")
                     mnQueueSupported = Not (InStr(1, QueueResult$, "NOT_SUPPORTED") > 0)
                  Else
                     mnQueueSupported = False
                  End If
                  If mnQueueConfigurable Or mnQueueSupported Then
                     If Not (InStr(1, msScanSupport, "AISCAN:QUEUE") > 0) Then
                        MsgBox "GetSupportedMessages indicates no queue support " & _
                        "but Reflection indicates queue is supported." & vbCrLf & _
                        "Assuming queue is supported.", vbCritical, _
                        "Conflicting Device Properties"
                     End If
                  End If
                  Component$ = "AI"
                  msAiSupport = MsgLibrary.GetSupportedMessages(Component$)
                  mnTempSupport = (InStr(1, msAiSupport, "CJC") > 0) Or (InStr(1, msAiSupport, "SENSOR=TC") > 0)
                  Scales% = mnuScale.Count - 1
                  For ScaleMenu% = 0 To 2 'Scales%
                     mnuScale(ScaleMenu%).ENABLED = mnTempSupport
                  Next
                  If msScanSupport = "" Then
                     If Not msAiSupport = "" Then
                        mnuFuncArray_Click (0)
                     Else
                     End If
                  End If
               Else
                  Component$ = "AOSCAN"
                  msScanSupport = MsgLibrary.GetSupportedMessages(Component$)
               End If
            Else
               Exit Sub
            End If
         End If
   End Select
   msBoardName = BoardName$
   msDisplayName = DisplayName$
   If mnRefreshProps And (mnLibType = MSGLIB) Then
      lpFileName$ = "UniBoards.ini"
   
      'check for custom strings for MsgDaq
      lpApplicationName$ = DisplayName$
      lpKeyName$ = "ExtClockOn"
      lpDefault$ = "ENABLE"
      nSize% = 32
      CustomParam$ = Space$(nSize%)
      StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
      msClockOn = Left$(CustomParam$, StringSize%)
      
      lpKeyName$ = "ExtClockOff"
      lpDefault$ = "DISABLE"
      nSize% = 32
      CustomParam$ = Space$(nSize%)
      StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
      msClockOff = Left$(CustomParam$, StringSize%)
      Select Case mnFuncType
         Case AIn, AIn32
            msStatusType = "AI"
            NumProps% = GetAIProps(msBoardName, MsgLibrary, PropsList)
            UpdateFormProps PropsList, NumProps%
         Case AInScan
            msStatusType = "AI"
            NumProps% = GetAIProps(msBoardName, MsgLibrary, PropsList)
            UpdateFormProps PropsList, NumProps%
            msStatusType = "AISCAN"
            NumProps% = GetAIScanProps(msBoardName, MsgLibrary, PropsList)
            If masDefaultScanProps(mnBoardIndex) = "" Then
               For PropIndex% = 0 To NumProps%
                  masDefaultScanProps(mnBoardIndex) = _
                  masDefaultScanProps(mnBoardIndex) & Separator$ & PropsList(PropIndex%)
                  Separator$ = ";"
               Next
            End If
            ResetOpts
            UpdateFormProps PropsList, NumProps%
         Case AOut
            msStatusType = "AO"
            NumProps% = GetAOProps(msBoardName, MsgLibrary, PropsList)
            UpdateFormProps PropsList, NumProps%
         Case AOutScan
            msStatusType = "AOSCAN"
      End Select
   End If
   'find out if the custom trigger ini file exists
   mnCustomInfo = False
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = "CustomTrigSet"
   lpKeyName$ = "FileExists"
   lpDefault$ = "0"
   nSize% = 4   'this will either be 0 or -1
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   If CustomParam$ = "-1" Then mnCustomInfo = True

   'see if this device is simultaneous input
   lpApplicationName$ = DisplayName$
   lpKeyName$ = "SimultaneousIn"
   lpDefault$ = "0"
   nSize% = 4   'this will either be 0 or -1
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   mnSimIn = (Not CustomParam$ = "0")

   InitMathModule DisplayName$
   InitLocModule DisplayName$
   FuncStat = 0
   A1 = DisplayName$
   If gnScriptSave And (Not gnInitializing) Then
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetBoardName, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   SetResolution
   If mnFormType = ANALOG_IN Then
      If Not (mnLibType = MSGLIB) Then
         CurRange& = GetManualRange(DisplayName$, mnBoardNum, FixedRange&)
      Else
         ChanString$ = "{" & Format(Chan%, "0") & "}"
         RangeString$ = ":RANGE"
         If mnFuncType = AInScan Then
            RangeMsg$ = "?" & msStatusType & RangeString$ & ChanString$
         Else
            StatType$ = "?" & msStatusType
            If msStatusType = "" Then StatType$ = "?AI"
            RangeMsg$ = StatType$ & ChanString$ & RangeString$
         End If
         MsgResult$ = MsgLibrary.SendMessage(RangeMsg$)
         If Left(MsgResult$, 36) = "Warning- The device does not support" Then
            lpApplicationName$ = DisplayName$
            lpKeyName$ = "FixedAInRange"
            lpDefault$ = "0"
            nSize% = 4   'this will be one of the gain codes
            CustomParam$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
            CustomParam$ = Left$(CustomParam$, StringSize%)
            FixedRange& = Val(CustomParam$)
         Else
            RangeRes = Split(MsgResult$, "=")
            If UBound(RangeRes) > 0 Then RangeString$ = RangeRes(1)
            FixedRange& = GetRangeCodeFromMsg(RangeString$)
         End If
      End If
      lpKeyName$ = "DefaultAInRange"
   Else
      Select Case mnLibType
         Case UNILIB
            CurRange& = GetManualAOutRange(DisplayName$, mnBoardNum, FixedRange&)
         Case MSGLIB
            RangeString$ = ":RANGES"
            If mnFuncType = AOutScan Then
               RangeMsg$ = "@" & msStatusType & RangeString$ & ChanString$
            Else
               StatType$ = "@" & msStatusType
               If msStatusType = "" Then StatType$ = "@AO"
               RangeMsg$ = StatType$ & ChanString$ & RangeString$
            End If
            MsgResult$ = MsgLibrary.SendMessage(RangeMsg$)
            FixedRange& = -2
            If InStr(1, MsgResult$, "FIXED%") > 0 Then
               GainArray = Split(MsgResult$, "%")
               If UBound(GainArray) > 0 Then
                  FixedRangeString$ = GainArray(1)
                  FixedRange& = GetRangeCodeFromMsg(FixedRangeString$)
                  DontSetRange% = True
                  mlFixedOutRange = FixedRange&
               End If
            End If
            If FixedRange& = -2 Then
               ChanString$ = "{" & Format(Chan%, "0") & "}"
               RangeString$ = ":RANGE"
               If mnFuncType = AOutScan Then
                  RangeMsg$ = msStatusType & RangeString$ & ChanString$
               Else
                  StatType$ = "?" & msStatusType
                  If msStatusType = "" Then StatType$ = "?AO"
                  RangeMsg$ = StatType$ & ChanString$ & RangeString$
               End If
               MsgResult$ = MsgLibrary.SendMessage(RangeMsg$)
               If Left(MsgResult$, 36) = "Warning- The device does not support" Then
               Else
                  'to do - add code when programmable output range devices are available
               End If
            Else
               mlFixedOutRange = FixedRange&
            End If
      End Select
      lpKeyName$ = "DefaultAOutRange"
   End If
   If FixedRange& = NOTUSED Then
      'see if this device requires default range other than ±5V
      lpApplicationName$ = DisplayName$
      lpDefault$ = "0"
      nSize% = 4   'this will be one of the gain codes
      CustomParam$ = Space$(nSize%)
      StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
      CustomParam$ = Left$(CustomParam$, StringSize%)
      DefaultRange& = Val(CustomParam$)
   Else
      DefaultRange& = FixedRange&
   End If
   SelectRangeUsingCode DefaultRange&  'If Not DontSetRange% Then
   mnFixedRange = FixedRange&
   CustomRange = GetCustomRange(DisplayName$)
   If Not (CustomRange = "") Then
      'undefined range
      mvCustomRange = GetCustomRange(DisplayName$)
   Else
      mvCustomRange = Null
   End If
   SetControlRes
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

Private Sub mnuBoardTitle_Click()

   mnuBoardTitle.Checked = Not mnuBoardTitle.Checked
   If mnuBoardTitle.Checked Then TitleType% = 1
   PlotOptions% = mnuRetainPlot.Checked
   SetPlotOptions PlotOptions%, TitleType%
   If gnScriptSave Then
      FuncStat = 0
      A1 = PlotOptions%
      A2 = TitleType%
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetPlotOpts, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuBufInfo_Click()

   ShowBufferInfo 0
   mnCalledLocally = True
   DoEvents

End Sub

Private Sub mnuCal_Click()

   mnuCal.Checked = Not mnuCal.Checked
   
   If mnuCal.Checked Then
      If Not mlCalHandle = 0 Then
         ULStat = cbWinBufFree(mlCalHandle)
         mlCalHandle = 0
      End If
      If mnuCal.Checked Then
         mlCalHandle = cbWinBufAlloc32(mlCount + 1)
         mlAcqPoints = mlCount
         mnAcqDataType = vbSingle
         BufType% = 4 Or &H20
         SetBufferType BufType%
      End If
      mnPlotType = PLOT_CHUNKS
      SetPlotType mnPlotType, Me
      If mlCalInterval = 0 Then mlCalInterval = 1000
      mfmUniTest.Picture1.Visible = False
      If Not gnScriptRun Then OpenGPIB
      If mnFormType = ANALOG_IN Then
         'mnCalConst = 1
         Dev$ = "DP8200"
         'BoardName$ = GetNameOfBoard(mnBoardNum)
         If Not gnScriptRun Then
            mnuLoop.Checked = True  'execute Go by timer event
            mnLoopType = TIMERLOOP
         End If
         tmrGoLoop.Interval = mlCalInterval   'interval for Go
         tmrDelay.Interval = 500
         If mlAvgVal = 0 Then mlAvgVal = 100
         Select Case msDisplayName
            Case "PMD-1208LS", "miniLAB 1008", "PMD-1208FS"
               mnUseTwoSources = True
            Case "USB-1208LS", "USB-1208FS", "USB-1408FS"
               mnUseTwoSources = True
            Case "USB-1208HS-4AO", "USB-1208HS-2AO", "USB-1208HS"
               mnUseTwoSources = True
         End Select
         If mnUseTwoSources Then
            If (mnRange = BIP20VOLTS) And (mnuCal.Checked) Then
               If Not mnCalWarn Then
                  MsgBox "Use two DP8200s at the same GPIB " & _
                  "address connected with opposite polarities, " & _
                  "one to CH IN LO, the other to CH IN HI", , "Two Sources Required"
                  mnCalWarn = True
               End If
            ElseIf (Not mnRange = BIP20VOLTS) And (mnuCal.Checked) Then
               If mnCalWarn Then
                  MsgBox "Remove DP8200 from CH IN LO and connect CH IN LO to GND", _
                  , "Remove One Source"
                  mnCalWarn = False
               End If
            End If
         End If
         Me.txtAmpl.Text = 2 ^ mnResolution
         Me.txtOffset.Text = 0
         ConfigureControls
         If mnLibType = MSGLIB Then SetMsgCount
      Else
         mnCalConst = 0
         If Not gnScriptRun Then Dev$ = InputBox$ _
         ("Enter meter type (F8840 or 34401): ", _
         "Meter Used", "F8840")
         If Dev$ = "" Then Dev$ = "F8840"
         msDevName = Dev$
         If Not gnScriptRun Then
            mnuLoop.Checked = True  'execute Go by timer event
            mnLoopType = TIMERLOOP
         End If
         tmrGoLoop.Interval = mlCalInterval   'interval for Go
         tmrDelay.Interval = 500
         If mlAvgVal = 0 Then mlAvgVal = 100
         If gnScriptSave Then
            A1 = msDevName
            FuncStat = 0
            For ArgNum% = 1 To 14
               ArgVar = Choose(ArgNum%, Me.Tag, SSetDevName, _
               FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
               If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
                  PrintString$ = PrintString$ & ", "
               Else
                  PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
               End If
            Next
            Print #2, PrintString$; Format$(AuxHandle, "0")
         End If
         txtAmpl.Text = (2 ^ mnResolution) - 1
         tmrDelay.Interval = 2000
         mlAvgVal = 4
      End If
      If Not gnScriptRun Then
         If Not GPIBSetDevice(0, Dev$) Then
            mnuCal.Checked = False
            MsgBox "Couldn't find " & Dev$ & ". If you aren't using " & _
            "MCC devices for testing, make sure app is not configured " & _
            "for 'Force MCC Control'."
            Instance% = Val(Right$(Me.Tag, 2))
            'CloseGPIB Instance%
         End If
         ConfigureControls
         cmdGo.Visible = mnuCal.Checked
      End If
      If Not gnScriptRun Then
         If InStr(1, msDevName, "34401") Then
            Setup$ = "CONF:VOLT:DC"
            If mnRange > 199 Then Setup$ = "CONF:CURR:DC"
            LinkCommand 0, "L1"  'set up for SCPI mode (not 8840 compatible)
            LinkStart 0
            LinkCommand 0, Setup$
            LinkStart 0
            DoEvents
         End If
      End If
      'mnuFuncArray_Click (0)     'AIn() or AOut()
      'If Not mnuPlotType(SINGLE_POINT).Checked Then mnuPlotType_Click (SINGLE_POINT)
   Else
      'mnuCal.Checked = False
      If Not mlCalHandle = 0 Then
         ULStat = cbWinBufFree(mlCalHandle)
         mlCalHandle = 0
         mlAcqPoints = mlCount
         mnAcqDataType = vbSingle
      End If
      mnCalConst = 0
      Instance% = Val(Right$(Me.Tag, 2))
      If Not gnScriptRun Then CloseGPIB Instance%
      mlAvgVal = 0
      tmrDelay.Interval = 500
      mnuLoop.Checked = False
      tmrGoLoop.Interval = 1000
      mnuFuncArray_Click (1)
      If Not mnuPlotType(VOLTS_VS_TIME).Checked _
      Then mnuPlotType_Click (VOLTS_VS_TIME)
   End If
   cmdStop.Visible = mnuLoop.Checked
   A1 = mnuCal.Checked
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SCalCheck, FuncStat, _
         A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuCalData_Click()

   mnuCalData.Checked = Not mnuCalData.Checked
   SetVal% = mnuCalData.Checked
   SetCalData mnBoardNum, mnRange, SetVal%
   A1 = mnuCalData.Checked
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SCalData, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuClearOnGo_Click()

   mnuClearOnGo.Checked = Not mnuClearOnGo.Checked

End Sub

Private Sub mnuClearOpts_Click()

   ResetOpts

End Sub

Private Sub mnuClearPlot_Click()

   ClearPlot

End Sub

Private Sub mnuClose_Click()

   Unload Me

End Sub

Private Sub mnuContPlot_Click()

   mnuContPlot.Checked = Not mnuContPlot.Checked
   A1 = mnuContPlot.Checked
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SContPlot, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuConvPT_Click()

   mnuConvPT.Checked = Not mnuConvPT.Checked
   SetVal% = mnuConvPT.Checked
   SetConvertData mnBoardNum, SetVal%
   A1 = mnuConvPT.Checked
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SConvPT, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuCurRange_Click(Index As Integer)

   SetRange 2, Index

End Sub

Private Sub mnuDrift_Click()

   frmSetTimer.Caption = "Drift Measurement Configuraton"
   frmSetTimer.optTimerMode(0).Visible = False
   frmSetTimer.optTimerMode(1).Visible = False
   frmSetTimer.chkEnableTimer.Caption = "Enable drift test"
   frmSetTimer.Label1.Caption = "Timer Interval"
   If mnuDrift.Checked Then
      frmSetTimer.chkEnableTimer.value = 1
      TestInProgress% = True
   End If
   frmSetTimer.txtInterval.Text = mlSettleStep
   frmSetTimer.optTimerMode(1).value = True
   frmSetTimer.Show 1
   mnuDrift.Checked = Not (frmSetTimer.chkEnableTimer.value = 0)
   'mnSettlePos = frmSetTimer.optTimerMode(0).Value
   'mlSettleStep = Val(frmSetTimer.txtInterval.Text)
   mnuLoop.Checked = (frmSetTimer.chkEnableTimer.value = 1)
   mnLoopType = TIMERLOOP
   If frmSetTimer.optLoopType(1).value Then mnLoopType = PROGLOOP
   mnLoop = frmSetTimer.optLoopType(1).value And (frmSetTimer.chkEnableTimer.value = 1)
   Unload frmSetTimer

   If LoopRate& > 0 Then tmrGoLoop.Interval = LoopRate&
   Unload frmSetTimer
   If mnuDrift.Checked Then
      cmdStop.Visible = True
      If Not mnuPlotType(SINGLE_POINT).Checked Then mnuPlotType_Click (SINGLE_POINT)
      PlotThis% = mnPlot
      'mnPlot = False
      
      'get first sample to establish initial value
      mfRate = Val(txtRate.Text)
      mlAvgVal = Val(txtCount.Text) / NumChans%
      AnalogIn
      If gnCancel Or mnCancel Then
           gnCancel = False
           mnCancel = False
           cmdStop_Click
           Exit Sub
      End If
      ReDim AvgData%(NumChans% - 1, mlAvgVal)
      If Not IsEmpty(mlHandle) Then
         ULStat = cbWinBufToArray(mlHandle, AvgData%(0, 0), 0, mlCount - 1)
      End If
      
      For SampleNum& = 0 To mlAvgVal - 1
         DataCum& = DataCum& + IntValToULong(AvgData%(mnLastChan, SampleNum&))
      Next SampleNum&
      
      'this is the initial sample to which the rest will be compared
      msSettleVal = DataCum& / mlAvgVal
      'mnPlot = PlotThis%
   Else
      mfRate = Val(txtRate.Text)
      'Instance% = Val(Right$(Me.Tag, 2))
      'If Not gnScriptRun Then CloseGPIB Instance%
      If Not mnuPlotType(VOLTS_VS_TIME).Checked Then mnuPlotType_Click (VOLTS_VS_TIME)
      'SetSettleStep 0, mfRate
      cmdStop.Visible = False
   End If

End Sub

Private Sub mnuErrCount_Click()

   mnuErrCount.Checked = Not mnuErrCount.Checked
   If mnuErrCount.Checked Then
      If Not mnuPlotType(SINGLE_POINT).Checked Then mnuPlotType_Click (SINGLE_POINT)
   Else
      If Not mnuPlotType(VOLTS_VS_TIME).Checked Then mnuPlotType_Click (VOLTS_VS_TIME)
   End If

End Sub

Private Sub mnuEvalData_Click()

   NumChans% = (mnLastChan - mnFirstChan) + 1
   If (mnFuncType = TIn) Or (mnFuncType = VIn) Then
      frmEvalData.txtVMin.Text = "20.0"
      frmEvalData.txtVMax.Text = "30.0"
      frmEvalData.fraDeltaT.ENABLED = False
      frmEvalData.chkMaxDelta.ENABLED = False
      frmEvalData.chkShowMinMax.ENABLED = False
   Else
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
   End If

   ParamValue = GetEvalParams(-1)
   frmEvalData.chkSamplePairs.value = 0
   If ParamValue Then frmEvalData.chkSamplePairs.value = 1
   For i& = 0 To 18
      CurParam& = 2 ^ i&
      ParamValue = GetEvalParams(CurParam&)
      Select Case CurParam&
         Case geENABLEEVAL
         Case geEVALDELTA
            DeltaEval = ParamValue
            frmEvalData.chkMaxDelta.value = 0
            If DeltaEval Then frmEvalData.chkMaxDelta.value = 1
         Case geEVALMINDELTA
            DeltaEval = ParamValue
            frmEvalData.chkMinDelta.value = 0
            If DeltaEval Then frmEvalData.chkMinDelta.value = 1
         Case geEVALINWINDOW
            WinEval = ParamValue
            frmEvalData.chkInWindow.value = 0
            If WinEval Then frmEvalData.chkInWindow.value = 1
         Case geEVALRANGE
            RangeEval = ParamValue
            frmEvalData.chkMinMaxStop.value = 0
            If RangeEval Then frmEvalData.chkMinMaxStop.value = 1
         Case geSTART
            EvalStart& = ParamValue
            frmEvalData.txtFirstPoint.Text = EvalStart&
         Case geNUMPOINTS
            ECount& = ParamValue
            PerChan& = (mlAcqPoints / NumChans%)
            If (ECount& = 0) Or (ECount& > PerChan&) Then
               'TCount& = Me.txtCount.Text
               ECount& = PerChan&   'TCount& / NumChans%
               InitialRun = True
            End If
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
            frmEvalData.txtDeltaVMin.Text = DvMin&
         Case geINWINDOW
            InWind& = ParamValue
            frmEvalData.txtWindowPercentage.Text = InWind&
         Case geSHOWPASTEMSG
            frmEvalData.chkShowMsg.value = 0
            If ParamValue Or InitialRun Then frmEvalData.chkShowMsg.value = 1
         Case geNUMMSGSAMPS
            If Not InitialRun Then frmEvalData.txtNumMsgSamps.Text = ParamValue
         Case geWRITEFILEERRS
            frmEvalData.chkSaveData.value = 0
            If ParamValue Or InitialRun Then frmEvalData.chkSaveData.value = 1
      End Select
   Next i&
   If mbEventCount Then
      frmEvalData.txtCountTrig.Text = Format(mlEvalCount, "0")
      frmEvalData.chkCountStop.value = 1
      frmEvalData.txtCountThen.Text = mlTrapCount
      frmEvalData.optCountFirst(mnCountFirst).value = True
      frmEvalData.optCountLast(mnCountLast).value = True
      frmEvalData.txtCountIterations.Text = mnCountIterations
   End If
      
   frmEvalData.Show 1
   mbEval = (frmEvalData.chkEnableEval.value = 1)
   mbEventCount = (frmEvalData.chkCountStop.value = 1)
   mlEvalCount = Val(frmEvalData.txtCountTrig.Text)
   If mbEval Then mlEvalCount = frmEvalData.txtNumSamps.Text
   If mbEventCount Then mlTrapCount = frmEvalData.txtCountThen.Text
   mnCountIterations = Val(frmEvalData.txtCountIterations.Text)
   For CountOpt% = 0 To 2
      If frmEvalData.optCountFirst(CountOpt%).value Then mnCountFirst = CountOpt%
      If frmEvalData.optCountLast(CountOpt%).value Then mnCountLast = CountOpt%
   Next
   mbEvalEnabled = mbEval
   Unload frmEvalData
   If gnScriptSave Then
      ThisFunction& = SEvalEnable
      FuncStat = 0
      A1 = mbEval
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, ThisFunction&, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
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
      Else
         EventParam = mlEventSize
      End If
      If Not (EventParam > 0) Then
         EnablingEvent = False
         mnuEvent(Index).Checked = False
      Else
         mlEventSize = EventParam
         EnablingEvent = True
         mnuEvent(Index).Checked = True
         If (mlEventType And ON_DATA_AVAILABLE) = ON_DATA_AVAILABLE Then
            'if already enabled, disable it so it can
            'be re-enabled with the new EventSize
            If MsgEvent Then
               MsgResult$ = MsgLibrary.DisableEvents(MsgEventType&)
               x% = SaveMsg(Me, "DisableEvents(" & _
               Format(MsgEventType&, "0") & ")", MsgResult$)
            Else
               ULStat = UninstallEvent(mnBoardNum, EventType)
               x% = SaveFunc(Me, DisableEvent, ULStat, mnBoardNum, EventType, _
                  A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            End If
         End If
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

Private Sub mnuFileRead_Click()

   Static HoldFunc As Integer

   mnuFileRead.Checked = Not mnuFileRead.Checked
   HoldFunc = Choose(mnFuncType, 0, 1, x%, 3, 2, 4, x%, 5, 6, 7)
   If mnuFileRead.Checked Then
      HoldFunc = mnFuncType
      mnuFuncArray(HoldFunc).Checked = False
      'mnFuncType = FileRead
   Else
      mnuFuncArray(HoldFunc).Checked = True
      'mnFuncType = HoldFunc
   End If
   ConfigureControls

End Sub

Private Sub mnuFuncArray_Click(Index As Integer)

   'find out which index is checked
   For CurIndex% = 0 To mnuFuncArray.Count - 1
      If mnuFuncArray(CurIndex%).Checked Then
         OldMenuIndex% = CurIndex%
         Exit For
      End If
   Next
   mnuFileRead.Checked = False
   If (mnFormType = ANALOG_IN) Then
      mnFuncType = Choose(Index + 1, AIn, AInScan, ATrig, _
      APretrig, FileAInScan, FilePretrig, TIn, TInScan, VIn)
      If mnFuncType = AIn And mnUse32 Then mnFuncType = AIn32
   Else
      mnFuncType = Choose(Index + 1, AOut, AOutScan, VOut)
   End If
   mnuFuncArray(OldMenuIndex%).Checked = False
   mnuFuncArray(Index%).Checked = True
   'to do - change prop funcs to take functype arg
   If (mnLibType = MSGLIB) Then
      If mnMessaging Then
         Select Case mnFuncType
            Case AIn, AIn32
               msStatusType = "AI"
               If (Not gnScriptRun) Then NumProps% = GetAIProps(msBoardName, MsgLibrary, PropsList)
            Case AInScan
               msStatusType = "AISCAN"
               If (Not gnScriptRun) Then
                  NumProps% = GetAIScanProps(msBoardName, MsgLibrary, PropsList)
                  ResetOpts
                  UpdateFormProps PropsList, NumProps%
               End If
            Case AOut
               msStatusType = "AO"
            Case AOutScan
               msStatusType = "AOSCAN"
         End Select
      End If
   End If
   mnuContPlot.Checked = (mnuContPlot.Checked) And (mnFuncType = AInScan)
   msConfig = mnuFuncArray(Index%).Caption
   ConfigureControls
   SetOptions
   ConfigureValidOptions
   mnDetailsNeedUpdate = True
   
   UpdateStatBar False
   FloatType% = (mnFuncType = TIn) Or (mnFuncType = VIn) Or (mnFuncType = TInScan)
   mnuScale(0).ENABLED = FloatType% Or mnIOTFunc Or mnTempSupport
   mnuScale(1).ENABLED = FloatType% Or mnIOTFunc Or mnTempSupport
   mnuScale(2).ENABLED = FloatType% Or mnIOTFunc Or mnTempSupport
   'If (mnFuncType = APretrig) And (mvOptions And EXTMEMORY) Then
   '   Obsolete EXTMEMORY replaced with NOCLEAR
   '   MsgBox "The MegaFifo must be fully populated to use cbAPretrigger() with the EXTMEMORY option."
   'End If
   SetControlRes

End Sub

Private Sub mnuGetStatus_Click()

   mnuGetStatus.Checked = Not mnuGetStatus.Checked
   If mnuGetStatus.Checked Then mnuGetSubSystemStatus.Checked = False
   A1 = mnuGetStatus.Checked
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SGetStatus, _
            FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuGetSubSystemStatus_Click()

   'mnuGetSubSystemStatus.Checked = Not mnuGetSubSystemStatus.Checked
   'If mnuGetSubSystemStatus.Checked Then
   '   Subsystem& = mnFormType
   '   ULStat = cbGetSubSystemStatus(mnBoardNum, Subsystem&, Status%, CurCount&, CurIndex&)
   '   If Not (ULStat = 0) Then
   '      mnuGetSubSystemStatus.Checked = False
   '      mnuGetStatus.Checked = True
   '      mlSubSystem = 0
   '   Else
   '      mnuGetStatus.Checked = False
   '      mlSubSystem = mnFormType
   '   End If
   'End If
   
End Sub

Private Sub mnuSetSEChanMode_Click()

   Dim ConfigurableMode As Boolean
   Dim SetChanDiff As Boolean
   
   If Not LibSupportsFunction(AChanInputMode) Then Exit Sub
   ConfigurableMode = True
   SetChanDiff = False
   ULStat& = cbErrHandling(DONTPRINT, DONTSTOP)
   LocRep% = gnLocalErrDisp
   gnLocalErrDisp = False
   
   ULStat& = cbAInputMode(mnBoardNum, DIFFERENTIAL)
   xx% = SaveFunc(Me, AInputMode, ULStat, mnBoardNum, _
      DIFFERENTIAL, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   If ULStat& = AIINPUTMODENOTCONFIGURABLE Then
      ConfigurableMode = False
   End If
   ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
   gnLocalErrDisp = LocRep%
   frmComposite.Caption = "Select Channels"
   frmComposite.fraChanSel.Visible = True
   frmComposite.chkSetUnselected.Visible = True
   frmComposite.lstChanSelect.Clear
   If SetChanDiff Then
      frmComposite.chkSetUnselected.value = 1
      frmComposite.txtChanList.Text = msBoardName & _
         " reports non-configurable when used with " & _
         "'cbAInputMode'. All chan modes unchanged."
   Else
      frmComposite.txtChanList.Text = "All channels have " & _
         "been set to differential using 'cbAInputMode'."
   End If
   If ConfigurableMode Then
      ULStat& = cbGetConfig(BOARDINFO, mnBoardNum, 0, BINUMADCHANS, ConfigVal&)
      For ChanNum& = 0 To ConfigVal& - 1
         frmComposite.lstChanSelect.AddItem "Channel " & Format(ChanNum&, "0")
      Next
   End If
   frmComposite.Show 1
   ChansAvailable& = frmComposite.lstChanSelect.ListCount
   ChansSelected& = 0
   SetChanDiff = _
      (frmComposite.chkSetUnselected.value = 1)
   For ChanItem& = 0 To ChansAvailable& - 1
      If frmComposite.lstChanSelect.Selected(ChanItem&) Then
         ULStat& = cbAChanInputMode(mnBoardNum, ChanItem&, SINGLE_ENDED)
         If SaveFunc(Me, AChanInputMode, ULStat, mnBoardNum, ChanItem&, _
            SINGLE_ENDED, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
            Exit For
         End If
      ElseIf SetChanDiff Then
         ULStat& = cbAChanInputMode(mnBoardNum, ChanItem&, DIFFERENTIAL)
         If SaveFunc(Me, AChanInputMode, ULStat, mnBoardNum, ChanItem&, _
            DIFFERENTIAL, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
            Exit For
         End If
      End If
   Next
   Unload frmComposite

End Sub

Private Sub mnuSetSEMode_Click()

   If Not LibSupportsFunction(AInputMode) Then Exit Sub
   mnuSetSEMode.Checked = Not mnuSetSEMode.Checked
   SEMode& = Abs(mnuSetSEMode.Checked)
   ULStat& = cbAInputMode(mnBoardNum, SEMode&)
   If SaveFunc(Me, AInputMode, ULStat, mnBoardNum, SEMode&, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
      mnuSetSEMode.Checked = False
   End If
   
End Sub

Private Sub mnuStatusRate_Click()

   Default$ = Format(Me.tmrCheckStatus.Interval, "0")
   SelRate$ = InputBox("Milliseconds between status checks:", "Status Rate", Default$)
   If Not SelRate$ = "" Then
      tmrCheckStatus.Interval = Val(SelRate$)
   End If
   
End Sub

Private Sub mnuUseIdle_Click()

   mnuUseIdle.Checked = Not mnuUseIdle.Checked

End Sub

Private Sub mnuLoop_Click()

   If mnuLoop.Checked Then frmSetTimer.chkEnableTimer.value = 1
   If mnTimerStop Then frmSetTimer.chkTimerStopBG = 1
   If mnDelayRestart Then
      frmSetTimer.chkDelayRestart.value = 1
      frmSetTimer.txtDelayTime.Text = Format(mlDelayTime, "0")
   End If
   
   frmSetTimer.txtInterval.Text = tmrGoLoop.Interval
   frmSetTimer.optTimerMode(1).value = mnTimerTillCount
   If mnLoopType = PROGLOOP Then frmSetTimer.optLoopType(1).value = True
   frmSetTimer.Show 1

   LoopRate& = Val(frmSetTimer.txtInterval.Text)
   If frmSetTimer.optTimerMode(1).value Then
      'mlCount = 1
      mnTimerTillCount = True
      If Me.mnuCal.Checked Then If mlAvgVal = 0 Then mlAvgVal = 100
      'If Not (mnuPlotType(SINGLE_POINT).Checked) Then mnuPlotType_Click (SINGLE_POINT)
      If Not (mnuPlotType(PLOT_CHUNKS).Checked) Then mnuPlotType_Click (PLOT_CHUNKS)
   Else
      mlCount = Val(txtCount.Text)
      mnTimerTillCount = False
      If Not mnuCal.Checked Then
         mlAvgVal = 0
         If mnuPlotType(SINGLE_POINT).Checked Or mnuPlotType(PLOT_CHUNKS).Checked Then mnuPlotType_Click (VOLTS_VS_TIME)
      End If
   End If
   mnTimerStop = (frmSetTimer.chkTimerStopBG = 1)
   mnuLoop.Checked = (frmSetTimer.chkEnableTimer.value = 1)
   mnDelayRestart = (frmSetTimer.chkDelayRestart.value = 1)
   mlDelayTime = Val(frmSetTimer.txtDelayTime.Text)
   mnLoopType = TIMERLOOP
   If frmSetTimer.optLoopType(1).value Then mnLoopType = PROGLOOP
   mnLoop = frmSetTimer.optLoopType(1).value And (frmSetTimer.chkEnableTimer.value = 1)
   Unload frmSetTimer

   If LoopRate& > 0 Then tmrGoLoop.Interval = LoopRate&
   'tmrGoLoop.Enabled = mnuLoop.Checked And (mnLoopType = TIMERLOOP)
   
   cmdStop.Visible = mnuLoop.Checked
   cmdStop.ENABLED = False
   mnDetailsNeedUpdate = True
   
End Sub

Private Sub mnuNextBlock_Click()

   NextBlock (0)
   If gnScriptSave Then
      A1 = 0
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SNextBlock, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuNoise_Click()

   mnuNoise.Checked = Not mnuNoise.Checked
   Calculate% = mnuNoise.Checked
   CalcNoise Calculate%
   A1 = Calculate%
   If gnScriptSave And (Not gnInitializing) Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SCalcNoise, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuOpt_Click(Index As Integer)

   'option menu choices are arranged according to the
   'bit position of its value
   'the script uses the bit position to locate the
   'menu selection - don't change the order
   'new options should be placed at the end

   OriginalOpts& = mvOptions
   
   Select Case Index
      'BURSTIO, SINGLEIO, BLOCKIO, and DMAIO are mutually exclusive
      Case 5
         mnuOpt(7).Checked = False
         mnuOpt(6).Checked = False
         mnuOpt(16).Checked = False
      Case 6
         mnuOpt(7).Checked = False
         mnuOpt(5).Checked = False
         mnuOpt(16).Checked = False
      Case 7
         mnuOpt(6).Checked = False
         mnuOpt(5).Checked = False
         mnuOpt(16).Checked = False
      'Case 12
      '   Obsolete EXTMEMORY replaced by NOCLEAR
      '   If (mnFuncType = APretrig) And (Not mnuOpt(Index).Checked) Then
      '      MsgBox "The MegaFifo must be fully populated to use cbAPretrigger() with the EXTMEMORY option."
      '   End If
      Case 16
         mnuOpt(6).Checked = False
         mnuOpt(5).Checked = False
         mnuOpt(7).Checked = False
   End Select
   mnuOpt(Index).Checked = Not mnuOpt(Index).Checked
   If Not mnScriptCommand Then
      If (mnBoardNum < 0) Then
         If Not (mnuContPlot.Checked = (mnuOpt(5).Checked And mnuContPlot.Checked)) _
         Then mnuContPlot_Click
      Else
         If Not (mnuContPlot.Checked = (mnuOpt(1).Checked And mnuContPlot.Checked)) _
         Then mnuContPlot_Click
      End If
   End If
   mnuContPlot.ENABLED = mnuOpt(1).Checked Or ((mlEventType And ON_DATA_AVAILABLE) = ON_DATA_AVAILABLE)
   If mnuOpt(1).Checked Then
      mnuLoop.Checked = False
      'mnuLoop.Enabled = False
      mnLoop = False
   Else
      mnuLoop.ENABLED = True
   End If
   If Index = mnuOpt.Count - 1 Then
      DefaultValue$ = Format(mlCustomOption, "0")
      varOption = InputBox("Enter custom option code.", "Custom Option", DefaultValue$)
      If IsNumeric(varOption) Then
         mlCustomOption = Val(varOption)
         mnuOpt(Index).Checked = Not (mlCustomOption = 0)
      End If
   End If
   SetOptions
   mnDetailsNeedUpdate = True
   mnRefreshProps = (OriginalOpts& <> mvOptions)
   
   UpdateStatBar False

End Sub

Private Sub mnuPlotChan_Click()
   
   Resp = InputBox("Channel to plot (enter ""a"" for all channels)", "Select Channel to Plot", "a")
   Select Case Resp
      Case "a", "A"
         PlotChan% = -1
      Case ""
         Exit Sub
      Case Else
         PlotChan% = Val(Resp)
   End Select
   SetChannel PlotChan%
   RePlot False
   If gnScriptSave Then
      A1 = PlotChan%
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetPlotChan, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuPlotType_Click(Index As Integer)

   mbJustReadBuffer = True
   If Index = PRINT_TEXT Then
      mnuPlotType(Index).Checked = Not mnuPlotType(Index).Checked
      If mnuPlotType(Index).Checked Then mlBlockSave = GetBlockSize()
      If mnuPlotType(VOLTS_VS_TIME).Checked Or mnuPlotType(DERIVATIVE).Checked Then
         If Not mnuPlotType(PRINT_TEXT).Checked Then SetBlockSize mlBlockSave, False
      End If
      TFVal% = mnuPlotType(Index).Checked
      ShowText TFVal%
      
      x% = SaveFunc(Me, SShowText, 0, 0, TFVal%, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

      If IsEmpty(mlHandle) Or (mlHandle = 0) Then
         'check if there's an output buffer
         If Not ((mlGenHandle = 0) Or (mlGenPoints = 0)) Then
            PlotGenData
            PrintMain "Buffer handle = 0x" & Hex$(mlGenHandle) & ",  buffer size = " & mlGenPoints
            Exit Sub
         End If
      Else
         If mnLibType = MSGLIB Then
            RePlot False
         Else
            If Not mnuSettle.Checked Then DisplayData
         End If
         Exit Sub
      End If
   End If

   mnPlot = False
   If Index = mnPlotType Then
      mnuPlotType(mnPlotType).Checked = Not mnuPlotType(mnPlotType).Checked
   Else
      If mnuPlotType(DERIVATIVE).Checked Then SetAutoScale 0
      mnuPlotType(mnPlotType).Checked = False
      mnuPlotType(Index).Checked = True
   End If
   DoEvents
   
   If mnuPlotType(Index).Checked Then
      mnPlotType = Index
      SetPlotType mnPlotType + mnHardCopy, Me
      mnPlot = True
   End If
   If Index = SINGLE_POINT Then
      mlCount = Val(txtCount.Text)
      mnuContPlot.ENABLED = mnuPlotType(SINGLE_POINT).Checked
      BlockSize& = GetBlockSize()
      If BlockSize& <> mlCount Then
         mlBlockSave = BlockSize&
         SetBlockSize mlCount + mnCalConst, False
      End If
   End If
   If mnuPlotType(DERIVATIVE).Checked Then SetAutoScale 1
   If Not mnuSettle.Checked Then DisplayData
   mbJustReadBuffer = False
   A1 = mnPlotType
   If gnScriptSave And (Not gnInitializing) Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SPlotType, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   Me.mnuNoise.ENABLED = Me.mnuPlotType(HISTOGRAM).Checked

End Sub

Private Sub mnuPrint_Click()

   ExecutePrint

End Sub

Private Sub mnuQueue_Click()

   fraSetQueue.Top = 0
   fraSetQueue.Left = frmNewAnalog(mnThisInstance).ScaleHeight * 0.04
   fraSetQueue.Width = frmNewAnalog(mnThisInstance).Width * 0.95
   fraSetQueue.Height = frmNewAnalog(mnThisInstance).ScaleHeight - lblStatus.Height
   If mnQueueConfigurable Then
      Me.chkEnableSPoint.Visible = False
      Me.cmbDataRate.Visible = mnQueueDataRate
      Me.cmbType.Visible = mnQueueChanMode
   Else
      Me.chkEnableSPoint.Visible = (mnLibType = UNILIB)
      Me.cmbDataRate.Visible = False
      Me.cmbType.Visible = (mnLibType = UNILIB)
   End If
   fraSetQueue.Visible = True

End Sub

Private Sub mnuQuit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub mnuRes_Click(Index As Integer)

   mnuRes(Index).Checked = Not mnuRes(Index).Checked
   If mnuRes(Index).Checked Then
      For ResMenu% = 0 To 10
         If Not (ResMenu% = Index) Then mnuRes(ResMenu%).Checked = False
      Next ResMenu%
      Resolution% = Choose(Index + 1, 0, 4, 8, 10, 12, 13, 14, 16, 24, 32, 48)
      manForcedRes(mnBoardIndex) = Resolution%
   Else
      manForcedRes(mnBoardIndex) = 0
      UseList% = mnuUseList.Checked
      'BoardName$ = GetNameOfBoard(mnBoardNum)
      Select Case mnFormType
         Case ANALOG_IN
            If mnLibType = MSGLIB Then
               Resolution% = GetMsgADResolution(MsgLibrary)
            Else
               Resolution% = GetADResolution(msDisplayName, mnBoardNum)
            End If
         Case ANALOG_OUT
            If mnLibType = MSGLIB Then
            Else
               Resolution% = GetDAResolution(msDisplayName, mnBoardNum, mnRange)
            End If
      End Select
   End If
   If mnResolution <> Resolution% Then
      mnResolution = Resolution%
      SetControlRes
   End If
   UpdateMainStatus ""

End Sub

Private Sub mnuRetainPlot_Click()

   mnuRetainPlot.Checked = Not mnuRetainPlot.Checked
   If mnuBoardTitle.Checked Then TitleType% = 1
   PlotOptions% = mnuRetainPlot.Checked
   SetPlotOptions PlotOptions%, TitleType%
   If gnScriptSave Then
      A1 = PlotOptions%
      A2 = TitleType%
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetPlotOpts, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub mnuSaveStatus_Click()

   mnuSaveStatus.Checked = Not mnuSaveStatus.Checked
   If mnuSaveStatus.Checked Then
      Size = InputBox("Enter buffer size for storing CurCount history (default = 1000)", "Status History", "1000")
      If Size = "" Then
         mnuSaveStatus.Checked = False
         Exit Sub
      End If
      ArraySize& = CLng(Size)
      ReDim malStat(1, ArraySize&)
   Else
      Erase malStat
   End If

End Sub

Private Sub mnuScale_Click(Index As Integer)

   CurrentSelection% = Switch(mnScale = CELSIUS, 0, mnScale = FAHRENHEIT, 1, _
   mnScale = KELVIN, 2, mnScale = VOLTS, 3, mnScale = NOSCALE, 4, mnScale = RAW, 5)
   mnuScale(CurrentSelection%).Checked = False
   mnScale = Choose(Index + 1, CELSIUS, FAHRENHEIT, KELVIN, VOLTS, NOSCALE, RAW)
   NewSelection% = Switch(mnScale = CELSIUS, 0, mnScale = FAHRENHEIT, 1, _
   mnScale = KELVIN, 2, mnScale = VOLTS, 3, mnScale = NOSCALE, 4, mnScale = RAW, 5)
   mnuScale(NewSelection%).Checked = True
   mnDetailsNeedUpdate = True
   'If mnScale = RAW Then
      SetOptions
      mnDetailsNeedUpdate = True
      UpdateStatBar False
   'End If

End Sub

Private Sub mnuSetBlockSize_Click()

   Size = InputBox("Number of points to plot: ", "Set Block Size", Format$(mlBlockSize, "0"))
   If Size <> "" Then
      mlBlockSize = Val(Size)
      SetBlockSize mlBlockSize, True
      InitBlock True
      If Not mlHandle = 0 Then TransferBlock
      If gnScriptSave Then
         A1 = mlBlockSize
         FuncStat = 0
         For ArgNum% = 1 To 14
            ArgVar = Choose(ArgNum%, Me.Tag, SSetBlock, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
            If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
               PrintString$ = PrintString$ & ", "
            Else
               PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
            End If
         Next
         Print #2, PrintString$; Format$(AuxHandle, "0")
      End If
   End If

End Sub

Private Sub mnuSettle_Click()

   frmSetTimer.Caption = "Settling Time Configuraton"
   frmSetTimer.optTimerMode(0).Caption = "Settling to positive value"
   frmSetTimer.optTimerMode(1).Caption = "Settling to negative value"
   frmSetTimer.chkEnableTimer.Caption = "Enable settling test"
   frmSetTimer.Label1.Caption = "Frequency step"
   If mnuSettle.Checked Then
      frmSetTimer.chkEnableTimer.value = 1
      TestInProgress% = True
   End If
   frmSetTimer.txtInterval.Text = mlSettleStep
   If Not mnSettlePos Then frmSetTimer.optTimerMode(1).value = True
   frmSetTimer.Show 1
   mnuSettle.Checked = Not (frmSetTimer.chkEnableTimer.value = 0)
   mnSettlePos = frmSetTimer.optTimerMode(0).value
   mlSettleStep = Val(frmSetTimer.txtInterval.Text)
   Unload frmSetTimer
   If mnuSettle.Checked Then
      cmdStop.Visible = True
      If Not TestInProgress% Then
         If Not gnScriptRun Then OpenGPIB
         If Not gnScriptRun Then
            Dev$ = "DP8200N"
            mnUniSettle = True
            If Not GPIBSetDevice(0, Dev$) Then
               'if this dev doesn't exist, both DP8200s
               'should be at same address - no unipolar settling
               mnUniSettle = False
            End If
            Dev$ = "DP8200"
            If Not GPIBSetDevice(0, Dev$) Then
               mnuSettle.Checked = False
               MsgBox "Couldn't find " & Dev$
               Instance% = Val(Right$(Me.Tag, 2))
               CloseGPIB Instance%
               cmdStop.Visible = False
               Exit Sub
            Else
               If mnUniSettle Then
                  MsgBox "Connect DP8200s to adjacent channels.  The device named 'DP8200N' should be connected to the first channel in the scan.", , "Two DP8200s at separate addresses"
               Else
                  MsgBox "Connect DP8200s to adjacent channels with opposite polarity.", , "Two DP8200s at same address"
               End If
            End If
         End If
      End If
      If Not mnuPlotType(SINGLE_POINT).Checked Then mnuPlotType_Click (SINGLE_POINT)
   Else
      mfRate = Val(txtRate.Text)
      Instance% = Val(Right$(Me.Tag, 2))
      If Not gnScriptRun Then CloseGPIB Instance%
      If Not mnuPlotType(VOLTS_VS_TIME).Checked Then mnuPlotType_Click (VOLTS_VS_TIME)
      SetSettleStep 0, mfRate
      cmdStop.Visible = False
   End If

End Sub

Private Sub mnuStopBG_Click()

   mnuStopBG.Checked = Not mnuStopBG.Checked
   If mnuStopBG.Checked Then mnuStopSubSystemBackground.Checked = False

End Sub

Private Sub mnuStopSubSystemBackground_Click()

   mnuStopSubSystemBackground.Checked = Not mnuStopSubSystemBackground.Checked
   DoEvents
   If mnuStopSubSystemBackground.Checked Then
      mlSubSystem = mnFormType
      mnuStopBG.Checked = False
   End If

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

Private Sub mnuToEng_Click()

   mnuToEng.Checked = Not mnuToEng.Checked
   If (Abs(mnResolution) = 16) Or (mnuToEng.Checked) Then
      hsbData.Min = -32768
      hsbData.Max = 32767
   Else
      hsbData.Max = 4095
      hsbData.Min = 0
   End If
   If mnLibType = UNILIB Then
      x% = SaveFunc(Me, SToEng, 0, mnuToEng.Checked, _
      A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   End If
   Convert% = mnuToEng.Checked
   If Convert% Then
      If Me.mnuPlotType(HISTOGRAM).Checked Then
         mnuPlotType_Click (VOLTS_VS_TIME)
      End If
   End If
   mnuPlotType(HISTOGRAM).ENABLED = Not Convert%
   If mnLibType = UNILIB Then
      ConvertToEng mnBoardNum, Convert%
   Else
      ScaleString$ = "DISABLE"
      If Convert% Then ScaleString$ = "ENABLE"
      MessageUpdate$ = msStatusType & ":SCALE=" & ScaleString$
      MsgResult$ = MsgLibrary.SendMessage(MessageUpdate$)
      If SaveMsg(Me, "SendMessage(" & MessageUpdate$ & ")", MsgResult$) Then Exit Sub
      If (mnFuncType = AOutScan) Then
         mbDataGenRequired = True 'mnuBuildData_Click
      End If
   End If
   mnRefreshProps = True
   
End Sub

Private Sub mnuGetTC_Click()

   mnuGetTC.Checked = Not mnuGetTC.Checked
   NumChans% = mnLastChan - mnFirstChan + 1
   TotalCount& = Val(txtCount.Text)
   If mnuGetTC.Checked Then
      If Not (TotalCount& Mod NumChans%) = 0 Then
         txtCount.Text = Format((TotalCount& \ NumChans%) * NumChans%, "0")
      End If
   End If
   x% = SaveFunc(Me, SGetTC, 0, mnuGetTC.Checked, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

End Sub

Private Sub mnuTrigType_Click(Index As Integer)
   
   For TrigMenu% = 0 To mnuTrigType.Count - 1
      mnuTrigType(TrigMenu%).Checked = False
   Next TrigMenu%
   
   SetTrigVal Index
   
   UpdateStatBar False

End Sub

Private Sub mnuULBuf_Click()

   UseWinAPI% = Not mnuULBuf.Checked
   If Not mlGenHandle = 0 Then
      If BufFree(Me, mlGenHandle, UseWinAPI%) Then
         mlGenHandle = 0
         ClearHandle
      End If
   End If
   If Not mlHandle = 0 Then
      If BufFree(Me, mlHandle, UseWinAPI%) Then mlHandle = 0
   End If
   mnuULBuf.Checked = Not (mnuULBuf.Checked)

End Sub

Private Sub mnuUniRange_Click(Index As Integer)

   SetRange 1, Index

End Sub

Private Sub mnuUseList_Click()

   mnuUseList.Checked = Not mnuUseList.Checked
   UseList% = mnuUseList.Checked
   BoardName$ = msDisplayName
   Select Case mnFormType
      Case ANALOG_IN
         Resolution% = GetADResolution(BoardName$, mnBoardNum)
      Case ANALOG_OUT
         Resolution% = GetDAResolution(BoardName$, mnBoardNum, mnRange)
   End Select

End Sub

Private Sub OutScan()
   
   CBRate& = mfRate
   CBCount& = Val(Me.txtCount.Text)
   Do
      If mnLoop Then
         DelayVal& = tmrGoLoop.Interval * gnTimeFactor
         For LoopDelay& = 0 To DelayVal&
            DoEvents
            If Not cmdStop.ENABLED Then Exit Do
         Next LoopDelay&
         If mnTimerStop Then
            Select Case mnLibType
               Case UNILIB
                  ULStat = StopBackground520(mnBoardNum, AOFUNCTION)
                  Stat = SaveFunc(Me, StopBackground, ULStat, mnBoardNum, AOFUNCTION, _
                  A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
               Case MSGLIB
                  Status% = MsgCheckStatus(Me, msBoardName, _
                  MsgLibrary, msStatusType, CurCount&, CurIndex&)
                  If Status% = OVERRUN Then ULStat = OVERRUN
            End Select
            cmdGo.FontBold = Not cmdGo.FontBold
            DoEvents
         End If
      End If
      If mnIOTFunc Then
         ULStat = cbDaqOutScan(mnBoardNum, manChanArray(0), manTypeArray(0), _
         manGainArray(0), mlQCount, CBRate&, CBCount&, mlGenHandle, mvOptions)
         If SaveFunc(Me, DaqOutScan, ULStat, mnBoardNum, manChanArray(0), manTypeArray(0), _
         manGainArray(0), mlQCount, CBRate&, CBCount&, mlGenHandle, mvOptions, A10, A11, 0) Then
            cmdGo.FontBold = True
            cmdStop.ENABLED = False
            Exit Sub
         End If
      Else
         Select Case mnLibType
            Case UNILIB
               ULStat = cbAOutScan(mnBoardNum, mnFirstChan, mnLastChan, _
               CBCount&, CBRate&, mnRange, mlGenHandle, mvOptions)
               mlTotalCountReturn = mlCount
               mfRateReturned = CBRate&
               If SaveFunc(Me, AOutScan, ULStat, mnBoardNum, mnFirstChan, mnLastChan, _
               CBCount&, CBRate&, mnRange, mlGenHandle, mvOptions, A9, A10, A11, 0) Then
                  cmdGo.FontBold = True
                  cmdStop.ENABLED = False
                  Exit Sub
               End If
            Case MSGLIB
               'to do - remove when EXTPACER bug is fixed
               'If InStr(1, msOpt, "EXTCLOCK") > 0 Then
               '   MsgResult$ = MsgLibrary.SendMessage("AOSCAN:EXTPACER=ENABLE")
               'End If
               AOStartMsg$ = msStatusType & ":START"
               MsgResult$ = MsgLibrary.SendMessage(AOStartMsg$)
               If SaveMsg(Me, "SendMessage(" & AOStartMsg$ & ")", MsgResult$) Then
                  cmdStop = True
                  mnBufferEmpty = True
                  Exit Sub
               End If
               mnBufferEmpty = False
               'If SaveMsg(Me, "SendMessage(" & AOStartMsg$ & ")", MsgResult$) Then
               '   cmdGo.FontBold = True
               '   cmdStop.Enabled = False
               '   Exit Sub
               'End If
         End Select
      End If
      mfRateReturned = CBRate&
      Select Case mnLibType
         Case UNILIB
            ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
            If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, AOFUNCTION, _
            A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Case MSGLIB
            Status% = MsgCheckStatus(Me, msBoardName, _
            MsgLibrary, msStatusType, CurCount&, CurIndex&)
            If Status% = OVERRUN Then ULStat = OVERRUN
      End Select

      If Not mnThreading Then
         DevActive% = (Status% = RUNNING)
         If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
         If DevActive% Then
            If mnuGetStatus.Checked Or mnuGetSubSystemStatus.Checked Then tmrCheckStatus.ENABLED = True
            cmdStop.Visible = True
            cmdStop.ENABLED = True
            If gnScriptSave Then
               ScrDelayVal% = CInt((mlCount / (mfRate * ((mnLastChan - mnFirstChan) + 1)))) + 0.6
               Delay$ = Format$(ScrDelayVal%, "0")
               'SDelay = 3000
               Print #2, "0, 3000, 0, " & Delay$ & ",,,,,,,,,,,"
            End If
         End If
      End If
   Loop While mnLoop
   If mnLoopType = PROGLOOP Then cmdGo.FontBold = True

End Sub

Private Sub ReleaseSource()

   GPIBWrite 0, "L"
   mnFirstSourceSet = True

End Sub

Private Sub ResetData()

   If (gnNumBoards = 0) Then Exit Sub
   
   NumChans% = (mnLastChan - mnFirstChan) + 1
   If NumChans% < 1 Then Exit Sub
   If Not mnuCal.Checked Then
      If mlGenPoints > NumChans% Then
         Samples& = Val(txtCount.Text)
         If Samples& > mlGenPoints Then
            mbDataGenRequired = True
            'MsgBox "There are only " & Format(mlGenPoints, "0") & _
            '" data points in the buffer.", vbOKOnly, "Buffer Too Small"
            Exit Sub  'data handled elsewhere
         End If
      Else
         If mnLibType = MSGLIB Then
            If mlGenHandle = 0 Then mlGenHandle = ScaledBufAlloc(Me, NumChans%)
            If Not mlGenHandle = 0 Then
               ReDim DblDAData(NumChans% - 1, 0) As Double
               mnGenDataType = vbDouble
               If mnFormType = ANALOG_OUT Then
                  If mnuToEng.Checked Then
                     DataVal = Val(txtData.Text)
                     If gbULLoaded Then
                     ULStat = cbFromEngUnits(mnBoardNum, mnRange, DataVal, DataValue%)
                     If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, mnRange, DataVal, DataValue%, _
                     A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                     Else
                       'DataValue% = GetHiResVolts(mnResolution, mnRange, DataVal)
                     End If
                  Else
                     If mnFuncType = VOut Then
                        FloatData! = 1
                     Else
                        DblDataValue# = Val(txtData.Text)
                     End If
                  End If
               End If
               For CurChan% = 0 To NumChans% - 1
                  DblDAData(CurChan%, 0) = DblDataValue#
               Next
               ULStat = WDblArrayToBuf(Me, mlGenHandle, DblDAData(), NumChans%)
               If ULStat = 0 Then mlGenPoints = NumChans%
            End If
         Else
            UseWinAPI% = Not mnuULBuf.Checked
            If mlGenHandle = 0 Then mlGenHandle = BufAlloc16(Me, NumChans%, UseWinAPI%)
            If Not mlGenHandle = 0 Then
               ReDim DAData(NumChans% - 1, 0) As Integer
               mnGenDataType = vbInteger
               If mnFormType = ANALOG_OUT Then
                  If mnuToEng.Checked Then
                     DataVal = Val(txtData.Text)
                     If gbULLoaded Then
                     ULStat = cbFromEngUnits(mnBoardNum, mnRange, DataVal, DataValue%)
                     If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, mnRange, DataVal, DataValue%, _
                     A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                     Else
                       'DataValue% = GetHiResVolts(mnResolution, mnRange, DataVal)
                     End If
                  Else
                     If mnFuncType = VOut Then
                        FloatData! = 1
                     Else
                        DataValue% = ULongValToInt(Val(txtData.Text))
                     End If
                  End If
               End If
               For CurChan% = 0 To NumChans% - 1
                  DAData(CurChan%, 0) = DataValue%
               Next
               ULStat = WArrayToBuf(Me, mlGenHandle, DAData(), NumChans%)
               'ULStat = cbWinArrayToBuf(DAData(0, 0), mlGenHandle, 0, 1)
               If ULStat = 0 Then mlGenPoints = NumChans%
            End If
         End If
      End If
      mnDetailsNeedUpdate = True
      Exit Sub
   Else
      'to do - verify no issue in following line
      'added for cal crash
      mlCount = Val(txtCount.Text)
      CalPoints& = NumChans% * mlCount
      If Not mlCalHandle = 0 Then
         ULStat = cbWinBufFree(mlCalHandle)
         mlCalHandle = 0
      End If
      If mnLibType = MSGLIB Then
         mlCalHandle = cbWinBufAlloc64(CalPoints&)
      Else
         mlCalHandle = cbWinBufAlloc32(CalPoints&)
      End If
      'Exit Sub
   End If
   
   If mnDataChange Then
      mlCount = Val(txtCount.Text)
      mlTotalCountReturn = mlCount
      If mnuCal.Checked And (mnFuncType = AInScan) Then
         CIFactor% = (NumChans% \ 32) + 1
         mlCalInterval = 1000 * CIFactor%
         tmrGoLoop.Interval = mlCalInterval
      End If
      If mnuPlotType(SINGLE_POINT).Checked Then
         SetPlotType mnPlotType + mnHardCopy, Me
         SetBlockSize mlCount + mnCalConst, False
         'to do for output
      End If
      If mnFormType = ANALOG_OUT Then
         If mnLibType = MSGLIB Then
            DblDataValue# = Val(txtData.Text)
            If mnuCal.Checked Then
               DataType& = 5
               If mlCount < 5 Then
                  MsgBox "Use at least 5 data points for calibration.", , _
                  "Insufficient Number of Points"
                  Exit Sub
               End If
            End If
            Amplitude = Val(txtAmpl.Text)
            Offset = Val(txtOffset.Text)
            If mnRange < 99 Then Offset = Val(txtOffset.Text) + 2 ^ mnResolution / 2
            mnDataChange = False
            Chans% = mnLastChan - mnFirstChan
            If Not mnuULBuf.Checked Then
               UseLibrary% = 0
            Else
               UseLibrary% = 2
               UseLibrary% = 1
            End If
            If mnuCal.Checked Then
               DataType& = 5
               If mlCount < 5 Then
                  MsgBox "Use at least 5 data points for calibration.", , _
                  "Insufficient Number of Points"
                  Exit Sub
               End If
            End If
            If mlCount = 1 Then
               mlGenHandle = ScaledBufAlloc(Me, 1)
               If Not mlGenHandle = 0 Then
                  ReDim DblDAData(0, 0) As Double
                  mnGenDataType = vbDouble
                  DblDAData(0, 0) = DblDataValue#
                  ULStat = WDblArrayToBuf(Me, mlGenHandle, DblDAData(), 1)
                  If ULStat = 0 Then mlGenPoints = 1
               End If
            Else
               Select Case mnGenDataType
                  Case vbInteger
                     DataType& = 1
                  Case vbLong
                     DataType& = 2
                  Case vbDouble
                     DataType& = 6
               End Select
               SignalType% = 6
               mlGenHandle = GenerateData(DataType&, 1, mlCount, 1, Amplitude, Offset, SignalType%)
               If Not mlGenHandle = 0 Then mlGenPoints = mlCount
            End If
         Else
            If mnuToEng.Checked Then
               DataVal = Val(txtData.Text)
               ULStat = cbFromEngUnits(mnBoardNum, mnRange, DataVal, DataValue%)
               If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, mnRange, DataVal, _
               DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            Else
               If mnFuncType = VOut Then
                  FloatData! = 1
               Else
                  DataValue% = ULongValToInt(Val(txtData.Text))
               End If
            End If
            Amplitude = Val(txtAmpl.Text)
            Offset = Val(txtOffset.Text)
            If mnRange < 99 Then Offset = Val(txtOffset.Text) + 2 ^ mnResolution / 2
            mnDataChange = False
            Chans% = mnLastChan - mnFirstChan
            If Not mnuULBuf.Checked Then
               UseLibrary% = 0
            Else
               UseLibrary% = 2
               UseLibrary% = 1
            End If
            If mnuCal.Checked Then
               DataType& = 5
               If mlCount < 5 Then
                  MsgBox "Use at least 5 data points for calibration.", , _
                  "Insufficient Number of Points"
                  Exit Sub
               End If
            End If
            If mlCount = 1 Then
               mlGenHandle = cbWinBufAlloc(1)
               If Not mlGenHandle = 0 Then
                  ReDim DAData(0, 0) As Integer
                  mnGenDataType = vbInteger
                  DAData(0, 0) = DataValue%
                  ULStat = cbWinArrayToBuf(DAData(0, 0), mlGenHandle, 0, 1)
                  If ULStat = 0 Then mlGenPoints = 1
               End If
            Else
               Select Case mnGenDataType
                  Case vbInteger
                     DataType& = 1
                  Case vbLong
                     DataType& = 2
               End Select
               SignalType% = 6
               mlGenHandle = GenerateData(DataType&, 1, mlCount, 1, Amplitude, Offset, SignalType%)
               If Not mlGenHandle = 0 Then mlGenPoints = mlCount
            End If
         End If
      End If
   End If

End Sub

Private Sub ResetOpts()

   'reset options to default
   For i% = 0 To mnuOpt.Count - 1
      mnuOpt(i%).Checked = False
   Next i%
   mnDetailsNeedUpdate = True
   SetOptions
   UpdateStatBar False

End Sub

Private Sub SendData()
   
   If mnuToEng.Checked Then
      DataVal = Val(txtData.Text)
      ULStat = cbFromEngUnits(mnBoardNum, mnRange, DataVal, DataValue%)
      If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, mnRange, DataVal, DataValue%, _
      A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   Else
      DataValue% = ULongValToInt(Val(txtData.Text))
   End If

   For Chan% = mnFirstChan To mnLastChan
      ULStat = cbAOut(mnBoardNum, Chan%, mnRange, DataValue%)
      If SaveFunc(Me, AOut, ULStat, mnBoardNum, Chan%, mnRange, DataValue%, _
      A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      If mnuCal.Checked Then
         hsbData.ENABLED = False
         ReDim RealArray!(0, 0)
         tmrDelay.ENABLED = True
         Do
            DoEvents
         Loop While tmrDelay.ENABLED
         DataCum! = 0
         For Iteration! = 1 To mlAvgVal
            LinkCommand 0, "Read"
            LinkStart 0
            DoEvents
            Voltage! = LinkRetrieve(0)
            ValRead! = GetFloatCounts(mnResolution, mnRange, Voltage!)
            DataCum! = DataCum! + ValRead!
         Next Iteration!
         AveragedData! = DataCum! / CSng(mlAvgVal)
         mlValSet = IntValToULong(DataValue%)
         PointValue = mlValSet - AveragedData!
         RealArray!(Chan%, 0) = PointValue
         x% = PlotRealArray(RealArray!(), mlCount)
         'x% = PlotPoint(PointValue, Chan%)
      End If
   Next Chan%
   hsbData.ENABLED = True

End Sub

Private Sub SetControlRes()

   If gnScriptRun Then Exit Sub
   Select Case Abs(mnResolution)
      Case 16
         hsbData.Min = -32768
         hsbData.Max = 32767
      Case 12
         hsbData.Min = 0
         hsbData.Max = 4095
      Case 11
         hsbData.Min = 0
         hsbData.Max = 2047
      Case 10
         hsbData.Min = 0
         hsbData.Max = 1023
      Case Else
         hsbData.Min = 0
         hsbData.Max = 255
   End Select

   If mnuToEng.Checked Or (mnFuncType = VOut) Then
      hsbData.Min = -32768
      hsbData.Max = 32767
   End If
   Select Case mnFormType
      Case ANALOG_IN
      Case ANALOG_OUT
         txtAmpl.Text = 2 ^ Abs(mnResolution) / 2
         txtOffset.Text = "0" '2 ^ Abs(mnResolution) / 2
         mnDataChange = True
         If hsbData.Max = 32767 Then
            DataValue% = hsbData.value Xor &H8000
            Me.txtData.Text = IntValToULong(DataValue%)
         Else
            Me.txtData.Text = hsbData.value
         End If
         ResetData
   End Select

End Sub

Private Sub SetOptions()
   
   'Dim OptionMsg() As String
   mvOptions = FOREGROUND  'init to 0
   msOpt = ""
   MsgTrig$ = "DISABLE"
   'MsgPacer$ = msClockOff
   
   If Not mnLibType = MSGLIB Then
      Select Case mnFuncType
         Case AIn, ATrig, VIn, AIn32, VIn32
            If mnFuncType = VIn Then
               If mnuScale(5).Checked Then msOpt = "RAWDATA"
               ConfigureControls
            End If
            Exit Sub
         Case TIn
            'Exit Sub
      End Select
   Else
      Component$ = "AITRIG"
      TrigSupport$ = MsgLibrary.GetSupportedMessages(Component$)
      Select Case mnFuncType
         Case AIn, ATrig, VIn, AIn32, VIn32
            NonPaced% = True
            If 0 Then
            'to do - change temp configuration for msg boards
            If mnTempSupport Then
               Scales% = mnuScale.Count - 1
               For ScaleMenu% = 0 To Scales%
                  If mnuScale(ScaleMenu%).Checked Then
                     msOpt = mnuScale(ScaleMenu%).Caption
                     Exit For
                  End If
               Next
            End If
            End If
      End Select
      ConfScanOptions% = (mnLibType = MSGLIB) And (Not NonPaced%)
      If Not ConfScanOptions% Then Exit Sub
   End If
   
   'count changes if options are added
   SimWord$ = "DWORDXFER"
   If (Not mnIOTFunc) And (mnFuncType = AOutScan) Then SimWord$ = "SIMULTANEOUS"
   For i% = 0 To 9
      CurMsgOption$ = ""
      Options = Choose(i% + 1, BACKGROUND, CONTINUOUS, EXTCLOCK, _
         CONVERTDATA, SCALEDATA, SINGLEIO, DMAIO, BLOCK_IO, WORDXFER, DWORDXFER)
      opt$ = Choose(i% + 1, "BACKGROUND", "CONTINUOUS", "EXTCLOCK", _
         "CONVERTDATA", "SCALEDATA", "SINGLEIO", "DMAIO", "BLOCKIO", _
         "WORDXFER", SimWord$)
      If mnuOpt(i%).Checked Then
         Select Case i%
            Case 2
               CurMsgOption$ = ""
               If (mnLibType = MSGLIB) Then ClockOn$ _
                  = GetPacerStrings(MsgLibrary, msStatusType, 1)
               If Not ClockOn$ = "" Then CurMsgOption$ _
                  = msStatusType & ":EXTPACER=" & ClockOn$
            Case 5, 6, 7
               If msStatusType = "AO" Then
                  CurMsgOption$ = ""
               Else
                  CurMsgOption$ = msStatusType & ":XFRMODE=" & opt$
               End If
               TransferModeSet% = MessageSet%
         End Select
         If Not IsNull(Options) Then mvOptions = mvOptions + Options
         If Not IsNull(opt$) Then msOpt = msOpt & opt$ & " "
      Else
         Select Case i%
            Case 2
               If (msStatusType = "AO") Or (msStatusType = "AOSCAN") Then
                  CurMsgOption$ = ""
               Else
                  CurMsgOption$ = ""
                  If (mnLibType = MSGLIB) Then
                     PropLoc& = InStr(1, masDefaultScanProps(mnBoardIndex), "EXTPACER")
                     If PropLoc& > 0 Then
                        EndProp& = InStr(PropLoc&, masDefaultScanProps(mnBoardIndex), ";")
                        If EndProp& > PropLoc& Then
                           ClockOff$ = Mid(masDefaultScanProps(mnBoardIndex), _
                              PropLoc&, EndProp& - PropLoc&)
                        Else
                           ClockOff$ = Mid(masDefaultScanProps(mnBoardIndex), PropLoc&)
                        End If
                        CurMsgOption$ = msStatusType & ":" & ClockOff$
                     End If
                  End If
                  If Not ClockOff$ = "" Then CurMsgOption$ = msStatusType & ":" & ClockOff$
               End If
         End Select
      End If
      If ConfScanOptions% And (Not CurMsgOption$ = "") Then GoSub SetMsgOption
   Next i%
   
   For i% = 10 To mnuOpt.Count - 2
      CurMsgOption$ = ""
      opt$ = ""
      Options = Choose(i% - 9, NOFILTER, NOCLEAR, _
      BURSTMODE, WAITFORNEWDATA, EXTTRIGGER, NOCALIBRATEDATA, BURSTIO, _
      RETRIGMODE, NONSTREAMEDIO, ADCCLOCKTRIG, ADCCLOCK, HIGHRESRATE, SHUNTCAL)
      opt$ = Choose(i% - 9, "NOFILTER", "NOCLEAR", _
      "BURSTMODE", "WAITFORNEWDATA", "EXTTRIGGER", "NOCALIBRATEDATA", _
      "BURSTIO", "RETRIGMODE", "NONSTREAMEDIO", "ADCCLOCKTRIG", _
      "ADCCLOCK", "HIGHRESRATE", "SHUNTCAL")
      If mnuOpt(i%).Checked Then
         Select Case i%
            Case 12
               If mnMessaging Then
                  XfrModeMsg$ = "@" & msStatusType & ":BURSTMODE"
                  MsgResult$ = MsgLibrary.SendMessage(XfrModeMsg$)
                  If Not (InStr(1, MsgResult$, "NOT_SUPPORTED") > 0) Then
                     If msStatusType = "AO" Then
                        CurMsgOption$ = ""
                     Else
                        CurMsgOption$ = msStatusType & ":" & opt$ & "=ENABLE"
                     End If
                     TransferModeSet% = MessageSet%
                  End If
               End If
            Case 16
               If msStatusType = "AO" Then
                  CurMsgOption$ = ""
               Else
                  CurMsgOption$ = msStatusType & ":XFRMODE=" & opt$
               End If
               TransferModeSet% = MessageSet%
            Case 14
               If msStatusType = "AO" Then
                  CurMsgOption$ = ""
               Else
                  CurMsgOption$ = msStatusType & ":TRIG=ENABLE"
               End If
            Case 17
               If msStatusType = "AISCAN" Then
                  CurMsgOption$ = "AITRIG:REARM=ENABLE"
               End If
         End Select
         If Not IsNull(Options) Then mvOptions = mvOptions + Options
         If Not IsNull(opt$) Then msOpt = msOpt & opt$ & " "
      Else
         If Not (opt$ = "") Then
         If msStatusType = "AO" Then
            CurMsgOption$ = ""
         Else
            Select Case i%
            Case 12
               If mnMessaging Then
                  XfrModeMsg$ = "@" & msStatusType & ":BURSTMODE"
                  MsgResult$ = MsgLibrary.SendMessage(XfrModeMsg$)
                  If Not (InStr(1, MsgResult$, "NOT_SUPPORTED") > 0) Then
                     If msStatusType = "AO" Then
                        CurMsgOption$ = ""
                     Else
                        CurMsgOption$ = msStatusType & ":" & opt$ & "=DISABLE"
                     End If
                     TransferModeSet% = MessageSet%
                  End If
               End If
               Case 14
                  If (mnLibType = MSGLIB) Then
                     PropLoc& = InStr(1, masDefaultScanProps(mnBoardIndex), "TRIG")
                     If PropLoc& > 0 Then
                        EndProp& = InStr(PropLoc&, masDefaultScanProps(mnBoardIndex), ";")
                        If EndProp& > PropLoc& Then
                           DefTrig$ = Mid(masDefaultScanProps(mnBoardIndex), _
                              PropLoc&, EndProp& - PropLoc&)
                        Else
                           DefTrig$ = Mid(masDefaultScanProps(mnBoardIndex), PropLoc&)
                        End If
                        CurMsgOption$ = msStatusType & ":" & DefTrig$
                     End If
                  End If
               Case 16
                  If msStatusType = "AO" Then
                     CurMsgOption$ = ""
                  Else
                     If (mnLibType = MSGLIB) Then
                        If Not TransferModeSet% Then
                           PropLoc& = InStr(1, masDefaultScanProps(mnBoardIndex), "XFRMODE")
                           If PropLoc& > 0 Then
                              EndProp& = InStr(PropLoc&, masDefaultScanProps(mnBoardIndex), ";")
                              If EndProp& > PropLoc& Then
                                 DefXfer$ = Mid(masDefaultScanProps(mnBoardIndex), PropLoc&, EndProp& - PropLoc&)
                              Else
                                 DefXfer$ = Mid(masDefaultScanProps(mnBoardIndex), PropLoc&)
                              End If
                              CurMsgOption$ = msStatusType & ":" & DefXfer$
                           End If
                        End If
                     End If
                  End If
               Case 17
                  If (mnLibType = MSGLIB) Then
                     PropLoc& = InStr(1, masDefaultScanProps(mnBoardIndex), "AITRIG:REARM")
                     If PropLoc& > 0 Then
                        EndProp& = InStr(PropLoc&, masDefaultScanProps(mnBoardIndex), ";")
                        If EndProp& > PropLoc& Then
                           DefXfer$ = Mid(masDefaultScanProps(mnBoardIndex), _
                              PropLoc&, EndProp& - PropLoc&)
                        Else
                           DefXfer$ = Mid(masDefaultScanProps(mnBoardIndex), PropLoc&)
                        End If
                        CurMsgOption$ = msStatusType & ":" & DefXfer$
                     End If
                     RetrigSupport% = (InStr(1, TrigSupport$, "REARM") > 0)
                     If (msStatusType = "AISCAN") And RetrigSupport% Then
                        CurMsgOption$ = "AITRIG:REARM=DISABLE"
                     End If
                  End If
            End Select
         End If
         End If
      End If
      If ConfScanOptions% And (Not CurMsgOption$ = "") Then GoSub SetMsgOption
   Next i%
   If Not (mlCustomOption = 0) Then
      msOpt = msOpt & " Custom"
      mvOptions = (mvOptions Or mlCustomOption)
   End If
   If (mvOptions And SCALEDATA) = SCALEDATA Then
      If (mnFuncType = AOutScan) Then
         If Not mnUse64 Then
            mbDataGenRequired = True
            mnGenDataType = vbDouble
         End If
      End If
      mnUse64 = True
      '3 sets resolution, 0x10 sets use of cbScaleWinBufToArray
      '0x20 sets float type
      BufRes% = 3
      BufType% = BufRes% Or &H30
      SetBufferType BufType%
   ElseIf mnUse32 Then
      If (mnFuncType = AOutScan) Then
         If mnUse64 Then
            mbDataGenRequired = True
            mnGenDataType = vbInteger
         End If
      End If
      mnUse64 = False
      BufType% = 2
      SetBufferType BufType%
   Else
      If (mnFuncType = AOutScan) Then
         If mnUse64 Then
            mbDataGenRequired = True
            mnGenDataType = vbInteger
         End If
      End If
      mnUse64 = False
      BufType% = 1
      SetBufferType BufType%
   End If
   
   ConfigureControls
   If Len(msOpt) = 0 Then msOpt = "Options = Default  "
   Exit Sub
   
SetMsgOption:
   If mnMessaging Then
      OptionMsg$ = CurMsgOption$
      MsgResult$ = MsgLibrary.SendMessage(OptionMsg$)
      x% = SaveMsg(Me, "SendMessage(" & OptionMsg$ & ")", MsgResult$)
      MessageSet% = Not x%
   End If
   Return
   
GetMsgOption:
   If mnMessaging Then
      If Not ((msStatusType = "AO") Or (msStatusType = "AOSCAN")) Then
         XfrString$ = ":XFRMODE"
         XfrModeMsg$ = "?" & msStatusType & XfrString$
         MsgResult$ = MsgLibrary.SendMessage(XfrModeMsg$)
         x% = SaveMsg(Me, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$)
         NumLoc& = InStr(1, MsgResult$, "=")
         If NumLoc& > 0 Then
            XferModeString$ = Mid(MsgResult$, NumLoc& + 1)
            msOpt = msOpt & XferModeString$
         End If
      End If
   End If
   Return
   
End Sub

Private Sub SetRange(RangeType As Integer, RangeIndex As Integer)

   'If Not mnFormInitialized Then Exit Sub
   mnDetailsNeedUpdate = True
   Select Case RangeType
      Case -1   'NotUsed
         mnRange = NOTUSED
         GoSub ClearMenus
         mnuNoRange.Checked = True
         RangeDivisor! = 2
         Prefix$ = ""
         Suffix$ = ""
      Case 0   'Bipolar
         varRange = Choose(RangeIndex + 1, BIP30VOLTS, BIP20VOLTS, BIP10VOLTS, BIP5VOLTS, _
         BIP4VOLTS, BIP2PT5VOLTS, BIP2VOLTS, BIP1PT25VOLTS, BIP1VOLTS, _
         BIPPT625VOLTS, BIPPT5VOLTS, BIPPT25VOLTS)
         If IsNull(varRange) Then
            NewRange = RangeIndex - 11
            varRange = Choose(NewRange, BIPPT2VOLTS, BIPPT1VOLTS, BIPPT05VOLTS, _
            BIPPT01VOLTS, BIPPT005VOLTS, BIP1PT67VOLTS, BIPPT312VOLTS, BIPPT156VOLTS, _
            BIPPT078VOLTS, BIP60VOLTS, BIP15VOLTS, BIPPT125VOLTS, _
            BIPPT025VOLTSPERVOLT, BIPPT073125VOLTS)
         End If
         If ((IsNull(varRange)) Or (varRange = "")) Then
            MsgBox "Range not supported before revision " & _
            Format(CURRENTREVNUM, "General Number") & ".", , "UL Update Required"
            Exit Sub
         End If
         mnRange = varRange
         Select Case msDisplayName
            Case "PMD-1208LS", "miniLAB 1008", "PMD-1208FS"
               mnUseTwoSources = True
            Case "USB-1208LS", "USB-1208FS", "USB-1408FS"
               mnUseTwoSources = True
         End Select
         If mnUseTwoSources Then
            If (mnRange = BIP20VOLTS) And (mnuCal.Checked) Then
               If Not mnCalWarn Then
                  MsgBox "Use two DP8200s at the same GPIB address connected with " & _
                  "opposite polarities, one to CH IN LO, the other to CH IN HI", , "Two Sources Required"
                  mnCalWarn = True
               End If
            ElseIf (Not mnRange = BIP20VOLTS) And (mnuCal.Checked) Then
               If mnCalWarn Then
                  MsgBox "Remove DP8200 from CH IN LO and connect CH IN LO to GND", , "Remove One Source"
                  mnCalWarn = False
               End If
            End If
         End If
         GoSub ClearMenus
         mnuBipRange(RangeIndex).Checked = True
         RangeDivisor! = 2
         Prefix$ = "±"
         Suffix$ = "V"
      Case 1   'Unipolar
         varRange = Choose(RangeIndex + 1, UNI10VOLTS, UNI5VOLTS, UNI4VOLTS, UNI2PT5VOLTS, _
         UNI2VOLTS, UNI1PT25VOLTS, UNI1VOLTS, UNIPT5VOLTS, UNIPT25VOLTS, UNIPT2VOLTS)
         If IsNull(varRange) Then
            NewRange = RangeIndex - 9
            varRange = Choose(NewRange, UNIPT1VOLTS, UNIPT05VOLTS, UNIPT01VOLTS, UNIPT02VOLTS, UNI1PT67VOLTS)
         End If
         If ((IsNull(varRange)) Or (varRange = "")) Then
            MsgBox "Range not supported.", , "UL Update Required"
            Exit Sub
         End If
         mnRange = varRange
         GoSub ClearMenus
         mnuUniRange(RangeIndex).Checked = True
         RangeDivisor! = 1
         Prefix$ = "0 to "
         Suffix$ = "V"
      Case 2   'Current
         mnRange = Choose(RangeIndex + 1, MA4TO20, MA2to10, MA1TO5, _
         MAPT5TO2PT5, MA0TO20, BIPPT025AMPS)
         GoSub ClearMenus
         mnuCurRange(RangeIndex).Checked = True
         RangeVolts! = GetRangeVolts(mnRange)
         RangeLow! = 0
         RangeHigh! = RangeVolts!
         RangeDivisor! = 1
         If Not (mnRange > MAPT5TO2PT5) Then
            RangeLow! = RangeVolts! / 4
            RangeHigh! = RangeVolts! + RangeLow!
            Prefix$ = RangeLow! & " to "
            Suffix$ = "mA"
         End If
         If mnRange = BIPPT025AMPS Then
            Prefix$ = "±"
            Suffix$ = "A"
            RangeDivisor! = 2
         End If
      Case 3
         'custom range value
         varRange = InputBox("Enter custom range code.", "Custom Range", "400")
         If IsNumeric(varRange) Then
            mnRange = Val(varRange)
            GoSub ClearMenus
            Me.mnuCustRange.Checked = True
         End If
   End Select
   msRange = GetRangeString(mnRange)
   
   Select Case RangeType
      Case 2
         'mnuRange.Caption = "&Range (" & RangeLow! & " to " & RangeHigh! & "mA)"
         RangeVolts! = RangeVolts! / RangeDivisor!
         mnuRange.Caption = "&Range (" & Prefix$ & RangeVolts! & Suffix$ & ")"
      Case 3
         mnuRange.Caption = "&Range (Custom)"
      Case Else
         If Not mnRange = NOTUSED Then
            RangeVolts! = GetRangeVolts(mnRange)
            If Not IsNull(mvCustomRange) Then RangeVolts! = Val(mvCustomRange)
            RangeVolts! = RangeVolts! / RangeDivisor!
            mnuRange.Caption = "&Range (" & Prefix$ & RangeVolts! & Suffix$ & ")"
         Else
            mnuRange.Caption = "&Range (NOTUSED)"
         End If
   End Select
   
   If (mnFuncType = AIn) Or (mnFuncType = AInScan) Then CheckQueue
   mnDetailsNeedUpdate = True
   
   If mnFormType = ANALOG_OUT Then
      mnDataChange = True
      ResetData
   End If
   msMsgRange = "=" & GetMsgRange(mnRange)
   If mnLibType = MSGLIB Then
      If Not (mnMessaging And mnFormInitialized) Then Exit Sub
      QueueCapable% = mnQueueSupported Or mnQueueConfigurable
      If QueueCapable% And Not (mnFuncType = AOut) Then
         'don't mess with range if queue is enabled
         QueueMsg$ = "?AISCAN:QUEUE"
         MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
         QError% = SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$)
         If Not QError% = 0 Then
            QEnabled% = False
         Else
            QResult = Split(MsgResult$, "=")
            If UBound(QResult) > 0 Then
               QEnabled% = (QResult(1) = "ENABLE")
            Else
               QEnabled% = False
            End If
         End If
      End If
      If QEnabled% Then Exit Sub
      LastElement% = UBound(manChanArray)
      For Chan% = mnFirstChan To mnLastChan
         ChanString$ = "{" & Format(Chan%, "0") & "}"
         RangeString$ = ":RANGE"
         If mnFormType = ANALOG_OUT Then
            RangeMsg$ = msStatusType & ChanString$ & RangeString$ & msMsgRange
            If Not (mlFixedOutRange < 0) Then DontSetRange% = True
         Else
            If mnFuncType = AInScan Then
               DontSetRange% = InStr(1, msScanSupport, "AISCAN:RANGE{*}=") = 0
               RangeMsg$ = msStatusType & RangeString$ & ChanString$ & msMsgRange
               If DontSetRange% Then
                  MsgResult$ = MsgLibrary.SendMessage(RangeMsg$)
                  If Not (Left(MsgResult$, 5) = "Error") Then
                     DontSetRange% = False
                     If Not mbGSMConflictReported Then
                        MsgBox "Range format 'AISCAN:RANGE{*}=' not listed in GSM, but appears to be supported.", vbInformation, "GetSupportedMessages Conflict"
                        mbGSMConflictReported = True
                     End If
                  End If
               End If
            Else
               DontSetRange% = InStr(1, msAiSupport, "AI{*}:RANGE=") = 0
               RangeMsg$ = msStatusType & ChanString$ & RangeString$ & msMsgRange
            End If
         End If
         If Not DontSetRange% Then
            MsgResult$ = MsgLibrary.SendMessage(RangeMsg$)
            If SaveMsg(Me, "SendMessage(" & RangeMsg$ & ")", MsgResult$) Then
               Exit Sub
            End If
         End If
         If Not (Chan% > LastElement%) Then
            manChanArray(Chan%) = Chan%
            manGainArray(Chan%) = mnRange
         End If
      Next
   End If
   Exit Sub

ClearMenus:
   mnuNoRange.Checked = False
   For i% = 0 To mnuBipRange.Count - 1
      mnuBipRange(i%).Checked = False
   Next i%
   For i% = 0 To mnuUniRange.Count - 1
      mnuUniRange(i%).Checked = False
   Next i%
   For i% = 0 To mnuCurRange.Count - 1
      mnuCurRange(i%).Checked = False
   Next i%
   Return

End Sub

Private Sub SetResolution()

   If manForcedRes(mnBoardIndex) = 0 Then
      'If UseList% Then
         'boardstring$ = GetNameOfBoard(mnBoardNum)
      'Else
        ' boardstring$ = Str$(mnBoardNum)
      'End If
      Select Case mnFormType
         Case ANALOG_IN
            If mnLibType = MSGLIB Then
               Resolution% = GetMsgADResolution(MsgLibrary)
            Else
               Resolution% = GetADResolution(msDisplayName, mnBoardNum)
            End If
         Case ANALOG_OUT
            If mnLibType = MSGLIB Then
               Resolution% = GetMsgDAResolution(MsgLibrary)
            Else
               Resolution% = GetDAResolution(msDisplayName, mnBoardNum, mnRange)
            End If
      End Select
   Else
      Resolution% = manForcedRes(mnBoardIndex)
   End If
   If (mnResolution <> Resolution%) Or (mnBoardNum < 0) Then
      If mnFormType = ANALOG_OUT Then Unload frmBuildData
      mnResolution = Resolution%
      'If (mnBoardNum < 0) And (mnRange = GAIN1_SE) Then mnResolution = 11
        'as of 8/03, both UL and OEM libs treat data as 12 bit (though res is actually 11 bit)
      SetControlRes
   End If
   UpdateMainStatus ""
   mnDetailsNeedUpdate = True
   ConfigureControls

End Sub

Private Sub SetSource(ResetCount As Integer)

   Static CountVal&, CountIndex&, StartingVolts$

   If ResetCount Then
      CountVal& = 0
      CountIndex& = 0
      Success% = GPIBSetDevice(0, "DP8200")
      If Success% Then ReleaseSource
      Exit Sub
   End If
   If Not mnuCal.Checked Then Exit Sub
   If mlCount < 2 Then Exit Sub
   If mnFirstSourceSet Then
      SaveGoTimer% = tmrGoLoop.ENABLED
      cmdGo.FontBold = True
      tmrGoLoop.ENABLED = False
      tmrDelay.ENABLED = True
      Do
         'wait to allow source to stabilize
         'Don't allow go timer to fire
         DoEvents
      Loop While tmrDelay.ENABLED
      tmrGoLoop.ENABLED = SaveGoTimer%
      mnFirstSourceSet = False
   End If
   
   'R& = 2 ^ mnResolution
   FSR& = 2 ^ mnResolution
   MidScale& = FSR& / 2
   R& = Val(Me.txtAmpl.Text)
   CalOffSet& = Val(Me.txtOffset.Text)
   StepSize& = R& / (mlCount + 1)

   CountVal& = (MidScale& - ((StepSize& * mlCount) \ 2) + (CountIndex& * StepSize&)) + CalOffSet&
   CountIndex& = CountIndex& + 1

   If CountIndex& > mlCount Then
      ResetCount% = True
      CountIndex& = 0
'      CountVal& = CountVal& + StepSize&
'      'tmrGoLoop.Enabled = False
'      'Exit Sub
'   ElseIf CountVal& = 0 Then
'      CountVal& = (StepSize& \ mlCount) * 2
'      If CountVal& = 0 Then CountVal& = 1
'   Else
'      CountVal& = CountVal& + StepSize&
   End If
   IntCount% = ULongValToInt(CountVal&)
   mlValSet = CountVal&
   
   CheckRange% = mnRange
   Divisor! = 1
   If mnUseTwoSources Then
      If mnRange = BIP20VOLTS Then
         If CheckRange% = BIP20VOLTS Then CheckRange% = BIP10VOLTS
         DoubleSource% = True
      End If
   End If
   If (mnFuncType = VIn) Or mnUse32 Then
      If mnFixedRange = CheckRange% Then
         SourceVal! = GetHiResVolts(mnResolution, CheckRange%, CountVal&, mvCustomRange)
         MaxVal! = GetHiResVolts(mnResolution, CheckRange%, FSR&, mvCustomRange)
      Else
         SourceVal! = GetHiResVolts(mnResolution, CheckRange%, CountVal&)
         MaxVal! = GetHiResVolts(mnResolution, CheckRange%, FSR&)
      End If
      mfValSet = SourceVal!
      'find out if high voltage required
   Else
      SourceVal! = GetVolts(mnResolution, CheckRange%, IntCount%)
      IntCount% = ULongValToInt(FSR& - 1)
      MaxVal! = GetVolts(mnResolution, CheckRange%, IntCount%)
   End If
   If CheckRange% > 100 Then MaxVal! = MaxVal! / 2
   If MaxVal! > 10.48 Then
      Prefix$ = "V2"
      Divisor! = 10
   Else
      Prefix$ = "V1"
      Divisor! = 1
   End If
   SourceVal! = SourceVal! * mnInvertCal
   SourceStrVal! = SourceVal! / Divisor!
   Pol$ = "+"
   If SourceStrVal! < 0 Then Pol$ = ""
   SourceString$ = Format(SourceStrVal!, "00.00000")
   GPIBWrite 0, Prefix$ & Pol$ & SourceString$
   If CountIndex& = 1 Then
      StartingVolts$ = Format(SourceVal!, "00.00000")
      If DoubleSource% Then
         DoubleVal! = SourceVal! * 2
         StartingVolts$ = Format(DoubleVal!, "00.00000")
      End If
   End If
   If CountIndex& = 2 Then
      EndVal& = (MidScale& - ((StepSize& * mlCount) \ 2) + (mlCount * StepSize&)) + CalOffSet&
      IntCount% = ULongValToInt(EndVal&)
      IntRange% = ULongValToInt(FSR&)
      RangeVolts! = GetVolts(mnResolution, CheckRange%, 0)
      If RangeVolts! < 0 Then
         'if its less than 0 then it's bipolar so double it
         RangeVolts! = RangeVolts! * -2
      Else
         'if it's 0 then it's unipolar so get the other end
         If (mnFuncType = VIn) Or mnUse32 Then
            If mnFixedRange = CheckRange% Then
               RangeVolts! = GetHiResVolts(mnResolution, CheckRange%, EndVal&, mvCustomRange)
            Else
               RangeVolts! = GetHiResVolts(mnResolution, CheckRange%, EndVal&)
            End If
         Else
            RangeVolts! = GetVolts(mnResolution, CheckRange%, IntRange%)
         End If
      End If
      'now find LSB value in volts
      StepVolts! = StepSize& * (RangeVolts! / (2 ^ mnResolution))
      If (mnFuncType = VIn) Or mnUse32 Then
         If mnFixedRange = CheckRange% Then
            FinalVal! = GetHiResVolts(mnResolution, CheckRange%, EndVal&, mvCustomRange)
         Else
            FinalVal! = GetHiResVolts(mnResolution, CheckRange%, EndVal&)
         End If
      Else
         FinalVal! = GetVolts(mnResolution, CheckRange%, IntCount%)
      End If
      'FinalVal! = FinalVal! '/ Divisor!
      frmPlot.picDetails.Line (32, 0.8)-(64, 3.6), &HFFFFFF, BF
      DoEvents
      EndingVolts$ = Format(FinalVal!, "00.00000")
      IncrementVolts$ = Format(StepVolts!, "0.0000000V")
      If DoubleSource% Then
         DoubleVal! = FinalVal! * 2
         EndingVolts$ = Format(DoubleVal!, "00.00000")
         DoubleVal! = StepVolts! * 2
         IncrementVolts$ = Format(DoubleVal!, "0.0000000V")
      End If
      frmPlot.picDetails.CurrentX = 32
      frmPlot.picDetails.CurrentY = 0.8
      frmPlot.picDetails.Print "Linearity from " & StartingVolts$ & " to " & EndingVolts$
      frmPlot.picDetails.CurrentX = 32
      frmPlot.picDetails.CurrentY = 1.8
      frmPlot.picDetails.Print "Stepping by " & IncrementVolts$
   End If
   If ResetCount% Then CountVal& = 0

   tmrDelay.ENABLED = True
   Do
      'wait to allow source to stabilize
      DoEvents
   Loop While tmrDelay.ENABLED

End Sub

Private Sub Settling()

   PlotThis% = mnPlot
   mnPlot = False
   NumChans% = mnLastChan - mnFirstChan + 1
   cmdStop.ENABLED = True
   'Set the two DP8200s
   TestVolts! = GetRangeVolts(mnRange)
   If mnRange < 100 Then TestVolts! = TestVolts! / 2
   
   SourceVal! = TestVolts! - (TestVolts! * 0.01)
   SourceVal1! = TestVolts! * 0.01
   If Not mnUniSettle Then
      'only one external voltage available
      'no unipolar testing possible
      If Not mnSettlePos Then SourceVal! = SourceVal! * -1
      SourceVal1! = SourceVal!
      If Not (mnRange < 100) Then SourceVal1! = TestVolts! * 0.01
   Else
      If mnRange < 100 Then SourceVal1! = SourceVal! * -1
   End If
   
   Prefix$ = "V1+"
   Prefix1$ = "V1+"
   If mnSettlePos Then
      SourceString$ = Format(SourceVal!, "00.00000")
      SourceString1$ = Format(SourceVal1!, "00.00000")
      If SourceVal! < 0 Then Prefix$ = "V1"
      If SourceVal1! < 0 Then Prefix1$ = "V1"
   Else
      SourceString1$ = Format(SourceVal!, "00.00000")
      SourceString$ = Format(SourceVal1!, "00.00000")
      If SourceVal! < 0 Then Prefix1$ = "V1"
      If SourceVal1! < 0 Then Prefix$ = "V1"
   End If
   Dev$ = "DP8200"
   x = GPIBSetDevice(0, Dev$)
   GPIBWrite 0, Prefix$ & SourceString$
   
   If mnRange < 100 Then
      If mnUniSettle Then
         'write same value as above to second DP8200
         Dev$ = "DP8200N"
         x = GPIBSetDevice(0, Dev$)
         GPIBWrite 0, Prefix1$ & SourceString1$
      End If
   Else
      If Not mnUniSettle Then
         'can't run unipolar test with same voltage on both DP8200s
         MsgBox "No Unipolar Test", , "Unipolar test requires two DP8200s at " & _
         "separate addresses. Install another DP8200 named 'DP8200N'."
      Else
         Dev$ = "DP8200N"
         x = GPIBSetDevice(0, Dev$)
         GPIBWrite 0, Prefix1$ & SourceString1$
      End If
   End If

   tmrDelay.ENABLED = True
   Do
      'wait to allow source to stabilize
      DoEvents
   Loop While tmrDelay.ENABLED

   'set sampling rate at 1kHz to establish settled value
   mfRate = 1000
   AvgVal% = Val(txtCount.Text) / NumChans%
   'If Not mlAvgVal = AvgVal% Then
      mlAvgVal = Val(txtCount.Text) / NumChans%
      ResetData
   'End If
   'establish settled value (msSettleVal)
   AnalogIn
   If gnCancel Or mnCancel Then
        gnCancel = False
        mnCancel = False
        cmdStop_Click
        Exit Sub
   End If
   ReDim AvgData%(NumChans% - 1, mlAvgVal)
   If Not IsEmpty(mlHandle) Then
      ULStat = cbWinBufToArray(mlHandle, AvgData%(0, 0), 0, mlCount - 1)
   End If
   
   For SampleNum& = 0 To mlAvgVal - 1
      DataCum& = DataCum& + IntValToULong(AvgData%(mnLastChan, SampleNum&))
   Next SampleNum&
   
   msSettleVal = DataCum& / mlAvgVal
   mnPlot = PlotThis%
   
   'check results from txtRate down to 10k
   'at equal intervals
   mfRate = Val(txtRate.Text)
   mnSettling = True
   SetSettleStep mlSettleStep, mfRate

   NumPoints& = mfRate / mlSettleStep
   SetBlockSize NumPoints&, False
   RealRate& = mfRate * 2
   StartTime! = (1 / RealRate&) * 1000000
   EndTime! = (1 / (mfRate - mlSettleStep * (NumPoints& - 1))) * 500000
   TimeSpan$ = "t = " & Format$(StartTime!, "0.000") & " to " & Format$(EndTime!, "0.000") & "µs"
   BoardNameNum = mnuBoard(mnBoardIndex).Caption
   SettlingTo$ = "Settling to +FS"
   If Not mnSettlePos Then SettlingTo$ = "Settling to -FS"
   SetDetails mnFuncType, mnFirstChan, mnLastChan - mnFirstChan, _
   mlTotalCountReturn, TimeSpan$, mnRange, BoardNameNum, SettlingTo$

   For i& = 0 To NumPoints& - 1
      AnalogIn
      mfRate = mfRate - mlSettleStep
      If gnCancel Or mnCancel Then
           cmdStop_Click
           gnCancel = False
           mnCancel = False
           Exit For
      End If
   Next i&
   mnSettling = False
   cmdStop.ENABLED = False
   SourceString$ = "L"
   If mnUniSettle Then
      'write same value as above to second DP8200
      Dev$ = "DP8200N"
      x = GPIBSetDevice(0, Dev$)
      GPIBWrite 0, SourceString$
   End If
   Dev$ = "DP8200"
   x = GPIBSetDevice(0, Dev$)
   GPIBWrite 0, SourceString$

End Sub

Private Sub tmrCheckStatus_Timer()

   Dim TrapTrig As Boolean
   If gnScriptRun And gnScriptPaused Then
      tmrCheckStatus.ENABLED = False
      Exit Sub
   End If
   Static Do2ndHalf%, DoFirstHalf%, Element&, StillZero%, MsgBlock&
   StopMsg$ = msStatusType & ":STOP"
   RunningBackground% = ((mvOptions And BACKGROUND) = BACKGROUND)
   RunningContinuous% = ((mvOptions And CONTINUOUS) = CONTINUOUS)
   If mnThreading Then
      Me.tmrCheckStatus.ENABLED = False
      Exit Sub
   End If
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
      Exit Sub
   End If
   
   If mnuGetStatus.Checked Then
      Select Case mnLibType
         Case UNILIB
            ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
         Case NETLIB
            ULStat = NetLibrary.GetStatus(Status%, CurCount&, CurIndex&, mlStatusType)
         Case MSGLIB
            Status% = MsgCheckStatus(Me, msBoardName, _
            MsgLibrary, msStatusType, CurCount&, CurIndex&)
            If Status% = OVERRUN Then ULStat = OVERRUN
      End Select
      
      A2 = Status%
      A3 = CurCount&
      A4 = CurIndex&
      EvaluateCount% = (frmEvalData.chkCountStop.value = 1) _
         Or (frmEvalData.chkCountDelta.value = 1)
      
      If mbHalfRoll Then
         If CurCount& > 0 Then
            mnCountRolls = mnCountRolls + 1
            mbHalfRoll = False
         End If
      Else
         mbHalfRoll = (CurCount < 0)
      End If
      If mbEvalEnabled And EvaluateCount% Then EvalResult% = EvalCount(CurCount&)
      If mbEventCount Then
         If mbStartTrap Then
            'TestCount& = Abs(CurCount&)
            'If mlTrapCount < -1 Then TestCount& = CurCount&
            Select Case mnCountLast
               Case 0
                  If (CurCount& < mlTrapCount) Then TrapTrig = True
               Case 1
                  If (CurCount& = mlTrapCount) Then TrapTrig = True
               Case 2
                  If (CurCount& > mlTrapCount) Then TrapTrig = True
            End Select
            If TrapTrig Then
               mnCountInstances = mnCountInstances + 1
               If mnCountInstances < mnCountIterations Then
                  mbStartTrap = False
               Else
                  TrapCount! = CurCount&
                  TrapOnCount TrapCount!, mnCountInstances
               End If
            End If
         Else
            Select Case mnCountFirst
               Case 0
                  If (CurCount& < mlEvalCount) Then mbStartTrap = True
               Case 1
                  If (CurCount& = mlEvalCount) Then mbStartTrap = True
               Case 2
                  If (CurCount& > mlEvalCount) Then mbStartTrap = True
            End Select
         End If
      End If
   End If
   'If CurCount& Mod 5 <> 0 Then Stop
   If mnuSaveStatus.Checked Then
      malStat(0, Element&) = CurCount&
      malStat(1, Element&) = CurIndex&
      If (malStat(1, Element&) + 1) Mod 5 <> 0 Then Stop
      Element& = Element& + 1
      If Element& > UBound(malStat, 2) Then Element& = 0
   End If
   If (Not gnScriptSave) Or (ULStat <> 0) Then
      Select Case mnLibType
         Case UNILIB, NETLIB
            If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, A2, A3, A4, mlStatusType, _
            A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
         Case MSGLIB
            If Status% = -2 Then Failure% = True
      End Select
      If Failure% And Not (ULStat = TOOFEW) Then
         If mnuStopBG.Checked Then
            Select Case mnLibType
               Case UNILIB
                  ULStat = StopBackground520(mnBoardNum, mlStatusType)
                  If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, _
                  A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               Case NETLIB
                  ULStat = NetLibrary.StopBackground(mlStatusType)
               Case MSGLIB
                  If mnMessaging Then
                     MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
                     If MsgResult$ = "AISCAN:STATUS=INTERRUPTED" Then
                        ULStat = INTERRUPTED
                        MsgQualifier$ = "Error: "
                        MsgResult$ = MsgQualifier$ & MsgResult$
                     End If
                     If SaveMsg(Me, "SendMessage(" & StopMsg$ & _
                        ")", MsgResult$) Then Exit Sub
                  End If
            End Select
         End If
         cmdStop.Visible = False
         Status% = 0
         Do2ndHalf% = False
         DoFirstHalf% = False
         tmrCheckStatus.ENABLED = False
         DoEvents
         mnPlot = True
         Exit Sub
      End If
   End If
   
   If RunningBackground% Then UpdateStatBar False
   DataA% = ((mlEventType And ON_DATA_AVAILABLE) = ON_DATA_AVAILABLE)
   If mnuContPlot.Checked And Not DataA% Then
      'If mnLibType = MSGLIB Then
      '   ContMsg% = True
      '   If CurCount& > mlCount Then
      '      If (CurCount& / mlCount) > MsgBlock& Then
      '         MsgBlock& = MsgBlock& + 1
      '         DoFirstHalf% = True
      '      End If
      '   End If
      'Else
         BufferSplit& = mlCount \ 2
         If (Not Do2ndHalf%) And (Not DoFirstHalf%) Then
            mlFirstPoint = 0
            If CurIndex& > BufferSplit& Then
               DoFirstHalf% = True
               mnPlot = mnuPlotType(0).Checked _
                  Or mnuPlotType(DERIVATIVE).Checked
               mbEval = mbEvalEnabled
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
      'End If
   End If

   If Status% = 0 Then
      Element& = 0
      If mnuStopBG.Checked Then
         Select Case mnLibType
            Case UNILIB
               ULStat = StopBackground520(mnBoardNum, mlStatusType)
            Case NETLIB
               ULStat = NetLibrary.StopBackground(mlStatusType)
            Case MSGLIB
               If mnMessaging Then
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
                  If SaveFunc(Me, StopBackground, ULStat, _
                     mnBoardNum, mlStatusType, A3, A4, _
                     A5, A6, A7, A8, A9, A10, A11, 0) Then
                        mbCatastrophicError = ((ULStat = BADBOARD) Or _
                           (ULStat = DEADADDEV) Or (ULStat = NO_USB_BOARD))
                        If mbCatastrophicError Then
                           tmrCheckStatus.ENABLED = False
                           Exit Sub
                        End If
                  End If
               Case MSGLIB
                  If SaveMsg(Me, "SendMessage(" & StopMsg$ & _
                     ")", MsgResult$) Then
                        mbCatastrophicError = ((ULStat = BADBOARD) Or _
                           (ULStat = DEADADDEV) Or (ULStat = NO_USB_BOARD))
                        If mbCatastrophicError Then
                           tmrCheckStatus.ENABLED = False
                           Exit Sub
                        End If
                  End If
            End Select
         End If
         LoopCondition% = mnDelayRestart Or mnLoop Or tmrGoLoop.ENABLED
         If Not LoopCondition% Then cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      tmrCheckStatus.ENABLED = False
      If RunningBackground% Then UpdateStatBar False
      DoEvents
      mnPlot = mnuPlotType(0).Checked _
         Or mnuPlotType(1).Checked _
         Or mnuPlotType(2).Checked _
         Or mnuPlotType(DERIVATIVE).Checked
      mbEval = mbEvalEnabled
      If (Not (mnFormType = ANALOG_OUT)) And _
         ((Not mbEval) Or (Not mnuContPlot.Checked)) Then
            If Not ((mlEventType And ON_DATA_AVAILABLE) _
            = ON_DATA_AVAILABLE) Then DisplayData
         MsgBlock& = 0
      End If
   Else
      If ((CurIndex& < 1) And (CurCount& > mlCount) _
         And Not (mnBoardNum < 0)) And mnuContPlot.Checked Then
            'if CurIndex remains 0 because continuous is being run with a size
            'below block size, check status again to see if still 0, if so, plot
            'doesn't apply to miniLAB
            If mnuGetStatus.Checked Then
               Select Case mnLibType
                  Case UNILIB
                     ULStat = GetStatus520(mnBoardNum, Status%, _
                        CurCount&, CurIndex&, mlStatusType)
                  Case NETLIB
                     ULStat = NetLibrary.GetStatus(Status%, _
                        CurCount&, CurIndex&, mlStatusType)
                  Case MSGLIB
                     MsgResult$ = MsgLibrary.SendMessage(StatMsg$)
               End Select
            End If
         If StillZero% Then
            mnPlot = mnuPlotType(0).Checked _
               Or mnuPlotType(1).Checked _
               Or mnuPlotType(2).Checked _
               Or mnuPlotType(DERIVATIVE).Checked
            mbEval = mbEvalEnabled
            If (Not (mnFormType = ANALOG_OUT)) And (Not mbEval) Then
               If ((Not mnuContPlot.Checked)) Then
                  DisplayData
               ElseIf mnuContPlot.Checked And ContMsg% Then
                  If DoFirstHalf% Then
                     DisplayData
                     DoFirstHalf% = False
                  End If
               End If
            End If
         End If
         If (CurIndex& < 1) And (CurCount& > mlCount) Then StillZero% = True
      Else
         StillZero% = False
      End If
   End If

End Sub

Private Sub tmrDelay_Timer()

   tmrDelay.ENABLED = False

End Sub

Private Sub tmrGoLoop_Timer()

   If mnSemaphore Then Exit Sub
   cmdGo.FontBold = Not cmdGo.FontBold
   'Static TimerCount&
   UpdateStatBar False
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
      If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
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
   cmdGo = True
   mlTimerCount = mlTimerCount + 1
   If mnTimerTillCount Then
      NumChans% = mnLastChan - mnFirstChan + 1
      If mlTimerCount >= (mlCount / NumChans%) - 1 Then  'replaced Val(txtCount.Text) with mlCount
         tmrGoLoop.ENABLED = False
         cmdGo.FontBold = True
         mlTimerCount = 0
      End If
   End If
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
   cmdStop.ENABLED = tmrGoLoop.ENABLED

End Sub

Private Sub txtAmpl_KeyPress(KeyAscii As Integer)

   mnDataChange = True
   mnSaveAmpl = True
   If Me.mnuCal.Checked Then
      Me.optFSN.Visible = True
      Me.optFSP.Visible = True
      Me.cmdOS.Visible = True
   End If
   If gnScriptSave And (mnFormType = ANALOG_OUT) And (KeyAscii = 13) Then
      A1 = Val(txtAmpl.Text)
      A3 = Val(txtAmpl.Text) 'backward compatibility
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetAmplitude, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
      mnSaveAmpl = False
   End If

End Sub

Private Sub txtAmpl_LostFocus()

   ResetData
   If gnScriptSave And (mnFormType = ANALOG_OUT) And mnSaveAmpl Then
      A1 = Val(txtAmpl.Text)
      A3 = Val(txtAmpl.Text)    'backward compatibility
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetAmplitude, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
      mnSaveAmpl = False
   End If

End Sub

Private Sub txtChan_Change()
   
   CheckQueue

End Sub


Private Sub txtCount_KeyPress(KeyAscii As Integer)

   mnDataChange = True
   mnDetailsNeedUpdate = True
   'If KeyAscii = 13 Then KeyCode = 0

End Sub

Private Sub txtCount_LostFocus()

   'For OpenForm% = 0 To Forms.Count - 1
   '   If Forms(OpenForm%).Name = "frmBuildData" Then
   '      SettingCount% = True
   '      frmBuildData.txtSamples.Text = txtCount.Text
   '   End If
   'Next
   If Not SettingCount% Then ResetData
   A1 = txtCount.Text
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SCountSet, FuncStat, _
         A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   If mnLibType = MSGLIB And Not mnuCal.Checked Then SetMsgCount

End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      mnDataChange = True
      ResetData
      AnalogOut
   End If

End Sub

Private Sub txtHighChan_Change()

   mnDataChange = True
   NumChans% = (mnLastChan - mnFirstChan) + 1
   If mlGenPoints = NumChans% Then MaintainDataCount% = True
   mnLastChan = Val(txtHighChan.Text)
   NumChans% = (mnLastChan - mnFirstChan) + 1
   If MaintainDataCount% Then mlGenPoints = NumChans%
   TotalCount& = Val(txtCount.Text)
   If mnuGetTC.Checked Then
      If Not (TotalCount& Mod NumChans%) = 0 Then
         txtCount.Text = Format((TotalCount& \ NumChans%) * NumChans%, "0")
      End If
   End If
   If (mnLibType = MSGLIB) And Not mlQEnabled Then
      If mnMessaging And Not (msScanSupport = "") Then
         If Not InStr(1, msStatusType, "SCAN") = 0 Then
            StatusType$ = msStatusType
         Else
            If mnFormType = ANALOG_OUT Then
               StatusType$ = "AOSCAN"
            Else
               StatusType$ = "AISCAN"
            End If
         End If
         HighChanMsg$ = StatusType$ & ":HIGHCHAN=" & Val(mnLastChan)
         MsgResult$ = MsgLibrary.SendMessage(HighChanMsg$)
         x% = SaveMsg(Me, "SendMessage(" & HighChanMsg$ & ")", MsgResult$)
      End If
   End If
   
   mnDetailsNeedUpdate = True
   
End Sub

Private Sub txtHighChan_LostFocus()

   ResetData

End Sub

Private Sub txtLowChan_Change()

   mnDataChange = True
   NumChans% = (mnLastChan - mnFirstChan) + 1
   If mlGenPoints = NumChans% Then MaintainDataCount% = True
   mnFirstChan = Val(txtLowChan.Text)
   NumChans% = mnLastChan - mnFirstChan + 1
   If MaintainDataCount% Then mlGenPoints = NumChans%
   TotalCount& = Val(txtCount.Text)
   If mnuGetTC.Checked Then
      If Not (TotalCount& Mod NumChans%) = 0 Then
         txtCount.Text = Format((TotalCount& \ NumChans%) * NumChans%, "0")
      End If
   End If
   If (mnLibType = MSGLIB) And Not mlQEnabled Then
      If mnMessaging Then
         If Not InStr(1, msStatusType, "SCAN") = 0 Then
            StatusType$ = msStatusType
         Else
            If mnFormType = ANALOG_OUT Then
               If msScanSupport = "" Then
                  DontSetChannel% = True
               Else
                  StatusType$ = "AOSCAN"
               End If
            Else
               StatusType$ = "AISCAN"
            End If
         End If
         If Not DontSetChannel% Then
            LowChanMsg$ = StatusType$ & ":LOWCHAN=" & Val(mnFirstChan)
            MsgResult$ = MsgLibrary.SendMessage(LowChanMsg$)
            x% = SaveMsg(Me, "SendMessage(" & LowChanMsg$ & ")", MsgResult$)
         End If
      End If
   End If
   mnDetailsNeedUpdate = True

End Sub

Private Sub txtLowChan_LostFocus()

   ResetData

End Sub

Private Sub txtOffset_KeyDown(KeyCode As Integer, Shift As Integer)

   mnSaveOffset = True
   If KeyCode = 13 Then
      If (mnFuncType = FileAInScan) Or (mnFuncType = FilePretrig) Then
         lpFileName$ = "UniTest.ini"
         lpApplicationName$ = "DataFile"
         lpKeyName$ = "FileName"
         DataFile$ = txtOffset.Text
         StringSize% = WritePrivateProfileString(lpApplicationName$, _
         lpKeyName$, DataFile$, lpFileName$)
      End If
      mnSaveOS = True
      If gnScriptSave And (mnFormType = ANALOG_OUT) Then
         A1 = Val(txtOffset.Text)
         A3 = Val(txtOffset.Text)       'backward compatibility
         FuncStat = 0
         For ArgNum% = 1 To 14
            ArgVar = Choose(ArgNum%, Me.Tag, SSetOffset, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
            If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
               PrintString$ = PrintString$ & ", "
            Else
               PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
            End If
         Next
         Print #2, PrintString$; Format$(AuxHandle, "0")
         mnSaveOffset = False
      End If
   End If

End Sub

Private Sub txtOffset_KeyPress(KeyAscii As Integer)

   mnDataChange = True
   mnSaveOffset = True

End Sub

Private Sub txtOffset_LostFocus()

   ResetData
   If gnScriptSave And (mnFormType = ANALOG_OUT) And mnSaveOffset Then
      A1 = Val(txtOffset.Text)
      A3 = Val(txtOffset.Text)  'backward compatibility
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetOffset, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
      mnSaveOffset = False
   End If

End Sub

Private Sub txtQCount_Change()

   SetQueueCount
   
End Sub

Private Sub txtRate_Change()

   '9 chars prevents overflow error
   mnDetailsNeedUpdate = True
   If mnLibType = MSGLIB Then
      SetMsgRate
   Else
      mfRate = Val(txtRate.Text)   'Val(Left(txtRate.Text, 9))
   End If
   
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Then
      If mnuLoop.Checked Then tmrGoLoop.Interval = Val(txtRate.Text)
      KeyAscii = 0
      Select Case mnLibType
         Case UNILIB
            mfRate = Val(txtRate.Text)
            mfRateReturned = mfRate
         Case MSGLIB
            SetMsgRate
      End Select
   End If

End Sub

Private Sub txtRate_LostFocus()

   '9 chars limit removed when mfRate changed from single to double
   mnDetailsNeedUpdate = True
   Select Case mnLibType
      Case UNILIB
         mfRate = Val(txtRate.Text)
         mfRateReturned = mfRate
      Case MSGLIB
         SetMsgRate
   End Select

End Sub

Private Sub UpdateMainStatus(FunctionStat As String)

   If Len(FunctionStat) Then FunctionStat = "  [" & FunctionStat & "]"
   board$ = mnuBoard(mnBoardIndex).Caption
   PrintMain "Current board: " & board$ & FunctionStat
   
End Sub

Private Sub UpdateStatBar(CycleStatus As Integer)
   
   Static sStatus As Integer
   If CycleStatus Then
      sStatus = sStatus + 1
      If sStatus > 3 Then sStatus = 0
      StatType% = sStatus
   ElseIf tmrCheckStatus.ENABLED Then
      StatType% = 2
   Else
      StatType% = 0 'sStatus
   End If
   
   Select Case StatType%
      Case 0
         Stat$ = msConfig & DDERange$ & " " & msOpt
      Case 1
         Stat$ = "Range = " & GetRangeString(mnRange)
         If Not TestRange(mnFormType, mnBoardNum, mnRange) Then Stat$ = Stat$ & " (invalid)"
         Stat$ = Stat$ & "   Rate = " & mfRateReturned
         If mnFuncType = APretrig Then
            Stat$ = Stat$ & "   Count = " & mlTotalCount
         Else
            Stat$ = Stat$ & "   Count = " & mlCount
         End If
      Case 2
         If (gnNumBoards > 0) Then
            If mnuGetStatus.Checked Then
               Select Case mnLibType
                  Case UNILIB
                     ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
                     FuncID% = GetStatus
                  Case NETLIB
                     ULStat = NetLibrary.GetStatus(Status%, CurCount&, CurIndex&, mlStatusType)
                     FuncID% = GetStatus
                  Case MSGLIB
                     Status% = MsgCheckStatus(Me, msBoardName, _
                     MsgLibrary, msStatusType, CurCount&, CurIndex&)
               End Select
               A2 = Status%
               A3 = CurCount&
               A4 = CurIndex&
            End If
            If (Not gnScriptSave) Or (ULStat <> 0) Then
               If mnuGetStatus.Checked Then
                  Select Case mnLibType
                     Case UNILIB, NETLIB
                        If SaveFunc(Me, FuncID%, ULStat, mnBoardNum, A2, A3, A4, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
                     Case MSGLIB
                        If Status% = -2 Then Exit Sub
                  End Select
               End If
            End If
            DevActive% = (Status% = RUNNING)
            If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
            If DevActive% Then
               BGOp$ = "Background task in progress..."
            Else
               BGOp$ = "No background task running..."
            End If
            InstString$ = ""
            If mnCountRolls > 1 Then RollString$ = _
               " (X" & mnCountRolls & ")"
            Stat$ = BGOp$ & " Count = " & CurCount& & _
               "  Index = " & CurIndex& & RollString$
         End If
      Case 3
         Stat$ = "Trigger type = " & GetTrigString(mnTrigType)
   End Select
   lblStatus.Caption = Stat$
   
End Sub

Private Sub vsbChan_Change()

   txtChan.Text = vsbChan.Max - vsbChan.value
   CheckQueue

End Sub

Private Sub cmbType_Click()

   CheckQueue

End Sub

Private Sub mnuResolution_Click(Index As Integer)

   For i% = 0 To mnuResolution.Count - 1
      mnuResolution(i%).Checked = False
   Next i%
   mnuResolution(Index).Checked = True
   ResVal% = Choose(Index + 1, -1, 0, 4, 8, 12, 13, 14, 16, 24, 32, 48, 64)
   SetPlotResolution ResVal%
   A1 = ResVal%
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SSetResolution, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub V_Out()

   NumChans% = (mnLastChan - mnFirstChan) + 1

   VoltageVal! = Val(txtData.Text)
   ULStat = cbVOut(mnBoardNum, Chan% + mnFirstChan, mnRange, VoltageVal!, mvOptions)
   If SaveFunc(Me, VOut, ULStat, mnBoardNum, Chan% + mnFirstChan, mnRange, VoltageVal!, mvOptions, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

   Exit Sub
   Do
      For Element% = 0 To PerChan% - 1
         For Chan% = 0 To NumChans% - 1
            ULStat = cbAOut(mnBoardNum, Chan% + mnFirstChan, mnRange, DataValue%)
            'following required for scripting info
            A5 = mnFirstChan
            A6 = mnLastChan
            A7 = mlCount
            If SaveFunc(Me, AOut, ULStat, mnBoardNum, Chan% + mnFirstChan, mnRange, DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            If mnuCal.Checked Then
               cmdStop.Visible = True
               cmdStop.ENABLED = True
               tmrDelay.ENABLED = True
               Do
                  DoEvents
                  If Not cmdStop.ENABLED Then
                     cmdStop.Visible = False
                     Exit Sub
                  End If
               Loop While tmrDelay.ENABLED
               DataCum! = 0
               Instance% = 0
               For Iteration! = 1 To mlAvgVal
                  If InStr(1, msDevName, "34401") Then
                     LinkCommand Instance%, "Read?"
                     LinkStart Instance%
                     DoEvents
                  End If
                  'next line required due to 32 bit GPIB bug
                  'need to read 8840 twice to get updated reading
                  If InStr(1, msDevName, "8840") Then LinkCommand Instance%, "Read"
                  LinkCommand Instance%, "Read"
                  'LinkStart
                  DoEvents
                  Voltage! = LinkRetrieve(Instance%)
                  NewRange% = mnRange
                  'If mnRange = MA0TO20 Then NewRange% = UNI10VOLTS
                  ValRead! = GetFloatCounts(mnResolution, NewRange%, Voltage!)
                  DataCum! = DataCum! + ValRead!
               Next Iteration!
               If mnDetailsNeedUpdate Then
                  BoardName = mnuBoard(mnBoardIndex).Caption
                  SetDetails mnFuncType, mnFirstChan, mnLastChan - mnFirstChan, mlAvgVal, mfRateReturned, mnRange, BoardName, mlPTCountReturn
                  mnDetailsNeedUpdate = False
               End If
               AveragedData! = DataCum! / CSng(mlAvgVal)
               If gnScriptRun Then
                  Voltage! = Val(cmdPlot.Caption)
                  mlValSet = GetCounts(mnResolution, mnRange, Voltage!)
               Else
                  mlValSet = IntValToULong(DataValue%)
               End If
               PointValue = mlValSet - AveragedData!
               x% = PlotPoint(PointValue, Chan%)
            End If
         Next Chan%
      Next Element%
      If mnuCal.Checked Then
         cmdStop.Visible = False
         cmdStop.ENABLED = False
      End If
      DoEvents
   Loop While mnLoop

End Sub

Private Sub SelectRangeUsingCode(RangeCode As Long)

   RangeName$ = GetRangeString(RangeCode)
   Select Case RangeCode
      Case -1
         mnuNoRange_Click
      Case 0 To 99
         For RangeMenuIndex& = 0 To mnuBipRange.Count - 1
            If mnuBipRange(RangeMenuIndex&).Caption = RangeName$ Then Exit For
         Next RangeMenuIndex&
         mnuBipRange_Click (RangeMenuIndex&)
      Case 100 To 199
         For RangeMenuIndex& = 0 To mnuUniRange.Count - 1
            If mnuUniRange(RangeMenuIndex&).Caption = RangeName$ Then Exit For
         Next RangeMenuIndex&
         mnuUniRange_Click (RangeMenuIndex&)
      Case 200 To 299
         For RangeMenuIndex& = 0 To mnuCurRange.Count - 1
            If mnuCurRange(RangeMenuIndex&).Caption = RangeName$ Then Exit For
         Next RangeMenuIndex&
         mnuCurRange_Click (RangeMenuIndex&)
      Case Else
         'covers Unipolar / Bipolar fixed
         If RangeCode = UNIPOLAR Then mnuUniRange_Click (0)
         If RangeCode = BIPOLAR Then mnuBipRange_Click (2)
   End Select

End Sub

Public Function GetQueueList(QChans As Variant, QGains As Variant, QTypes As Variant) As Long

   If mnLibType = MSGLIB Then
      If mlQCount > 0 Then
         ChanCount% = mlQCount
      Else
         If mnFormType = ANALOG_IN Then
            If mnNumAIChans > 0 Then ChanCount% = mnNumAIChans
         Else
            If mnNumAOChans > 0 Then ChanCount% = mnNumAOChans
         End If
      End If
   Else
      ChanCount% = mlQCount
   End If
   QChans = manChanArray()
   QGains = manGainArray()
   GetQueueList = ChanCount%
   
End Function

Public Sub SetupData(DataHandle&, NumPoints&, Channels%, DataType&)

   mlGenHandle = DataHandle&
   mlGenPoints = NumPoints&: mnGenChans = Channels%
   If mlGenPoints > 1 Then cmdGo.Visible = True
   If mnLibType = MSGLIB Then mlGenPoints = NumPoints& * Channels%

   If Not DataHandle& = 0 Then
      Select Case DataType&
         Case 1
            mnGenDataType = vbInteger
         Case 2
            mnGenDataType = vbLong
         Case 4
            mnGenDataType = vbSingle
         Case 6
            mnGenDataType = vbDouble
            If mnLibType = MSGLIB Then
               'Dim Ranges(0) As String
               ReDim DatArray(Channels% - 1, NumPoints& - 1) As Double
               ULStat = LoadDblArrayFromWinBuf64(Me, mlGenHandle, DatArray(), _
               FirstPoint&, mlGenPoints)
               If Not ULStat = 0 Then Exit Sub
               'Ranges(0) = "UNI4.096V"
               If Not (mnFuncType = AOut) Then
                  mlWriteTimeout = Val(frmBuildData.txtTimeOut.Text)
                  MsgResult$ = MsgLibrary.WriteOutScan(DatArray, NumPoints&, mlWriteTimeout)
                  x& = SaveMsg(Me, "WriteOutScan(" & DatArray(0, 0) & ", " & NumPoints& & ", " & _
                  mlWriteTimeout & ")", MsgResult$)
               End If
            End If
      End Select
   Else
      mnGenDataType = vbEmpty
   End If
   If ((mvOptions And SCALEDATA) = SCALEDATA) And (mlGenPoints > 0) Then
      If Not mnGenDataType = vbDouble Then
         MsgBox "Make sure you have the right data type specified for the SCALEDATA option. " & _
         "The scan will expect a double data type, but that type was not generated.", _
         vbCritical, "Possible Data Type Problem"
      End If
   End If

End Sub

Public Function GetDataHandle(AcqOrGen As Integer, DataType As Long, NumSamples As Long) As Long

   Select Case AcqOrGen
      Case ACQUIREDDATA
         GetDataHandle = mlHandle
         If mnLibType = MSGLIB Then
            If (Not (Me.mnuToEng.Checked)) And _
            (mnAcqDataType = vbDouble) Then Convert& = &H10
         End If
         DataType = mnAcqDataType Or Convert&
         NumSamples = mlAcqPoints
      Case GENERATEDDATA
         GetDataHandle = mlGenHandle
         DataType = mnGenDataType
         NumSamples = mlGenPoints
   End Select
   
End Function

Public Sub PlotGenData()

   If (mlGenHandle = 0) Or (mlGenPoints = 0) Then Exit Sub
   Select Case mnGenDataType
      Case vbInteger
         BufType% = 1
      Case vbLong
         BufType% = 2
      Case vbSingle
         BufType% = 4 Or &H30
      Case vbDouble
         BufType% = 3 Or &H30
   End Select
   SetPlotType mnPlotType, Me
   '0x40 = FROMMESSAGE, 0x20 = float, 3 = double
   '0x10 = scaled, 0x80 = BUFFER, 4 = single
   SetBufferType BufType%
   ShowBufferInfo 3
   PlotBuffer mlGenHandle, mlGenPoints, mnGenChans - 1
   mbDataGenRequired = False
   
End Sub

Public Sub PlotAcquiredData()

   If (mlHandle = 0) Or (mlCount = 0) Then Exit Sub
   Select Case mnAcqDataType
      Case vbInteger
         BufType% = 1
      Case vbLong
         BufType% = 2
      Case vbSingle
         BufType% = 4 Or &H20
      Case vbDouble
         ScalingData% = mnuToEng.Checked
         SpecifyMsgLib& = &H100
         '0x40 = FROMMESSAGE, 0x20 = float, 3 = double
         '0x10 = scaled, 0x80 = BUFFER
         If ScalingData% Then
            BufType% = 3 Or &H20 Or &H10
         Else
            BufType% = 3 Or &H20
         End If
   End Select
   SetBufferType BufType%
   NumChans% = (mnLastChan - mnFirstChan)
   If mnMessaging And Not mnBufferEmpty Then _
   Success% = TransferDataFromMsg(mlCount, False)
   PlotBuffer mlHandle, mlCount, NumChans%

End Sub

Public Sub Set32Bit(BitArg As Integer)

   Select Case BitArg
      Case 0
         'resets to native resolution
         If (mnResolution > 16) Then
            If Not Me.mnuBuff32.Checked Then mnuBuff32_Click
         Else
            If Me.mnuBuff32.Checked Then mnuBuff32_Click
         End If
      Case 1
         'forces 16 bit function calls
         If mnuBuff32.Checked Then mnuBuff32_Click
      Case 2
         'forces 32 bit function calls
         If Not mnuBuff32.Checked Then mnuBuff32_Click
   End Select
   
End Sub

Private Sub GetMeterValue()

   TimerState% = tmrGoLoop.ENABLED
   tmrGoLoop.ENABLED = False
   GPIBSetDevice 0, "DP8200"

   If mnuCal.Checked Or mnTimerTillCount Then TotalCount& = mlAvgVal
   If mnSampleNum = 0 Then
      mnFirstSourceSet = True
   End If
   If Not gnScriptRun Then SetSource False

   GPIBSetDevice 0, "HP34401"
   
   ReDim LongVals(0, mlAvgVal - 1) As Long
   ReDim LongBack(0, mlAvgVal - 1) As Long
   For Rdng% = 0 To mlAvgVal - 1
      GPIBWrite 0, "Read?"
      'DoEvents
      Reading$ = GPIBRead
      If mnUse32 Then
         DValRead# = Val(Reading$)
         CountValue& = GetHiResCounts(mnResolution, mnRange, DValRead#)
      Else
         ValueRead! = Val(Reading$)
         CountValue& = GetCounts(mnResolution, mnRange, ValueRead!)
      End If
      LongVals(0, Rdng%) = CountValue&
      If gnCancel Or mnCancel Then Exit Sub
   Next Rdng%
   If (mlHandle = 0) Then
      If mnUse32 Then
         mlHandle = cbWinBufAlloc32(TotalCount&)
         If SaveFunc(Me, WinBufAlloc32, mlHandle, TotalCount&, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mnAcqDataType = vbLong: mlAcqPoints = TotalCount&
      Else
         mlHandle = cbWinBufAlloc(TotalCount&)
         If SaveFunc(Me, WinBufAlloc, mlHandle, TotalCount&, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mnAcqDataType = vbInteger: mlAcqPoints = TotalCount&
      End If
   End If
   SourceRef& = mlHandle
   ULStat = cbWinBufToArray32(SourceRef&, LongBack(0, 0), 0, TotalCount&)
   
   CopyMemory ByVal SourceRef&, LongVals(0, 0), (TotalCount& * 4)
   tmrGoLoop.ENABLED = TimerState%
   ULStat = cbWinBufToArray32(SourceRef&, LongBack(0, 0), 0, TotalCount&)

End Sub

Public Sub SetVoltageValue(VoltVal As String)

   txtData.Text = VoltVal
   mnDataChange = True
   'ResetData
   
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

Public Sub GetEvent(EventType As Long, EventData As Long, EventParam As Long)

   EventType = mlEventType
   EventData = mlEventData 'mlEventSize
   EventParam = mlEventParam
   DoEvents
   
End Sub

Public Function GetInScanConfig() As Integer

   'true if input is simultaneous rather than multiplexed
   GetInScanConfig = mnSimIn
   
End Function

Public Function GetInstance() As Integer
   
   GetInstance = mnThisInstance
   
End Function

Public Function GetStaticOption() As Long

   GetStaticOption = mlStaticOpt
   
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
      Case "amplitude"
         msAmplitude = PropVal
      Case "cancel"
         mnCancel = PropVal
      Case "genoffset"
         msGenOffset = PropVal
      Case "genvolts"
         msGenVolts = PropVal
      Case "staticrange"
         mlStaticRange = PropVal
      'Case "resolution"
         'not set here - calculated or ini setting
      'Case "siminput"
         'not set here - calculated or ini setting
      Case "threshold"
         msThreshold = PropVal
      Case Else
         MsgBox "Attempt was made to set a property (" & PropName & _
         ") that cannot be set by a script.", vbOKOnly, "Read Only Property"
         SetFormProperty = False
   End Select

End Function

Public Sub SetStaticOption(NewOption As Long)

   If NewOption = 0 Then
      mlStaticOpt = NewOption
   Else
      mlStaticOpt = mlStaticOpt Or NewOption
   End If
   mnRefreshProps = True
   If (mnFuncType = AOut) Then
      If (NewOption = BACKGROUND) Or (NewOption = _
      CONTINUOUS) Then mnFuncType = AOutScan
   End If
   
End Sub

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

Private Sub MsgLibrary_ErrorAvailable(ByVal DeviceID As String, ByVal ErrCode As Long)

   gnCancel = True
   MsgResult$ = "Error- " & MsgLibrary.ConvertErrorCode(ErrCode)
   x% = SaveMsg(Me, "CallbackOnError (" & Format(ErrCode, "0") & ")", MsgResult$)
   tmrCheckStatus.ENABLED = False
   tmrGoLoop.ENABLED = False
   cmdGo.FontBold = True
   cmdStop = True
   cmdStop.Visible = False

End Sub

Private Sub MsgLibrary_DataAvailable(ByVal DeviceID As String, _
ByVal EventType As Long, ByRef EventData)

   MsgEventType& = Switch(EventType = 1, ON_DATA_AVAILABLE, _
   EventType = 2, ON_END_OF_INPUT_SCAN, EventType = 4, ON_SCAN_ERROR)

   MsgResult$ = "Event Occurred"
   If SaveMsg(Me, "MsgEvent(" & DeviceID & ", " & Format(EventType, "0") _
   & ", " & Format(EventData, "0") & ")", MsgResult$) Then
      Exit Sub
   End If
   If Not IsNull(EventData) Then
      'Chans% = (mnLastChan - mnFirstChan) + 1
      Select Case MsgEventType&
         Case ON_END_OF_INPUT_SCAN
            Me.cmdStop.Visible = False
            NumSamples& = EventData
            mnBufferEmpty = True
            Success% = TransferDataFromMsg(NumSamples&)
            mlEventData = NumSamples&
         Case ON_DATA_AVAILABLE
            If Me.mnuContPlot.Checked Then
               NumSamples& = EventData
               If Not mnEventTimeout Then Success% = TransferDataFromMsg(NumSamples&)
               If Not Success% Then
                  mnEventTimeout = True
                  'MsgResult$ = MsgLibrary.DisableEvents(1)
                  'If SaveMsg(Me, "DisableEvents(1)", MsgResult$) Then Exit Sub
               End If
            End If
            lblStatus = "Samples available = " & Format(EventData, "0")
            lblStatus.BackColor = lblStatus.BackColor Xor &HF00000
         Case Else
         
      End Select
   End If
   'Chans% = UBound(DataArray, 1)
   'SampsPerChan& = UBound(DataArray, 2) + 1
   'BufCount& = SampsPerChan& * (Chans% + 1)
   'If Not WDblArrayToBuf(Me, mlHandle, DataArray(), BufCount&, UseWinAPI%) Then
   '   ScalingData% = mnuToEng.Checked
   '   SpecifyMsgLib& = &H100
   '   '0x40 = FROMMESSAGE, 0x20 = float, 3 = double
   '   '0x10 = scaled, 0x80 = BUFFER
   '   If ScalingData% Then
   '      BufType% = 3 Or &H20 Or &H10
   '   Else
   '      BufType% = 3 Or &H20
   '   End If
   '   SetBufferType BufType%
   '   PlotBuffer mlHandle, BufCount&, Chans%
   'End If
   'PlotFromCallback CBArray, BufCount&, Chans%
   
End Sub

Private Sub NetLibrary_FunctionResultsAvailable(ByVal contents As String)

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
      Case AIn
         mvDataPoint = A(2)
      Case AInScan
         mfRateReturned = A(3)
      Case APretrig
         mlPTCountReturn = A(2)
         mfRateReturned = A(4)
      Case TIn
         mvDataPoint = A(2)
      Case TInScan
         'mvDataPoint = A(3)
      Case AOutScan
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
                  'DisplayData
               Else
                  'If mnuGetStatus.Checked Or mnuGetSubSystemStatus.Checked Then tmrCheckStatus.Enabled = True
                  'cmdStop.Visible = True
                  'cmdStop.Enabled = True
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
      Case VIn
         mvDataPoint = A(2)
      Case AIn32
         mvDataPoint = A(2)
      Case VIn32
         mvDataPoint = A(2)
   End Select
   mnThreadState = TS_IDLE

End Sub

Private Sub NetLibrary_FunctionCalled(ByVal FuncName As String)

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
      Case AInScan
      Case TInScan
         NumPoints% = UBound(Elements) - 2
         ReDim TempData(NumPoints%) As Single
         For i% = 0 To NumPoints%
            TempData(i%) = Val(Elements(i% + 1))
         Next
         mvDataPoint = TempData
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
    
   Static Do2ndHalf%, DoFirstHalf%, Element&, StillZero%
   If mnCancel Or gnCancel Then
      tmrGoLoop.ENABLED = False
      cmdGo.FontBold = True
      cmdStop = True
      gnCancel = False
      mnCancel = False
      If mnSetActive Then SetContActive False
      Element& = 0
      If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
      If DevActive% Then
         BGOp$ = "Background task in progress..."
      Else
         BGOp$ = "No background task running..."
      End If
      Stat$ = msStringID & BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
      lblStatus.Caption = Stat$
      If mnuStopBG.Checked Then
         If mnThreading Then
            functionToCall$ = GetCallAsString(StopBackground, mnBoardNum, mlStatusType)
            'AcqThread.CallULFunction functionToCall$
         Else
            ULStat = StopBackground520(mnBoardNum, mlStatusType)
         End If
         If (Not gnScriptSave) Or (ULStat <> 0) Then
            If Not mnThreading Then
               If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
            End If
         End If
         cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      DoEvents
      mnPlot = True
      Exit Function
   End If
   'If mnuGetStatus.Checked Then
   '   If mnThreading Then
   '      functionToCall$ = GetCallAsString(GetStatus, mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
   '      AcqThread.CallULFunction functionToCall$
   '      Do
   '         DoEvents
   '      Loop While Not (mnThreadState = TS_IDLE)
   '   Else
   '      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
   '   End If
      
      A2 = Status%
      A3 = CurCount&
      A4 = CurIndex&
   '   If mbEvalEnabled And (frmEvalData.chkCountStop.value = 1) Then EvalResult% = EvalCount(CurCount&)
   'End If
   'If CurCount& Mod 5 <> 0 Then Stop
   If mnuSaveStatus.Checked Then
      malStat(0, Element&) = CurCount&
      malStat(1, Element&) = CurIndex&
      If (malStat(1, Element&) + 1) Mod 5 <> 0 Then Stop
      Element& = Element& + 1
      If Element& > UBound(malStat, 2) Then Element& = 0
   End If
   If (Not gnScriptSave) Or (ULStat <> 0) Then
      If Not mnThreading Then
         If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, A2, A3, A4, mlStatusType, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
      End If
      If Failure% And Not (ULStat = TOOFEW) Then
         If mnuStopBG.Checked Then
            If mnThreading Then
               functionToCall$ = GetCallAsString(StopBackground, mnBoardNum, mlStatusType)
               'AcqThread.CallULFunction functionToCall$
            Else
               ULStat = StopBackground520(mnBoardNum, mlStatusType)
               If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
            End If
         End If
         cmdStop.Visible = False
         Status% = 0
         Do2ndHalf% = False
         DoFirstHalf% = False
         tmrCheckStatus.ENABLED = False
         DoEvents
         mnPlot = True
         Exit Function
      End If
   End If
   
   If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
   If DevActive% Then
      BGOp$ = "Background task in progress..."
      Stat$ = msStringID & BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
      lblStatus.Caption = Stat$
   End If
   If mnuContPlot.Checked Then
      BufferSplit& = mlCount \ 2
      If (Not Do2ndHalf%) And (Not DoFirstHalf%) Then
         mlFirstPoint = 0
         If CurIndex& > BufferSplit& Then
            DoFirstHalf% = True
            mnPlot = mnuPlotType(0).Checked
            mbEval = mbEvalEnabled
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
      Element& = 0
      lblStatus.Caption = msStringID & msConfig & " " & msOpt
      If mnuStopBG.Checked Then
         If mnThreading Then
            functionToCall$ = GetCallAsString(StopBackground, mnBoardNum, mlStatusType)
            'AcqThread.CallULFunction functionToCall$
         Else
            ULStat = StopBackground520(mnBoardNum, mlStatusType)
         End If
         If (Not gnScriptSave) Or (ULStat <> 0) Then
            If Not mnThreading Then
               If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mlStatusType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
            End If
         End If
         If Not tmrGoLoop.ENABLED And (Not mnLoop) Then cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
      If DevActive% Then
         BGOp$ = "Background task in progress..."
         Stat$ = msStringID & BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
         lblStatus.Caption = Stat$
         DoEvents
      End If
      mnPlot = mnuPlotType(0).Checked _
         Or mnuPlotType(1).Checked _
         Or mnuPlotType(2).Checked _
         Or mnuPlotType(DERIVATIVE).Checked
      mbEval = mbEvalEnabled
      If (Not (mnFormType = ANALOG_OUT)) And ((Not mbEval) Or (Not mnuContPlot.Checked)) Then
         DisplayData
      End If
   Else
      Me.cmdStop.Visible = True
      Me.cmdStop.ENABLED = True
      If ((CurIndex& < 1) And (CurCount& > mlCount)) And mnuContPlot.Checked Then
         'if CurIndex remains 0 because continuous is being run with a size
         'below block size, check status again to see if still 0, if so, plot
         'doesn't apply to miniLAB
         If mnuGetStatus.Checked Then
            If mnThreading Then
               functionToCall$ = GetCallAsString(GetStatus, mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
               'AcqThread.CallULFunction functionToCall$
               Do
                  DoEvents
               Loop While Not (mnThreadState = TS_IDLE)
            Else
               ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
            End If
         End If
         
         If StillZero% Then
            mnPlot = mnuPlotType(0).Checked _
               Or mnuPlotType(1).Checked _
               Or mnuPlotType(2).Checked _
               Or mnuPlotType(DERIVATIVE).Checked
            mbEval = mbEvalEnabled
            If (Not (mnFormType = ANALOG_OUT)) And ((Not mbEval) Or (Not mnuContPlot.Checked)) Then
               DisplayData
            End If
         End If
         If (CurIndex& < 1) And (CurCount& > mlCount) Then StillZero% = True
      Else
         If Status% = RUNNING Then
            BGOp$ = "Background task in progress..."
            Stat$ = msStringID & BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
            lblStatus.Caption = Stat$
         End If
         StillZero% = False
      End If
   End If

End Function

Public Sub ShowBufferInfo(ByVal Source As Integer)
   
   If (mnCalledLocally And (Source = 1)) Then Exit Sub 'mnAIScanDataRead Or
   If IsEmpty(mlHandle) Or (mlHandle = 0) Or (mnFuncType = AOutScan) Then
      'check if there's an output buffer
      If Not ((mlGenHandle = 0) Or (mlGenPoints = 0)) Then
         Range% = mnRange
         lblStatus.Caption = msStringID & msConfig & " " & msOpt
         If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
            RateVal = mfRateReturned / 1000 & "Hz"
         Else
            RateVal = mfRateReturned / 1000 & "kHz"
         End If
         SetDetails mnFuncType, mnFirstChan, mnLastChan, mlCount, _
         RateVal, Range%, msBoardName, mlPTCountReturn
         PrintMain "Buffer handle = 0x" & Hex$(mlGenHandle) & _
         ",  buffer size = " & mlGenPoints
         'Source = 3 means called from PlotGenData
         If Not Source = 3 Then PlotGenData
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
         Exit Sub
      Case BUFFERTOOSMALL
         MsgBox "Windows buffer too small (size = " & BufSize& _
         & ", data requested = " & mlCount * 2 & ")", , "Buffer Status"
         Exit Sub
      Case NOBUFFER
         MsgBox "No Windows buffer has been allocated", , "Buffer Status"
         Exit Sub
      Case Else
         BufInfoString$ = "Buffer handle = 0x" & Hex$(mlHandle) & ",  buffer size = " & BufSize&
         PlotData% = True
   End Select
   SavePlotState% = mnPlot
   mnPlot = True
   'Conversions are already done and saved in buffer
   'don't repeat
   Select Case mnLibType
      Case UNILIB
         PrintMain BufInfoString$
         PTConvertState% = mnuConvPT.Checked
         mnuConvPT.Checked = False
         If PlotData% Then DisplayData
         mnPlot = SavePlotState%
         mnuConvPT.Checked = PTConvertState%
      Case MSGLIB
         If mnBufferEmpty Then
            CurCount& = 0
            CurIndex& = 0
         Else
            Status% = MsgCheckStatus(Me, msBoardName, _
            MsgLibrary, msStatusType, CurCount&, CurIndex&)
         End If
         PrintMain BufInfoString$ & " (" & CurCount& & ", " & CurIndex& & ")"
         NumSamps& = mlCount
         If CurCount& < mlCount Then NumSamps& = CurCount&
         'NumSamps& = 300
         If mnMessaging And Not mnBufferEmpty Then Success% = TransferDataFromMsg(NumSamps&)
   End Select
   
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SBufInfo, FuncStat, _
         A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Public Sub InitForm(FunctionInit As Integer)

   Select Case mnFuncType
      Case AIn, AIn32
         msStatusType = "AI"
      Case AInScan
         msStatusType = "AISCAN"
      Case AOut
         msStatusType = "AO"
      Case AOutScan
         msStatusType = "AOSCAN"
   End Select
   If Not mnFormInitialized Then mnuBoard_Click (0)
   If mnLibType = MSGLIB Then
      If Not mnMessaging Then Exit Sub
      Select Case mnFormType
         Case ANALOG_IN
            msStatusType = "AI"
            NumProps% = GetAIProps(msBoardName, MsgLibrary, PropList)
            UpdateFormProps PropList, NumProps%
            If InStr(1, msScanSupport, "QUEUE") > 0 Then QChanged% = ConfigureMsgQueue()
            msStatusType = "AISCAN"
            NumProps% = GetAIScanProps(msBoardName, MsgLibrary, PropList)
            If NumProps% < 0 Then
               msStatusType = "AI"
            Else
               For i% = 0 To NumProps%
                  masDefaultScanProps(mnBoardIndex) = _
                  masDefaultScanProps(mnBoardIndex) & Separator$ & PropList(i%)
                  Separator$ = ";"
               Next
               UpdateFormProps PropList, NumProps%
            End If
      End Select
   End If
   If Not mnFormInitialized Then
      If mnFuncType = 0 Then mnuFuncArray_Click (FunctionInit)
   End If
   mnFormInitialized = True
   If Not (gnNumBoards > 0) Then
      For Ctl& = 0 To Me.Controls.Count - 1
         ContType$ = Left(Controls(Ctl&).Name, 3)
         If (ContType$ = "cmd") Or _
            (ContType$ = "chk") Or _
            (ContType$ = "hsb") Then _
            Controls(Ctl&).ENABLED = False
      Next
      Exit Sub
   End If
   
End Sub

Public Sub ConfigurationChange(ConfigType As String)

   ParseType$ = Mid(ConfigType, 2, 6)
   TypeSupported% = (InStr(1, msAiSupport, ParseType$) > 0)
   TypeSupported% = TypeSupported% Or (InStr(1, msScanSupport, ParseType$) > 0)
   If TypeSupported% Then
      Select Case ParseType$
         Case "CAL"
            QueryMsg = "?" & msStatusType & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewCal$ = ParseMsg(1)
            For OptIndex% = 0 To mnuOpt.Count - 1
               If mnuOpt(OptIndex%).Caption = "NOCALIBRATEDATA" Then Exit For
            Next
            NotCalibrating% = (NewCal$ = "DISABLE")
            mnuOpt(OptIndex%).Checked = NotCalibrating%
            Options = Split(msOpt, "NOCALIBRATEDATA")
            If NotCalibrating% Then
               mvOptions = (mvOptions Or NOCALIBRATEDATA)
               If Not UBound(Options) > 0 Then msOpt = Options(0) & "NOCALIBRATEDATA"
            Else
               mvOptions = (&HFFFF7FFF And mvOptions)
               If UBound(Options) > 0 Then msOpt = Options(0) & Options(1)
            End If
            UpdateStatBar False
         Case "CHMODE"
            QueryMsg = "?AI" ' & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewChanCount$ = ParseMsg(1)
            mnNumAIChans = Val(NewChanCount$)
            LastElement& = mnNumAIChans - 1
            ReDim Preserve manGainArray(LastElement&)
            ReDim Preserve manChanArray(LastElement&)
            ReDim Preserve manModeArray(LastElement&)
            ReDim Preserve manDRateArray(LastElement&)
         Case "EXTPAC"
            QueryMsg = "?" & msStatusType & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewCal$ = ParseMsg(1)
            For OptIndex% = 0 To mnuOpt.Count - 1
               If mnuOpt(OptIndex%).Caption = "EXTCLOCK" Then Exit For
            Next
            ParseCfg = Split(NewCal$, "/")
            If UBound(ParseCfg) > 0 Then
               ParsedPacer$ = ParseCfg(1)
               State$ = ParseCfg(0)
               ExtPacing% = ((ParsedPacer$ = "SLAVE") Or (ParsedPacer$ = "GSLAVE"))
               ExtPacing% = (ExtPacing% And (State$ = "ENABLE"))
            Else
               ExtPacing% = (NewCal$ = "ENABLE")
            End If
            mnuOpt(OptIndex%).Checked = ExtPacing%
            Options = Split(msOpt, "EXTCLOCK")
            If ExtPacing% Then
               mvOptions = (mvOptions Or EXTCLOCK)
               If Not UBound(Options) > 0 Then msOpt = Options(0) & "EXTCLOCK"
            Else
               mvOptions = (&HFFFFFFFB And mvOptions)
               If UBound(Options) > 0 Then msOpt = Options(0) & Options(1)
            End If
            UpdateStatBar False
         Case "LOWCHA", "HIGHCH"
            QueryMsg = "?AISCAN" & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then
               If ParseMsg(0) = "AISCAN:LOWCHAN" Then
                  NewLow$ = ParseMsg(1)
                  ChanChanged% = (mnFirstChan <> Val(NewLow$))
                  If ChanChanged% Then txtLowChan.Text = NewLow$
               End If
               If ParseMsg(0) = "AISCAN:HIGHCHAN" Then
                  NewHigh$ = ParseMsg(1)
                  ChanChanged% = (mnLastChan <> Val(NewHigh$))
                  If ChanChanged% Then txtHighChan.Text = NewHigh$
               End If
            End If
         Case "RATE"
            QueryMsg = "?AISCAN" & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewRate$ = ParseMsg(1)
            mfRate = Val(NewRate$)
            Me.txtRate.Text = NewRate$
         Case "SAMPLE"
            QueryMsg = "?AISCAN" & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewCount$ = ParseMsg(1)
            Me.txtCount.Text = NewCount$
            ResetData
         Case "SCALE"
            QueryMsg = "?" & msStatusType & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewScale$ = ParseMsg(1)
            mnuToEng.Checked = (NewScale$ = "ENABLE")
         Case "TRIG"
            QueryMsg = "?" & msStatusType & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then NewCal$ = ParseMsg(1)
            For OptIndex% = 0 To mnuOpt.Count - 1
               If mnuOpt(OptIndex%).Caption = "EXTTRIGGER" Then Exit For
            Next
            'If UBound(ParseCfg) > 0 Then
            '   ParsedPacer$ = ParseCfg(1)
            '   State$ = ParseCfg(0)
            '   EXTTRIGGER% = ((ParsedPacer$ = "SLAVE") Or (ParsedPacer$ = "GSLAVE"))
            '   ExtPacing% = (ExtPacing% And (State$ = "ENABLE"))
            'Else
               ExtPacing% = (NewCal$ = "ENABLE")
            'End If
            mnuOpt(OptIndex%).Checked = ExtPacing%
            Options = Split(msOpt, "EXTTRIGGER")
            If ExtPacing% Then
               mvOptions = (mvOptions Or EXTTRIGGER)
               If Not UBound(Options) > 0 Then msOpt = Options(0) & "EXTTRIGGER"
            Else
               mvOptions = (&HFFFFBFFF And mvOptions)
               If UBound(Options) > 0 Then msOpt = Options(0) & Options(1)
            End If
            UpdateStatBar False
         Case "QUEUE"
            QChanged% = ConfigureMsgQueue()
            If QChanged% Then
               txtLowChan.Text = "0"
               txtHighChan.Text = "0"
               If mlQCount > 0 Then txtHighChan.Text = Format$(mlQCount - 1, "0")
            End If
         Case "RANGE{"
            QChanged% = ConfigureMsgQueue()
            If QChanged% Then
               txtLowChan.Text = "0"
               txtHighChan.Text = "0"
               If mlQCount > 0 Then txtHighChan.Text = Format$(mlQCount - 1, "0")
            End If
         Case "XFRMOD"
            QueryMsg = "?" & msStatusType & ConfigType
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then XMode$ = ParseMsg(1)
            For OptIndex% = 0 To mnuOpt.Count - 1
               If mnuOpt(OptIndex%).Caption = "SINGLEIO" Then SioIndex% = OptIndex%
               If mnuOpt(OptIndex%).Caption = "BLOCKIO" Then BlockIndex% = OptIndex%
               If mnuOpt(OptIndex%).Caption = "BURSTIO" Then BurstIndex% = OptIndex%
            Next
            Select Case XMode$
               Case "BLOCKIO"
                  mnuOpt(BurstIndex%).Checked = False
                  mnuOpt(SioIndex%).Checked = False
                  mnuOpt(BlockIndex%).Checked = True
               Case "BURSTIO"
                  mnuOpt(BurstIndex%).Checked = True
                  mnuOpt(SioIndex%).Checked = False
                  mnuOpt(BlockIndex%).Checked = False
               Case "SINGLEIO"
                  mnuOpt(BurstIndex%).Checked = False
                  mnuOpt(SioIndex%).Checked = True
                  mnuOpt(BlockIndex%).Checked = False
            End Select
            SetOptions
            'mnuOpt(OptIndex%).Checked = NotCalibrating%
            'Options = Split(msOpt, "NOCALIBRATEDATA")
            'If NotCalibrating% Then
            '   mvOptions = (mvOptions Or NOCALIBRATEDATA)
            '   If Not UBound(Options) > 0 Then msOpt = Options(0) & "NOCALIBRATEDATA"
            'Else
            '   mvOptions = (&HFFFF7FFF And mvOptions)
            '   If UBound(Options) > 0 Then msOpt = Options(0) & Options(1)
            'End If
            UpdateStatBar False
      End Select
   End If
   
End Sub

Private Sub UpdateFormProps(PropList As Variant, ListSize As Integer)

   'ResetOpts
   For i% = 0 To ListSize
      ListedProp$ = PropList(i%)
      ParseProp = Split(ListedProp$, "=")
      If UBound(ParseProp) > 0 Then
         Prop$ = ParseProp(0)
         PropVal$ = ParseProp(1)
         Select Case Prop$
            Case "AI"
               If IsNumeric(PropVal$) Then mnNumAIChans = Val(PropVal$)
            Case "AO"
               If IsNumeric(PropVal$) Then mnNumAOChans = Val(PropVal$)
            Case "CAL"
               If PropVal$ = "DISABLE" Then
                  OptionVal& = GetOptionCodeFromMsg(Prop$, OptionIndex%)
                  If Not mnuOpt(OptionIndex%).Checked Then
                     If Not OptionIndex% = -1 Then mnuOpt_Click (OptionIndex%)
                  End If
               End If
            Case "XFRMODE"
               OptionVal& = GetOptionCodeFromMsg(PropVal$, OptionIndex%)
               If Not OptionIndex% = -1 Then mnuOpt_Click (OptionIndex%)
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
                  mfRateReturned = mfRate
                  Me.txtRate.Text = Format(mfRate, "0.0####")
               End If
            Case "SAMPLES"
               If Not (txtCount.Text = PropVal$) Then
                  If IsNumeric(PropVal$) Then
                     Me.txtCount.Text = PropVal$
                  Else
                     PropVal$ = "1"
                     Me.txtCount.Text = PropVal$
                  End If
                  mlCount = PropVal$
                  mnDetailsNeedUpdate = True
               End If
            Case "LOWCHAN"
               If IsNumeric(PropVal$) Then
                  Me.hsbLowChan.value = Val(PropVal$)
               End If
            Case "HIGHCHAN"
               If IsNumeric(PropVal$) Then
                  HCVal& = Val(PropVal$)
                  If HCVal& < hsbHighChan.Max Then
                     Me.hsbHighChan.value = Val(PropVal$)
                  Else
                     Me.txtHighChan.Text = Format(HCVal&, "0")
                  End If
               End If
            Case "QUEUE"
               Select Case PropVal$
                  Case "ENABLE", "DISABLE", "RESET"
                     If Not mnQueueConfigurable Then mnQueueSupported = True
                  Case "NOTSUPPORTED"
                     mnQueueSupported = False
               End Select
               'ConfigureMsgQueue
            Case "RANGE{0}", "AI{0}:RANGE"
               mnRange = GetRangeCodeFromMsg(PropVal$)
               If (Not mnRange = 0) And (mnNumAIChans > 0) Then QChanged = ConfigureMsgQueue()
            Case "AITRIG:TYPE"
               mnTrigType = GetTrigTypeFromMsg(PropVal$)
               If Not mnuTrigType(mnTrigType).Checked Then
                  For TrigMenu% = 0 To mnuTrigType.Count - 1
                     mnuTrigType(TrigMenu%).Checked = False
                  Next
                  mnuTrigType(mnTrigType).Checked = True
               End If
            Case "AITRIG:REARM"
               RetrigSet% = (PropVal$ = "ENABLE")
               If Not mnuOpt(17).Checked = RetrigSet% Then mnuOpt_Click (17) 'retrigger
         End Select
      End If
   Next
   UpdateStatBar False
   
End Sub

Sub WritePropsToDevice()

   ScalingData% = mnuToEng.Checked
   ScaleString$ = "=DISABLE"
   If ScalingData% Then ScaleString$ = "=ENABLE"
   NotCalibratingData% = (mvOptions And NOCALIBRATEDATA) = NOCALIBRATEDATA
   CalString$ = "=ENABLE"
   If NotCalibratingData% Then CalString$ = "=DISABLE"
   Prop$ = ":SCALE"
   MessageUpdate$ = msStatusType & Prop$ & ScaleString$
   'to do - implement ToEng and remove following conditional
   If Not (msStatusType = "AO") Then GoSub DeliverMessage
   If Not ((msStatusType = "AO") Or (msStatusType = "AOSCAN")) Then
      'to do - this may be legit on future devices
      Prop$ = ":CAL"
      MessageUpdate$ = msStatusType & Prop$ & CalString$
      GoSub DeliverMessage
   End If
   
   If (mnFuncType = AInScan) Or (mnFuncType = AOutScan) Then
      If (mvOptions And CONTINUOUS) = CONTINUOUS Then
         SamplesString$ = "0"
      Else
         SamplesString$ = txtCount.Text
         If mnuCal.Checked Then SamplesString$ = Format(mlAvgVal, "0")
      End If
      MessageUpdate$ = msStatusType & ":SAMPLES=" & SamplesString$
      GoSub DeliverMessage
   End If
   If Not ((msStatusType = "AO") Or (msStatusType = "AOSCAN")) Then
      'to do - this may be legit on future devices
      If mlQCount = 0 Then
         For Chan% = mnFirstChan To mnLastChan
            If Chan% > UBound(manGainArray) Then Exit Sub
            RangeCode% = manGainArray(Chan%)
            MsgRange$ = GetMsgRange(RangeCode%)
            RangeString$ = "=" & MsgRange$
            ChanString$ = "{" & Format(Chan%, "0") & "}"
            Prop$ = ":RANGE"
            If mnFuncType = AInScan Then
               If InStr(1, msScanSupport, "AISCAN:RANGE{*}=") > 0 Then WriteRange% = True
               MessageUpdate$ = msStatusType & Prop$ & ChanString$ & RangeString$
            Else
               If InStr(1, msAiSupport, "AI{*}:RANGE=") > 0 Then WriteRange% = True
               MessageUpdate$ = msStatusType & ChanString$ & Prop$ & RangeString$
            End If
            If WriteRange% Then
               GoSub DeliverMessage
            Else
               Exit For
            End If
         Next
      End If
   End If
   
   Select Case mnFuncType
      Case AIn, ATrig, TIn, TInScan, VIn, AIn32
         'to do - all chans set alike now - may want per-chan settings
         'For Chan% = mnFirstChan To mnLastChan
         '   ChanString$ = "{" & Format(Chan%, "0") & "}"
         '   Prop$ = ":SCALE="
         '   MessageUpdate$ = msStatusType & ChanString$ & Prop$ & ScaleString$
         '   GoSub DeliverMessage
         '   Prop$ = ":CAL="
         '   MessageUpdate$ = msStatusType & ChanString$ & Prop$ & CalString$
         '   GoSub DeliverMessage
         'Next
      Case AInScan
         'If (mvOptions And CONTINUOUS) = CONTINUOUS Then
         '   SamplesString$ = "0"
         'Else
         '   SamplesString$ = txtCount.Text
         'End If
         'MessageUpdate$ = msStatusType & ":SAMPLES=" & SamplesString$
         'GoSub DeliverMessage
         If Not mlQEnabled Then
            MessageUpdate$ = msStatusType & ":LOWCHAN=" & Format(mnFirstChan, "0")
            GoSub DeliverMessage
            MessageUpdate$ = msStatusType & ":HIGHCHAN=" & Format(mnLastChan, "0")
            GoSub DeliverMessage
         End If
         SetMsgRate
         NotCalibratingData% = (mvOptions And NOCALIBRATEDATA) = NOCALIBRATEDATA
         CalString$ = "ENABLE"
         If NotCalibratingData% Then CalString$ = "DISABLE"
         MessageUpdate$ = msStatusType & ":CAL=" & CalString$
         GoSub DeliverMessage
         ScalingData% = mnuToEng.Checked
         ScaleString$ = "DISABLE"
         If ScalingData% Then ScaleString$ = "ENABLE"
         MessageUpdate$ = msStatusType & ":SCALE=" & ScaleString$
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

Private Function ConfigureMsgQueue(Optional SetQueue As Integer = False) As Integer

   If mnQueueSupported Or mnQueueConfigurable Then
      QueueMsg$ = "?AISCAN:QUEUE"
      MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
      QError% = SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$)
      If Not QError% = 0 Then
         QEnabled% = False
      Else
         QResult = Split(MsgResult$, "=")
         If UBound(QResult) > 0 Then
            QEnabled% = (QResult(1) = "ENABLE")
         Else
            QEnabled% = False
         End If
      End If
      QChanged% = Not (mlQEnabled = QEnabled%)
      mlQEnabled = QEnabled%
   End If
   
   If Not SetQueue Then
      If mnQueueSupported Or mnQueueConfigurable Then
         'don't mess with range if queue is enabled
         If QEnabled% Then 'Exit Function
            'find out if device has a queue set
            If mnQueueSupported Then
               QueueString$ = ":RANGE"
               QueueMsg$ = "?" & msStatusType & QueueString$
               MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            ElseIf mnQueueConfigurable Then
               QueueMsg$ = "?AIQUEUE:COUNT"
               MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            End If
            'to do - after queue fixed, remove this
            If Left(MsgResult$, 36) = "Warning- The device does not support" Then
               If mnNumAIChans = 0 Then Exit Function
               LastElement& = mnNumAIChans - 1
            Else
               If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then Exit Function
               ValueLoc& = InStr(1, MsgResult$, "=")
               If ValueLoc& > 0 Then
                  QueueState$ = Mid(MsgResult$, ValueLoc& + 1)
                  If IsNumeric(QueueState$) Then
                     QCount& = Val(QueueState$)
                     QChanged% = QChanged% Or (Not (mlQCount = QCount&))
                     mlQCount = QCount
                     'if device does have a queue set, then populate the
                     'queue arrays from the values rec'd from the device
                     If mlQCount > 0 Then
                        LastElement& = mlQCount - 1
                        GetQueueFromDevice% = True
                     Else
                        LastElement& = mnNumAIChans - 1
                     End If
                  Else
                     'no queue is set up on device
                     LastElement& = mnNumAIChans - 1
                  End If
               End If
            End If
         Else
            LastElement& = mnNumAIChans - 1
         End If
      Else
         If mnNumAIChans = 0 Then Exit Function
         LastElement& = mnNumAIChans - 1
      End If
      If Not LastElement& < 0 Then
         ReDim Preserve manGainArray(LastElement&)
         ReDim Preserve manChanArray(LastElement&)
         ReDim Preserve manModeArray(LastElement&)
         ReDim Preserve manDRateArray(LastElement&)
      End If
   Else
      If Not QEnabled% Then
         QueueString$ = ":QUEUE=ENABLE"
         QueueMsg$ = msStatusType & QueueString$
         MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
         If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then Failure% = True
         QChanged% = QChanged% Or (Not Failure%)
      End If
      LastElement& = mlQCount - 1
   End If
      
   For Element% = 0 To LastElement&
      If (Not SetQueue) Or GetQueueFromDevice% Then
         If mnQueueSupported Then
            ElementString$ = "{" & Format(Element%, "0") & "}"
            QueueString$ = ":RANGE"
            If (mnFuncType = AIn) Or (msStatusType = "AI") Then
               RangeMsg$ = msStatusType & ElementString$ & QueueString$
            Else
               RangeMsg$ = msStatusType & QueueString$ & ElementString$
            End If
            QueueMsg$ = "?" & RangeMsg$
            MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            If (msDisplayName = "USB-7204") And (Not Element% < _
                mnNumAIChans) And (Element% < 17) Then
                If (Not mbIgnore578045) Then
                    Dim Resp As VbMsgBoxResult
                    Resp = MsgBox("DAQFlex bug reading valid queue elements." & vbCrLf & _
                        "Ignore subsequent instances of this error?", vbYesNo, "CAR 578045")
                    If Resp = vbYes Then mbIgnore578045 = True
                End If
            Else
                If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then Exit Function
            End If
            SplitMsg = Split(MsgResult$, "/")
            If UBound(SplitMsg) > 0 Then
               QueueVals$ = SplitMsg(1)
               SplitVals = Split(QueueVals$, "}=")
               If UBound(SplitVals) > 0 Then
                  ChanElement% = Val(SplitVals(0))
                  RangeString$ = SplitVals(1)
                  RangeElement% = GetRangeCodeFromMsg(RangeString$)
               End If
            End If
         ElseIf mnQueueConfigurable Then
            'don't load values if queue is not enabled
            'these devices don't keep same queue for both en/disabled
            If Not QEnabled% Then
               ConfigureMsgQueue = False
               Exit Function
            End If
            QueueString$ = "?AIQUEUE{" & Format(Element%, "0") & "}:"
            For QPart% = 1 To 4
               ItemValid% = Choose(QPart%, -1, _
               -1, mnQueueChanMode, mnQueueDataRate)
               If ItemValid% Then
                  QueueItem$ = Choose(QPart%, "CHAN", _
                  "RANGE", "CHMODE", "DATARATE")
                  QueueMsg$ = QueueString$ & QueueItem$
                  MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
                  If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then
                     Failure% = True
                     Exit For
                  End If
                  SplitMsg = Split(MsgResult$, "=")
                  If UBound(SplitMsg) > 0 Then
                     QueueVals$ = SplitMsg(1)
                     Select Case QPart%
                        Case 1
                           ChanElement% = Val(QueueVals$)
                        Case 2
                           RangeString$ = QueueVals$
                           RangeElement% = GetRangeCodeFromMsg(RangeString$)
                        Case 3
                           ModeElement% = Val(QueueVals$)
                        Case 4
                           DRateElement% = Val(QueueVals$)
                     End Select
                  End If
               End If
            Next
            QChanged% = QChanged% Or (manChanArray(Element%) <> ChanElement%)
            QChanged% = QChanged% Or (manGainArray(Element%) <> RangeElement%)
            'QChanged% = QChanged% Or (manModeArray(Element%) <> ModeElement%)
            'QChanged% = QChanged% Or (manDRateArray(Element%) <> DRateElement%)
            If Not gnScriptRun Then
                'should not be required if script is running
                manChanArray(Element%) = ChanElement%
                manGainArray(Element%) = RangeElement%
            End If
            'manModeArray(Element%) = ModeElement%
            'manDRateArray(Element%) = DRateElement%
         End If
      Else
         QueueRange$ = GetMsgRange(manGainArray(Element%))
         If mnQueueSupported Then
            ElementString$ = Format(Element%, "0") & "/"
            ChannelString$ = Format(manChanArray(Element%), "0") & "}="
            QueueString$ = ":RANGE{"
            QueueMsg$ = msStatusType & QueueString$ & ElementString$ & _
            ChannelString$ & QueueRange$
            MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
            If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then Failure% = True
         ElseIf mnQueueConfigurable Then
            QueueString$ = "AIQUEUE{" & Format(Element%, "0") & "}:"
            ChannelString$ = "CHAN=" & Format(manChanArray(Element%), "0")
            RangeString$ = "RANGE=" & QueueRange$
            QueueMode$ = GetMsgChMode(manModeArray(Element%))
            ModeString$ = "CHMODE=" & QueueMode$
            DRateString$ = "DATARATE=" & Format(manDRateArray(Element%), "0")
            DontSetRange% = InStr(1, msScanSupport, "AISCAN:RANGE{*}=") = 0
            For QPart% = 1 To 4
               ItemValid% = Choose(QPart%, -1, _
               Not DontSetRange%, mnQueueChanMode, mnQueueDataRate)
               If ItemValid% Then
                  QueueItem$ = Choose(QPart%, ChannelString$, _
                  RangeString$, ModeString$, DRateString$)
                  QueueMsg$ = QueueString$ & QueueItem$
                  MsgResult$ = MsgLibrary.SendMessage(QueueMsg$)
                  If SaveMsg(Me, "SendMessage(" & QueueMsg$ & ")", MsgResult$) Then
                     Failure% = True
                     Exit For
                  End If
               End If
            Next
         End If
         QChanged% = QChanged% Or (Not Failure%)
      End If
   Next
   ConfigureMsgQueue = QChanged%
   
End Function

Sub SetMsgRate()

   If (msStatusType = "AI") Or (msStatusType = "AO") Then Exit Sub
   RateText$ = txtRate.Text
   RateParse = Split(RateText$, "e")
   If UBound(RateParse) > 0 Then
      Exponent$ = RateParse(1)
      Suffix$ = "e" & Exponent$
   End If
   RateString$ = Left(txtRate.Text, 7) & Suffix$
   NewRate! = Val(RateString$)
   
   If ((Not mnRefreshProps) And (mfRate = NewRate!)) Or (NewRate! = 0) Then Exit Sub
   If IsNumeric(RateString$) Then mfRate = Val(RateString$)
   RateMsg$ = msStatusType & ":RATE=" & RateString$
   MsgResult$ = MsgLibrary.SendMessage(RateMsg$)
   If SaveMsg(Me, "SendMessage(" & RateMsg$ & ")", MsgResult$) Then Exit Sub
   RateMsg$ = "?" & msStatusType & ":RATE"
   MsgResult$ = MsgLibrary.SendMessage(RateMsg$)
   If Not SaveMsg(Me, "SendMessage(" & RateMsg$ & ")", MsgResult$) Then
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         RateString$ = Mid(MsgResult$, ValueLoc& + 1)
         If IsNumeric(RateString$) Then
            mfRateReturned = Val(RateString$)
         End If
      End If
   End If

End Sub

Sub SetMsgCount()
   
   If msStatusType = "AI" Then Exit Sub
   If mnLibType = MSGLIB Then
      NumChans% = (mnLastChan - mnFirstChan) + 1
      'TotalSamples& = mlAvgVal * NumChans%
      If mnuCal.Checked Then
         SamplesString$ = Format(mlAvgVal, "0")
      Else
         SamplesString$ = txtCount.Text
         NewCount& = Val(SamplesString$)
         If NewCount& = mlCount Then Exit Sub
      End If
      If (mvOptions And CONTINUOUS) = CONTINUOUS _
      Then SamplesString$ = "0"
      
      SamplesMsg$ = msStatusType & ":SAMPLES=" & SamplesString$
      MsgResult$ = MsgLibrary.SendMessage(SamplesMsg$)
      If SaveMsg(Me, "SendMessage(" & SamplesMsg$ & ")", MsgResult$) Then
         SamplesMsg$ = "?" & msStatusType & ":SAMPLES=" & SamplesString$
         MsgResult$ = MsgLibrary.SendMessage(SamplesMsg$)
         ValueLoc& = InStr(1, MsgResult$, "=")
         If ValueLoc& > 0 Then
            SampsReturned$ = Mid(MsgResult$, ValueLoc& + 1)
            Me.txtCount.Text = SampsReturned$
         End If
      End If
   End If

End Sub

Public Sub SetOptionsMenuByValue(OptionVal As Long)

   For i% = 0 To mnuOpt.Count - 1
      mnuOpt(i%).Checked = False
   Next i%
   For i% = 0 To mnuOpt.Count - 1
      If ((2 ^ i%) And OptionVal) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 5 Then
            'if apparently SINGLEIO
            If ((2 ^ 6) And OptionVal) = (2 ^ 6) Then
               'check if also DMAIO (if both SIO & DIO then
               MenuIndex% = 7 'its actually BLOCKIO)
               i% = 6
            End If
         End If
         mnuOpt(MenuIndex%).Checked = True
      End If
   Next i%
   mnDetailsNeedUpdate = True
   SetOptions
   UpdateStatBar False

End Sub

Private Function TransferDataFromMsg(Optional ByVal DataAvailable As Variant, Optional PlotData As Integer = True) As Integer

   mlReadTimeout = Val(Me.txtOffset.Text)
   Chans% = (mnLastChan - mnFirstChan)
   NumChans% = (mnLastChan - mnFirstChan) + 1
   If mnDetailsNeedUpdate Then
      If Me.mnuCalLoopback.Checked Then
         BoardName = "HP34401 loopback"
      Else
         BoardName = mnuBoard(mnBoardIndex).Caption
      End If
      lblStatus.Caption = msStringID & msConfig & " " & msOpt
      If Not InStr(1, msOpt, "HIGHRESRATE") = 0 Then
         RateVal = mfRateReturned / 1000 & "Hz"
      Else
         If mfRateReturned < 1000 Then
            RateVal = mfRateReturned & "Hz"
         Else
            RateVal = mfRateReturned / 1000 & "kHz"
         End If
      End If
      SetDetails mnFuncType, mnFirstChan, Chans%, mlAvgVal, _
      RateVal, mnRange, BoardName, mlPTCountReturn
      mnDetailsNeedUpdate = False
      If mbEvalEnabled Then SetRate mfRateReturned
   End If
   If IsMissing(DataAvailable) Then
      NumSamples& = mlCount '* (Chans%)
      CfgMessage$ = "?AISCAN:BUFSIZE"
      MBDResponse$ = MsgLibrary.SendMessage(CfgMessage$)
      If Not SaveMsg(Me, "SendMessage(" & CfgMessage$ & ")", MBDResponse$) Then
         SplitVals = Split(MBDResponse$, "=")
         If UBound(SplitVals) > 0 Then BSize& = Val(SplitVals(1))
         MaxSamps& = (BSize& / 2) / NumChans%
         If mlCount > MaxSamps& Then NumSamples& = MaxSamps&
      End If
      mnBufferEmpty = True
   Else
      If mlEventType = 0 Then
         NumSamples& = DataAvailable
         CfgMessage$ = "?AISCAN:BUFSIZE"
         MBDResponse$ = MsgLibrary.SendMessage(CfgMessage$)
         If Not SaveMsg(Me, "SendMessage(" & CfgMessage$ & ")", MBDResponse$) Then
            SplitVals = Split(MBDResponse$, "=")
            If UBound(SplitVals) > 0 Then BSize& = Val(SplitVals(1))
            MaxSamps& = (BSize& / 2) / NumChans%
            If DataAvailable > MaxSamps& Then NumSamples& = MaxSamps&
         End If
      Else
        'number of samples should be returned by event
        NumSamples& = DataAvailable
      End If
   End If
   BufferExists% = Not (mlHandle = 0)
   TotalSamples& = NumSamples& * NumChans%
   
   If Not mbJustReadBuffer Then
      If (NumSamples& > mlCount) Or (mlHandle = 0) Then
         If Not (mlHandle = 0) Then
            If BufFree(Me, mlHandle, UseWinAPI%) Then
               mlHandle = 0: mnAcqDataType = vbEmpty
               mlAcqPoints = 0
            End If
         End If
         mlHandle = ScaledBufAlloc(Me, TotalSamples&, False)
         mnAcqDataType = vbDouble: mlAcqPoints = TotalSamples&
      End If
      If NumSamples& = 0 Then
         mnBufferEmpty = True
         Exit Function
      End If
   
      ReDim MultiDatArray(Chans%, NumSamples& - 1) As Double
      MsgResult$ = MsgLibrary.ReadInScanMultiChan(MultiDatArray(), NumSamples&, mlReadTimeout)
      ChanDim% = UBound(MultiDatArray, 1)
      If ChanDim% < Chans% Then
         MsgBox "Less data than expected returned. Expected " & _
            Chans% + 1 & " but received " & ChanDim% + 1 & _
            " channels.", vbCritical, "Incorrect Data"
         TransferDataFromMsg = False
         Exit Function
      End If
      For check% = 0 To Chans%
         If MultiDatArray(check%, 0) > 2 ^ 64 Then
            MsgBox "Bad data returned. Element " & check% & ", " & _
            "0 = " & MultiDatArray(check%, 0), vbCritical, "Bad Data"
            TransferDataFromMsg = False
            Exit Function
         End If
      Next
      If MsgResult$ = "No error occurred." Then
         'check for infinity or not a number
         CurArg = MultiDatArray(0, 0)
         If IsNumeric(CurArg) Then
            If CurArg = 0 Then
               CurArgString$ = Format(CurArg, "0")
               If Not Val(CurArgString$) = 0 Then
                  CurArg = "Undefined"
                  'CurArg is infinity or not a number
                  MsgResult$ = "Error- Data contains values that " & _
                    "are not numbers, infinite, or out of range."
               End If
            End If
         End If
      End If
      If InStr(1, MsgResult$, "Error- There are no more input samples available to read") > 0 Then
         SplitMsg = Split(MsgResult$, "Error- ")
         MsgResult$ = SplitMsg(1)
         BadRead% = True
         lblStatus.BackColor = &HF0B0FF
         Me.lblStatus.Caption = Format(NumSamples&, "0") & " requested are not available to read."
      End If
      If SaveMsg(Me, "ReadInScanMultiChan(" & Format(NumSamples&, "0") _
      & ", " & Format(mlReadTimeout, "0") & ")", MsgResult$) Then
         Exit Function
      End If
      
      If BadRead% Then Exit Function
      If UBound(MultiDatArray, 1) < Chans% Then
         MsgBox "Although " & Chans% + 1 & " channels of data were expected, only " & _
         NumChans% + 1 & " channels were returned.", vbInformation, "Incomplete Data"
      End If
      If mlHandle = 0 Then Exit Function
      
      If WDblArrayToBuf(Me, mlHandle, MultiDatArray(), TotalSamples&) Then Exit Function
   End If
   
   If PlotData Then
      ScalingData% = mnuToEng.Checked
      SpecifyMsgLib& = &H100
      '0x40 = FROMMESSAGE, 0x20 = float, 3 = double
      '0x10 = scaled, 0x80 = BUFFER
      If ScalingData% Then
         BufType% = 3 Or &H20 Or &H10
      Else
         BufType% = 3 Or &H20
      End If
      SetBufferType BufType%
      PlotBuffer mlHandle, TotalSamples&, Chans%
   End If

   TransferDataFromMsg = True
      
End Function

Public Sub SetLibType(ByVal LibType As Integer)

   mnuLibrary_Click (LibType)
   
End Sub

Public Sub SetTrigVal(Index As Integer, Optional LThresh As Variant, Optional HThresh As Variant, _
Optional TrigOption As Variant)

   Dim OtherTrigForms As Form
   Dim response As VbMsgBoxResult
   'select the range used to set trigger level
   lpFileName$ = "UniBoards.ini"

   IsAnalog% = (Index < 10) Or ((Index > 13) And (Index < 16))
   If IsAnalog% Then
      'if it's an analog trigger function
      '(trigger types above 7 are digital)
      If mnCustomInfo Then
         'if the ini file exists, use parameters from that
         'otherwise, use the hard-coded info in Else clause below
         'If Not mnIniChecked Then
         If 1 Then
            'if the file exists, get the custom settings for this board
            lpApplicationName$ = msDisplayName
            lpTrigResName$ = "TrigRes"
            lpTrigRangeName$ = "TrigRange"
            lpKeyName$ = "Query"
            'first, get standard entries, then check for custom entries
            lpDefault$ = "Not Listed"
            nSize% = 12
            StdParam$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpTrigResName$, lpDefault$, StdParam$, nSize%, lpFileName$)
            StdParam$ = Left$(StdParam$, StringSize%)
            StdTrigRes& = Val(StdParam$)
            lpDefault$ = "Not Listed"
            nSize% = 12
            StdParam$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpTrigRangeName$, lpDefault$, StdParam$, nSize%, lpFileName$)
            StdParam$ = Left$(StdParam$, StringSize%)
            StdTrigRange& = Val(StdParam$)
            'check for special conditions
            lpDefault$ = ""
            nSize% = 80
            CustomParam$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
            CustomParam$ = Left$(CustomParam$, StringSize%)
            If Len(CustomParam$) Then
               If IsMissing(TrigOption) Then
                  response = MsgBox(CustomParam$, vbYesNo, "Instacal Configuration")
                  If response = vbYes Then
                     lpTrigResName$ = "AltTrigRes"
                     lpTrigRangeName$ = "AltTrigRange"
                  End If
               Else
                  If (TrigOption And 1) = 1 Then
                     lpTrigResName$ = "AltTrigRes"
                     lpTrigRangeName$ = "AltTrigRange"
                  End If
               End If
            End If
            lpDefault$ = Format(StdTrigRes&, "0")
            nSize% = 12
            CustomParam$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpTrigResName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
            CustomParam$ = Left$(CustomParam$, StringSize%)
            mnCustomRes = Val(CustomParam$)
            mlCustomRange = -1
            lpDefault$ = Format(StdTrigRange&, "0")
            nSize% = 8
            CustomParam$ = Space$(nSize%)
            StringSize% = GetPrivateProfileString(lpApplicationName$, lpTrigRangeName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
            CustomParam$ = Left$(CustomParam$, StringSize%)
            mlCustomRange = Val(CustomParam$)
            mnIniChecked = True
         End If
         If mlCustomRange = -2 Then
            TrigRange% = mnRange
         Else
            TrigRange% = mlCustomRange
         End If
         Resolution% = mnResolution
         If Not mnCustomRes = 0 Then
            SpecialRes% = True
            Resolution% = mnCustomRes
         End If
      Else
         If mnFuncType = ATrig Then
            TrigRange% = mnRange
         Else
            'this is dependant on the range available on the DACs
            'for boards with analog trigger capability.
            'as of UL5.02, all use 10V bipolar
            'as of 5.10, some boards use 5V DACs for analog
            'trigger thresholds
         
            Select Case msDisplayName
               Case "PCI-DAS64/M1/16", "PCI-DAS64/M2/16", "PCI-DAS64/M3/16"
                  TrigRange% = BIP5VOLTS
               Case "PCI-DAS4020/12"
                  '4020 uses the actual acquired signal for trigger so the
                  'range is the same as the A/D range
                  TrigRange% = mnRange
               Case Else
                  TrigRange% = BIP10VOLTS
            End Select
         End If

         'as of 5.30, some boards use separate DACs that have resolutions
         'unrelated to the boards output DACs
         Select Case msDisplayName
            Case "PCI-DAS6070", "PCI-DAS6040", "PCI-DAS6071"
               SpecialRes% = True
               Resolution% = 8
               If IsMissing(TrigOption) Then
                  Prompt$ = "Is board configured for first channel in scan as analog trigger input?"
                  If response = 0 Then response = MsgBox(Prompt$, 4, "Instacal Configuration")
                  If response = vbYes Then TrigRange% = mnRange
               Else
                  If (TrigOption And 2) = 2 Then TrigRange% = mnRange
               End If
            Case "PCI-DAS6052", "PCI-DAS6030", "PCI-DAS6032", "PCI-DAS6031", "PCI-DAS6033"
               SpecialRes% = True
               Resolution% = 12
               If IsMissing(TrigOption) Then
                  Prompt$ = "Is board configured for first channel in scan as analog trigger input?"
                  If response = vbCancel Then response = MsgBox(Prompt$, 4, "Instacal Configuration")
                  If response = vbYes Then TrigRange% = mnRange
               Else
                  If (TrigOption And 2) = 2 Then TrigRange% = mnRange
               End If
            Case "PCI-2511", "PCI-2513", "PCI-2515", "PCI-2517", "USB-2523", "USB-2527", "USB-2533", "USB-2537"
               If mnFuncType = APretrig Then
                  SpecialRes% = False
                  Resolution% = 16
               Else
                  SpecialRes% = True
                  Resolution% = 12
               End If
               TrigRange% = mnRange
            Case Else
               UseList% = mnuUseList.Checked
               If UseList% Then
                  BoardName$ = msDisplayName
               Else
                  BoardName$ = Str$(mnBoardNum)
               End If
               Resolution% = GetDAResolution(BoardName$, mnBoardNum, mnRange)
         End Select
      'tentative endif
      End If

      If (Index <> 1) Or (mnFuncType = ATrig) Then
         If SpecialRes% Then
            HighThreshold! = GetVolts(Resolution%, TrigRange%, mnHighThreshold)
         ElseIf Resolution% <> mnResolution Then
            'for boards that have different AD and DA resolutions
            HighThreshold! = GetVolts(mnResolution, mnRange, mnHighThreshold)
         Else
            If mnLibType = UNILIB Then
               ULStat = cbToEngUnits(mnBoardNum, TrigRange%, mnHighThreshold, HighThreshold!)
            End If
         End If
         If Not (Index = GATEBELOW) Then
            If IsMissing(HThresh) Then
               sDefault$ = Format$(HighThreshold!, "0.00")
               HighT = InputBox("Enter high threshold in volts", "Set Analog Trigger Threshold", sDefault$)
            Else
               HighT = HThresh
            End If
         Else
            'trigger is GATEBELOW so high threshold not used
            HighT = "10"
         End If
         If Len(HighT) Then
            HighThreshold! = Val(HighT)
            If SpecialRes% Then
               LongCounts& = GetCounts(Resolution%, TrigRange%, HighThreshold!)
               mnHighThreshold = ULongValToInt(LongCounts&)
            ElseIf Resolution% <> mnResolution Then
               LongCounts& = GetCounts(mnResolution, mnRange, HighThreshold!)
               mnHighThreshold = ULongValToInt(LongCounts&)
            Else
               If mnLibType = UNILIB Then
                  ULStat = cbFromEngUnits(mnBoardNum, TrigRange%, HighThreshold!, mnHighThreshold)
                  If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, TrigRange%, HighThreshold!, mnHighThreshold, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               End If
            End If
         Else
            Exit Sub
         End If
      End If
      If (Index > 0) Or (mnFuncType = ATrig) Then
         If SpecialRes% Then
            LowThreshold! = GetVolts(Resolution%, TrigRange%, mnLowThreshold)
         ElseIf Resolution% <> mnResolution Then
            'for boards that have different AD and DA resolutions
            LowThreshold! = GetVolts(mnResolution, mnRange, mnLowThreshold)
         Else
            If mnLibType = UNILIB Then
               ULStat = cbToEngUnits(mnBoardNum, TrigRange%, mnLowThreshold, LowThreshold!)
            End If
         End If
         If Not (Index = GATEABOVE) Then
            If IsMissing(LThresh) Then
               sDefault$ = Format$(LowThreshold!, "0.00")
               LowT = InputBox("Enter low threshold in volts", "Set Analog Trigger Threshold", sDefault$)
            Else
               LowT = LThresh
            End If
         Else
            'trigger is GATEABOVE so low threshold not used
            LowT = "-10"
         End If
         If Len(LowT) Then
            LowThreshold! = Val(LowT)
            If SpecialRes% Then
               LongCounts& = GetCounts(Resolution%, TrigRange%, LowThreshold!)
               mnLowThreshold = ULongValToInt(LongCounts&)
            ElseIf Resolution% <> mnResolution Then
               LongCounts& = GetCounts(mnResolution, mnRange, LowThreshold!)
               mnLowThreshold = ULongValToInt(LongCounts&)
            Else
               If mnLibType = UNILIB Then
                  ULStat = cbFromEngUnits(mnBoardNum, TrigRange%, LowThreshold!, mnLowThreshold)
                  If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, TrigRange%, LowThreshold!, mnLowThreshold, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               End If
            End If
         Else
            Exit Sub
         End If
      End If
   End If
      
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


   If (mnFuncType = ATrig) And Not (Index < 2) Then
      If Index > 1 Then ULStat = BADTRIGTYPE
   ElseIf Not (mnFuncType = ATrig) Then
      If mnLibType = UNILIB Then
         ULStat = cbSetTrigger(mnBoardNum, Index, mnLowThreshold, mnHighThreshold)
         x% = SaveFunc(Me, SetTrigger, ULStat, mnBoardNum, Index, mnLowThreshold, mnHighThreshold, A5, A6, A7, A8, A9, A10, A11, 0)
      Else
         Prop$ = "AITRIG:TYPE="
         TrigString$ = GetMsgTrigTypeString(Index)
         TrigTypeMsg$ = Prop$ & TrigString$
         MsgResult$ = MsgLibrary.SendMessage(TrigTypeMsg$)
         If SaveMsg(Me, "SendMessage(" & TrigTypeMsg$ & ")", MsgResult$) Then ULStat = BADTRIGTYPE
      End If
      mbThresholdSet = True
   End If

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
      mnuTrigType(mnTrigType).Checked = False
      mnTrigSet = True
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
      AmpMax! = GetRangeVolts(mnRange) '/ 2
      If mnRange < 100 Then
         If AsOffset Then
            NegFS! = AmpMax! * -0.5
            CalcOffset! = (AmpMax! * FSFactor!) '/ 2
            CalcAmplitude! = NegFS! + CalcOffset!
         Else
            CalcAmplitude! = (AmpMax! * FSFactor!) / 2
         End If
      Else
         CalcAmplitude! = AmpMax! * FSFactor!
      End If
   Else
      If Not InStr(DataString, ".") = 0 Then
         CalcAmplitude! = Val(DataString)
      Else
         LongCounts& = Val(DataString)
         GetCountsFromUnits = LongCounts&
         Exit Function
      End If
   End If
   LongCounts& = GetHiResCounts(mnResolution, mnRange, CalcAmplitude!)
   
   GetCountsFromUnits = LongCounts&

End Function

Public Function PlotOnInterval(ByVal CountValue As Long, ByVal TimeLimit As Long) As Long

   ULStat = 0
   StartTime! = Timer()
   Do
      If mnLibType = MSGLIB Then
         Status% = MsgCheckStatus(Me, msBoardName, _
         MsgLibrary, msStatusType, CurCount&, CurIndex&)
         If Status% = OVERRUN Then ULStat = OVERRUN
      Else
         ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
         x% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, _
         mlStatusType, A6, A7, A8, A9, A10, A11, 0)
      End If
      StatReport$ = Str(ULStat) & "," & Str(Status%) & "," & Str(CurCount&) & "," & _
      Str(CurIndex&)
      Me.cmdConfigure.Caption = StatReport$
      If (Status% = RUNNING) Then
         BGOp$ = "Background task in progress..."
         ScanActive% = (ULStat = 0)
      Else
         BGOp$ = "No background task running..."
         ScanActive% = False
      End If
      Stat$ = BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
      lblStatus.Caption = Stat$
      ET! = Timer() - StartTime!
      TimedOut% = (ET! > TimeLimit)
      If ScanActive% Then ScanActive% = (ET! < TimeLimit)
      If (ET! - Iteration!) > 0.5 Then
         frmScript.PostTimeoutStatus TimeLimit - ET!
         Iteration! = ET!
      End If
      If ScanActive% And ((CurCount& - PreviousCount&) > CountValue) Then
         PreviousCount& = CurCount&
         ShowBufferInfo 0
      End If
      If gnScriptPaused Then ScanActive% = False
      DoEvents
   Loop While ScanActive%
   PlotOnInterval = TimedOut%

End Function

Public Function StopOnCount(ByVal CountValue As Long, ByVal Timeout As Long) As Long

   ULStat = 0
   StartTime! = Timer()
   Do
      If mnLibType = MSGLIB Then
         Status% = MsgCheckStatus(Me, msBoardName, _
         MsgLibrary, msStatusType, CurCount&, CurIndex&)
         If Status% = OVERRUN Then ULStat = OVERRUN
      Else
         ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
         x% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, CurIndex&, _
         mlStatusType, A6, A7, A8, A9, A10, A11, 0)
      End If
      StatReport$ = Str(ULStat) & "," & Str(Status%) & "," & Str(CurCount&) & "," & _
      Str(CurIndex&)
      Me.cmdConfigure.Caption = StatReport$
      If (Status% = RUNNING) Then
         BGOp$ = "Background task in progress..."
         ScanActive% = (CurCount& < CountValue) And (ULStat = 0)
      Else
         BGOp$ = "No background task running..."
         ScanActive% = False
      End If
      Stat$ = BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
      lblStatus.Caption = Stat$
      ET! = Timer() - StartTime!
      TimedOut% = (ET! > Timeout)
      If ScanActive% Then ScanActive% = (Timeout > ET!)
      If (ET! - Iteration!) > 0.5 Then
         frmScript.PostTimeoutStatus Timeout - ET!
         Iteration! = ET!
      End If
      DoEvents
      If gnScriptPaused Then
         ScanActive% = False
         tmrCheckStatus.ENABLED = False
      End If
   Loop While ScanActive%
   StopOnCount = TimedOut%
   
   If Not gnScriptPaused Then
      Select Case mnLibType
         Case UNILIB
            ULStat = StopBackground520(mnBoardNum, mlStatusType)
            If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnFormType, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
         Case MSGLIB
            If Not (msScanSupport = "") Then
               StopMsg$ = msStatusType & ":STOP"
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
               If MsgResult$ = "AISCAN:STATUS=INTERRUPTED" Then
                  ULStat = INTERRUPTED
                  MsgQualifier$ = "Error: "
                  MsgResult$ = MsgQualifier$ & MsgResult$
               End If
               If SaveMsg(Me, "SendMessage(" & StopMsg$ & ")", MsgResult$) Then Failure% = True
               mnEventTimeout = False
            End If
      End Select
   End If
   
End Function

Public Function ScriptGenData(ByVal DataType&, ByVal Cycles%, ByVal TotalCount&, _
ByVal NumChans%, ByVal Amplitude, ByVal Offset, ByVal SigType%, ByVal NewData%, _
ByVal Channel%, ByVal FirstPoint&) As Long

   SampleMultiplier% = 1
   If mnLibType = MSGLIB Then SampleMultiplier% = NumChans%
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
   DataHandle& = GenerateData(DataType&, Cycles%, TotalSamples&, NumChans%, _
   Amplitude, Offset, SigType%, NewData%, Channel%, FirstPoint&)
   
   SetupData DataHandle&, Samples&, NumChans%, DataType&
   PlotGenData
   ScriptGenData = DataHandle&
   
End Function

Public Sub SetChanInputMode(ByVal Chan As Long, ByVal InputMode As Long)

   If mnLibType = MSGLIB Then
      MsgBox "This method has not been implemented for DAQFlex", _
         vbCritical, "Not Implemented for DAQFlex"
   Else
      ULStat& = cbAChanInputMode(mnBoardNum, Chan, InputMode)
      x% = SaveFunc(Me, AChanInputMode, ULStat, mnBoardNum, Chan, _
         InputMode, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      End If
   
End Sub

Public Sub SetInputMode(ByVal SetSE As Boolean)

   mnuSetSEMode.Checked = SetSE
   SEMode& = Abs(mnuSetSEMode.Checked)
   If mnLibType = MSGLIB Then
      SetMode$ = "DIFF"
      If SetSE Then SetMode$ = "SE"
      ModeMsg$ = "AI:CHMODE=" & SetMode$
      MsgResult$ = MsgLibrary.SendMessage(ModeMsg$)
      Error% = SaveMsg(Me, "SendMessage(" & ModeMsg$ & ")", MsgResult$)
   Else
      ULStat& = cbAInputMode(mnBoardNum, SEMode&)
      If SaveFunc(Me, AInputMode, ULStat, mnBoardNum, SEMode&, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
         mnuSetSEMode.Checked = False
      End If
   End If
   
End Sub

Private Sub ConfigureValidOptions()

   Dim OptionFound As Boolean
   
   If mnFormType = ANALOG_OUT Then
      ConfigItem& = BIDACSCANOPTIONS
      If mnIOTFunc Then ConfigItem& = BIDAQOSCANOPTIONS
   Else
      ConfigItem& = BIADSCANOPTIONS
      If mnIOTFunc Then ConfigItem& = BIDAQISCANOPTIONS
   End If
   FormType% = mnFormType
   If (InStr(1, msBoardName, "CTR") > 1) Then FormType% = COUNTERS
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
      OptString$ = GetOptionsString(ConfigVal&, FormType%, mnIOTFunc)
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

Public Function GetTCInfo(ByVal InfoValue As Variant) As String

   Select Case InfoValue
      Case -11111
         InfoString$ = " Disabled "
      Case -9999
         InfoString$ = "  Open    "
      Case -9000
         InfoString$ = "Not Ready "
      Case -7777
         InfoString$ = "CMV Error "
      Case -6666
         InfoString$ = "Over Range"
   End Select
   GetTCInfo = InfoString$
   
End Function
