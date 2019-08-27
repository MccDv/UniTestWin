VERSION 5.00
Begin VB.Form frmConfiguration 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Board Configuration"
   ClientHeight    =   5460
   ClientLeft      =   2760
   ClientTop       =   1740
   ClientWidth     =   6045
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   6045
   Begin VB.ComboBox cmbList 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   90
      Style           =   1  'Simple Combo
      TabIndex        =   48
      Text            =   "Combo1"
      Top             =   3690
      Visible         =   0   'False
      Width           =   5820
   End
   Begin VB.Frame fraSetPoint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Setpoint Setup"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   28
      Top             =   1620
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cmbSPFlags 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1020
         Width           =   3135
      End
      Begin VB.ComboBox cmbSPOutput 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1380
         Width           =   2595
      End
      Begin VB.TextBox txtLimitA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         ToolTipText     =   "Limit A"
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtLimitB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "Limit B"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtOut1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         ToolTipText     =   "Output 1"
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtOut2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Output 2"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtMask1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Mask 1"
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtMask2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         ToolTipText     =   "Mask 2"
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton cmdLoadArray 
         Appearance      =   0  'Flat
         Caption         =   "Load"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4740
         TabIndex        =   17
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdDone 
         Appearance      =   0  'Flat
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   336
         Left            =   4740
         TabIndex        =   18
         Top             =   1080
         Width           =   915
      End
      Begin VB.ListBox lstElement 
         Appearance      =   0  'Flat
         Height          =   1005
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
         Width           =   855
      End
      Begin VB.TextBox txtQCount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   29
         Text            =   "0"
         Top             =   300
         Width           =   435
      End
      Begin VB.CheckBox chkLatch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Latch"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1380
         Width           =   1035
      End
   End
   Begin VB.Frame fraDaqTrig 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Daq Trigger Configuration"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   3900
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtDTVar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   25
         Text            =   "0.0"
         Top             =   1020
         Width           =   795
      End
      Begin VB.TextBox txtDTLevel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Text            =   "0.0"
         Top             =   660
         Width           =   795
      End
      Begin VB.OptionButton optDTEvent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop Event"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   27
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optDTEvent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Start Event"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4140
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox cmbDTSource 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Text            =   "TrigSource"
         Top             =   300
         Width           =   2115
      End
      Begin VB.ComboBox cmbDTSense 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "cfg.frx":0000
         Left            =   180
         List            =   "cfg.frx":0002
         TabIndex        =   21
         Text            =   "TrigSense"
         Top             =   660
         Width           =   2115
      End
      Begin VB.ComboBox cmbDTType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Text            =   "ChanType"
         Top             =   1020
         Width           =   2115
      End
      Begin VB.CommandButton cmdDTTrig 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Height          =   375
         Left            =   4800
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDTChan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Text            =   "0"
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lblDTVariance 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Variance"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblDTLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblDTChan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Channel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2940
         TabIndex        =   33
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame fraATCC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Advanced Timing && Control"
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   120
      TabIndex        =   34
      Top             =   1980
      Width           =   5775
      Begin VB.OptionButton optSelIO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Disabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4500
         TabIndex        =   35
         Top             =   660
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox lstInvert 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   60
         TabIndex        =   36
         Top             =   300
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkInvert 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Invert"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4500
         TabIndex        =   15
         Top             =   960
         Width           =   1155
      End
      Begin VB.OptionButton optSelIO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Outputs"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4500
         TabIndex        =   37
         Top             =   420
         Width           =   1095
      End
      Begin VB.OptionButton optSelIO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inputs"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4500
         TabIndex        =   38
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdATCC 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Height          =   375
         Left            =   4740
         TabIndex        =   39
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox lstSignals 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   2460
         MultiSelect     =   2  'Extended
         TabIndex        =   40
         Top             =   300
         Width           =   1935
      End
      Begin VB.ListBox lstATCCIn 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   600
         MultiSelect     =   2  'Extended
         TabIndex        =   41
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Frame fraConfiguration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Board Configuration"
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox chkFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Legacy"
         ForeColor       =   &H80000008&
         Height          =   195
         HelpContextID   =   4
         Index           =   5
         Left            =   4560
         TabIndex        =   55
         Top             =   1260
         Width           =   1155
      End
      Begin VB.CheckBox chkFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Std"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   1260
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Temp"
         ForeColor       =   &H80000008&
         Height          =   195
         HelpContextID   =   4
         Index           =   4
         Left            =   3540
         TabIndex        =   53
         Top             =   1260
         Width           =   915
      End
      Begin VB.CheckBox chkFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DCtr"
         ForeColor       =   &H80000008&
         Height          =   195
         HelpContextID   =   3
         Index           =   3
         Left            =   2640
         TabIndex        =   52
         Top             =   1260
         Width           =   855
      End
      Begin VB.CheckBox chkFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DA"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   51
         Top             =   1260
         Width           =   735
      End
      Begin VB.CheckBox chkFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AD"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   50
         Top             =   1260
         Width           =   675
      End
      Begin VB.CheckBox chkReQuery 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ReQuery"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   49
         Top             =   480
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cmbDevNum2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3690
         TabIndex        =   47
         Text            =   "DevNum"
         ToolTipText     =   "DevNum"
         Top             =   840
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdConfigure 
         Appearance      =   0  'Flat
         Height          =   195
         Left            =   3720
         TabIndex        =   43
         Top             =   465
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkHex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hex"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   44
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtConfigVal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   45
         ToolTipText     =   "Press 'Insert' for multiline."
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Height          =   375
         Left            =   4800
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbDevNum 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "cfg.frx":0004
         Left            =   2655
         List            =   "cfg.frx":000B
         TabIndex        =   4
         ToolTipText     =   "DevNum"
         Top             =   855
         Width           =   870
      End
      Begin VB.ComboBox cmbConfigItem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Text            =   "ConfigItem"
         ToolTipText     =   "ConfigItem"
         Top             =   855
         Width           =   2310
      End
      Begin VB.ComboBox cmbInfoType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Text            =   "InfoType"
         ToolTipText     =   "InfoType"
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lblConfigVal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ConfigVal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   5715
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
      Begin VB.Menu mnuSep2 
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
         Caption         =   "cbGetConfig()"
         Index           =   0
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbSetConfig()"
         Index           =   1
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbGetSignal()"
         Index           =   2
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbSelectSignal()"
         Index           =   3
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDaqSetTrigger()"
         Index           =   4
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbGetConfigString()"
         Index           =   5
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbSetConfigString()"
         Index           =   6
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDaqDeviceVersion()"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDaqSetSetpoints()"
         Index           =   8
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadConfig 
         Caption         =   "cbLoadConfig()"
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
      Begin VB.Menu mnuSep5 
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
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the following constants added here until
'they become part of the header
'Const BIDACRANGE = 114
'Const BIDISOFILTER = 122

Const BITCCHANTYPE = 169
Const BIFWVERSION = 170
Const BIAIWAVETYPE = 202
Const BISERIALNUM = 214
Const BIRELAYLOGIC = 228
Const BIOPENRELAYLEVEL = 229
'Const BIADTRIGCOUNT = 219
'Const BIADFIFOSIZE = 220
Const BIMFGSERIALNUM = 224
'Const BIDACTRIGCOUNT = 284

Const BIHASEXTINFO = 309                 '/* Indicates if devices has extended info */
Const BINUMIODEVS = 310                  '/* Number of IO devices */
Const BIIODEVTYPE = 311                  '/* Type of IO device */
'Const BIADNUMCHANMODES = 312             '/* Number of channel modes */
'Const BIADCHANMODE = 313                 '/* Channel mode */
'Const BIADNUMDIFFRANGES = 314            '/* Number of differncial ranges supported by devide */
'Const BIADDIFFRANGE = 315
'Const BIADNUMSERANGES = 316              '/* Number of Single-Ended ranges supported by devide */
'Const BIADSERANGE = 317
Const BIADNUMTRIGTYPES = 318
Const BIADTRIGTYPE = 319
Const BIADMAXRATE = 320
Const BIADMAXTHROUGHPUT = 321
Const BIADMAXBURSTRATE = 322
Const BIADMAXBURSTTHROUGHPUT = 323
Const BIADHASPACER = 324
Const BIADCHANTYPES = 325
'Const BIADSCANOPTIONS = 326
Const BIADMAXSEQUEUELENGTH = 327
Const BIADMAXDIFFQUEUELENGTH = 328
Const BIADQUEUETYPES = 329
Const BIADQUEUELIMITS = 330

Const BIDACHASPACER = 331
'Const BIDACSCANOPTIONS = 332
Const BIDACFIFOSIZE = 333
'Const BIDACNUMRANGES = 334               '/* Number of ranges supported by dac devide */
'Const BIDACDEVRANGE = 335               '// BIDACRANGE is already defined
Const BIDACNUMTRIGTYPES = 336
Const BIDACTRIGTYPE = 337
Const BIDAQAMISUPPORTED = 338

Const BIDISCONNECT = 340
'Const BINETCONNECTCODE = 341
'Const BIDEVVERSION = 359
Const BINETBIOSNAME = 366
'Const BIADAIMODE = 373
Const BIDEVIPADDR = 374
'Const BIADCHANAIMODE = 249

Const BIDAQINUMCHANTYPES = 376
Const BIDAQICHANTYPE = 377
Const BIDAQONUMCHANTYPES = 378
Const BIDAQOCHANTYPE = 379
Const BICTRZACTIVEMODE = 380

Const DEMO_SRC_NONE = 0
Const DEMO_SRC_SINE = 1, DEMO_SRC_SQUARE = 2
Const DEMO_SRC_SAWTOOTH = 3, DEMO_SRC_RAMP = 4
Const DEMO_SRC_DAMPED_SINE = 5, DEMO_SRC_FILE = 1000
'Const TYPEB = 7, TYPEN = 8

Const GET_CONF = 0
Const SET_CONF = 1
Const GET_SIGNAL = 2
Const SELECT_SIGNAL = 3
Const DAQ_TRIG = 4
Const GET_STRING = 5
Const SET_STRING = 6
'Const GET_DEVVER = 7
Const SET_SETPOINTS = 8
Const LOAD_CONF = 16

Const LISTSTD = 1
Const LISTAD = 2
Const LISTDA = 4
Const LISTICAL = 8
Const LISTIDCTR = 8
Const LISTTC = 16
Const LISTLEGACY = 32

#If MSGOPS Then
   Dim WithEvents MsgLibrary As MBDClass.MBDComClass
Attribute MsgLibrary.VB_VarHelpID = -1
#Else
   Dim MsgLibrary As Object
#End If

#If NETOPS Then
   'Dim ULNetLibrary As AcqThread.AcqThread
#Else
   Dim ULNetLibrary As Object
#End If

Dim mnThisInstance As Integer

Dim mnFormType As Integer, msTitle As String

Dim mnFuncType As Integer, mnPlotType As Integer

Dim mnPlot As Integer, mnLoaded As Integer

Dim gsConfig As String

Dim mnBoardIndex As Integer
Dim mnBoardNum As Integer

Dim mnADChans As Integer, mnDigDevs As Integer
Dim mnCtrDevs As Integer, mnDAChans As Integer
Dim mnIOPorts As Integer
Dim mlDirection As Long, mlPolarity As Long
Dim mnCurSelect As Integer

Dim mnRange As Integer
Dim mlQCount As Long
Dim mafLimAArray() As Single
Dim mafLimBArray() As Single
Dim mafOut1Array() As Single
Dim mafOut2Array() As Single
Dim mafMask1Array() As Single
Dim mafMask2Array() As Single
Dim malSPFlags() As Long, malOutputs() As Long
Dim msBoardName As String

Dim mnMessaging As Integer, mnLibType As Integer
Dim mnNumBoards As Integer, mnFormInitialized As Integer
Dim mnUnloading As Integer, mnLoading As Integer
Dim msDisplayName As String, msAiSupport As String
Dim mnTempSupport As Integer, mnScale As Integer

Dim mnDevList1 As Integer, mnDevList2 As Integer
Dim mbHoldoffUpdate As Boolean

Private Sub chkFilter_Click(Index As Integer)

   If Not mbHoldoffUpdate Then LoadULItems
   
End Sub

Private Sub chkHex_Click()

   If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
   If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus

End Sub

Private Sub chkInvert_Click()

   If chkInvert.value = 1 Then
      mlPolarity = INVERTED
   Else
      mlPolarity = NONINVERTED
   End If

End Sub

Private Sub cmbConfigItem_Click()

   ConfItemString$ = cmbConfigItem.Text
   SetDefaultCfgValue
   If cmbInfoType.ENABLED Then  'And (gnNumBoards > 0)
      If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
      If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus
   End If

End Sub

Private Sub cmbConfigItem_KeyDown(KeyCode As Integer, Shift As Integer)

   If (KeyCode = 13) Then

   ConfItemString$ = Me.cmbConfigItem.Text
      If IsNumeric(ConfItemString$) Then
         CfgVal& = Val(ConfItemString$)
         ttText$ = GetCfgItemStrFromVal(CfgVal&)
         cmbConfigItem.ToolTipText = ttText$
      End If
      SetDefaultCfgValue
      If ((mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING)) Then UpdateStatus
      If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus
   End If

End Sub

Sub SetDefaultCfgValue()

   Dim CurUpdateStatus As Boolean
   
   CurUpdateStatus = mbHoldoffUpdate
   mbHoldoffUpdate = True
   ConfItemString$ = cmbConfigItem.Text
   cmbDevNum.Clear
   cmbDevNum2.Clear
   'cmbDevNum.Visible = False
   'cmbDevNum2.Visible = False
   cmbDevNum.AddItem "0"
   cmbDevNum.ListIndex = 0

   Select Case ConfItemString$
      Case "AInputMode", "BIADAIMODE"
         txtConfigVal.Text = "1"
      Case "AChanInputMode", "BIADCHANAIMODE"
         txtConfigVal.Text = "1"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For Chan% = 0 To 7
            cmbDevNum.AddItem Format(Chan%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "BIAIWAVETYPE"
         txtConfigVal.Text = "5"
      Case "BIINPUTPACEROUT"
         txtConfigVal.Text = "1"
      Case "BIOUTPUTPACEROUT"
         txtConfigVal.Text = "1"
      Case "BIEXTINPACEREDGE", "BIEXTOUTPACEREDGE"
         txtConfigVal.Text = "2"
      Case "BIADCSETTLETIME"
         txtConfigVal.Text = "4"
      Case "BICHANTCTYPE"
         txtConfigVal.Text = "1"
      Case "BIDACFORCESENSE"
         txtConfigVal.Text = "1"
      Case "BIDISOFILTER"
         txtConfigVal.Text = "1"
      Case "BISYNCMODE"
         txtConfigVal.Text = "1"
      Case "BIDEVVERSION"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         cmbDevNum.AddItem "0"
         cmbDevNum.AddItem "1"
         cmbDevNum.AddItem "2"
         cmbDevNum.AddItem "3"
         cmbDevNum.AddItem "4"
         cmbDevNum.ListIndex = 0
      Case "BIDIALARMMASK"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         cmbDevNum.AddItem "0"
         cmbDevNum.AddItem "1"
         cmbDevNum.AddItem "2"
         cmbDevNum.AddItem "3"
         cmbDevNum.AddItem "4"
         cmbDevNum.ListIndex = 0
      Case "BIDISCONNECT"
         txtConfigVal.Text = "54211"
      Case "BIEXTCLKTYPE"
         txtConfigVal.Text = "2"
      Case "BINETCONNECTCODE"
         txtConfigVal.Text = "0"
      Case "BIDITRIGCOUNT"
         txtConfigVal.Text = "100"
      Case "BIDOTRIGCOUNT"
         txtConfigVal.Text = "100"
      Case "BIPATTERNTRIGPORT"
         txtConfigVal.Text = "1"
      Case "BITEMPREJFREQ"
         txtConfigVal.Text = "60"
      Case "BUFSIZE"
         txtConfigVal.Text = "4096"
      Case "BUFOVERWRITE"
         txtConfigVal.Text = "ENABLE"
      Case "BURSTMODE"
         txtConfigVal.Text = "ENABLE"
      Case "CAL"
         txtConfigVal.Text = "ENABLE"
      Case "CHAN"
         txtConfigVal.Text = "0"
      Case "CHMODE"
         txtConfigVal.Text = "DIFF"
      Case "CICTRNUM"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         NumCtrs% = mnCtrDevs
         If mnCtrDevs < 1 Then NumCtrs% = 3
         For Dev% = 0 To NumCtrs% - 1
            ItemStr$ = Format(Dev%, "0")
            cmbDevNum.AddItem ItemStr$
         Next
         cmbDevNum.ListIndex = 0
      Case "CICTRTYPE"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         NumCtrs% = mnCtrDevs
         If mnCtrDevs < 1 Then NumCtrs% = 3
         For Dev% = 0 To NumCtrs% - 1
            ItemStr$ = Format(Dev%, "0")
            cmbDevNum.AddItem ItemStr$
         Next
         cmbDevNum.ListIndex = 0
      Case "DATARATE"
         txtConfigVal.Text = "3750"
      Case "DATATYPE"
         txtConfigVal.Text = "ENABLE"
      Case "DEBUG"
         txtConfigVal.Text = "ENABLE"
      Case "DELAY"
         txtConfigVal.Text = "500"
      Case "DIDEVTYPE", "DICONFIG", "DINUMBITS"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         NumDio% = mnDigDevs
         If mnDigDevs < 1 Then NumDio% = 3
         For Dev% = 0 To NumDio% - 1
            ItemStr$ = Format(Dev%, "0")
            cmbDevNum.AddItem ItemStr$
         Next
         cmbDevNum.ListIndex = 0
      Case "DICURVAL", "DIINMASK", "DIOUTMASK"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         NumDio% = mnDigDevs
         If mnDigDevs < 1 Then NumDio% = 3
         For Dev% = 0 To NumDio% - 1
            ItemStr$ = Format(Dev%, "0")
            cmbDevNum.AddItem ItemStr$
         Next
         cmbDevNum.ListIndex = 0
      Case "DIR"
         txtConfigVal.Text = "OUT"
      Case "DIDISABLEDIRCHECK"
         txtConfigVal.Text = "1"
      Case "DUTYCYCLE"
         txtConfigVal.Text = "20"
      Case "EXTPACER"
         txtConfigVal.Text = "ENABLE"
      Case "FLASHLED"
         txtConfigVal.Text = "3"
      Case "GIINIT"
         txtConfigVal.Text = "1"
      Case "HIGHCHAN"
         txtConfigVal.Text = "0"
      Case "LATCH"
         txtConfigVal.Text = "0"
      Case "LOWCHAN"
         txtConfigVal.Text = "0"
      Case "IDLESTATE"
         txtConfigVal.Text = "HIGH"
      Case "PERIOD"
         txtConfigVal.Text = "0.5"
      Case "PULSECOUNT"
         txtConfigVal.Text = "100"
      Case "QUEUE"
         txtConfigVal.Text = "ENABLE"
      Case "RANGE"
         txtConfigVal.Text = "BIP10V"
      Case "RANGE{ch}"
         txtConfigVal.Text = "BIP10V"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "RANGE{el/ch}"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
         cmbDevNum2.Visible = True
         cmbDevNum2.Clear
         For i% = 0 To 15
            cmbDevNum2.AddItem Format(i%, "0")
         Next
         cmbDevNum2.ListIndex = 0
         txtConfigVal.Text = "BIP10V"
      Case "RATE"
         txtConfigVal.Text = "1000"
      Case "REARM"
         txtConfigVal.Text = "ENABLE"
      Case "REG"
         If mnLibType = MSGLIB Then
            Resolution% = GetMsgDAResolution(MsgLibrary)
         Else
            Resolution% = GetDAResolution(msDisplayName, mnBoardNum, mnRange)
         End If
         DefaultVal& = (2 ^ Resolution%) / 2
         txtConfigVal.Text = Format(DefaultVal&, "0")
      Case "RESET"
         If Me.cmbInfoType.Text = "DEV" Then
            txtConfigVal.Text = "DEFAULT"
         Else
            txtConfigVal.Text = ""
         End If
      Case "SAMPLES"
         txtConfigVal.Text = "1000"
      Case "SAVESTATE"
         txtConfigVal.Text = "ENABLE/TRUE"
      Case "SCALE"
         txtConfigVal.Text = "ENABLE"
      Case "SENSOR"
         txtConfigVal.Text = "TC/J"
      Case "SRC"
         txtConfigVal.Text = "HWSTART/DIG"
      Case "STALL"
         txtConfigVal.Text = "ENABLE"
      Case "TEMP{ch}"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "TRIG"
         txtConfigVal.Text = "ENABLE"
      Case "TYPE"
         txtConfigVal.Text = "EDGE/FALLING"
      Case "VALUE"
         txtConfigVal.Text = "0"
      Case "XFRMODE"
         txtConfigVal.Text = "SINGLEIO"
      Case Else
         txtConfigVal.Text = ""
   End Select
   InfoString$ = Me.cmbInfoType.Text
   Select Case InfoString$
      Case "DIO{port/bit}"
         cmbDevNum2.Visible = True
         For i% = 0 To 15
            cmbDevNum2.AddItem Format(i%, "0")
         Next
         cmbDevNum2.ListIndex = 0
   End Select
   mbHoldoffUpdate = CurUpdateStatus
   
End Sub

Private Sub cmbDevNum_Change()

   If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
   If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus

End Sub

Private Sub cmbDevNum_Click()

   If cmbDevNum.ENABLED Then
      If cmbInfoType.ENABLED Then
         If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
         If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus
      End If
   End If
   
End Sub

Private Sub cmbDevNum2_Change()

   If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
   If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus

End Sub

Private Sub cmbDevNum2_Click()

   If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
   If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus

End Sub

Private Sub cmbInfoType_Click()

   ConfigureControls

End Sub

Private Sub cmbInfoType_KeyDown(KeyCode As Integer, Shift As Integer)

   If (KeyCode = 13) Then
      If ((mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING)) Then UpdateStatus
      If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus
   End If

End Sub

Private Sub cmdATCC_Click()

   For SignalItem% = 0 To lstSignals.ListCount - 1
      If lstSignals.Selected(SignalItem%) Then
         If mlDirection = SIGNAL_IN Then
            SelectedSignals& = SelectedSignals& + Choose(SignalItem% + 1, ADC_CONVERT, ADC_GATE, ADC_START_TRIG, ADC_STOP_TRIG, ADC_TB_SRC, DAC_UPDATE, DAC_TB_SRC, DAC_START_TRIG, SYNC_CLK)
         ElseIf mlDirection = CBDISABLED Then
            SelectedSignals& = SelectedSignals& + Choose(SignalItem% + 1, ADC_TB_SRC, DAC_TB_SRC, SYNC_CLK)
         Else
            SelectedSignals& = SelectedSignals& + Choose(SignalItem% + 1, ADC_CONVERT, ADC_START_TRIG, ADC_STOP_TRIG, ADC_SCANCLK, ADC_SSH, ADC_STARTSCAN, ADC_SCAN_STOP, DAC_UPDATE, DAC_START_TRIG, SYNC_CLK, CTR1_CLK, CTR2_CLK, DGND)
         End If
      End If
   Next SignalItem%
   
   If mnFuncType = SELECT_SIGNAL Then
      For IOItem% = 0 To lstATCCIn.ListCount - 1
         If lstATCCIn.Selected(IOItem%) Then
            If (SelectedIO& > 0) And (Not MultiIO_OK%) Then
               response& = MsgBox("Multiple I/O pins are connected to one signal.  Continue?", 4, "Multiple I/O Warning")
               If response& = 6 Then
                  MultiIO_OK% = True
               Else
                  Exit For
               End If
            End If
            If mlDirection = SIGNAL_IN Then
               SelectedIO& = SelectedIO& + Choose(IOItem% + 1, AUXIN0, AUXIN1, AUXIN2, AUXIN3, AUXIN4, AUXIN5, DS_CONNECTOR)
            ElseIf mlDirection = CBDISABLED Then
               'SelectedIO& = SelectedIO& + Choose(IOItem% + 1, AUXOUT0, AUXOUT1, AUXOUT2, DS_CONNECTOR)
               'no connection options
            Else
               SelectedIO& = SelectedIO& + Choose(IOItem% + 1, AUXOUT0, AUXOUT1, AUXOUT2, DS_CONNECTOR)
            End If
         End If
      Next IOItem%
      ULStat = SelectSignal524(mnBoardNum, mlDirection, SelectedSignals&, SelectedIO&, mlPolarity)
      If SaveFunc(Me, SelectSignal, ULStat, mnBoardNum, mlDirection, SelectedSignals&, SelectedIO&, mlPolarity, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   Else
      For IOItem% = 0 To lstATCCIn.ListCount - 1
         lstATCCIn.Selected(IOItem%) = False
      Next IOItem%
      Do
         ULStat = GetSignal524(mnBoardNum, mlDirection, SelectedSignals&, Index&, SelectedIO&, Polarity&)
         If ULStat = BADINDEX Then Exit Do
         If SaveFunc(Me, GetSignal, ULStat, mnBoardNum, mlDirection, SelectedSignals&, Index&, SelectedIO&, Polarity&, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Index& = Index& + 1
         If mlDirection = SIGNAL_IN Then
            Select Case SelectedIO&
               Case AUXIN0
                  lstATCCIn.Selected(0) = True
                  lstInvert.RemoveItem (0)
                  lstInvert.AddItem Str$(Polarity&), 0
               Case AUXIN1
                  lstATCCIn.Selected(1) = True
                  lstInvert.RemoveItem (1)
                  lstInvert.AddItem Str$(Polarity&), 1
               Case AUXIN2
                  lstATCCIn.Selected(2) = True
                  lstInvert.RemoveItem (2)
                  lstInvert.AddItem Str$(Polarity&), 2
               Case AUXIN3
                  lstATCCIn.Selected(3) = True
                  lstInvert.RemoveItem (3)
                  lstInvert.AddItem Str$(Polarity&), 3
               Case AUXIN4
                  lstATCCIn.Selected(4) = True
                  lstInvert.RemoveItem (4)
                  lstInvert.AddItem Str$(Polarity&), 4
               Case AUXIN5
                  lstATCCIn.Selected(5) = True
                  lstInvert.RemoveItem (5)
                  lstInvert.AddItem Str$(Polarity&), 5
               Case DS_CONNECTOR
                  lstATCCIn.Selected(6) = True
                  lstInvert.RemoveItem (6)
                  lstInvert.AddItem Str$(Polarity&), 6
            End Select
         Else
            Select Case SelectedIO&
               Case AUXOUT0
                  lstATCCIn.Selected(0) = True
                  lstInvert.RemoveItem (0)
                  lstInvert.AddItem Str$(Polarity&), 0
               Case AUXOUT1
                  lstATCCIn.Selected(1) = True
                  lstInvert.RemoveItem (1)
                  lstInvert.AddItem Str$(Polarity&), 1
               Case AUXOUT2
                  lstATCCIn.Selected(2) = True
                  lstInvert.RemoveItem (2)
                  lstInvert.AddItem Str$(Polarity&), 2
               Case DS_CONNECTOR
                  lstATCCIn.Selected(3) = True
                  lstInvert.RemoveItem (3)
                  lstInvert.AddItem Str$(Polarity&), 3
            End Select
         End If
      Loop While (SelectedIO& > 0) And (mlDirection = SIGNAL_OUT)
   End If

End Sub

Private Sub cmdConfigure_Click()

   'this exists to give menu access to the scripting
   'form when running scripts
   CmdStr$ = Left$(cmdConfigure.Caption, 1)
   value& = Val(Mid$(cmdConfigure.Caption, 2))
   Select Case CmdStr$
      Case "$" 'return board name
         cmdConfigure.Caption = msBoardName
      Case "8"
         CurOption$ = ""
         Me.cmdConfigure.Caption = CurOption$
      Case "B" 'set board number
         SearchName$ = Mid$(cmdConfigure.Caption, 2)
         BoardParams = Split(SearchName$, ",")
         If UBound(BoardParams) = 1 Then DupeIndex% = Val(BoardParams(1))
         For MenuIndex% = 0 To mnNumBoards - 1
            NotFound% = True
            NameStart% = InStr(mnuBoard(MenuIndex%).Caption, ") ") + 2
            If Mid$(mnuBoard(MenuIndex%).Caption, NameStart%) = BoardParams(0) Then
               If DupeFound% = DupeIndex% Then
                  mnuBoard_Click (MenuIndex%)
                  NotFound% = False
                  Exit For
               End If
               DupeFound% = DupeFound% + 1
            End If
         Next MenuIndex%
         If NotFound% Then
            MsgBox SearchName$ & " not available in list of currently installed boards. Aborting script.", , "Requested Board Not Available"
            gnScriptRun = False
         End If
      Case "C" 'set connection
         If mlDirection = SIGNAL_IN Then
            Select Case value&
               Case AUXIN0
                  ItemToSelect& = 0
               Case AUXIN1
                  ItemToSelect& = 1
               Case AUXIN2
                  ItemToSelect& = 2
               Case AUXIN3
                  ItemToSelect& = 3
               Case AUXIN4
                  ItemToSelect& = 4
               Case AUXIN5
                  ItemToSelect& = 5
               Case DS_CONNECTOR
                  ItemToSelect& = 6
            End Select
         ElseIf mlDirection = CBDISABLED Then
            'SelectedIO& = SelectedIO& + Choose(IOItem% + 1, AUXOUT0, AUXOUT1, AUXOUT2, DS_CONNECTOR)
            'no connection options
         Else
            Select Case value&
               Case AUXOUT0
                  ItemToSelect& = 0
               Case AUXOUT1
                  ItemToSelect& = 1
               Case AUXOUT2
                  ItemToSelect& = 2
               Case DS_CONNECTOR
                  ItemToSelect& = 3
            End Select
         End If
         If lstATCCIn.ListCount > 0 Then lstATCCIn.Selected(ItemToSelect&) = True
      Case "D" 'set direction
         mlDirection = value&
         Select Case value&
            Case 0
               optSelIO(2).value = True
            Case 2
               optSelIO(0).value = True
            Case 4
               optSelIO(1).value = True
         End Select
      Case "E" 'set DevNum
         cmbDevNum.Text = Format(value&, "0")
      Case "F" 'set function
         mnuFuncArray_Click (value&)
      Case "I" 'set ConfigItem
         LastDefinedItem& = cmbConfigItem.ListCount
         If value& < LastDefinedItem& Then
            cmbConfigItem.ListIndex = value&
         Else
            cmbConfigItem.Text = Format(value&, "0")
         End If
      Case "P" 'set polarity
         'to do
      Case "R" 'reset all list boxes to none selected
         If lstATCCIn.ListCount > 0 Then
            For i% = 0 To lstATCCIn.ListCount - 1
               lstATCCIn.Selected(i%) = False
            Next i%
         End If
         If lstSignals.ListCount > 0 Then
            For i% = 0 To lstSignals.ListCount - 1
               lstSignals.Selected(i%) = False
            Next i%
         End If
      Case "S" 'set Signal
         If mlDirection = SIGNAL_IN Then
            Select Case value&
               Case ADC_CONVERT
                  ItemToSelect& = 0
               Case ADC_GATE
                  ItemToSelect& = 1
               Case ADC_START_TRIG
                  ItemToSelect& = 2
               Case ADC_STOP_TRIG
                  ItemToSelect& = 3
               Case ADC_TB_SRC
                  ItemToSelect& = 4
               Case DAC_UPDATE
                  ItemToSelect& = 5
               Case DAC_TB_SRC
                  ItemToSelect& = 6
               Case DAC_START_TRIG
                  ItemToSelect& = 7
               Case SYNC_CLK
                  ItemToSelect& = 8
            End Select
         ElseIf mlDirection = CBDISABLED Then
            Select Case value&
               Case ADC_TB_SRC
                  ItemToSelect& = 0
               Case DAC_TB_SRC
                  ItemToSelect& = 1
               Case SYNC_CLK
                  ItemToSelect& = 2
            End Select
         Else
            Select Case value&
               Case ADC_CONVERT
                  ItemToSelect& = 0
               Case ADC_START_TRIG
                  ItemToSelect& = 1
               Case ADC_STOP_TRIG
                  ItemToSelect& = 2
               Case ADC_SCANCLK
                  ItemToSelect& = 3
               Case ADC_SSH
                  ItemToSelect& = 4
               Case ADC_STARTSCAN
                  ItemToSelect& = 5
               Case ADC_SCAN_STOP
                  ItemToSelect& = 6
               Case DAC_UPDATE
                  ItemToSelect& = 7
               Case DAC_START_TRIG
                  ItemToSelect& = 8
               Case SYNC_CLK
                  ItemToSelect& = 9
               Case CTR1_CLK
                  ItemToSelect& = 10
               Case CTR2_CLK
                  ItemToSelect& = 11
               Case DGND
                  ItemToSelect& = 12
            End Select
         End If
         If lstSignals.ListCount > 0 Then lstSignals.Selected(ItemToSelect&) = True
      Case "T" 'set InfoType
         'to do - update for new list options
         cmbInfoType.ListIndex = value&
      Case "V" 'set ConfigVal
         chkHex.value = 0  'ignored when script running
         txtConfigVal.Text = value&
      Case "X" 'execute the function
         cmdATCC = True
      Case "a" 'set limitA
         txtLimitA.Text = Mid$(cmdConfigure.Caption, 2)
      Case "b" 'set limitA
         txtLimitB.Text = Mid$(cmdConfigure.Caption, 2)
      Case "c" 'set Out1
         txtOut1.Text = Mid$(cmdConfigure.Caption, 2)
      Case "d" 'set Out2
         txtOut2.Text = Mid$(cmdConfigure.Caption, 2)
      Case "e" 'set Mask1
         txtMask1.Text = Mid$(cmdConfigure.Caption, 2)
      Case "f" 'set Mask2
         txtMask2.Text = Mid$(cmdConfigure.Caption, 2)
      Case "i" 'load the queue element
         cmdLoadArray = True
      Case "j" 'finish the queue setup
         cmdDone = True
      Case "k" 'set Latch
         optDTEvent(0).value = True
         If value& = STOP_EVENT Then optDTEvent(1).value = True
      Case "l" 'set Latch
         chkLatch.value = value&
      Case "m" 'set trigger source
         cmbDTSource.ListIndex = value&
      Case "n" 'set trigger sense
         cmbDTSense.ListIndex = value&
      Case "o" 'set trigger channel
         txtDTChan.Text = value&
      Case "p" 'set channel type
         cmbDTType.ListIndex = value&
      Case "q" 'quit - stop background task
         mnCancel = True
      Case "r" 'set range
         Select Case value&
            Case Is < 0
               mnuNoRange_Click
            Case Is < 100
               varRange = Choose(value& + 1, 2, 1, 4, 6, 7, 8, 9, 12, 13)
               If IsNull(varRange) Then
                  NewRange = value& - 8
                  varRange = Choose(NewRange, 14, 15, 16, 10, 11, 5, 0, 3)
               End If
               MenuIndex% = varRange
               mnuBipRange_Click (MenuIndex%)
            Case Is < 200
               varRange = Choose(value& - 99, 0, 1, 3, 4, 5, 6, 10, 12, 13, 14)
               If IsNull(varRange) Then
                  NewRange = value& - 109
                  varRange = Choose(NewRange, 7, 8, 9, 11, 2)
               End If
               MenuIndex% = varRange
               mnuUniRange_Click (MenuIndex%)
            Case Is < 300
               mnuCurRange_Click (value& - 200)
         End Select
      Case "s" 'set trigger range (changed from Case "q" 6/09)
         cmbDTType.ListIndex = value&
      Case "t" 'set trigger level (float)
         txtDTLevel.Text = Mid$(cmdConfigure.Caption, 2)
      Case "u" 'set number of queue elements
         txtQCount.Text = value&
      Case "v" 'select a queue element
         lstElement.ListIndex = value&
      Case "w" 'select a flag type and latch or update on true and false
         FlagValue& = value& And 7
         LatchValue& = 1
         If (value& And 8) = 8 Then LatchValue& = 0
         cmbSPFlags.ListIndex = FlagValue&
         chkLatch.value = LatchValue&
      Case "x" 'execute set trigger
         cmdDTTrig = True
      Case "y" 'set trigger level variance (float)
         txtDTVar.Text = Mid$(cmdConfigure.Caption, 2)
      Case "z" 'select an output type
         cmbSPOutput.ListIndex = value&
   End Select

End Sub

Private Sub cmdDTTrig_Click()

   DTSource& = Choose(cmbDTSource.ListIndex + 1, _
      TRIG_IMMEDIATE, TRIG_EXTTTL, TRIG_ANALOG_HW, _
      TRIG_ANALOG_SW, TRIG_DIGPATTERN, TRIG_COUNTER, TRIG_SCANCOUNT)
   DTSense& = Choose(cmbDTSense.ListIndex + 1, RISING_EDGE, _
      FALLING_EDGE, HIGH_LEVEL, LOW_LEVEL, ABOVE_LEVEL, BELOW_LEVEL, EQ_LEVEL, NE_LEVEL)
   DTType& = Choose(cmbDTType.ListIndex + 1, ANALOG, DIGITAL8, _
      DIGITAL16, CTR16, CTR32LOW, CTR32HIGH, CJC, TC)
   DTChan& = Val(txtDTChan.Text)
   DTLevel! = Val(txtDTLevel.Text)
   DTVariance! = Val(txtDTVar.Text)
   DTTrigEvent& = START_EVENT
   If optDTEvent(1).value Then DTTrigEvent& = STOP_EVENT
   ULStat = IOTDaqSetTrigger(mnBoardNum, DTSource&, DTSense&, _
      DTChan&, DTType&, mnRange, DTLevel!, DTVariance!, DTTrigEvent&)
   If SaveFunc(Me, DaqSetTrigger, ULStat, mnBoardNum, DTSource&, _
      DTSense&, DTChan&, DTType&, mnRange, DTLevel!, DTVariance!, _
      DTTrigEvent&, A10, A11, 0) Then Exit Sub

End Sub

Private Sub cmdOK_Click()

   UpdateStatus

End Sub

Private Sub cmdOK_DragDrop(Source As Control, x As Single, y As Single)

   'lblChipFunctions.Caption = ""
   'lblQuery.Visible = True

End Sub

Private Sub ConfigureControls()
   
   Dim EnableItems As Boolean
   
   txtConfigVal.Width = 855
   fraSetPoint.Visible = False
   Me.chkHex.Caption = "Hex"
   If mnLibType = MSGLIB Then
      txtConfigVal.Width = 1795
      For MenuIndex% = 3 To mnuFuncArray.Count - 1
         mnuFuncArray(MenuIndex%).Checked = False
         mnuFuncArray(MenuIndex%).ENABLED = False
      Next
      mnuFuncArray(2).Caption = "Reflection"
      mnuLoadConfig.ENABLED = False
   Else
      For MenuIndex% = GET_SIGNAL To mnuFuncArray.Count - 1
         mnuFuncArray(MenuIndex%).ENABLED = True
      Next
      mnuFuncArray(GET_SIGNAL).Caption = "cbGetSignal()"
      mnuLoadConfig.ENABLED = True
   End If
   Me.cmbInfoType.Visible = True
   Me.cmbDevNum.Visible = True
   Me.chkHex.Visible = True
   If gnNumBoards = 0 Then
      EnableItems = (cmbInfoType.ListIndex = 0)
      Me.cmbConfigItem.ENABLED = EnableItems
      Me.cmbDevNum.ENABLED = EnableItems
      If Not EnableItems Then Exit Sub
   End If
   Select Case mnFuncType
      Case GET_CONF, SET_CONF
         fraATCC.Visible = False
         fraDaqTrig.Visible = False
         fraConfiguration.Visible = True
         cmbConfigItem.ENABLED = True
         Select Case mnLibType
            Case UNILIB
               LoadULItems
            Case MSGLIB
               LoadMsgItems
         End Select
      Case GET_SIGNAL
         Select Case mnLibType
            Case UNILIB
               optSelIO(2).Visible = False
               chkInvert.Visible = False
               lstATCCIn.Left = 600
               lstInvert.Visible = True
               lstSignals.Left = 2460
               chkInvert.Visible = False
               fraDaqTrig.Visible = False
               fraATCC.Visible = True
               fraConfiguration.Visible = False
               fraATCC.Top = 120
               lblStatus.Caption = gsConfig
               lstATCCIn.MousePointer = 12
            Case MSGLIB
               LoadMsgItems
               Me.chkHex.Caption = "Numeric"
         End Select
      Case DAQ_TRIG
         fraATCC.Visible = False
         fraDaqTrig.Visible = True
         fraConfiguration.Visible = False
         fraDaqTrig.Top = 120
      Case GET_STRING, SET_STRING
         fraATCC.Visible = False
         fraDaqTrig.Visible = False
         fraConfiguration.Visible = True
         txtConfigVal.Width = 1845
         Me.cmbInfoType.ListIndex = 1
         Select Case mnLibType
            Case UNILIB
               LoadULItems
            Case MSGLIB
               LoadMsgItems
         End Select
      Case SELECT_SIGNAL
         lstInvert.Visible = False
         lstATCCIn.Left = 60
         lstSignals.Left = 1920
         If mlDirection = SIGNAL_OUT Then chkInvert.Visible = True
         fraDaqTrig.Visible = False
         fraATCC.Visible = True
         fraConfiguration.Visible = False
         fraATCC.Top = 120
         lblStatus.Caption = gsConfig
         lstATCCIn.MousePointer = vbDefault
         optSelIO(2).Visible = True
      Case SET_SETPOINTS
         fraSetPoint.Top = 0
         fraSetPoint.Left = frmNewCfg(mnThisInstance).ScaleHeight * 0.04
         fraSetPoint.Width = frmNewCfg(mnThisInstance).Width * 0.955
         fraSetPoint.Height = frmNewCfg(mnThisInstance).ScaleHeight - lblStatus.Height
         fraSetPoint.Visible = True
   End Select

End Sub

Private Sub Form_Activate()

   UpdateMainStatus

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
            'add configuration type if applicable
            Case ANALOG_IN
               NumProps% = GetAIProps(msBoardName, MsgLibrary, PropList)
               'UpdateFormProps PropList, NumProps%
               'QChanged% = ConfigureMsgQueue()
         End Select
      End If
   End If
   If gnScriptRun Then
      For cBox% = 0 To Me.chkFilter.Count - 1
         chkFilter(cBox%).value = 1
      Next
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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
   ElseIf (Shift And 2) = 2 Then
      Select Case KeyCode
         Case 13  'Enter
            cmdOK = True
      End Select
   Else
      Select Case KeyCode
         Case 118 'F7 - set default child size
            Me.Height = 2650
            Me.Width = 6200
         Case 120 'F9 - set height to 1/3 screen
            mfmUniTest.Height = Screen.Height / 3
         Case 122 'F11 - set to screen bottom
            mfmUniTest.Move 0, Screen.Height - mfmUniTest.Height, Screen.Width
         Case 123 'F12 - set to screen top
            mfmUniTest.Move 0, 0, Screen.Width
      End Select
   End If
   'KeyCode = 0

End Sub

Private Sub Form_Load()
   
#If MSGOPS Then
   If gnLibType = MSGLIB Then
      Set MsgLibrary = New MBDClass.MBDComClass
   End If
#End If

#If NETOPS Then
   If gnLibType = NETLIB Then
      'Set ULNetLibrary = New AcqThread.AcqThread
   End If
#End If
   
   mnLibType = gnLibType
   mnNumBoards = gnNumBoards
   mnuLibrary(gnLibType).Checked = True
   If Not gbULLoaded Then
      mnuLibrary(UNILIB).Checked = False
      mnuLibrary(UNILIB).ENABLED = False
   End If
   lstATCCIn.Left = 60
   lstSignals.Left = 1920
   For i% = 0 To gnNumBoards - 1
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
      msDisplayName = DisplayName$
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   
   Select Case gnLibType
      Case UNILIB
         LoadULLists
      Case MSGLIB
         LoadMsgLists
   End Select
   
End Sub

Private Sub Form_Resize()

   lblStatus.Width = ScaleWidth
   lblStatus.Top = ScaleHeight - lblStatus.Height
   cmbList.Top = lblStatus.Top - 80
   cmbList.Width = lblStatus.Width - 200
   If gnInitializing Then
      'if the form is just loading, sets the form type
      'and sets the default function (cbC8254Config())
      mnFormType = (Val("&H" & Tag) And &HF00&) / &H100
      mnThisInstance = Val("&H" & Tag) And &HFF
      If gbULLoaded Then
         mnuFuncArray_Click (GET_CONF)
         'msTitle = Caption
         'mnuBoard_Click (0)
      End If
      msTitle = Caption
      gnInitializing = False
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Select Case mnLibType
      Case UNILIB
         'If mnNumBoards > 0 Then ULStat = StopBackground520(mnBoardNum, mlStatusType)
      Case NETLIB
         Set AcqThread = Nothing
         mnThreading = False
      Case MSGLIB
         If Not MsgLibrary Is Nothing Then
            If mnMessaging Then
               If Not (MsgLibrary.DeviceID = "") Then _
               MsgLibrary.ReleaseDevice (msBoardName)
               Set MsgLibrary = Nothing
               mnMessaging = False
            End If
         End If
   End Select

   UnLoadChildForm Me, mnFormType, mnThisInstance
   gnCfgForms = gnCfgForms - 1

End Sub

Private Sub lblStatus_DblClick()

   Clipboard.SetText (lblStatus.Caption)
   MsgBox "Status text copied to clipboard.", vbOKOnly, "Data Copied"
   
End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash

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
         If Not MsgLibrary Is Nothing Then
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
            End If
         Else
            mnLibType = gnLibType
            Me.mnuLibrary(MSGLIB).Checked = False
            Me.mnuLibrary(MSGLIB).ENABLED = False
         End If
   End Select
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
         LoadMsgLists
      Else
         BoardNum% = gnBoardEnum(i%)
         BoardName$ = GetNameOfBoard(BoardNum%)
         DisplayName$ = BoardName$
         LoadULLists
      End If
      CurrentName$ = BoardNum% & ") " & DisplayName$
      msDisplayName = DisplayName$
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
   If mnNumBoards > 0 Then
      mnuBoard_Click (0)
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

End Sub

Private Sub mnuBoard_Click(Index As Integer)

   Caption = msTitle & " board " & mnuBoard(Index).Caption
   If Not (gnNumBoards > 0) Then Exit Sub
   If mnLibType = UNILIB Then
      BoardNum% = gnBoardEnum(Index)
   Else
      BoardNum% = Index
   End If
   If Not gnInitializing Then mnuBoard(mnBoardIndex).Checked = False
   Caption = msTitle & " Board " & mnuBoard(Index).Caption
   If mnNumBoards = 0 Then Exit Sub
   mnBoardNum = BoardNum%
   mnBoardIndex = Index
   mnuBoard(Index).Checked = True
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
         Component$ = "AI"
         msAiSupport = MsgLibrary.GetSupportedMessages(Component$)
         mnTempSupport = (InStr(1, msAiSupport, "CJC") > 0) Or (InStr(1, msAiSupport, "SENSOR=TC") > 0)
         Scales% = mnuScale.Count - 1
         For ScaleMenu% = 0 To 2 'Scales%
            mnuScale(ScaleMenu%).ENABLED = mnTempSupport
         Next
         mnuScale(RAW - 1).ENABLED = mnTempSupport
      End If
   Else
      BoardName$ = GetNameOfBoard(mnBoardNum)
      DisplayName$ = BoardName$
   End If
   msBoardName = BoardName$
   msDisplayName = DisplayName$
   UpdateMainStatus
   If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then
      ConfigureControls
      UpdateStatus
   End If
   If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then
      ConfigureControls
      UpdateStatus
   End If
   A1 = msDisplayName
   If gnScriptSave And (Not gnInitializing) Then
      FuncStat = 0
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

End Sub

Private Sub mnuClose_Click()

   Unload Me

End Sub

Private Sub mnuCurRange_Click(Index As Integer)

   SetRange 2, Index

End Sub

Private Sub mnuExit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub mnuFuncArray_Click(Index As Integer)

   mnuFuncArray(mnFuncType).Checked = False
   mnFuncType = Index
   gsConfig = mnuFuncArray(mnFuncType).Caption
   mnuFuncArray(mnFuncType).Checked = True
   If (Index = GET_STRING) Or (Index = SET_STRING) Then
      chkFilter(0).value = 1
      For CheckFilter% = 1 To Me.chkFilter.Count - 1
         chkFilter(CheckFilter%).value = 0
      Next
   End If
   ConfigureControls
   If mnLibType = MSGLIB Then
      LoadMsgLists
      Me.chkReQuery.Visible = (Index = SET_CONF)
   Else
      LoadULLists
   End If
   txtConfigVal.Visible = ((Index = SET_CONF) Or (Index = SET_STRING))
   lblConfigVal.Visible = ((Index = SET_CONF) Or (Index = SET_STRING))
   cmdOK.Visible = (Index = SET_CONF) Or (Index = SET_STRING)
   If cmbInfoType.ENABLED Then
      If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
      If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus
   End If
   
End Sub

Private Sub mnuLoadConfig_Click()

   Filename$ = ""
   ULStat = GetCfgFile(Filename$)
   
End Sub

Private Sub mnuScale_Click(Index As Integer)

   CurrentSelection% = Switch(mnScale = CELSIUS, 0, mnScale = FAHRENHEIT, 1, _
   mnScale = KELVIN, 2, mnScale = VOLTS, 3, mnScale = NOSCALE, 4, mnScale = RAW, 5)
   mnuScale(CurrentSelection%).Checked = False
   mnScale = Choose(Index + 1, CELSIUS, FAHRENHEIT, KELVIN, VOLTS, NOSCALE, RAW)
   NewSelection% = Switch(mnScale = CELSIUS, 0, mnScale = FAHRENHEIT, 1, _
   mnScale = KELVIN, 2, mnScale = VOLTS, 3, mnScale = NOSCALE, 4, mnScale = RAW, 5)
   mnuScale(NewSelection%).Checked = True
   If mnLibType = MSGLIB Then LoadMsgItems
   
End Sub

Private Sub mnuUniRange_Click(Index As Integer)

   SetRange 1, Index

End Sub

Private Sub optSelIO_Click(Index As Integer)

   lstSignals.Clear
   lstATCCIn.Clear
   lstInvert.Clear
   
   If optSelIO(0).value Then
      mlDirection = SIGNAL_IN
      'chkInvert.Visible = False
      lstATCCIn.AddItem "AUXIN0"
      lstATCCIn.AddItem "AUXIN1"
      lstATCCIn.AddItem "AUXIN2"
      lstATCCIn.AddItem "AUXIN3"
      lstATCCIn.AddItem "AUXIN4"
      lstATCCIn.AddItem "AUXIN5"
      lstATCCIn.AddItem "DS_CONNECTOR"
      lstSignals.AddItem "ADC_CONVERT"
      lstSignals.AddItem "ADC_GATE"
      lstSignals.AddItem "ADC_START_TRIG"
      lstSignals.AddItem "ADC_STOP_TRIG"
      lstSignals.AddItem "ADC_TB_SRC"
      lstSignals.AddItem "DAC_UPDATE"
      lstSignals.AddItem "DAC_TB_SRC"
      lstSignals.AddItem "DAC_START_TRIG"
      lstSignals.AddItem "SYNC_CLK"
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
   ElseIf optSelIO(1).value Then
      mlDirection = SIGNAL_OUT
      If mnFuncType = GET_SIGNAL Then
         lstInvert.Visible = True
         lstATCCIn.Left = 600
         lstSignals.Left = 2460
      Else
         chkInvert.Visible = True
      End If
      lstSignals.AddItem "ADC_CONVERT"
      lstSignals.AddItem "ADC_START_TRIG"
      lstSignals.AddItem "ADC_STOP_TRIG"
      lstSignals.AddItem "ADC_SCANCLK"
      lstSignals.AddItem "ADC_SSH"
      lstSignals.AddItem "ADC_STARTSCAN"
      lstSignals.AddItem "ADC_SCAN_STOP"
      lstSignals.AddItem "DAC_UPDATE"
      lstSignals.AddItem "DAC_START_TRIG"
      lstSignals.AddItem "SYNC_CLK"
      lstSignals.AddItem "CTR1_CLK"
      lstSignals.AddItem "CTR2_CLK"
      lstSignals.AddItem "DGND"
      lstATCCIn.AddItem "AUXOUT0"
      lstATCCIn.AddItem "AUXOUT1"
      lstATCCIn.AddItem "AUXOUT2"
      lstATCCIn.AddItem "DS_CONNECTOR"
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
   Else
      mlDirection = CBDISABLED
      chkInvert.Visible = False
      lstSignals.AddItem "ADC_TB_SRC"
      lstSignals.AddItem "DAC_TB_SRC"
      lstSignals.AddItem "SYNC_CLK"
   End If
   lstSignals.Selected(0) = True

End Sub

Private Sub SetRange(RangeType As Integer, RangeIndex As Integer)

   If (msBoardName = "") Or (msDisplayName = "") Then
      'this shouldn't happen
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
      msBoardName = GetNameOfBoard(mnBoardNum)
   End If
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
            BIPPT01VOLTS, BIPPT005VOLTS, BIP1PT67VOLTS, BIPPT156VOLTS, _
            BIPPT312VOLTS, BIPPT078VOLTS, BIP60VOLTS, BIP15VOLTS, _
            BIPPT125VOLTS, BIPPT025VOLTSPERVOLT, BIPPT073125VOLTS)
         End If
         If ((IsNull(varRange)) Or (varRange = "")) Then
            MsgBox "Range not supported before revision " & _
            Format(CURRENTREVNUM, "General Number") & ".", , "UL Update Required"
            Exit Sub
         End If
         mnRange = varRange
         GoSub ClearMenus
         mnuBipRange(RangeIndex).Checked = True
         RangeDivisor! = 2
         Prefix$ = "±"
         Suffix$ = "V"
      Case 1   'Unipolar
         varRange = Choose(RangeIndex + 1, UNI10VOLTS, UNI5VOLTS, UNI4VOLTS, UNI2PT5VOLTS, UNI2VOLTS, UNI1PT25VOLTS, UNI1VOLTS, UNIPT5VOLTS, UNIPT25VOLTS, UNIPT2VOLTS)
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
         mnRange = Choose(RangeIndex + 1, MA4TO20, MA2to10, _
         MA1TO5, MAPT5TO2PT5, MA0TO20, BIPPT025AMPS)
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
         End If
         If mnRange = BIPPT025AMPS Then
            Prefix$ = "±"
            Suffix$ = "A"
            RangeDivisor! = 2
         End If
   End Select
   msRange = GetRangeString(mnRange)
   If RangeType = 2 Then
      'mnuRange.Caption = "&Range (" & Prefix$ = RangeLow! & " to " & RangeHigh! & "mA)"
      RangeVolts! = RangeVolts! / RangeDivisor!
      mnuRange.Caption = "&Range (" & Prefix$ & RangeVolts! & Suffix$ & ")"
   Else
      If Not mnRange = NOTUSED Then
         RangeVolts! = GetRangeVolts(mnRange)
         RangeVolts! = RangeVolts! / RangeDivisor!
         mnuRange.Caption = "&Range (" & Prefix$ & RangeVolts! & Suffix$ & ")"
      Else
         mnuRange.Caption = "&Range (NOTUSED)"
      End If
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

Private Sub UpdateMainStatus()

   board$ = mnuBoard(mnBoardIndex).Caption
   PrintMain "Current board: " & board$
   
End Sub

Private Sub UpdateStatus()

   If Not mbHoldoffUpdate Then
      Select Case mnLibType
         Case UNILIB
            UpdateULStatus
         Case MSGLIB
            UpdateMsgStatus
      End Select
   Else
      'mbHoldoffUpdate = False
   End If

End Sub

Private Sub hsbQCount_Change()

   txtQCount.Text = hsbQCount.value
   CheckQueue

End Sub

Private Sub CheckQueue()
   
   CurElement& = lstElement.ListIndex
   If CurElement& < 0 Then Exit Sub
   CurFlag% = Choose(cmbSPFlags.ListIndex + 1, SF_EQUAL_LIMITA, SF_LESSTHAN_LIMITA, _
   SF_INSIDE_LIMITS, SF_GREATERTHAN_LIMITB, SF_OUTSIDE_LIMITS, SF_HYSTERESIS)

   cmdLoadArray.ENABLED = Not (Val(txtLimitA.Text) = mafLimAArray(CurElement&))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (Val(txtLimitB.Text) = mafLimBArray(CurElement&))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (Val(txtOut1.Text) = mafOut1Array(CurElement&))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (Val(txtOut2.Text) = mafOut2Array(CurElement&))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (Val(txtMask1.Text) = mafMask1Array(CurElement&))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (Val(txtMask1.Text) = mafMask2Array(CurElement&))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (cmbSPFlags.ListIndex = (malSPFlags(CurElement&) And 7))
   cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (cmbSPOutput.ListIndex = malOutputs(CurElement&))
   If ((malSPFlags(CurElement&) And 8) = 0) Then
      cmdLoadArray.ENABLED = (cmdLoadArray.ENABLED Or (chkLatch.value = 0))
   Else
      cmdLoadArray.ENABLED = cmdLoadArray.ENABLED Or Not (chkLatch.value = 0)
   End If

End Sub

Private Sub txtConfigVal_KeyDown(KeyCode As Integer, Shift As Integer)

   If (KeyCode = 13) And ((mnFuncType = GET_CONF) _
      Or (mnFuncType = GET_STRING)) Then UpdateStatus
   If KeyCode = 45 Then
      frmComposite.Caption = "Multiline ConfigVal"
      frmComposite.txtShow.Top = 60
      frmComposite.txtShow.Left = 120
      frmComposite.txtShow.Visible = True
      frmComposite.chkComposite.Visible = False
      frmComposite.chkConsecutive.Visible = False
      frmComposite.chkMaskFirst.Visible = False
      frmComposite.chkMaskSecond.Visible = False
      frmComposite.Show 1
      DisplayText$ = frmComposite.txtShow.Text
      If Len(DisplayText$) > 12 Then
         Me.txtConfigVal.Text = Left(DisplayText$, 12) & "..."
      Else
         Me.txtConfigVal.Text = DisplayText$
      End If
   End If

End Sub

Private Sub txtConfigVal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
   End If

End Sub

Private Sub txtQCount_Change()

   QCount& = Val(txtQCount.Text) - 1
   
   If Not (QCount& < 0) Then
      ReDim Preserve mafLimAArray(QCount&)
      ReDim Preserve mafLimBArray(QCount&)
      ReDim Preserve mafOut1Array(QCount&)
      ReDim Preserve mafOut2Array(QCount&)
      ReDim Preserve mafMask1Array(QCount&)
      ReDim Preserve mafMask2Array(QCount&)
      ReDim Preserve malSPFlags(QCount&)
      ReDim Preserve malOutputs(QCount&)
   Else
      txtLimitA.Text = ""
      txtLimitB.Text = ""
      txtOut1.Text = ""
      txtOut2.Text = ""
      txtMask1.Text = ""
      txtMask2.Text = ""
      cmbSPFlags.ListIndex = 0
      cmbSPOutput.ListIndex = 0
   End If
   cmdDone.ENABLED = (mlQCount > QCount&)
   mlQCount = QCount& + 1
   lstElement.Clear
   For Element% = 0 To mlQCount - 1
      lstElement.AddItem Format$(Element%, "0"), Element%
   Next Element%
   
End Sub

Private Sub cmdLoadArray_Click()

   CurElement& = lstElement.ListIndex
   If Not (CurElement& < 0) Then
      CurOut& = Choose(cmbSPOutput.ListIndex + 1, SO_NONE, SO_FIRSTPORTC, _
      SO_DIGITALPORT, SO_DAC0, SO_DAC1, SO_DAC2, SO_DAC3, SO_TMR0, SO_TMR1)

      CurFlag& = Choose(cmbSPFlags.ListIndex + 1, SF_EQUAL_LIMITA, SF_LESSTHAN_LIMITA, _
      SF_INSIDE_LIMITS, SF_GREATERTHAN_LIMITB, SF_OUTSIDE_LIMITS, SF_HYSTERESIS)
      If Not (chkLatch.value = 1) Then CurFlag& = CurFlag& Or SF_UPDATEON_TRUEANDFALSE

      mafLimAArray(CurElement&) = Val(txtLimitA.Text)
      mafLimBArray(CurElement&) = Val(txtLimitB.Text)
      mafOut1Array(CurElement&) = Val(txtOut1.Text)
      mafOut2Array(CurElement&) = Val(txtOut2.Text)
      mafMask1Array(CurElement&) = Val(txtMask1.Text)
      mafMask2Array(CurElement&) = Val(txtMask2.Text)
      malSPFlags(CurElement&) = CurFlag&
      malOutputs(CurElement&) = CurOut&
   End If
   cmdDone.ENABLED = True
   cmdLoadArray.ENABLED = False
   If gnScriptSave Then
      FuncStat = 0
      A1 = 0
      A2 = Val(txtQCount.Text)
      A3 = CurElement&
      A4 = Val(txtLimitA.Text)
      A5 = Val(txtLimitB.Text)
      A6 = Val(txtOut1.Text)
      A7 = Val(txtOut2.Text)
      A8 = Val(txtMask1.Text)
      A9 = Val(txtMask2.Text)
      A10 = CurFlag&
      A11 = CurOut&
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, DaqSetSetpoints, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If

End Sub

Private Sub cmdDone_Click()

   If mlQCount > 0 Then
      ULStat = cbDaqSetSetpoints(mnBoardNum, mafLimAArray(0), mafLimBArray(0), Reserved!, malSPFlags(0), malOutputs(0), mafOut1Array(0), mafOut2Array(0), mafMask1Array(0), mafMask2Array(0), mlQCount)
      If SaveFunc(Me, DaqSetSetpoints, ULStat, mnBoardNum, mafLimAArray(0), mafLimBArray(0), Reserved!, malSPFlags(0), malOutputs(0), mafOut1Array(0), mafOut2Array(0), mafMask1Array(0), mafMask2Array(0), mlQCount, 0) Then Exit Sub
   Else
      ULStat = cbDaqSetSetpoints(mnBoardNum, NullFlt!, NullFlt!, Reserved!, NullLong&, NullLong&, NullFlt!, NullFlt!, NullFlt!, NullFlt!, mlQCount)
      If SaveFunc(Me, DaqSetSetpoints, ULStat, mnBoardNum, NullFlt!, NullFlt!, Reserved!, NullLong&, NullLong&, NullFlt!, NullFlt!, NullFlt!, NullFlt!, mlQCount, 0) Then Exit Sub
   End If
   cmdDone.ENABLED = False

End Sub

Private Sub lstElement_Click()

   CurElement& = lstElement.ListIndex
   If Not CurElement& < 0 Then
      chkLatch.value = 0
      txtLimitA.Text = mafLimAArray(CurElement&)
      txtLimitB.Text = mafLimBArray(CurElement&)
      txtOut1.Text = mafOut1Array(CurElement&)
      txtOut2.Text = mafOut2Array(CurElement&)
      txtMask1.Text = mafMask1Array(CurElement&)
      txtMask2.Text = mafMask2Array(CurElement&)
      cmbSPFlags.ListIndex = malSPFlags(CurElement&) And 7
      cmbSPOutput.ListIndex = malOutputs(CurElement&)
      If Not ((malSPFlags(CurElement&) And 8) = 0) Then
         chkLatch.value = 0
      Else
         chkLatch.value = 1
      End If
   End If
   CheckQueue

End Sub

Private Sub chkLatch_Click()

   CheckQueue

End Sub

Private Sub cmbSPFlags_Change()

   CheckQueue

End Sub

Private Sub cmbSPOutput_Change()

   CheckQueue

End Sub

Private Sub txtLimitA_LostFocus()

   CheckQueue
   
End Sub

Private Sub txtLimitB_LostFocus()

   CheckQueue

End Sub

Private Sub txtMask1_LostFocus()

   CheckQueue

End Sub

Private Sub txtMask2_LostFocus()

   CheckQueue

End Sub

Private Sub txtOut1_LostFocus()

   CheckQueue

End Sub

Private Sub txtOut2_LostFocus()

   CheckQueue

End Sub

Private Sub cmbSPOutput_Click()

   CheckQueue
   
End Sub

Private Sub cmbSPFlags_Click()

   CheckQueue

End Sub

Public Function GetInstance() As Integer
   
   GetInstance = mnThisInstance
   
End Function

Sub LoadULLists()
   
   Dim NotDone As Boolean
   
   cmbInfoType.Clear
   If Not gbULLoaded Then
      fraConfiguration.ENABLED = False
      Exit Sub
   Else
      Do
         NotDone = GetInfoTypes(Index%, 31, CfgVal&, CfgName$)
         cmbInfoType.AddItem CfgName$
         cmbInfoType.ItemData(cmbInfoType.NewIndex) = CfgVal&
         Index% = Index% + 1
      Loop While NotDone
      If (mnFuncType = SET_STRING) Or (mnFuncType = GET_STRING) Then
         cmbInfoType.ListIndex = 1
      Else
         cmbInfoType.ListIndex = 0
      End If
      mnDigDevs = -1
      mnCtrDevs = -1

      lstATCCIn.Clear
      lstATCCIn.AddItem "AUXIN0"
      lstATCCIn.AddItem "AUXIN1"
      lstATCCIn.AddItem "AUXIN2"
      lstATCCIn.AddItem "AUXIN3"
      lstATCCIn.AddItem "AUXIN4"
      lstATCCIn.AddItem "AUXIN5"
      lstATCCIn.AddItem "DS_CONNECTOR"

      lstInvert.Clear
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""
      lstInvert.AddItem ""

      lstSignals.Clear
      lstSignals.AddItem "ADC_CONVERT"
      lstSignals.AddItem "ADC_GATE"
      lstSignals.AddItem "ADC_START_TRIG"
      lstSignals.AddItem "ADC_STOP_TRIG"
      lstSignals.AddItem "ADC_TB_SRC"
      lstSignals.AddItem "DAC_UPDATE"
      lstSignals.AddItem "DAC_TB_SRC"
      lstSignals.AddItem "DAC_START_TRIG"
      lstSignals.AddItem "SYNC_CLK"
      lstSignals.Selected(0) = True

      mlDirection = SIGNAL_IN
      mlPolarity = NONINVERTED

      cmbDTType.Clear
      cmbDTType.AddItem "ANALOG"
      cmbDTType.AddItem "DIGITAL8"
      cmbDTType.AddItem "DIGITAL16"
      cmbDTType.AddItem "CTR16"
      cmbDTType.AddItem "CTR32LOW"
      cmbDTType.AddItem "CTR32HIGH"
      cmbDTType.AddItem "CJC"
      cmbDTType.AddItem "TC"
      cmbDTType.ListIndex = 0
   
      cmbDTSource.Clear
      cmbDTSource.AddItem "TRIG_IMMEDIATE"
      cmbDTSource.AddItem "TRIG_EXTTTL"
      cmbDTSource.AddItem "TRIG_ANALOG_HW"
      cmbDTSource.AddItem "TRIG_ANALOG_SW"
      cmbDTSource.AddItem "TRIG_DIGPATTERN"
      cmbDTSource.AddItem "TRIG_COUNTER"
      cmbDTSource.AddItem "TRIG_SCANCOUNT"
      cmbDTSource.ListIndex = 0

      cmbDTSense.Clear
      cmbDTSense.AddItem "RISING_EDGE"
      cmbDTSense.AddItem "FALLING_EDGE"
      cmbDTSense.AddItem "HIGH_LEVEL"
      cmbDTSense.AddItem "LOW_LEVEL"
      cmbDTSense.AddItem "ABOVE_LEVEL"
      cmbDTSense.AddItem "BELOW_LEVEL"
      cmbDTSense.AddItem "EQ_LEVEL"
      cmbDTSense.AddItem "NE_LEVEL"
      cmbDTSense.ListIndex = 0

      cmbSPFlags.Clear
      cmbSPFlags.AddItem "SF_EQUAL_LIMITA"
      cmbSPFlags.AddItem "SF_LESSTHAN_LIMITA"
      cmbSPFlags.AddItem "SF_INSIDE_LIMITS"
      cmbSPFlags.AddItem "SF_GREATERTHAN_LIMITB"
      cmbSPFlags.AddItem "SF_OUTSIDE_LIMITS"
      cmbSPFlags.AddItem "SF_HYSTERESIS"
      cmbSPFlags.ListIndex = 0

      cmbSPOutput.Clear
      cmbSPOutput.AddItem "SO_NONE"
      cmbSPOutput.AddItem "SO_FIRSTPORTC"
      cmbSPOutput.AddItem "SO_DIGITALPORT"
      cmbSPOutput.AddItem "SO_DAC0"
      cmbSPOutput.AddItem "SO_DAC1"
      cmbSPOutput.AddItem "SO_DAC2"
      cmbSPOutput.AddItem "SO_DAC3"
      cmbSPOutput.AddItem "SO_TMR0"
      cmbSPOutput.AddItem "SO_TMR1"
      cmbSPOutput.ListIndex = 0
   End If

End Sub

Sub LoadMsgLists()

   cmbInfoType.Clear
   If Not (mnFuncType = GET_SIGNAL) Then cmbInfoType.AddItem "DEV"
   cmbInfoType.AddItem "AI"
   cmbInfoType.AddItem "AI{ch}"
   'cmbInfoType.AddItem "AICAL"
   cmbInfoType.AddItem "AISCAN"
   cmbInfoType.AddItem "AITRIG"
   cmbInfoType.AddItem "AIQUEUE"
   cmbInfoType.AddItem "AIQUEUE{el}"
   cmbInfoType.AddItem "AO"
   cmbInfoType.AddItem "AO{ch}"
   cmbInfoType.AddItem "AOSCAN"
   cmbInfoType.AddItem "DIO"
   cmbInfoType.AddItem "DIO{port}"
   If Not (mnFuncType = GET_SIGNAL) Then cmbInfoType.AddItem "DIO{port/bit}"
   cmbInfoType.AddItem "CTR"
   cmbInfoType.AddItem "CTR{ch}"
   cmbInfoType.AddItem "TMR"
   cmbInfoType.AddItem "TMR{ch}"
   cmbInfoType.AddItem "AICAL"
   cmbInfoType.AddItem "AICAL{ch}"
   cmbInfoType.AddItem "AOCAL"
   cmbInfoType.AddItem "AOCAL{ch}"
   cmbInfoType.ListIndex = 0

End Sub

Sub LoadULItems()

   Dim NotDone As Boolean, FilterOpts As Boolean
   Dim varCfgVal() As Variant, varCfgName() As Variant
   
   mnDevList1 = False
   cmbConfigItem.Clear
   
   If ((mnFuncType = SET_CONF) Or (mnFuncType = GET_CONF)) Then
      If Not (cmbInfoType.ItemData(cmbInfoType.ListIndex) = BOARDINFO) Then
         mbHoldoffUpdate = True
         If Not chkFilter(0).value = 1 Then chkFilter(0).value = 1
         mbHoldoffUpdate = False
      End If
      For CheckFilter% = 0 To Me.chkFilter.Count - 1
         If chkFilter(CheckFilter%).value = 1 Then
            CfgFilter& = (2 ^ CheckFilter%)
            Items& = GetCfgItems(CfgFilter&, varCfgVal(), varCfgName())
            For Itm& = 0 To Items&
               cmbConfigItem.AddItem varCfgName(Itm&)
               cmbConfigItem.ItemData(cmbConfigItem.NewIndex) = varCfgVal(Itm&)
            Next
         End If
      Next
   Else
      For CheckFilter% = 0 To Me.chkFilter.Count - 1
         If chkFilter(CheckFilter%).value = 1 Then
            CfgFilter& = (2 ^ CheckFilter%)
            Items& = GetCfgItems(CfgFilter&, varCfgVal(), varCfgName())
            For Itm& = 0 To Items&
               cmbConfigItem.AddItem varCfgName(Itm&)
               cmbConfigItem.ItemData(cmbConfigItem.NewIndex) = varCfgVal(Itm&)
            Next
         End If
      Next
   End If
   For CheckFilter% = 0 To Me.chkFilter.Count - 1
      chkFilter(CheckFilter%).Visible = False
   Next
   mbHoldoffUpdate = True
   If cmbConfigItem.ListCount > 0 Then cmbConfigItem.ListIndex = 0
   Select Case cmbInfoType.ItemData(cmbInfoType.ListIndex)
      Case GLOBALINFO '(1)
      Case BOARDINFO
         If ((mnFuncType = SET_CONF) Or (mnFuncType = GET_CONF)) Then
            For CheckFilter% = 0 To Me.chkFilter.Count - 1
               chkFilter(CheckFilter%).Visible = True
            Next
         Else
            chkFilter(0).value = 1
            chkFilter(0).Visible = True
            chkFilter(5).Visible = True
         End If
      Case DIGITALINFO, COUNTERINFO, EXPANSIONINFO
         chkFilter(0).Visible = True
         chkFilter(5).Visible = True
   End Select
   mbHoldoffUpdate = False

   If cmbInfoType.ENABLED Then
      If (mnFuncType = GET_CONF) Or (mnFuncType = GET_STRING) Then UpdateStatus
      If mnLibType = MSGLIB And (mnFuncType = GET_SIGNAL) Then UpdateStatus
   End If
   cmbDevNum.ListIndex = 0

End Sub

Sub LoadMsgItems()

   mnDevList1 = False
   cmbConfigItem.Clear
   cmbDevNum.Clear
   cmbDevNum.Visible = False
   cmbDevNum.AddItem "0"
   Component$ = cmbInfoType.Text
   cmbDevNum2.Visible = False
   Select Case Component$
      Case "DEV"
         cmbConfigItem.AddItem "GetSupportedMsgs"
         cmbConfigItem.AddItem "MFGSER"
         cmbConfigItem.AddItem "FWV"
         cmbConfigItem.AddItem "FPGAV"
         cmbConfigItem.AddItem "ID"
         cmbConfigItem.AddItem "MFGCAL"
         cmbConfigItem.AddItem "MFGCAL{YEAR}"
         cmbConfigItem.AddItem "MFGCAL{MONTH}"
         cmbConfigItem.AddItem "MFGCAL{DAY}"
         cmbConfigItem.AddItem "MFGCAL{HOUR}"
         cmbConfigItem.AddItem "MFGCAL{MINUTE}"
         cmbConfigItem.AddItem "MFGCAL{SECOND}"
         cmbConfigItem.AddItem "DATATYPE"
         cmbConfigItem.AddItem "FLASHLED"
         cmbConfigItem.AddItem "RESET"
         cmbConfigItem.AddItem "LOADCAPS"
         cmbConfigItem.AddItem "FPGACFG"
         cmbConfigItem.AddItem "TEMP{ch}"
         cmbConfigItem.AddItem "STATUS/ISO"
         cmbConfigItem.AddItem "SAVESTATE"
         cmbConfigItem.ListIndex = 0
      Case "AI"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "CHANNELS"
            cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "MAXRATE"
            cmbConfigItem.AddItem "RANGES"
            cmbConfigItem.AddItem "RES"
            cmbConfigItem.AddItem "INPUTS"
            cmbConfigItem.AddItem "CHMODES"
            cmbConfigItem.AddItem "DATARATES"
            cmbConfigItem.AddItem "SELFCAL"
            cmbConfigItem.AddItem "FACCAL"
            cmbConfigItem.AddItem "FIELDCAL"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem ""
            cmbConfigItem.AddItem "CHMODE"
            cmbConfigItem.AddItem "RES"
            cmbConfigItem.AddItem "VALIDCHANS"
            cmbConfigItem.AddItem "VALIDCHANS/CHMODE"
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "DATARATE"
            cmbConfigItem.AddItem "SCALE"
            cmbConfigItem.AddItem "CAL"
            cmbConfigItem.AddItem "SLOPE"
            cmbConfigItem.AddItem "OFFSET"
            cmbConfigItem.AddItem "ADCAL/START"
            cmbConfigItem.AddItem "ADCAL/STATUS"
         End If
         cmbConfigItem.ListIndex = 0
      Case "AI{ch}"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            'cmbConfigItem.AddItem "CHANNELS"
            'cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "RANGES"
            cmbConfigItem.AddItem "CHMODES"
            cmbConfigItem.AddItem "INPUTS"
            cmbConfigItem.AddItem "SENSORS"
            cmbConfigItem.AddItem "SENSORCONFIG"
            cmbConfigItem.AddItem "TCTYPES"
            cmbConfigItem.AddItem "CJC"
            'cmbConfigItem.AddItem "FACCAL"
            'cmbConfigItem.AddItem "FIELDCAL"
         Else
            If mnTempSupport Then
               Scales% = mnuScale.Count - 1
               For ScaleMenu% = 0 To Scales%
                  If mnuScale(ScaleMenu%).Checked Then
                     UnitString$ = Choose(ScaleMenu% + 1, _
                     "/DEGC", "/DEGF", "/KELVIN", _
                     "/VOLTS", "", "/RAW")
                     Exit For
                  End If
               Next
               cmbConfigItem.AddItem "VALUE" & UnitString$
            Else
               cmbConfigItem.AddItem "VALUE"
               cmbConfigItem.AddItem "VALUE/VOLTS"
               cmbConfigItem.AddItem "VALUE/RAW"
            End If
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "CHMODE"
            cmbConfigItem.AddItem "SENSOR"
            cmbConfigItem.AddItem "CJC" & UnitString$
            cmbConfigItem.AddItem "SLOPE"
            cmbConfigItem.AddItem "OFFSET"
            cmbConfigItem.AddItem "STATUS"
            cmbConfigItem.AddItem "DATARATE"
         End If
         cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "AISCAN"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "MAXSCANTHRUPUT"
            cmbConfigItem.AddItem "MINSCANRATE"
            cmbConfigItem.AddItem "MAXSCANRATE"
            cmbConfigItem.AddItem "BURSTMODE"
            cmbConfigItem.AddItem "MAXBURSTTHRUPUT"
            cmbConfigItem.AddItem "MINBURSTRATE"
            cmbConfigItem.AddItem "MAXBURSTRATE"
            cmbConfigItem.AddItem "QUEUECONFIG"
            cmbConfigItem.AddItem "QUEUESEQ"
            cmbConfigItem.AddItem "QUEUELEN"
            cmbConfigItem.AddItem "XFRMODES"
            cmbConfigItem.AddItem "EXTPACER"
            cmbConfigItem.AddItem "TRIG"
            cmbConfigItem.AddItem "FIFOSIZE"
            cmbConfigItem.AddItem "XFRSIZE"
            cmbConfigItem.ListIndex = 0
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "XFRMODE"
            cmbConfigItem.AddItem "RATE"
            cmbConfigItem.AddItem "SAMPLES"
            cmbConfigItem.AddItem "BUFSIZE"
            cmbConfigItem.AddItem "EXTPACER"
            cmbConfigItem.AddItem "BURSTMODE"
            cmbConfigItem.AddItem "BUFOVERWRITE"
            'cmbConfigItem.AddItem "EXTSYNC" 'obs
            cmbConfigItem.AddItem "LOWCHAN"
            cmbConfigItem.AddItem "HIGHCHAN"
            cmbConfigItem.AddItem "TRIG"
            cmbConfigItem.AddItem "QUEUE"
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "RANGE{ch}"
            cmbConfigItem.AddItem "RANGE{el/ch}"
            cmbConfigItem.AddItem "STATUS"
            cmbConfigItem.AddItem "COUNT"
            cmbConfigItem.AddItem "INDEX"
            cmbConfigItem.AddItem "SCALE"
            cmbConfigItem.AddItem "CAL"
            cmbConfigItem.AddItem "DEBUG"
            cmbConfigItem.AddItem "STALL"
            cmbConfigItem.AddItem "RESET"
            cmbConfigItem.AddItem "START"
            cmbConfigItem.AddItem "STOP"
            cmbConfigItem.ListIndex = 0
         End If
      Case "AITRIG"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "SRCS"
            cmbConfigItem.AddItem "TYPES"
            cmbConfigItem.AddItem "RANGES"
            cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "REARM"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "TYPE"
            cmbConfigItem.AddItem "REARM"
            cmbConfigItem.AddItem "SRC"
         End If
         cmbConfigItem.ListIndex = 0
      Case "AIQUEUE"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "DUMMY1"
            cmbConfigItem.AddItem "DUMMY2"
            cmbConfigItem.AddItem "DUMMY3"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "COUNT"
            cmbConfigItem.AddItem "CLEAR"
         End If
         cmbConfigItem.ListIndex = 0
      Case "AIQUEUE{el}"
         cmbConfigItem.AddItem "GetSupportedMsgs"
         cmbConfigItem.AddItem "CHAN"
         cmbConfigItem.AddItem "CHMODE"
         cmbConfigItem.AddItem "RANGE"
         cmbConfigItem.AddItem "DATARATE"
         cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "AO"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "CHANNELS"
            cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "MAXRATE"
            cmbConfigItem.AddItem "RANGES"
            cmbConfigItem.AddItem "OUTPUTS"
            cmbConfigItem.AddItem "SELFCAL"
            cmbConfigItem.AddItem "FACCAL"
            cmbConfigItem.AddItem "FIELDCAL"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem ""
            cmbConfigItem.AddItem "SCALE"
            cmbConfigItem.AddItem "UPDATE"
            cmbConfigItem.AddItem "RES"
         End If
         cmbConfigItem.ListIndex = 0
         UpdateMsgStatus
      Case "AO{ch}"
         If mnFuncType = GET_SIGNAL Then
            cmbConfigItem.AddItem "RANGES"
         Else
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "VALUE"
            cmbConfigItem.AddItem "REG"
            cmbConfigItem.AddItem "SLOPE"
            cmbConfigItem.AddItem "OFFSET"
         End If
         cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "AOSCAN"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "MAXSCANTHRUPUT"
            cmbConfigItem.AddItem "MINSCANRATE"
            cmbConfigItem.AddItem "MAXSCANRATE"
            cmbConfigItem.AddItem "EXTPACER" 'PACERSRC
            cmbConfigItem.AddItem "TRIG"  'EXTTRIG
            cmbConfigItem.AddItem "ADCLKTRIG"
            cmbConfigItem.AddItem "SIMUL"
            cmbConfigItem.AddItem "FIFOSIZE"
            cmbConfigItem.AddItem "XFRSIZE"
            cmbConfigItem.ListIndex = 0
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "RATE"
            cmbConfigItem.AddItem "SAMPLES"
            cmbConfigItem.AddItem "LOWCHAN"
            cmbConfigItem.AddItem "HIGHCHAN"
            cmbConfigItem.AddItem "BUFSIZE"
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "RANGE{ch}"
            cmbConfigItem.AddItem "RANGE{el/ch}"
            cmbConfigItem.AddItem "STATUS"
            cmbConfigItem.AddItem "COUNT"
            cmbConfigItem.AddItem "INDEX"
            cmbConfigItem.AddItem "SIMUL"
            cmbConfigItem.AddItem "TRIG"
            cmbConfigItem.AddItem "EXTPACER"
            cmbConfigItem.AddItem "XFRMODE"
            'cmbConfigItem.AddItem "EXTSYNC" 'obs
            cmbConfigItem.AddItem "SCALE"
            cmbConfigItem.AddItem "CAL"
            cmbConfigItem.AddItem "DEBUG"
            cmbConfigItem.AddItem "STALL"
            cmbConfigItem.AddItem "RESET"
            cmbConfigItem.AddItem "START"
            cmbConfigItem.AddItem "STOP"
            cmbConfigItem.ListIndex = 0
            mnDevList1 = True
            cmbDevNum.Visible = True
            cmbDevNum.Clear
            For i% = 0 To 15
               cmbDevNum.AddItem Format(i%, "0")
            Next
            cmbDevNum.ListIndex = 0
         End If
      Case "DIO"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "CHANNELS"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem ""
         End If
         cmbConfigItem.ListIndex = 0
         UpdateMsgStatus
      Case "DIO{port}"
         If mnFuncType = GET_SIGNAL Then
            cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "CONFIG"
            cmbConfigItem.AddItem "LATCH"
            cmbConfigItem.AddItem "FILTER"
            cmbConfigItem.AddItem "FILTTIME"
         Else
            cmbConfigItem.AddItem ""
            cmbConfigItem.AddItem "DIR"
            cmbConfigItem.AddItem "VALUE"
            cmbConfigItem.AddItem "LATCH"
            cmbConfigItem.AddItem "FILTER"
            cmbConfigItem.AddItem "FILTTIME"
         End If
         cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "DIO{port/bit}"
         mnDevList1 = True
         mnDevList2 = True
         cmbDevNum2.Visible = True
         txtConfigVal.Text = "0"
         cmbConfigItem.AddItem "VALUE"
         cmbConfigItem.AddItem "DIR"
         cmbConfigItem.AddItem "LATCH"
         cmbConfigItem.AddItem "FILTER"
         cmbConfigItem.AddItem "FILTTIME"
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
            cmbDevNum2.AddItem Format(i%, "0")
         Next
         cmbConfigItem.ListIndex = 0
         cmbDevNum2.ListIndex = 0
         cmbDevNum.ListIndex = 0
      Case "CTR"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "CHANNELS"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem ""
         End If
         cmbConfigItem.ListIndex = 0
         UpdateMsgStatus
      Case "CTR{ch}"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "TYPE"
            cmbConfigItem.AddItem "EDGE"
            cmbConfigItem.AddItem "LDMIN"
            cmbConfigItem.AddItem "LDMAX"
         Else
            cmbConfigItem.AddItem "VALUE"
            cmbConfigItem.AddItem "START"
            cmbConfigItem.AddItem "STOP"
         End If
         cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
      Case "TMR"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "CHANNELS"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem ""
         End If
         cmbConfigItem.ListIndex = 0
         UpdateMsgStatus
      Case "TMR{ch}"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            cmbConfigItem.AddItem "MAXCOUNT"
            cmbConfigItem.AddItem "CLKSRC"
            cmbConfigItem.AddItem "BASEFREQ"
            cmbConfigItem.AddItem "TYPE"
            cmbConfigItem.AddItem "DUTYCYCLE"
            cmbConfigItem.AddItem "DELAY"
            cmbConfigItem.AddItem "PULSECOUNT"
            cmbConfigItem.AddItem "IDLESTATE"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "PERIOD"
            cmbConfigItem.AddItem "DUTYCYCLE"
            cmbConfigItem.AddItem "DELAY"
            cmbConfigItem.AddItem "PULSECOUNT"
            cmbConfigItem.AddItem "IDLESTATE"
            cmbConfigItem.AddItem "START"
            cmbConfigItem.AddItem "STOP"
         End If
         cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
         UpdateMsgStatus
      Case "AICAL"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            'cmbConfigItem.AddItem "CHANNELS"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "DATETIME"
            cmbConfigItem.AddItem "DATETIME{YEAR}"
            cmbConfigItem.AddItem "DATETIME{MONTH}"
            cmbConfigItem.AddItem "DATETIME{DAY}"
            cmbConfigItem.AddItem "DATETIME{HOUR}"
            cmbConfigItem.AddItem "DATETIME{MINUTE}"
            cmbConfigItem.AddItem "DATETIME{SECOND}"
            cmbConfigItem.AddItem "LOCK"
            cmbConfigItem.AddItem "UNLOCK"
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "REF"
            cmbConfigItem.AddItem "REFVAL"
            cmbConfigItem.AddItem "MODE"
            cmbConfigItem.AddItem "RES"
         End If
         If cmbConfigItem.ListCount > 0 Then cmbConfigItem.ListIndex = 0
         UpdateMsgStatus
      Case "AICAL{ch}"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            'cmbConfigItem.AddItem "MAXCOUNT"
         Else
            cmbConfigItem.AddItem "RANGE"
            cmbConfigItem.AddItem "VALUE"
            cmbConfigItem.AddItem "SLOPE"
            cmbConfigItem.AddItem "OFFSET"
         End If
         If cmbConfigItem.ListCount > 0 Then cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
         UpdateMsgStatus
      Case "AOCAL"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            'cmbConfigItem.AddItem "CHANNELS"
         Else
            cmbConfigItem.AddItem "GetSupportedMsgs"
            cmbConfigItem.AddItem "LOCK"
            cmbConfigItem.AddItem "UNLOCK"
            cmbConfigItem.AddItem "RES"
            cmbConfigItem.AddItem "AIRES"
         End If
         If cmbConfigItem.ListCount > 0 Then cmbConfigItem.ListIndex = 0
         UpdateMsgStatus
      Case "AOCAL{ch}"
         If mnFuncType = GET_SIGNAL Then
            'load reflection messages
            'cmbConfigItem.AddItem "MAXCOUNT"
         Else
            cmbConfigItem.AddItem "VALUE"
            cmbConfigItem.AddItem "AIVALUE"
            cmbConfigItem.AddItem "AIRANGE"
            cmbConfigItem.AddItem "SLOPE"
            cmbConfigItem.AddItem "OFFSET"
            cmbConfigItem.AddItem "AISLOPE"
            cmbConfigItem.AddItem "AIOFFSET"
         End If
         If cmbConfigItem.ListCount > 0 Then cmbConfigItem.ListIndex = 0
         mnDevList1 = True
         cmbDevNum.Visible = True
         cmbDevNum.Clear
         For i% = 0 To 15
            cmbDevNum.AddItem Format(i%, "0")
         Next
         cmbDevNum.ListIndex = 0
         UpdateMsgStatus
   End Select
   cmbDevNum.ListIndex = 0

End Sub

Sub UpdateULStatus()

   If mnNumBoards = 0 Then Exit Sub
   
   If IsNumeric(cmbInfoType.Text) Then
      InfoType& = Val(cmbInfoType.Text)
      cmbInfoType.ToolTipText = GetCfgInfoTypeString(InfoType&)
   Else
      InfoType& = cmbInfoType.ItemData(cmbInfoType.ListIndex)
   End If
   
   ConfigString$ = cmbConfigItem.Text
   If IsNumeric(ConfigString$) Then
      ConfigItem& = Val(ConfigString$)
      ConfigString$ = cmbConfigItem.ToolTipText '= GetCfgItemString(ArgVal&)
   Else
      cmbConfigItem.ToolTipText = "ConfigItem"
      If cmbConfigItem.ListCount > 0 Then _
         ConfigItem& = cmbConfigItem.ItemData(cmbConfigItem.ListIndex)
   End If
   
   DevNum% = Val(cmbDevNum.Text)

   Select Case mnFuncType
      Case SET_CONF
         If (Not gnScriptRun) And chkHex.value Then
            ValConfig& = Val("&H" & txtConfigVal.Text)
         Else
            ValConfig& = Val(txtConfigVal.Text)
         End If
         Select Case ConfigString$
            Case "AInputMode"
               ULStat = cbAInputMode(mnBoardNum, ValConfig&)
               If SaveFunc(Me, AInputMode, ULStat, mnBoardNum, ValConfig&, A3, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            Case "AChanInputMode"
               ULStat = cbAChanInputMode(mnBoardNum, DevNum%, ValConfig&)
               If SaveFunc(Me, AChanInputMode, ULStat, mnBoardNum, DevNum%, ValConfig&, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
                  Prefix$ = "Error setting channel " & Format(DevNum%, "0")
                  Exit Sub
               Else
                  Prefix$ = "C"
               End If
            Case Else
               ULStat = SetConfig520(InfoType&, mnBoardNum, DevNum%, ConfigItem&, ValConfig&)
               If SaveFunc(Me, SetConfig, ULStat, InfoType&, mnBoardNum, DevNum%, _
               ConfigItem&, ValConfig&, A6, A7, A8, A9, A10, A11, 0) Then
                  If (ConfigString$ = "BIADCHANAIMODE") Then
                     Prefix$ = "Error setting c"
                  Else
                     Exit Sub
                  End If
               Else
                  If (ConfigString$ = "BIADCHANAIMODE") Then
                     Prefix$ = "C"
                  End If
               End If
         End Select
      Case SET_STRING
         ReturnString$ = txtConfigVal.Text
         If Not (frmComposite.txtShow.Text = "") Then ReturnString$ = frmComposite.txtShow.Text
         ConfigLen& = Len(ReturnString$)
         lblStatus.Caption = "Sent " & ConfigLen&
         ULStat = SetConfigString573(InfoType&, mnBoardNum, DevNum%, _
         ConfigItem&, ReturnString$, ConfigLen&)
         lblStatus.Caption = lblStatus.Caption & "  Returned " & ConfigLen&
         If SaveFunc(Me, SetConfigString, ULStat, InfoType&, mnBoardNum, DevNum%, _
         ConfigItem&, ReturnString$, ConfigLen&, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Unload frmComposite
         Exit Sub
   End Select
   
   If Not ConfigString$ = "" Then
      Select Case mnFuncType
         Case GET_CONF, SET_CONF
            If ConfigString$ = "AInputMode" Then
               ConfigItem& = BIADAIMODE
               ULStat = GetConfig520(InfoType&, mnBoardNum, DevNum%, ConfigItem&, ConfigVal&)
            ElseIf (ConfigString$ = "AChanInputMode") Or (ConfigString$ = "BIADCHANAIMODE") Then
               ULStat = GetConfig520(InfoType&, mnBoardNum, DevNum%, BIADCHANAIMODE, ConfigVal&)
               ConfigString$ = GetAiChanModeString(ConfigVal&)
               lblStatus.Caption = "Channel " & DevNum% & " mode is " & ConfigString$
               Exit Sub
            Else
               If ConfigItem& = GIINIT Then
                  NoReturnCfg% = True
                  'If ULStat = 0 Then
               Else
                  ULStat = GetConfig520(InfoType&, mnBoardNum, DevNum%, ConfigItem&, ConfigVal&)
                  'ConfigVal& = ValConfig&
                  If mnFuncType = GET_CONF Then
                     If SaveFunc(Me, GetConfig, ULStat, InfoType&, mnBoardNum, DevNum%, _
                     ConfigItem&, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then
                        'If ULStat = CFGFILENOTFOUND Then
                        '   Me.cmbInfoType.ENABLED = False
                        '   Exit Sub
                        'Else
                        '   Exit Sub
                        'End If
                     End If
                  Else
                     If Not (ULStat = 0) Then NoReturnCfg% = True
                  End If
               End If
            End If
         Case GET_STRING, SET_STRING
            StLen& = ERRSTRLEN   '64
            ReturnString$ = Space$(StLen&)
            ConfigLen& = StLen&
            lblStatus.Caption = "Sent " & ConfigLen&
            ULStat = GetConfigString573(InfoType&, mnBoardNum, DevNum%, _
            ConfigItem&, ReturnString$, ConfigLen&)
            lblStatus.Caption = lblStatus.Caption & "  Returned " & ConfigLen&
            ReturnString$ = Left$(ReturnString$, ConfigLen&)
            Select Case ConfigItem&
               Case BIDEVVERSION
                  Select Case DevNum%
                     Case VER_FW_MAIN
                        CItem$ = "Firmware version: "
                     Case VER_FW_MEASUREMENT
                        CItem$ = "Measurement firmware version: "
                     Case VER_FW_RADIO
                        CItem$ = "Radio firmware version: "
                     Case VER_FPGA
                        CItem$ = "FPGA version: "
                     Case VER_FW_MEASUREMENT_EXP
                        CItem$ = "Expansion firmware version: "
                     Case Else
                        CItem$ = ""
                  End Select
               Case Else
                  CItem$ = ""
            End Select
            If SaveFunc(Me, GetConfigString, ULStat, InfoType&, mnBoardNum, DevNum%, _
            ConfigItem&, ReturnString$, ConfigLen&, A7, A8, A9, A10, A11, 0) Then Exit Sub
            SetPlotType PRINT_LIST, Me
            frmPlot.txtShow.Text = ""
            TextList CItem$ & ReturnString$
            Exit Sub
      End Select
      
      If InfoType& = BOARDINFO Then
         Select Case ConfigItem&
            Case BINUMADCHANS
               mnADChans = ConfigVal&
               LoadDevs% = True
            Case BIDINUMDEVS
               mnDigDevs = ConfigVal&
               LoadDevs% = True
            Case BICINUMDEVS
               mnCtrDevs = ConfigVal&
               LoadDevs% = True
            Case BINUMDACHANS
               mnDAChans = ConfigVal&
               LoadDevs% = True
            Case BINUMIOPORTS
               mnIOPorts = ConfigVal&
               LoadDevs% = True
            Case BIADSCANOPTIONS, BIDACSCANOPTIONS, BIDISCANOPTIONS, _
                 BIDOSCANOPTIONS, BICTRSCANOPTIONS, BIDAQISCANOPTIONS, _
                 BIDAQOSCANOPTIONS
               FormType% = ANALOG_IN
               If ConfigItem& = BICTRSCANOPTIONS Then FormType% = COUNTERS
               If (ConfigItem& = BIDAQISCANOPTIONS) And (InStr(1, msBoardName, "CTR") > 1) Then FormType% = COUNTERS
               OptString$ = GetOptionsString(ConfigVal&, FormType%)
               OptString$ = Replace(OptString$, " ", vbCrLf)
               SetPlotType PRINT_LIST, Me
               frmPlot.txtShow.Text = ""
               TextList msBoardName & " " & Me.cmbConfigItem.Text & _
                  vbCrLf & vbCrLf & OptString$
         End Select
      End If
   Else
      lblStatus.Caption = ConfigString$
      Exit Sub
   End If
   
   If Not ((mnFuncType = SET_STRING) Or (mnFuncType = GET_STRING)) Then
      If Not NoReturnCfg% Then
         If chkHex.value Then
            lblStatus.Caption = ConfigString$ & " = 0x" & Hex$(ConfigVal&)
         Else
            Select Case ConfigItem&
               Case BICHANTCTYPE
                  ConfigValString$ = GetTcTypeString(ConfigVal&)
               Case BIADCHANTYPE
                  ConfigValString$ = GetAiChanTypeString(ConfigVal&)
               Case BIADAIMODE, BIADCHANAIMODE
                  ConfigValString$ = GetAiChanModeString(ConfigVal&)
               Case Else
                  ConfigValString$ = Format(ConfigVal&, "0")
            End Select
            lblStatus.Caption = ConfigString$ & " = " & ConfigValString$
         End If
      Else
         lblStatus.Caption = ConfigString$ & " Invalid item for GetConfig"
      End If
   End If
   
End Sub

Sub UpdateMsgStatus()

   If Not mnMessaging Then Exit Sub
   ComponentSelected$ = cmbInfoType.Text
   Component$ = ComponentSelected$
   Prop$ = cmbConfigItem.Text
   
   If Not Prop$ = "" Then Prop$ = ":" & Prop$
   If cmbConfigItem.Text = "GetSupportedMsgs" Then
      cmbList.Visible = True
      lblStatus.Visible = False
      x$ = MsgLibrary.GetSupportedMessages(Component$)
      Reslt% = SaveMsg(Me, "GetSupportedMessages(" & Component$ & ")", x$)
      MsgsSupported = Split(x$, "|")
      cmbList.Clear
      For MsgNum& = 0 To UBound(MsgsSupported) - 1
         cmbList.AddItem MsgsSupported(MsgNum&)
      Next
      If cmbList.ListCount > 0 Then cmbList.ListIndex = 0
      Exit Sub
   Else
      cmbList.Visible = False
      lblStatus.Visible = True
   End If
   Select Case ComponentSelected$
      Case "DEV"
         ItemText$ = cmbConfigItem.Text
         Select Case ItemText$
            Case "DATATYPE"
            Case "FLASHLED"
               AltForm$ = "/" & txtConfigVal.Text
               If txtConfigVal.Text = "" Then AltForm$ = ""
               DontRead% = True
               UseAltFormat% = True
               If Not (cmbDevNum.Text = "") Then
                  AltParam$ = "{" & cmbDevNum.Text & "}"
               End If
            Case "RESET"
               AltForm$ = "/" & txtConfigVal.Text
               DontRead% = True
               UseAltFormat% = True
               If Not (cmbDevNum.Text = "") Then
                  AltParam$ = "{" & cmbDevNum.Text & "}"
               End If
            Case "TEMP{ch}"
               If cmbDevNum.Text = "" Then Exit Sub
               Prop$ = ":TEMP{" & cmbDevNum.Text & "}"
         End Select
      Case "AI"
      Case "AI{ch}"
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "AI{" & cmbDevNum.Text & "}"
      Case "AIQUEUE"
         UseAltFormat% = True
         If Prop = ":CLEAR" Then DontRead% = True
      Case "AIQUEUE{el}"
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "AIQUEUE{" & cmbDevNum.Text & "}"
      Case "AO"
         ItemText$ = cmbConfigItem.Text
         Select Case ItemText$
            Case "UPDATE"
               DontRead% = True
               UseAltFormat% = True
         End Select
      Case "AO{ch}"
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "AO{" & cmbDevNum.Text & "}"
      Case "DIO"
      Case "DIO{port}"
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "DIO{" & cmbDevNum.Text & "}"
      Case "DIO{port/bit}"
         'to do - add bit input
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "DIO{" & cmbDevNum.Text & "/" & cmbDevNum2.Text & "}"
      Case "AISCAN"
         Select Case Prop$
            Case ":RANGE{ch}"
               If cmbDevNum.Text = "" Then Exit Sub
               Prop$ = ":RANGE{" & cmbDevNum.Text & "}"
            Case ":RANGE{el/ch}"
               If (cmbDevNum.Text = "") Or (cmbDevNum2.Text = "") Then Exit Sub
               Prop$ = ":RANGE{" & cmbDevNum.Text & "/" & cmbDevNum2.Text & "}"
               UseResponse% = True
         End Select
      Case "AITRIG"
      Case "AOSCAN"
         Select Case Prop$
            Case ":RANGE{ch}"
               If cmbDevNum.Text = "" Then Exit Sub
               Prop$ = ":RANGE{" & cmbDevNum.Text & "}"
            Case ":RANGE{el/ch}"
               If (cmbDevNum.Text = "") Or (cmbDevNum2.Text = "") Then Exit Sub
               Prop$ = ":RANGE{" & cmbDevNum.Text & "/" & cmbDevNum2.Text & "}"
               UseResponse% = True
         End Select
      Case "CTR"
      Case "CTR{ch}"
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "CTR{" & cmbDevNum.Text & "}"
      Case "TMR"
      Case "TMR{ch}"
         If cmbDevNum.Text = "" Then Exit Sub
         Component$ = "TMR{" & cmbDevNum.Text & "}"
         ItemText$ = cmbConfigItem.Text
         Select Case ItemText$
            Case "START", "STOP"
               DontRead% = True
               UseAltFormat% = True
         End Select
   End Select
   
   Select Case mnFuncType
      Case GET_CONF
         CfgMessage$ = "?" & Component$ & Prop$ & AltParam$
         MBDResponse$ = MsgLibrary.SendMessage(CfgMessage$)
         If Not SaveMsg(Me, "SendMessage(" & CfgMessage$ & ")", MBDResponse$) Then
            lblStatus.Caption = MBDResponse$
         End If
      Case SET_CONF
         If UseAltFormat% Then
            SetCfgMessage$ = Component$ & Prop$ & AltForm$ '& "=" & txtConfigVal.Text
         Else
            SetText$ = ""
            If txtConfigVal.Text <> "" Then SetText$ = "="
            SetCfgMessage$ = Component$ & Prop$ & AltParam$ & SetText$ & txtConfigVal.Text
         End If
         MBDResponse$ = MsgLibrary.SendMessage(SetCfgMessage$)
         Error% = SaveMsg(Me, "SendMessage(" & SetCfgMessage$ & ")", MBDResponse$)
         DontRead% = (DontRead% Or (chkReQuery.value = 0))
         If Not (UseResponse% Or DontRead%) Then
            If Not UseAltFormat% Then
               CfgMessage$ = "?" & Component$ & Prop$
            Else
               CfgMessage$ = "?" & Component$ & Prop$ & AltParam$
            End If
            MBDResponse$ = MsgLibrary.SendMessage(CfgMessage$)
            Error% = SaveMsg(Me, "SendMessage(" & CfgMessage$ & ")", MBDResponse$)
         End If
         lblStatus.Caption = MBDResponse$
         If Not Error% Then
            UpdateAssocForms Component$, Prop$
         End If
      Case GET_SIGNAL
         CfgMessage$ = "@" & Component$ & Prop$ & AltParam$
         If Me.chkHex.value = 0 Then
            MBDResponse$ = MsgLibrary.SendMessage(CfgMessage$)
            If Not SaveMsg(Me, "SendMessage(" & CfgMessage$ & ")", MBDResponse$) Then
               lblStatus.Caption = MBDResponse$
            End If
         Else
            NumResponse = MsgLibrary.SendMessageN(CfgMessage$)
            lblStatus.Caption = NumResponse
         End If
   End Select
   
End Sub

Private Sub UpdateAssocForms(Component As String, Prop As String)

   ReleventForms% = -1
   CompArray = Split(Component, "{")
   CompType$ = CompArray(0)
   Select Case CompType$
      Case "AISCAN", "AI"
         ReleventForms% = GetFormsOfType("100", FormArray)
      Case "DEV"
         ReleventForms% = GetFormsOfType("-1", FormArray)
   End Select
   For i% = 0 To ReleventForms%
      FormIndex% = FormArray(i%)
      formTitle$ = Forms(FormIndex%).Caption
      If InStr(1, formTitle$, msDisplayName) > 0 Then
         If (Component = "DEV") And (Prop = ":RESET") Then
            FormTag$ = Left(Forms(FormIndex%).Tag, 1)
            Select Case FormTag$
               Case "1"
                  Forms(FormIndex%).ConfigurationChange ":CHMODE"
                  Forms(FormIndex%).ConfigurationChange ":EXTPACER"
                  Forms(FormIndex%).ConfigurationChange ":LOWCHAN"
                  Forms(FormIndex%).ConfigurationChange ":HIGHCHAN"
                  Forms(FormIndex%).ConfigurationChange ":RATE"
                  Forms(FormIndex%).ConfigurationChange ":SAMPLES"
                  Forms(FormIndex%).ConfigurationChange ":TRIG"
                  Forms(FormIndex%).ConfigurationChange ":QUEUE"
                  'Forms(FormIndex%).ConfigurationChange ":RANGE{"
                  Forms(FormIndex%).ConfigurationChange ":XFRMODE"
               Case "5"
                  Forms(FormIndex%).ConfigurationChange "TMR"
            End Select
         Else
            Forms(FormIndex%).ConfigurationChange Prop
         End If
      End If
   Next
   
End Sub

Public Sub InitForm(FunctionInit As Integer)

   mnuBoard_Click (0)
   mnuFuncArray_Click (FunctionInit)
   If gnLibType = MSGLIB Then
      If Not mnMessaging Then Exit Sub
   End If
   'If Not (gnNumBoards > 0) Then
   '   For Ctl& = 0 To Me.Controls.Count - 1
   '      ContType$ = Left(Controls(Ctl&).Name, 3)
   '      If (ContType$ = "cmd") Or _
   '         (ContType$ = "chk") Or _
   '         (ContType$ = "cmb") Then _
   '         Controls(Ctl&).ENABLED = False
   '   Next
   '   Exit Sub
   'End If
   
End Sub

Public Sub SetConfigValues(ByVal ListItem As Integer, _
ByVal ConfigItem As String, ByVal ConfigVal As String)

   cmbInfoType.ListIndex = ListItem
   cmbConfigItem.Text = ConfigItem
   txtConfigVal.Text = ConfigVal

End Sub

Public Function GetMsgDevice() As String

   BoardName$ = ""
   If Not MsgLibrary Is Nothing Then
      BoardName$ = msBoardName
   End If
   GetMsgDevice = BoardName$
   
End Function

Public Sub ConfigurationChange(ConfigType As String)

   ParseType$ = Mid(ConfigType, 2, 6)
   'TypeSupported% = (InStr(1, msAiSupport, ParseType$) > 0)
   'TypeSupported% = TypeSupported% Or (InStr(1, msScanSupport, ParseType$) > 0)
   'If TypeSupported% Then
   Select Case ParseType$
      Case "CHMODE"
         QueryMsg = "?AI" ' & ConfigType
         MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
         If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
         ParseMsg = Split(MsgResult$, "=")
         If UBound(ParseMsg) > 0 Then NewChanCount$ = ParseMsg(1)
         mnNumAIChans = Val(NewChanCount$)
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
   End Select
   
End Sub

Private Function GetConfigList(ByVal ValType As Long, ByVal FILTER As Long, _
   ByRef CfgVals() As Long, ByRef CfgNames() As String) As Long

   
End Function

Private Function GetInfoTypes(ByVal Index As Integer, ByVal CfgFilter As Integer, _
   ByRef CfgVal As Long, ByRef CfgName As String) As Boolean
   
   Dim NotDone As Boolean
   
   NotDone = True
   Select Case Index
      Case 0
         CfgName = "GLOBALINFO"
         CfgVal = GLOBALINFO
      Case 1
         CfgName = "BOARDINFO"
         CfgVal = BOARDINFO
      Case 2
         CfgName = "DIGITALINFO"
         CfgVal = DIGITALINFO
      Case 3
         CfgName = "COUNTERINFO"
         CfgVal = COUNTERINFO
      Case 4
         CfgName = "EXPANSIONINFO"
         CfgVal = EXPANSIONINFO
         NotDone = False
   End Select
   GetInfoTypes = NotDone
   
End Function

Private Function GetCfgItems(ByVal CfgFilter As Integer, _
   ByRef CfgVal() As Variant, ByRef CfgName() As Variant) As Long
   
   Dim StdFeatures As Boolean, ADFeatures As Boolean
   Dim DAFeatures As Boolean, DCtrFeatures As Boolean
   Dim TCFeatures As Boolean, LegacyFeatures As Boolean
      
   StdFeatures = ((CfgFilter And LISTSTD) = LISTSTD)
   ADFeatures = ((CfgFilter And LISTAD) = LISTAD)
   DAFeatures = ((CfgFilter And LISTDA) = LISTDA)
   DCtrFeatures = ((CfgFilter And LISTIDCTR) = LISTIDCTR)
   TCFeatures = ((CfgFilter And LISTTC) = LISTTC)
   LegacyFeatures = ((CfgFilter And LISTLEGACY) = LISTLEGACY)
   
   InfoType& = cmbInfoType.ItemData(cmbInfoType.ListIndex)
   Select Case InfoType&
      Case GLOBALINFO
         CfgName = Array("GIVERSION", "GIINIT", "GINUMBOARDS", "GINUMEXPBOARDS")
         CfgVal = Array(GIVERSION, GIINIT, GINUMBOARDS, GINUMEXPBOARDS)
      Case BOARDINFO
         If StdFeatures Then
            If (mnFuncType = GET_CONF) Then
               CfgName = Array("BIBOARDTYPE", "BIINTLEVEL", "BIINTEDGE", "BICLOCK", _
                              "BINUMADCHANS", "BINUMTEMPCHANS", "BIUSESEXPS", _
                              "BIDINUMDEVS", "BICINUMDEVS", "BINUMDACHANS", "BINUMEXPS", _
                              "BIINPUTPACEROUT", "BIOUTPUTPACEROUT", "BIEXTINPACEREDGE", _
                              "BIEXTOUTPACEREDGE", "BIEXTCLKTYPE", "BINETCONNECTCODE", _
                              "BINETIOTIMEOUT", "BIUSERDEVIDNUM", "BICALTABLETYPE", _
                              "BIDAQISCANOPTIONS", "BIDAQOSCANOPTIONS", "BIDAQITRIGCOUNT", _
                              "BIDAQINUMCHANTYPES", "BIDAQICHANTYPE", "BIDAQONUMCHANTYPES", "BIDAQOCHANTYPE", _
                              "BICTRZACTIVEMODE", "BIDAQAMISUPPORTED", "BIHASEXTINFO", "BINUMIODEVS", "BIIODEVTYPE")
               CfgVal = Array(BIBOARDTYPE, BIINTLEVEL, BIINTEDGE, BICLOCK, _
                              BINUMADCHANS, BINUMTEMPCHANS, BIUSESEXPS, _
                              BIDINUMDEVS, BICINUMDEVS, BINUMDACHANS, BINUMEXPS, _
                              BIINPUTPACEROUT, BIOUTPUTPACEROUT, BIEXTINPACEREDGE, _
                              BIEXTOUTPACEREDGE, BIEXTCLKTYPE, BINETCONNECTCODE, _
                              BINETIOTIMEOUT, BIUSERDEVIDNUM, BICALTABLETYPE, _
                              BIDAQISCANOPTIONS, BIDAQOSCANOPTIONS, BIDAQITRIGCOUNT, _
                              BIDAQINUMCHANTYPES, BIDAQICHANTYPE, BIDAQONUMCHANTYPES, BIDAQOCHANTYPE, _
                              BICTRZACTIVEMODE, BIDAQAMISUPPORTED, BIHASEXTINFO, BINUMIODEVS, BIIODEVTYPE)
            End If
            If (mnFuncType = SET_CONF) Then
               CfgName = Array("BIINTLEVEL", "BIINTEDGE", "BICLOCK", _
                              "BIINPUTPACEROUT", "BIOUTPUTPACEROUT", _
                              "BIEXTINPACEREDGE", "BIEXTOUTPACEREDGE", _
                              "BIEXTCLKTYPE", "BINETCONNECTCODE", _
                              "BICALTABLETYPE", "BIDAQITRIGCOUNT", "BICTRZACTIVEMODE")
               CfgVal = Array(BIINTLEVEL, BIINTEDGE, BICLOCK, _
                              BIINPUTPACEROUT, BIOUTPUTPACEROUT, _
                              BIEXTINPACEREDGE, BIEXTOUTPACEREDGE, _
                              BIEXTCLKTYPE, BINETCONNECTCODE, _
                              BICALTABLETYPE, BIDAQITRIGCOUNT, BICTRZACTIVEMODE)
            End If
            If (mnFuncType = GET_STRING) Then
               CfgName = Array("BIDEVUNIQUEID", "BIDEVSERIALNUM", "BIDEVMACADDR", _
                              "BIDEVVERSION", "BIDEVNOTES", "BIUSERDEVID", _
                              "BINETBIOSNAME", "BIDEVIPADDR")
               CfgVal = Array(BIDEVUNIQUEID, BIDEVSERIALNUM, BIDEVMACADDR, _
                              BIDEVVERSION, BIDEVNOTES, BIUSERDEVID, _
                              BINETBIOSNAME, BIDEVIPADDR)
            End If
            If (mnFuncType = SET_STRING) Then
               CfgName = Array("BIDEVNOTES", "BIUSERDEVID")
               CfgVal = Array(BIDEVNOTES, BIUSERDEVID)
            End If
         End If
         If ADFeatures Then
            If (mnFuncType = GET_CONF) Then
               CfgName = Array("BIADAIMODE", "BINUMADCHANS", "BIRANGE", "BIADTRIGCOUNT", _
                              "BIADRES", "BIADTRIGSRC", "BIADXFERMODE", "BIADCSETTLETIME", _
                              "BIADTIMINGMODE", "BIADCHANTYPE", "BIADCHANAIMODE", _
                              "BIADDATARATE", "BIADSCANOPTIONS", "BIADNUMCHANMODES", _
                              "BIADCHANMODE", "BIADNUMDIFFRANGES", "BIADDIFFRANGE", _
                              "BIADNUMSERANGES", "BIADSERANGE", "BIADNUMTRIGTYPES", _
                              "BIADTRIGTYPE", "BIADMAXRATE", "BIADMAXTHROUGHPUT", _
                              "BIADMAXBURSTRATE", "BIADMAXBURSTTHROUGHPUT", "BIADHASPACER", _
                              "BIADCHANTYPES", "BIADMAXSEQUEUELENGTH", "BIADMAXDIFFQUEUELENGTH", _
                              "BIADQUEUETYPES", "BIADQUEUELIMITS")
               CfgVal = Array(BIADAIMODE, BINUMADCHANS, BIRANGE, BIADTRIGCOUNT, _
                              BIADRES, BIADTRIGSRC, BIADXFERMODE, BIADCSETTLETIME, _
                              BIADTIMINGMODE, BIADCHANTYPE, BIADCHANAIMODE, _
                              BIADDATARATE, BIADSCANOPTIONS, BIADNUMCHANMODES, _
                              BIADCHANMODE, BIADNUMDIFFRANGES, BIADDIFFRANGE, _
                              BIADNUMSERANGES, BIADSERANGE, BIADNUMTRIGTYPES, _
                              BIADTRIGTYPE, BIADMAXRATE, BIADMAXTHROUGHPUT, _
                              BIADMAXBURSTRATE, BIADMAXBURSTTHROUGHPUT, BIADHASPACER, _
                              BIADCHANTYPES, BIADMAXSEQUEUELENGTH, BIADMAXDIFFQUEUELENGTH, _
                              BIADQUEUETYPES, BIADQUEUELIMITS)
            End If
            If (mnFuncType = SET_CONF) Then
               CfgName = Array("AInputMode", "AChanInputMode", "BIRANGE", "BIADTRIGCOUNT", _
                              "BIADTRIGSRC", "BIADXFERMODE", "BIADCSETTLETIME", _
                              "BIADTIMINGMODE", "BIADCHANTYPE", "BIADDATARATE", _
                              "BIADAIMODE", "BIADCHANAIMODE")
               CfgVal = Array(SINGLE_ENDED, DIFFERENTIAL, BIRANGE, BIADTRIGCOUNT, _
                              BIADTRIGSRC, BIADXFERMODE, BIADCSETTLETIME, _
                              BIADTIMINGMODE, BIADCHANTYPE, BIADDATARATE, _
                              BIADAIMODE, BIADCHANAIMODE)
            End If
         End If
         If DAFeatures Then
            If (mnFuncType = GET_CONF) Then
               CfgName = Array("BINUMDACHANS", "BIDACRES", "BISYNCMODE", "BIDACUPDATEMODE", _
                              "BIDACSTARTUP", "BIDACTRIGCOUNT", "BIDACFORCESENSE", _
                              "BIDACRANGE", "BIDACSCANOPTIONS", "BIDACHASPACER", _
                              "BIDACFIFOSIZE", "BIDACNUMRANGES", "BIDACDEVRANGE", _
                              "BIDACNUMTRIGTYPES", "BIDACTRIGTYPE")
               CfgVal = Array(BINUMDACHANS, BIDACRES, BISYNCMODE, BIDACUPDATEMODE, _
                              BIDACSTARTUP, BIDACTRIGCOUNT, BIDACFORCESENSE, _
                              BIDACRANGE, BIDACSCANOPTIONS, BIDACHASPACER, _
                              BIDACFIFOSIZE, BIDACNUMRANGES, BIDACDEVRANGE, _
                              BIDACNUMTRIGTYPES, BIDACTRIGTYPE)
            End If
            If (mnFuncType = SET_CONF) Then
               CfgName = Array("BISYNCMODE", "BIDACUPDATEMODE", _
                              "BIDACSTARTUP", "BIDACTRIGCOUNT", _
                              "BIDACFORCESENSE", "BIDACRANGE", "BIDACUPDATECMD")
               CfgVal = Array(BISYNCMODE, BIDACUPDATEMODE, _
                              BIDACSTARTUP, BIDACTRIGCOUNT, _
                              BIDACFORCESENSE, BIDACRANGE, BIDACUPDATECMD)
            End If
         End If
         If DCtrFeatures Then
            If (mnFuncType = GET_CONF) Then
               CfgName = Array("BIDINUMDEVS", "BICINUMDEVS", "BICTRTRIGCOUNT", "BIDITRIGCOUNT", _
                              "BIDOTRIGCOUNT", "BIPATTERNTRIGPORT", "BITERMCOUNTSTATBIT", _
                              "BIDISOFILTER", "BIDISCANOPTIONS", "BIDOSCANOPTIONS", _
                              "BICTRSCANOPTIONS", "BIRELAYLOGIC", "BIOPENRELAYLEVEL")
               CfgVal = Array(BIDINUMDEVS, BICINUMDEVS, BICTRTRIGCOUNT, BIDITRIGCOUNT, _
                              BIDOTRIGCOUNT, BIPATTERNTRIGPORT, BITERMCOUNTSTATBIT, _
                              BIDISOFILTER, BIDISCANOPTIONS, BIDOSCANOPTIONS, _
                              BICTRSCANOPTIONS, BIRELAYLOGIC, BIOPENRELAYLEVEL)
            End If
            If (mnFuncType = SET_CONF) Then
               CfgName = Array("BICTRTRIGCOUNT", "BIDITRIGCOUNT", "BIDOTRIGCOUNT", _
                              "BIPATTERNTRIGPORT", "BITERMCOUNTSTATBIT", _
                              "BIDISOFILTER", "BIOPENRELAYLEVEL")
               CfgVal = Array(BICTRTRIGCOUNT, BIDITRIGCOUNT, BIDOTRIGCOUNT, _
                              BIPATTERNTRIGPORT, BITERMCOUNTSTATBIT, _
                              BIDISOFILTER, BIOPENRELAYLEVEL)
            End If
         End If
         If TCFeatures Then
            If (mnFuncType = GET_CONF) Then
               CfgName = Array("BINUMTEMPCHANS", "BICHANTCTYPE", "BITEMPREJFREQ", _
                              "BITEMPAVG", "BIEXCITATION", "BICHANBRIDGETYPE", "BITEMPSCALE", _
                              "BICHANRTDTYPE", "BIDETECTOPENTC")
               CfgVal = Array(BINUMTEMPCHANS, BICHANTCTYPE, BITEMPREJFREQ, _
                              BITEMPAVG, BIEXCITATION, BICHANBRIDGETYPE, BITEMPSCALE, _
                              BICHANRTDTYPE, BIDETECTOPENTC)
            End If
            If (mnFuncType = SET_CONF) Then
               CfgName = Array("BICHANTCTYPE", "BITEMPREJFREQ", _
                              "BITEMPAVG", "BIEXCITATION", "BICHANBRIDGETYPE", "BITEMPSCALE", _
                              "BICHANRTDTYPE", "BIDETECTOPENTC")
               CfgVal = Array(BICHANTCTYPE, BITEMPREJFREQ, _
                              BITEMPAVG, BIEXCITATION, BICHANBRIDGETYPE, BITEMPSCALE, _
                              BICHANRTDTYPE, BIDETECTOPENTC)
            End If
         End If
         If LegacyFeatures Then
            If (mnFuncType = GET_CONF) Then
               CfgName = Array("BIBASEADR", "BIDMACHAN", "BIWAITSTATE", "BINUMIOPORTS", _
                              "BIDTBOARD", "BIPANID", "BIRFCHANNEL", "BIRSS")
               CfgVal = Array(BIBASEADR, BIDMACHAN, BIWAITSTATE, BINUMIOPORTS, _
                              BIDTBOARD, BIPANID, BIRFCHANNEL, BIRSS)
            End If
            If (mnFuncType = SET_CONF) Then
               CfgName = Array("BIBASEADR", "BINUMADCHANS", "BIDMACHAN", _
                              "BIWAITSTATE", "BIPANID", "BIRFCHANNEL")
               CfgVal = Array(BIBASEADR, BINUMADCHANS, BIDMACHAN, _
                              BIWAITSTATE, BIPANID, BIRFCHANNEL)
            End If
            If (mnFuncType = GET_STRING) Then
               CfgName = Array("BIMFGSERIALNUM", "BISERIALNUM", "BINODEID")
               CfgVal = Array(BIMFGSERIALNUM, BISERIALNUM, BINODEID)
            End If
            If (mnFuncType = SET_STRING) Then
               CfgName = Array("BIMFGSERIALNUM", "BISERIALNUM", "BINODEID")
               CfgVal = Array(BIMFGSERIALNUM, BISERIALNUM, BINODEID)
            End If
         End If
      Case DIGITALINFO
         If StdFeatures Then
            CfgName = Array("DIDEVTYPE", "DINUMBITS", "DIINMASK", "DIOUTMASK", _
                           "DIDISABLEDIRCHECK", "DICONFIG", "DICURVAL")
            CfgVal = Array(DIDEVTYPE, DINUMBITS, DIINMASK, DIOUTMASK, _
                           DIDISABLEDIRCHECK, DICONFIG, DICURVAL)
         End If
         If LegacyFeatures Then
            CfgName = Array("DIBASEADR", "DIINITIALIZED", "DIMASK", "DIREADWRITE")
            CfgVal = Array(DIBASEADR, DIINITIALIZED, DIMASK, DIREADWRITE)
         End If
      Case COUNTERINFO
         If StdFeatures Then
            CfgName = Array("CICTRTYPE", "CICTRNUM")
            CfgVal = Array(CICTRTYPE, CICTRNUM)
         End If
         If LegacyFeatures Then
            CfgName = Array("CIBASEADR", "CIINITIALIZED", "CICONFIGBYTE")
            CfgVal = Array(CIBASEADR, CIINITIALIZED, CICONFIGBYTE)
         End If
      Case EXPANSIONINFO
         If StdFeatures Then
            CfgName = Array("XIBOARDTYPE", "XIMUXADCHAN1", "XIMUXADCHAN2", _
                           "XIRANGE1", "XIRANGE2", "XICJCCHAN", "XITHERMTYPE", _
                           "XINUMEXPCHANS", "XIPARENTBOARD")
            CfgVal = Array(XIBOARDTYPE, XIMUXADCHAN1, XIMUXADCHAN2, _
                           XIRANGE1, XIRANGE2, XICJCCHAN, XITHERMTYPE, _
                           XINUMEXPCHANS, XIPARENTBOARD)
         End If
         If LegacyFeatures Then
            CfgName = Array("XISPARE0")
            CfgVal = Array(XISPARE0)
         End If
   End Select
   
   ArraySize& = -1
   If CfgFilter > 0 Then ArraySize& = UBound(CfgVal)
   GetCfgItems = ArraySize&
   
End Function

Private Function GetCfgItemsEx(ByVal Index As Integer, ByVal CfgFilter As Integer, _
   ByRef CfgVal As Long, ByRef CfgName As String) As Boolean

   'obsolete
   MsgBox "This function (GetCfgItemsEx) shouldn't be called", _
      vbInformation, "Obsolete Function"
   Dim NotDone As Boolean, ADFeatures As Boolean
   Dim DAFeatures As Boolean
   Dim LegacyFeatures As Boolean
   
   NotDone = True
   If CfgFilter = -1 Then
      Select Case Index
         Case 0
            CfgName = "VER_FW_MAIN"
            CfgVal = VER_FW_MAIN
         Case 1
            CfgName = "VER_FW_ISOLATED"
            CfgVal = VER_FW_ISOLATED
         Case 2
            CfgName = "VER_FW_RADIO"
            CfgVal = VER_FW_RADIO
         Case 3
            CfgName = "VER_FPGA"
            CfgVal = VER_FPGA
            NotDone = False
      End Select
   Else
      ADFeatures = ((CfgFilter And LISTAD) = LISTAD)
      DAFeatures = ((CfgFilter And LISTDA) = LISTDA)
      IcalFeatures = ((CfgFilter And LISTICAL) = LISTICAL)
      TCFeatures = ((CfgFilter And LISTTC) = LISTTC)
      LegacyFeatures = ((CfgFilter And LISTLEGACY) = LISTLEGACY)
      
      InfoType& = cmbInfoType.ItemData(cmbInfoType.ListIndex)
      Select Case InfoType&
         Case GLOBALINFO
            Select Case Index
               Case 0
                  CfgName = "GIVERSION"
                  CfgVal = GIVERSION
               Case 1
                  CfgName = "GINUMBOARDS"
                  CfgVal = GINUMBOARDS
               Case 2
                  CfgName = "GINUMEXPBOARDS"
                  CfgVal = GINUMEXPBOARDS
                  NotDone = False
            End Select
         Case BOARDINFO
            Select Case Index
               Case 0
                  CfgName = "BIBOARDTYPE"
                  CfgVal = BIBOARDTYPE
               Case 1
                  CfgName = "BIFWVERSION"
                  CfgVal = BIFWVERSION
               Case 2
                  CfgName = "BISERIALNUM"
                  CfgVal = BISERIALNUM
               Case 3
                  CfgName = "BINUMADCHANS"
                  CfgVal = BINUMADCHANS
               Case 4
                  CfgName = "BIDINUMDEVS"
                  CfgVal = BIDINUMDEVS
               Case 5
                  CfgName = "BICINUMDEVS"
                  CfgVal = BICINUMDEVS
               Case 6
                  CfgName = "BINUMDACHANS"
                  CfgVal = BINUMDACHANS
                  If CfgFilter = 0 Then NotDone = False
               Case Else
                  If CfgVal = 0 Then CfgVal = -1
            End Select
            If ADFeatures Then
               Select Case Index
                  Case 7
                     CfgName = "BIADRES"
                     CfgVal = BIADRES
                  Case 8
                     CfgName = "BIADTRIGCOUNT"
                     CfgVal = BIADTRIGCOUNT
                  Case 9
                     CfgName = "BIADFIFOSIZE"
                     CfgVal = BIADFIFOSIZE
                  Case 10
                     CfgName = "BIADSOURCE"
                     CfgVal = BIADSOURCE
                  Case 11
                     CfgName = "BIADTRIGSRC"
                     CfgVal = BIADTRIGSRC
                  Case 12
                     CfgName = "BIADXFERMODE"
                     CfgVal = BIADXFERMODE
                  Case 13
                     CfgName = "BIADCHANTYPE"
                     CfgVal = BIADCHANTYPE
                  Case 14
                     CfgName = "BIAIWAVETYPE"
                     CfgVal = BIAIWAVETYPE
                  Case 15
                     CfgName = "BITEMPREJFREQ"
                     CfgVal = BITEMPREJFREQ
                     FilterMask& = LISTDA Or LISTICAL _
                        Or LISTTC Or LISTLEGACY
                     If ((CfgFilter And FilterMask&) = 0) _
                        Then NotDone = False
                  Case Else
                     If CfgVal = 0 Then CfgVal = -1
               End Select
            End If
            If DAFeatures Then
               Select Case Index
                  Case 16
                     CfgName = "BIDACRANGE"
                     CfgVal = BIDACRANGE
                  Case 17
                     CfgName = "BIDACRES"
                     CfgVal = BIDACRES
                  Case 18
                     CfgName = "BIDACTRIGCOUNT"
                     CfgVal = BIDACTRIGCOUNT
                  Case 19
                     CfgName = "BIDACUPDATEMODE"
                     CfgVal = BIDACUPDATEMODE
                  Case 20
                     CfgName = "BIDACUPDATECMD"
                     CfgVal = BIDACUPDATECMD
                  Case 21
                     CfgName = "BIDACSTARTUP"
                     CfgVal = BIDACSTARTUP
                     FilterMask& = LISTICAL _
                        Or LISTTC Or LISTLEGACY
                     If ((CfgFilter And FilterMask&) = 0) _
                        Then NotDone = False
                  Case Else
                     If CfgVal = 0 Then CfgVal = -1
               End Select
            End If
            If IcalFeatures Then
               Select Case Index
                  Case 22
                     CfgName = "BINETCONNECTCODE"
                     CfgVal = BINETCONNECTCODE
                  Case 23
                     CfgName = "BINETIOTIMEOUT"
                     CfgVal = BINETIOTIMEOUT
                  Case 24
                     CfgName = "BIDISCONNECT"
                     CfgVal = BIDISCONNECT
                  Case 25
                     CfgName = "BIDIALARMMASK"
                     CfgVal = BIDIALARMMASK
                  Case 26
                     CfgName = "BIDITRIGCOUNT"
                     CfgVal = BIDITRIGCOUNT
                  Case 27
                     CfgName = "BIDOTRIGCOUNT"
                     CfgVal = BIDOTRIGCOUNT
                  Case 28
                     CfgName = "BIPATTERNTRIGPORT"
                     CfgVal = BIPATTERNTRIGPORT
                  Case 29
                     CfgName = "BIDIDEBOUNCESTATE"
                     CfgVal = BIDIDEBOUNCESTATE
                  Case 30
                     CfgName = "BIDIDEBOUNCETIME"
                     CfgVal = BIDIDEBOUNCETIME
                  Case 31
                     CfgName = "BIINPUTPACEROUT"
                     CfgVal = BIINPUTPACEROUT
                  Case 32
                     CfgName = "BIOUTPUTPACEROUT"
                     CfgVal = BIOUTPUTPACEROUT
                  Case 33
                     CfgName = "BISRCADPACER"
                     CfgVal = BISRCADPACER
                  Case 34
                     CfgName = "BIEXTCLKTYPE"
                     CfgVal = BIEXTCLKTYPE
                  Case 35
                     CfgName = "BIDACFORCESENSE"
                     CfgVal = BIDACFORCESENSE
                  Case 36
                     CfgName = "BISYNCMODE"
                     CfgVal = BISYNCMODE
                  Case 37
                     CfgName = "BIEXTINPACEREDGE"
                     CfgVal = BIEXTINPACEREDGE
                  Case 38
                     CfgName = "BIEXTOUTPACEREDGE"
                     CfgVal = BIEXTOUTPACEREDGE
                  Case 39
                     CfgName = "BIADCSETTLETIME"
                     CfgVal = BIADCSETTLETIME
                  Case 40
                     CfgName = "BICALOUTPUT"
                     CfgVal = BICALOUTPUT
                     FilterMask& = LISTTC Or LISTLEGACY
                     If ((CfgFilter And FilterMask&) = 0) _
                        Then NotDone = False
                  Case Else
                     If CfgVal = 0 Then CfgVal = -1
               End Select
            End If
            If TCFeatures Then
               Select Case Index
                  Case 41
                     CfgName = "BICHANTCTYPE"
                     CfgVal = BICHANTCTYPE
                  Case 42
                     CfgName = "BITEMPAVG"
                     CfgVal = BITEMPAVG
                  Case 43
                     CfgName = "BIEXCITATION"
                     CfgVal = BIEXCITATION
                  Case 44
                     CfgName = "BICHANBRIDGETYPE"
                     CfgVal = BICHANBRIDGETYPE
                  Case 45
                     CfgName = "BICHANRTDTYPE"
                     CfgVal = BICHANRTDTYPE
                  Case 46
                     CfgName = "BINUMTEMPCHANS"
                     CfgVal = BINUMTEMPCHANS
                  Case 47
                     CfgName = "BIUSESEXPS"
                     CfgVal = BIUSESEXPS
                  Case 48
                     CfgName = "BINUMEXPS"
                     CfgVal = BINUMEXPS
                     FilterMask& = LISTLEGACY
                     If ((CfgFilter And FilterMask&) = 0) _
                        Then NotDone = False
                  Case Else
                     If CfgVal = 0 Then CfgVal = -1
               End Select
            End If
            If LegacyFeatures Then
               Select Case Index
                  Case 49
                     CfgName = "BIHIDELOGINDLG"
                     CfgVal = BIHIDELOGINDLG
                  Case 50
                     CfgName = "BIPANID"
                     CfgVal = BIPANID
                  Case 51
                     CfgName = "BIRFCHANNEL"
                     CfgVal = BIRFCHANNEL
                  Case 52
                     CfgName = "BIRSS"
                     CfgVal = BIRSS
                  Case 53
                     CfgName = "BIBASEADR"
                     CfgVal = BIBASEADR
                  Case 54
                     CfgName = "BIINTLEVEL"
                     CfgVal = BIINTLEVEL
                  Case 55
                     CfgName = "BIDMACHAN"
                     CfgVal = BIDMACHAN
                  Case 56
                     CfgName = "BICLOCK"
                     CfgVal = BICLOCK
                  Case 57
                     CfgName = "BIRANGE"
                     CfgVal = BIRANGE
                  Case 58
                     CfgName = "BINUMIOPORTS"
                     CfgVal = BINUMIOPORTS
                  Case 59
                     CfgName = "BIWAITSTATE"
                     CfgVal = BIWAITSTATE
                  Case 60
                     CfgName = "BIDTBOARD"
                     CfgVal = BIDTBOARD
                     NotDone = False
               End Select
            End If
         Case DIGITALINFO
            Select Case Index
               Case 0
                  CfgName = "DIDEVTYPE"
                  CfgVal = DIDEVTYPE
               Case 1
                  CfgName = "DINUMBITS"
                  CfgVal = DINUMBITS
               Case 2
                  CfgName = "DIINMASK"
                  CfgVal = DIINMASK
               Case 3
                  CfgName = "DIOUTMASK"
                  CfgVal = DIOUTMASK
               Case 4
                  CfgName = "DICURVAL"
                  CfgVal = DICURVAL
                  NotDone = False
            End Select
         Case COUNTERINFO
            Select Case Index
               Case 0
                  CfgName = "CICTRTYPE"
                  CfgVal = CICTRTYPE
               Case 1
                  CfgName = "CICTRNUM"
                  CfgVal = CICTRNUM
                  NotDone = False
            End Select
         Case EXPANSIONINFO
            Select Case Index
               Case 0
                  CfgName = "XIBOARDTYPE"
                  CfgVal = XIBOARDTYPE
               Case 1
                  CfgName = "XIMUXADCHAN1"
                  CfgVal = XIMUXADCHAN1
               Case 2
                  CfgName = "XIMUXADCHAN2"
                  CfgVal = XIMUXADCHAN2
               Case 3
                  CfgName = "XIRANGE1"
                  CfgVal = XIRANGE1
               Case 4
                  CfgName = "XIRANGE2"
                  CfgVal = XIRANGE2
               Case 5
                  CfgName = "XICJCCHAN"
                  CfgVal = XICJCCHAN
               Case 6
                  CfgName = "XITHERMTYPE"
                  CfgVal = XITHERMTYPE
               Case 7
                  CfgName = "XINUMEXPCHANS"
                  CfgVal = XINUMEXPCHANS
                  NotDone = False
            End Select
      End Select
   End If
   GetCfgItemsEx = NotDone
   
End Function


Private Function GetCfgStringItems(ByVal Index As Integer, ByVal CfgFilter As Integer, _
   ByRef CfgVal As Long, ByRef CfgName As String) As Boolean

   Dim NotDone As Boolean
   
   NotDone = True
   InfoType& = cmbInfoType.ItemData(cmbInfoType.ListIndex)
   Select Case InfoType&
      Case GLOBALINFO
            NotDone = False
      Case BOARDINFO
         Select Case Index
            Case 0
               CfgName = "BIDEVUNIQUEID"
               CfgVal = BIDEVUNIQUEID
            Case 1
               CfgName = "BIUSERDEVID"
               CfgVal = BIUSERDEVID
            Case 2
               CfgName = "BIDEVVERSION"
               CfgVal = BIDEVVERSION
            Case 3
               CfgName = "BINODEID"
               CfgVal = BINODEID
            Case 4
               CfgName = "BIFACTORYID"
               CfgVal = BIFACTORYID
            Case 5
               CfgName = "BIMFGSERIALNUM"
               CfgVal = BIMFGSERIALNUM
            Case 6
               CfgName = "BIDEVNOTES"
               CfgVal = BIDEVNOTES
               NotDone = False
         End Select
   End Select
   GetCfgStringItems = NotDone
   
End Function

Public Function GetCfgItemIndex(CfgItem As Integer) As Integer

   ItemIndex% = 0
   For CItem% = 0 To Me.cmbConfigItem.ListCount - 1
      CurVal% = Me.cmbConfigItem.ItemData(CItem%)
      If CurVal% = CfgItem Then
         ItemIndex% = CItem%
         Exit For
      End If
   Next
   GetCfgItemIndex = CItem%
   
End Function

Private Function GetCfgItemStrFromVal(ByVal CfgVal As Long) As String

   Dim varCfgVal() As Variant, varCfgName() As Variant
   Dim FoundIt As Boolean

   StringFound$ = "Undefined config item"
   For Filt& = 0 To 5
      CfgFilter& = 2 ^ Filt&
      Items& = GetCfgItems(CfgFilter&, varCfgVal(), varCfgName())
      For Itm& = 0 To Items&
         If CfgVal = varCfgVal(Itm&) Then
            StringFound$ = varCfgName(Itm&)
            FoundIt = True
            Exit For
         End If
         If FoundIt Then
            GetCfgItemStrFromVal = StringFound$
            Exit Function
         End If
      Next
   Next
   GetCfgItemStrFromVal = StringFound$

End Function
