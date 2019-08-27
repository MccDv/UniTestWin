VERSION 5.00
Begin VB.Form frmCounter 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Counter Functions"
   ClientHeight    =   6420
   ClientLeft      =   1200
   ClientTop       =   4215
   ClientWidth     =   10605
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
   LinkTopic       =   "frmCounter"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   10605
   Visible         =   0   'False
   Begin VB.CheckBox chkManStart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Manual Start"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3510
      TabIndex        =   131
      Top             =   2295
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Start"
      Height          =   360
      Left            =   3285
      TabIndex        =   130
      Top             =   1890
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3555
      Top             =   4320
   End
   Begin VB.TextBox txtCount 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   2355
      TabIndex        =   128
      Text            =   "1"
      Top             =   1980
      Width           =   795
   End
   Begin VB.CommandButton cmdConfigure 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5400
      TabIndex        =   110
      Top             =   5940
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3975
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "&Go"
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   1890
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Frame fraCStat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   5400
      TabIndex        =   77
      Top             =   4080
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmdStopPls0 
         Appearance      =   0  'Flat
         Caption         =   "Stop"
         Height          =   312
         Left            =   1320
         TabIndex        =   121
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtIdle0 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   120
         Text            =   "0"
         ToolTipText     =   "Idle State"
         Top             =   1440
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtDly0 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   119
         Text            =   "0.50"
         ToolTipText     =   "Initial Delay"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTmr0Count 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   118
         Text            =   "1000"
         ToolTipText     =   "Number of pulses"
         Top             =   1140
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTmr0Dty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   116
         Text            =   "0.50"
         ToolTipText     =   "Duty Cycle"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTmr0Freq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Text            =   "100"
         ToolTipText     =   "Frequency"
         Top             =   540
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdStartTmr0 
         Appearance      =   0  'Flat
         Caption         =   "Start"
         Height          =   312
         Left            =   180
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Index Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   27
         Left            =   3840
         TabIndex        =   105
         Top             =   1470
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Counting Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   26
         Left            =   3840
         TabIndex        =   104
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   25
         Left            =   3840
         TabIndex        =   103
         Top             =   1050
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sign"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   24
         Left            =   3840
         TabIndex        =   102
         Top             =   840
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Compare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   23
         Left            =   3840
         TabIndex        =   101
         Top             =   630
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OverFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   22
         Left            =   3840
         TabIndex        =   100
         Top             =   420
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "UnderFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   21
         Left            =   3840
         TabIndex        =   99
         Top             =   210
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Index Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   20
         Left            =   2580
         TabIndex        =   98
         Top             =   1470
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Counting Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   19
         Left            =   2580
         TabIndex        =   97
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   18
         Left            =   2580
         TabIndex        =   96
         Top             =   1050
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sign"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   17
         Left            =   2580
         TabIndex        =   95
         Top             =   840
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Compare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   16
         Left            =   2580
         TabIndex        =   94
         Top             =   630
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OverFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   15
         Left            =   2580
         TabIndex        =   93
         Top             =   420
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "UnderFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   14
         Left            =   2580
         TabIndex        =   92
         Top             =   210
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Index Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   13
         Left            =   1380
         TabIndex        =   91
         Top             =   1470
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Counting Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   12
         Left            =   1380
         TabIndex        =   90
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   11
         Left            =   1380
         TabIndex        =   89
         Top             =   1050
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sign"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   10
         Left            =   1380
         TabIndex        =   88
         Top             =   840
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Compare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   9
         Left            =   1380
         TabIndex        =   87
         Top             =   630
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OverFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   8
         Left            =   1380
         TabIndex        =   86
         Top             =   420
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "UnderFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   7
         Left            =   1380
         TabIndex        =   85
         Top             =   210
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Index Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   6
         Left            =   180
         TabIndex        =   84
         Top             =   1470
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Counting Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   5
         Left            =   180
         TabIndex        =   83
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   4
         Left            =   180
         TabIndex        =   82
         Top             =   1050
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sign"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   3
         Left            =   180
         TabIndex        =   81
         Top             =   840
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Compare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   2
         Left            =   180
         TabIndex        =   80
         Top             =   630
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OverFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   1
         Left            =   180
         TabIndex        =   79
         Top             =   420
         Width           =   1200
      End
      Begin VB.CheckBox chkCStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "UnderFlow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   0
         Left            =   180
         TabIndex        =   78
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label lblTimerStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2400
         TabIndex        =   141
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblCount0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1380
         TabIndex        =   123
         Top             =   1140
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblDelay0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1740
         TabIndex        =   122
         Top             =   1440
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblDuty0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   117
         Top             =   1440
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label lblTmr0Freq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdNext1 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Height          =   315
      Left            =   840
      TabIndex        =   50
      Top             =   1950
      Width           =   615
   End
   Begin VB.CommandButton cmdBack1 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   49
      Top             =   1950
      Width           =   615
   End
   Begin VB.Timer tmrGoLoop 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   4320
   End
   Begin VB.Timer tmrCheckStatus 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3960
      Top             =   4320
   End
   Begin VB.Frame fraChipInit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5340
      TabIndex        =   37
      Top             =   60
      Visible         =   0   'False
      Width           =   5200
      Begin VB.ListBox lstCountCfg 
         Height          =   1035
         ItemData        =   "CTR.frx":0000
         Left            =   180
         List            =   "CTR.frx":001C
         MultiSelect     =   1  'Simple
         TabIndex        =   138
         Top             =   600
         Width           =   1515
      End
      Begin VB.ListBox lstGateOut 
         Height          =   1035
         ItemData        =   "CTR.frx":0073
         Left            =   1800
         List            =   "CTR.frx":0092
         MultiSelect     =   1  'Simple
         TabIndex        =   137
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox lstModeChange 
         Height          =   840
         ItemData        =   "CTR.frx":0104
         Left            =   3720
         List            =   "CTR.frx":0117
         MultiSelect     =   1  'Simple
         TabIndex        =   136
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txtMode 
         Height          =   285
         Left            =   1140
         TabIndex        =   135
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbTickSize 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CTR.frx":015C
         Left            =   1920
         List            =   "CTR.frx":015E
         TabIndex        =   127
         Text            =   "TickSize"
         Top             =   1380
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cmbCtrType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CTR.frx":0160
         Left            =   1920
         List            =   "CTR.frx":0162
         TabIndex        =   126
         Text            =   "CounterType"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Height          =   255
         Left            =   3960
         TabIndex        =   43
         Top             =   1500
         Width           =   975
      End
      Begin VB.ComboBox cmbTimeOfDay 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   40
         Text            =   "TimeOf Day"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbFOutSource 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   39
         Text            =   "FOut Source"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbFOutDiv 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3780
         TabIndex        =   38
         Text            =   "Div"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtMapChan 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   180
         TabIndex        =   51
         Text            =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblQuery 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4860
         TabIndex        =   42
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblMapChan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Map"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   111
         Top             =   300
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.Timer tmrCIn 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4380
      Top             =   4320
   End
   Begin VB.PictureBox picReadVal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1785
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5175
      TabIndex        =   15
      Top             =   4110
      Visible         =   0   'False
      Width           =   5200
   End
   Begin VB.Frame fraLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   5340
      TabIndex        =   11
      Top             =   1860
      Visible         =   0   'False
      Width           =   5205
      Begin VB.HScrollBar hsbPrescaler 
         Height          =   255
         LargeChange     =   4096
         Left            =   2820
         Min             =   -32768
         TabIndex        =   71
         Top             =   150
         Value           =   -32758
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.HScrollBar hsbAlarm2 
         Height          =   255
         LargeChange     =   4096
         Left            =   2820
         Min             =   -32768
         TabIndex        =   108
         Top             =   450
         Value           =   -32758
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtAlarm2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3660
         TabIndex        =   107
         Text            =   "10"
         ToolTipText     =   "Allows exponential notation (2^12)"
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPrescaler 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3660
         TabIndex        =   72
         Text            =   "10"
         ToolTipText     =   "Allows exponential notation (2^12)"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtHoldVal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Text            =   "10"
         ToolTipText     =   "Allows exponential notation (2^12)"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtLoadVal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   113
         Text            =   "10"
         ToolTipText     =   "Allows exponential notation (2^12)"
         Top             =   150
         Width           =   1095
      End
      Begin VB.HScrollBar hsbLoadVal 
         Height          =   255
         LargeChange     =   4096
         Left            =   120
         Min             =   -32768
         TabIndex        =   12
         Top             =   150
         Value           =   -32758
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.HScrollBar hsbHoldVal 
         Height          =   255
         LargeChange     =   4096
         Left            =   120
         Min             =   -32768
         TabIndex        =   45
         Top             =   480
         Value           =   -32758
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtRegLoadValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   134
         Text            =   "10"
         ToolTipText     =   "Generic load - valid for all registers. Allows exponential notation (2^12)"
         Top             =   1440
         Width           =   2835
      End
      Begin VB.CommandButton cmdRegLoad 
         Appearance      =   0  'Flat
         Caption         =   "Load Register"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   133
         ToolTipText     =   "Generic load - valid for all registers. Use parameters to the left."
         Top             =   900
         Width           =   1395
      End
      Begin VB.ComboBox cmbRegSelector 
         Height          =   315
         ItemData        =   "CTR.frx":0164
         Left            =   120
         List            =   "CTR.frx":019B
         TabIndex        =   132
         Text            =   "Register Selector"
         ToolTipText     =   "Generic load - valid for all registers. Select or type register number."
         Top             =   1440
         Width           =   2115
      End
      Begin VB.CheckBox chkEnableLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enable Load"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3600
         TabIndex        =   125
         Top             =   600
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox chkAlarm2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   420
         TabIndex        =   106
         Top             =   780
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkPrescaler 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   76
         Top             =   780
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkHold 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   75
         Top             =   450
         Width           =   255
      End
      Begin VB.CheckBox chkLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   74
         Top             =   120
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton cmdLoad 
         Appearance      =   0  'Flat
         Caption         =   "&Load"
         Height          =   312
         Left            =   4080
         TabIndex        =   112
         Top             =   210
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblScanStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   180
         TabIndex        =   139
         Top             =   900
         Width           =   3375
      End
      Begin VB.Label lblAlarm2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alarm 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   109
         Top             =   450
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblPrescaler 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prescaler"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   73
         Top             =   150
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblHold 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hold Reg"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1260
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Load Reg"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1260
         TabIndex        =   114
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Frame fraCounter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   5200
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   3420
         TabIndex        =   140
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkLoad64 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   124
         Top             =   1260
         Width           =   255
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3660
         TabIndex        =   52
         Text            =   "1000"
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdSetAll 
         Appearance      =   0  'Flat
         Caption         =   "Set All"
         Height          =   315
         Left            =   3900
         TabIndex        =   48
         Top             =   1020
         Width           =   975
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2700
         TabIndex        =   35
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   2700
         TabIndex        =   34
         Top             =   720
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   2700
         TabIndex        =   33
         Top             =   480
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   2700
         TabIndex        =   32
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   2040
         TabIndex        =   31
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   30
         Top             =   720
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   29
         Top             =   480
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   1380
         TabIndex        =   27
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1380
         TabIndex        =   26
         Top             =   720
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1380
         TabIndex        =   25
         Top             =   480
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   1380
         TabIndex        =   24
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   780
         TabIndex        =   23
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   780
         TabIndex        =   22
         Top             =   720
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   780
         TabIndex        =   21
         Top             =   480
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   780
         TabIndex        =   20
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   18
         Top             =   720
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   480
         Width           =   555
      End
      Begin VB.CheckBox chkCounter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkLoad32 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Use 16 bit value for read && load"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   420
         TabIndex        =   70
         Top             =   1260
         Width           =   3315
      End
   End
   Begin VB.Frame fraConfig 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Visible         =   0   'False
      Width           =   5200
      Begin VB.Frame fraQuadMode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   1410
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton optQuadMode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "X4"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   2580
            TabIndex        =   69
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optQuadMode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "X2"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   1860
            TabIndex        =   68
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optQuadMode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "X1"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   1140
            TabIndex        =   67
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optQuadMode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CheckBox chkClearOutput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Clear Output"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   2220
         TabIndex        =   63
         Top             =   1500
         Visible         =   0   'False
         Width           =   1392
      End
      Begin VB.CheckBox chkRetrigger 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Retrig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   62
         Top             =   1500
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkRecycle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Recycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   61
         Top             =   1500
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdLoadConfig 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Height          =   312
         Left            =   3900
         TabIndex        =   41
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkYesNoOpts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   3240
         TabIndex        =   60
         Top             =   1200
         Width           =   252
      End
      Begin VB.CheckBox chkYesNoOpts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   3540
         TabIndex        =   59
         Top             =   1200
         Width           =   252
      End
      Begin VB.CheckBox chkYesNoOpts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   3840
         TabIndex        =   58
         Top             =   1200
         Width           =   252
      End
      Begin VB.CheckBox chkYesNoOpts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   4140
         TabIndex        =   57
         Top             =   1200
         Value           =   1  'Checked
         Width           =   252
      End
      Begin VB.CheckBox chkYesNoOpts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   4440
         TabIndex        =   56
         Top             =   1200
         Width           =   252
      End
      Begin VB.CheckBox chkYesNoOpts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   4740
         TabIndex        =   55
         Top             =   1200
         Width           =   252
      End
      Begin VB.ComboBox cmbCountSource 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   53
         Text            =   "CountSource / FlagPins"
         Top             =   540
         Width           =   2535
      End
      Begin VB.ComboBox cmbGateControl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   115
         Text            =   "Gate Control / Index"
         Top             =   180
         Width           =   2535
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dummy"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1980
         TabIndex        =   36
         Top             =   1980
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HIGHONLASTCOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   1995
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ONESHOT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   1995
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "RATEGENERATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1995
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SQUAREWAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   1692
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SOFTWARESTROBE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   990
         Width           =   2055
      End
      Begin VB.OptionButton optConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HARDWARESTROBE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label lblCheckFunction 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check to set misc options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2220
         TabIndex        =   54
         Top             =   960
         Width           =   2895
      End
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1575
      TabIndex        =   129
      Top             =   1980
      Width           =   675
   End
   Begin VB.Label lblDDEStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5100
      TabIndex        =   64
      Top             =   3480
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   6030
      Width           =   5235
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
         Caption         =   "Select Counters"
         Index           =   0
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbC9513Init()"
         Index           =   1
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbC8254Config()"
         Index           =   2
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbCLoad()"
         Index           =   3
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbCIn()"
         Index           =   4
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbCStoreOnInt()"
         Index           =   5
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbCFreqIn()"
         Index           =   6
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbCStatus()"
         Index           =   7
      End
      Begin VB.Menu mnuSep1 
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
      Begin VB.Menu mnuClearOpts 
         Caption         =   "Reset Options"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep5 
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
         Caption         =   "SINGLEIO"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "BLOCKIO"
         Index           =   3
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "EXTTRIGGER"
         Index           =   4
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "EXTCLOCK"
         Index           =   5
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "CTR32BIT"
         Index           =   6
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "CTR48BIT"
         Index           =   7
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "CTR64BIT"
         Index           =   8
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "NOCLEAR"
         Index           =   9
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "RETRIGMODE"
         Index           =   10
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "HIGHRESRATE"
         Index           =   11
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContPlot 
         Caption         =   "Continuous Plot"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLoop 
         Caption         =   "Program Loop"
         Index           =   1
      End
      Begin VB.Menu mnuLoop 
         Caption         =   "Timer Loop"
         Index           =   2
      End
      Begin VB.Menu mnuHexVals 
         Caption         =   "Show Hex Values"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuStatusRate 
         Caption         =   "Status Rate..."
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
      Begin VB.Menu mnuPlotType 
         Caption         =   "Chunks"
         Enabled         =   0   'False
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
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuBufInfo 
         Caption         =   "Buffer Info"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSep4 
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
Attribute VB_Name = "frmCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SEL_CTR = 0
Const CB_CINIT = 1
Const CB_CONFIG = 2
Const CB_CLOAD = 3
Const CB_CIN = 4
Const CB_STOREINT = 5
Const CB_FREQIN = 6
Const CB_CSTAT = 7
Const CB_CCFGSCAN = 8
Const CB_CCLEAR = 9
Const CB_CSCAN = 10
Const CB_CSTART = 11
Const CB_CSTOP = 12

Const C8254 = 1
Const C9513 = 2
Const C8536 = 3
Const C7266 = 4
Const CEvent = 5
Const CIOTechScn = 6
Const CIOTechTmr = 7
Const CIOTechQuad = 8
Const CIOTechPulse = 9
'Const Cmini = 5 'obsolete (for miniLAB custom USB library)

#If MSGOPS Then
   Dim WithEvents MsgLibrary As MBDClass.MBDComClass
Attribute MsgLibrary.VB_VarHelpID = -1
#Else
   Dim MsgLibrary As Object
#End If

Private mfrmThisForm As Form
Private mnThisInstance As Integer

Private mnFormType As Integer, msTitle As String
Private mnDDEControl As Integer, mnDDEStat As Integer

Private mnFuncType As Integer, mnPlotType As Integer
Private mnFuncIndex As Integer

Private mlEvalCount As Long, mlTrapCount As Long
Private mbEvalEnabled As Boolean, mbEval As Boolean
Private mbEventCount As Boolean, mbStartTrap As Boolean
Private mnCountFirst As Integer, mnCountLast As Integer
Private mnCountIterations As Integer, mnCountInstances As Integer

Private mnPlot As Integer, mnLoaded As Integer

Private msConfig As String, msOpt As String
Private mnLoop As Integer

Private mnBoardIndex As Integer
Private mnBoardNum As Integer

Private mlCount As Long, mlRate As Long, mlFirstPoint As Long

Private mnCounterNum As Integer, mnNum8254s As Integer
Private mnNum9513s As Integer, mnNum8536s As Integer
Private mnNum7266s As Integer, mnNumIOTechScns As Integer
Private mnNumIOTechTmrs As Integer, mnNumEventCtrs As Integer
Private mnNumIOTechPulsers As Integer, mnNumIOTechQuads As Integer
Private mnCounterType As Integer, msCtrList As String

Private mnOutCtl As Integer, mnCountersEnabled As Integer

Private mvHandle As Variant, mvOptions As Variant
Private mvAuxHandle As Variant

Private mnTimerState As Integer, mnTimerNum As Integer
Private mlEventSize As Long, mlEventType As Long
Private mlEventParam As Long
Private mlBlockSave As Long
Private mnFirstChan As Integer, mnLastChan As Integer
Private mnEncoder As Integer, mlStatusType As Long

Private mnResolution As Integer, mlRateReturned As Long
Private mnRange As Integer, mnFixedRange As Integer
Private msBoardName As String, mnCounterBase As Integer
Private mnAcqDataType As VbVarType, mlStaticOpt As Long, mlAqcPoints As Long
Private mnHexVals As Integer, mlCountReturn As Long
Private mnCounterRes As Integer   'will be 1, 2, or 3 (number of choices for resolution)
Private mnNumEnabled As Integer   '1 based item for asynch reads (number of ctr check boxes checked)
Private maCtrList() As Integer
Private mnTimerStop As Integer, mnDelayRestart As Integer
Private mnCancel As Integer

Private mnLibType As Integer, mnNumBoards As Integer
Private msStatusType As String, mnMessaging As Integer   'added for msgDaq
Private mnUnloading As Integer, mnLoading As Integer
Private mnRefreshProps As Integer
Private mnFormInitialized As Integer, msDisplayName As String
Private msCTRSupport As String

Private mnFreq0Set As Integer, mnDelay0Set As Integer
Private mnDuty0Set As Integer, mnTmr0Running As Integer
Private mnTmrCount0Set As Integer, mnTmrCount1Set As Integer
Private mnIdleState0Set As Integer, mnIdleState1Set As Integer
Private mnFreq1Set As Integer, mnDelay1Set As Integer
Private mnDuty1Set As Integer, mnTmr1Running As Integer
Private mbBit48 As Boolean
Private mnCountRolls As Integer, mbHalfRoll As Integer
Private mnTrigType As Integer
Private mnLowThreshold As Integer, mnHighThreshold As Integer
Private mbThresholdSet As Boolean

Private Sub chkEnableLoad_Click()

   If chkEnableLoad.value = 1 Then
      Me.cmdLoad.ENABLED = True
      Me.cmdLoad.ToolTipText = "Right click to load without disabling button."
      Me.cmdGo.ENABLED = False
      lblLoad.Caption = "Load Reg"
      lblHold.Caption = "Hold Reg"
      lblPrescaler.Caption = "Prescaler"
      lblAlarm2.Caption = "Alarm 2"
   Else
      Me.cmdLoad.ENABLED = False
      Me.cmdLoad.ToolTipText = ""
      Me.cmdGo.ENABLED = True
      lblLoad.Caption = "Rate"
      lblHold.Caption = "Total Samples"
      lblPrescaler.Caption = "Low Chan"
      lblAlarm2.Caption = "High Chan"
   End If
   ConfigureControls
   
End Sub

Private Sub chkLoad32_Click()

   CounterBits% = GetCounterBits
   Select Case CounterBits%
      Case 16
         Me.chkLoad32.Caption = "Use 16 bit value for read && load"
      Case 32
         Me.chkLoad32.Caption = "Use 32 bit value for read && load"
      Case 64
         Me.chkLoad32.Caption = "Use 64 bit value for read && load"
   End Select
   
End Sub

Private Function GetCounterBits() As Integer

   If mnLibType = MSGLIB Then
      CounterBits% = 32
      mnuFuncArray(CB_CIN).Caption = "cbCIn32()"
   Else
      If chkLoad32.value = 1 Then
         If Me.chkLoad64.value = 1 Then
            Me.chkLoad32.Caption = "Use 64 bit value for read && load"
            CounterBits% = 64
            mnuFuncArray(CB_CIN).Caption = "cbCIn64()"
         Else
            Me.chkLoad32.Caption = "Use 32 bit value for read && load"
            CounterBits% = 32
            mnuFuncArray(CB_CIN).Caption = "cbCIn32()"
         End If
      Else
         If Me.chkLoad64.value = 1 Then
            Me.chkLoad32.Caption = "Use 32 bit value for read && load"
            CounterBits% = 32
            mnuFuncArray(CB_CIN).Caption = "cbCIn32()"
         Else
            Me.chkLoad32.Caption = "Use 16 bit value for read && load"
            CounterBits% = 16
            mnuFuncArray(CB_CIN).Caption = "cbCIn()"
         End If
      End If
   End If
   GetCounterBits = CounterBits%
   
End Function

Private Sub chkLoad64_Click()

   CounterBits% = GetCounterBits
   Select Case CounterBits%
      Case 16
         Me.chkLoad32.Caption = "Use 16 bit value for read && load"
      Case 32
         Me.chkLoad32.Caption = "Use 32 bit value for read && load"
      Case 64
         Me.chkLoad32.Caption = "Use 64 bit value for read && load"
   End Select

End Sub

Private Sub chkManStart_Click()

   cmdStart.Visible = (chkManStart.value = 1)
   cmdStop.Visible = (chkManStart.value = 1)
   cmdStop.ENABLED = (chkManStart.value = 1)

End Sub

Private Sub chkCounter_Click(Index As Integer)
   
   RefreshCtrList
   UpdateStatBar 0

End Sub

Private Sub chkCounter_MouseUp(Index As Integer, _
   Button As Integer, Shift As Integer, x As Single, y As Single)

   NewVal% = 0
   If Shift And vbShiftMask Then
      CurVal% = chkCounter(Index).value
      If Index > 0 Then
         For CControl% = Index - 1 To 0 Step -1
            If chkCounter(CControl%).value = CurVal% Then
               Exit For
            Else
               chkCounter(CControl%).value = CurVal%
            End If
         Next
      End If
   End If

End Sub

Private Sub chkCounter_MouseDown(Index As Integer, _
   Button As Integer, Shift As Integer, x As Single, y As Single)
   
End Sub

Private Sub chkYesNoOpts_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   StatString$ = " (Disabled)"
   Select Case Index
      Case 0
         If mnCounterType = C9513 Then
            FuncString$ = "Count Edge"
            StatString$ = " (Positive Edge)"
            If chkYesNoOpts(Index).value Then StatString$ = " (Negative Edge)"
         Else
            FuncString$ = "Data Encoding"
            StatString$ = " (BCD)"
            If chkYesNoOpts(Index).value Then StatString$ = " (Binary)"
         End If
      Case 1
         If mnCounterType = C9513 Then
            FuncString$ = "Special Gate"
            If chkYesNoOpts(Index).value Then StatString$ = " (Enabled)"
         Else
            FuncString$ = "Invert Index"
            If chkYesNoOpts(Index).value Then StatString$ = " (Enabled)"
         End If
      Case 2
         If mnCounterType = C9513 Then
            FuncString$ = "Reload"
            StatString$ = " (From Load)"
            If chkYesNoOpts(Index).value Then StatString$ = " (From Load and Hold)"
         Else
            FuncString$ = "Gating"
            If chkYesNoOpts(Index).value Then StatString$ = " (Enabled)"
         End If
      Case 3
         If mnCounterType = C9513 Then
            FuncString$ = "Recycle"
            StatString$ = " (One Time)"
            If chkYesNoOpts(Index).value Then StatString$ = " (Recycle)"
         Else
            FuncString$ = ""
            StatString$ = ""
         End If
      Case 4
         If mnCounterType = C9513 Then
            FuncString$ = "BCD Mode"
            If chkYesNoOpts(Index).value Then StatString$ = " (Enabled)"
         Else
            FuncString$ = ""
            StatString$ = ""
         End If
      Case 5
         If mnCounterType = C9513 Then
            FuncString$ = "Count Direction"
            StatString$ = " (Count Down)"
            If chkYesNoOpts(Index).value Then StatString$ = " (Count Up)"
         Else
            FuncString$ = ""
            StatString$ = ""
         End If
   End Select
   lblCheckFunction.Caption = FuncString$ & StatString$

End Sub

Private Sub cmbCtrType_Change()

   ConfigureControls
   
End Sub

Private Sub cmbCtrType_Click()

   ConfigureControls

End Sub

Private Sub cmbCtrType_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""
   
End Sub

Private Sub cmbCtrType_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)
   
   SetHelpText "cmbCtrType"

End Sub

Private Sub cmbFOutDiv_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmbFOutDiv_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)

   SetHelpText "cmbFOutDiv"

End Sub

Private Sub cmbFOutSource_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmbFOutSource_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)

   SetHelpText "cmbFOutSource"

End Sub

Private Sub cmbRegSelector_Change()

   Me.cmdRegLoad.ENABLED = True
   txtRegLoadValue.Text = ReadRegisters(0)

End Sub

Private Sub cmbRegSelector_Click()

   Me.cmdRegLoad.ENABLED = True
   If mnCounterType = CIOTechScn Then txtRegLoadValue.Text = ReadRegisters(0)

End Sub

Private Sub cmbRegSelector_KeyUp(KeyCode As Integer, Shift As Integer)

   If (KeyCode = 13) Or (KeyCode = 40) Or (KeyCode = 38) Then
      txtRegLoadValue.Text = ReadRegisters(0)
   End If

End Sub

Private Sub cmbRegSelector_Scroll()

   'txtRegLoadValue.Text = ReadRegisters(0)

End Sub

Private Sub cmbTickSize_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmbTickSize_DragOver(Source As Control, x As Single, y As Single, State As Integer)

   SetHelpText "cmbTickSize"

End Sub

Private Sub cmbTimeOfDay_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmbTimeOfDay_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)

   SetHelpText "cmbTimeOfDay"

End Sub

Private Sub cmdBack1_Click()

    On Error GoTo MenuBegin
    MenuIndex% = mnFuncIndex
    cmdNext1.ENABLED = True
    Do
        ValidMenu% = mnuFuncArray(MenuIndex% - 1).Visible
        MenuIndex% = MenuIndex% - 1
    Loop While Not ValidMenu%
    mnuFuncArray_Click (MenuIndex%)
    Do
        ValidMenu% = mnuFuncArray(MenuIndex% - 1).Visible
        MenuIndex% = MenuIndex% - 1
    Loop While Not ValidMenu%
    Exit Sub

MenuBegin:
    cmdBack1.ENABLED = False
    Exit Sub

End Sub

Private Sub cmdBack1_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmdConfigure_Click()

   'this exists to give menu access to the scripting
   'form when running scripts
   CmdStr$ = Left$(cmdConfigure.Caption, 1)
   value& = Val(Mid$(cmdConfigure.Caption, 2))
   Select Case CmdStr$
      Case "$" 'return board name
         cmdConfigure.Caption = msBoardName
      Case "1" 'return handle to data
         Me.cmdConfigure.Caption = Format(mvHandle, "0")
      Case "2" 'return data resolution
         Me.cmdConfigure.Caption = Format(mnResolution, "0")
      Case "3" 'return range at which data was collected
         Range% = -2
         Me.cmdConfigure.Caption = Format(Range%, "0")
      Case "4" 'return rate requested
         Me.cmdConfigure.Caption = Format(mlRate, "0")
      Case "5" 'return rate returned
         Me.cmdConfigure.Caption = Format(mlRateReturned, "0")
      Case "8"
         CurOption$ = msOpt
         If msOpt = "Options = Default  " Then CurOption$ = ""
         Me.cmdConfigure.Caption = CurOption$
      Case "9" 'set Compare1%
         'chkCompare1.value = value&
      Case "=" 'return number of samples
         Me.cmdConfigure.Caption = mlAqcPoints
      Case "?" 'return number of channels
         If mnFuncType = CB_CIN Then
            NumChans% = mnNumEnabled
         Else
            NumChans% = (mnLastChan - mnFirstChan) + 1
         End If
         Me.cmdConfigure.Caption = NumChans%
      Case "@" 'set Compare2%
         chkCompare2.value = value&
      Case "A" 'set TimeOfDay%
         cmbTimeOfDay.ListIndex = value&
      Case "B" 'find board name
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
            MsgBox SearchName$ & " not available in list of " & _
               "currently installed boards. Aborting script.", , _
               "Requested Board Not Available"
            gnScriptRun = False
         End If
      Case "C" 'configure counter
         optCtl% = -1
         Select Case mnCounterType
            Case C8254
               optCtl% = Choose(value& + 1, 0, 1, 2, 3, 4, 5)
            Case C9513
               optCtl% = Choose(value& + 1, 0, 1, 2, 0, 3, 4)
            Case C8536
               optCtl% = Choose(value& + 1, 0, 1, 2, 0, 0, 0, 3)
            Case C7266
               optCtl% = Choose(value& + 1, 0, 1, 2, 3)
         End Select
         If Not (optCtl% < 0) Then optConfig(optCtl%).value = True
         cmdLoadConfig = True
      Case "D" 'reload (or 7266 gate)
         chkYesNoOpts(2).value = value&
      Case "E" 'set edge (or 7266 decoding)
         If Left(msConfig, 7) = "cbC7266" Then
            chkYesNoOpts(0).value = value& - 1
         Else
            chkYesNoOpts(0).value = value&
         End If
      Case "F" 'set function
         mnuFuncArray_Click (value&)
      Case "G" 'set gate control
         If Not cmbGateControl.ListCount = 0 Then
            'values up to AHLPREVGATE are valid for 7266
            cmbGateControl.ListIndex = Choose(value& + 1, _
               NOGATE, AHLTCPREVCTR, AHLNEXTGATE, AHLPREVGATE, _
               AHLGATE, ALLGATE, AHEGATE, ALEGATE)
         End If
'         mnTrigType = Value&
      Case "H" 'set chip number
         For CheckBoxNum% = 0 To 19
            'clear each checkbox
            chkCounter(0).value = 0
         Next
         If mnCounterType = C9513 Then
            FirstCtrOnChip% = (value& - 1) * 5
            chkCounter(FirstCtrOnChip%).value = 1
         Else
            FirstCtrOnChip% = (value& - 1) * 3
            chkCounter(FirstCtrOnChip%).value = 1
         End If
      Case "I" 'read counter
         tmrCIn_Timer
      Case "J" 'set for 16, 32, or 64 bit IO (CIn, CIn32, CIn64, etc)
         chkLoad32.value = (value& And 1)
         chkLoad64.value = Switch((value& And 2) = 2, 1, (value& And 2) = 0, 0)
      Case "K"
         txtCount.Text = value&
      Case "L" 'load the selected counter register
         If Me.chkAlarm2.value = 1 Then Me.txtAlarm2.Text = value&
         If Me.chkPrescaler.value = 1 Then Me.txtPrescaler.Text = value&
         If Me.chkHold.value = 1 Then Me.txtHoldVal.Text = value&
         If Me.chkLoad.value = 1 Then txtLoadVal.Text = value&
         cmdLoad = True
      Case "M" 'set BCD Mode
         chkYesNoOpts(4).value = value&
      Case "N" 'set counter number
         For CheckBoxNum% = 0 To 20
            'clear each checkbox
            chkCounter(CheckBoxNum%).value = 0
         Next
         CtrIndex& = (value& Mod 20)
         If CtrIndex& < (value& And 1) Then CtrIndex& = 1
         If (CtrIndex& - mnCounterBase) < 0 Then
            MsgBox "Script is requesting an invalid counter number (" & _
            Format(value&, "0") & ").", vbCritical, "Bad Script Value"
            Exit Sub
         End If
         chkCounter(CtrIndex&).value = 1  ' - mnCounterBase
         chkLoad.value = 0
         Select Case value&
            Case Is > 800
               chkPrescaler.value = 1
            Case Is > 700
               chkHold.value = 1
            Case Is > 600
               'this one isn't necessary unless
               'script uses the COUNT1, etc values
               chkLoad.value = 1
            Case Is > 200
               chkAlarm2.value = 1
            Case Else
               chkLoad.value = 1
         End Select
      Case "O" 'set Options
         If value& = -1 Then
            ResetOpts
         Else
            mnuOpt_Click (value&)
         End If
      Case "P" 'set special gate (or 7266 index invert)
         chkYesNoOpts(1).value = value&
      Case "Q" 'get background status
         ULStat = GetStatus520(mnBoardNum, Status%, _
            CurCount&, CurIndex&, mlStatusType)
         x% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, _
            Status%, CurCount&, CurIndex&, mlStatusType, _
            A6, A7, A8, A9, A10, A11, 0)
         If (Status% = RUNNING) Then
            BGOp$ = "Background task in progress..."
         Else
            BGOp$ = "No background task running..."
         End If
         StatReport$ = Str(ULStat) & "," & Str(Status%) _
            & "," & Str(CurCount&) & "," & Str(CurIndex&)
         Me.cmdConfigure.Caption = StatReport$
         Stat$ = BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex&
         lblScanStatus.Caption = Stat$
         DoEvents
      Case "R" 'set the counter register
         'set all reg selections to off
         chkLoad.value = 0
         chkHold.value = 0
         chkPrescaler.value = 0
         chkAlarm2.value = 0
         Select Case value&
            Case Is < 100
               chkLoad.value = 1
            Case Is < 200
               chkHold.value = 1
            Case 201, 301, 401, 501
               chkPrescaler.value = 1
            Case 202, 302, 402, 502
               chkAlarm2.value = 1
         End Select
      Case "S" 'set counter source
         If Left(msConfig, 7) = "cbC7266" Then
            cmbCountSource.ListIndex = value& - 1
         Else
            cmbCountSource.ListIndex = value&
         End If
      Case "T" 'set count direction
         chkYesNoOpts(5).value = value&
      Case "U" 'set FOutSource
         cmbFOutSource.ListIndex = value&
      Case "V" 'set GateInterval
         cmbTimeOfDay.ListIndex = Log(value&) / Log(10)
      Case "W" 'set FOutDivider (or 7266 debounce)
         cmbFOutDiv.ListIndex = value&
      Case "X" 'select multiple counters
         G1 = Val(Mid$(cmdConfigure.Caption, 2))
         For CheckBoxNum% = 0 To 19
            chkCounter(CheckBoxNum%).value = 0
            If (G1 And 2 ^ CheckBoxNum%) = 2 ^ CheckBoxNum% Then
               chkCounter(CheckBoxNum%).value = 1
            End If
         Next
      Case "Y" 'set recycle mode
         chkYesNoOpts(3).value = value&
      Case "Z" 'update buffer info
         mnuBufInfo_Click
      Case "a" 'stop timer output
         Tmr% = value&
         ULStat = TmrOutStop(mnBoardNum, Tmr%)
         If SaveFunc(Me, TimerOutStop, ULStat, mnBoardNum, _
            Tmr%, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case "b" 'set timer number
         mnTimerNum = value&
      Case "c" 'set timer output
         Frequency# = Val(Mid$(cmdConfigure.Caption, 2))
         ULStat = TmrOutStart(mnBoardNum, mnTimerNum, Frequency#)
         If SaveFunc(Me, TimerOutStart, ULStat, mnBoardNum, mnTimerNum, _
            Frequency#, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case "d" 'set IOTech debounce
         If cmbFOutDiv.ListCount > value& Then cmbFOutDiv.ListIndex = value&
      Case "e" 'set IOTech edge detection
         'chkEdge.value = 0
         'Me.chkBEdge.value = 0
         'Me.chkZEdge.value = 0
         'If (value& And 1) = 1 Then chkEdge.value = 1
         'If (value& And 2) = 2 Then chkBEdge.value = 1
         'If (value& And 4) = 4 Then chkZEdge.value = 1
      Case "f" 'clear the counter register check boxes, set load checkbox
         chkLoad.value = 1
         chkHold.value = 0
         chkPrescaler.value = 0
         chkAlarm2.value = 0
      Case "g" 'set IOTech debounce trigger
         'chkTrigger.value = value&
      Case "h" 'set CInScan high channel
         txtAlarm2.Text = value&
      Case "i" 'set CInScan count
         txtHoldVal.Text = value&
      Case "l" 'set CInScan low channel
         txtPrescaler.Text = value&
      Case "m" 'set IOTech Counter modes
         If Left(msConfig, 7) = "cbC7266" Then
            Me.optConfig(value&).value = True
         Else
            For i% = 1 To 5
               OptVal& = Choose(i%, ENCODER, PERIOD, PULSEWIDTH, TIMING, 0)
               If (&HF00 And value&) = OptVal& Then
                  If OptVal& = 0 Then Exit For
                  Me.cmbCtrType.ListIndex = i%
                  Exit For
               End If
            Next
            For CtrlNum% = 0 To Me.Controls.Count - 1
               'uncheck all checkboxes in the ChipInit container
               Skip% = TypeOf Controls(CtrlNum%) Is Timer
               Skip% = Skip% Or (TypeOf Controls(CtrlNum%) Is Menu)
               If Not Skip% Then
                  If Controls(CtrlNum%).Container.Name = "fraChipInit" Then
                     If TypeOf Controls(CtrlNum%) Is CheckBox _
                        Then Controls(CtrlNum%).value = 0
                  End If
               End If
               Skip% = False
            Next
         Select Case OptVal&
            Case TOTALIZE
               'If (value& And CLEAR_ON_READ) = CLEAR_ON_READ Then chkCompare1.value = 1
               'If (value& And STOP_AT_MAX) = STOP_AT_MAX Then chkCompare2.value = 1
               'If (value& And LATCH_ON_MAP) = LATCH_ON_MAP Then chkLatchMap.value = 1
               'If (value& And DECREMENT_ON) = DECREMENT_ON Then chkCompare3.value = 1
            Case PERIOD, PULSEWIDTH
               'If (value& And PERIOD_MODE_X1) = PERIOD_MODE_X1 Then chkCompare1.value = 1
               'If (value& And PERIOD_MODE_X10) = PERIOD_MODE_X10 Then chkCompare2.value = 1
               'If (value& And PERIOD_MODE_X100) = PERIOD_MODE_X100 Then chkLatchMap.value = 1
               'If (value& And PERIOD_MODE_X1000) = PERIOD_MODE_X1000 Then chkClearOnZ.value = 1
            Case TIMING
            Case ENCODER
               'If (value& And ENCODER_MODE_X2) = ENCODER_MODE_X2 Then chkCompare1.value = 1
               'If (value& And ENCODER_MODE_X4) = ENCODER_MODE_X4 Then chkCompare2.value = 1
               'If (value& And LATCH_ON_Z) = LATCH_ON_Z Then chkLatchMap.value = 1
               'If (value& And CLEAR_ON_Z_ON) = CLEAR_ON_Z_ON Then chkClearOnZ.value = 1
               'If (value& And UPDOWN_ON) = UPDOWN_ON Then Me.chkUpDown.value = 1
         End Select
            'If (value& And MODULO_N_ON) = MODULO_N_ON Then chkModulo.value = 1
            'If (value& And GATING_ON) = GATING_ON Then chkCompare4.value = 1
            'If (value& And NO_RECYCLE_ON) = NO_RECYCLE_ON Then chkNoRecycle.value = 1
            'If (value& And RANGE_LIMIT_ON) = RANGE_LIMIT_ON Then Me.chkRangeLim.value = 1
            'If (value& And UPDOWN_ON) = UPDOWN_ON Then Me.chkUpDown.value = 1
            'If (value& And BIT_32) = BIT_32 Then chkBit32.value = 1
            'If (value& And BIT_48) = BIT_48 Then chkBit48.value = 1
         End If
      Case "n" 'set 7266 quad X1, 2, or 4
         optQuadMode(value&).value = True
      Case "p" 'set IOTech Counter map channel
         txtMapChan.Text = Format(value&, "0")
      Case "q" 'quit - stop background task
         mnCancel = True
      Case "r" 'set CInScan rate
         txtLoadVal.Text = value&
      Case "s" 'execute CInScan
         cmdGo = True
      Case "t" 'disable/enable counter read timer
         mnuLoop(2).Checked = value&
         tmrGoLoop.ENABLED = value&
         If Not value& Then cmdGo.FontBold = True
      Case "x" 'execute counter configuration
         cmdOK = True
      Case "y" 'execute cmdLoadConfig
         cmdLoadConfig = True
   End Select

End Sub

Private Sub cmdGo_Click()

   mnCancel = False
   SetPlotType mnPlotType + mnHardCopy, Me
   mnCountRolls = 1
   Select Case mnFuncType
      Case CB_STOREINT
         tmrGoLoop.ENABLED = mnuLoop(2).Checked
         cmdGo.FontBold = Not tmrGoLoop.ENABLED
         StartCounter
      Case CB_FREQIN
         tmrGoLoop.ENABLED = mnuLoop(2).Checked
         cmdGo.FontBold = Not tmrGoLoop.ENABLED
         StartFreq
      Case CB_CSTAT
      Case Else
         If mnFuncType = CB_CIN Then
            mnLoop = mnuLoop(1).Checked
            ReadByTimer% = mnuLoop(2).Checked
            DoEvents
            If (mnLibType = MSGLIB) And (chkManStart.value = 0) Then
               If ReadByTimer% And Not tmrCIn.ENABLED Then
                  NumEnabled% = mnNumEnabled - 1
                  For CtrIndex% = 0 To NumEnabled%
                     CtrNum% = maCtrList(CtrIndex%)
                     Failure% = ArmCounter(CtrNum%, True)
                  Next
               End If
               If Failure% Then Exit Sub
            End If
            ReadCounters   'read at least once
            cmdStop.Visible = mnuLoop(2).Checked Or (chkManStart.value = 1)
            tmrCIn.ENABLED = mnuLoop(2).Checked
            cmdGo.FontBold = Not tmrCIn.ENABLED
         Else
            mnEncoder = cmbCtrType.ListIndex = 1   '(chkEncoder.value = 1)
            tmrGoLoop.ENABLED = mnuLoop(2).Checked
            StartScan
         End If
   End Select
   cmdStop.ENABLED = (mnuLoop(1).Checked _
      Or mnuLoop(2).Checked) Or (chkManStart.value = 1)

End Sub

Private Sub cmdGo_MouseUp(Button As Integer, Shift As Integer, _
   x As Single, y As Single)

   If Button = 2 Then
      If mnuLoop(2).Checked Then
         'if already checked, just change loop rate
         LoopInterval% = tmrCIn.Interval
         If Shift > 0 Then
            LoopInterval% = LoopInterval% * 2
         Else
            LoopInterval% = LoopInterval% / 2
         End If
         tmrCIn.Interval = LoopInterval%
      End If
      'if right click, enable timer
      mnuLoop(2).Checked = True  'Not mnuLoop(2).Checked
      cmdGo = True
   End If

End Sub

Private Sub cmdLoad_Click()

   Bit32% = (chkLoad32.value = 1) Xor (Me.chkLoad64.value = 1)
   Bit64% = (chkLoad32.value = 1) And (Me.chkLoad64.value = 1)
   
   For CtrControl% = 0 To mnNumEnabled - 1
      CtrNum% = maCtrList(CtrControl%)
      If chkLoad.value Then
         RegNum% = CtrNum%
         If mnLibType = UNILIB Then
            If mnCounterType = C7266 Then RegNum% = _
            Choose(CtrNum%, COUNT1, COUNT2, COUNT3, COUNT4, 605, 606)
            CounterBits% = GetCounterBits()
            LoadValString$ = txtLoadVal.Text
            If Not (InStr(1, LoadValString$, "^") = 0) Then
               ExNote% = True
               Notation = Split(LoadValString$, "^")
               Mant& = Notation(0)
               Expnt& = Notation(1)
            End If
            Select Case CounterBits%
               Case 16
                  TempVal& = Val(txtLoadVal.Text)
                  If ExNote% Then TempVal& = Mant& ^ Expnt&
                  LoadValue% = ULongValToInt(TempVal&)
                  ULStat = cbCLoad(mnBoardNum, RegNum%, LoadValue%)
                  If SaveFunc(Me, CLoad, ULStat, mnBoardNum, _
                     RegNum%, LoadValue%, A4, A5, A6, A7, A8, _
                     A9, A10, A11, 0) Then Exit Sub
               Case 32
                  TempVal& = Val(txtLoadVal.Text)
                  If ExNote% Then TempVal& = Mant& ^ Expnt&
                  ULStat = cbCLoad32(mnBoardNum, RegNum%, TempVal&)
                  If SaveFunc(Me, CLoad32, ULStat, mnBoardNum, _
                     RegNum%, TempVal&, A4, A5, A6, A7, A8, _
                     A9, A10, A11, 0) Then Exit Sub
               Case 64
                  CtrLoadValue@ = Val(txtLoadVal.Text)
                  If ExNote% Then CtrLoadValue@ = Mant& ^ Expnt&
                  ULStat = CLoad64Bit(mnBoardNum, RegNum%, CtrLoadValue@)
                  If SaveFunc(Me, CLoad64, ULStat, mnBoardNum, _
                     RegNum%, CtrLoadValue@, A4, A5, A6, A7, _
                     A8, A9, A10, A11, 0) Then Exit Sub
            End Select
         Else
            LoadValString$ = txtLoadVal.Text
            CtrLoadMsg$ = msStatusType & "{" & _
               Format(RegNum%, "0") & "}:VALUE=" & LoadValString$
            MsgResult$ = MsgLibrary.SendMessage(CtrLoadMsg$)
            x% = SaveMsg(Me, "SendMessage(" & CtrLoadMsg$ & ")", MsgResult$)
         End If
      End If
      If chkHold.value Then
         If mnCounterType = C7266 Then
            RegNum% = Choose(CtrNum%, PRESET1, PRESET2, PRESET3, PRESET4, 705, 706)
         Else
            RegNum% = CtrNum% + 100
         End If
         
         HoldValString$ = txtHoldVal.Text
         If Not (InStr(1, HoldValString$, "^") = 0) Then
            ExNote% = True
            Notation = Split(HoldValString$, "^")
            Mant& = Notation(0)
            Expnt& = Notation(1)
         End If
         If Bit32% Then
            TempVal& = Val(HoldValString$)
            If ExNote% Then TempVal& = Mant& ^ Expnt&
            ULStat = cbCLoad32(mnBoardNum, RegNum%, TempVal&)
            If SaveFunc(Me, CLoad32, ULStat, _
               mnBoardNum, RegNum%, TempVal&, A4, _
               A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Else
            TempVal& = Val(HoldValString$)
            If ExNote% Then TempVal& = Mant& ^ Expnt&
            HoldValue% = ULongValToInt(TempVal&)
            ULStat = cbCLoad(mnBoardNum, RegNum%, HoldValue%)
            If SaveFunc(Me, CLoad, ULStat, _
               mnBoardNum, RegNum%, HoldValue%, A4, _
               A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
      End If
      
      If chkPrescaler.value Then
         If mnCounterType = C7266 Then
            RegNum% = Choose(CtrNum%, PRESCALER1, _
               PRESCALER2, PRESCALER3, PRESCALER4, 805, 806)
         Else
            Control% = 1
            Select Case CtrNum%
               Case Is > 5
                  Control% = 2
               Case Is > 10
                  Control% = 3
               Case Is > 15
                  Control% = 4
            End Select
            RegNum% = Choose(Control%, ALARM1CHIP1, _
               ALARM1CHIP2, ALARM1CHIP3, ALARM1CHIP4)
         End If

         PreValString$ = txtPrescaler.Text
         If Not (InStr(1, PreValString$, "^") = 0) Then
            ExNote% = True
            Notation = Split(PreValString$, "^")
            Mant& = Notation(0)
            Expnt& = Notation(1)
         End If
         If Bit32% Then
            TempVal& = Val(PreValString$)
            If ExNote% Then TempVal& = Mant& ^ Expnt&
            ULStat = cbCLoad32(mnBoardNum, RegNum%, TempVal&)
            If SaveFunc(Me, CLoad32, ULStat, _
               mnBoardNum, RegNum%, TempVal&, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0) _
               Then Exit Sub
         Else
            TempVal& = Val(PreValString$)
            If ExNote% Then TempVal& = Mant& ^ Expnt&
            PrescalerValue% = ULongValToInt(TempVal&)
            ULStat = cbCLoad(mnBoardNum, RegNum%, PrescalerValue%)
            If SaveFunc(Me, CLoad, ULStat, _
               mnBoardNum, RegNum%, PrescalerValue%, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0) _
               Then Exit Sub
         End If
      End If
      
      If chkAlarm2.value Then
         Control% = 1
         Select Case CtrNum%
            Case Is > 5
               Control% = 2
            Case Is > 10
               Control% = 3
            Case Is > 15
               Control% = 4
         End Select
         RegNum% = Choose(Control%, ALARM2CHIP1, _
            ALARM2CHIP2, ALARM2CHIP3, ALARM2CHIP4)
         
         AlarmValString$ = txtAlarm2.Text
         If Not (InStr(1, AlarmValString$, "^") = 0) Then
            ExNote% = True
            Notation = Split(AlarmValString$, "^")
            Mant& = Notation(0)
            Expnt& = Notation(1)
         End If
         If Bit32% Then
            TempVal& = Val(AlarmValString$)
            If ExNote% Then TempVal& = Mant& ^ Expnt&
            ULStat = cbCLoad32(mnBoardNum, RegNum%, TempVal&)
            If SaveFunc(Me, CLoad32, ULStat, _
               mnBoardNum, RegNum%, TempVal&, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0) _
               Then Exit Sub
         Else
            TempVal& = Val(AlarmValString$)
            If ExNote% Then TempVal& = Mant& ^ Expnt&
            Alarm2Value% = ULongValToInt(TempVal&)
            ULStat = cbCLoad(mnBoardNum, RegNum%, Alarm2Value%)
            If SaveFunc(Me, CLoad, ULStat, _
               mnBoardNum, RegNum%, Alarm2Value%, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0) _
               Then Exit Sub
         End If
      End If
   Next CtrControl%

End Sub

Private Sub cmdLoad_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   cmdLoad = True
   If Not Button = 2 Then
      If chkEnableLoad.value = 1 Then
         Me.chkEnableLoad.value = 0
      End If
   End If

End Sub

Private Sub cmdLoadConfig_Click()

   For CtrSelected% = 0 To mnNumEnabled - 1
      CounterNum% = maCtrList(CtrSelected%)
      If IsNumeric(chkCounter(CounterNum%).Tag) _
         Then CounterType% = Me.chkCounter(CounterNum%).Tag
      If (CounterType% = 0) And ((CounterNum% - mnCounterBase) > 0) Then
         If IsNumeric(chkCounter(0).Tag) Then CounterType% = Me.chkCounter(0).Tag
      End If
      Select Case CounterType%
         Case C8254
            ULStat = cbC8254Config(mnBoardNum, CounterNum%, mnOutCtl)
            x% = SaveFunc(Me, C8254Config, ULStat, _
               mnBoardNum, CounterNum%, mnOutCtl, _
               A4, A5, A6, A7, A8, A9, A10, A11, 0)
         Case C9513
            CounterEdge% = Choose(chkYesNoOpts(0).value + 1, _
               POSITIVEEDGE, NEGATIVEEDGE)
            If cmbCountSource.ListIndex = -1 Then
               CountSource% = Val(cmbCountSource.Text)
            Else
               CSource = Choose(cmbCountSource.ListIndex + 1, _
                  TCPREVCTR, CTRINPUT1, CTRINPUT2, CTRINPUT3, _
                  CTRINPUT4, CTRINPUT5, GATE1, GATE2, GATE3)
               If IsNull(CSource) Then CSource = _
                  Choose(cmbCountSource.ListIndex - 8, _
                  GATE4, GATE5, FREQ1, FREQ2, FREQ3, FREQ4, FREQ5)
               If Not IsNull(CSource) Then CountSource% = Val(CSource)
            End If
            SpecialGate% = Choose(chkYesNoOpts(1).value + 1, CBDISABLED, CBENABLED)
            Reload% = Choose(chkYesNoOpts(2).value + 1, LOADREG, LOADANDHOLDREG)
            RecycleMode% = Choose(chkYesNoOpts(3).value + 1, ONETIME, RECYCLE)
            BCDMode% = Choose(chkYesNoOpts(4).value + 1, CBDISABLED, CBENABLED)
            CountDirec% = Choose(chkYesNoOpts(5).value + 1, COUNTDOWN, COUNTUP)
            OutputCtrl% = mnOutCtl
            If cmbGateControl.ListIndex = -1 Then
               GateControl% = Val(cmbGateControl.Text)
            Else
               GateControl% = Choose(cmbGateControl.ListIndex + 1, _
                  NOGATE, AHLTCPREVCTR, AHLNEXTGATE, _
                  AHLPREVGATE, AHLGATE, ALLGATE, AHEGATE, ALEGATE)
            End If
            ULStat = cbC9513Config(mnBoardNum, CounterNum%, _
               GateControl%, CounterEdge%, CountSource%, _
               SpecialGate%, Reload%, RecycleMode%, _
               BCDMode%, CountDirec%, OutputCtrl%)
            If SaveFunc(Me, C9513Config, ULStat, mnBoardNum, _
               CounterNum%, GateControl%, CounterEdge%, _
               CountSource%, SpecialGate%, Reload%, _
               RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%, 0) _
               Then Exit Sub
            If chkClearOutput.value Then
               PortVal% = &HE0 Or (CtrControl% + mnCounterBase)
               PortNum% = (CtrControl% \ 5) * 2 + 1
               ULStat = cbOutByte(mnBoardNum, PortNum%, PortVal%)
               If SaveFunc(Me, OutByte, ULStat, _
                  mnBoardNum, PortNum%, PortVal%, _
                  A4, A5, A6, A7, A8, A9, A10, A11, 0) _
                  Then Exit Sub
            End If
         Case C8536
            RecycleMode% = Choose(chkRecycle.value + 1, ONETIME, RECYCLE)
            Retrigger% = Choose(cmbGateControl.ListIndex + 1, _
               HW_START_TRIG, HW_RETRIG, SW_START_TRIG)
            ULStat = cbC8536Config(mnBoardNum, CtrControl% _
               + mnCounterBase, mnOutCtl, RecycleMode%, Retrigger%)
            If SaveFunc(Me, C8536Config, ULStat, mnBoardNum, _
               CtrControl% + mnCounterBase, mnOutCtl, RecycleMode%, _
               Retrigger%, A6, A7, A8, A9, A10, A11, 0) _
               Then Exit Sub
         Case C7266
            For OptControl% = 0 To 3
               If optQuadMode(OptControl%).value Then
                  Quadrature& = Choose(OptControl% + 1, _
                     NO_QUAD, X1_QUAD, X2_QUAD, X4_QUAD)
                  Exit For
               End If
            Next OptControl%
            CountingMode& = mnOutCtl
            DataEncoding& = Choose(chkYesNoOpts(0).value + 1, _
               BCD_ENCODING, BINARY_ENCODING)
            If cmbGateControl.ListIndex < 0 Then
               MsgBox "Invalid Gating value specified in script.", _
                  vbCritical, "Bad Script Argument"
               Exit Sub
            End If
            IndexMode& = Choose(cmbGateControl.ListIndex + 1, _
               INDEX_DISABLED, LOAD_CTR, LOAD_OUT_LATCH, RESET_CTR)
            InvertIndex& = Choose(chkYesNoOpts(1).value + 1, CBDISABLED, CBENABLED)
            If cmbCountSource.ListIndex < 0 Then
               MsgBox "Invalid Flag Pin value specified in script.", _
                  vbCritical, "Bad Script Argument"
               Exit Sub
            End If
            FlagPins& = Choose(cmbCountSource.ListIndex + 1, _
               CARRY_BORROW, COMPARE_BORROW, CARRYBORROW_UPDOWN, INDEX_ERROR)
            GateEnable& = Choose(chkYesNoOpts(2).value + 1, CBDISABLED, CBENABLED)
            ULStat = cbC7266Config(mnBoardNum, CounterNum%, _
               Quadrature&, CountingMode&, DataEncoding&, _
               IndexMode&, InvertIndex&, FlagPins&, GateEnable&)
            If SaveFunc(Me, C7266Config, ULStat, mnBoardNum, _
               CounterNum%, Quadrature&, CountingMode&, _
               DataEncoding&, IndexMode&, InvertIndex&, FlagPins&, _
               GateEnable&, A10, A11, 0) _
               Then Exit Sub
         Case CIOTechScn, CEvent, CIOTechQuad
            A3 = CounterNum% + 1  'required for script, which assumes 1 based counters
            ULStat = CtrClear(mnBoardNum, CounterNum%)
            If SaveFunc(Me, CClear, ULStat, _
               mnBoardNum, CounterNum%, A3, A4, _
               A5, A6, A7, A8, A9, A10, A11, 0) _
               Then Exit Sub
      End Select
   Next
   
End Sub

Private Sub cmdNext1_Click()

    On Error GoTo MenuEnd
    cmdBack1.ENABLED = True
    MenuIndex% = mnFuncIndex
    Do
        ValidMenu% = mnuFuncArray(MenuIndex% + 1).Visible
        MenuIndex% = MenuIndex% + 1
    Loop While Not ValidMenu%
    mnuFuncArray_Click (MenuIndex%)
    Do
        ValidMenu% = mnuFuncArray(MenuIndex% + 1).Visible
        MenuIndex% = MenuIndex% + 1
    Loop While Not ValidMenu%
    Exit Sub

MenuEnd:
    cmdNext1.ENABLED = False
    Exit Sub

End Sub

Private Sub cmdNext1_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmdOK_Click()

   For CtrSelected% = 0 To 20
      If chkCounter(CtrSelected%).value = 1 Then
         If IsNumeric(chkCounter(CtrSelected%).Tag) _
            Then CounterType% = Me.chkCounter(CtrSelected%).Tag
         CtrNum% = CtrSelected% '+ mnCounterBase
         Select Case CounterType%
            Case C9513
               ThisChip% = ((CtrSelected% - mnCounterBase) \ 5) + 1
               If Not (ChipNum% = ThisChip%) Then
                  ChipNum% = ThisChip%
                  If cmbFOutDiv.ListIndex = -1 Then
                     FOutDivider% = Val(cmbFOutDiv.Text)
                  Else
                     FOutDivider% = cmbFOutDiv.ListIndex
                  End If
                  If cmbFOutSource.ListIndex = -1 Then
                     FOutSource% = Val(cmbFOutSource.Text)
                  Else
                     FOutSource% = cmbFOutSource.ListIndex + 1
                  End If
                  Compare1% = lstGateOut.Selected(7) 'chkCompare1.value
                  Compare2% = lstGateOut.Selected(8) 'chkCompare2.value
                  If cmbTimeOfDay.ListIndex = -1 Then
                     TimeOfDay% = Val(cmbTimeOfDay.Text)
                  Else
                     TimeOfDay% = cmbTimeOfDay.ListIndex
                  End If
                  ULStat = cbC9513Init(mnBoardNum, ChipNum%, _
                     FOutDivider%, FOutSource%, _
                     Compare1%, Compare2%, TimeOfDay%)
                  If SaveFunc(Me, C9513Init, ULStat, mnBoardNum, _
                     ChipNum%, FOutDivider%, FOutSource%, Compare1%, _
                     Compare2%, TimeOfDay%, A8, A9, A10, A11, 0) _
                     Then Exit Sub
               End If
            Case CIOTechScn, CIOTechQuad
               ModeType& = Choose(cmbCtrType.ListIndex + 1, _
                  TOTALIZE, ENCODER, PERIOD, PULSEWIDTH, TIMING)
               Mode& = ModeType&
               Select Case ModeType&
                  Case TOTALIZE
                     If lstModeChange.Selected(0) _
                        Then Mode& = Mode& Or CLEAR_ON_READ
                     If lstModeChange.Selected(1) _
                        Then Mode& = Mode& Or STOP_AT_MAX
                     If lstModeChange.Selected(2) _
                        Then Mode& = Mode& Or LATCH_ON_MAP
                     If lstModeChange.Selected(3) _
                        Then Mode& = Mode& Or DECREMENT_ON
                  Case ENCODER
                     If lstModeChange.Selected(0) _
                        Then Mode& = Mode& Or ENCODER_MODE_X2
                     If lstModeChange.Selected(1) _
                        Then Mode& = Mode& Or ENCODER_MODE_X4
                     If lstModeChange.Selected(2) _
                        Then Mode& = Mode& Or LATCH_ON_Z
                     If lstModeChange.Selected(3) _
                        Then Mode& = Mode& Or CLEAR_ON_Z_ON
                  Case PERIOD, PULSEWIDTH
                     If lstModeChange.Selected(0) _
                        Then Mode& = Mode& Or PERIOD_MODE_X1
                     If lstModeChange.Selected(1) _
                        Then Mode& = Mode& Or PERIOD_MODE_X10
                     If lstModeChange.Selected(2) _
                        Then Mode& = Mode& Or PERIOD_MODE_X100
                     If lstModeChange.Selected(3) _
                        Then Mode& = Mode& Or PERIOD_MODE_X1000
                  Case TIMING
               End Select
               If lstGateOut.Selected(0) _
                  Then Mode& = Mode& Or GATING_ON
               If lstGateOut.Selected(1) _
                  Then Mode& = Mode& Or INVERT_GATE
               If lstGateOut.Selected(2) _
                  Then Mode& = Mode& Or GATE_CONTROLS_DIR
               If lstGateOut.Selected(3) _
                  Then Mode& = Mode& Or GATE_CLEARS_CTR
               If lstGateOut.Selected(4) _
                  Then Mode& = Mode& Or GATE_TRIG_SRC
               If lstGateOut.Selected(5) _
                  Then Mode& = Mode& Or OUTPUT_ON
               If lstGateOut.Selected(6) _
                  Then Mode& = Mode& Or OUTPUT_INITIAL_STATE_HIGH
               If lstCountCfg.Selected(4) _
                  Then Mode& = Mode& Or MODULO_N_ON 'chkModulo.value
               If lstCountCfg.Selected(0) _
                  Then Mode& = Mode& Or NO_RECYCLE_ON 'chkNoRecycle.value
               If lstCountCfg.Selected(1) _
                  Then Mode& = Mode& Or RANGE_LIMIT_ON 'chkRangeLim.value
               If lstCountCfg.Selected(2) _
                  Then Mode& = Mode& Or UPDOWN_ON 'chkUpDown.value
               If lstCountCfg.Selected(3) _
                  Then Mode& = Mode& Or COUNT_DOWN_ON
               If lstCountCfg.Selected(5) _
                  Then Mode& = Mode& Or BIT_32 'chkBit32.value
               If lstCountCfg.Selected(6) _
                  Then Mode& = Mode& Or BIT_48 'chkBit48.value
               mbBit48 = lstCountCfg.Selected(6)
               MapChannel& = Val(txtMapChan.Text)
               DebounceTime& = cmbFOutDiv.ListIndex
               If cmbFOutDiv.ListIndex = -1 _
                  Then DebounceTime& = Val(cmbFOutDiv.Text)
               If lstCountCfg.Selected(7) Then _
                  DebounceTrigger& = CTR_TRIGGER_BEFORE_STABLE
               LastInList% = lstModeChange.ListCount - 1
               For CurItem% = 4 To LastInList%
                  If lstModeChange.Selected(CurItem%) Then _
                     EdgeDetection& = EdgeDetection& Or (2 ^ (CurItem% - 4))
               Next
               Ticks& = Me.cmbTickSize.ListIndex + 1
               TickSize& = Choose(Ticks&, CTR_TICK20PT83ns, _
                  CTR_TICK208PT3ns, CTR_TICK2083PT3ns, CTR_TICK20833PT3ns, _
                  CTR_TICK20ns, CTR_TICK200ns, CTR_TICK2000ns, CTR_TICK20000ns)
               A9 = CtrNum% + 1    'compensates for these being different from all other counters which are base 1
               ULStat = CtrConfigScan(mnBoardNum, CtrNum%, Mode&, _
                  DebounceTime&, DebounceTrigger&, _
                  EdgeDetection&, TickSize&, MapChannel&)
               If SaveFunc(Me, CConfigScan, ULStat, mnBoardNum, _
                  CtrNum%, Mode&, DebounceTime&, DebounceTrigger&, _
                  EdgeDetection&, TickSize&, MapChannel&, A9, A10, A11, 0) _
                  Then Exit Sub
            Case C8536
               ChipNum% = (CtrSelected% \ 3) + 1
               If cmbFOutSource.ListIndex = -1 Then
                  Ctr1Output% = Val(cmbFOutSource.Text)
               Else
                  Ctr1Output% = Choose(cmbFOutSource.ListIndex + 1, _
                     NOTLINKED, GATECTR2, TRIGCTR2, INCTR2)
               End If
               ULStat = cbC8536Init(mnBoardNum, ChipNum%, Ctr1Output%)
               If SaveFunc(Me, C8536Init, ULStat, mnBoardNum, _
                  ChipNum%, Ctr1Output%, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
                  Then Exit Sub
         End Select
      End If
   Next

End Sub

Private Sub cmdOK_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub cmdRegLoad_Click()

   CounterBase& = 0
   Bit32% = (chkLoad32.value = 1) Xor (Me.chkLoad64.value = 1)
   Bit64% = (chkLoad32.value = 1) And (Me.chkLoad64.value = 1)
   
   Bit32% = Me.mnuOpt(6).Checked Or Bit32%
   Bit64% = Me.mnuOpt(8).Checked Or Bit64%
   
   ConfItemString$ = Me.cmbRegSelector.Text
   If IsNumeric(ConfItemString$) Then
      RegSelected& = -1
      RegIndex& = 8
   Else
      RegIndex& = Me.cmbRegSelector.ListIndex + 1
      RegSelected& = Choose(RegIndex&, LOADREG0, HOLDREG1, COUNT1, _
      PRESET1, PRESCALER1, MINLIMITREG0, MAXLIMITREG0, OUTPUTVAL0REG0, _
      OUTPUTVAL1REG0, ALARM1CHIP1, ALARM2CHIP1, ALARM1CHIP2, ALARM2CHIP2, _
      ALARM1CHIP3, ALARM2CHIP3, ALARM1CHIP4, ALARM2CHIP4)
   End If
   
   CounterBits% = 16
   If Bit32% Then CounterBits% = 32
   If Bit64% Then CounterBits% = 64
   Resolution% = CounterBits%
   Instruct$ = " Select a data type capable of holding larger values."
   If Resolution% = 64 Then Instruct$ = " This value is not valid. " & _
      "Range is - 9223372036854775808 to 9223372036854775807"
   
   DataString$ = GetStringValueAsUType(txtRegLoadValue, Resolution%)
   If Resolution% = 0 Then
      MsgBox DataString$ & Instruct$, vbExclamation, "Data Type Inadequate for Value"
      Exit Sub
   End If
   
   LastRegIndex& = mnNumEnabled - 1
   For CtrControl% = 0 To LastRegIndex&
      CtrNum% = maCtrList(CtrControl%)
      Select Case RegSelected&
         Case -1
            RegNum& = Val(ConfItemString$)
         Case LOADREG0
            'CounterBase& = mnCounterBase
            RegNum& = CtrNum%
         Case HOLDREG1
            RegNum& = HOLDREG1 + (CtrNum% - mnCounterBase)
         Case COUNT1
            RegNum& = COUNT1 + (CtrNum% - mnCounterBase)
         Case PRESET1
            RegNum& = PRESET1 + (CtrNum% - mnCounterBase)
         Case PRESCALER1
            RegNum& = PRESCALER1 + (CtrNum% - mnCounterBase)
         Case MINLIMITREG0
            RegNum& = MINLIMITREG0 + (CtrNum% - mnCounterBase)
         Case MAXLIMITREG0
            RegNum& = MAXLIMITREG0 + (CtrNum% - mnCounterBase)
         Case OUTPUTVAL0REG0
            RegNum& = OUTPUTVAL0REG0 + (CtrNum% - mnCounterBase)
         Case OUTPUTVAL1REG0
            RegNum& = OUTPUTVAL1REG0 + (CtrNum% - mnCounterBase)
         Case Else
            LastRegIndex& = 0
            RegNum& = RegSelected&
      End Select
   
      Select Case CounterBits%
         Case 16
            LoadValue& = Val(DataString$) 'txtRegLoadValue.Text
            'If ExNote% Then LoadValue& = Mant& ^ Expnt&
            ULStat = cbCLoad(mnBoardNum, RegNum&, LoadValue&)
            If SaveFunc(Me, CLoad, ULStat, mnBoardNum, RegNum&, _
               LoadValue&, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Case 32
            'TestVal = Val(txtRegLoadValue.Text)
            'LoadValue& = Val(txtRegLoadValue.Text)
            'If ExNote% Then LoadValue& = Mant& ^ Expnt&
            LoadValue& = Val(DataString$)
            ULStat = cbCLoad32(mnBoardNum, RegNum&, LoadValue&)
            If SaveFunc(Me, CLoad32, ULStat, mnBoardNum, RegNum&, _
               LoadValue&, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Case 64
            'LoadString$ = txtRegLoadValue.Text
            'CurrencyString$ = Left(LoadString$, Len(LoadString$) - 4) & "." & Right(LoadString$, 4)
            CtrLoadValue@ = CCur(DataString$)
            ULStat = CLoad64Bit(mnBoardNum, RegNum&, CtrLoadValue@)
            If SaveFunc(Me, CLoad64, ULStat, mnBoardNum, RegNum&, _
               CtrLoadValue@, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      End Select
   Next
   
End Sub

Private Sub cmdSetAll_Click()

   If cmdSetAll.Caption = "Set All" Then
      cmdSetAll.Caption = "Clear All"
      CtlVal% = 1
   Else
      cmdSetAll.Caption = "Set All"
      CtlVal% = 0
   End If
   For i% = 0 To 20
      If chkCounter(i%).FontUnderline Then chkCounter(i%).value = CtlVal%
   Next i%

End Sub

Private Sub cmdStart_Click()

   NumEnabled% = mnNumEnabled - 1
   For CtrIndex% = 0 To NumEnabled%
      CtrNum% = maCtrList(CtrIndex%)
      Failure% = ArmCounter(CtrNum%, True)
   Next
   
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

   mnCancel = True
   mnLoop = False
   If mnLibType = MSGLIB Then
      NumEnabled% = mnNumEnabled - 1
      For CtrIndex% = 0 To NumEnabled%
         CtrNum% = maCtrList(CtrIndex%)
         Failure% = ArmCounter(CtrNum%, False)
      Next
   Else
      ULStat = StopBackground520(mnBoardNum, CTRFUNCTION)
      If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, _
         CTRFUNCTION, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
         Then Exit Sub
   End If
   tmrCIn.ENABLED = False
   tmrGoLoop.ENABLED = False
   cmdGo.FontBold = True
   cmdStartTmr0.FontBold = True
   cmdGo.ENABLED = True
   cmdStop.ENABLED = (chkManStart.value = 1)

End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   'if right click, turn off timer
   If Button = 2 Then
      mnLoop = False
      If mnLibType = MSGLIB Then
         NumEnabled% = mnNumEnabled - 1
         For CtrIndex% = 0 To NumEnabled%
            CtrNum% = maCtrList(CtrIndex%)
            Failure% = ArmCounter(CtrNum%, False)
         Next
      Else
         ULStat = StopBackground520(mnBoardNum, CTRFUNCTION)
         If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, CTRFUNCTION, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      End If
      mnuLoop(1).Checked = False
      mnuLoop(2).Checked = False
      tmrGoLoop.ENABLED = False
      tmrCIn.ENABLED = False
      tmrCheckStatus.ENABLED = False
      cmdStop.Visible = mnuLoop(2).Checked Or (chkManStart.value = 1)
      cmdGo.FontBold = True
   End If

End Sub

Private Sub ConfigureControls()
   
   'display appropriate controls and
   'hide others
   Show8254Funcs% = False
   Show9513Funcs% = False
   Show8536Funcs% = False
   Show7266Funcs% = False
   ShowEventFuncs% = False
   ShowClearFuncs% = False
   ShowIOTechScnFuncs% = False
   ShowIOTechTmrFuncs% = False
   ShowIOTechPlsFuncs% = False
   mnuFuncArray(CB_CLOAD).Caption = "cbCLoad()"
   mnuFuncArray(CB_CSTAT).Caption = "cbCStatus()"
   lblLoad.Caption = "Load Reg"
   Me.cmbTickSize.Visible = False
   lstGateOut.Height = 1035
   'Me.chkNoRecycle.Visible = True
   For CtrType% = 1 To 9
      NumOfType% = 0
      NumOfType% = Choose(CtrType, mnNum8254s, mnNum9513s, _
         mnNum8536s, mnNum7266s, mnNumEventCtrs, mnNumIOTechScns, _
         mnNumIOTechTmrs, mnNumIOTechQuads, mnNumIOTechPulsers)
      If NumOfType% > 0 Then
         Select Case CtrType%
            Case C8254
               Show8254Funcs% = True
            Case C9513
               Show9513Funcs% = True
               mnuFuncArray(CB_CINIT).Caption = "cbC9513Init()"
               mnuFuncArray(CB_CONFIG).Caption = "cbC9513Config()"
            Case C8536
               Show8536Funcs% = True
               mnuFuncArray(CB_CINIT).Caption = "cbC8536Init()"
               mnuFuncArray(CB_CONFIG).Caption = "cbC8536Config()"
            Case C7266
               Show7266Funcs% = True
               mnuFuncArray(CB_CONFIG).Caption = "cbC7266Config()"
               mnuFuncArray(CB_CLOAD).Caption = "cbCLoad()"
            Case CEvent
               'Show8254Funcs% = True
               ShowEventFuncs% = True
               ShowClearFuncs% = True
               mnuFuncArray(CB_CONFIG).Caption = "cbCClear()"
            Case CIOTechScn
               ShowIOTechScnFuncs% = True
               ShowClearFuncs% = True
               mnuFuncArray(CB_CINIT).Caption = "cbCConfigScan()"
               mnuFuncArray(CB_CONFIG).Caption = "cbCClear()"
               mnuFuncArray(CB_CLOAD).Caption = "cbCInScan()"
               If chkEnableLoad.value = 1 Then
                  CurBits% = GetCounterBits()
                  Select Case CurBits%
                     Case 16
                        mnuFuncArray(CB_CLOAD).Caption = "cbCLoad()"
                     Case 32
                        mnuFuncArray(CB_CLOAD).Caption = "cbCLoad32()"
                     Case 64
                        mnuFuncArray(CB_CLOAD).Caption = "cbCLoad64()"
                  End Select
               End If
            Case CIOTechTmr
               ShowIOTechTmrFuncs% = True
               mnuFuncArray(CB_CSTAT).Caption = "cbTimerOutStart()"
            Case CIOTechPulse
               ShowIOTechTmrFuncs% = True
               ShowIOTechPlsFuncs% = True
               mnuFuncArray(CB_CSTAT).Caption = "cbPulseOutStart()"
         End Select
      End If
   Next CtrType%
   
   mnuFuncArray(CB_CINIT).Visible = Show9513Funcs% Or Show8536Funcs% Or ShowIOTechScnFuncs%
   'Me.chkEnableLoad.Visible = Not Show9513Funcs%
   mnuFuncArray(CB_STOREINT).Visible = Show9513Funcs%
   mnuFuncArray(CB_FREQIN).Visible = Show9513Funcs%
   mnuFuncArray(CB_CSTAT).Visible = Show7266Funcs% Or ShowIOTechTmrFuncs%
   'mnuFuncArray(CB_CONFIG).Visible = Not ShowEventFuncs%
   'mnuStopBG.Checked = (mnFuncType = CB_STOREINT)
   cmbFOutSource.Visible = Show9513Funcs% Or Show8536Funcs%
   cmbFOutDiv.Visible = Show9513Funcs% Or ShowIOTechScnFuncs%
   cmbTimeOfDay.Visible = Show9513Funcs%
   'txtMode.Visible = ShowIOTechScnFuncs%
   
   cmdOK.Visible = Not (mnFuncType = CB_FREQIN)
   
   cmbGateControl.Visible = Show9513Funcs% Or Show7266Funcs% Or Show8536Funcs%
   cmbCountSource.Visible = Show9513Funcs% Or Show7266Funcs%

   fraQuadMode.Visible = Show7266Funcs%
   For ChkYesNo% = 0 To 5
      chkYesNoOpts(ChkYesNo%).Visible = Show9513Funcs% Or (Show7266Funcs% And (ChkYesNo% < 3))
   Next ChkYesNo%
   lblCheckFunction.Visible = Show9513Funcs% Or Show7266Funcs%

   cmbCtrType.Visible = ShowIOTechScnFuncs%
   IOTType& = cmbCtrType.ListIndex
   txtHoldVal.Visible = Show9513Funcs% Or Show7266Funcs% Or ShowIOTechScnFuncs%
   hsbHoldVal.Visible = Show9513Funcs% Or Show7266Funcs% Or ShowIOTechScnFuncs%
   chkHold.Visible = Show9513Funcs% Or Show7266Funcs%
   lblHold.Visible = Show9513Funcs% Or Show7266Funcs% Or ShowIOTechScnFuncs%
   hsbPrescaler.Visible = Show9513Funcs% Or Show7266Funcs% Or ShowIOTechScnFuncs%
   chkPrescaler.Visible = Show9513Funcs% Or Show7266Funcs%
   txtPrescaler.Visible = Show9513Funcs% Or Show7266Funcs% Or ShowIOTechScnFuncs%
   lblPrescaler.Visible = Show9513Funcs% Or Show7266Funcs% Or ShowIOTechScnFuncs%
   hsbAlarm2.Visible = Show9513Funcs% Or ShowIOTechScnFuncs%
   txtAlarm2.Visible = Show9513Funcs% Or ShowIOTechScnFuncs%
   lblAlarm2.Visible = Show9513Funcs% Or ShowIOTechScnFuncs%
   chkAlarm2.Visible = Show9513Funcs%
   chkLoad.Visible = Not ShowIOTechScnFuncs%
   
   cmdLoad.ENABLED = (Not ShowIOTechScnFuncs%) Or (Me.chkEnableLoad.value = 1)
   
   txtMapChan.Visible = ShowIOTechScnFuncs%
   lblMapChan.Visible = ShowIOTechScnFuncs%

   fraChipInit.Visible = mnuFuncArray(CB_CINIT).Checked Or mnuFuncArray(CB_FREQIN).Checked
   cmbFOutDiv.Visible = ((mnuFuncArray(CB_CINIT).Checked) And Show9513Funcs%) Or ShowIOTechScnFuncs%
   
   lblStatus.Caption = msCtrList '& msConfig & " " & msOpt
   fraCounter.Visible = mnuFuncArray(SEL_CTR).Checked Or (mnFuncType = CB_STOREINT)
   txtCount.Visible = (mnFuncType = CB_STOREINT) Or (mnFuncType = CB_CIN)
   lblCount.Visible = (mnFuncType = CB_STOREINT) Or (mnFuncType = CB_CIN)
   fraConfig.Visible = mnuFuncArray(CB_CONFIG).Checked
   'optConfig(6).Value = mnuFuncArray(CB_CONFIG).Checked
   fraLoad.Visible = mnuFuncArray(CB_CLOAD).Checked
   fraCStat.Visible = mnuFuncArray(CB_CSTAT).Checked
   ShowGo% = ((mnuFuncArray(CB_CIN).Checked) _
      Or ((mnuFuncArray(CB_CSTAT).Checked) _
      And Show7266Funcs%)) _
      Or (mnuFuncArray(CB_CLOAD).Checked _
      And ShowIOTechScnFuncs%)
   cmdGo.Visible = ShowGo% _
      Or (mnFuncType = CB_CSCAN) _
      Or (mnFuncType = CB_FREQIN) _
      Or (mnFuncType = CB_STOREINT)
   ShowStop% = (mnuFuncArray(CB_CLOAD).Checked And ShowIOTechScnFuncs%)
   cmdStop.Visible = ShowStop% _
      Or (mnFuncType = CB_STOREINT) _
      Or ((mnFuncType = CB_CIN) _
      And (mnuLoop(2).Checked _
      Or mnuLoop(1).Checked)) _
      Or (chkManStart.value = 1)
   picReadVal.Visible = mnuFuncArray(CB_CIN).Checked
   chkManStart.Visible = (mnLibType = MSGLIB) And mnuFuncArray(CB_CIN).Checked

   
   For ChkNum% = 0 To 20
      chkCounter(ChkNum%).Caption = Format(ChkNum%, "0") ' + mnCounterBase
   Next ChkNum%
   optConfig(0).Visible = True
   optConfig(1).Visible = True
   optConfig(2).Visible = True
   optConfig(3).Visible = True
   optConfig(4).Visible = True
   If Show9513Funcs% Then
      optConfig(0).Visible = True
      optConfig(1).Visible = True
      optConfig(2).Visible = True
      optConfig(3).Visible = True
      optConfig(4).Visible = True
      optConfig(5).Visible = False
      optConfig(0).Caption = "ALWAYSLOW"
      optConfig(1).Caption = "HIGHPULSEONTC"
      optConfig(2).Caption = "TOGGLEONTC"
      optConfig(3).Caption = "DISCONNECTED"
      optConfig(4).Caption = "LOWPULSEONTC"
      lblHold.Caption = "Hold Reg"
      lblPrescaler.Caption = "Alarm1"
   End If
   If Show8254Funcs% Then
      mnuFuncArray(CB_CONFIG).Caption = "cbC8254Config()"
      optConfig(3).Visible = True
      optConfig(4).Visible = True
      optConfig(5).Visible = True
      optConfig(0).Caption = "HIGHONLASTCOUNT"
      optConfig(1).Caption = "ONESHOT"
      optConfig(2).Caption = "RATEGENERATOR"
      optConfig(3).Caption = "SQUAREWAVE"
      optConfig(4).Caption = "SOFTWARESTROBE"
      optConfig(5).Caption = "HARDWARESTROBE"
   End If
   If Show8536Funcs% Then
      mnuFuncArray(CB_CONFIG).Caption = "cbC8536Config()"
      optConfig(0).Caption = "HIGHPULSEONTC"
      optConfig(1).Caption = "TOGGLEONTC"
      optConfig(2).Caption = "HIGHUNTILTC"
      optConfig(3).Visible = False
      optConfig(4).Visible = False
      optConfig(5).Visible = False
   End If
   If Show7266Funcs% Then
      optConfig(3).Visible = True
      optConfig(4).Visible = False
      optConfig(5).Visible = False
      optConfig(0).Caption = "NORMALMODE"
      optConfig(1).Caption = "RANGELIMIT"
      optConfig(2).Caption = "NORECYCLE"
      optConfig(3).Caption = "MODULON"
      'chkLoad32.value = 1
      lblHold.Caption = "Preset Value"
   End If
   If ShowClearFuncs% Then
      optConfig(0).Visible = False
      optConfig(1).Visible = False
      optConfig(2).Visible = False
      optConfig(3).Visible = False
      optConfig(4).Visible = False
      optConfig(5).Visible = False
   End If
   If ShowIOTechScnFuncs% Then
      hsbAlarm2.Max = 20
      hsbAlarm2.Min = 0
      hsbAlarm2.LargeChange = 1
      hsbAlarm2.SmallChange = 1
      hsbPrescaler.Max = 20
      hsbPrescaler.Min = 0
      hsbPrescaler.LargeChange = 1
      chkPrescaler.value = 0
      chkHold.value = 0
      mnuFuncArray(CB_CINIT).Caption = "cbCConfigScan()"
      optConfig(0).Visible = False
      optConfig(1).Visible = False
      optConfig(2).Visible = False
      optConfig(3).Visible = False
      optConfig(4).Visible = False
      optConfig(5).Visible = False
      'chkCompare3.Caption = "DecrementOn"
      'chkCompare4.Caption = "GatingOn"
      CounterMode& = Choose(IOTType& + 1, _
        TOTALIZE, ENCODER, PERIOD, PULSEWIDTH, TIMING)
   
      AdvancedCtr% = ((msBoardName = "USB-CTR04") _
         Or (msBoardName = "USB-CTR08"))
      If Not AdvancedCtr% Then
         'moved following to ChangeCounterType IOTType& sub
         ChangeCounterType IOTType&
         Select Case CounterMode&
            Case TOTALIZE
               'Me.lstModeChange.Clear
               'lstModeChange.AddItem "ClearOnRead"
               'lstModeChange.AddItem "StopAtMax"
               'lstModeChange.AddItem "LatchOnMap"
               'lstModeChange.AddItem "DecrementOn"
               'lstModeChange.AddItem "CtrFallingEdge"
            Case ENCODER
               'Me.lstModeChange.Clear
               'lstModeChange.AddItem "EncoderX2"
               'lstModeChange.AddItem "EncoderX4"
               'lstModeChange.AddItem "LatchOnZ"
               'lstModeChange.AddItem "ClearOnZ"
               'lstModeChange.AddItem "FallEdgeA"
               'lstModeChange.AddItem "FallEdgeB"
               'lstModeChange.AddItem "FallEdgeZ"
            Case PERIOD
               'Me.lstModeChange.Clear
               'lstModeChange.AddItem "PeriodX1"
               'lstModeChange.AddItem "PeriodX10"
               'lstModeChange.AddItem "PeriodX100"
               'lstModeChange.AddItem "PeriodX1000"
               'lstModeChange.AddItem "CtrFallingEdge"
               cmbTickSize.Visible = True
               lstGateOut.Height = 645
            Case PULSEWIDTH
               'Me.lstModeChange.Clear
               'lstModeChange.AddItem "ClearOnRead"
               'lstModeChange.AddItem "StopAtMax"
               'lstModeChange.AddItem "LatchOnMap"
               'lstModeChange.AddItem "DecrementOn"
               'lstModeChange.AddItem "CtrFallingEdge"
               cmbTickSize.Visible = True
               lstGateOut.Height = 645
            Case TIMING
               'Me.lstModeChange.Clear
               'lstModeChange.AddItem "ClearOnRead"
               'lstModeChange.AddItem "StopAtMax"
               'lstModeChange.AddItem "LatchOnMap"
               'lstModeChange.AddItem "DecrementOn"
               'lstModeChange.AddItem "CtrFallingEdge"
               For CtrlNum% = 0 To Me.Controls.Count - 1
                  'set all checkboxes in the ChipInit container for FontBold False
                  Skip% = TypeOf Controls(CtrlNum%) Is Timer
                  Skip% = Skip% Or (TypeOf Controls(CtrlNum%) Is Menu)
                  If Not Skip% Then
                     If Controls(CtrlNum%).Container.Name = "fraChipInit" Then
                        If TypeOf Controls(CtrlNum%) Is CheckBox _
                           Then Controls(CtrlNum%).FontBold = False
                     End If
                  End If
                  Skip% = False
               Next
               Me.cmbTickSize.Visible = True
               lstGateOut.Height = 645
         End Select
      Else
         ChangeCounterType IOTType&
      End If
      
      LowCtr% = 20
      HighCtr% = 0
      'following two do nothing?
      'LowVal& = Val(txtPrescaler.Text)
      'HighVal& = Val(txtAlarm2.Text)
      
      For ChkNum% = 0 To 20
         'chkCounter(ChkNum%).Caption = Format(ChkNum% + mnCounterBase, "0")
         If chkCounter(ChkNum%).value = 1 Then
            'high and low channels for scan
            If LowCtr% > ChkNum% Then LowCtr% = ChkNum%
            If HighCtr% < ChkNum% Then HighCtr% = ChkNum%
            CtrsSelected% = True
         End If
      Next ChkNum%
      If hsbPrescaler.value = hsbAlarm2.value Then
         If (hsbPrescaler.value < LowCtr%) Or (hsbAlarm2.value > HighCtr%) Then
            hsbPrescaler.value = LowCtr%
            hsbAlarm2.value = HighCtr%
         End If
      End If
      If Not CtrsSelected% Then
         hsbPrescaler.value = 0
         hsbAlarm2.value = 0
      End If
      If Not chkEnableLoad.value = 1 Then
         lblLoad.Caption = "Rate"
         lblHold.Caption = "Total Samples"
         lblPrescaler.Caption = "Low Chan"
         lblAlarm2.Caption = "High Chan"
         'txtPrescaler.Text = "0"
         'txtAlarm2.Text = "0"
      End If
   End If
   
   If Show7266Funcs% Or ShowIOTechTmrFuncs% Then
      For ChkStatCtl% = 0 To 27
         chkCStat(ChkStatCtl%).Visible = Show7266Funcs%
      Next ChkStatCtl%
      lblTmr0Freq.Visible = ShowIOTechTmrFuncs% 'And Not ShowIOTechPlsFuncs%
      lblDelay0.Visible = ShowIOTechTmrFuncs%
      lblCount0.Visible = ShowIOTechTmrFuncs%
      Me.lblDuty0.Visible = ShowIOTechPlsFuncs%
      Me.lblTimerStat.Visible = ShowIOTechPlsFuncs%
      Me.cmdStopPls0.Visible = ShowIOTechPlsFuncs%
      txtTmr0Freq.Visible = ShowIOTechTmrFuncs% Or ShowIOTechPlsFuncs%
      Me.txtTmr0Dty.Visible = ShowIOTechPlsFuncs%
      Me.txtIdle0.Visible = ShowIOTechPlsFuncs%
      Me.txtTmr0Count.Visible = ShowIOTechPlsFuncs%
      Me.txtDly0.Visible = ShowIOTechPlsFuncs%
      cmdStartTmr0.Visible = ShowIOTechTmrFuncs%
      fraCStat.ENABLED = ShowIOTechTmrFuncs%
      If Show7266Funcs% Then lblPrescaler.Caption = "Prescaler"
      If ShowIOTechTmrFuncs% Then
         lblTmr0Freq.Left = 1300
         lblDuty0.Left = 1300
         lblTmr0Freq.Width = 1400
         lblDuty0.Width = 1400
         lblTmr0Freq.Top = Me.txtTmr0Freq.Top
         lblDuty0.Top = Me.txtTmr0Dty.Top
      End If
   End If
   Me.lblDuty0.Visible = Me.txtTmr0Dty.Visible
   Me.lblDelay0.Visible = Me.txtDly0.Visible
   Me.lblCount0.Visible = Me.txtTmr0Count.Visible
   msConfig = mnuFuncArray(mnFuncIndex).Caption
   Me.Caption = msConfig & ": " & " Board " & mnuBoard(mnBoardIndex).Caption
   cmdStart.Visible = (mnLibType = MSGLIB) And (Me.chkManStart.value = 1)

End Sub

Private Sub ChangeCounterType(IOTType&)
   
   CounterMode& = Choose(IOTType& + 1, _
     TOTALIZE, ENCODER, PERIOD, PULSEWIDTH, TIMING)

   Select Case CounterMode&
      Case TOTALIZE
         Me.lstModeChange.Clear
         lstModeChange.AddItem "ClearOnRead"
         lstModeChange.AddItem "StopAtMax"
         lstModeChange.AddItem "LatchOnMap"
         lstModeChange.AddItem "DecrementOn"
         lstModeChange.AddItem "CtrFallingEdge"
      Case ENCODER
         Me.lstModeChange.Clear
         lstModeChange.AddItem "EncoderX2"
         lstModeChange.AddItem "EncoderX4"
         lstModeChange.AddItem "LatchOnZ"
         lstModeChange.AddItem "ClearOnZ"
         lstModeChange.AddItem "FallEdgeA"
         lstModeChange.AddItem "FallEdgeB"
         lstModeChange.AddItem "FallEdgeZ"
      Case PERIOD
         Me.lstModeChange.Clear
         lstModeChange.AddItem "PeriodX1"
         lstModeChange.AddItem "PeriodX10"
         lstModeChange.AddItem "PeriodX100"
         lstModeChange.AddItem "PeriodX1000"
         lstModeChange.AddItem "CtrFallingEdge"
         cmbTickSize.Visible = True
         lstGateOut.Height = 645
      Case PULSEWIDTH
         Me.lstModeChange.Clear
         lstModeChange.AddItem "ClearOnRead"
         lstModeChange.AddItem "StopAtMax"
         lstModeChange.AddItem "LatchOnMap"
         lstModeChange.AddItem "DecrementOn"
         lstModeChange.AddItem "CtrFallingEdge"
         cmbTickSize.Visible = True
         lstGateOut.Height = 645
      Case TIMING
         Me.lstModeChange.Clear
         lstModeChange.AddItem "ClearOnRead"
         lstModeChange.AddItem "StopAtMax"
         lstModeChange.AddItem "LatchOnMap"
         lstModeChange.AddItem "DecrementOn"
         lstModeChange.AddItem "CtrFallingEdge"
         For CtrlNum% = 0 To Me.Controls.Count - 1
            'set all checkboxes in the ChipInit container for FontBold False
            Skip% = TypeOf Controls(CtrlNum%) Is Timer
            Skip% = Skip% Or (TypeOf Controls(CtrlNum%) Is Menu)
            If Not Skip% Then
               If Controls(CtrlNum%).Container.Name = "fraChipInit" Then
                  If TypeOf Controls(CtrlNum%) Is CheckBox Then Controls(CtrlNum%).FontBold = False
               End If
            End If
            Skip% = False
         Next
         Me.cmbTickSize.Visible = True
         lstGateOut.Height = 645
   End Select

End Sub

Private Sub DisplayData()

   SetPlotType mnPlotType + mnHardCopy, Me
   Select Case mnAcqDataType
      Case vbDecimal
         SetBufferType 3
      Case vbVariant
         SetBufferType 5
      Case vbDouble, vbLong
         SetBufferType 2
      Case vbInteger
         SetBufferType 1
   End Select

   If Not mnPlot Then Exit Sub
   DoEvents
   BoardName = mnuBoard(mnBoardIndex).Caption
   Select Case mnFuncType
      Case CB_CIN
         Bitness% = GetCounterBits()
         Select Case Bitness%
            Case 16
               ULFunction% = CIn
            Case 32
               ULFunction% = CIn32
            Case 64
               ULFunction% = CIn64
         End Select
         A2 = Trim(msCtrList)
         SetDetails ULFunction%, mlCount, A2, A3, A4, A5, BoardName, A6
      Case CB_STOREINT
         StoreCount& = mlCount
         If mlCount = 0 Then StoreCount& = mnNum9513s
         SetDetails CStoreOnInt, StoreCount&, A2, A3, A4, A5, BoardName, A6
      Case CB_CSCAN
         SetDetails CInScan, mnFirstChan, mnLastChan - mnFirstChan, _
            mnAvgVal, mlRateReturned, mnRange, BoardName, mlCountReturn
   End Select
   mnDetailsNeedUpdate = False
   'If mnEvalEnabled Then SetRate mlRateReturned
   Select Case mnCounterType
      Case C9513
         LongHandle& = mvHandle
         NumChans% = mnNumEnabled - 1
         If mnFuncType = CB_STOREINT Then NumChans% = mnNum9513s - 1
         PlotBuffer LongHandle&, mlCount, NumChans%
      Case CIOTechScn, CIOTechQuad, C7266, CIOTechPulse
         'SetBufferType 2
         SetThenPlot = False
         If (Not mnuPlotType(PRINT_TEXT).Checked) And _
            (Not mnEncoder) Then SetPlotResolution mnResolution, False, SetThenPlot
         LongHandle& = mvHandle
         If mnFuncType = CB_CIN Then
            NumChans% = mnNumEnabled - 1
         Else
            NumChans% = mnLastChan - mnFirstChan
         End If
         
         CountMode$ = cmbCtrType.Text
         Ticks& = cmbTickSize.ListIndex + 1
         TickSize! = Choose(Ticks&, 0.00000002083, _
            0.0000002083, 0.0000020833, 0.0000208333, _
            0.00000002, 0.0000002, 0.000002, 0.00002)
         Select Case CountMode$
            Case "PERIOD"
               ConvertToEng mnBoardNum, True, TickSize!
            Case "PULSEWIDTH"
               ConvertToEng mnBoardNum, True, TickSize!
            Case "TIMING"
               ConvertToEng mnBoardNum, True, TickSize!
               'TextValue$ = CalcMeasurement(3)
            Case Else
               ConvertToEng mnBoardNum, False
         End Select
   
         If mnuContPlot.Checked Then
            LongHandle& = mvHandle
            PlotPartialBuf LongHandle&, (mlCount), mlFirstPoint, NumChans%
         Else
            If mnPlot Then PlotBuffer LongHandle&, mlCount, NumChans%
         End If
      Case Else
         If mnuContPlot.Checked Then
            LongHandle& = mvHandle
            PlotPartialBuf LongHandle&, (mlCount) \ 2, mlFirstPoint, mnLastChan - mnFirstChan
         Else
         LongHandle& = mvHandle
         NumChans% = mnNumEnabled - 1
         If mnPlot Then PlotBuffer LongHandle&, mlCount, NumChans%
         End If
   End Select
   DoEvents
   If mnDDEControl And mnDDEStat Then
      NumFuncs% = GetHistory() - 1
      ReDim MyArray(NumFuncs%, 14)
      GetHistoryArray MyArray()
      ArrayText$ = "History " & Chr$(13) & Chr$(10)
      For Func% = 0 To NumFuncs%
         For Arg% = 0 To 14
            ArrayText$ = ArrayText$ & Format$(MyArray(Func%, Arg%), "0\,")
         Next Arg%
         ArrayText$ = ArrayText$ & Chr$(13) & Chr$(10)
      Next
      lblDDEStatus.Caption = ArrayText$
   End If

End Sub

Private Sub cmdStopPls0_Click()

   mnCancel = True
   mnLoop = False
   For CurCtr% = 0 To mnNumEnabled - 1
      Tmr% = maCtrList(CurCtr%)
      Select Case mnLibType
         Case UNILIB
            ULStat = PlsOutStop(mnBoardNum, Tmr%)
            If SaveFunc(Me, PulseOutStop, ULStat, _
               mnBoardNum, Tmr%, A3, A4, A5, A6, _
               A7, A8, A9, A10, A11, 0) Then Exit Sub
            Me.lblTmr0Freq.Caption = "Off"
         Case MSGLIB
            ChanString$ = Format(Tmr%, "0")
            TmrMsg$ = "TMR{" & ChanString$ & "}:STOP"
            MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
            If SaveMsg(Me, "SendMessage(" & TmrMsg$ & _
               ")", MsgResult$) Then Exit Sub
            mnTmr0Running = False
            Me.lblTmr0Freq.Caption = "Off"
      End Select
   Next
   tmrGoLoop.ENABLED = False
   Me.cmdStartTmr0.FontBold = True
   
End Sub

Private Sub Form_Activate()

   UpdateMainStatus
   SetOptions

End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub Form_Initialize()
   
   mnLoading = True
   Me.mnuLibrary(MSGLIB).ENABLED = frmMain.mnuLibrary(MSGLIB).ENABLED
   Me.mnuLibrary(UNILIB).ENABLED = frmMain.mnuLibrary(UNILIB).ENABLED
   mlStatusType = CTRFUNCTION

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = 27 Then cmdStop = True
   Select Case Shift
      Case vbAltMask
         'If Shift = vbAltMask Then
         'if alt key pressed
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
      Case vbCtrlMask
         CtrCtl% = -1
         Select Case KeyCode
            Case vbKeyAdd
               For CtrCtl% = 0 To 19
                  chkCounter(CtrCtl%).value = 1
               Next
               CtrCtl% = -1
            Case vbKeySubtract
               For CtrCtl% = 0 To 19
                  chkCounter(CtrCtl%).value = 0
               Next
               CtrCtl% = -1
            Case Asc("0")
               CtrCtl% = 0
            Case Asc("1")
               CtrCtl% = 1
            Case Asc("2")
               CtrCtl% = 2
            Case Asc("3")
               CtrCtl% = 3
            Case Asc("4")
               CtrCtl% = 4
            Case Asc("5")
               CtrCtl% = 5
            Case Asc("6")
               CtrCtl% = 6
            Case Asc("7")
               CtrCtl% = 7
            Case Asc("8")
               CtrCtl% = 8
            Case Asc("9")
               CtrCtl% = 9
         End Select
         If Not (CtrCtl% < 0) Then
            CurVal% = chkCounter(CtrCtl%).value
            chkCounter(CtrCtl%).value = Switch((CurVal% = 1), 0, _
               (CurVal% = 0), 1)
         End If
      Case (vbCtrlMask Or vbShiftMask)
         CtrCtl% = -1
         Select Case KeyCode
            Case Asc("0")
               CtrCtl% = 10
            Case Asc("1")
               CtrCtl% = 11
            Case Asc("2")
               CtrCtl% = 12
            Case Asc("3")
               CtrCtl% = 13
            Case Asc("4")
               CtrCtl% = 14
            Case Asc("5")
               CtrCtl% = 15
            Case Asc("6")
               CtrCtl% = 16
            Case Asc("7")
               CtrCtl% = 17
            Case Asc("8")
               CtrCtl% = 18
            Case Asc("9")
               CtrCtl% = 19
         End Select
         If Not (CtrCtl% < 0) Then
            CurVal% = chkCounter(CtrCtl%).value
            chkCounter(CtrCtl%).value = Switch((CurVal% = 1), 0, _
               (CurVal% = 0), 1)
         End If
      Case Else
         Select Case KeyCode
            Case 118 'F7 - set default child size
               Me.Height = 3000
               Me.Width = 5600
            Case 120 'F9 - set height to 1/3 screen
               mfmUniTest.Height = Screen.Height / 3
            Case 122 'F11 - set to screen bottom
               mfmUniTest.Move 0, Screen.Height - mfmUniTest.Height, Screen.Width
            Case 123 'F12 - set to screen top
               mfmUniTest.Move 0, 0, Screen.Width
         End Select
   End Select

End Sub

Private Sub Form_Load()
   
#If MSGOPS Then
   If gnLibType = MSGLIB Then
      Set MsgLibrary = New MBDClass.MBDComClass
   End If
#Else
   Dim MsgLibrary As Object
#End If
   
   mnNumBoards = gnNumBoards
   mnLibType = gnLibType
   mnuLibrary(gnLibType).Checked = True
   If Not gbULLoaded Then
      Me.mnuULBuf.Checked = False
      mnuLibrary(UNILIB).Checked = False
      mnuLibrary(UNILIB).ENABLED = False
   End If
   If Not gbULLoaded Then Me.mnuULBuf.Checked = False
   If (gnLibType = UNILIB) And Not gbULLoaded Then
      mnuFunc.ENABLED = False
      cmdNext1.ENABLED = False
      mnuEvents.ENABLED = False
   End If
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
      msDisplayName = DisplayName$
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      Else
         mnBoardNum = BoardNum%
         mnBoardIndex = 0
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   mnCounterNum = 1
   fraLoad.Top = fraCounter.Top
   fraLoad.Left = fraCounter.Left
   picReadVal.Top = fraCounter.Top
   fraConfig.Top = fraCounter.Top
   fraChipInit.Left = fraCounter.Left
   fraCStat.Top = fraCounter.Top
   fraCStat.Left = fraCounter.Left
   msCtrList = "NO COUNTERS SELECTED"
   
   'added 5/2008 for script purposes
   '(no meaning except for data interpretation
   mnResolution = 16
   mnRange = BIP5VOLTS: mnFixedRange = BIP5VOLTS
   'end of meaningless variables
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   mnUnloading = True

End Sub

Private Sub Form_Resize()

   lblStatus.Width = ScaleWidth
   lblStatus.Top = ScaleHeight - lblStatus.Height
   If gnInitializing Then
      'if the form is just loading, sets the form type
      'and sets the default function (cbC8254Config())
      mnFormType = (Val("&H" & Tag) And &HF00&) / &H100
      mnThisInstance = Val("&H" & Tag) And &HFF
      'mnTimerState = True
      'mnuFuncArray_Click (SEL_CTR)
      mlCount = 1000
      'mlRate = 1000
      msTitle = Caption
      'mnuBoard_Click (0)
      msStatusType = "CTR"
      gnInitializing = False
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   mnLoop = False
   tmrCIn.ENABLED = False
   tmrCheckStatus.ENABLED = False
   'tmrGoLoop.Enabled = False (useless timer)
   
   Select Case mnLibType
      Case UNILIB
         If mnNumBoards > 0 Then ULStat = StopBackground520(mnBoardNum, mlStatusType)
         If mvHandle Then
            UseLibrary% = mnuULBuf.Checked
            If UseLibrary% Then
               ULStat = cbWinBufFree(mvHandle)
               If SaveFunc(Me, WinBufFree, ULStat, mvHandle, A2, _
                  A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               mvHandle = 0: mlAqcPoints = 0
               LongHandle& = mvHandle
               PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
            Else
               If FreeMemory(mvHandle) Then mvHandle = 0
            End If
         End If
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
   gnCTRForms = gnCTRForms - 1

End Sub

Private Sub fraChipInit_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText 0
   
End Sub

Private Sub hsbAlarm2_Change()

   txtAlarm2.Text = hsbAlarm2.value
   
End Sub

Private Sub hsbHoldVal_Change()

   HoldVal% = CInt(hsbHoldVal.value) Xor &H8000
   txtHoldVal.Text = IntValToULong(HoldVal%)

End Sub

Private Sub hsbLoadVal_Change()

   LoadVal% = CInt(hsbLoadVal.value) Xor &H8000
   txtLoadVal.Text = IntValToULong(LoadVal%)

End Sub

Private Sub hsbPrescaler_Change()

   If Not (mnCounterType = C7266) Then
      PreScalerVal% = hsbPrescaler.value
      txtPrescaler.Text = PreScalerVal%
   Else
      PreScalerVal% = CInt(hsbPrescaler.value) Xor &H8000
      txtPrescaler.Text = IntValToULong(PreScalerVal%)
   End If
   
End Sub

Private Sub InitCounters()

   On Error GoTo CheckRevision
   For CtrCheck% = 0 To 20
      Me.chkCounter(CtrCheck%).FontUnderline = False
      chkCounter(CtrCheck%).ToolTipText = "Invalid"
      chkCounter(CtrCheck%).Tag = 0
   Next CtrCheck%
   
   mnNum8254s = 0
   mnNum9513s = 0
   mnNum8536s = 0
   mnNum7266s = 0
   mnNumEventCtrs = 0
   mnNumIOTechTmrs = 0
   mnNumIOTechScns = 0
   mnNumIOTechPulsers = 0
   mnNumIOTechQuads = 0
   mnCounterType = 0
   Options& = mvOptions
      
   If mnLibType = UNILIB Then
      InfoType% = BOARDINFO
      ConfigItem% = BICINUMDEVS
      ULStat = GetConfig520(InfoType%, _
         mnBoardNum, DevNum%, ConfigItem%, ValConfig&)
      ConfigVal% = ValConfig&
      For DevNum% = 0 To ConfigVal% - 1
         InfoType% = COUNTERINFO
         ConfigItem% = CICTRTYPE
         ULStat = GetConfig520(InfoType%, _
            mnBoardNum, DevNum%, ConfigItem%, ValConfig&)
         Select Case ValConfig&
            Case C8254
               mnNum8254s = mnNum8254s + 1
               mnCounterType = C8254
               TypeString = "8254"
            Case C9513
               mnNum9513s = mnNum9513s + 1
               mnCounterType = C9513
               TypeString = "9513"
            Case C8536
               mnNum8536s = mnNum8536s + 1
               mnCounterType = C8536
               TypeString = "8536"
            Case C7266
               mnNum7266s = mnNum7266s + 1
               mnCounterType = C7266
               TypeString = "7266"
            'Case Cmini
            Case CEvent
               'mnNum8254s = mnNum8254s + 1
               mnNumEventCtrs = mnNumEventCtrs + 1
               mnCounterType = CEvent
               TypeString = "Event"
            Case CIOTechScn
               mnNumIOTechScns = mnNumIOTechScns + 1
               mnCounterType = CIOTechScn
               TypeString = "Scan"
            Case CIOTechTmr
               mnNumIOTechTmrs = mnNumIOTechTmrs + 1
               mnCounterType = CIOTechTmr
               TypeString = "Timer"
            Case CIOTechQuad
               mnNumIOTechQuads = mnNumIOTechQuads + 1
               mnCounterType = CIOTechQuad
               TypeString = "Quad"
            Case CIOTechPulse
               mnNumIOTechPulsers = mnNumIOTechPulsers + 1
               mnCounterType = CIOTechPulse
               TypeString = "Pulse"
            Case Else
               TypeString = ""
         End Select
         ConfigItem% = CICTRNUM
         ULStat = GetConfig520(InfoType%, _
            mnBoardNum, DevNum%, ConfigItem%, ValConfig&)
         If mnCounterType > 0 Then
            CounterNum% = ValConfig& '- mnCounterBase
            chkCounter(DevNum% + mnCounterBase).FontUnderline = True
            chkCounter(DevNum% + mnCounterBase).Tag = mnCounterType
            chkCounter(DevNum% + mnCounterBase).ToolTipText = TypeString
         End If
      Next DevNum%
   Else
      'to do - if msgdaq ever returns counter types, add here
      NumProps% = GetCtrProps(msBoardName, MsgLibrary, PropsList)
      UpdateFormProps PropsList, NumProps%
      For CounterNum% = 0 To mnNumEventCtrs - 1
         mnCounterType = CEvent
         chkCounter(CounterNum%).FontUnderline = True
         chkCounter(CounterNum%).Tag = mnCounterType
         TypeString = "Event"
         chkCounter(CounterNum%).ToolTipText = TypeString
      Next
      NumTProps% = GetTmrProps(msBoardName, MsgLibrary, PropsList)
      UpdateFormProps PropsList, NumTProps%
      For CounterNum% = mnNumEventCtrs _
         To mnNumEventCtrs + mnNumIOTechPulsers - 1
            mnCounterType = CIOTechPulse
            chkCounter(CounterNum%).FontUnderline = True
            chkCounter(CounterNum%).Tag = mnCounterType
            TypeString = "Pulse"
            chkCounter(CounterNum%).ToolTipText = TypeString
      Next
   End If
   LoadControls
   ConfigureControls
   If mnNumEventCtrs = 1 Then
      If IsNumeric(chkCounter(1).Tag) Then
         FirstECtr% = chkCounter(1).Tag
         If FirstECtr% = CEvent Then chkCounter(0).Tag = CEvent
      End If
   End If
   Exit Sub
      
   'commented out exit sub below and moved above
   'nothing seems to happen in following lines 6/9/09
   SaveErrFlow% = geErrFlow
   SaveLocalErrDisp% = gnLocalErrDisp
   geErrFlow = 0
   gnLocalErrDisp = False
   If mnLibType = UNILIB Then
      ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
      If SaveFunc(Me, ErrHandling, ULStat, _
         DONTPRINT, DONTSTOP, A3, A4, A5, A6, _
         A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
   geErrFlow = SaveErrFlow%
   gnLocalErrDisp = SaveLocalErrDisp%
   ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
   x% = SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, _
      gnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   'Exit Sub

   SaveErrFlow% = geErrFlow
   SaveLocalErrDisp% = gnLocalErrDisp
   geErrFlow = 0
   gnLocalErrDisp = False
   ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
   If SaveFunc(Me, ErrHandling, ULStat, DONTPRINT, _
      DONTSTOP, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
      Then Exit Sub
   mnNum8254s = 0
   mnNum9513s = 0
   mnNum8536s = 0
   mnNum7266s = 0
   mnNumIOTechTmrs = 0
   mnNumIOTechScns = 0
   mnCounterType = 0
   
   For i% = 0 To 15
      CtrNum% = i% + mnCounterBase
      ULStat = cbC8254Config(mnBoardNum, CtrNum%, HIGHONLASTCOUNT)
      If Not gnScriptSave Then
         x% = SaveFunc(Me, C8254Config, ULStat, _
            mnBoardNum, CtrNum%, HIGHONLASTCOUNT, _
            A4, A5, A6, A7, A8, A9, A10, A11, 0)
      End If
      If Not ((ULStat = BADCOUNTERDEVNUM) Or _
         (ULStat = NOT8254CTR) Or (ULStat = BADBOARDTYPE)) Then
            mnCounterType = C8254
            mnNum8254s = mnNum8254s + 1
      End If
   Next i%
   For i% = 0 To 20
      CtrNum% = i% + mnCounterBase
      ULStat = cbC9513Config(mnBoardNum, _
         CtrNum%, NOGATE, POSITIVEEDGE, CtrNum%, _
         0, LOADREG, RECYCLE, 0, COUNTDOWN, DISCONNECTED)
      If Not gnScriptSave Then
         x% = SaveFunc(Me, C9513Config, _
            ULStat, mnBoardNum, CtrNum%, NOGATE, _
            POSITIVEEDGE, CtrNum%, 0, LOADREG, _
            RECYCLE, 0, COUNTDOWN, DISCONNECTED, 0)
      End If
      If Not ((ULStat = BADCOUNTERDEVNUM) Or _
         (ULStat = NOT9513CTR) Or (ULStat = BADBOARDTYPE)) Then
            mnCounterType = C9513
            mnNum9513s = mnNum9513s + 1
      End If
   Next i%
   For i% = 0 To 6
      CtrNum% = i% + mnCounterBase
      ULStat = cbC8536Config(mnBoardNum, _
         CtrNum%, HIGHPULSEONTC, RECYCLE, 0)
      If Not gnScriptSave Then
         x% = SaveFunc(Me, C8536Config, ULStat, _
            mnBoardNum, CtrNum%, HIGHPULSEONTC, _
            RECYCLE, 0, A6, A7, A8, A9, A10, A11, 0)
      End If
      If Not ((ULStat = BADCOUNTERDEVNUM) Or _
         (ULStat = NOT8536CTR) Or (ULStat = BADBOARDTYPE)) Then
            mnCounterType = C8536
            mnNum8536s = mnNum8536s + 1
      End If
   Next i%
   For i% = 0 To 4
      CtrNum% = i% + mnCounterBase
      ULStat = cbC7266Config(mnBoardNum, CtrNum%, _
         NO_QUAD, NORMAL_MODE, BCD_ENCODING, INDEX_DISABLED, _
         InvertIndex&, CARRY_BORROW, GateEnable&)
      If Not gnScriptSave Then
         x% = SaveFunc(Me, C7266Config, ULStat, _
            mnBoardNum, CtrNum%, NO_QUAD, NORMAL_MODE, _
            BCD_ENCODING, INDEX_DISABLED, InvertIndex&, _
            CARRY_BORROW, GateEnable&, A10, A11, 0)
      End If
      If Not ((ULStat = BADCOUNTERDEVNUM) Or _
         (ULStat = NOT7266CTR) Or (ULStat = BADBOARDTYPE) _
         Or (ULStat = NOTWINDOWSFUNC)) Then
            mnCounterType = C7266
            mnNum7266s = mnNum7266s + 1
      End If
   Next i%
   For i% = 0 To 5
      CtrNum% = i% + mnCounterBase
      ULStat = CtrClear(mnBoardNum, CtrNum%)
      If Not gnScriptSave Then
         x% = SaveFunc(Me, CClear, ULStat, _
            mnBoardNum, CtrNum%, A3, A4, _
            A5, A6, A7, A8, A9, A10, A11, 0)
      End If
      If Not ((ULStat = NOCTRSTATUS) Or _
         (ULStat = BADCOUNTERDEVNUM) Or _
         (ULStat = NOT7266CTR) Or (ULStat _
         = BADBOARDTYPE) Or (ULStat = NOTWINDOWSFUNC)) Then
            mnCounterType = CIOTechScn
            mnNumIOTechScns = mnNumIOTechScns + 1
      End If
   Next i%
   For i% = 0 To 4
      CtrNum% = i% + mnCounterBase
      Frequency# = 1000
      Select Case mnLibType
         Case UNILIB
            ULStat = TmrOutStart(mnBoardNum, CtrNum%, Frequency#)
            If Not gnScriptSave Then
               x% = SaveFunc(Me, TimerOutStart, ULStat, mnBoardNum, _
               CtrNum%, Frequency#, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            End If
            If Not ((ULStat = NOCTRSTATUS) Or (ULStat = BADCOUNTERDEVNUM) _
            Or (ULStat = NOT7266CTR) Or (ULStat = BADBOARDTYPE) Or _
            (ULStat = NOTWINDOWSFUNC)) Then
               mnNumIOTechPulsers = mnNumIOTechPulsers + 1
            End If
            ULStat = TmrOutStop(mnBoardNum, CtrNum%)
         Case MSGLIB
            ChanString$ = Format(CtrNum%, "0")
            'PeriodVal# = (1 / Frequency#) * 1000
            TmrMsg$ = "@TMR{" & ChanString$ & "}:TYPE"
            MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
            If Not InStr(1, MsgResult$, "PULSE") = 0 Then
               mnNumIOTechPulsers = mnNumIOTechPulsers + 1
            End If
            mnNumIOTechPulsers = mnNumIOTechPulsers + 1
      End Select
      If Not ULStat = 0 Then Exit For
   Next i%
   For i% = 0 To 4
      CtrNum% = i% + mnCounterBase
      Frequency# = 1000
      DutyCycle# = 50
      PulseCount& = 1000
      Select Case mnLibType
         Case UNILIB
            ULStat = PlsOutStart(mnBoardNum, CtrNum%, Frequency#, DutyCycle#, _
            PulseCount&, InitialDelay#, IdleState&, Options&)
            If Not gnScriptSave Then
               x% = SaveFunc(Me, PulseOutStart, ULStat, mnBoardNum, _
               CtrNum%, Frequency#, DutyCycle#, _
               PulseCount&, InitialDelay#, IdleState&, Options&, A9, A10, A11, 0)
               If Not x% = 0 Then Exit For
            End If
            If Not ((ULStat = NOCTRSTATUS) Or (ULStat = BADCOUNTERDEVNUM) _
            Or (ULStat = NOT7266CTR) Or (ULStat = BADBOARDTYPE) Or _
            (ULStat = NOTWINDOWSFUNC)) Then
               mnNumIOTechPulsers = mnNumIOTechPulsers + 1
            End If
            ULStat = PlsOutStop(mnBoardNum, CtrNum%)
            If Not ULStat = 0 Then Exit For
         Case MSGLIB
            ChanString$ = Format(CtrNum%, "0")
            TmrMsg$ = "@TMR{" & ChanString$ & "}:DUTYCYCLE"
            MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
            If InStr(1, MsgResult$, "NOT_SUPPORTED") = 0 Then
               mnNumIOTechPulsers = mnNumIOTechPulsers + 1
            End If
      End Select
   Next i%
   
   geErrFlow = SaveErrFlow%
   gnLocalErrDisp = SaveLocalErrDisp%
   ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
   x% = SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, _
      gnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   
DoneDetecting:
   LoadControls
   ConfigureControls

Exit Sub

CheckRevision:
   If Err = 453 Then
      String1 = " This function was not included in the Universal Library revision you are using."
      String2 = " Continuing will not cause a problem unless you are using a board that requires this function."
      MsgBox Error$(Err) & String1 & String2, vbInformation, "Unsupported Function Called"
      Resume DoneDetecting
   Else
      MsgBox Error$(Err), vbCritical, "Crappy Program Bug"
   End If

End Sub

Private Sub lblMapChan_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""
   
End Sub

Private Sub lblMapChan_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  SetHelpText "txtMapChan"

End Sub

Private Sub lblQuery_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   lblQuery.Visible = False
   lblQuery.Drag
   
End Sub

Private Sub lblStatus_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""

End Sub

Private Sub LoadControls()

   cmbGateControl.Clear
   cmbCountSource.Clear
   cmbFOutDiv.Clear
   cmbFOutSource.Clear
   cmbTimeOfDay.Clear
   Me.txtLoadVal.Text = "10"

   Select Case mnCounterType
      Case C9513
         cmbGateControl.AddItem "NOGATE", NOGATE
         cmbGateControl.AddItem "AHLTCPREVCTR", AHLTCPREVCTR
         cmbGateControl.AddItem "AHLNEXTGATE", AHLNEXTGATE
         cmbGateControl.AddItem "AHLPREVGATE", AHLPREVGATE
         cmbGateControl.AddItem "AHLGATE", AHLGATE
         cmbGateControl.AddItem "ALLGATE", ALLGATE
         cmbGateControl.AddItem "AHEGATE", AHEGATE
         cmbGateControl.AddItem "ALEGATE", ALEGATE
         cmbGateControl.ListIndex = NOGATE
      
         cmbCountSource.AddItem "TCPREVCTR", TCPREVCTR
         cmbCountSource.AddItem "CTRINPUT1", CTRINPUT1
         cmbCountSource.AddItem "CTRINPUT2", CTRINPUT2
         cmbCountSource.AddItem "CTRINPUT3", CTRINPUT3
         cmbCountSource.AddItem "CTRINPUT4", CTRINPUT4
         cmbCountSource.AddItem "CTRINPUT5", CTRINPUT5
         cmbCountSource.AddItem "GATE1", GATE1
         cmbCountSource.AddItem "GATE2", GATE2
         cmbCountSource.AddItem "GATE3", GATE3
         cmbCountSource.AddItem "GATE4", GATE4
         cmbCountSource.AddItem "GATE5", GATE5
         cmbCountSource.AddItem "FREQ1", FREQ1
         cmbCountSource.AddItem "FREQ2", FREQ2
         cmbCountSource.AddItem "FREQ3", FREQ3
         cmbCountSource.AddItem "FREQ4", FREQ4
         cmbCountSource.AddItem "FREQ5", FREQ5
         cmbCountSource.ListIndex = FREQ3
      
         cmbFOutDiv.AddItem "16", 0
         cmbFOutDiv.AddItem "1", 1
         cmbFOutDiv.AddItem "2", 2
         cmbFOutDiv.AddItem "3", 3
         cmbFOutDiv.AddItem "4", 4
         cmbFOutDiv.AddItem "5", 5
         cmbFOutDiv.AddItem "6", 6
         cmbFOutDiv.AddItem "7", 7
         cmbFOutDiv.AddItem "8", 8
         cmbFOutDiv.AddItem "9", 9
         cmbFOutDiv.AddItem "10", 10
         cmbFOutDiv.AddItem "11", 11
         cmbFOutDiv.AddItem "12", 12
         cmbFOutDiv.AddItem "13", 13
         cmbFOutDiv.AddItem "14", 14
         cmbFOutDiv.AddItem "15", 15
         cmbFOutDiv.ListIndex = 1
      
         If mnFuncType = CB_FREQIN Then
            cmbFOutSource.AddItem "CTRINPUT1"
            cmbFOutSource.AddItem "CTRINPUT2"
            cmbFOutSource.AddItem "CTRINPUT3"
            cmbFOutSource.AddItem "CTRINPUT4"
            cmbFOutSource.AddItem "CTRINPUT5"
            cmbFOutSource.AddItem "GATE1"
            cmbFOutSource.AddItem "GATE2"
            cmbFOutSource.AddItem "GATE3"
            cmbFOutSource.AddItem "GATE4"
            cmbFOutSource.AddItem "FREQ1"
            cmbFOutSource.AddItem "FREQ2"
            cmbFOutSource.AddItem "FREQ3"
            cmbFOutSource.AddItem "FREQ4"
            cmbFOutSource.AddItem "FREQ5"
            If mnNum9513s > 5 Then
               cmbFOutSource.AddItem "CTRINPUT6"
               cmbFOutSource.AddItem "CTRINPUT7"
               cmbFOutSource.AddItem "CTRINPUT8"
               cmbFOutSource.AddItem "CTRINPUT9"
               cmbFOutSource.AddItem "CTRINPUT10"
               cmbFOutSource.AddItem "GATE6"
               cmbFOutSource.AddItem "GATE7"
               cmbFOutSource.AddItem "GATE8"
               cmbFOutSource.AddItem "GATE9"
               cmbFOutSource.AddItem "GATE10"
               cmbFOutSource.AddItem "FREQ6"
               cmbFOutSource.AddItem "FREQ7"
               cmbFOutSource.AddItem "FREQ8"
               cmbFOutSource.AddItem "FREQ9"
               cmbFOutSource.AddItem "FREQ10"
            End If
            If mnNum9513s > 10 Then
               cmbFOutSource.AddItem "CTRINPUT11"
               cmbFOutSource.AddItem "CTRINPUT12"
               cmbFOutSource.AddItem "CTRINPUT13"
               cmbFOutSource.AddItem "CTRINPUT14"
               cmbFOutSource.AddItem "CTRINPUT15"
               cmbFOutSource.AddItem "GATE11"
               cmbFOutSource.AddItem "GATE12"
               cmbFOutSource.AddItem "GATE13"
               cmbFOutSource.AddItem "GATE14"
               cmbFOutSource.AddItem "GATE15"
               cmbFOutSource.AddItem "FREQ11"
               cmbFOutSource.AddItem "FREQ12"
               cmbFOutSource.AddItem "FREQ13"
               cmbFOutSource.AddItem "FREQ14"
               cmbFOutSource.AddItem "FREQ15"
               cmbFOutSource.AddItem "CTRINPUT16"
               cmbFOutSource.AddItem "CTRINPUT17"
               cmbFOutSource.AddItem "CTRINPUT18"
               cmbFOutSource.AddItem "CTRINPUT19"
               cmbFOutSource.AddItem "CTRINPUT20"
               cmbFOutSource.AddItem "GATE16"
               cmbFOutSource.AddItem "GATE17"
               cmbFOutSource.AddItem "GATE18"
               cmbFOutSource.AddItem "GATE19"
               cmbFOutSource.AddItem "GATE20"
               cmbFOutSource.AddItem "FREQ16"
               cmbFOutSource.AddItem "FREQ17"
               cmbFOutSource.AddItem "FREQ18"
               cmbFOutSource.AddItem "FREQ19"
               cmbFOutSource.AddItem "FREQ20"
            End If
         Else
            cmbFOutSource.AddItem "CTRINPUT1", CTRINPUT1 - 1
            cmbFOutSource.AddItem "CTRINPUT2", CTRINPUT2 - 1
            cmbFOutSource.AddItem "CTRINPUT3", CTRINPUT3 - 1
            cmbFOutSource.AddItem "CTRINPUT4", CTRINPUT4 - 1
            cmbFOutSource.AddItem "CTRINPUT5", CTRINPUT5 - 1
            cmbFOutSource.AddItem "GATE1", GATE1 - 1
            cmbFOutSource.AddItem "GATE2", GATE2 - 1
            cmbFOutSource.AddItem "GATE3", GATE3 - 1
            cmbFOutSource.AddItem "GATE4", GATE4 - 1
            cmbFOutSource.AddItem "GATE5", GATE5 - 1
            cmbFOutSource.AddItem "FREQ1", FREQ1 - 1
            cmbFOutSource.AddItem "FREQ2", FREQ2 - 1
            cmbFOutSource.AddItem "FREQ3", FREQ3 - 1
            cmbFOutSource.AddItem "FREQ4", FREQ4 - 1
            cmbFOutSource.AddItem "FREQ5", FREQ5 - 1
         End If
         cmbFOutSource.ListIndex = FREQ1 - 1
      
         If mnFuncType = CB_FREQIN Then
            cmbTimeOfDay.AddItem "1", 0
            cmbTimeOfDay.AddItem "10", 1
            cmbTimeOfDay.AddItem "100", 2
            cmbTimeOfDay.AddItem "1000", 3
            cmbTimeOfDay.AddItem "10000", 4
            cmbTimeOfDay.ListIndex = 2
         Else
            cmbTimeOfDay.AddItem "DISABLED", 0
            cmbTimeOfDay.AddItem "50Hz Source", 1
            cmbTimeOfDay.AddItem "60Hz Source", 2
            cmbTimeOfDay.AddItem "100Hz Source", 3
            cmbTimeOfDay.ListIndex = 0
         End If
         cmbGateControl.ENABLED = True
         Me.cmbCountSource.ENABLED = True
      Case C8536
         cmbGateControl.AddItem "HW_START_TRIG", 0
         cmbGateControl.AddItem "HW_RETRIG", 1
         cmbGateControl.AddItem "SW_START_TRIG", 2
         cmbGateControl.ListIndex = 0
         cmbFOutSource.AddItem "NOTLINKED", NOTLINKED
         cmbFOutSource.AddItem "GATECTR2", GATECTR2
         cmbFOutSource.AddItem "TRIGCTR2", TRIGCTR2
         cmbFOutSource.AddItem "INCTR2", INCTR2
         cmbFOutSource.ListIndex = 0
      Case C7266
         cmbGateControl.AddItem "INDEXDISABLED", INDEX_DISABLED
         cmbGateControl.AddItem "LOADCTR", LOAD_CTR
         cmbGateControl.AddItem "LOADOUTLATCH", LOAD_OUT_LATCH
         cmbGateControl.AddItem "RESETCTR", RESET_CTR
         cmbGateControl.ListIndex = 0
         cmbCountSource.AddItem "CARRYBORROW", CARRY_BORROW - 1
         cmbCountSource.AddItem "COMPAREBORROW", COMPARE_BORROW - 1
         cmbCountSource.AddItem "CARRYBORROWUPDOWN", CARRYBORROW_UPDOWN - 1
         cmbCountSource.AddItem "INDEXERROR", INDEX_ERROR - 1
         cmbCountSource.ListIndex = 0
'      Case Cmini
'         cmbGateControl.Enabled = False
'         cmbFOutSource.Enabled = False
'         cmbCountSource.Enabled = False
      Case CEvent
         cmbGateControl.ENABLED = False
         cmbFOutSource.ENABLED = False
         cmbCountSource.ENABLED = False
         Me.txtLoadVal.Text = "0"
      Case CIOTechScn, CIOTechQuad
         cmbFOutDiv.AddItem "500ns", 0
         cmbFOutDiv.AddItem "1500ns", 1
         cmbFOutDiv.AddItem "3500ns", 2
         cmbFOutDiv.AddItem "7500ns", 3
         cmbFOutDiv.AddItem "15500ns", 4
         cmbFOutDiv.AddItem "31500ns", 5
         cmbFOutDiv.AddItem "63500ns", 6
         cmbFOutDiv.AddItem "127500ns", 7
         cmbFOutDiv.AddItem "100us", 8
         cmbFOutDiv.AddItem "300us", 9
         cmbFOutDiv.AddItem "700us", 10
         cmbFOutDiv.AddItem "1500us", 11
         cmbFOutDiv.AddItem "3100us", 12
         cmbFOutDiv.AddItem "6300us", 13
         cmbFOutDiv.AddItem "12700us", 14
         cmbFOutDiv.AddItem "25500us", 15
         cmbFOutDiv.AddItem "NONE", 16
         cmbFOutDiv.ListIndex = 16
         'cmbFOutSource.AddItem "
   End Select
   If (mnNumIOTechQuads > 0) Or (mnNumIOTechScns > 0) Then
      cmbFOutDiv.AddItem "500ns", 0
      cmbFOutDiv.AddItem "1500ns", 1
      cmbFOutDiv.AddItem "3500ns", 2
      cmbFOutDiv.AddItem "7500ns", 3
      cmbFOutDiv.AddItem "15500ns", 4
      cmbFOutDiv.AddItem "31500ns", 5
      cmbFOutDiv.AddItem "63500ns", 6
      cmbFOutDiv.AddItem "127500ns", 7
      cmbFOutDiv.AddItem "100us", 8
      cmbFOutDiv.AddItem "300us", 9
      cmbFOutDiv.AddItem "700us", 10
      cmbFOutDiv.AddItem "1500us", 11
      cmbFOutDiv.AddItem "3100us", 12
      cmbFOutDiv.AddItem "6300us", 13
      cmbFOutDiv.AddItem "12700us", 14
      cmbFOutDiv.AddItem "25500us", 15
      cmbFOutDiv.AddItem "NONE", 16
      cmbFOutDiv.ListIndex = 16
   End If
   cmbCtrType.Clear
   cmbCtrType.AddItem "TOTALIZE", 0
   cmbCtrType.AddItem "ENCODER", 1
   cmbCtrType.AddItem "PERIOD", 2
   cmbCtrType.AddItem "PULSEWIDTH", 3
   cmbCtrType.AddItem "TIMING", 4
   
   cmbCtrType.ListIndex = 0
   
   cmbTickSize.Clear
   cmbTickSize.AddItem "Tick 20.83ns"
   cmbTickSize.AddItem "Tick 208.3ns"
   cmbTickSize.AddItem "Tick 2083.3ns"
   cmbTickSize.AddItem "Tick 20833.3ns"
   cmbTickSize.AddItem "Tick 20ns"
   cmbTickSize.AddItem "Tick 200ns"
   cmbTickSize.AddItem "Tick 2000ns"
   cmbTickSize.AddItem "Tick 20000ns"
   cmbTickSize.ListIndex = 0

End Sub

Private Sub lstCountCfg_DragDrop(Source As Control, _
   x As Single, y As Single)

   SetHelpText ""
   
End Sub

Private Sub lstCountCfg_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)

   SetHelpText "lstCountCfg"
   
End Sub

Private Sub lstGateOut_DragDrop(Source As Control, _
   x As Single, y As Single)

   SetHelpText ""
   
End Sub

Private Sub lstGateOut_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)

   SetHelpText "lstGateOut"
   
End Sub

Private Sub lstModeChange_DragDrop(Source As Control, _
   x As Single, y As Single)

   SetHelpText ""
   
End Sub

Private Sub lstModeChange_DragOver(Source As Control, _
   x As Single, y As Single, State As Integer)

SetHelpText "lstModeChange"

End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash

End Sub

Private Sub mnuBoard_Click(Index As Integer)

   Dim EventConflict As Boolean
   Dim LookupBase As Boolean
   Dim eventForm As Form
   
   If Not (gnNumBoards > 0) Then Exit Sub
   If mnLibType = UNILIB Then
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
   If Not gnInitializing Then mnuBoard(mnBoardIndex).Checked = False
   Caption = msTitle & " Board " & mnuBoard(Index).Caption
   If mnNumBoards = 0 Then Exit Sub
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
         Component$ = "CTR"
         msCTRSupport = MsgLibrary.GetSupportedMessages(Component$)
      End If
      lpFileName$ = "UniBoards.ini"
      lpApplicationName$ = DisplayName$
      mnCounterBase = 0
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
      BusType$ = Left$(BoardName$, 4)
      Select Case BusType$
         Case "PCI-"
            mnCounterBase = 1
            If Left$(BoardName$, 6) = "PCI-25" Then mnCounterBase = 0
            If Left$(BoardName$, 8) = "PCI-DAS1" Then
               LookupBase = True
               lpDefault$ = "1"
            End If
         Case "CIO-", "PC10", "PPIO", "PCM-"
            mnCounterBase = 1
            LookupBase = False
         Case "PCIM", "PCIe", "PC-C"
            mnCounterBase = 1
            LookupBase = False
         Case Else
            LookupBase = True
            lpDefault$ = "0"
      End Select
      If LookupBase Then
         lpFileName$ = "UniBoards.ini"
         lpApplicationName$ = BoardName$
         lpKeyName$ = "CounterBase"
         nSize% = 4   'this will either be 0 or -1
         CustomParam$ = Space$(nSize%)
         StringSize% = GetPrivateProfileString(lpApplicationName$, _
            lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
         CustomParam$ = Left$(CustomParam$, StringSize%)
         mnCounterBase = Val(CustomParam$)
      End If
   End If
   msBoardName = BoardName$
   msDisplayName = DisplayName$
   
   lpKeyName$ = "CounterResolution"
   lpDefault$ = "1"   '1 for fixed, or higher for multiple resolutions
   nSize% = 4
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   mnCounterRes = Val(CustomParam$)
   If (Not gnInitializing) Then
      InitCounters
      ConfigureControls
   End If
   UpdateMainStatus
   A1 = msDisplayName
   FuncID% = SSetBoardName
   If gnScriptSave And (Not gnInitializing) Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, FuncID%, _
            FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   RefreshCtrList
   UpdateStatBar 0
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

Private Sub mnuClearOpts_Click()

   ResetOpts

End Sub

Private Sub mnuClose_Click()

   Unload Me

End Sub

Private Sub mnuExit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub mnuFuncArray_Click(Index As Integer)

   If (mnFuncType = CB_FREQIN) Or (Index = CB_FREQIN) Then Reload% = True
   mnuFuncArray(mnFuncIndex).Checked = False
   mnFuncIndex = Index
   'mnCounterType = Val(Me.chkCounter(0).Tag)
   Select Case mnCounterType
      Case C8254, C9513, C8536, C7266, CEvent   '1, 2, 3, 4 ,5
         mnFuncType = Choose(Index + 1, SEL_CTR, CB_CINIT, CB_CONFIG, _
         CB_CLOAD, CB_CIN, CB_STOREINT, CB_FREQIN, CB_CSTAT)
      Case CIOTechScn, CIOTechTmr, CIOTechQuad, CIOTechPulse '6, 7, 8, 9
         mnFuncType = Choose(Index + 1, SEL_CTR, CB_CCFGSCAN, CB_CCLEAR, _
         CB_CSCAN, CB_CIN, CB_STOREINT, CB_FREQIN, CB_CSTART)
      Case Else
         mnFuncType = Choose(Index + 1, SEL_CTR, CB_CINIT, CB_CONFIG, _
         CB_CLOAD, CB_CIN, CB_STOREINT, CB_FREQIN, CB_CSTAT)
   End Select
   msConfig = mnuFuncArray(mnFuncIndex).Caption
   mnuFuncArray(mnFuncIndex).Checked = True
   If (Not gnInitializing) Then
      ConfigureControls
      If Reload% Then LoadControls
   End If
   If (mnFuncType = CB_STOREINT) Then
      If mnCounterType = C9513 Then
         'turn off the Go timer
         'no longer save state and turn back on when returning to other functions
         If (mnuLoop(2).Checked) Then
            mnuLoop(2).Checked = False
         End If
      End If
   End If
   UpdateStatBar 0

End Sub

Private Sub mnuGetStatus_Click()

   mnuGetStatus.Checked = Not mnuGetStatus.Checked
   'If mnuGetStatus.Checked Then mnuGetSubSystemStatus.Checked = False
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

Private Sub mnuHexVals_Click()

   mnuHexVals.Checked = Not mnuHexVals.Checked
   mnHexVals = mnuHexVals.Checked
   
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
            If (mvOptions And BACKGROUND) = BACKGROUND Then
               MsgResult$ = MsgLibrary.StopBackgroundScan(msStatusType)
            Else
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
            End If
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
   Else
      Caption = msTitle & " Board " & mnuBoard(0).Caption
   End If
   gnNumBoards = mnNumBoards

End Sub

Private Sub mnuLoop_Click(Index As Integer)

   If mnTimerStop Then frmSetTimer.chkTimerStopBG = 1

   If mnDelayRestart Then
      frmSetTimer.chkDelayRestart.value = 1
      frmSetTimer.txtDelayTime.Text = Format(mlDelayTime, "0")
   End If
   
   If Index = 1 Then
      mnuLoop(2).Checked = False
      mnuLoop(Index).Checked = Not mnuLoop(Index).Checked
   Else
      mnuLoop(1).Checked = False
      If mnuLoop(2).Checked Then frmSetTimer.chkEnableTimer.value = 1
   End If
   mnLoop = mnuLoop(1).Checked
   
   If Index = 2 Then
      frmSetTimer.txtInterval.Text = tmrCIn.Interval
      frmSetTimer.optTimerMode(1).value = mnTimerTillCount
      frmSetTimer.Show 1
      LoopRate& = Val(frmSetTimer.txtInterval.Text)
      If frmSetTimer.optTimerMode(1).value Then
         mlCount = 1
         mnTimerTillCount = True
         If Not (mnuPlotType(SINGLE_POINT).Checked) Then mnuPlotType_Click (SINGLE_POINT)
      Else
         mlCount = Val(txtCount.Text)
         mnTimerTillCount = False
         If mnuPlotType(SINGLE_POINT).Checked Then mnuPlotType_Click (SINGLE_POINT)
      End If
      mnuLoop(2).Checked = (frmSetTimer.chkEnableTimer.value = 1)
      mnTimerStop = (frmSetTimer.chkTimerStopBG = 1)
      mnDelayRestart = (frmSetTimer.chkDelayRestart.value = 1)
      mlDelayTime = Val(frmSetTimer.txtDelayTime.Text)
      Unload frmSetTimer
      If LoopRate& > 0 Then
         tmrGoLoop.Interval = LoopRate&
         tmrCIn.Interval = LoopRate&
      Else
         mnuLoop(2).Checked = False
      End If
      'mnTimerState = mnuLoop(2).Checked
      If Not mnuLoop(2).Checked Then
         tmrGoLoop.ENABLED = False
         tmrCIn.ENABLED = False
         cmdGo.FontBold = True
      End If
   End If
   'cmdStop.Visible = mnuLoop(1).Checked Or mnuLoop(2).Checked

End Sub

Private Sub mnuPlotType_Click(Index As Integer)

   If Index = PRINT_TEXT Then
      mnuPlotType(Index).Checked = Not mnuPlotType(Index).Checked
      If mnuPlotType(Index).Checked Then mlBlockSave = GetBlockSize()
      If mnuPlotType(VOLTS_VS_TIME).Checked Or mnuPlotType(DERIVATIVE).Checked Then
         If Not mnuPlotType(PRINT_TEXT).Checked Then SetBlockSize mlBlockSave, False
      End If
      TFVal% = mnuPlotType(Index).Checked
      ShowText TFVal%
      
      x% = SaveFunc(Me, SShowText, 0, 0, TFVal%, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
      
      DisplayData
      Exit Sub
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
   
   'If (Index = VOLTS_VS_TIME) And Not gnInitializing Then
   '   If Not mnuPlotType(PRINT_TEXT).Checked Then SetBlockSize mlBlockSave, False
   'End If
   'If (Index = HISTOGRAM) And Not gnInitializing Then SetBlockSize mlBlockSave, False
   'If TFVal% Then
   '   mnPlotType = Index
   '   SetPlotType mnPlotType + mnHardCopy, Me
   '   mnPlot = True
   'End If
   If mnuPlotType(DERIVATIVE).Checked Then SetAutoScale 1
   DisplayData
   A1 = mnPlotType
   If gnScriptSave Then
      FuncStat = 0
      For ArgNum% = 1 To 14
         ArgVar = Choose(ArgNum%, Me.Tag, SPlotType, FuncStat, _
            A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
         If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
            PrintString$ = PrintString$ & ", "
         Else
            PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
         End If
      Next
      Print #2, PrintString$; Format$(AuxHandle, "0")
   End If
   Exit Sub
   
   mnPlot = False
   If Index = mnPlotType Then
      mnuPlotType(mnPlotType).Checked = Not mnuPlotType(mnPlotType).Checked
   Else
      mnuPlotType(mnPlotType).Checked = False
   End If
   mnPlotType = Index
   mnuPlotType(mnPlotType).Checked = True
   
   If mnuPlotType(mnPlotType).Checked Then
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

Private Sub mnuStatusRate_Click()

   Default$ = Format(Me.tmrCheckStatus.Interval, "0")
   SelRate$ = InputBox("Milliseconds between status checks:", "Status Rate", Default$)
   If Not SelRate$ = "" Then
      tmrCheckStatus.Interval = Val(SelRate$)
   End If

End Sub

Private Sub mnuStopBG_Click()

   mnuStopBG.Checked = Not mnuStopBG.Checked

End Sub

Private Sub mnuULBuf_Click()

   mnuULBuf.Checked = Not (mnuULBuf.Checked)
   
End Sub

Private Sub optConfig_Click(Index As Integer)

   If Index = 6 Then
      If mnOutCtl Then
         Select Case mnCounterType
            Case C8254
               optCtl% = Choose(mnOutCtl + 1, 0, 1, 2, 3, 4, 5)
            Case C9513
               optCtl% = Choose(mnOutCtl + 1, 0, 1, 2, 0, 3, 4)
            Case C8536
               optCtl% = Choose(mnOutCtl + 1, 0, 1, 2, 0, 0, 0, 3)
            Case C7266
               optCtl% = Choose(mnOutCtl + 1, 0, 1, 2, 3)
         End Select
      End If
   
      optConfig(optCtl%).value = True
      Exit Sub
   End If
   Select Case mnCounterType
      Case C8254
         mnOutCtl = Choose(Index + 1, HIGHONLASTCOUNT, ONESHOT, _
            RATEGENERATOR, SQUAREWAVE, SOFTWARESTROBE, HARDWARESTROBE)
      Case C9513
         mnOutCtl = Choose(Index + 1, ALWAYSLOW, _
            HIGHPULSEONTC, TOGGLEONTC, DISCONNECTED, LOWPULSEONTC)
         chkClearOutput.Visible = mnuFuncArray(CB_CONFIG).Checked _
            And (Index = TOGGLEONTC)
      Case C8536
         mnOutCtl = Choose(Index + 1, HIGHPULSEONTC, TOGGLEONTC, HIGHUNTILTC)
         chkRecycle.Visible = True
         chkRetrigger.Visible = True
      Case C7266
         mnOutCtl = Choose(Index + 1, NORMAL_MODE, RANGE_LIMIT, NO_RECYCLE, MODULO_N)
   End Select
   
End Sub

Private Function ReadRegisters(RegSelect As Integer) As String

   ConfItemString$ = Me.cmbRegSelector.Text
   If IsNumeric(ConfItemString$) Then
      RegSelected& = -1
      RegIndex& = 8
   Else
      RegIndex& = Me.cmbRegSelector.ListIndex + 1
      If RegIndex& = 0 Then
         MsgBox "Select A Register", vbOKOnly, "Select A Register"
      Else
         RegSelected& = Choose(RegIndex&, LOADREG0, HOLDREG1, COUNT1, _
            PRESET1, PRESCALER1, MINLIMITREG0, MAXLIMITREG0, OUTPUTVAL0REG0, _
            OUTPUTVAL1REG0, ALARM1CHIP1, ALARM2CHIP1, ALARM1CHIP2, ALARM2CHIP2, _
            ALARM1CHIP3, ALARM2CHIP3, ALARM1CHIP4, ALARM2CHIP4)
      End If
   End If

   If RegSelect = -1 Then
      LastRegIndex& = mnNumEnabled - 1
      If RegIndex& > 7 Then LastRegIndex& = 0
   Else
      'requires only one counter selected
      LastRegIndex& = 0
   End If
   
   For CtrControl% = 0 To LastRegIndex&
      CtrNum% = maCtrList(CtrControl%)
      Select Case RegSelected&
         Case -1
            RegNum& = Val(ConfItemString$)
         Case LOADREG0
            'CounterBase& = mnCounterBase
            RegNum& = CtrNum%
         Case HOLDREG1
            RegNum& = HOLDREG1 + (CtrNum% - mnCounterBase)
         Case COUNT1
            RegNum& = COUNT1 + (CtrNum% - mnCounterBase)
         Case PRESET1
            RegNum& = PRESET1 + (CtrNum% - mnCounterBase)
         Case PRESCALER1
            RegNum& = PRESCALER1 + (CtrNum% - mnCounterBase)
         Case MINLIMITREG0
            RegNum& = MINLIMITREG0 + (CtrNum% - mnCounterBase)
         Case MAXLIMITREG0
            RegNum& = MAXLIMITREG0 + (CtrNum% - mnCounterBase)
         Case OUTPUTVAL0REG0
            RegNum& = OUTPUTVAL0REG0 + (CtrNum% - mnCounterBase)
         Case OUTPUTVAL1REG0
            RegNum& = OUTPUTVAL1REG0 + (CtrNum% - mnCounterBase)
         Case Else
            LastRegIndex& = 0
            RegNum& = RegSelected&
      End Select
   Next
   
   'ULStat = cbCIn32(mnBoardNum, RegNum&, lngCount&)
   ULStat = cbCIn64(mnBoardNum, RegNum&, CBCountVal@)
   If SaveFunc(Me, CIn64, ULStat, mnBoardNum, CtrNum%, _
      CBCountVal@, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
      Then Failure% = True
   CMult# = 10000
   DblVal# = Fix(CBCountVal@)
   TempDbl# = (CBCountVal@ - DblVal#) * CMult#
   UStringVal$ = Format((DblVal#), "#")
   If DblVal# = 0 Then
      LStringVal$ = Format(TempDbl#, "0")
   Else
      LStringVal$ = Format(Abs(TempDbl#), "0000")
   End If
   StringVal$ = UStringVal$ & LStringVal$
   ReadRegisters = StringVal$

End Function

Private Sub ReadCounters()
   
   If mnNumEnabled = 0 Then Exit Sub
   CountType% = cmbCtrType.ListIndex - 1
   mnPlot = False
   Bit32% = (chkLoad32.value = 1) Xor (Me.chkLoad64.value = 1)
   Bit64% = (chkLoad32.value = 1) And (Me.chkLoad64.value = 1)
   Bit48% = (chkLoad32.value = 1) And (Me.chkLoad64.value = 1) And mbBit48
   'If Not Bit32% Then Bit32% = Me.mnuOpt(6).Checked
   'Bit64% = Me.mnuOpt(8).Checked
   mnResolution = 16
   If Bit32% Then mnResolution = 32
   If Bit48% Then mnResolution = 48
   If Bit64% Then mnResolution = 64
   If mnLibType = MSGLIB Then
      mnResolution = 32
      Bit32% = True
   End If
   
   Select Case mnResolution
      Case 16
         Dim IntArray() As Integer
      Case 32
         Dim LongArray() As Long
      Case 64, 48
         Dim VariantArray() As Variant
         CMult# = 10000
   End Select
   
   mlCount = Val(Me.txtCount.Text)
   If mlCount = 0 Then mlCount = 1
   
   If mlCount < 0 Then
      'used for scripting repeated separate reads
      'without loosing data from last read
      'number of reads stored in last element
      LoopCount& = 1
      AppendData% = True
      NumSamples& = Abs(mlCount) * mnNumEnabled
      TotalCount& = NumSamples& + 1
      SamplesPerCtr& = Abs(mlCount)
      ChanElement% = mnNumEnabled - 1
      If Bit48% Then
         SampSize& = SamplesPerCtr& * 2
         If SampSize& = 0 Then SampSize& = 1
         ReDim LongArray(ChanElement%, SampSize&)
      Else
         ReDim LongArray(ChanElement%, SamplesPerCtr&)
      End If
      If Not IsEmpty(mvHandle) Then
         If UseLibrary% Then
            Select Case mnResolution
               Case 16
                  ErrCode& = cbWinBufToArray(mvHandle, _
                     IntArray(0, 0), FirstPoint&, NumSamples& + 1)
                  SampleSet& = IntArray(ChanElement%, SamplesPerCtr&)
                  If SampleSet& = SamplesPerCtr& Then SampleSet& = 0
               Case 32
                  ErrCode& = cbWinBufToArray32(mvHandle, _
                     LongArray(0, 0), FirstPoint&, NumSamples& + 1)
                  SampleSet& = LongArray(ChanElement%, SamplesPerCtr&)
                  If SampleSet& = SamplesPerCtr& Then SampleSet& = 0
               Case 64, 48
            End Select
         Else
         End If
      Else
         SampleSet& = 0
      End If
   Else
      LoopCount& = mlCount
      Select Case mnResolution
         Case 64, 48
            ReDim VariantArray(mnNumEnabled - 1, LoopCount& - 1)
         Case 32
            SampSize& = (LoopCount& * 2) - 1
            If SampSize& = 0 Then SampSize& = 1
            ReDim LongArray(mnNumEnabled - 1, SampSize&)
         Case Else
            ReDim IntArray(mnNumEnabled - 1, LoopCount& - 1)
      End Select
      TotalCount& = mnNumEnabled * LoopCount&
   End If
   mnLoop = True
   
   If mvHandle Then
      If UseLibrary% Then
         ULStat = cbWinBufFree(mvHandle)
         If SaveFunc(Me, WinBufFree, ULStat, mvHandle, _
            A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mvHandle = 0: mlAqcPoints = 0
      Else
         If FreeMemory(mvHandle) Then
            mvHandle = 0: mlAqcPoints = 0
            mnAcqDataType = vbEmpty
         End If
      End If
   End If
   
   NumEnabled% = mnNumEnabled - 1
   Me.cmdGo.ENABLED = False
   If Not (tmrCIn.ENABLED Or gnScriptRun) Then
      If (mnLibType = MSGLIB) And (chkManStart.value = 0) Then
         For CtrIndex% = 0 To NumEnabled%
            CtrNum% = maCtrList(CtrIndex%)
            Failure% = ArmCounter(CtrNum%, True)
         Next
      End If
      If Failure% Then Exit Sub
   End If
   Do
      For CtrIndex% = 0 To NumEnabled%
         CtrNum% = maCtrList(CtrIndex%)
         If mnFuncType = CB_CSTAT Then
            ULStat = cbCStatus(mnBoardNum, CtrNum%, StatusBits&)
            If SaveFunc(Me, CStatus, ULStat, mnBoardNum, CtrNum%, _
               StatusBits&, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
               cmdStop = True
               Exit Sub
            End If
            For BitWeight% = 0 To 6
               CtrIndex% = CtrNum% * 7
               BitMask& = Choose(BitWeight% + 1, C_UNDERFLOW, C_OVERFLOW, _
                  C_COMPARE, C_SIGN, C_ERROR, C_UP_DOWN, C_INDEX)
               chkCStat(BitWeight% + CtrIndex%) = Abs((StatusBits& _
                  And 2 ^ BitWeight%) = BitMask&)
            Next BitWeight%
         Else
            Select Case True
               Case Bit64%, Bit48%
                  ULStat = cbCIn64(mnBoardNum, CtrNum%, CBCountVal@)
                  If SaveFunc(Me, CIn64, ULStat, mnBoardNum, CtrNum%, _
                     CBCountVal@, A4, A5, A6, A7, A8, A9, A10, A11, 0) _
                     Then Failure% = True
                  CountString$ = GetStringValueFromCur(CBCountVal@)
                  CMult# = 10000
                  'VarVal = CBCountVal@ - 4294967296# 'Fix(CBCountVal@)
                  VariantArray(CtrIndex%, BiSample& + SampleSet&) = CBCountVal@ * CDec(10000) 'WholeVal
               'Case Bit48%
               '   ULStat = cbCIn64(mnBoardNum, CtrNum%, CBCountVal@)
               '   If SaveFunc(Me, CIn64, ULStat, mnBoardNum, CtrNum%, CBCountVal@, _
               '      A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
               '   BiSample& = LoopNum& * 2
               '   VariantArray(CtrIndex%, BiSample& + SampleSet&) = CBCountVal@ * CMult#
               Case Bit32%
                  Select Case mnLibType
                     Case UNILIB
                        ULStat = cbCIn32(mnBoardNum, CtrNum%, CBCount&)
                        If SaveFunc(Me, CIn32, ULStat, mnBoardNum, CtrNum%, CBCount&, _
                           A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
                     Case MSGLIB
                        CtrString$ = Format(CtrNum%, "0")
                        CtrReadMsg$ = "?CTR{" & CtrString$ & "}:VALUE"
                        MsgResult$ = MsgLibrary.SendMessage(CtrReadMsg$)
                        If SaveMsg(Me, "SendMessage(" & CtrReadMsg$ & _
                           ")", MsgResult$) Then Failure% = True
                        MeasLoc& = InStr(1, MsgResult$, "=")
                        If MeasLoc& > 0 Then
                           CBCount& = Val(Mid(MsgResult$, MeasLoc& + 1))
                        End If
                  End Select
                  If Failure% Then
                     cmdStop = True
                     Exit Sub
                  End If
                  LongArray(CtrIndex%, LoopNum& + SampleSet&) = CBCount&
               Case Else
                  ULStat = CIn520(mnBoardNum, CtrNum%, CBCount&)
                  If SaveFunc(Me, CIn, ULStat, mnBoardNum, CtrNum%, CBCount&, _
                     A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Failure% = True
                  If Failure% Then
                     cmdStop = True
                     Exit Sub
                  End If
                  If mnLibType = UNILIB Then
                     IntData% = CBCount& 'ULongValToInt(CBCount&)
                     IntArray(CtrIndex%, LoopNum& + SampleSet&) = IntData%
                     'IntValToULong(IntData%)
                  Else
                     IntArray(CtrIndex%, LoopNum& + SampleSet&) = CBCount&
                  End If
            End Select
         End If
      Next CtrIndex%
      DoEvents
      LoopNum& = LoopNum& + 1
      If LoopNum& = LoopCount& Then mnLoop = False
   Loop While mnLoop
   
   If (LoopCount& > 1) Or (gnScriptRun) Then
      If Not (tmrCIn.ENABLED Or gnScriptRun) Then
         If (mnLibType = MSGLIB) And (chkManStart.value = 0) Then
            For CtrIndex% = 0 To NumEnabled%
               CtrNum% = maCtrList(CtrIndex%)
               Failure% = ArmCounter(CtrNum%, False)
            Next
         End If
      End If
      If AppendData% Then
         SampleSet& = SampleSet& + 1
         LongArray(ChanElement%, SamplesPerCtr&) = SampleSet&
      End If
      Select Case mnResolution
         Case 16
            mvHandle = BufAlloc16(Me, TotalCount&)
            SourceRef& = mvHandle
            mnAcqDataType = vbInteger: mlAqcPoints = TotalCount&
            CopyMemory ByVal SourceRef&, IntArray(0, 0), TotalCount& * 2
            mnPlot = True
         Case 32, 48
            'If Not Bit48% Then
            FirstPoint& = 0
            'mvHandle = cbWinBufAlloc32(TotalCount&)
            mvHandle = BufAlloc32(Me, TotalCount&)
            SourceRef& = mvHandle
            mnAcqDataType = vbLong: mlAqcPoints = TotalCount&
            CopyMemory ByVal SourceRef&, LongArray(0, 0), TotalCount& * 4
            mnPlot = True
         Case Else
            FirstPoint& = 0
            mvHandle = cbWinBufAlloc32(TotalCount& * 2)
            SourceRef& = mvHandle
            mnAcqDataType = vbDecimal: mlAqcPoints = TotalCount&
            CopyMemory ByVal SourceRef&, LongArray(0, 0), TotalCount& * 8
            mnPlot = True
      End Select
   End If
   
   picReadVal.Cls
   picReadVal.CurrentY = 1
   x% = 0: y% = 0
   For CtrIndex% = 0 To NumEnabled%
      CtrNum% = maCtrList(CtrIndex%)
      y% = y% + 1
      Select Case mnResolution
         Case 64, 48
            VarValue = VariantArray(CtrIndex%, LoopCount& - 1)
            TextValue$ = Format(VarValue, "0") 'StringVal$ '
            'If Not (TextValue$ = CountString$) Then
               'gap converting currency to double between
               '9223372036854770000 and 9223372036854775807
            '   Diff& = Val(Right(TextValue$, 4)) - Val(Right(CountString$, 4))
            '   TextValue$ = CountString$
            '   If mnHexVals Then
            '      TempValue$ = GetHexValue(VarValue, mnResolution)
            '      CompValue$ = GetHexValue(Abs(Diff&), mnResolution)
            '      TempVal# = CDbl("&H" & Right(TempValue$, 4)) _
            '         - CDbl("&H" & CompValue$)
            '      EndVal$ = GetHexValue(TempVal#, mnResolution)
            '      TextValue$ = Left(TempValue$, Len(TempValue$) - 4) & EndVal$
            '   End If
            'Else
               If mnHexVals Then
                  HexValue$ = GetHexValue(VarValue, mnResolution)
                  TextValue$ = "0x" & HexValue$ & _
                     " (" & Len(HexValue$) & ")"
                  If Len(HexValue$) <> 16 Then Stop
               End If
            'End If
         Case 32
            VarValue = LongArray(CtrIndex%, LoopCount& - 1)
            TextValue$ = Format(VarValue, "0")
            If mnHexVals Then TextValue$ _
               = GetHexValue(VarValue, mnResolution)
            If y% = 6 Then
               y% = 0: x% = x% + 1800
               picReadVal.CurrentY = 0
            End If
         Case Else
            CtrVal& = IntArray(CtrIndex%, LoopCount& - 1)
            TextValue$ = Format(CtrVal&, "0")
            If mnHexVals Then TextValue$ _
               = GetHexValue(CtrVal&, mnResolution)
            If y% = 6 Then
               y% = 0: x% = x% + 1200
               picReadVal.CurrentY = 0
            End If
      End Select
      If CountType% > 0 Then
         ValueRead = Val(TextValue$)
         TextValue$ = CalcMeasurement(CountType%, ValueRead)
      End If
      picReadVal.CurrentX = x%
      picReadVal.Print CtrNum%; ") "; TextValue$
   Next
   If mnPlot Then DisplayData
   Me.cmdGo.ENABLED = True

End Sub

Private Sub StartCounter()

   MaxCounters% = 20    'maximum number
   mlCount = Val(txtCount.Text)
   IntCount& = mlCount
   NumCounters% = mnNum9513s
   If IntCount& > 0 Then ReDim CounterData%(NumCounters% - 1, IntCount& - 1)
   ReDim CntrControl%(NumCounters%)
   UseLibrary% = mnuULBuf.Checked

   If mvHandle Then
      If UseLibrary% Then
         ULStat = cbWinBufFree(mvHandle)
         If SaveFunc(Me, WinBufFree, ULStat, mvHandle, A2, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mvHandle = 0: mlAqcPoints = 0
         LongHandle& = mvHandle
         PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
      Else
         If FreeMemory(mvHandle) Then mvHandle = 0
      End If
   End If
   BufferSize& = NumCounters%
   If IntCount& > 1 Then BufferSize& = NumCounters% * IntCount&

   If UseLibrary% Then
      mvHandle = cbWinBufAlloc(BufferSize&)
      If SaveFunc(Me, WinBufAlloc, mvHandle, BufferSize&, A2, _
         A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      mnAcqDataType = vbInteger: mlAqcPoints = BufferSize&
   Else
      mvHandle = AllocateMemory(BufferSize&)
   End If
   
   mnCountersEnabled = 0
   For i% = 0 To NumCounters% - 1
      CntrControl%(i%) = chkCounter(i%).value
      If chkCounter(i%).value <> 0 Then mnCountersEnabled = i%
      If gnScriptSave Then
         If chkCounter(i%).value Then
            CtrBitField& = (2 ^ i%) Or CtrBitField&
         End If
      End If
   Next i%
   'following used by script
   A5 = CtrBitField&
   
   ReDim CounterCtl(20) As Integer
   ULStat = cbCStoreOnInt(mnBoardNum, IntCount&, CntrControl%(0), mvHandle)
   If SaveFunc(Me, CStoreOnInt, ULStat, mnBoardNum, _
      IntCount&, CntrControl%(0), mvHandle, A5, A6, _
      A7, A8, A9, A10, A11, mvAuxHandle) Then Exit Sub
   If mnuGetStatus.Checked Then
      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, CTRFUNCTION)
      If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, _
         CurIndex&, CTRFUNCTION, A6, A7, A8, A9, A10, A11, 0) Then
         ULStat = StopBackground520(mnBoardNum, CTRFUNCTION)
         If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, _
            CTRFUNCTION, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         tmrCheckStatus.ENABLED = False
         cmdStop.ENABLED = tmrCIn.ENABLED Or (chkManStart.value = 1)
         DoEvents
         mnPlot = True
         Exit Sub
      End If
   End If
   If (Status% = RUNNING) Then
      lblDDEStatus.Caption = "Background Started"
      tmrCheckStatus.ENABLED = True
      cmdStop.ENABLED = tmrCIn.ENABLED Or (chkManStart.value = 1)
   Else
      mnPlot = True
      DisplayData
   End If

End Sub

Private Sub StartScan()

   IntCount& = Val(txtHoldVal.Text)
   IntRate& = Val(txtLoadVal.Text)
   mlCount = IntCount&
   mlRate = IntRate&
   mnFirstChan = Val(txtPrescaler.Text)
   mnLastChan = Val(txtAlarm2.Text)
   mlStatusType = CTRFUNCTION
   mnCounterType = chkCounter(mnFirstChan).Tag  ' + mnCounterBase
   mnResolution = 16
   Bit32% = (chkLoad32.value = 1) Xor (Me.chkLoad64.value = 1)
   If Not Bit32% Then Bit32% = Me.mnuOpt(6).Checked
   Bit48% = (chkLoad32.value = 1) And (Me.chkLoad64.value = 1)
   Bit64% = Me.mnuOpt(8).Checked
   If Bit32% Then mnResolution = 32
   If Bit48% Then mnResolution = 48
   If Bit64% Then mnResolution = 64

   NumCounters% = mnNumIOTechScns
   If NumCounters% = 0 Then
      MsgBox "No counters selected.", , "NumCounters Invalid"
   End If
   UseLibrary% = mnuULBuf.Checked
        
   If mvHandle Then
      If UseLibrary% Then
         ULStat = cbWinBufFree(mvHandle)
         If SaveFunc(Me, WinBufFree, ULStat, mvHandle, _
            A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mvHandle = 0: mnAcqDataType = vbEmpty
         LongHandle& = mvHandle: mlAqcPoints = 0
         PlotBuffer LongHandle&, mlCount, mnLastChan - mnFirstChan
      Else
         If FreeMemory(mvHandle) Then
            mvHandle = 0: mlAqcPoints = BufferSize&
         End If
      End If
   End If
   
   If mnCounterType = C9513 Then
      BufferSize& = NumCounters% * IntCount&
   Else
      BufferSize& = IntCount&
   End If
   If UseLibrary% Then
      Select Case mnResolution
         Case 64
            mvHandle = WBufAlloc64(BufferSize&)
            If SaveFunc(Me, WinBufAlloc64, mvHandle, BufferSize&, _
               A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            mnAcqDataType = vbVariant: mlAqcPoints = BufferSize&
         Case 48
            mvHandle = WBufAlloc64(BufferSize&)
            If SaveFunc(Me, WinBufAlloc64, mvHandle, BufferSize&, _
               A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            mnAcqDataType = vbDecimal: mlAqcPoints = BufferSize&
         Case 32
            mvHandle = WBufAlloc32(BufferSize&)
            If SaveFunc(Me, WinBufAlloc32, mvHandle, BufferSize&, _
               A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            mnAcqDataType = vbLong: mlAqcPoints = BufferSize&
         Case Else
            mvHandle = cbWinBufAlloc(BufferSize&)
            If SaveFunc(Me, WinBufAlloc, mvHandle, BufferSize&, _
               A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            mnAcqDataType = vbInteger: mlAqcPoints = BufferSize&
      End Select
   Else
      mvHandle = AllocateMemory(BufferSize&)
   End If
   
   Options& = mvOptions

   ULStat = CtrInScan(mnBoardNum, mnFirstChan, _
      mnLastChan, IntCount&, IntRate&, mvHandle, Options&)
   If SaveFunc(Me, CInScan, ULStat, mnBoardNum, mnFirstChan, _
      mnLastChan, IntCount&, IntRate&, mvHandle, Options&, _
      A8, A9, A10, A11, mvAuxHandle) Then Exit Sub
   mlRateReturned = IntRate&
   mlCountReturn = IntCount&
   If mnuGetStatus.Checked Then
      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, CTRFUNCTION)
      If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, Status%, CurCount&, _
         CurIndex&, CTRFUNCTION, A6, A7, A8, A9, A10, A11, 0) Then
         ULStat = StopBackground520(mnBoardNum, CTRFUNCTION)
         If SaveFunc(Me, StopBackground, ULStat, mnBoardNum, CTRFUNCTION, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         tmrCheckStatus.ENABLED = False
         cmdStop.ENABLED = tmrCIn.ENABLED
         DoEvents
         mnPlot = True
         Exit Sub
      End If
   End If
   If Status% = RUNNING Then
      lblDDEStatus.Caption = "Background Started"
      tmrCheckStatus.ENABLED = True
      cmdStop.Visible = True
      cmdStop.ENABLED = True
   Else
      mnPlot = True
      DisplayData
   End If

End Sub

Private Sub StartFreq()

   If cmbFOutSource.ListIndex = -1 Then
      SigSource% = Val(cmbFOutSource.Text)
   Else
      TempVal = Choose(cmbFOutSource.ListIndex + 1, _
         CTRINPUT1, CTRINPUT2, CTRINPUT3, CTRINPUT4, _
         CTRINPUT5, GATE1, GATE2, GATE3, GATE4)
      If IsNull(TempVal) Then TempVal = _
         Choose(cmbFOutSource.ListIndex - 8, FREQ1, _
         FREQ2, FREQ3, FREQ4, FREQ5, CTRINPUT6, _
         CTRINPUT7, CTRINPUT8, CTRINPUT9, CTRINPUT10)
      If IsNull(TempVal) Then TempVal = _
         Choose(cmbFOutSource.ListIndex - 18, _
         GATE6, GATE7, GATE8, GATE9, GATE10, _
         FREQ6, FREQ7, FREQ8, FREQ9, FREQ10)
      If IsNull(TempVal) Then TempVal = _
         Choose(cmbFOutSource.ListIndex - 28, _
         CTRINPUT11, CTRINPUT12, CTRINPUT13, _
         CTRINPUT14, CTRINPUT15, GATE11, _
         GATE12, GATE13, GATE14, GATE15)
      If IsNull(TempVal) Then TempVal = _
         Choose(cmbFOutSource.ListIndex - 38, _
         FREQ11, FREQ12, FREQ13, FREQ14, _
         FREQ15, CTRINPUT16, CTRINPUT17, _
         CTRINPUT18, CTRINPUT19, CTRINPUT20)
      If IsNull(TempVal) Then TempVal = _
         Choose(cmbFOutSource.ListIndex - 48, _
         GATE16, GATE17, GATE18, GATE19, GATE20, _
         FREQ16, FREQ17, FREQ18, FREQ19, FREQ20)
      If IsNull(TempVal) Then
         MsgBox "Error determining source from listbox.", _
            , "Listbox Error"
         Exit Sub
      End If
      SigSource% = Val(TempVal)
   End If
   GateInterval% = Val(cmbTimeOfDay.Text)

   Do
      ULStat = cbCFreqIn(mnBoardNum, SigSource%, GateInterval%, CBCount%, Freq&)
      If SaveFunc(Me, CFreqIn, ULStat, mnBoardNum, SigSource%, GateInterval%, _
         CBCount%, Freq&, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      LongCount& = IntValToULong(CBCount%)
      lblScanStatus.Caption = msConfig & "  Count: " & _
         LongCount& & "  Frequency: " & Freq&
      DoEvents
   Loop While mnLoop

End Sub

Private Sub tmrCheckStatus_Timer()
    
   Static Do2ndHalf%, DoFirstHalf%, Element&
   ULStat = GetStatus520(mnBoardNum, Status%, _
      CurCount&, CurIndex&, CTRFUNCTION)
   If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, _
      Status%, CurCount&, CurIndex&, CTRFUNCTION, _
      A6, A7, A8, A9, A10, A11, 0) Then
         ULStat = StopBackground520(mnBoardNum, CTRFUNCTION)
         If SaveFunc(Me, StopBackground, ULStat, _
            mnBoardNum, CTRFUNCTION, A3, A4, A5, _
            A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Status% = 0
         Do2ndHalf% = False
         DoFirstHalf% = False
         tmrCheckStatus.ENABLED = False
         cmdStop.ENABLED = (mnuLoop(2).Checked) _
            Or (chkManStart.value = 1)
         DoEvents
         mnPlot = True
         Exit Sub
   End If
   
   If mbHalfRoll Then
      If CurCount& > 0 Then
         mnCountRolls = mnCountRolls + 1
         mbHalfRoll = False
      End If
   Else
      mbHalfRoll = (CurCount < 0)
   End If
   
   If mnCancel Or gnCancel Then
      tmrCheckStatus.ENABLED = False
      tmrGoLoop.ENABLED = False
      cmdGo.FontBold = True
      cmdStop = True
      gnCancel = False
      mnCancel = False
      'If mnSetActive Then SetContActive False
      Element& = 0
      UpdateStatBar False
      If mnuStopBG.Checked Then
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
         If (Not gnScriptSave) Or (ULStat <> 0) Then
            If SaveFunc(Me, StopBackground, ULStat, _
               mnBoardNum, mlStatusType, A3, A4, A5, A6, _
               A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
         cmdStop.Visible = False
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      DoEvents
      mnPlot = True
      DisplayData
      Exit Sub
   End If
   
   UpdateStatBar False
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

   If Status% = 0 Then
      Element& = 0
      If mnuStopBG.Checked Then
         ULStat = StopBackground520(mnBoardNum, CTRFUNCTION)
         If SaveFunc(Me, StopBackground, ULStat, _
            mnBoardNum, CTRFUNCTION, A3, A4, A5, _
            A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      End If
      Do2ndHalf% = False
      DoFirstHalf% = False
      tmrCheckStatus.ENABLED = False
      UpdateStatBar False
      DoEvents
      mnPlot = True
      DisplayData
      cmdStop.ENABLED = (mnuLoop(2).Checked) Or (chkManStart.value = 1)
   Else
      cmdStop.ENABLED = True
   End If

End Sub

Private Sub tmrCIn_Timer()

   cmdGo.FontBold = Not cmdGo.FontBold
   ReadCounters

End Sub

Private Sub tmrDelay_Timer()

   tmrDelay.ENABLED = False

End Sub

Private Sub tmrGoLoop_Timer()

   If mnFuncType = CB_CSTART Then
      cmdStartTmr0.FontBold = Not cmdStartTmr0.FontBold
   Else
      cmdGo.FontBold = Not cmdGo.FontBold
   End If
   If mnTimerStop Then
      ULStat = GetStatus520(mnBoardNum, Status%, CurCount&, CurIndex&, mlStatusType)
      DevActive% = (Status% = RUNNING)
      'If mnuUseIdle.Checked Then DevActive% = Not (Status% = IDLE)
      If DevActive% Then
         ULStat = StopBackground520(mnBoardNum, mlStatusType)
      End If
   End If
   If gnCancel Or mnCancel Then
      If tmrGoLoop.ENABLED Then
         tmrGoLoop.ENABLED = False
         cmdGo.FontBold = True
         cmdStop = True
      End If
      gnCancel = False
      mnCancel = False
      lblStatus = "{Loop #" & mlTimerCount & "} " & lblStatus
      'mlTimerCount = 0
      Exit Sub
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
   If mnFuncType = CB_CSTART Then
      Me.cmdStartTmr0 = True
   Else
      cmdGo = True
   End If

End Sub

Private Sub UpdateMainStatus()

   board$ = mnuBoard(mnBoardIndex).Caption
   PrintMain "Current board: " & board$
   
End Sub

Private Sub UpdateStatBar(CycleStatus As Integer)
   
   Static sStatus As Integer
   msCtrList = ""
   For i% = 0 To mnNumEnabled - 1
      msCtrList = msCtrList & Str(maCtrList(i%))
   Next i%
   If msCtrList = "" Then msCtrList = "NO COUNTERS SELECTED"
   If CycleStatus Then
      sStatus = sStatus + 1
      If sStatus > 1 Then sStatus = 0
      StatType% = sStatus
   ElseIf tmrCheckStatus.ENABLED Then
      StatType% = 1
   Else
      StatType% = 0 'sStatus
   End If
   
   Select Case StatType%
      Case 0
         If mnFuncType = CB_CSCAN Then _
            Stat$ = msConfig & " " & msOpt
         lblScanStatus.Caption = Stat$
         lblStatus.Caption = msCtrList
      Case 1
         If mnuGetStatus.Checked Then
            ULStat = GetStatus520(mnBoardNum, Status%, _
               CurCount&, CurIndex&, CTRFUNCTION)
            If SaveFunc(Me, GetStatus, ULStat, mnBoardNum, _
               Status%, CurCount&, CurIndex&, CTRFUNCTION, _
               A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            If Status% Then
               BGOp$ = "Background task in progress..."
            Else
               BGOp$ = "No background task running..."
            End If
            If mnCountRolls > 1 Then RollString$ = " (X" & mnCountRolls & ")"
            Stat$ = BGOp$ & " Count = " & CurCount& & "  Index = " & CurIndex& & RollString$
            lblScanStatus.Caption = Stat$
            lblStatus.Caption = msCtrList
         End If
   End Select
   
End Sub

Private Sub mnuOpt_Click(Index As Integer)

   mnuOpt(Index).Checked = Not mnuOpt(Index).Checked
   mnuContPlot.Checked = (mnuContPlot.Checked And mnuOpt(1).Checked)
   mnuContPlot.ENABLED = mnuOpt(1).Checked
   Select Case Index
      'SINGLEIO, BLOCKIO are mutually exclusive
      'CTR32BIT, CTR64BIT are mutually exclusive
      Case 2
         mnuOpt(3).Checked = False
      Case 3
         mnuOpt(2).Checked = False
      Case 6
         mnuOpt(8).Checked = False
      Case 8
         mnuOpt(6).Checked = False
   End Select
   SetOptions
   lblTimerStat.Caption = "PulseOutStart() " & msOpt

End Sub

Private Sub mnuContPlot_Click()

   mnuContPlot.Checked = Not mnuContPlot.Checked

End Sub


Private Sub cmdStartTmr0_Click()

   mnCancel = False
   tmrGoLoop.ENABLED = mnuLoop(2).Checked
   cmdGo.FontBold = Not tmrGoLoop.ENABLED
   For CurCtr% = 0 To mnNumEnabled - 1
      CNum% = maCtrList(CurCtr%)
      SetupTimer (CNum%)
   Next

End Sub

Private Sub WriteTimerPulser()

   FindTmrBase% = True
   FindPulseBase% = True
   Frequency# = Val(txtTmr0Freq.Text)
   DutyCycle# = Val(txtTmr0Dty.Text)
   PulseCount& = Val(txtTmr0Count.Text)
   InitialDelay# = Val(txtDly0.Text)
   IdleState& = Val(txtIdle0.Text)
   Options& = mvOptions
   
   For CtrSelected% = 0 To mnNumEnabled - 1
      CounterNum% = maCtrList(CtrSelected%)
      CtrCheckBox% = CounterNum% '- mnCounterBase
      If IsNumeric(chkCounter(CtrCheckBox%).Tag) _
        Then CounterType% = Me.chkCounter(CtrCheckBox%).Tag
      If (CounterType% = 0) And (CtrCheckBox% > 0) Then
         If IsNumeric(chkCounter(0).Tag) Then CounterType% = chkCounter(0).Tag
      End If
      If FindTmrBase% And (CounterType% = CIOTechTmr) Then
         TmrBase% = CounterNum%
         FindTmrBase% = False
      End If
      If FindPulseBase% And (CounterType% = CIOTechPulse) Then
         PulseBase% = CounterNum%
         FindPulseBase% = False
      End If
      Pulser% = CounterNum% - PulseBase%
      Tmr% = CounterNum% - TmrBase%
      Select Case CounterType%
         Case CIOTechTmr
         Case CIOTechPulse
            Select Case mnLibType
               Case UNILIB
                  ULStat = PlsOutStart(mnBoardNum, Pulser%, Frequency#, DutyCycle#, _
                  PulseCount&, InitialDelay#, IdleState&, Options&)
                  lblTmr0Freq.Caption = Format(Frequency#, "0.0####") & "Hz"
                  lblDuty0.Caption = Format(DutyCycle# * 100, "0.0######") & "%"
                  lblDelay0.Caption = Format(InitialDelay#, "0.0##")
                  lblCount0.Caption = PulseCount&
                  x% = SaveFunc(Me, PulseOutStart, ULStat, mnBoardNum, Tmr%, Frequency#, _
                  DutyCycle#, PulseCount&, InitialDelay#, IdleState&, Options&, A9, A10, A11, 0)
                  Me.lblTimerStat.Caption = "PulseOutStart() " & msOpt
               Case MSGLIB
                  ChanString$ = Format(Pulser%, "0")
                  If Not mnFreq0Set Then
                     FreqReturned$ = SetMsgFreq(Frequency#, ChanString$)
                     If Not (FreqReturned$ = "") Then
                        FreqSet% = True
                        lblTmr0Freq.Caption = FreqReturned$
                     End If
                  End If
                  If Not mnDelay0Set Then
                     DelayReturned$ = SetMsgDelay(InitialDelay#, ChanString$)
                     If Not (DelayReturned$ = "") Then
                        DelaySet% = True
                        lblDelay0.Caption = DelayReturned$
                     End If
                  End If
                  If Not mnDuty0Set Then
                     DutyReturned$ = SetMsgDuty(DutyCycle#, ChanString$)
                     If Not (DutyReturned$ = "") Then
                        DutySet% = True
                        lblDuty0.Caption = DutyReturned$
                     End If
                  End If
                  If Not mnIdleState0Set Then
                     IdleReturned$ = SetMsgIdleState(IdleState&, ChanString$)
                     If Not (IdleReturned$ = "") Then IdleStateSet% = True
                  End If
                  If Not mnTmrCount0Set Then
                     CountReturned$ = SetMsgCountVal(PulseCount&, ChanString$)
                     If Not (CountReturned$ = "") Then
                        TmrCountSet% = True
                        lblCount0.Caption = CountReturned$
                     End If
                  End If
                  If Not mnTmr0Running Then
                     TmrMsg$ = "TMR{" & ChanString$ & "}:START"
                     MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                     If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                     TmrRunning% = True
                  End If
               Case Else
            End Select
         Case Else
      End Select
   Next
   
   mnFreq0Set = FreqSet%
   mnDelay0Set = DelaySet%
   mnDuty0Set = DutySet%
   mnIdleState0Set = IdleStateSet%
   mnTmrCount0Set = TmrCountSet%
   mnTmr0Running = TmrRunning%
      
End Sub

Private Sub SetupTimer(ByVal Tmr As Integer)
   
   Frequency# = Val(txtTmr0Freq.Text)
   DutyCycle# = Val(txtTmr0Dty.Text)
   PulseCount& = Val(txtTmr0Count.Text)
   InitialDelay# = Val(txtDly0.Text)
   IdleState& = Val(txtIdle0.Text)
   Options& = mvOptions

   StartText$ = "Start " & Format(Tmr, "0")
   TmrOffText$ = "Timer " & Format(Tmr, "0") & " Off"
   StopText$ = "Stop " & Format(Tmr, "0")
   If PulseCount& > 0 Then TmrRunning% = False
   
   
   If mnNumIOTechTmrs > 0 Then
      DoStart% = (cmdStartTmr0.Caption = StartText$)
   Else
      DoStart% = True
   End If
   If DoStart% Then
      If mnNumIOTechPulsers > 0 Then
         'added Options& at UL 5.88 (only 5.87 didn't have it)
         Select Case mnLibType
            Case UNILIB
               ULStat = PlsOutStart(mnBoardNum, Tmr%, Frequency#, DutyCycle#, _
               PulseCount&, InitialDelay#, IdleState&, Options&)
               lblTmr0Freq.Caption = Format(Frequency#, "0.0####") & "Hz"
               lblDuty0.Caption = Format(DutyCycle# * 100, "0.0######") & "%"
               lblDelay0.Caption = Format(InitialDelay#, "0.0##")
               lblCount0.Caption = PulseCount&
               x% = SaveFunc(Me, PulseOutStart, ULStat, mnBoardNum, Tmr%, Frequency#, _
               DutyCycle#, PulseCount&, InitialDelay#, IdleState&, Options&, A9, A10, A11, 0)
               lblTimerStat.Caption = "PulseOutStart() " & msOpt
            Case MSGLIB
               If Not FreqSet% Then
                  PeriodVal# = (1 / Frequency#) * 1000
                  TmrMsg$ = "TMR{" & ChanString$ & "}:PERIOD=" & Format(PeriodVal#, "0.0#########")
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                  FreqSet% = True
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:PERIOD"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then PerSet# = Val(ResultArr(1))
                  If PerSet# > 0 Then FreqRead# = 1 / (PerSet# / 1000)
                  lblTmr0Freq.Caption = Format(FreqRead#, "0.0#") & "Hz"
               End If
               If Not DelaySet% Then
                  DelayVal# = InitialDelay# * 1000
                  TmrMsg$ = "TMR{" & ChanString$ & "}:DELAY=" & Format(DelayVal#, "0.0########")
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                  DelaySet% = True
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:DELAY"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
                  DelayRead# = DSet# / 1000
                  lblDelay0.Caption = Format(DelayRead#, "0.0#")
               End If
               If Not DutySet% Then
                  DutyVal# = DutyCycle# * 100
                  TmrMsg$ = "TMR{" & ChanString$ & "}:DUTYCYCLE=" & Format(DutyVal#, "0.0###")
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                  DutySet% = True
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:DUTYCYCLE"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
                  DutyRead# = DSet#
                  lblDuty0.Caption = Format(DutyRead#, "0.0#") & "%"
               End If
               If Not IdleStateSet% Then
                  Select Case IdleState&
                     Case 0
                        StateString$ = "LOW"
                     Case 1
                        StateString$ = "HIGH"
                  End Select
                  TmrMsg$ = "TMR{" & ChanString$ & "}:IDLESTATE=" & StateString$
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                  IdleStateSet% = True
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:IDLESTATE"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then IdleRead$ = ResultArr(1)
                  Select Case IdleRead$
                     Case "LOW"
                        PropInt& = 0
                     Case "HIGH"
                        PropInt& = 1
                  End Select
                  'lblDuty.Caption = Format(PropInt&, "0")
               End If
               If Not TmrCountSet% Then
                  TmrMsg$ = "TMR{" & ChanString$ & "}:PULSECOUNT=" & Format(PulseCount&, "0")
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                  TmrCountSet% = True
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:PULSECOUNT"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then CountRead$ = ResultArr(1)
                  lblCount0.Caption = CountRead$
               End If
               If Not TmrRunning% Then
                  TmrMsg$ = "TMR{" & ChanString$ & "}:START"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
                  TmrRunning% = True
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:PERIOD"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then PerSet# = Val(ResultArr(1))
                  If PerSet# > 0 Then FreqRead# = 1 / (PerSet# / 1000)
                  lblTmr0Freq.Caption = Format(FreqRead#, "0.0#") & "Hz"
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:DELAY"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
                  DelayRead# = DSet# / 1000
                  lblDelay0.Caption = Format(DelayRead#, "0.0#")
                  TmrMsg$ = "?TMR{" & ChanString$ & "}:DUTYCYCLE"
                  MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
                  ResultArr = Split(MsgResult$, "=")
                  If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
                  DutyRead# = DSet#
                  lblDuty0.Caption = Format(DutyRead#, "0.0#") & "%"
               End If
         End Select
      Else
         Select Case mnLibType
            Case UNILIB
               ULStat = TmrOutStart(mnBoardNum, Tmr%, Frequency#)
               lblTmr0Freq.Caption = Format(Frequency#, "0.0####") & "Hz"
               x% = SaveFunc(Me, TimerOutStart, ULStat, mnBoardNum, Tmr%, _
               Frequency#, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            Case MSGLIB
               'PeriodVal# = (1 / Frequency#) * 1000
               'TmrMsg$ = "TMR{" & ChanString$ & "}:PERIOD=" & Format(PeriodVal#, "0.0#")
               'MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
               'If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
               'DelayVal# = InitialDelay# * 1000
               'TmrMsg$ = "TMR{" & ChanString$ & "}:DELAY=" & Format(DelayVal#, "0.0#")
               'MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
               'If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
               'TmrMsg$ = "TMR{" & ChanString$ & "}:START"
               'MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
               'If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
         End Select
      End If
      TmrText$ = "Tmr" & Format(Tmr, "0") & " Freq"
      If Not ULStat = 0 Then
         lblTmr0Freq.Caption = TmrText$
         Exit Sub
      End If
      If Not (mnNumIOTechPulsers > 0) Then cmdStartTmr0.Caption = StopText$
   Else
      'pulsers use separate button
      If mnNumIOTechPulsers > 0 Then
         Select Case mnLibType
            Case UNILIB
               ULStat = PlsOutStop(mnBoardNum, Tmr%)
               If SaveFunc(Me, PulseOutStop, ULStat, mnBoardNum, _
               Tmr%, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            Case MSGLIB
               TmrMsg$ = "TMR{" & ChanString$ & "}:STOP"
               MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
               If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
               TmrRunning% = False
         End Select
      Else
         Select Case mnLibType
            Case UNILIB
               ULStat = TmrOutStop(mnBoardNum, Tmr%)
               If SaveFunc(Me, TimerOutStop, ULStat, mnBoardNum, _
               Tmr%, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            Case MSGLIB
               'TmrMsg$ = "TMR{" & ChanString$ & "}:STOP"
               'MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
               'If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Sub
         End Select
      End If
      cmdStartTmr0.Caption = StartText$
      lblTmr0Freq.Caption = TmrOffText$
   End If

   'Select Case Tmr
   '   Case 0
         mnFreq0Set = FreqSet%
         mnDelay0Set = DelaySet%
         mnDuty0Set = DutySet%
         mnIdleState0Set = IdleStateSet%
         mnTmrCount0Set = TmrCountSet%
         mnTmr0Running = TmrRunning%
   '   Case 1
   '      mnFreq1Set = FreqSet%
   '      mnDelay1Set = DelaySet%
   '      mnDuty1Set = DutySet%
   '      mnIdleState1Set = IdleStateSet%
   '      mnTmrCount1Set = TmrCountSet%
   '      mnTmr1Running = TmrRunning%
   'End Select

End Sub

Private Function SetMsgFreq(ByVal Freq As Double, _
ByVal ChanString As String) As String

   PeriodVal# = (1 / Freq) * 1000
   TmrMsg$ = "TMR{" & ChanString & "}:PERIOD=" & Format(PeriodVal#, "0.0#########")
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Function
   TmrMsg$ = "?TMR{" & ChanString & "}:PERIOD"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then PerSet# = Val(ResultArr(1))
   If PerSet# > 0 Then FreqRead# = 1 / (PerSet# / 1000)
   SetMsgFreq = Format(FreqRead#, "0.0#") & "Hz"

End Function

Private Function SetMsgDelay(ByVal InitialDelay As Double, _
ByVal ChanString As String) As String

   DelayVal# = InitialDelay * 1000
   TmrMsg$ = "TMR{" & ChanString & "}:DELAY=" & Format(DelayVal#, "0.0########")
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Function
   TmrMsg$ = "?TMR{" & ChanString & "}:DELAY"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
   DelayRead# = DSet# / 1000
   SetMsgDelay = Format(DelayRead#, "0.0#")

End Function

Private Function SetMsgDuty(ByVal DutyCycle As Double, _
ByVal ChanString As String) As String

   DutyVal# = DutyCycle * 100
   TmrMsg$ = "TMR{" & ChanString & "}:DUTYCYCLE=" & Format(DutyVal#, "0.0###")
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Function
   TmrMsg$ = "?TMR{" & ChanString & "}:DUTYCYCLE"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
   DutyRead# = DSet#
   SetMsgDuty = Format(DutyRead#, "0.0#") & "%"
   
End Function

Private Function SetMsgIdleState(ByVal IdleState As Long, _
ByVal ChanString As String) As String

   Select Case IdleState
      Case 0
         StateString$ = "LOW"
      Case 1
         StateString$ = "HIGH"
   End Select
   TmrMsg$ = "TMR{" & ChanString & "}:IDLESTATE=" & StateString$
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Function
   TmrMsg$ = "?TMR{" & ChanString & "}:IDLESTATE"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then IdleRead$ = ResultArr(1)
   Select Case IdleRead$
      Case "LOW"
         PropInt$ = "0"
      Case "HIGH"
         PropInt$ = "1"
   End Select
   SetMsgIdleState = PropInt$

End Function

Private Function SetMsgCountVal(ByVal PulseCount As Long, _
ByVal ChanString As String) As String

   TmrMsg$ = "TMR{" & ChanString & "}:PULSECOUNT=" & Format(PulseCount, "0")
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   If SaveMsg(Me, "SendMessage(" & TmrMsg$ & ")", MsgResult$) Then Exit Function
   TmrMsg$ = "?TMR{" & ChanString & "}:PULSECOUNT"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then CountRead$ = ResultArr(1)
   SetMsgCountVal = CountRead$

   TmrMsg$ = "?TMR{" & ChanString$ & "}:PERIOD"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then PerSet# = Val(ResultArr(1))
   If PerSet# > 0 Then FreqRead# = 1 / (PerSet# / 1000)
   lblTmr0Freq.Caption = Format(FreqRead#, "0.0#") & "Hz"
   TmrMsg$ = "?TMR{" & ChanString$ & "}:DELAY"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
   DelayRead# = DSet# / 1000
   lblDelay.Caption = Format(DelayRead#, "0.0#")
   TmrMsg$ = "?TMR{" & ChanString$ & "}:DUTYCYCLE"
   MsgResult$ = MsgLibrary.SendMessage(TmrMsg$)
   ResultArr = Split(MsgResult$, "=")
   If UBound(ResultArr) > 0 Then DSet# = Val(ResultArr(1))
   DutyRead# = DSet#
   lblDuty.Caption = Format(DutyRead#, "0.0#") & "%"

End Function

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

Private Sub mnuTrigType_Click(Index As Integer)
   
   For TrigMenu% = 0 To mnuTrigType.Count - 1
      'If mnuTrigType(TrigMenu%).Checked Then TempTrigIndex% = TrigMenu%
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

Private Sub ResetOpts()

   'reset options to default
   For i% = 0 To mnuOpt.Count - 1
      mnuOpt(i%).Checked = False
   Next i%
   mvOptions = 0
   SetOptions
   UpdateStatBar False

End Sub

Sub SetHelpText(SourceName As String)

   Select Case SourceName
      Case "chkBEdge"
         Me.fraChipInit.Caption = "Select negative edge for B-phase quadrature input."
      Case "chkBit32"
         Me.fraChipInit.Caption = "Select 32-bit ctr for cbCIn32"
      Case "chkCompare1"
         If mnCounterType = C9513 Then
            fraChipInit.Caption = "Comparitor 1"
         Else
            fraChipInit.Caption = "Totalize or ClearOnRead"
         End If
      Case "chkCompare2"
         If mnCounterType = C9513 Then
            fraChipInit.Caption = "Comparitor 2"
         Else
            fraChipInit.Caption = "Rollover or StopAtMax"
         End If
      Case "chkCompare3"
         fraChipInit.Caption = "Count down (check) or up"
      Case "chkCompare4"
         fraChipInit.Caption = "Gate on or off"
      Case "chkEdge"
         fraChipInit.Caption = "Trig on rising or falling (check) edge"
      Case "chkNoRecycle"
         Me.fraChipInit.Caption = "Continuously count"
      Case "chkTrigger"
         fraChipInit.Caption = "Trig before or after stable"
      Case "chkUpDown"
         fraChipInit.Caption = "Enable up/down counting"
      Case "chkZEdge"
         Me.fraChipInit.Caption = "Select negative edge for Index quadrature input."
      Case "cmbFOutDiv"
         If mnCounterType = C9513 Then
            fraChipInit.Caption = "FOut Divider"
         Else
            fraChipInit.Caption = "Debounce Time"
         End If
      Case "cmbFOutSource"
         If mnCounterType = C9513 Then
            If mnFuncType = CB_FREQIN Then
               fraChipInit.Caption = "Freq In Source"
            Else
               fraChipInit.Caption = "FOut Source"
            End If
         Else
            fraChipInit.Caption = "Control Output"
         End If
      Case "cmbCtrType"
         fraChipInit.Caption = "Set counter mode"
      Case "cmbTickSize"
         fraChipInit.Caption = "Tick size"
      Case "cmbTimeOfDay"
         If mnFuncType = CB_FREQIN Then
            fraChipInit.Caption = "Gate Interval (ms)"
         Else
            fraChipInit.Caption = "Time of Day"
         End If
      Case "txtMapChan"
         fraChipInit.Caption = "Mapped Channel"
      Case "lstCountCfg"
         fraChipInit.Caption = "Configure Counter"
      Case "lstModeChange"
         fraChipInit.Caption = "Configure Inputs"
      Case "lstGateOut"
         fraChipInit.Caption = "Configure Gate / Ouput"
      Case Else
         'lblChipFunctions.Caption = ""
         fraChipInit.Caption = ""
         lblQuery.Visible = True
   End Select
   
End Sub

Private Sub txtDly0_Change()

   mnDelay0Set = False
   
End Sub

Private Sub txtIdle0_Change()

   mnIdleState0Set = False

End Sub

Private Sub txtMapChan_DragDrop(Source As Control, x As Single, y As Single)

   SetHelpText ""
   
End Sub

Private Sub txtMapChan_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  SetHelpText "txtMapChan"
  
End Sub

Private Sub SetOptions()
   
   msOpt = ""
   mvOptions = 0
   For i% = 0 To mnuOpt.Count - 1
      If mnuOpt(i%).Checked Then
         Options = Choose(i% + 1, BACKGROUND, CONTINUOUS, SINGLEIO, _
         BLOCK_IO, EXTTRIGGER, EXTCLOCK, CTR32BIT, CTR48BIT, CTR64BIT, _
         NOCLEAR, RETRIGMODE, HIGHRESRATE)
         opt$ = Choose(i% + 1, "BACKGROUND", "CONTINUOUS", "SINGLEIO", _
         "BLOCKIO", "EXTTRIGGER", "EXTCLOCK", "CTR32BIT", "CTR48BIT", "CTR64BIT", _
         "NOCLEAR", "RETRIGMODE", "HIGHRESRATE")
         If Not IsNull(Options) Then mvOptions = mvOptions + Options
         If Not IsNull(opt$) Then msOpt = msOpt & opt$ & " "
      End If
   Next
   msOpt = Trim(msOpt)
   If Len(msOpt) = 0 Then msOpt = "Options = Default  "
   UpdateStatBar 0
   
End Sub

Public Sub ShowBufferInfo(ByVal Source As Integer)

   DoEvents
   
   BufSize& = mlCount
   BufferStat% = CheckBuffer(mvHandle, BufSize&)
   Select Case BufferStat%
      Case INVALIDBUFFER
         MsgBox "Windows buffer no longer valid", , "Buffer Status"
      Case BUFFERTOOSMALL
         MsgBox "Windows buffer too small (size = " & BufSize& & _
            ", data requested = " & mlCount * 2 & ")", , "Buffer Status"
      Case NOBUFFER
         MsgBox "No Windows buffer has been allocated", , "Buffer Status"
      Case Else
         PrintMain "Buffer handle = 0x" & Hex$(mvHandle) & ",  buffer size = " & BufSize&
         PlotData% = True
         mnPlot = True
   End Select
   If PlotData% Then DisplayData

End Sub

Private Sub RefreshCtrList()

   mnNumEnabled = 0
   ReDim maCtrList(mnNumEnabled)
   For ChkNum% = 0 To 20
      TypeOfCtr% = chkCounter(ChkNum%).Tag
      If chkCounter(ChkNum%).value = 1 Then
         ReDim Preserve maCtrList(mnNumEnabled)
         If (TypeOfCtr% = 7) Or (TypeOfCtr% = 9) Then
            CurCtrNumber% = TmrPulserNum%
         Else
            CurCtrNumber% = ChkNum% '+ mnCounterBase
         End If
         maCtrList(mnNumEnabled) = CurCtrNumber%
         mnNumEnabled = mnNumEnabled + 1
      End If
      If (TypeOfCtr% = 7) Or (TypeOfCtr% = 9) Then
         TmrPulserNum% = TmrPulserNum% + 1
      End If
   Next ChkNum%

End Sub

Public Function GetDataHandle(AcqOrGen As Integer, DataType As VbVarType, NumSamples As Long) As Long

   Select Case AcqOrGen
      Case ACQUIREDDATA
         Handle& = mvHandle
         GetDataHandle = Handle&
         DataType = mnAcqDataType
         NumSamples = mlAqcPoints
      Case GENERATEDDATA
         GetDataHandle = mlGenHandle
         DataType = mnGenDataType
         NumSamples = mlGenPoints
   End Select
   
End Function

Function CalcMeasurement(MeasureType As Integer, Optional ReadVal As Variant) As String
         
   UseBuffer% = False
   If IsMissing(ReadVal) Then UseBuffer% = True
   Multiplier% = 1
   If MeasureType = 1 Then
      For CurItem% = 0 To 3
         If lstModeChange.Selected(CurItem%) Then _
            Multiplier% = 10 ^ CurItem%
      Next
   End If
   
   If UseBuffer% Then
      NumChans% = (mnLastChan - mnFirstChan)
      If mlCount > 11 Then
         FirstPoint& = mlCount - 10
         If mnResolution = 48 Then FirstPoint& = FirstPoint& * 2
         NumPoints& = 10
      Else
         FirstPoint& = 2
         NumPoints& = mlCount - 2
      End If
      Select Case mnResolution
         Case 48
            DimensionTwo& = (NumPoints& * 2) - 1
            ReDim LongArray(NumChans%, DimensionTwo&) As Long
            ReDim DWordArray(NumChans%, NumPoints) As Variant
            ULStat = WBufToArray32(mvHandle, LongArray(0, 0), _
               FirstPoint&, NumPoints& * 2)
            For Samp& = 0 To ((NumPoints& * 2) - 1) Step 2
               LWord = LongArray(Chan%, Samp&)
               If LWord < 0 Then LWord = LWord + 2 ^ 32
               DWordArray(Chan%, CombinedSample&) _
                  = LWord + (LongArray(Chan%, Samp& + 1) * 2 ^ 32)
               CumVal = CumVal + DWordArray(0, CombinedSample&)
               CombinedSample& = CombinedSample& + 1
            Next
         Case 32
            ReDim LongArray(NumChans%, NumPoints&) As Long
            ULStat = WBufToArray32(mvHandle, LongArray(0, 0), FirstPoint&, NumPoints&)
            For Samp& = 0 To NumPoints& - 1
               TempVal = LongArray(0, Samp&)
               If TempVal < 0 Then
                  CumVal = CumVal + (2 ^ 32 + LongArray(Ch%, l&))
               Else
                  CumVal = CumVal + LongArray(0, Samp&)
               End If
            Next
         Case Else
            ReDim IntArray(NumChans%, NumPoints&) As Integer
            ULStat = cbWinBufToArray(mvHandle, IntArray(0, 0), FirstPoint&, NumPoints&)
            For Samp& = 0 To NumPoints& - 1
               LongVal& = IntValToULong(IntArray(0, Samp&))
               CumVal = CumVal + LongVal&
            Next
      End Select
      AvgVal! = CumVal / NumPoints&
   Else
      AvgVal! = ReadVal
   End If
   
   TickSel% = Me.cmbTickSize.ListIndex
   Ticks& = Me.cmbTickSize.ListIndex + 1
   TickSize! = Choose(Ticks&, 0.00000002083, _
      0.0000002083, 0.0000020833, 0.0000208333, _
      0.00000002, 0.0000002, 0.000002, 0.00002)
   PerVal! = ((TickSize! * AvgVal!) * 1000) / Multiplier%
   If PerVal! > 999 Then
      PerVal! = PerVal! / 1000
      CalcResult$ = Format(PerVal!, "0.00###") & "s"
   Else
      CalcResult$ = Format(PerVal!, "0.00###") & "ms"
   End If
   CalcMeasurement = CalcResult$

End Function

Public Sub InitForm(FunctionInit As Integer)

   mnuBoard_Click (0)
   mnuFuncArray_Click (FunctionInit)
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
   InitCounters
   
End Sub


Private Sub UpdateFormProps(PropList As Variant, ListSize As Integer)

   'to do - set properties if any (other than number of counters) can be obtained
   'ResetOpts
   For i% = 0 To ListSize
      ListedProp$ = PropList(i%)
      ParseProp = Split(ListedProp$, "=")
      If UBound(ParseProp) > 0 Then
         Prop$ = ParseProp(0)
         PropVal$ = ParseProp(1)
         Select Case Prop$
            Case "CTR"
               If IsNumeric(PropVal$) Then
                  mnNumEventCtrs = Val(PropVal$)
               End If
            Case "DELAY"
               PropFloat# = Val(PropVal$) / 1000
               txtDly0.Text = Format(PropFloat#, "0.0#")
               mnDelay0Set = True
            Case "DUTYCYCLE"
               PropFloat# = Val(PropVal$) / 100
               txtTmr0Dty.Text = Format(PropFloat#, "0.0#")
               mnDuty0Set = True
            Case "IDLESTATE"
               Select Case PropVal$
                  Case "LOW"
                     PropInt& = 0
                  Case "HIGH"
                     PropInt& = 1
               End Select
                  txtIdle0.Text = Format(PropInt&, "0")
                  mnIdleState0Set = True
            Case "PERIOD"
               PropFloat# = Val(PropVal$)
               If PropFloat# > 0 Then
                  FreqVal# = 1 / (PropFloat# / 1000)
               End If
               txtTmr0Freq.Text = Format(FreqVal#, "0.0#")
               If PropFloat# > 0 Then mnFreq0Set = True
            Case "PULSECOUNT"
               PropInt& = Val(PropVal$)
               txtTmr0Count.Text = Format(PropInt&, "0")
               mnTmrCount0Set = True
            Case "TMR"
               TNum% = -1
               If IsNumeric(PropVal$) Then
                  mnNumIOTechPulsers = Val(PropVal$)
                  TNum% = TNum% + 1
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
         End Select
      End If
   Next
   UpdateStatBar False
   
End Sub

Public Function ArmCounter(CounterNum As Integer, Enable As Integer) As Integer

   CtrString$ = Format(CounterNum, "0")
   ArmString$ = "STOP"
   If Enable Then ArmString$ = "START"
   CtrStartMsg$ = "CTR{" & CtrString$ & "}:" & ArmString$
   MsgResult$ = MsgLibrary.SendMessage(CtrStartMsg$)
   x% = SaveMsg(Me, "SendMessage(" & CtrStartMsg$ & ")", MsgResult$)
   ArmCounter = x%

End Function

Public Function GetMsgDevice() As String

   BoardName$ = ""
   If Not MsgLibrary Is Nothing Then
      BoardName$ = msBoardName
   End If
   GetMsgDevice = BoardName$
   
End Function

Public Function GetCountsFromUnits(ByVal DataString As String, _
ByVal AsOffset As Integer) As Long

   GetCountsFromUnits = Val(DataString)
   
End Function

Private Sub txtTmr0Count_Change()

   mnTmrCount0Set = False

End Sub

Private Sub txtTmr0Dty_Change()

   mnDuty0Set = False
   
End Sub

Private Sub txtTmr0Freq_Change()

   mnFreq0Set = False
   
End Sub

Public Sub ConfigurationChange(ConfigType As String)

   ParseType$ = ConfigType
   If InStr(1, ConfigType, ":") > 0 Then ParseType$ = Mid(ConfigType, 2, 6)
   'Will need to change this for multiple TMRs
   Select Case ParseType$
      Case "TMR"
         QueryMsg = "?TMR"
         MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
         If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
         ParseMsg = Split(MsgResult$, "=")
         If UBound(ParseMsg) > 0 Then
            If ParseMsg(0) = "TMR" Then
               Timers$ = ParseMsg(1)
               NumTimers% = Val(Timers$)
            End If
         End If
         For Tmr% = 0 To NumTimers% - 1
            TmrString$ = Format(Tmr%, "0")
            QueryMsg = "?TMR{" & TmrString$ & "}:PERIOD"
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then
               If InStr(1, ParseMsg(0), ":PERIOD") > 0 Then
                  NewPeriod$ = ParseMsg(1)
                  PerSet# = Val(NewPeriod$)
                  If PerSet# > 0 Then FreqSet# = 1 / (PerSet# / 1000)
                  txtTmr0Freq.Text = Format(FreqSet#, "0.0#")
               End If
            End If
            QueryMsg = "?TMR{" & TmrString$ & "}:DUTYCYCLE"
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then
               If InStr(1, ParseMsg(0), ":DUTYCYCLE") > 0 Then
                  NewPeriod$ = ParseMsg(1)
                  DSet# = Val(NewPeriod$)
                  DutySet# = DSet# / 100
                  txtTmr0Dty.Text = Format(DutySet#, "0.0#")
               End If
            End If
            QueryMsg = "?TMR{" & TmrString$ & "}:DELAY"
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then
               If InStr(1, ParseMsg(0), ":DELAY") > 0 Then
                  NewPeriod$ = ParseMsg(1)
                  DSet# = Val(NewPeriod$)
                  DelaySet# = DSet# / 1000
                  txtDly0.Text = Format(DelaySet#, "0.0#")
               End If
            End If
            QueryMsg = "?TMR{" & TmrString$ & "}:PULSECOUNT"
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then
               If InStr(1, ParseMsg(0), ":PULSECOUNT") > 0 Then
                  NewCount$ = ParseMsg(1)
                  CountRead& = Val(NewCount$)
                  txtTmr0Count.Text = Format(CountRead&, "0")
               End If
            End If
            QueryMsg = "?TMR{" & TmrString$ & "}:IDLESTATE"
            MsgResult$ = MsgLibrary.SendMessage(QueryMsg)
            If SaveMsg(Me, "SendMessage(" & QueryMsg & ")", MsgResult$) Then Exit Sub
            ParseMsg = Split(MsgResult$, "=")
            If UBound(ParseMsg) > 0 Then
               If InStr(1, ParseMsg(0), ":IDLESTATE") > 0 Then
                  NewState$ = ParseMsg(1)
                  Select Case NewState$
                     Case "LOW"
                        PropInt& = 0
                     Case "HIGH"
                        PropInt& = 1
                  End Select
                  txtIdle0.Text = Format(PropInt&, "0")
               End If
            End If
         Next
   End Select
   
End Sub

Private Sub ConfigureValidOptions()

   Dim OptionFound As Boolean
   
   ConfigItem& = BICTRSCANOPTIONS
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

Public Function GetFormProperty(PropName As String) As Variant

   Select Case PropName
      Case "rate"
         GetFormProperty = mfRateReturned
      Case "resolution"
         GetFormProperty = mnResolution
      Case Else
         GetFormProperty = "Invalid"
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

