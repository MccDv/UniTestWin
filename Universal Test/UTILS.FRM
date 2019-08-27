VERSION 5.00
Begin VB.Form frmUtilities 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Miscellaneous Functions"
   ClientHeight    =   3120
   ClientLeft      =   1125
   ClientTop       =   1785
   ClientWidth     =   6960
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
   LinkTopic       =   "frmUtilities"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   6960
   Tag             =   "AIn"
   Begin VB.OptionButton optFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Characters"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hex"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   25
      Top             =   360
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Integers"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdConfigure 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5220
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ComboBox cmbSurrogateBoard 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5220
      TabIndex        =   22
      Text            =   "Surrogate"
      Top             =   1260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkSurrogate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Surrogate"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4860
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbDevNum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4140
      TabIndex        =   20
      Text            =   "DevNum"
      Top             =   60
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtOffset 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      TabIndex        =   17
      Text            =   "2048"
      Top             =   2340
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtAmpl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      TabIndex        =   16
      Text            =   "2048"
      Top             =   1980
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      Caption         =   "&Go"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer tmrReadPorts 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   2280
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Caption         =   "&Stop"
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtNumSecAddr 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1500
      TabIndex        =   13
      Text            =   "1"
      Top             =   1620
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.HScrollBar hsbNumSecAddr 
      Height          =   252
      LargeChange     =   4
      Left            =   180
      Max             =   64
      Min             =   1
      TabIndex        =   12
      Top             =   1620
      Value           =   1
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.HScrollBar hsbSecondAddr 
      Height          =   252
      LargeChange     =   256
      Left            =   180
      Max             =   4096
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox txtSecondAddr 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1500
      TabIndex        =   9
      Text            =   "0"
      Top             =   1260
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1500
      TabIndex        =   6
      Text            =   "0"
      Top             =   900
      Width           =   852
   End
   Begin VB.HScrollBar hsbData 
      Height          =   252
      LargeChange     =   512
      Left            =   180
      Max             =   4095
      TabIndex        =   5
      Top             =   900
      Width           =   1332
   End
   Begin VB.Label lblOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Offset"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2820
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblAmpl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Amplitude"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2820
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblXNumSecAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Number of secondary addresses"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2460
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblXSecondAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Offset of first address"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2460
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblError 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   3015
   End
   Begin VB.Label lblReturned 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label lblSent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sent to cbToEngUnits"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Label lblXReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Returned from cbToEngUnits"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3300
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblXError 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Error"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3300
      TabIndex        =   2
      Top             =   480
      Width           =   915
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
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnuFunc 
      Caption         =   "F&unction"
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbToEngUnits()"
         Index           =   0
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbFromEngUnits()"
         Index           =   1
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbFileRead()"
         Index           =   2
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbMemRead()"
         Index           =   3
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbMemWrite()"
         Index           =   4
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbInByte()"
         Index           =   5
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbOutByte()"
         Index           =   6
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbWinBufAlloc/Free()"
         Index           =   7
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbStopBackground()"
         Index           =   8
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbFlashLED()"
         Index           =   9
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBGetSerialNum"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBSetSerialNum"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBBlink"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBReset"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBWatchdog"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBMemRead"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBMemWrite"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbUSBSaveConfig"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbGetErrMsg()"
         Index           =   18
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbGetStatus()"
         Index           =   19
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDeviceLogin()"
         Index           =   20
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbDeviceLogout()"
         Index           =   21
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbTEDSRead()"
         Index           =   22
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbGetBoardName()"
         Index           =   23
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAConv 
         Caption         =   "cbAConvertData()"
      End
      Begin VB.Menu mnuBoardName 
         Caption         =   "cbGetBoardName()"
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuWord 
         Caption         =   "&Word IO"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComposite 
         Caption         =   "Composite.."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMask 
         Caption         =   "Apply Mask.."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTimer 
         Caption         =   "Timer Loop"
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResolution 
         Caption         =   "Resolution"
         Begin VB.Menu mnuUseList 
            Caption         =   "Use List"
         End
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
            Caption         =   "BIP20VOLTS"
            Index           =   0
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP10VOLTS"
            Index           =   1
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP5VOLTS"
            Checked         =   -1  'True
            Index           =   2
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP4VOLTS"
            Index           =   3
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP2PT5VOLTS"
            Index           =   4
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP2VOLTS"
            Index           =   5
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP1PT25VOLTS"
            Index           =   6
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP1VOLTS"
            Index           =   7
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT625VOLTS"
            Index           =   8
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT5VOLTS"
            Index           =   9
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT25VOLTS"
            Index           =   10
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT2VOLTS"
            Index           =   11
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT1VOLTS"
            Index           =   12
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT05VOLTS"
            Index           =   13
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT01VOLTS"
            Index           =   14
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT005VOLTS"
            Index           =   15
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP1PT67VOLTS"
            Index           =   16
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT312VOLTS"
            Index           =   17
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT156VOLTS"
            Index           =   18
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT078VOLTS"
            Index           =   19
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP60VOLTS"
            Index           =   20
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIP15VOLTS"
            Index           =   21
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT125VOLTS"
            Index           =   22
         End
         Begin VB.Menu mnuBipRange 
            Caption         =   "BIPPT025VOLTSPERVOLT"
            Index           =   23
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
   End
   Begin VB.Menu mnuPlotSel 
      Caption         =   "&Plot Type"
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
   End
   Begin VB.Menu mnuDataSel 
      Caption         =   "&Data"
      Begin VB.Menu mnuData 
         Caption         =   "Single"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuData 
         Caption         =   "Square"
         Index           =   1
      End
      Begin VB.Menu mnuData 
         Caption         =   "Sine"
         Index           =   2
      End
      Begin VB.Menu mnuData 
         Caption         =   "Ramp"
         Index           =   3
      End
      Begin VB.Menu mnuData 
         Caption         =   "Build"
         Index           =   4
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowData 
         Caption         =   "Show Data"
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
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mnFuncType's
Const TO_ENG = 0
Const FROM_ENG = 1
Const FILE_READ = 2
Const MEM_READ = 3
Const MEM_WRITE = 4
Const REG_IN = 5
Const REG_OUT = 6
Const WIN_BUF = 7
Const STOP_BG = 8
Const FLASH_LED = 9
Const USB_GETSERIAL = 10
Const USB_SETSERIAL = 11
Const USB_BLINK = 12
Const USB_RESET = 13
Const USB_WATCH = 14
Const USB_MEMREAD = 15
Const USB_MEMWRITE = 16
Const SAVE_CONFIG = 17
Const GET_ERROR = 18
Const GET_STATUS = 19
'Const LOAD_CONFIG = 20
Const DEV_LOGIN = 20
Const DEV_LOGOUT = 21
Const READ_TEDS = 22
Const READ_BOARDNAME = 23

#If MSGOPS Then
   Dim WithEvents MsgLibrary As MBDClass.MBDComClass
Attribute MsgLibrary.VB_VarHelpID = -1
#Else
   Dim MsgLibrary As Object
#End If

Dim mnThisInstance As Integer
Dim mnFormType As Integer, msTitle As String
Dim mnBoardNum As Integer, mnBoardIndex As Integer
Dim mnFuncType As Integer

Dim mnPlot As Integer, mnLoaded As Integer

Dim msConfig As String, msOpt As String

Dim mnRange As Integer, mnResolution As Integer
Dim mnWordMode As Integer, mDAData() As Integer
Dim mlCount As Long, mnDataType As Integer
Dim mlFirstPoint As Long

Dim mvHandle As Variant, mvOptions As Variant

Dim mnSurrogateBoard As Integer, mnSurrogateIndex As Integer
Dim mnMaskVal As Integer, mlCurAddress As Long

Dim mnMaskFirst As Integer, mnMaskSecond As Integer, mnConsecutive As Integer
Dim mnOpType As Integer
Dim msBoardName As String
Dim mlCustomRange As Long
Dim mvCustomRange As Variant
Dim mnFixedRange As Integer

Dim mnLibType As Integer, mnNumBoards As Integer
Dim msStatusType As String   'added for msgDaq
Dim mnUnloading As Integer, mnLoading As Integer
Dim mnRefreshProps As Integer, mnMessaging As Integer
Dim mnFormInitialized As Integer, msDisplayName As String

Private Sub chkSurrogate_Click()

   cmbSurrogateBoard.Visible = (chkSurrogate.value = 1)
   SetOpts  'and ConfigureControls

End Sub

Private Sub cmbDevNum_Change()

   If FuncType = STOP_BG Then
      mnOpType = cmbDevNum.ListIndex + 1
   Else
      InfoType% = BOARDINFO
      ConfigItem% = BIBASEADR
      DevNum% = cmbDevNum.ListIndex
      board% = mnBoardNum
      If chkSurrogate.value = 1 Then board% = mnSurrogateBoard
      ULStat = GetConfig520(InfoType%, board%, DevNum%, ConfigItem%, ValConfig&)
      ConfigVal& = ValConfig&
      'ULStat = cbGetConfig(InfoType%, Board%, DevNum%, ConfigItem%, ConfigVal&)
      If SaveFunc(Me, GetConfig, ULStat, InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      lblReturned.Caption = "&&H" & Hex$(ConfigVal&)
   End If

End Sub

Private Sub cmbDevNum_Click()

   If mnFuncType = STOP_BG Then
      mnOpType = cmbDevNum.ListIndex + 1
   Else
      If mnLibType = UNILIB Then
         InfoType% = BOARDINFO
         ConfigItem% = BIBASEADR
         DevNum% = cmbDevNum.ListIndex
         board% = mnBoardNum
         If chkSurrogate.value = 1 Then board% = mnSurrogateBoard
         ULStat = GetConfig520(InfoType%, board%, DevNum%, ConfigItem%, ValConfig&)
         ConfigVal& = ValConfig&
         'ULStat = cbGetConfig(InfoType%, Board%, DevNum%, ConfigItem%, ConfigVal&)
         If SaveFunc(Me, GetConfig, ULStat, InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         lblReturned.Caption = "&&H" & Hex$(ConfigVal&)
      End If
   End If

End Sub

Private Sub cmbSurrogateBoard_Change()

   mnSurrogateIndex = cmbSurrogateBoard.ListIndex
   mnSurrogateBoard = gnBoardEnum(mnSurrogateIndex)
   'If chkSurrogate.Value = 1 Then
   '   ConfigureControls
   '   Caption = msTitle & " surrogate " & mnuBoard(mnSurrogateIndex).Caption
   'End If
   SetOpts  'and ConfigureControls

End Sub

Private Sub cmbSurrogateBoard_Click()

   mnSurrogateIndex = cmbSurrogateBoard.ListIndex
   mnSurrogateBoard = gnBoardEnum(mnSurrogateIndex)
   'If chkSurrogate.Value = 1 Then
   '   ConfigureControls
   '   Caption = msTitle & " surrogate " & mnuBoard(mnSurrogateIndex).Caption
   'End If
   SetOpts  'and ConfigureControls

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
         CurOption$ = msOpt
         If msOpt = "Options = Default  " Then CurOption$ = ""
         Me.cmdConfigure.Caption = CurOption$
      Case "A" 'composite consecutive or same address
         mnConsecutive = value&
      Case "B" 'set board number
         SearchName$ = Mid$(cmdConfigure.Caption, 2)
         BoardParams = Split(SearchName$, ",")
         If UBound(BoardParams) = 1 Then DupeIndex% = Val(BoardParams(1))
         For MenuIndex% = 0 To gnNumBoards - 1
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
      Case "C" 'turn composite mode on or off
         'mnSurrogateBoard, mnSurrogateIndex
         'mlCurAddress
         mnuComposite.Checked = (Not value& = 0)
      Case "D" 'set DevNum (used also for PCI BADR on surrogate)
         cmbDevNum.ListIndex = value&
      Case "F" 'set function
         'If Value& > 6 Then
         '   Select Case Value&
         '      Case 7
         '         mnuStopBackground_Click
         '   End Select
         'Else
            mnuFuncArray_Click (value&)
         'End If
      Case "M" 'set mask value
         mnMaskVal = CInt(value&)
      Case "N" 'where to apply mask
         mnMaskFirst = Not ((value& And 1) = 0)
         mnMaskSecond = Not ((value& And 2) = 0)
      Case "R" 'register to write to
         txtSecondAddr.Text = "&H" & Hex$(CInt(value&))
      Case "S" 'enable surrogate board
         chkSurrogate.value = value&
      Case "U" 'set surrogate board number
         cmbSurrogateBoard.ListIndex = value&
      Case "V" 'set value to write
         txtData.Text = "&H" & Hex$(CInt(value&))
      Case "W" 'use cbInWord or cbOutWord (not Byte)
         If Not mnuWord.Checked Then mnuWord_Click
      Case "X" 'stop background
         'ULStat = cbStopBackground(mnBoardNum)
         'will need additional code for selection of subsystem
         ULStat = StopBackground520(mnBoardNum, 0)
      Case "Y" 'use cbInByte or cbOutByte (not Word)
         If mnuWord.Checked Then mnuWord_Click
   End Select

End Sub

Private Sub cmdGo_Click()

   Static stErrStart%

   txtData.ForeColor = &H80000008
   txtSecondAddr.ForeColor = &H80000008
   txtNumSecAddr.ForeColor = &H80000008
   If mnuTimer.Checked Then
      tmrReadPorts.ENABLED = True
      cmdStop.Visible = True
   End If
   Select Case mnFuncType
      Case FILE_READ
         Filename$ = txtData.Text
         PlotFile Filename$
      Case MEM_READ
         NumPoints& = Val(txtSecondAddr.Text)
         FirstPoint& = Val(txtData.Text)
         Chans% = 0
         PlotMemory mnBoardNum, mnBoardNum, NumPoints&, FirstPoint&, Chans%
      Case MEM_WRITE
         WriteMem
      Case REG_IN
         ReadRegisters
      Case REG_OUT
         WriteRegisters
      Case WIN_BUF
         CycleBuffer
      Case STOP_BG
         If mnLibType = MSGLIB Then
            Component$ = Choose(mnOpType, "AISCAN", "AOSCAN", "DISCAN", "DOSCAN", "CTRSCAN", "", "")
            StopMsg$ = Component$ & ":STOP"
            If Not MsgLibrary Is Nothing Then
               MsgResult$ = MsgLibrary.SendMessage(StopMsg$)
               If SaveMsg(Me, "SendMessage(" & StopMsg$ & ")", MsgResult$) Then Exit Sub
               lblStatus.Caption = StopMsg$ & " (" & msBoardName & ") = " & MsgResult$
            End If
         Else
            ULStat = StopBackground520(mnBoardNum, mnOpType)
            x% = SaveFunc(Me, StopBackground, ULStat, mnBoardNum, mnOpType, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            lblStatus.Caption = "cbStopBackground(" & mnBoardNum & ", " & mnOpType & ") = " & ULStat
         End If
      Case FLASH_LED
         Select Case mnLibType
            Case UNILIB
               ULStat = LEDFlash(mnBoardNum)
               If SaveFunc(Me, FlashLED, ULStat, mnBoardNum, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               lblStatus.Caption = "cbFlashLED(" & mnBoardNum & ") = " & ULStat
            Case MSGLIB
               NumFlash$ = Me.txtData.Text
               FlashMsg$ = "DEV:FLASHLED/" & NumFlash$
               MsgResult$ = MsgLibrary.SendMessage(FlashMsg$)
               If SaveMsg(Me, "SendMessage(" & FlashMsg$ & ")", MsgResult$) Then Exit Sub
               lblStatus.Caption = FlashMsg$ & " = " & MsgResult$
         End Select
      Case USB_GETSERIAL
      Case USB_SETSERIAL
      Case USB_BLINK
      Case USB_RESET
      Case USB_WATCH
      Case USB_MEMREAD
      Case USB_MEMWRITE
      Case SAVE_CONFIG
      Case GET_ERROR
         If stErrStart% > 900 Then stErrStart% = 0
         SetPlotType PRINT_LIST, Me
         frmPlot.txtShow.Text = ""
         LineFeed$ = Chr$(13) + Chr$(10)
         
         'Get each consecutive error message in list
         Screen.MousePointer = vbHourglass
         If mnLibType = MSGLIB Then
            For ErrCode% = -1 To 2
               ErrMessage$ = MsgLibrary.ConvertErrorCode(ErrCode%)
               Messages$ = Messages$ & "Code " & Format(ErrCode%, "0") & " (" & ErrMessage$ & ")" & LineFeed$
            Next
            For ErrCode% = 2000 To 2065
               ErrMessage$ = MsgLibrary.ConvertErrorCode(ErrCode%)
               Messages$ = Messages$ & "Code " & Format(ErrCode%, "0") & " (" & ErrMessage$ & ")" & LineFeed$
            Next
         Else
            For ErrCode% = stErrStart% To stErrStart% + 49
               ErrMessage$ = Space$(ERRSTRLEN)
               ULStat = cbGetErrMsg(ErrCode%, ErrMessage$)
               If SaveFunc(Me, GetErrMsg, ULStat, ErrCode%, ErrMessage$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               ErrMessage$ = RTrim$(ErrMessage$)   'Drop the space characters
               ErrMessage$ = Left$(ErrMessage$, Len(ErrMessage$) - 1)
               Messages$ = Messages$ & Str$(ErrCode%) & " (" & GetErrorConst(ErrCode%) & "): " & ErrMessage$ & LineFeed$
            Next ErrCode%
         End If
         TextList Messages$
         Screen.MousePointer = vbDefault
         stErrStart% = stErrStart% + 50
      Case GET_STATUS
         Selection% = cmbDevNum.ListIndex + 1
         'initialize so user can tell when library changes values
         '(otherwise, VB sets these to zero)
         BGStatus% = -7
         CurCount& = -7
         CurIndex& = -7
         If mnLibType = MSGLIB Then
            Component$ = Choose(Selection%, "AISCAN", "AOSCAN", "DISCAN", "DOSCAN", "CTRSCAN", "", "")
            BGStatus% = MsgCheckStatus(Me, msBoardName, MsgLibrary, Component$, CurCount&, CurIndex&)
         Else
            StatType& = Choose(Selection%, AIFUNCTION, AOFUNCTION, DIFUNCTION, DOFUNCTION, CTRFUNCTION, DAQIFUNCTION, DAQOFUNCTION)
            ULStat = GetStatus520(mnBoardNum, BGStatus%, CurCount&, CurIndex&, StatType&)
            x% = SaveFunc(Me, GetStatus, ULStat, mnBoardNum, BGStatus%, CurCount&, CurIndex&, StatType&, A6, A7, A8, A9, A10, A11, 0)
         End If
         txtData.ForeColor = &HFF0000  '&H80000008&
         txtSecondAddr.ForeColor = &HFF0000
         txtNumSecAddr.ForeColor = &HFF0000
         txtData.Text = Format$(BGStatus%, "0")
         txtSecondAddr.Text = Format$(CurCount&, "0")
         txtNumSecAddr.Text = Format$(CurIndex&, "0")
      Case DEV_LOGIN '= 20
         AccountName$ = txtData.Text
         Password$ = txtSecondAddr.Text
         ULStat = cbDeviceLogin(mnBoardNum, AccountName$, Password$)
         x% = SaveFunc(Me, DeviceLogin, ULStat, mnBoardNum, AccountName$, Password$, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      Case DEV_LOGOUT '= 21
         ULStat = cbDeviceLogout(mnBoardNum)
         x% = SaveFunc(Me, DeviceLogout, ULStat, mnBoardNum, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      Case READ_TEDS '22
         Chan& = Val(txtData.Text)
         CBCount& = Val(txtSecondAddr.Text)
         ReDim DataBuffer(CBCount&) As Byte
         Screen.MousePointer = vbHourglass
         ULStat = cbTEDSRead(mnBoardNum, Chan&, DataBuffer(0), CBCount&, Options&)
         If Not (ULStat = 0) Then Screen.MousePointer = vbDefault
         If SaveFunc(Me, TEDSRead, ULStat, mnBoardNum, Chan&, DataBuffer(0), CBCount&, Options&, A6, A7, A8, A9, A10, A11, 0) Then
            lblStatus.Caption = ""
            Exit Sub
         End If
         lblStatus.Caption = "Returned from cbTEDSRead(): " & Format(CBCount&, "0") & " items"
         SetPlotType PRINT_LIST, Me
         frmPlot.txtShow.Text = ""
         'For TEDsCode% = 0 To CBCount& - 1
         Do
            TMessage$ = ""
            TCount$ = Format(TEDsCode%, "0") & vbTab
            For Mult% = 0 To 7
               TMessage$ = TMessage$ & Format(DataBuffer(TEDsCode% + Mult%), "0") & ", "
               Done% = Not ((TEDsCode% + Mult%) < CBCount&)
               If Done% Then Exit For
            Next
            TMessage$ = Left(TMessage$, Len(TMessage$) - 2)
            Messages$ = Messages$ & TCount$ & TMessage$ & vbCrLf
            TEDsCode% = TEDsCode% + 8
            Done% = Not (TEDsCode% < CBCount&)
         Loop While Not Done%
         'Next TEDsCode%
         TextList Messages$
         Screen.MousePointer = vbDefault
      Case READ_BOARDNAME
         ReadBoardName
   End Select
   cmdStop.ENABLED = tmrReadPorts.ENABLED

End Sub

Private Sub cmdGo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
      'if right click, change timer state
      mnLoopType = TIMERLOOP
      If Me.mnuTimer.Checked Then
         'if already checked, just change loop rate
         LoopInterval% = tmrReadPorts.Interval
         If Shift > 0 Then
            If LoopInterval% < 16384 Then LoopInterval% = LoopInterval% * 2
         Else
            If LoopInterval% > 1 Then LoopInterval% = LoopInterval% / 2
         End If
         Me.tmrReadPorts.Interval = LoopInterval%
      End If
      mnuTimer.Checked = True
      cmdGo = True
      cmdStop.Visible = True
   End If

End Sub

Private Sub cmdStop_Click()

   tmrReadPorts.ENABLED = False
   cmdStop.ENABLED = False

End Sub

Private Sub ConfigureControls()
   
   'display appropriate controls and
   'hide others
   If mnLibType = MSGLIB Then
      For MenuIndex% = 0 To mnuFuncArray.Count - 1
         mnuFuncArray(MenuIndex%).ENABLED = False
      Next
      mnuFuncArray(FLASH_LED).ENABLED = True
      mnuFuncArray(STOP_BG).ENABLED = True
      mnuFuncArray(GET_ERROR).ENABLED = True
      mnuFuncArray(GET_STATUS).ENABLED = True
      mnuAConv.ENABLED = False
      'txtData.Visible = ShowFlashLEDCtls%
   Else
      For MenuIndex% = 0 To mnuFuncArray.Count - 1
         mnuFuncArray(MenuIndex%).ENABLED = True
      Next
      mnuAConv.ENABLED = True
   End If
   txtData.Visible = False

   ShowConvCtls% = mnuFuncArray(TO_ENG).Checked Or mnuFuncArray(FROM_ENG).Checked
   ShowFileCtls% = mnuFuncArray(FILE_READ).Checked
   ShowRegInCtls% = mnuFuncArray(REG_IN).Checked
   ShowRegOutCtls% = mnuFuncArray(REG_OUT).Checked
   ShowRegCtls% = ShowRegOutCtls% Or ShowRegInCtls%
   ShowMemCtls% = mnuFuncArray(MEM_READ).Checked Or mnuFuncArray(MEM_WRITE).Checked
   ShowMemWriteCtls% = mnuFuncArray(MEM_WRITE).Checked
   ShowMemReadCtls% = mnuFuncArray(MEM_READ).Checked
   ShowWinBufCtls% = mnuFuncArray(WIN_BUF).Checked
   ShowStopBGCtls% = mnuFuncArray(STOP_BG).Checked
   ShowFlashLEDCtls% = mnuFuncArray(FLASH_LED).Checked Or mnuFuncArray(READ_BOARDNAME).Checked
   'ShowLoadCfgCtls% = mnuFuncArray(LOAD_CONFIG).Checked
   ShowGetErrorCtls% = mnuFuncArray(GET_ERROR).Checked
   ShowGetStatusCtls% = mnuFuncArray(GET_STATUS).Checked
   ShowUSBCtls% = (mnBoardNum < 0)
   ShowLogIn% = mnuFuncArray(DEV_LOGIN).Checked
   ShowLogOut% = mnuFuncArray(DEV_LOGOUT).Checked
   ShowTeds% = mnuFuncArray(READ_TEDS).Checked

   ShowGo% = ShowFileCtls% Or ShowMemCtls% Or ShowRegCtls% Or ShowLoadCfgCtls% Or ShowWinBufCtls%
   ShowGo% = ShowGo% Or ShowLogIn% Or ShowLogOut% Or ShowTeds%
   cmdGo.Visible = ShowGo% Or ShowGetStatusCtls% Or ShowStopBGCtls% Or ShowUSBCtls% Or ShowFlashLEDCtls% Or ShowGetErrorCtls%
   'cmdLoadData.Visible = ShowMemWriteCtls% And Not mnuData(0).Checked
   cmdStop.Visible = mnuTimer.Checked
   lblError.Visible = ShowConvCtls%
   lblXError.Visible = ShowConvCtls%
   lblXReturn.Visible = ShowConvCtls% Or ShowRegCtls%
   lblSent.Visible = ShowMemCtls% Or ShowRegCtls% Or ShowWinBufCtls%
   DataShow% = ShowConvCtls% Or ShowRegCtls% Or ShowMemCtls% Or ShowWinBufCtls% Or ShowGetErrorCtls%
   DataShow% = DataShow% Or ShowGetErrorCtls% Or ShowTeds%
   hsbData.Visible = DataShow%
   txtData.Visible = DataShow% Or ShowFlashLEDCtls% 'Not (ShowStopBGCtls% Or ShowLoadCfgCtls% Or ShowFlashLEDCtls%)
   lblReturned.Visible = ShowConvCtls% Or ShowRegCtls%
   hsbSecondAddr.Visible = ShowRegCtls% Or ShowMemCtls% Or ShowTeds%
   txtSecondAddr.Visible = ShowRegCtls% Or ShowMemCtls% Or ShowTeds%
   lblXSecondAddr.Visible = ShowRegCtls% Or ShowMemCtls%
   txtOffset.Visible = ShowMemWriteCtls%
   lblOffset.Visible = ShowMemWriteCtls%
   lblAmpl.Visible = ShowMemWriteCtls%
   txtAmpl.Visible = ShowMemWriteCtls%
   mnuDataSel.Visible = ShowMemWriteCtls%

   lblXNumSecAddr.Visible = ShowMemWriteCtls%
   txtOffset.Visible = ShowMemWriteCtls%
   txtAmpl.Visible = ShowMemWriteCtls%
   
   hsbData.Top = 184
   cmbDevNum.Visible = False
   cmbDevNum.Left = 4140
   cmbDevNum.Width = 855
   chkSurrogate.Visible = False
   cmbSurrogateBoard.Visible = False
   txtNumSecAddr.Visible = False
   txtNumSecAddr.Text = "1"
   txtNumSecAddr.Left = 1500
   txtNumSecAddr.Width = 852
   txtNumSecAddr.Top = 1620
   lblSent.Visible = False
   lblSent.Top = 960
   lblSent.Left = 2460
   lblSent.Caption = "Sent to cbToEngUnits"
   txtData.Top = 900
   txtData.Left = 1500
   txtData.Width = 852
   txtData.Text = "0"
   lblXSecondAddr.Visible = False
   lblXSecondAddr.Left = 2460
   lblXSecondAddr.Top = 1320
   lblXSecondAddr.Caption = "Offset of first address"
   txtSecondAddr.Visible = False
   txtSecondAddr.Text = "0"
   txtSecondAddr.Top = 1260
   txtSecondAddr.Left = 1500
   txtSecondAddr.Width = 852
   lblXNumSecAddr.Visible = False
   lblXNumSecAddr.Left = 2460
   lblXNumSecAddr.Top = 1680
   lblXNumSecAddr.Caption = "Number of secondary addresses"

   Select Case True
      Case ShowFileCtls%
         lpFileName$ = "UniTest.ini"
         lpApplicationName$ = "DataFile"
         lpKeyName$ = "FileName"
         lpDefault$ = "C:\Demo.dat"
         nSize% = 80
         DataFile$ = Space$(nSize%)
         StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, DataFile$, nSize%, lpFileName$)
         DataFile$ = Left$(DataFile$, StringSize%)
         txtData.Width = 3500
         txtData.Width = 3500
         txtData.Top = lblReturned.Top
         txtData.Left = lblReturned.Left
         txtData.Text = DataFile$
         txtData.Visible = True
         hsbNumSecAddr.Visible = False
         txtNumSecAddr.Visible = False
         lblXNumSecAddr.Visible = False
      Case ShowRegInCtls%, ShowRegOutCtls%
         If mnWordMode Then
            mnuFuncArray(5).Caption = "cbInWord()"
            mnuFuncArray(6).Caption = "cbOutWord()"
         Else
            mnuFuncArray(5).Caption = "cbInByte()"
            mnuFuncArray(6).Caption = "cbOutByte()"
         End If
         ThisBoard$ = GetNameOfBoard(mnBoardNum)
         ShowPCICtls% = (UCase$(Left$(ThisBoard$, 3)) = "PCI") Or (UCase$(Left$(ThisBoard$, 4)) = "CPCI") Or (chkSurrogate.value = 1)
         ShowDemoCtls% = (UCase$(Left$(ThisBoard$, 4)) = "DEMO")
         chkSurrogate.Visible = ShowDemoCtls%
         cmbDevNum.Visible = ShowPCICtls% Or (chkSurrogate.value = 1)
         If ShowPCICtls% Then
            cmbDevNum.Left = 4140
            cmbDevNum.Width = 855
            cmbDevNum.Clear
            cmbDevNum.AddItem "0"
            cmbDevNum.AddItem "1"
            cmbDevNum.AddItem "2"
            cmbDevNum.AddItem "3"
            cmbDevNum.AddItem "4"
            cmbDevNum.ListIndex = 0
            mnOpType = 1
         End If
         If ShowDemoCtls% Then
            If chkSurrogate.value = 1 Then cmbSurrogateBoard.Visible = True
            chkSurrogate.Top = cmdGo.Top + 450
            cmbSurrogateBoard.Top = chkSurrogate.Top + 325
            If cmbSurrogateBoard.ListIndex < 0 Then
               cmbSurrogateBoard.Clear
               For BoardList% = 0 To gnNumBoards - 1
                  cmbSurrogateBoard.AddItem Format$(gnBoardEnum(BoardList%), "0")
               Next BoardList%
               cmbSurrogateBoard.ListIndex = 0
            End If
         End If
         board% = mnBoardNum
         InfoType% = BOARDINFO
         ConfigItem% = BIBASEADR
         If Not cmbDevNum.ListIndex < 0 Then DevNum% = cmbDevNum.ListIndex
         ULStat = GetConfig520(InfoType%, board%, DevNum%, ConfigItem%, ValConfig&)
         ConfigVal& = ValConfig&
         'ULStat = cbGetConfig(InfoType%, Board%, DevNum%, ConfigItem%, ConfigVal&)
         If SaveFunc(Me, GetConfig, ULStat, InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         mlCurAddress = ConfigVal&
         If chkSurrogate.value = 1 Then
            board% = mnSurrogateBoard
            ULStat = GetConfig520(InfoType%, board%, DevNum%, ConfigItem%, ValConfig&)
            ConfigVal& = ValConfig&
            'ULStat = cbGetConfig(InfoType%, Board%, DevNum%, ConfigItem%, ConfigVal&)
            If SaveFunc(Me, GetConfig, ULStat, InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
         hsbSecondAddr.Top = lblError.Top + 364
         txtSecondAddr.Top = lblError.Top + 364
         txtSecondAddr.Visible = True
         lblXSecondAddr.Top = lblError.Top + 432
         lblXSecondAddr.Visible = True
         lblXReturn.Caption = "Base address"
         lblReturned.Caption = "&&H" & Hex$(ConfigVal&)
         SentString$ = "Number of addresses to read"
         'SecAddrString$ = "Offset of secondary address"
         SecAddrString$ = "Offset of first address"
         If ShowRegOutCtls% Then
            SecAddrString$ = "Offset of address"
            SentString$ = "Data to write"
            txtSecondAddr.Visible = True
            txtSecondAddr.Text = "0"
            txtNumSecAddr.Visible = False
            hsbNumSecAddr.Visible = False
            lblXNumSecAddr.Visible = False
         'Else
         '   hsbNumSecAddr.Visible = True
         '   txtNumSecAddr.Visible = True
         '   lblXNumSecAddr.Visible = True
         '   NumSecAddr$ = "Number of secondary addresses"
         End If
         lblXNumSecAddr.Caption = NumSecAddr$
         lblXSecondAddr.Caption = SecAddrString$
         lblSent.Caption = SentString$
         lblSent.Top = lblError.Top + 132
         txtData.Visible = True
         txtData.Text = "1"
         txtData.Left = txtSecondAddr.Left
         txtData.Width = txtSecondAddr.Width
         txtData.Top = lblError.Top + 64
         hsbData.Top = lblError.Top + 64
         hsbData.Min = 0
         hsbData.Max = 64
         hsbData.LargeChange = 8
         hsbData.value = 1
         hsbNumSecAddr.Top = lblError.Top + 664
         txtNumSecAddr.Top = lblError.Top + 664
         lblXNumSecAddr.Top = lblError.Top + 732
         lblSent.Visible = True
         lblSent.Width = 2500
      Case ShowMemCtls%, ShowMemWriteCtls%, ShowTeds%
         hsbSecondAddr.Top = lblReturned.Top + 364
         txtSecondAddr.Top = lblReturned.Top + 364
         txtSecondAddr.Visible = True
         txtSecondAddr.Width = 1200
         lblXSecondAddr.Top = lblReturned.Top + 432
         txtData.Visible = True
         txtData.Text = "0"
         txtData.Left = txtSecondAddr.Left
         txtData.Width = txtSecondAddr.Width
         txtData.Top = lblReturned.Top + 64
         lblSent.Top = lblReturned.Top + 132
         lblSent.Caption = "First point"
         lblSent.Visible = True
         lblXSecondAddr.Caption = "Number of points"
         lblXSecondAddr.Visible = True
         hsbData.Top = lblReturned.Top + 64
         hsbData.Min = 0
         hsbData.Max = 32000
         hsbData.LargeChange = 2048
         hsbData.value = 0
         hsbSecondAddr.Min = 0
         hsbSecondAddr.Max = 32000
         hsbSecondAddr.value = 1000
         txtAmpl.Top = lblError.Top + 400
         txtOffset.Top = lblError.Top + 772
         lblAmpl.Top = lblError.Top + 472
         lblOffset.Top = lblError.Top + 864
         If ShowMemWriteCtls% Then lblXNumSecAddr.Caption = ""
         lblSent.Left = 2800
         lblXSecondAddr.Left = 2800
         'hsbData.Top = 900
         If ShowTeds% Then
            lblSent.Caption = "Channel"
            lblXSecondAddr.Caption = "Number of items"
            hsbData.Max = 64
            hsbData.LargeChange = 16
            hsbData.value = 0
            hsbSecondAddr.Min = 0
            hsbSecondAddr.Max = 4096
            hsbSecondAddr.value = 512
         End If
      Case ShowWinBufCtls%
         hsbData.Top = Me.txtData.Top
         lblSent.Caption = "Buffer Size"
         txtData.Visible = True
         txtData.Text = "1000"
      Case ShowStopBGCtls%
         cmbDevNum.Visible = True
         cmbDevNum.Left = 2000
         cmbDevNum.Width = 2000
         cmbDevNum.Clear
         cmbDevNum.AddItem "AIFUNCTION"
         cmbDevNum.AddItem "AOFUNCTION"
         cmbDevNum.AddItem "DIFUNCTION"
         cmbDevNum.AddItem "DOFUNCTION"
         cmbDevNum.AddItem "CTRFUNCTION"
         cmbDevNum.AddItem "DAQIFUNCTION"
         cmbDevNum.AddItem "DAQOFUNCTION"
         cmbDevNum.ListIndex = 0
      Case ShowGetErrorCtls%
         hsbData.Min = -32768
         hsbData.Max = 32767
         hsbData.Top = Me.txtData.Top
         txtData.Visible = True
      Case ShowGetStatusCtls%
         cmbDevNum.Visible = True
         cmbDevNum.Left = 200
         cmbDevNum.Width = 2500
         lblSent.Visible = True
         lblSent.Top = 480
         lblSent.Left = 1500
         lblSent.Caption = "Status"
         txtData.Visible = True
         txtData.Top = 420
         txtData.Left = 200
         txtData.Width = 1200
         txtData.Text = ""
         lblXSecondAddr.Visible = True
         lblXSecondAddr.Left = 1500
         lblXSecondAddr.Top = 840
         lblXSecondAddr.Caption = "Count"
         txtSecondAddr.Visible = True
         txtSecondAddr.Top = 780
         txtSecondAddr.Left = 200
         txtSecondAddr.Width = 1200
         txtSecondAddr.Text = ""
         lblXNumSecAddr.Visible = True
         lblXNumSecAddr.Left = 1500
         lblXNumSecAddr.Top = 1200
         lblXNumSecAddr.Caption = "Index"
         txtNumSecAddr.Visible = True
         txtNumSecAddr.Top = 1140
         txtNumSecAddr.Left = 200
         txtNumSecAddr.Width = 1200
         txtNumSecAddr.Text = ""
         cmbDevNum.Clear
         cmbDevNum.AddItem "AIFUNCTION"
         cmbDevNum.AddItem "AOFUNCTION"
         cmbDevNum.AddItem "DIFUNCTION"
         cmbDevNum.AddItem "DOFUNCTION"
         cmbDevNum.AddItem "CTRFUNCTION"
         cmbDevNum.AddItem "DAQIFUNCTION"
         cmbDevNum.AddItem "DAQOFUNCTION"
         cmbDevNum.ListIndex = 0
      Case ShowLogIn%
         txtData.Visible = True
         txtSecondAddr.Visible = True
         txtData.Width = 1500
         txtSecondAddr.Width = 1500
         txtData.Top = cmdGo.Top
         txtSecondAddr.Top = txtData.Top + 360
         txtData.Text = "user"
         txtSecondAddr.Text = "mccdaq"
      Case Else
         lblReturned.Caption = ""
         lblSent.Caption = "Sent to " & mnuFuncArray(mnFuncType).Caption
         lblSent.Top = lblReturned.Top + 900
         lblError.Caption = ""
         If ShowConvCtls% Then
            txtData.Width = 1500
            txtData.Visible = True
         Else
            txtData.Width = txtSecondAddr.Width
         End If
         txtData.Top = lblReturned.Top + 840
         txtData.Left = txtSecondAddr.Left
         hsbData.Top = lblReturned.Top + 840
         hsbData.LargeChange = 512
         If (Abs(mnResolution) > 15) Or (mnFuncType = FROM_ENG) Then
            hsbData.Min = -32768
            hsbData.Max = 32767
         ElseIf mnResolution > 12 Then
            hsbData.Min = 0
            hsbData.Max = (2 ^ mnResolution) - 1
         Else
            hsbData.Min = 0
            hsbData.Max = 4095
         End If
         hsbNumSecAddr.Visible = False
         txtNumSecAddr.Visible = False
         lblXNumSecAddr.Visible = False
   End Select
   boardindex% = mnBoardIndex
   If chkSurrogate.value = 1 Then boardindex% = mnSurrogateIndex
   msTitle = mnuFuncArray(mnFuncType).Caption
   Caption = msTitle & msOpt & mnuBoard(boardindex%).Caption
   optFormat(0).Visible = False
   optFormat(1).Visible = False
   optFormat(2).Visible = False
   
   If ShowUSBCtls% Then
      mnuBipRange(0).Caption = "GAIN1_DIFF"
      mnuBipRange(1).Caption = "GAIN2_DIFF"
      mnuBipRange(2).Caption = "GAIN4_DIFF"
      mnuBipRange(3).Caption = "GAIN5_DIFF"
      mnuBipRange(4).Caption = "GAIN8_DIFF"
      mnuBipRange(5).Caption = "GAIN10_DIFF"
      mnuBipRange(6).Caption = "GAIN16_DIFF"
      mnuBipRange(7).Caption = "GAIN20_DIFF"
      mnuBipRange(8).Caption = "GAIN1_SE"
      mnuBipRange(9).Caption = "GAIN1_DAC"
      For USBMenu% = 10 To 17
         mnuFuncArray(USBMenu%).Visible = True
      Next
      If mnFuncType = USB_GETSERIAL Then lblError.Visible = True
      If (mnFuncType = USB_GETSERIAL) Or (mnFuncType = USB_MEMREAD) Or (mnFuncType = USB_MEMWRITE) Then lblReturned.Visible = True
      If (mnFuncType = USB_SETSERIAL) Or (mnFuncType = USB_MEMREAD) Or (mnFuncType = USB_MEMWRITE) Or (mnFuncType = USB_WATCH) Then hsbData.Visible = True
      If (mnFuncType = USB_SETSERIAL) Or (mnFuncType = USB_MEMREAD) Or (mnFuncType = USB_MEMWRITE) Then txtData.Visible = True
      If (mnFuncType = USB_MEMREAD) Or (mnFuncType = USB_MEMWRITE) Or (mnFuncType = USB_WATCH) Then txtSecondAddr.Visible = True
      If (mnFuncType = USB_RESET) Or (mnFuncType = USB_BLINK) Or (mnFuncType = SAVE_CONFIG) Then txtData.Visible = False
      If (mnFuncType = USB_MEMREAD) Or (mnFuncType = USB_MEMWRITE) Or (mnFuncType = USB_WATCH) Then
         hsbSecondAddr.Visible = True
         optFormat(0).Visible = True
         optFormat(1).Visible = True
         optFormat(2).Visible = True
         txtNumSecAddr.Visible = (mnFuncType = USB_MEMWRITE)
         lblReturned.Visible = (mnFuncType = USB_MEMREAD)
         txtNumSecAddr.Left = lblError.Left
         txtNumSecAddr.Top = lblError.Top
         txtNumSecAddr.Width = lblError.Width
      End If
      If (mnFuncType = USB_WATCH) Then
         optFormat(0).Caption = "Reset USB"
         optFormat(1).Caption = "Set DOut"
         optFormat(2).Caption = "Disable"
      Else
         optFormat(0).Caption = "Integers"
         optFormat(1).Caption = "Hex"
         optFormat(2).Caption = "Characters"
      End If
   Else
      mnuBipRange(0).Caption = "BIP20VOLTS"
      mnuBipRange(1).Caption = "BIP10VOLTS"
      mnuBipRange(2).Caption = "BIP5VOLTS"
      mnuBipRange(3).Caption = "BIP4VOLTS"
      mnuBipRange(4).Caption = "BIP2PT5VOLTS"
      mnuBipRange(5).Caption = "BIP2VOLTS"
      mnuBipRange(6).Caption = "BIP1PT25VOLTS"
      mnuBipRange(7).Caption = "BIP1VOLTS"
      mnuBipRange(8).Caption = "BIPPT625VOLTS"
      mnuBipRange(9).Caption = "BIPPT5VOLTS"
      mnuBipRange(10).Caption = "BIPPT25VOLTS"
      mnuBipRange(11).Caption = "BIPPT2VOLTS"
      'If mnFuncType > 9 Then mnuFuncArray_Click (0)
      For USBMenu% = 10 To 16
         mnuFuncArray(USBMenu%).Visible = False
      Next
   End If
   If (mnLibType = MSGLIB) And ShowFlashLEDCtls% Then
      Me.txtData.Text = "3"
   End If

End Sub

Private Sub CycleBuffer()

   BufferSize& = Val(txtData.Text)
   If Not (mvHandle = 0) Then
      'in UL16, the following can be called even if the buffer doesn't exist
      'not true in UL32
      ULStat = cbWinBufFree(mvHandle)
      If SaveFunc(Me, WinBufFree, ULStat, mvHandle, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
   mvHandle = cbWinBufAlloc(BufferSize&)
   If SaveFunc(Me, WinBufAlloc, mvHandle, BufferSize&, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   ULStat = cbWinBufFree(mvHandle)
   If ULStat = 0 Then mvHandle = 0
   If SaveFunc(Me, WinBufFree, ULStat, mvHandle, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
      'if right click, disable timer
      cmdStop_Click
      Me.mnuTimer.Checked = False
      cmdStop.Visible = False
   End If

End Sub

Private Sub Form_Activate()

   UpdateMainStatus

End Sub

Private Sub Form_Initialize()
   
   mnLoading = True
   Me.mnuLibrary(MSGLIB).ENABLED = frmMain.mnuLibrary(MSGLIB).ENABLED
   Me.mnuLibrary(UNILIB).ENABLED = frmMain.mnuLibrary(UNILIB).ENABLED

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   If Shift And 4 Then
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
            Me.Height = 2000
            Me.Width = 6435
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

#If MSGOPS Then
   If gnLibType = MSGLIB Then
      Set MsgLibrary = New MBDClass.MBDComClass
   End If
#Else
   Dim MsgLibrary As Object
#End If
   
   mnLibType = gnLibType
   mnNumBoards = gnNumBoards
   mnuLibrary(gnLibType).Checked = True
   If Not gbULLoaded Then
      Me.mnuULBuf.Checked = False
      mnuLibrary(UNILIB).Checked = False
      mnuLibrary(UNILIB).ENABLED = False
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
      UseList% = mnuUseList.Checked
      If i% > 0 Then
         Load mnuBoard(i%)
         mnuBoard(i%).Checked = False
      Else
         mnBoardNum = BoardNum%
         mnBoardIndex = 0
         mnResolution = GetADResolution(DisplayName$, mnBoardNum)
         msBoardName = BoardName$
         ConfigureControls
         mnuBipRange_Click (2)
      End If
      mnuBoard(i%).Caption = CurrentName$
   Next i%
   mnPlot = True
   mnTrigType = TRIG_POS_EDGE
   mnMaskVal = &HFF

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   mnUnloading = True

End Sub

Private Sub Form_Resize()

   lblStatus.Width = ScaleWidth
   lblStatus.Top = ScaleHeight - lblStatus.Height
   If gnInitializing Then
      'if the form is just loading, sets the form type
      'and sets the default function (cbToEngUnits())
      mnFormType = (Val("&H" & Tag) And &HF00&) / &H100
      mnThisInstance = Val("&H" & Tag) And &HFF
      'mnuFuncArray_Click (TO_ENG)
      'msTitle = Caption
      'mnuBoard_Click (0)
      gnInitializing = False
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   tmrReadPorts.ENABLED = False
   Select Case mnLibType
      Case UNILIB
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
   gnUtilForms = gnUtilForms - 1

End Sub

Private Sub hsbData_Change()

   hsbData_Scroll

End Sub

Private Sub hsbData_Scroll()
   
   CheckRange% = mnRange
   Select Case mnFuncType
      Case FROM_ENG
         If (Abs(mnResolution) >= 16) Or (mnFuncType = FROM_ENG) Then Maskval& = &H8000
         DataValue% = hsbData.value Xor Maskval&
         EngUnits! = GetVolts(16, CheckRange%, DataValue%)
         txtData.Text = EngUnits!
         ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, DataValue%)
         If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, mnRange, EngUnits!, DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         lblReturned.Caption = DataValue%
         FSR! = GetRangeVolts(mnRange)
         LSB! = FSR! / 2 ^ Abs(mnResolution)
         EngUnits! = EngUnits! + LSB! / 2
         CountsLong& = GetCounts(mnResolution, CheckRange%, EngUnits!)
         CalcCounts% = ULongValToInt(CountsLong&)
         lblError.Caption = CLng(CalcCounts%) - DataValue%
      Case TO_ENG
         If (Abs(mnResolution) >= 16) Or (mnFuncType = FROM_ENG) Then Maskval& = &H8000
         DataValue% = hsbData.value Xor Maskval&
         LongVal& = IntValToULong(DataValue%)
         If (Abs(mnResolution) = 24) Then
            DataVal& = LongVal& * 256
            txtData.Text = Format(DataVal&, "0")
            ULStat = cbToEngUnits32(mnBoardNum, mnRange, DataVal&, dEngUnits#)
            If SaveFunc(Me, ToEngUnits32, ULStat, mnBoardNum, mnRange, DataVal&, dEngUnits#, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            lblReturned.Caption = dEngUnits#
         Else
            txtData.Text = Format(LongVal&, "0")
            ULStat = cbToEngUnits(mnBoardNum, mnRange, DataValue%, EngUnits!)
            'result can be infinite (NaN or 1.#INF)
            Dim NewVar As Variant
            NewVar = EngUnits!
            If ((NewVar - 1) = NewVar) Then
               'if above is true, number isn't real
               A4 = "Undefined"
            Else
               A4 = EngUnits!
            End If
            If SaveFunc(Me, ToEngUnits, ULStat, mnBoardNum, mnRange, DataValue%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            lblReturned.Caption = EngUnits!
         End If
         
         If (Abs(mnResolution) = 24) Then
            If CheckRange% = mnFixedRange Then
               CalcVal# = GetHiResVolts(mnResolution, CheckRange%, DataVal&, mvCustomRange)
            Else
               CalcVal# = GetHiResVolts(mnResolution, CheckRange%, DataVal&)
            End If
            lblError.Caption = CalcVal# - dEngUnits# & " volts"
         Else
            CalcVal# = GetVolts(mnResolution, CheckRange%, DataValue%)
            lblError.Caption = CalcVal# - EngUnits! & " volts"
         End If
         
      Case REG_IN, REG_OUT, MEM_READ ', GET_ERROR
         txtData.Text = hsbData.value
      Case USB_MEMREAD, USB_MEMWRITE, READ_TEDS
         txtData.Text = hsbData.value
      Case GET_ERROR
         ErrCode% = hsbData.value
         txtData.Text = ErrCode%
         If mnLibType = MSGLIB Then
            ErrMessage$ = MsgLibrary.ConvertErrorCode(ErrCode%)
            lblStatus.Caption = "Code " & Format(ErrCode%, "0") & " (" & ErrMessage$ & ")" & LineFeed$
         Else
            ErrMessage$ = Space$(ERRSTRLEN)
            ULStat = cbGetErrMsg(ErrCode%, ErrMessage$)
            If SaveFunc(Me, GetErrMsg, ULStat, ErrCode%, ErrMessage$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            ErrMessage$ = RTrim$(ErrMessage$)   'Drop the space characters
            If Len(ErrMessage$) Then ErrMessage$ = Left$(ErrMessage$, Len(ErrMessage$) - 1)
            lblStatus.Caption = Str$(ErrCode%) & " (" & GetErrorConst(ErrCode%) & "): " & ErrMessage$
         End If
   End Select

End Sub

Private Sub hsbNumSecAddr_Change()

   NumAddr% = hsbNumSecAddr.value
   txtNumSecAddr.Text = NumAddr%

End Sub

Private Sub hsbSecondAddr_Change()

   If (mnFuncType = USB_MEMWRITE) Or (mnFuncType = USB_MEMREAD) Then
      txtSecondAddr.Text = hsbSecondAddr.value
   Else
      StepVal% = 1
      If mnWordMode Then StepVal% = 2
      AddrOffset% = hsbSecondAddr.value * StepVal%
      'If mnFuncType = REG_IN Then
         'Addr2Offset% = hsbSecondAddr.Value
         txtSecondAddr.Text = AddrOffset%
         'hsbNumSecAddr.Visible = (Addr2Offset% > 0)
         'txtNumSecAddr.Visible = (Addr2Offset% > 0)
         'lblXNumSecAddr.Visible = (Addr2Offset% > 0)
      'Else
      '   WritePort% = hsbSecondAddr.Value
      '   txtSecondAddr.Text = WritePort%
      'End If
   End If

End Sub

Private Sub lblReturned_DblClick()

   Clipboard.SetText Mid$(lblReturned.Caption, 2)
   MsgBox "Address copied to clipboard"

End Sub

Private Sub LoadConfig()

   Filename$ = ""
   ULStat = GetCfgFile(Filename$)
   'the value returned by cbLoadConfig() is
   'not an error code

End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash

End Sub

Private Sub mnuAConv_Click()

   mnuAConv.Checked = Not mnuAConv.Checked
   SetVal% = mnuAConv.Checked
   SetConvertData mnBoardNum, SetVal%

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
               InitForm FLASH_LED
               InitNewInstance% = True
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
         msBoardName = BoardName$
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
   SetResolution
   
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
   'Caption = msTitle & " Board " & mnuBoard(Index).Caption
   If gnNumBoards = 0 Then Exit Sub
   chkSurrogate.value = 0
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
      If Not SaveMsg(Me, "CreateDevice(" & BoardName$ & ")", MBDResponse$) Then mnMessaging = True
   Else
      BoardName$ = GetNameOfBoard(mnBoardNum)
      DisplayName$ = BoardName$
   End If
   msBoardName = BoardName$
   msDisplayName = DisplayName$
   If mnLibType = UNILIB Then
      CurRange& = GetManualRange(msDisplayName, mnBoardNum, FixedRange&)
      mnFixedRange = FixedRange&
      CustomRange = GetCustomRange(msDisplayName)
      If Not (CustomRange = "") Then
         'undefined range
         mvCustomRange = GetCustomRange(msDisplayName)
      Else
         mvCustomRange = Null
      End If
      SetResolution
      SetOpts  'and ConfigureControls
      'ConfigureControls
      UpdateMainStatus
      InitMathModule msDisplayName
      InitLocModule msDisplayName
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
         'Print #2, Me.Tag & ","; FuncID%; ",0,"; A1; ","; A2; ","; A3; ","; A4; ","; A5; ","; A6; ","; A7; ","; A8; ","; A9; ","; A10; ","; A11; ","; AuxHandle
      End If
      mnuBipRange_Click (2)
   End If
   ConfigureControls

End Sub

Private Sub mnuBoardName_Click()

   SetPlotType PRINT_LIST, Me
   frmPlot.txtShow.Text = ""
   LineFeed$ = Chr$(13) & Chr$(10)
   
   If mnLibType = MSGLIB Then
      BoardNames$ = GetAllMsgBoardNames()
      NameArray = Split(BoardNames$, "|")
      For Element% = 0 To UBound(NameArray)
         BoardName$ = NameArray(Element%)
         BoardList$ = BoardList$ & BoardName$ & LineFeed$
      Next
   Else
      'Get the first board type in list of supported boards
      'The first string in the boardlist is "Not Supported"
      BoardName$ = Space$(BOARDNAMELEN) 'to do - replace BOARDNAMELEN when > 25
      
      BoardNum% = GETFIRST
      ULStat = cbGetBoardName(BoardNum%, BoardName$)
      If Len(BoardName$) Then
         'Drop the space characters
         BoardName$ = RTrim$(BoardName$)
         StringSize% = Len(BoardName$)
         'lop off the null
         If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
         'msBoardName = BoardName$
      End If
      If SaveFunc(Me, GetBoardName, ULStat, BoardNum%, BoardName$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      BoardList$ = BoardList$ & BoardName$ & LineFeed$
      
      'Get each consecutive board type in list
      Do
         BoardName$ = Space$(BOARDNAMELEN) 'to do - replace BOARDNAMELEN when > 25
         BoardNum% = GETNEXT
         ULStat = cbGetBoardName(BoardNum%, BoardName$)
         If Len(BoardName$) Then
            'Drop the space characters
            BoardName$ = RTrim$(BoardName$)
            StringSize% = Len(BoardName$)
            'lop off the null
            If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
         End If
         If SaveFunc(Me, GetBoardName, ULStat, BoardNum%, BoardName$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         BoardList$ = BoardList$ & BoardName$ & LineFeed$
      Loop While Len(BoardName$) > 3
   End If
   TextList BoardList$
   
End Sub

Private Sub mnuClose_Click()

   Unload Me

End Sub

Private Sub mnuComposite_Click()

   If mnuComposite.Checked Then frmComposite.chkComposite.value = 1
   If mnMaskFirst Then frmComposite.chkMaskFirst.value = 1
   If mnMaskSecond Then frmComposite.chkMaskSecond.value = 1
   If mnConsecutive Then frmComposite.chkConsecutive.value = 1
   frmComposite.Show 1
   mnuComposite.Checked = (frmComposite.chkComposite.value = 1)
   mnMaskFirst = (frmComposite.chkMaskFirst.value = 1)
   mnMaskSecond = (frmComposite.chkMaskSecond.value = 1)
   mnConsecutive = (frmComposite.chkConsecutive.value = 1)
   Unload frmComposite
   SetOpts  'and ConfigureControls

End Sub

Private Sub mnuCurRange_Click(Index As Integer)

   SetRange 2, Index

End Sub

Private Sub mnuData_Click(Index As Integer)
'select the pattern of data for output array

   mnuData(mnDataType).Checked = False
   mnDataType = Index
   mnuData(mnDataType).Checked = True
   If Not gnInitializing Then ConfigureControls
   If Index > 0 Then
      ResetData
      cmdGo.Visible = True
   End If
   
End Sub

Private Sub mnuFuncArray_Click(Index As Integer)

   Me.lblStatus.Caption = ""
   If Index < 2 Then mnuAConv.Checked = False
   If Index < 5 Then
      chkSurrogate.value = 0
      chkSurrogate.Visible = 0
      cmbSurrogateBoard.Visible = 0
   End If
   mnuFuncArray(mnFuncType).Checked = False
   mnFuncType = Index
   msConfig = mnuFuncArray(mnFuncType).Caption
   msTitle = msConfig
   mnuFuncArray(mnFuncType).Checked = True
   mnuComposite.Visible = mnuFuncArray(5).Checked Or mnuFuncArray(6).Checked
   mnuMask.Visible = mnuFuncArray(5).Checked Or mnuFuncArray(6).Checked
   
   lblXReturn.Caption = "Returned from " & mnuFuncArray(mnFuncType).Caption
   Resolution% = 12
   UseList% = mnuUseList.Checked

   If Not (gnNumBoards = 0) Then
      Select Case mnFuncType
         Case TO_ENG
            Resolution% = GetADResolution(msDisplayName, mnBoardNum)
         Case FROM_ENG
            Resolution% = GetDAResolution(msDisplayName, mnBoardNum, mnRange)
      End Select
   End If
   If mnResolution <> Resolution% Then
      mnResolution = Resolution%
   End If
   If Not gnInitializing Then ConfigureControls
   boardindex% = mnBoardIndex 'GetBoardIndex(mnBoardNum)
   If chkSurrogate.value = 1 Then boardindex% = mnSurrogateIndex

   SetOpts  'and ConfigureControls
   
End Sub

Private Sub mnuMask_Click()

   Maskval$ = InputBox$("Enter mask value (precede with '&H' if hex): ", "Mask Reads", "&H" & Hex$(mnMaskVal))
   If Len(Maskval$) Then mnMaskVal = Val(Maskval$)
   If mnWordMode Then
      If mnMaskVal <> &HFFFF Then
         If (Not mnuComposite.Checked) Then mnMaskFirst = True
      Else
         If (Not mnuComposite.Checked) Then mnMaskFirst = False
      End If
   Else
      If mnMaskVal <> &HFF Then
         If (Not mnuComposite.Checked) Then mnMaskFirst = True
      Else
         If (Not mnuComposite.Checked) Then mnMaskFirst = False
      End If
   End If

End Sub

Private Sub mnuPlotType_Click(Index As Integer)

   mnPlot = False
   mnuPlotType(Index).Checked = Not mnuPlotType(Index).Checked
   If mnuPlotType(Index).Checked Then
      SetPlotType Index, Me
      mnPlot = True
   End If
   For i% = 0 To 2
      If i% <> Index Then mnuPlotType(i%).Checked = False
   Next i%
   'DisplayData

End Sub

Private Sub mnuQuit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub mnuShowData_Click()

   BufSize& = mlCount
   BufferStat% = CheckBuffer(mvHandle, BufSize&)
   Select Case BufferStat%
      Case INVALIDBUFFER
         MsgBox "Windows buffer no longer valid"
      Case BUFFERTOOSMALL
         MsgBox "Windows buffer too small (size = " & BufSize& & ", data requested = " & mlCount * 2 & ")"
      Case NOBUFFER
         MsgBox "No Windows buffer has been allocated"
      Case Else
         PrintMain "Buffer handle = 0x" & Hex$(mvHandle) & ",  buffer size = " & BufSize&
         PlotData% = True
   End Select
   If PlotData% Then
      LongHandle& = mvHandle
      PlotBuffer LongHandle&, mlCount, 0
   End If

End Sub

Private Sub mnuTimer_Click()

   If mnuTimer.Checked Then frmSetTimer.chkEnableTimer.value = 1
   frmSetTimer.txtInterval.Text = tmrReadPorts.Interval
   frmSetTimer.optTimerMode(1).value = mnTimerTillCount
   frmSetTimer.Show 1
   LoopRate& = Val(frmSetTimer.txtInterval.Text)
   mnuTimer.Checked = (frmSetTimer.chkEnableTimer.value = 1)
   Unload frmSetTimer
   If LoopRate& > 0 Then
      tmrReadPorts.Interval = LoopRate&
   Else
      mnuTimer.Checked = False
   End If
   If Not mnuTimer.Checked Then tmrReadPorts.ENABLED = False
   cmdStop.Visible = mnuTimer.Checked

End Sub

Private Sub mnuULBuf_Click()

   mnuULBuf.Checked = Not (mnuULBuf.Checked)

End Sub

Private Sub mnuUniRange_Click(Index As Integer)

   SetRange 1, Index

End Sub

Private Sub mnuUseList_Click()

   mnuUseList.Checked = Not mnuUseList.Checked
   UseList% = mnuUseList.Checked
   If mnFuncType = FROM_ENG Then
      mnResolution = GetDAResolution(msDisplayName, UseList%, mnRange)
   Else
      mnResolution = GetADResolution(msDisplayName, mnBoardNum)
   End If

End Sub

Private Sub mnuWord_Click()

   mnuWord.Checked = Not mnuWord.Checked
   mnWordMode = mnuWord.Checked
   If mnWordMode Then
      If (mnMaskVal = &HFF) Then mnMaskVal = &HFFFF
   Else
      If (mnMaskVal = &HFFFF) Then mnMaskVal = &HFF
   End If
   ConfigureControls

End Sub

Private Sub ReadMem()

   FirstPoint& = Val(txtData.Text)
   NumPoints& = Val(txtSecondAddr.Text)
   PlotMemory mnBoardNum, mnBoardNum, NumPoints&, FirstPoint&, Chans%

End Sub

Private Sub ReadRegisters()

   SetPlotType PRINT_LIST, Me
   frmPlot.txtShow.Text = ""
   LineFeed$ = Chr$(13) + Chr$(10)
   StepVal% = 1
   If mnWordMode Then StepVal% = 2
   address& = mlCurAddress

   HowMany% = Val(txtData.Text)
   board% = mnBoardNum
   If chkSurrogate.value = 1 Then
      Surrogate& = Val(Mid$(lblReturned.Caption & "&", 2))
   End If
   FirstReg& = Val(txtSecondAddr.Text) + Surrogate&
   LastReg& = FirstReg& + ((HowMany% - 1) * StepVal%)
   'following info needed for scripting
   A2 = FirstReg&
   A3 = mnuComposite.Checked
   A4 = mnConsecutive
   A5 = mnMaskVal
   A6 = mnMaskFirst
   A7 = mnMaskSecond
   A8 = chkSurrogate.value
   A9 = cmbSurrogateBoard.ListIndex
   A10 = cmbDevNum.ListIndex
   For PrimReg& = FirstReg& To LastReg& Step StepVal%
      RegNum& = PrimReg&
      A2 = RegNum&
      If gnScriptSave Then
         A2 = Val(txtSecondAddr.Text) + Iterations%
         Iterations% = Iterations% + StepVal%
      End If
      If mnWordMode Then
         MaskVal1% = &HFFFF
         If mnMaskFirst Then MaskVal1% = mnMaskVal
         PortVal = cbInWord(board%, RegNum&) 'And MaskVal1%
         If SaveFunc(Me, InWord, PortVal, board%, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         PortVal = PortVal And MaskVal1%
         WordLen% = 4
         If (MaskVal1% = &HFF) Or (MaskVal1% = &HFF00) Then WordLen% = 2
         If VarType(PortVal) = 2 Then
            LSBVal$ = Hex$(IntValToULong(Int(PortVal)))
         Else
            LSBVal$ = Hex$(PortVal)
         End If
         LSBVal$ = Right$(String$(4 - Len(LSBVal$), "0") & LSBVal$, WordLen%)
         If mnuComposite.Checked Then
            If mnConsecutive Then NextReg% = 2
            MaskVal2% = &HFFFF
            If mnMaskSecond Then MaskVal2% = mnMaskVal
            PortVal2 = cbInWord(board%, RegNum& + NextReg%) 'And MaskVal2%
            If Not gnScriptSave Then
               If SaveFunc(Me, InWord, PortVal2, board%, A2 + NextReg%, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            End If
            PortVal2 = CInt(PortVal2 And MaskVal2%)
            WordLen% = 4
            If (MaskVal2% = &HFF) Or (MaskVal2% = &HFF00) Then WordLen% = 2
            If VarType(PortVal2) = 2 Then
               MSBVal$ = Hex$(IntValToULong(Int(PortVal2)))
            Else
               MSBVal$ = Hex$(PortVal2)
            End If
            MSBVal$ = Right$(String$(4 - Len(MSBVal$), "0") & MSBVal$, WordLen%)
         End If
      Else
         MaskVal1% = &HFF
         If mnMaskFirst Then MaskVal1% = mnMaskVal
         PortVal = InByte1632(board%, RegNum&) 'And MaskVal1%
         If SaveFunc(Me, InByte, PortVal, board%, RegNum&, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         PortVal = PortVal And MaskVal1%
         WordLen% = 2
         If (MaskVal1% = &HF0) Or (MaskVal1% = &HF) Then WordLen% = 1
         LSBVal$ = Hex$(PortVal)
         If (mnMaskVal = &HF0) Then
            LSBVal$ = Left$(String$(2 - Len(LSBVal$), "0") & LSBVal$, WordLen%)
         Else
            LSBVal$ = Right$(String$(2 - Len(LSBVal$), "0") & LSBVal$, WordLen%)
         End If
         If mnuComposite.Checked Then
            If mnConsecutive Then NextReg% = 1
            MaskVal2% = &HFF
            If mnMaskSecond Then MaskVal2% = mnMaskVal
            PortVal2 = InByte1632(board%, RegNum& + NextReg%) 'And MaskVal2%
            If Not gnScriptSave Then
               If SaveFunc(Me, InByte, PortVal2, board%, RegNum& + NextReg%, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            End If
            PortVal2 = PortVal2 And MaskVal2%
            WordLen% = 2
            If (MaskVal2% = &HF) Or (MaskVal2% = &HF0) Then WordLen% = 1
            MSBVal$ = Hex$(PortVal2)
            MSBVal$ = String$(2 - Len(MSBVal$), "0") & MSBVal$
         End If
      End If
      RegValString$ = RegValString$ & "Port 0x" & Hex$(address& + RegNum&) & ": 0x" & MSBVal$ & LSBVal$ & LineFeed$
   Next PrimReg&
   RegValString$ = RegValString$ & LineFeed$ & LineFeed$
   TextList RegValString$

End Sub

Private Sub ResetData()

   If txtSecondAddr.Text > "2147483647" Then Exit Sub
   mlCount = Val(txtSecondAddr.Text)
   If mnFuncType = MEM_WRITE Then
      Amplitude& = Val(txtAmpl.Text)
      Offset& = Val(txtOffset.Text)
      If mnDataType > 0 Then
         UseLibrary% = mnuULBuf.Checked
         InitOutputBuffer mvHandle, mnDataType, mlCount, 0, Amplitude&, Offset&, UseLibrary%
         txtSecondAddr.Text = mlCount
      End If
   End If

End Sub

Private Sub SendData(DataValue)

End Sub

Private Sub SetOpts()

   msOpt = ""
   If chkSurrogate.value = 1 Then msOpt = " Surrogate "
   If mnuComposite.Checked Then
      msOpt = msOpt & " Composite "
      If mnConsecutive Then msOpt = msOpt & " Consecutive "
   End If
   If mnMaskFirst Or mnMaskSecond Then msOpt = msOpt & " Masked "
   msOpt = msOpt & " Board "
   ConfigureControls

End Sub

Private Sub SetRange(RangeType As Integer, RangeIndex As Integer)

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
            BIPPT01VOLTS, BIPPT005VOLTS, BIP1PT67VOLTS, BIPPT156VOLTS, BIPPT312VOLTS, _
            BIPPT078VOLTS, BIP60VOLTS, BIP15VOLTS, BIPPT125VOLTS, _
            BIPPT025VOLTSPERVOLT, BIPPT073125VOLTS)
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
         'varRange = Choose(RangeIndex + 1, UNI10VOLTS, UNI5VOLTS, UNI2PT5VOLTS, UNI2VOLTS, UNI1PT25VOLTS, UNI1VOLTS, UNIPT5VOLTS, UNIPT25VOLTS, UNIPT2VOLTS)
         varRange = Choose(RangeIndex + 1, UNI10VOLTS, UNI5VOLTS, UNI4VOLTS, UNI2PT5VOLTS, UNI2VOLTS, UNI1PT25VOLTS, UNI1VOLTS, UNIPT5VOLTS, UNIPT25VOLTS, UNIPT2VOLTS)
         If IsNull(varRange) Then
            NewRange = RangeIndex - 9
            varRange = Choose(NewRange, UNIPT1VOLTS, UNIPT05VOLTS, UNIPT01VOLTS, UNIPT02VOLTS, UNI1PT67VOLTS)
         End If
         If ((IsNull(varRange)) Or (varRange = "")) Then
            MsgBox "Range not supported before revision 5.21.", , "UL Update Required"
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
            RangeVolts! = RangeVolts! / RangeDivisor!
            mnuRange.Caption = "&Range (" & Prefix$ & RangeVolts! & Suffix$ & ")"
         Else
            mnuRange.Caption = "&Range (NOTUSED)"
         End If
   End Select
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

   UseList% = mnuUseList.Checked
   If mnFuncType = FROM_ENG Then
      Resolution% = GetDAResolution(msDisplayName, mnBoardNum, mnRange)
   Else
      Resolution% = GetADResolution(msDisplayName, mnBoardNum)
   End If
   mnResolution = Resolution%
   'If (mnBoardNum < 0) And (mnRange = GAIN1_SE) Then mnResolution = 11
   'as of 8/03, both UL and OEM libs treat data as 12 bit (though res is actually 11 bit)

End Sub

Private Sub tmrReadPorts_Timer()

   cmdGo.FontBold = Not cmdGo.FontBold
   If mnFuncType = REG_IN Then
      ReadRegisters
   Else
      cmdGo_Click
   End If

End Sub

Private Sub txtAmpl_LostFocus()

   ResetData

End Sub

Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = 13 Then
      CheckRange% = mnRange
      Select Case mnFuncType
         Case FROM_ENG
            EngUnits! = txtData.Text
            ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, DataValue%)
            If SaveFunc(Me, FromEngUnits, ULStat, mnBoardNum, mnRange, EngUnits!, DataValue%, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            lblReturned.Caption = DataValue%
            FSR! = GetRangeVolts(mnRange)
            LSB! = FSR! / 2 ^ Abs(mnResolution)
            EngUnits! = EngUnits! + LSB! / 2
            CountsLong& = GetCounts(mnResolution, CheckRange%, EngUnits!)
            CalcCounts% = ULongValToInt(CountsLong&)
            lblError.Caption = CLng(CalcCounts%) - DataValue%
         Case TO_ENG
            'If (Abs(mnResolution) = 16) Or (mnFuncType = FROM_ENG) Then Maskval& = &H8000
            LongData& = Val(txtData.Text) 'Xor Maskval&
            DataValue% = ULongValToInt(LongData&)
            'txtData.Text = IntValToULong(DataValue%)
            ULStat = cbToEngUnits(mnBoardNum, mnRange, DataValue%, EngUnits!)
            If SaveFunc(Me, ToEngUnits, ULStat, mnBoardNum, mnRange, DataValue%, EngUnits!, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            lblReturned.Caption = EngUnits!
            CalcVal! = GetVolts(mnResolution, CheckRange%, DataValue%)
            lblError.Caption = CalcVal! - EngUnits! & " volts"
         Case REG_IN, REG_OUT, MEM_READ ', GET_ERROR
            txtData.Text = hsbData.value
         Case GET_ERROR
            ErrCode% = txtData.Text
            'txtData.Text = ErrCode%
            ErrMessage$ = Space$(ERRSTRLEN)
            ULStat = cbGetErrMsg(ErrCode%, ErrMessage$)
            If SaveFunc(Me, GetErrMsg, ULStat, ErrCode%, ErrMessage$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            ErrMessage$ = RTrim$(ErrMessage$)   'Drop the space characters
            If Len(ErrMessage$) Then ErrMessage$ = Left$(ErrMessage$, Len(ErrMessage$) - 1)
            lblStatus.Caption = Str$(ErrCode%) & " (" & GetErrorConst(ErrCode%) & "): " & ErrMessage$
      End Select
   End If

End Sub

Private Sub txtOffset_LostFocus()

   ResetData

End Sub

Private Sub txtSecondAddr_LostFocus()

   ResetData

End Sub

Private Sub UpdateMainStatus()
   
   board$ = mnuBoard(mnBoardIndex).Caption
   PrintMain "Current board: " & board$

End Sub

Private Sub WriteMem()

   If mlCount < 32768 Then
      ReDim MemData(0, mlCount) As Integer
   Else
      MsgBox "Not yet configured for large arrays"
      Exit Sub
   End If
   ULStat = cbWinBufToArray(mvHandle, MemData(0, 0), 0, mlCount)
   If SaveFunc(Me, WinBufToArray, ULStat, mvHandle, MemData(0, 0), 0, mlCount, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

   FirstPoint& = Val(txtData.Text)
   ULStat = cbMemWrite(mnBoardNum, MemData(0, 0), FirstPoint&, mlCount)
   If SaveFunc(Me, MemWrite, ULStat, mnBoardNum, MemData(0, 0), FirstPoint&, mlCount, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

End Sub

Private Sub WriteRegisters()

   board% = mnBoardNum
   If chkSurrogate.value = 1 Then
      Surrogate& = Val(Mid$(lblReturned.Caption, 2))
      'Board% = mnSurrogateBoard
   End If
   RegToWrite& = Val(txtSecondAddr.Text) + Surrogate&
   If (mnuComposite.Checked) And Not mnConsecutive Then
      Total% = Val(txtData.Text)
      LSB% = Total% And &HFF
      MSB% = (Total% And &HFF00&) / &H100
      ValToWrite% = LSB%
   Else
      ValToWrite% = Val(txtData.Text)
   End If
   A2 = RegToWrite&
   If gnScriptSave Then A2 = Val(txtSecondAddr.Text)
   A3 = ValToWrite%
   If gnScriptSave Then A3 = Val(txtData.Text)
   A4 = mnuComposite.Checked
   A5 = mnConsecutive
   A6 = mnMaskVal
   A7 = mnMaskFirst
   A8 = mnMaskSecond
   A9 = chkSurrogate.value
   A10 = cmbSurrogateBoard.ListIndex
   A11 = cmbDevNum.ListIndex

   If mnWordMode Then
      ULStat = OutWord1632(board%, RegToWrite&, ValToWrite%)
      If SaveFunc(Me, OutWord, ULStat, board%, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      If mnuComposite.Checked Then
         ULStat = cbOutWord(board%, RegToWrite&, MSB%)
         If Not gnScriptSave Then
            If SaveFunc(Me, OutWord, ULStat, board%, A2, MSB%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
      End If
   Else
      ULStat = OutByte1632(board%, RegToWrite&, ValToWrite%)
      If SaveFunc(Me, OutByte, ULStat, board%, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      If mnuComposite.Checked Then
         ULStat = cbOutByte(board%, RegToWrite&, MSB%)
         If Not gnScriptSave Then
            If SaveFunc(Me, OutByte, ULStat, board%, A2, MSB%, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         End If
      End If
   End If

End Sub

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
   
End Sub

Public Function GetMsgDevice() As String

   BoardName$ = ""
   If Not MsgLibrary Is Nothing Then
      BoardName$ = msBoardName
   End If
   GetMsgDevice = BoardName$
   
End Function

Sub ReadBoardName()
   
   BoardNum% = Me.txtData.Text
   BoardName$ = Space$(BOARDNAMELEN) 'to do - replace BOARDNAMELEN when > 25
   ULStat = cbGetBoardName(BoardNum%, BoardName$)
   If Len(BoardName$) Then
      'Drop the space characters
      BoardName$ = RTrim$(BoardName$)
      StringSize% = Len(BoardName$)
      'lop off the null
      If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
   End If
   x% = SaveFunc(Me, GetBoardName, ULStat, BoardNum%, BoardName$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   Me.lblStatus.Caption = BoardName$
   
End Sub
