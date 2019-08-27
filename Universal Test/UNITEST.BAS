Attribute VB_Name = "UNITEST1"
'Description
'  loads and configures child forms for the unitest
'  project

'Dependencies
'  UniTest.frm, AIn.frm
'____________________________________________________
'added for 5.88 but not in header as of 12/16/08
'not officially supported for VB due to data type manipulation required
Declare Function cbCLoad64 Lib "cbw32.dll" (ByVal BoardNum&, ByVal RegNum&, ByVal LoadValue As Currency) As Long
Declare Function cbCIn64 Lib "cbw32.dll" (ByVal BoardNum&, ByVal CounterNum&, CBCount As Any) As Long
'Declare Function cbWinBufToArray64 Lib "cbw32.dll" (ByVal MemHandle&, VarBuffer, ByVal FirstPoint&, ByVal CBCount&) As Long
Declare Function cbWinBufToArray64 Lib "cbw32.dll" (ByVal MemHandle&, DataBuffer As Currency, ByVal FirstPoint&, ByVal CBCount&) As Long


'added for DIO32
'Global Const DWORDXFER = &H200&
'Global Const TRIG_PATTERN_EQ = 16
'Global Const TRIG_PATTERN_NQ = 17
'Global Const TRIG_PATTERN_ABOVE = 18
'Global Const TRIG_PATTERN_BELOW = 19
'Global Const AUXPORT2 = 3

'ul constants not in VB header

'Global Const ON_END_OF_INPUT_SCAN = &H10
'Global Const ON_END_OF_OUTPUT_SCAN = &H20

'Global Const CTR_TICK20ns = 10
'Global Const CTR_TICK200ns = 11
'Global Const CTR_TICK2000ns = 12
'Global Const CTR_TICK20000ns = 13

'Global Const DIGITAL = 16
'Global Const CTR = 17
Global Const BICTR0SRC = 104     'CTR 0 source */
Global Const BICTR1SRC = 105     'CTR 1 source */
Global Const BICTR2SRC = 106     'CTR 2 source */
Global Const BIPACERCTR0SRC = 107 'Pacer CTR 0 source */
Global Const BITRIGEDGE = 113    'AD Pacer edge (1 rising, 2 falling)
Global Const BIADCFG = 117       'AD Config (SE/DIFF) (DevNo)
'Global Const BITEMPREJFREQ = 121 'DAS-TEMP reject frequency */
Global Const BICTR3SRC = 130     'CTR 3 source */
Global Const BICTR4SRC = 131     'CTR 4 source */
Global Const BICTR5SRC = 132     'CTR 5 source */
Global Const BIDACTRIG = 148     'DAC pacer edge (1 rising, 2 falling)
Global Const BITCCHANTYPE = 169  'J=1,K=2,T=3,E=4,R=5,S=6,B=7,N=8
Global Const BIAIWAVETYPE = 202  'analog input wave type (for demo board)
'Global Const BINUMTEMPCHANS = 208
'Global Const BIADTRIGSRC = 209   'Analog trigger source */
Global Const BIBNCSRC = 210      'BNC source */
Global Const BIBNCTHRESHOLD = 211 'BNC Threshold 2.5V or 0.0V */
Global Const BITEMPSENSORTYPE = 235
'Global Const BIADAIMODE = 249    'AD Config (SE/DIFF) (DevNo)

'Global Const BICTRTRIGCOUNT = 307
'Global Const BIDAQITRIGCOUNT = 308

Global Const BIADNUMCHANMODES = 312             '/* Number of channel modes */
Global Const BIADCHANMODE = 313                 '/* Channel mode */
Global Const BIADNUMDIFFRANGES = 314            '/* Number of differncial ranges supported by devide */
Global Const BIADDIFFRANGE = 315
Global Const BIADNUMSERANGES = 316              '/* Number of Single-Ended ranges supported by devide */
Global Const BIADSERANGE = 317

Global Const BIADSCANOPTIONS = 326
Global Const BIDACSCANOPTIONS = 332
Global Const BIDACNUMRANGES = 334
Global Const BIDACDEVRANGE = 335
Global Const BIDISCANOPTIONS = 367
Global Const BIDOSCANOPTIONS = 368
Global Const BICTRSCANOPTIONS = 369
Global Const BIDAQISCANOPTIONS = 370
Global Const BIDAQOSCANOPTIONS = 371
'Global Const DIDISABLEDIRCHECK = 13
'for eventual inclusion deprecating BLOCKIO
Global Const BLOCK_IO = &H40000000

'ranges used for message daq
Global Const BIPPT073125VOLTS = 73    ' -0.073125 to +0.073125 Volts
Global Const RAW = 6

Global Const ENCDR = 11
Global Const CTR48 = 12

Global Const ENABLED = 1
Global Const DISABLED = 0
' 25xx series counter tick size constants */
'Global Const CTR_TICK20PT83ns = 0
'Global Const CTR_TICK208PT3ns = 1
'Global Const CTR_TICK2083PT3ns = 2
'Global Const CTR_TICK20833PT3ns = 3

Global Const MAIN_FORM = 0
Global Const ANALOG_IN = 1
Global Const ANALOG_OUT = 2
Global Const ANALOG_IO = 1 'in and out are actually the same form
Global Const DIGITAL_IN = 3
Global Const DIGITAL_OUT = 4
Global Const DIGITAL_IO = 3 'in and out are actually the same form
Global Const COUNTERS = 5
Global Const UTILITIES = 6
Global Const Config = 7
Global Const GPIB_CTL = 8
Global Const LOGFUNC = 9

Global Const QUADENC = 11     'encoder channel
Global Const GIINIT = 44

'library used (gnLibType values)
Global Const INVALIDLIB = -1
Global Const UNILIB = 0
Global Const NETLIB = 1
Global Const MSGLIB = 2

'threading (.net) states
Global Const TS_IDLE = 0
Global Const TS_PRECALL = 1
Global Const TS_FIRSTCHECK = 2
Global Const TS_RUNNING = 4
Global Const TS_TMRCHK = 8
Global Const TS_ERROR = 256

Global gbULLoaded As Boolean
Global gnXternalCtl As Integer
Global gnInitializing As Integer
Global gnDIOForms As Integer, gnAIOForms As Integer, gnCTRForms As Integer
Global gnUtilForms As Integer, gnCfgForms As Integer, gnGPIBCtlForms As Integer
Global gnLoggerForms As Integer
Global gnBoardEnum() As Integer, gnNumBoards As Integer
Global gnTimeFactor As Long
Global gnThreading As Integer
Global gnLibType As Integer
Global gnICalDisable As Integer

Global frmNewAnalog() As Form, frmNewDigital() As Form, frmNewCounter() As Form
Global frmNewUtil() As Form, frmNewCfg() As Form, frmNewGPIBCtl() As Form
Global frmNewLogger() As Form
'Global frmScriptDIO() As Form
Global gn540 As Integer
Global gnIDERunning As Integer
Global gfDLLRev As Single

Dim smMainString As String

Dim mnGPIBError As Integer, mnAnalogInError As Integer
Dim mnAnalogOutError As Integer, mnDigitalInError As Integer
Dim mnDigitalOutError As Integer, mnCounterError As Integer
Dim mnUtilError As Integer, mnDigitalIOError As Integer
Dim mnConfigError As Integer, mnAnalogIOError As Integer

'Declare Function cbVIn Lib "cbw32.dll" (ByVal BoardNum&, ByVal Chan&, ByVal CBRange&, VoltValue!, ByVal Options&) As Long
'Declare Function cbVOut Lib "cbw32.dll" (ByVal BoardNum&, ByVal Chan&, ByVal CBRange&, VoltValue!, ByVal Options&) As Long


Function GetBoardIndex(BoardNum As Integer) As Integer

   For i% = 0 To gnNumBoards - 1
      If gnBoardEnum(i%) = BoardNum Then Exit For
   Next i%
   GetBoardIndex = i%

End Function

Sub LoadChildForm(Index As Integer)

   Dim FormLoading As Form
   mfmUniTest.MousePointer = vbHourglass
   NewForm% = Forms.Count
   
   
   Select Case Index
      Case ANALOG_IN
         ReDim Preserve frmNewAnalog(gnAIOForms)
         Set frmNewAnalog(gnAIOForms) = New frmAnalog
         Load frmNewAnalog(gnAIOForms)
         TagString$ = Hex$(Index * &H100 + gnAIOForms)
         frmNewAnalog(gnAIOForms).Tag = TagString$
         frmNewAnalog(gnAIOForms).Caption = "Analog Input"
         'frmNewAnalog(gnAIOForms).mnuDataSel.Visible = False
         frmNewAnalog(gnAIOForms).mnuFuncArray(2).Caption = "cbATrig()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(3).Caption = "cbAPreTrig()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(4).Caption = "cbFileAInScan()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(5).Caption = "cbFilePretrig()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(6).Caption = "cbTIn()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(7).Caption = "cbTInScan()"
         frmNewAnalog(gnAIOForms).Show
         Instance% = gnAIOForms
         FunctionIndex% = 1
         Set FormLoading = frmNewAnalog(gnAIOForms)
         gnAIOForms = gnAIOForms + 1
      Case ANALOG_OUT
         ReDim Preserve frmNewAnalog(gnAIOForms)
         Set frmNewAnalog(gnAIOForms) = New frmAnalog
         TagString$ = Hex$(Index * &H100 + gnAIOForms)
         frmNewAnalog(gnAIOForms).Tag = TagString$
         frmNewAnalog(gnAIOForms).Caption = "Analog Output"
         frmNewAnalog(gnAIOForms).mnuFuncArray(0).Caption = "cbAOut()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(1).Caption = "cbAOutScan()"
         frmNewAnalog(gnAIOForms).mnuFuncArray(2).Caption = "cbVOut()"
         For i% = 3 To 8
            frmNewAnalog(gnAIOForms).mnuFuncArray(i%).Visible = False
         Next i%
         frmNewAnalog(gnAIOForms).cmdGo.Visible = False
         frmNewAnalog(gnAIOForms).txtCount.Text = 100
         frmNewAnalog(gnAIOForms).hsbData.Visible = True
         frmNewAnalog(gnAIOForms).txtData.Visible = True
         frmNewAnalog(gnAIOForms).Show
         Instance% = gnAIOForms
         FunctionIndex% = 0
         Set FormLoading = frmNewAnalog(gnAIOForms)
         gnAIOForms = gnAIOForms + 1
      Case DIGITAL_IN
         ReDim Preserve frmNewDigital(gnDIOForms)
         Set frmNewDigital(gnDIOForms) = New frmDigital
         TagString$ = Hex$(Index * &H100 + gnDIOForms)
         frmNewDigital(gnDIOForms).Tag = TagString$
         frmNewDigital(gnDIOForms).Caption = "Digital Input"
         frmNewDigital(gnDIOForms).Show
         Instance% = gnDIOForms
         FunctionIndex% = 5
         Set FormLoading = frmNewDigital(gnDIOForms)
         gnDIOForms = gnDIOForms + 1
      Case DIGITAL_OUT
         ReDim Preserve frmNewDigital(gnDIOForms)
         Set frmNewDigital(gnDIOForms) = New frmDigital
         TagString$ = Hex$(Index * &H100 + gnDIOForms)
         frmNewDigital(gnDIOForms).Tag = TagString$
         frmNewDigital(gnDIOForms).Caption = "Digital Output"
         frmNewDigital(gnDIOForms).Show
         Instance% = gnDIOForms
         FunctionIndex% = 5
         Set FormLoading = frmNewDigital(gnDIOForms)
         gnDIOForms = gnDIOForms + 1
      Case COUNTERS
         ReDim Preserve frmNewCounter(gnCTRForms)
         Set frmNewCounter(gnCTRForms) = New frmCounter
         TagString$ = Hex$(Index * &H100 + gnCTRForms)
         frmNewCounter(gnCTRForms).Tag = TagString$
         frmNewCounter(gnCTRForms).Show
         Instance% = gnCTRForms
         FunctionIndex% = 0
         Set FormLoading = frmNewCounter(gnCTRForms)
         gnCTRForms = gnCTRForms + 1
      Case UTILITIES
         ReDim Preserve frmNewUtil(gnUtilForms)
         Set frmNewUtil(gnUtilForms) = New frmUtilities
         TagString$ = Hex$(Index * &H100 + gnUtilForms)
         frmNewUtil(gnUtilForms).Tag = TagString$
         frmNewUtil(gnUtilForms).Show
         Instance% = gnUtilForms
         FunctionIndex% = 9   'flash LED
         Set FormLoading = frmNewUtil(gnUtilForms)
         gnUtilForms = gnUtilForms + 1
      Case Config
         ReDim Preserve frmNewCfg(gnCfgForms)
         Set frmNewCfg(gnCfgForms) = New frmConfiguration
         TagString$ = Hex$(Index * &H100 + gnCfgForms)
         frmNewCfg(gnCfgForms).Tag = TagString$
         frmNewCfg(gnCfgForms).Show
         Instance% = gnCfgForms
         Set FormLoading = frmNewCfg(gnCfgForms)
         FunctionIndex% = 0
         gnCfgForms = gnCfgForms + 1
      'Case GPIB_CTL
      '   ReDim Preserve frmNewGPIBCtl(gnGPIBCtlForms)
      '   Set frmNewGPIBCtl(gnGPIBCtlForms) = New frmGPIBCtl
      '   TagString$ = Hex$(Index * &H100 + gnGPIBCtlForms)
      '   frmNewGPIBCtl(gnGPIBCtlForms).Tag = TagString$
      '   frmNewGPIBCtl(gnGPIBCtlForms).Show
      '   Instance% = gnGPIBCtlForms
      '   gnGPIBCtlForms = gnGPIBCtlForms + 1
      Case LOGFUNC
         ReDim Preserve frmNewLogger(gnLoggerForms)
         Set frmNewLogger(gnLoggerForms) = New frmLogger
         TagString$ = Hex$(Index * &H100 + gnLoggerForms)
         frmNewLogger(gnLoggerForms).Tag = TagString$
         frmNewLogger(gnLoggerForms).Show
         Instance% = gnLoggerForms
         FunctionIndex% = 14   'Read CSV
         Set FormLoading = frmNewLogger(gnLoggerForms)
         gnLoggerForms = gnLoggerForms + 1
   End Select
   gnInitializing = True
   PositionForm Index, Instance%

   If gnScriptSave Then
      'Print #2, TagString$; ","; 5001; ","; Index; ","; Instance%; ","; ","; ","; ","; ","; ","; ","; ","; ","; ","; ","
      Print #2, TagString$ & ", 5001, " & Format$(Index, "0") & ", " & Format$(Instance%, "0") & ",,,,,,,,,,,"
      BoardZero$ = GetNameOfBoard(gnBoardEnum(0))
      Print #2, TagString$ & ", " & Format$(SSetBoardName, "0") & ", 0, " & BoardZero$ & ", "; A2; ","; A3; ","; A4; ","; A5; ","; A6; ","; A7; ","; A8; ","; A9; ","; A10; ","; A11; ","; AuxHandle
   End If
   
   DoEvents
   If Not FormLoading Is Nothing Then FormLoading.InitForm FunctionIndex%
   mfmUniTest.MousePointer = vbDefault

End Sub

Sub Main()
   
   Dim s As String
   Set FS = CreateObject("Scripting.FileSystemObject")
   s = Space$(255)
   ModHandle& = GetModuleHandle(vbNullString)
   StrLength& = GetModuleFileName(ModHandle&, s, Len(s))
   If StrLength& > 1 Then ModulePath$ = UCase(Left(s, StrLength&))
   gnIDERunning = (InStr(1, ModulePath$, "VB6.EXE") > 0)

   'AppVersion$ = App.Major & "." & App.Minor & "." & App.Revision
   'frmSplash.lblVersion.Caption = "Version   " & AppVersion$
   frmSplash.Show
   'get system timing
   StartTime! = Timer
   Do
      DoEvents
      Iteration& = Iteration& + 1
   Loop While Timer - StartTime! < 0.1
   gnTimeFactor = Iteration& / 2
   SegPath = Split(App.Path, "\", 10)
   DriveName$ = SegPath(0)
   Set drv = FS.GetDrive(DriveName$)
   d& = drv.DriveType
   If d& = 3 Then
      mfmUniTest.Picture1.Visible = True
      mfmUniTest.Picture1.Print "    Do not run Universal Test from the Network."
      mfmUniTest.Picture1.Print " Download it to a local drive and run it from there.."
      mfmUniTest.picCommands.ENABLED = False
      frmMain.mnuScripter.ENABLED = False
   Else
      CmdString$ = Command()
      Quotes& = FindInString(CmdString$, Chr(34), Locs)
      If Not Quotes& < 0 Then CmdString$ = Mid(CmdString$, Locs(0) + 1, Locs(1) - (Locs(0) + 1))
      CheckFile = Split(CmdString$, ".")
      If Not IsEmpty(CheckFile) Then
         EndString& = UBound(CheckFile)
         If EndString& > 0 Then FileExt$ = LCase(CheckFile(EndString&))
      End If
      'MsgBox CmdString$, vbInformation, "No Quotes"
      If Not CmdString$ = "" Then
         Select Case FileExt$
            Case "uts", "utm", "uss"
               frmScript.Show
               gnUniScript = True
               frmMain.mnuScripter.Checked = True
            Case "csv"
               mfmUniTest.Show
               mfmUniTest.cmdFormType(8) = True
               DoEvents
               DirParse = Split(CmdString$, "\")
               If Not IsEmpty(DirParse) Then
                  EndString& = UBound(DirParse)
                  If EndString& > 0 Then
                     Found& = FindInString(CmdString$, "\", Locations)
                     If Not (Found& < 0) Then
                        Loca& = Locations(Found&)
                     End If
                     FP$ = Left$(CmdString$, Loca&)
                     FilePath$ = FP$
                     'If Left(FP$, 1) = Chr(34) Then FilePath$ = Mid(FP$, 2)
                     'MsgBox FilePath$, vbInformation, "Path"
                     frmNewLogger(0).Dir1.Path = FilePath$
                  End If
               End If
               frmNewLogger(0).txtCurFile.Text = CmdString$
               If Not CmdString$ = "" Then frmNewLogger(0).cmdGo = True
            Case Else
               mfmUniTest.Show
         End Select
      Else
         mfmUniTest.Show
      End If
   End If
   Unload frmSplash
   App.HelpFile = App.Path & "\UniScript.chm"
   Seed = Rnd(1)
   Randomize (Seed)

End Sub

Sub PositionForm(FormType As Integer, FormInstance As Integer)

   lpFileName$ = "UniTest.ini"
   lpHKeyName$ = "Height"
   lpWKeyName$ = "Width"
   lpTKeyName$ = "Top"
   nTDefault% = 0
   lpLKeyName$ = "Left"
   nLDefault% = 0
   nHDefault% = 2215
   nWDefault% = 6685
   
   Select Case FormType
      Case MAIN_FORM
         FormName$ = "MainForm"
         nHDefault% = Screen.Height / 3 '1815
         nWDefault% = Screen.Width '8685
      Case ANALOG_IN
         FormName$ = "AnalogIn"
      Case ANALOG_OUT
         FormName$ = "AnalogOut"
      Case DIGITAL_IN
         FormName$ = "DigitalIn"
      Case DIGITAL_OUT
         FormName$ = "DigitalOut"
      Case COUNTERS
         FormName$ = "Counter"
      Case UTILITIES
         FormName$ = "Utilities"
      Case Config
         FormName$ = "Configuration"
      Case LOGFUNC
         FormName$ = "LoggerFuncs"
   End Select
   lpApplicationName$ = FormName$ & FormInstance

   Select Case FormType
      Case MAIN_FORM
         mfmUniTest.Height = 1
         mfmUniTest.Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         mfmUniTest.Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         mfmUniTest.Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         mfmUniTest.Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
      Case ANALOG_IN, ANALOG_OUT
         frmNewAnalog(FormInstance).Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         frmNewAnalog(FormInstance).Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         frmNewAnalog(FormInstance).Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         frmNewAnalog(FormInstance).Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
      Case DIGITAL_IN, DIGITAL_OUT
         frmNewDigital(FormInstance).Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         frmNewDigital(FormInstance).Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         frmNewDigital(FormInstance).Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         frmNewDigital(FormInstance).Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
      Case COUNTERS
         frmNewCounter(FormInstance).Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         frmNewCounter(FormInstance).Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         frmNewCounter(FormInstance).Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         frmNewCounter(FormInstance).Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
      Case UTILITIES
         frmNewUtil(FormInstance).Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         frmNewUtil(FormInstance).Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         frmNewUtil(FormInstance).Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         frmNewUtil(FormInstance).Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
      Case Config
         frmNewCfg(FormInstance).Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         frmNewCfg(FormInstance).Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         frmNewCfg(FormInstance).Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         frmNewCfg(FormInstance).Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
      Case LOGFUNC
         frmNewLogger(FormInstance).Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
         frmNewLogger(FormInstance).Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
         frmNewLogger(FormInstance).Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
         frmNewLogger(FormInstance).Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
   End Select

End Sub

Sub PositionMain()

   frmMain.Width = mfmUniTest.ScaleWidth
   frmMain.Height = 10
   frmMain.Top = 0

End Sub

Sub PrintMain(MainString As String)

   If Len(MainString) Then smMainString = MainString
   mfmUniTest.picCommands.Cls
   mfmUniTest.picCommands.CurrentX = 20
   mfmUniTest.picCommands.CurrentY = mfmUniTest.picCommands.ScaleHeight / 20
   mfmUniTest.picCommands.Print smMainString

End Sub

Sub UnLoadChildForm(FormUnLoading As Form, FormType As Integer, Instance As Integer)

   If FormUnLoading.WindowState = 0 Then
      lpFileName$ = "UniTest.ini"
      Select Case FormType
         Case MAIN_FORM
            FormName$ = "MainForm"
         Case ANALOG_IN
            FormName$ = "AnalogIn"
         Case ANALOG_OUT
            FormName$ = "AnalogOut"
         Case DIGITAL_IN
            FormName$ = "DigitalIn"
         Case DIGITAL_OUT
            FormName$ = "DigitalOut"
         Case COUNTERS
            FormName$ = "Counter"
         Case UTILITIES
            FormName$ = "Utilities"
         Case Config
            FormName$ = "Configuration"
      End Select
      lpApplicationName$ = FormName$ & Instance
      If FormUnLoading.WindowState = 0 Then
         lpKeyName$ = "Height"
         lpString$ = Str$(FormUnLoading.Height)
         x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
         lpKeyName$ = "Width"
         lpString$ = Str$(FormUnLoading.Width)
         x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
         lpKeyName$ = "Top"
         lpString$ = Str$(FormUnLoading.Top)
         x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
         lpKeyName$ = "Left"
         lpString$ = Str$(FormUnLoading.Left)
         x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      End If
      TagString$ = FormUnLoading.Tag
      If gnScriptSave Then
         Print #2, TagString$ & ", 5002, " & Format$(FormType, "0") & ", " & Format$(Instance%, "0") & ", , , , , , , , , , , """
      End If
   End If

End Sub

Function GetFormsOfType(FormID As String, MatchingForms As Variant) As Integer

   Dim FormIndexs() As Integer
   If IsNumeric(FormID) Then
      FormName$ = Left(FormID, 2)
   Else
      FormName$ = FormID
   End If
   IDLength& = Len(FormName$)
   For i% = 0 To Forms.Count - 1
      FormTag$ = Forms(i%).Tag
      If Not ((FormTag$ = "UniTest") Or (FormTag$ = "Main")) Then
         If Len(FormTag$) > IDLength& Then FormTag$ = Left(FormTag$, IDLength&)
         If (FormTag$ = FormName$) Or (FormName$ = "-1") Then
            ReDim Preserve FormIndexs(FormsFound&)
            FormIndexs(FormsFound&) = i%
            FormsFound& = FormsFound& + 1
         End If
      End If
   Next
   MatchingForms = FormIndexs()
   GetFormsOfType = FormsFound& - 1
   
End Function

Function GetFormReference(FormID As String, FormRef As Form) As Integer

   'check for valid form
   For i% = 0 To Forms.Count - 1
      FormTag$ = Forms(i%).Tag
      If FormTag$ = FormID Then FormCheck% = True
   Next
   If Len(FormID) > 2 Then Instance% = Val(Right$(FormID, 2))
   FormType% = Val(Left$(FormID, 1))
   
   Select Case FormType%
      Case MAIN_FORM
         ReturnVal% = False
      Case ANALOG_IN
         If FormCheck% Then NumForms% = UBound(frmNewAnalog)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Analog Input"
            If Not mnAnalogInError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnAnalogInError = True
         Else
            Set FormRef = frmNewAnalog(Instance%)
         End If
      Case ANALOG_OUT
         If FormCheck% Then NumForms% = UBound(frmNewAnalog)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Analog Output"
            If Not mnAnalogOutError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnAnalogOutError = True
         Else
            Set FormRef = frmNewAnalog(Instance%)
         End If
      Case DIGITAL_IN
         If FormCheck% Then NumForms% = UBound(frmNewDigital)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Digital Input"
            If Not mnDigitalInError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnDigitalInError = True
         Else
            Set FormRef = frmNewDigital(Instance%)
         End If
      Case DIGITAL_OUT
         If FormCheck% Then NumForms% = UBound(frmNewDigital)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Digital Output"
            If Not mnDigitalOutError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnDigitalOutError = True
         Else
            Set FormRef = frmNewDigital(Instance%)
         End If
      Case COUNTERS
         If FormCheck% Then NumForms% = UBound(frmNewCounter)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Counter"
            If Not mnCounterError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnCounterError = True
         Else
            Set FormRef = frmNewCounter(Instance%)
         End If
      Case UTILITIES
         If FormCheck% Then NumForms% = UBound(frmNewUtil)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Counter"
            If Not mnUtilError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnUtilError = True
         Else
            Set FormRef = frmNewUtil(Instance%)
         End If
      Case Config
         If FormCheck% Then NumForms% = UBound(frmNewCfg)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "Configuration"
            If Not mnConfigError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnConfigError = True
         Else
            Set FormRef = frmNewCfg(Instance%)
         End If
      Case GPIB_CTL
         If FormCheck% Then NumForms% = UBound(frmNewGPIBCtl)
         If (NumForms% < Instance%) Or Not FormCheck% Then
            FormTypeName$ = "GPIB"
            If Not mnGPIBError Then BadInstance% = _
            FormNotOpen(FormTypeName$, FormID$)
            mnGPIBError = True
         Else
            Set FormRef = frmNewGPIBCtl(Instance%)
         End If
      Case LOGFUNC
         Set FormRef = frmNewLogger(Instance%)
      Case Else
         ReturnVal% = False
   End Select
   If Not BadInstance% Then
      If Not (FormRef Is Nothing) Then
         ReturnVal% = True
      Else
         'MsgBox "Could not get reference to form specified in script. Script value " & _
         "for form: " & FormID, vbCritical, "Form Reference Error in Script"
      End If
   End If
   GetFormReference = ReturnVal%
   
End Function

Function GetFormTypeName(FormID As String) As String

   FormType% = Val(Left$(FormID, 1))
   Select Case FormType%
      Case MAIN_FORM
         ReturnVal$ = "Main Form"
      Case ANALOG_IN
         ReturnVal$ = "Analog Input"
      Case ANALOG_OUT
         ReturnVal$ = "Analog Output"
      Case DIGITAL_IN
         ReturnVal$ = "Digital Input"
      Case DIGITAL_OUT
         ReturnVal$ = "Digital Output"
      Case COUNTERS
         ReturnVal$ = "Counter"
      Case UTILITIES
         ReturnVal$ = "Counter"
      Case Config
         ReturnVal$ = "Configuration"
      Case GPIB_CTL
         ReturnVal$ = "GPIB"
      Case LOGFUNC
         ReturnVal$ = "Data Logger"
      Case Else
         ReturnVal$ = "Undefined"
   End Select
   GetFormTypeName = ReturnVal$

End Function

Function FormNotOpen(FormTypeName As String, FormID As String) As Integer

   Dim response As VbMsgBoxResult
   
   response = MsgBox("This script is calling a " & FormTypeName & " form" & _
   " (" & FormID & ") " & "that isn't open. It may not work properly. " & _
   "Abort?", vbYesNo, "Missing Form")
   If response = vbYes Then
      gnScriptRun = False
      FormNotOpen = True
   End If

End Function

Public Sub ConfigureLibrary(LibType As Integer, Optional Initialize As Integer = False)

   Select Case LibType
      Case INVALIDLIB
         mfmUniTest.Caption = "Universal Test: Windows (No library selected)"
      Case UNILIB
         mfmUniTest.Caption = "Universal Test: Universal Library for Windows"
         If gbULLoaded Then
            gnLibType = UNILIB
            'LibRevString$ = GetLibDLLRev("cbw32.dll")
            'gfDLLRev = Val(LibRevString$)
         End If
         If Initialize Then gnNumBoards = GetNumInstalled()
      Case NETLIB
         mfmUniTest.Caption = "Universal Test: Universal Library for .Net"
         If Initialize And Not gnICalDisable Then gnNumBoards = GetNumInstalled()
      Case MSGLIB
         mfmUniTest.Caption = "Universal Test: DAQFlex for Windows"
         If Initialize Then gnNumBoards = GetNumMsgBoards()
   End Select
   mfmUniTest.picCommands.ENABLED = (gnLibType > INVALIDLIB)
   frmMain.mnuScripter.ENABLED = (gnLibType > INVALIDLIB)

End Sub

Public Sub RemoveDiscoveryForm()

   Unload frmDiscovery
   
End Sub

Public Function GetDiscoverOption() As Boolean

   Dim RemoveUndiscovered As Boolean
   
   RemoveUndiscovered = (frmDiscovery.chkRemoveUndisc.value = 1)
   GetDiscoverOption = RemoveUndiscovered
   
End Function
