VERSION 5.00
Begin VB.Form frmLogger 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Logger Functions"
   ClientHeight    =   6480
   ClientLeft      =   3480
   ClientTop       =   1710
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
   ScaleHeight     =   6480
   ScaleWidth      =   6045
   Tag             =   "900"
   Visible         =   0   'False
   Begin VB.Frame fraStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   1920
      Width           =   5775
      Begin VB.Label lblFileDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label lblFileSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2820
         TabIndex        =   17
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   120
         Width           =   2715
      End
   End
   Begin VB.Frame fraConfiguration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set Preferences"
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3100
      Begin VB.OptionButton optUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "°K"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   780
         Width           =   600
      End
      Begin VB.OptionButton optUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "°F"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   540
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "°C"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   300
         Width           =   600
      End
      Begin VB.CheckBox chkPrefs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "24 Hr"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   840
      End
      Begin VB.CheckBox chkPrefs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "GMT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   615
         Width           =   840
      End
   End
   Begin VB.ComboBox cmbFileNum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Text            =   "FileNumber"
      Top             =   1320
      Width           =   1995
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   2220
      MousePointer    =   9  'Size W E
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3720
      Width           =   3675
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtCurFile 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   2220
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1995
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdConfigure 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
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
   Begin VB.Menu mnuFunc 
      Caption         =   "F&unction"
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbGetFileName()"
         Index           =   0
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetFileInfo()"
         Index           =   1
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetSampleInfo()"
         Index           =   2
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetAIChannelCount()"
         Index           =   3
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetAIInfo()"
         Index           =   4
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetCJCInfo()"
         Index           =   5
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetDIOInfo()"
         Index           =   6
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogReadTimeTags()"
         Index           =   7
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogReadAIChannels()"
         Index           =   8
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogReadCJCChannels()"
         Index           =   9
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogReadDIOChannels()"
         Index           =   10
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogConvertFile()"
         Index           =   11
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogSetPreferences()"
         Index           =   12
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "cbLogGetPreferences()"
         Index           =   13
      End
      Begin VB.Menu mnuFuncArray 
         Caption         =   "Read CSV"
         Index           =   14
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadConfig 
         Caption         =   "cbLoadConfig()"
      End
   End
   Begin VB.Menu mnuPlot 
      Caption         =   "Plot Type"
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
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the following two constants added here until
'they become part of the header
'Const BIADTRIGCOUNT = 219
'Const BIADFIFOSIZE = 220

Const GETFILENAME = 0
Const GETFILEINFO = 1
Const GETSAMPLEINFO = 2
Const GETAICHANS = 3
Const GETAIINFO = 4
Const GETCJCINFO = 5
Const GETDIOINFO = 6
Const READTIMETAGS = 7
Const READAICHANS = 8
Const READCJCCHANS = 9
Const READDIOCHANS = 10
Const CONVERTFILE = 11
Const SETPREFERENCE = 12
Const GETPREFERENCE = 13
Const READ_CSV = 14

Private Type ChanConfig
   ChanType As Integer
   ChanNum As Integer
   ChanConfig As String
   ChanRange As Long
   ChanUnits As String
End Type

Dim ChanDefs() As ChanConfig

Dim mnThisInstance As Integer

Dim mnFormType As Integer, msTitle As String

Dim mnFuncType As Integer, mnPlotType As Integer

Dim mnPlot As Integer, mnLoaded As Integer

Dim gsConfig As String, msOpt As String

Dim mnNumBoards As Integer, mnBoardIndex As Integer
Dim mnBoardNum As Integer

Dim mnADChans As Integer, mnDigDevs As Integer
Dim mnCtrDevs As Integer, mnDAChans As Integer
Dim mnIOPorts As Integer
Dim mlDirection As Long, mlPolarity As Long
Dim mnCurSelect As Integer
Dim mnWide As Integer

Dim mlTimeFormat As Long
Dim mlTimeZone As Long
Dim mlUnits As Long
Dim msDestFile As String
Dim mlSampleCount As Long
Dim mlAIChans As Long
Dim mlCJCChans As Long
Dim mlDIOChans As Long
Dim malTimes() As Long
Dim malDates() As Long
Dim mafAIChans() As Single
Dim mafCJCChans() As Single
Dim malDIOChans()  As Long
Dim malChanNums() As Long
Dim malUnits() As Long

Dim mlQuadChans As Long
Dim msDisplayName As String, mnResolution As Integer
Dim mfRateReturned As Single, mfRate As Single
Dim mlHandle As Long, mlDataType As Long, mlCount As Long
Dim manGainArray() As Integer, manChanArray() As Integer, mlQCount As Long
Dim manTypeArray() As Integer
Dim mnLibType As Integer, mlBlockSave As Long
Dim mlTotalCount As Long
Dim mlPreTrigCount As Long, mlPTCountReturn As Long
Dim mnTimeStamped As Integer
Dim mlDataLine As Long, mlStaticOpt As Long
Dim mnAcqDataType As VbVarType, mnGenDataType As VbVarType
Dim mlAcqPoints As Long, mnEngApp As Integer

Private Sub chkPrefs_Click(Index As Integer)

   mlTimeFormat = Choose(chkPrefs(0).value + 1, TIMEFORMAT_12HOUR, TIMEFORMAT_24HOUR)
   mlTimeZone = Choose(chkPrefs(1).value + 1, TIMEZONE_LOCAL, TIMEZONE_GMT)

End Sub

Private Sub cmdConfigure_Click()

   'this exists to give menu access to the scripting
   'form when running scripts
   CmdStr$ = Left$(cmdConfigure.Caption, 1)
   value& = Val(Mid$(cmdConfigure.Caption, 2))
   Select Case CmdStr$
      Case "$" 'return board name
         cmdConfigure.Caption = msDisplayName
      Case "1" 'return handle to data
         Me.cmdConfigure.Caption = Format(mlHandle, "0")
      Case "2" 'return data resolution
         Me.cmdConfigure.Caption = Format(mnResolution, "0")
      Case "3" 'return range at which data was collected
         'If Not IsNull(mvCustomRange) Then
         '   Me.cmdConfigure.Caption = "C," & _
         '   Format(mnRange, "0") & "," & mvCustomRange
         '   Exit Sub
         'End If
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
               If Not (mlAIChans < 0) Then
                  'MSG boards may have queue or individual chans
                  If (mlQCount > 0) Then cmdConfigure.Caption = "Q"
               Else
                  Range% = mnFixedRange
                  Me.cmdConfigure.Caption = Format(Range%, "0")
               End If
            End Select
      Case "4" 'return rate requested
         Me.cmdConfigure.Caption = Format(mfRate, "0.0####")
      Case "5" 'return rate returned
         Me.cmdConfigure.Caption = Format(mfRateReturned, "0.0####")
      Case "7"
         Me.cmdConfigure.Caption = Format(mlPTCountReturn, "0")
      Case "8"
         CurOption$ = msOpt
         If msOpt = "Options = Default  " Then CurOption$ = ""
         Me.cmdConfigure.Caption = CurOption$
      Case "="
         TotalCount& = mlCount
         'If mlAcqPoints > mlCount Then TotalCount& = mlAcqPoints
         'If Not mlTotalCountReturn = 0 Then
         '   If mlTotalCountReturn < mlCount Then TotalCount& = mlTotalCountReturn
         'End If
         cmdConfigure.Caption = Format(TotalCount&, "0")
      Case "?" 'return number of channels
         If (mlQCount > 0) Then QueueSet% = True
         NumberOfChans% = (mnLastChan - mnFirstChan) + 1
         If QueueSet% Then cmdConfigure.Caption = "Q"
         'Me.cmdConfigure.Caption = NumChans%
      Case "F" 'set function
         mnuFuncArray_Click (value&)
      Case "X" 'execute the function
         cmdGo = True
   End Select

End Sub

Private Sub cmdGo_Click()

   On Error Resume Next
   Filename$ = txtCurFile.Text
   mlQCount = 0
   Select Case mnFuncType
      Case GETFILENAME
         DataFile$ = Space$(100)
         Path$ = Me.Dir1.Path & Chr$(0)
         FileNum& = cmbFileNum.ListIndex - 3 'GETFIRST
         ULStat = cbLogGetFileName(FileNum&, Path$, DataFile$)
         DataFile$ = TrimStrings(DataFile$)
         Path$ = TrimStrings(Path$)
         x& = SaveFunc(Me, LogGetFileName, ULStat, FileNum&, Path$, DataFile$, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = txtData.Text & DataFile$ & Chr$(13) & Chr$(10)
            txtData.SelStart = Len(txtData.Text)
         End If
      Case GETFILEINFO
         ULStat = cbLogGetFileInfo(Filename$, Vers&, Size&)
         x& = SaveFunc(Me, LogGetFileInfo, ULStat, Filename$, Vers&, Size&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = "Version: " & Vers& & Chr$(13) & Chr$(10)
            txtData.Text = txtData.Text & "Size: " & Size& & Chr$(13) & Chr$(10)
            txtData.SelStart = Len(txtData.Text)
         End If
      Case GETSAMPLEINFO
         txtData.Left = 120
         txtData.Width = Me.Width - 280
         mnWide = True
         ULStat = cbLogGetSampleInfo(Filename$, SampleInterval&, mlSampleCount, StartDate&, StartTime&)
         x& = SaveFunc(Me, LogGetSampleInfo, ULStat, Filename$, SampleInterval&, mlSampleCount, StartDate&, StartTime&, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = "Sample Interval: " & SampleInterval& & Chr$(13) & Chr$(10)
            txtData.Text = txtData.Text & "Sample Count: " & mlSampleCount & Chr$(13) & Chr$(10)
            txtData.Text = txtData.Text & "Start Date: " & StartDate&
            txtData.Text = txtData.Text & "  (" & ConvertDate(StartDate&) & ")" & Chr$(13) & Chr$(10)
            txtData.Text = txtData.Text & "Start Time: " & StartTime&
            txtData.Text = txtData.Text & "  (" & ConvertTime(StartTime&) & ")" & Chr$(13) & Chr$(10)
            txtData.SelStart = Len(txtData.Text)
         End If
      Case GETAICHANS
         ULStat = cbLogGetAIChannelCount(Filename$, mlAIChans)
         x& = SaveFunc(Me, LogGetAIChannelCount, ULStat, Filename$, mlAIChans, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = "AI Channel Count: " & mlAIChans & Chr$(13) & Chr$(10)
            txtData.SelStart = Len(txtData.Text)
         End If
      Case GETAIINFO
         txtData.Left = 120
         txtData.Width = Me.Width - 280
         mnWide = True
         TempChans& = mlAIChans
         ReDim malChanNums(TempChans&)
         ReDim malUnits(TempChans&)
         'ULStat = cbLogGetAIInfo(FileName$, ChannelMask&, UnitMask&, mlAIChans)
         ULStat = cbLogGetAIInfo(Filename$, malChanNums(0), malUnits(0))
         x& = SaveFunc(Me, LogGetAIInfo, ULStat, Filename$, malChanNums(0), malUnits(0), A4, A5, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = ""
            For Ch& = 0 To TempChans& - 1
               txtData.Text = txtData.Text & "Channel " & Format$(malChanNums(Ch&), "0") & " units: " & Choose(malUnits(Ch&) + 1, "Temperature", "Raw") & Chr$(13) & Chr$(10)
            Next
            txtData.SelStart = Len(txtData.Text)
         End If
      Case GETCJCINFO
         ULStat = cbLogGetCJCInfo(Filename$, mlCJCChans)
         x& = SaveFunc(Me, LogGetCJCInfo, ULStat, Filename$, mlCJCChans, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = "CJC Chan Count: " & mlCJCChans & Chr$(13) & Chr$(10)
            txtData.SelStart = Len(txtData.Text)
         End If
      Case GETDIOINFO
         ULStat = cbLogGetDIOInfo(Filename$, mlDIOChans)
         x& = SaveFunc(Me, LogGetDIOInfo, ULStat, Filename$, mlDIOChans, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         If ULStat = 0 Then
            txtData.Text = "DIO Chan Count: " & mlDIOChans & Chr$(13) & Chr$(10)
            txtData.SelStart = Len(txtData.Text)
         End If
      Case READTIMETAGS
         ActionCancelled% = GetSampleRange(SStart&, SCount&)
         If Not ActionCancelled% Then
            ReDim malDates(SCount&)
            ReDim malTimes(SCount&)
            txtData.Left = 120
            txtData.Width = Me.Width - 280
            mnWide = True
            ULStat = cbLogReadTimeTags(Filename$, SStart&, SCount&, malDates(0), malTimes(0))
            x& = SaveFunc(Me, LogReadTimeTags, ULStat, Filename$, SStart&, SCount&, malDates(0), malTimes(0), A6, A7, A8, A9, A10, A11, 0)
            If ULStat = 0 Then
               txtData.Text = "StartSample: " & SStart& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "SampleCount: " & SCount& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "Dates: " & malDates(0)
               txtData.Text = txtData.Text & "  (" & ConvertDate(malDates(0)) & ")" & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "Times: " & malTimes(0)
               txtData.Text = txtData.Text & "  (" & ConvertTime(malTimes(0)) & ")" & Chr$(13) & Chr$(10)
               txtData.SelStart = Len(txtData.Text)
               SetPlotType PRINT_LIST, Me
               frmPlot.txtShow.Text = ""
               LineFeed$ = Chr$(13) + Chr$(10)
               ListSize& = 500
               TotalSamps& = SCount& - 1
               If TotalSamps& < ListSize& Then ListSize& = TotalSamps&
               For Samp& = 0 To ListSize&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(Samp&, "0)") & Chr$(9) & ConvertDate(malDates(Samp&)) & Chr$(9) & ConvertTime(malTimes(Samp&)) & Chr$(13) & Chr$(10)
               Next Samp&
            End If
         End If
      Case READAICHANS
         If (mlAIChans = 0) Or (mlSampleCount = 0) Then
            MsgBox "Run cbLogGetSampleInfo and cbLogGetAIInfo before running this function.", , "No Size Data"
            Exit Sub
         End If
         ActionCancelled% = GetSampleRange(SStart&, SCount&)
         If Not ActionCancelled% Then
            ReDim mafAIChans(mlAIChans - 1, SCount& - 1)
            ULStat = cbLogReadAIChannels(Filename$, SStart&, SCount&, mafAIChans(0, 0))
            x& = SaveFunc(Me, LogReadAIChannels, ULStat, Filename$, SStart&, SCount&, mafAIChans(0, 0), A5, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = 0 Then
               txtData.Left = 120
               txtData.Width = Me.Width - 280
               mnWide = True
               txtData.Text = "Start Sample: " & SStart& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "Sample Count: " & SCount& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "AI Channel Data: " & mafAIChans(0, 0) & Chr$(13) & Chr$(10)
               txtData.SelStart = Len(txtData.Text)
               SetPlotType PRINT_LIST, Me
               frmPlot.txtShow.Text = ""
               LineFeed$ = Chr$(13) + Chr$(10)
               ListSize& = 500
               TotalSamps& = SCount& - 1
               If TotalSamps& < ListSize& Then ListSize& = TotalSamps&
               For Samp& = 0 To ListSize&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(Samp&, "0)") & Chr$(9)
                  For Chan& = 0 To mlAIChans - 1
                     frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(mafAIChans(Chan&, Samp&), "0.00000") & Chr$(9)
                  Next Chan&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & LineFeed$
               Next Samp&
            End If
         End If
      Case READCJCCHANS
         If mlCJCChans = 0 Then
            MsgBox "Either no CJC Chans were logged or you need to run cbLogGetCJCInfo before running this function.", , "No Size Data"
            Exit Sub
         End If
         ActionCancelled% = GetSampleRange(SStart&, SCount&)
         If Not ActionCancelled% Then
            ReDim mafCJCChans(mlCJCChans - 1, SCount& - 1)
            ULStat = cbLogReadCJCChannels(Filename$, SStart&, SCount&, mafCJCChans(0, 0))
            x& = SaveFunc(Me, LogReadCJCChannels, ULStat, Filename$, SStart&, SCount&, mafCJCChans(0, 0), A5, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = 0 Then
               txtData.Left = 120
               txtData.Width = Me.Width - 280
               mnWide = True
               txtData.Text = "Start Sample: " & SStart& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "Sample Count: " & SCount& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "CJC Channel Data: " & mafCJCChans(0, 0) & Chr$(13) & Chr$(10)
               txtData.SelStart = Len(txtData.Text)
               SetPlotType PRINT_LIST, Me
               frmPlot.txtShow.Text = ""
               LineFeed$ = Chr$(13) + Chr$(10)
               ListSize& = 500
               TotalSamps& = SCount& - 1
               If TotalSamps& < ListSize& Then ListSize& = TotalSamps&
               For Samp& = 0 To ListSize&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(Samp&, "0)") & Chr$(9)
                  For Chan& = 0 To mlCJCChans - 1
                     frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(mafCJCChans(Chan&, Samp&), "0.00000") & Chr$(9)
                  Next Chan&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & LineFeed$
               Next Samp&
            End If
         End If
      Case READDIOCHANS
         If (mlDIOChans = 0) Or (mlSampleCount = 0) Then
            MsgBox "Run cbLogGetSampleInfo and cbLogGetDIOInfo before running this function.", , "No Size Data"
            Exit Sub
         End If
         ActionCancelled% = GetSampleRange(SStart&, SCount&)
         If Not ActionCancelled% Then
            ReDim malDIOChans(mlDIOChans - 1, SCount& - 1)
            ULStat = cbLogReadDIOChannels(Filename$, SStart&, SCount&, malDIOChans(0, 0))
            x& = SaveFunc(Me, LogReadDIOChannels, ULStat, Filename$, SStart&, SCount&, malDIOChans(0, 0), A5, A6, A7, A8, A9, A10, A11, 0)
            If ULStat = 0 Then
               txtData.Text = "Start Sample: " & SStart& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "Sample Count: " & SCount& & Chr$(13) & Chr$(10)
               txtData.Text = txtData.Text & "DIO 0 Data: " & malDIOChans(0, 0) & Chr$(13) & Chr$(10)
               txtData.SelStart = Len(txtData.Text)
               SetPlotType PRINT_LIST, Me
               frmPlot.txtShow.Text = ""
               LineFeed$ = Chr$(13) + Chr$(10)
               ListSize& = 500
               TotalSamps& = SCount& - 1
               If TotalSamps& < ListSize& Then ListSize& = TotalSamps&
               For Samp& = 0 To ListSize&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(Samp&, "0)") & Chr$(9)
                  For Chan& = 0 To mlDIOChans - 1
                     frmPlot.txtShow.Text = frmPlot.txtShow.Text & Format$(malDIOChans(Chan&, Samp&), "0") & Chr$(9)
                  Next Chan&
                  frmPlot.txtShow.Text = frmPlot.txtShow.Text & LineFeed$
               Next Samp&
            End If
         End If
      Case CONVERTFILE
         ActionCancelled% = GetDestFile(FileType&, StartVal&, SampleCount&, Delimiter&)
         If Not ActionCancelled% Then
            ULStat = cbLogConvertFile(Filename$, msDestFile, StartVal&, SampleCount&, Delimiter&)
            x& = SaveFunc(Me, LogConvertFile, ULStat, Filename$, msDestFile, FileType&, StartVal&, SampleCount&, Delimiter&, A7, A8, A9, A10, A11, 0)
         End If
      Case SETPREFERENCE, GETPREFERENCE
         ULStat = cbLogSetPreferences(mlTimeFormat, mlTimeZone, mlUnits)
         x& = SaveFunc(Me, LogSetPreferences, ULStat, mlTimeFormat, mlTimeZone, mlUnits, A4, A5, A6, A7, A8, A9, A10, A11, 0)
         ShowPreferences
      Case READ_CSV
         mlHandle = ReadCSVtoMem(Filename$, mlDataType, mlCount, ADChans%)
         If Not mlHandle = 0 Then DisplayData
   End Select

End Sub

Private Sub ConfigureControls()
   
   If Not mnWide Then txtData.Text = ""
   mnWide = False
   Select Case mnFuncType
      Case GETFILENAME
         File1.Visible = False
         Dir1.Visible = True
         Drive1.Visible = True
         cmbFileNum.Visible = True
         txtCurFile.Visible = False
         txtData.Left = 2220
         txtData.Width = 3675
         fraConfiguration.Visible = False
         txtData.Visible = True
      Case SETPREFERENCE, GETPREFERENCE
         fraConfiguration.Visible = True
         File1.Visible = False
         Dir1.Visible = False
         Drive1.Visible = False
         cmbFileNum.Visible = False
         txtCurFile.Visible = False
         txtData.Left = 3340
         txtData.Width = 2540
         txtData.Visible = True
         ShowPreferences
      Case Else
         File1.Visible = True
         Dir1.Visible = True
         Drive1.Visible = True
         cmbFileNum.Visible = False
         txtCurFile.Visible = True
         txtData.Left = 4080
         txtData.Width = 1815
         fraConfiguration.Visible = False
         txtData.Visible = True
         If mnFuncType = READ_CSV Then
            File1.Pattern = "*.csv"
         Else
            File1.Pattern = "*.*"
         End If
   End Select

End Sub

Private Function ConvertDate(DateCode As Long) As String

   MyDay = (DateCode And 255)
   MyMonth = (DateCode / 256) And 255
   MyYear = (DateCode / 65536) And 65535
   StartDateStr$ = Format$(MyMonth, "00/") & Format$(MyDay, "00/") & Format$(MyYear, "0000")
   ConvertDate = StartDateStr$

End Function

Private Function ConvertTime(TimeCode As Long) As String

   MyHours = (TimeCode / 65536) And 255
   MyMinutes = (TimeCode / 256) And 255
   MySeconds = (TimeCode And 255)

   'Suffix$ = (TimeCode / 16777216) And 255
   TimeVal = (TimeCode / 16777216) And 255
   If TimeVal = 0 Then
      Suffix$ = " AM"
   ElseIf TimeVal = 1 Then
      Suffix$ = " PM"
   ElseIf TimeVal = -1 Then
      Suffix$ = ""
   Else
      Suffix$ = ""
   End If
  
   StartTimeStr$ = Format$(MyHours, "0:") & Format$(MyMinutes, "00:") & Format$(MySeconds, "00") & Suffix$
   ConvertTime = StartTimeStr$

End Function

Private Sub Dir1_Change()
   
   File1.Path = Dir1.Path   ' Set file path.

End Sub

Private Sub Dir1_KeyUp(KeyCode As Integer, Shift As Integer)
   
   File1.Path = Dir1.Path   ' Set file path.

End Sub

Private Sub Drive1_Change()
On Error GoTo NoDrive

   Dir1.Path = Drive1.Drive ' Set directory path.

   Exit Sub

NoDrive:
   MsgBox Error$(Err), , "Error Selecting Drive"
   Resume Next

End Sub

Private Sub File1_Click()
   
   
   FilePath$ = Dir1.Path
   'txtCurFile.Text = FilePath$
   If Not (Len(FilePath$) = 0) Then
      If Not (Right$(FilePath$, 1) = "\") Then FilePath$ = Dir1.Path & "\"
      Filename$ = FilePath$ & File1.Filename
      If txtCurFile.Text = Filename$ Then Exit Sub
      txtCurFile.Text = Filename$
      If mnFuncType = GETSAMPLEINFO Then
         mlSampleCount = 0
         mlAIChans = 0
         mlCJCChans = 0
         mlDIOChans = 0
         ULStat = cbLogGetFileInfo(Filename$, Vers&, Size&)
         If ULStat = 0 Then
            lblFileSize.Caption = "Size " & Size&
         Else
            lblFileSize.Caption = ""
         End If
         ULStat = cbLogGetSampleInfo(Filename$, SampleInterval&, mlSampleCount, StartDate&, StartTime&)
         If ULStat = 0 Then
            lblFileDate.Caption = ConvertDate(StartDate&)
         Else
            lblFileDate.Caption = ""
         End If
      End If
   Else
      MsgBox "No file path", , "No Path"
   End If

End Sub

Private Sub File1_DblClick()

   Me.cmdGo = True
   
End Sub

Private Sub File1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

   If Data.Files.Count > 0 Then
      FilePath$ = Data.Files(1)
      NumSubs& = FindInString(FilePath$, "\", Locations)
      Dir1.Path = Left(FilePath$, Locations(NumSubs&))
      File1.Refresh
      txtCurFile.Text = FilePath$
      Me.cmdGo = True
   End If
   
End Sub

Private Sub Form_Activate()

   UpdateMainStatus

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
            Me.Height = 2700
            Me.Width = 6200
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
   
   'PositionForm LOGFUNC, Instance%
   fraConfiguration.Top = 60
   txtData.Top = 60

   mlUnits = FAHRENHEIT

   cmbFileNum.AddItem "GETNEXT"
   cmbFileNum.AddItem "GETFIRST"
   cmbFileNum.AddItem "FROMHERE"
   cmbFileNum.AddItem "0"
   cmbFileNum.AddItem "1"
   cmbFileNum.AddItem "2"
   cmbFileNum.AddItem "3"
   cmbFileNum.AddItem "4"
   cmbFileNum.AddItem "5"
   cmbFileNum.AddItem "6"
   cmbFileNum.AddItem "7"
   cmbFileNum.AddItem "8"
   cmbFileNum.AddItem "9"
   cmbFileNum.AddItem "10"
   cmbFileNum.AddItem "11"
   cmbFileNum.AddItem "12"
   cmbFileNum.AddItem "13"
   cmbFileNum.AddItem "14"
   cmbFileNum.AddItem "15"
   cmbFileNum.ListIndex = 1
   mlAIChans = -1: mlCJCChans = -1: mlDIOChans = -1: mlQuadChans = -1

   mnuFuncArray_Click 1

   UpdateStatus

End Sub

Private Sub Form_Resize()

   fraStatus.Width = ScaleWidth
   fraStatus.Top = ScaleHeight - fraStatus.Height
   lblStatus.Width = fraStatus.Width * 0.47
   lblFileSize.Left = lblStatus.Width + 20
   lblFileSize.Width = fraStatus.Width * 0.304
   lblFileDate.Left = lblFileSize.Left + lblFileSize.Width + 20
   lblFileDate.Width = fraStatus.Width * 0.179
   If gnInitializing Then
      'if the form is just loading, sets the form type
      'and sets the default function (cbC8254Config())
      mnFormType = (Val("&H" & Tag) And &HF00&) / &H100
      mnThisInstance = 0
      'mnuFuncArray_Click (GET_CONF)
      msTitle = Caption
      mnuPlotType_Click (0)
      'mnuBoard_Click (0)
      gnInitializing = False
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   'UnLoadChildForm Me, mnFormType, mnThisInstance
   'gnCFGForms = gnCFGForms - 1

   If Me.WindowState = 0 Then
      lpFileName$ = "UniTest.ini"
      FormName$ = "LoggerFuncs"
      lpApplicationName$ = FormName$ & "0"
      lpKeyName$ = "Height"
      lpString$ = Str$(Me.Height)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      lpKeyName$ = "Width"
      lpString$ = Str$(Me.Width)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      lpKeyName$ = "Top"
      lpString$ = Str$(Me.Top)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      lpKeyName$ = "Left"
      lpString$ = Str$(Me.Left)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      TagString$ = Me.Tag
      If gnScriptSave Then
         'Print #2, TagString$; ","; 5002; ","; FormType; ","; 0; ","; ","; ","; ","; ","; ","; ","; ","; ","; ","; ","
         Print #2, TagString$ & ", 5002, " & Format$(FormType, "0") & ", 0"; ","; ","; ","; ","; ","; ","; ","; ","; ","; ","; ","
      End If
   End If

End Sub

Private Function GetDestFile(FileType&, StartVal&, SampleCount&, Delimiter&) As Integer
   
   GetDestFile = False
   frmConvertFile.txtCount = mlSampleCount
   If msDestFile <> "" Then
      'find the file name
      NameLen& = Len(msDestFile)
      For Posit& = NameLen& To 1 Step -1
         If Mid$(msDestFile, Posit&, 1) = "\" Then Exit For
      Next
      MyPath$ = Left$(msDestFile, Posit&)
      frmConvertFile.Dir1.Path = MyPath$
      frmConvertFile.txtDestFile.Text = msDestFile
   End If
   frmConvertFile.Show 1
   'check if form is loaded
   For ThisForm& = 0 To Forms.Count - 1
      'If Forms(ThisForm&).Name = "frmConvertFile" Then
      If Forms(ThisForm&).Tag = "frmConvertFile" Then
         FormIsOpen% = True
         Exit For
      End If
   Next
   If FormIsOpen% Then
      If Not frmConvertFile.txtDestFile.Text = "Cancel" Then
         msDestFile = frmConvertFile.txtDestFile.Text
         Select Case True
            Case frmConvertFile.optDelimiter(0).value
               Delimiter& = DELIMITER_COMMA
            Case frmConvertFile.optDelimiter(1).value
               Delimiter& = DELIMITER_SEMICOLON
            Case frmConvertFile.optDelimiter(2).value
               Delimiter& = DELIMITER_SPACE
            Case frmConvertFile.optDelimiter(3).value
               Delimiter& = DELIMITER_TAB
         End Select
         FileType& = FILETYPE_CSV
         'If frmConvertFile.optFileType(1).Value Then FileType& = FILETYPE_TEXT
         SampleCount& = Val(frmConvertFile.txtCount)
         StartVal& = Val(frmConvertFile.txtStart)
         Unload frmConvertFile
      End If
   Else
      GetDestFile = True
   End If

End Function

Private Function GetSampleRange(First As Long, Total As Long) As Integer

   GetSampleRange = False
   frmConvertFile.txtCount.Left = 180
   frmConvertFile.lblCount.Left = 180
   frmConvertFile.lblCount.Top = 1820
   frmConvertFile.lblCount.Alignment = 0
   frmConvertFile.txtStart.Left = 180
   frmConvertFile.txtStart.Top = 1400
   frmConvertFile.lblStart.Top = 1000
   frmConvertFile.lblStart.Left = 180
   frmConvertFile.lblStart.Alignment = 0
   frmConvertFile.Drive1.Visible = False
   frmConvertFile.Dir1.Visible = False
   frmConvertFile.File1.Visible = False
   frmConvertFile.fraConvert.Visible = False
   'frmConvertFile.fraFileType.Visible = False
   frmConvertFile.txtDestFile.Visible = False
   frmConvertFile.txtStart.Visible = True
   frmConvertFile.cmdCancel.Left = 2480
   frmConvertFile.cmdOK.Left = 3500
   frmConvertFile.cmdCancel.Top = 1860
   frmConvertFile.cmdOK.Top = 1860
   frmConvertFile.Width = 4700
   frmConvertFile.txtCount.Text = mlSampleCount
   frmConvertFile.txtStart.Text = "0"
   frmConvertFile.Show 1
   'check if form is loaded
   For ThisForm& = 0 To Forms.Count - 1
      'If Forms(ThisForm&).Name = "frmConvertFile" Then
      If Forms(ThisForm&).Tag = "frmConvertFile" Then
         FormIsOpen% = True
         Exit For
      End If
   Next
   If FormIsOpen% Then
      First = Val(frmConvertFile.txtStart.Text)
      Total = Val(frmConvertFile.txtCount.Text)
      Unload frmConvertFile
   Else
      GetSampleRange = True
   End If

End Function

Private Sub mnuClose_Click()

   Unload Me

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
   ConfigureControls
   UpdateStatus

End Sub

Private Sub mnuLoadConfig_Click()

   Filename$ = ""
   ULStat = GetCfgFile(Filename$)
   
End Sub

Private Sub mnuPlotType_Click(Index As Integer)

   If Index = PRINT_TEXT Then
      mnuPlotType(Index).Checked = Not mnuPlotType(Index).Checked
      If mnuPlotType(Index).Checked Then mlBlockSave = GetBlockSize()
      If mnuPlotType(VOLTS_VS_TIME).Checked Then
         If Not mnuPlotType(PRINT_TEXT).Checked Then SetBlockSize mlBlockSave, False
      End If
      TFVal% = mnuPlotType(Index).Checked
      ShowText TFVal%
      
      x% = SaveFunc(Me, SShowText, 0, 0, TFVal%, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

      If mnLibType = MSGLIB Then
         RePlot False
      Else
         DisplayData
      End If
      Exit Sub
   End If

   mnPlot = False
   If Index = mnPlotType Then
      mnuPlotType(mnPlotType).Checked = Not mnuPlotType(mnPlotType).Checked
   Else
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
   DisplayData
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
   'Me.mnuNoise.Enabled = Me.mnuPlotType(HISTOGRAM).Checked

End Sub

Private Sub optUnits_Click(Index As Integer)

   mlUnits = Choose(Index + 1, CELSIUS, FAHRENHEIT, KELVIN)

End Sub

Private Sub ShowPreferences()

   ULStat = cbLogGetPreferences(TimeFormat&, TimeZone&, Units&)
   x& = SaveFunc(Me, LogGetPreferences, ULStat, TimeFormat&, TimeZone&, Units&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   txtData.Text = Choose(TimeFormat& + 1, "TIMEFORMAT_12HOUR", "TIMEFORMAT_24HOUR") & " (" & TimeFormat& & ")" & Chr$(13) & Chr$(10)
   txtData.Text = txtData.Text & Choose(TimeZone& + 1, "TIMEZONE_LOCAL", "TIMEZONE_GMT") & " (" & TimeZone& & ")" & Chr$(13) & Chr$(10)
   txtData.Text = txtData.Text & Choose(Units& + 1, "CELSIUS", "FAHRENHEIT", "KELVIN") & " (" & Units& & ")" & Chr$(13) & Chr$(10)

End Sub

Private Function TrimStrings(NullTermString As String) As String
   
   'remove spaces
   ThisArg$ = Trim(NullTermString)
   'check for null term
   If Len(ThisArg$) = 0 Then Exit Function
   If Right$(ThisArg$, 1) = Chr$(0) Then ThisArg$ = Left$(ThisArg$, Len(ThisArg$) - 1)
   TrimStrings = ThisArg$

End Function

Private Sub txtData_DblClick()

   If mnWide Then
      ConfigureControls
      mnWide = False
   Else
      txtData.Left = 120
      txtData.Width = Me.Width - 280
      mnWide = True
   End If

End Sub

Private Sub UpdateMainStatus()

  ' Board$ = mnuBoard(mnBoardIndex).Caption
  ' PrintMain "Current board: " & Board$
   
End Sub

Private Sub UpdateStatus()

   CurFilePath$ = Me.txtCurFile.Text
   If Not (CurFilePath$ = "") Then
      FileParse = Split(CurFilePath$, "\")
      NumDirs& = UBound(FileParse)
      CurFile$ = FileParse(NumDirs&)
      For ListItem& = 0 To Me.File1.ListCount - 1
         If CurFile$ = Me.File1.List(ListItem&) Then
            Me.File1.Selected(ListItem&) = True
            Exit For
         End If
      Next
      'CurFile$ = Me.File1.Filename
      If Not (CurFile$ = "") Then Filename$ = "  " & CurFile$
   End If
   lblStatus.Caption = gsConfig & Filename$

End Sub

Private Sub DisplayData()

   If mnFuncType = READ_CSV Then
      UseChDef% = True
      Chans% = UBound(ChanDefs) + 1
   Else
      Chans% = mnADChans
   End If
   FuncType% = AInScan
   If mnAcqDataType = vbDouble Then
      BufType% = 3 Or &H30
   Else
      BufType% = 2
   End If
   SetBufferType BufType%
   'SetPlotType VOLTS_VS_TIME, Me
   BoardName = msDisplayName '"USB-1616-LGR"
   
   If UseChDef% Then
      For QElement& = 0 To Chans% - 1
         TypeVal% = ChanDefs(QElement&).ChanType + 1
         TypeStr$ = Choose(TypeVal%, "A", "d", "D", "C", "CL", _
         "CH", "J", "T", "SE", "DF", "SP", "EN", "C!")
         ChanVar = ChanVar & TypeStr$ & Format$(ChanDefs(QElement&).ChanNum, "0") & ", "
         RangeVar = RangeVar & Format$(ChanDefs(QElement&).ChanRange, "0") & ", "
      Next
      ChanVar = Left$(ChanVar, Len(ChanVar) - 2)
      If Len(RangeVar) > 0 Then RangeVar = Left$(RangeVar, Len(RangeVar) - 2)
      QVar = "Q"
   Else
      lblStatus.Caption = msStringID & msConfig & " " & msOpt
      If mlQCount > 0 Then
         'to do - configure for mixed data
         If mlAIChans > 0 Then ArrSize& = UBound(manTypeArray)
         If mlQuadChans > 0 Then Chans% = mlQuadChans
         For QElement& = 0 To mlQCount - 1
            If mlQuadChans > 0 Then
               QVar = "Q"
            End If
            If mlAIChans > 0 Then
               If (ArrSize& + 1) < mlQCount Then
                  TypeStr$ = "A"
               Else
                  If (Not manTypeArray(QElement&) < 0) Then
                     TypeVal% = (manTypeArray(QElement&) And &HFF) + 1
                     TypeStr$ = Choose(TypeVal%, "A", "d", "D", "C", "CL", _
                     "CH", "J", "T", "SE", "DF", "SP", "EN", "C!")
                  Else
                     TypeStr$ = "A"
                  End If
               End If
               ChanVar = ChanVar & TypeStr$ & Format$(ChanDefs(QElement&).ChanNum, "0") & ", "
               RangeVar = RangeVar & Format$(ChanDefs(QElement&).ChanRange, "0") & ", "
               'ChanVar = ChanVar & TypeStr$ & Format$(manChanArray(QElement&), "0") & ", "
               'RangeVar = RangeVar & Format$(manGainArray(QElement&), "0") & ", "
            End If
         Next QElement&
         ChanVar = Left$(ChanVar, Len(ChanVar) - 2)
         RangeVar = Left$(RangeVar, Len(RangeVar) - 2)
         QVar = "Q"
      ElseIf mnNumAIChans > 0 Then
         NumChans% = (mnLastChan - mnFirstChan)
         For ChanIndex% = 0 To NumChans%
            If ChanIndex% > UBound(manGainArray) Then Exit Sub
            CurChan% = ChanIndex% + mnFirstChan
            ChanVar = ChanVar & Format(CurChan%, "0") & ", "
            If UBound(manGainArray) >= (CurChan%) Then
            RangeVar = RangeVar & Format$(manGainArray(CurChan%), "0") & ", "
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
   End If
   If mfRate < 1000 Then
      RateVal = mfRate & "Hz"
   Else
      RateVal = mfRate / 1000 & "kHz"
   End If
   If Not (mfRate = mfRateReturned) Then
      If mfRateReturned < 1000 Then
         RateVal = RateVal & " / " & mfRateReturned & "Hz"
      Else
         RateVal = RateVal & " / " & mfRateReturned / 1000 & "kHz"
      End If
   End If
   
   TotalCount& = mlCount '* Chans%
   SetDetails FuncType%, ChanVar, QVar, TotalCount&, RateVal, RangeVar, BoardName, TotalCount&
   PlotBuffer mlHandle, TotalCount&, Chans% - 1
   
End Sub

Private Function ReadCSVtoMem(Filename As String, ByRef AcqDataType As Long, AcqPoints As Long, Chans As Integer) As Long

   On Error GoTo csvError
   Dim CSVLines() As String
   Dim QChans() As Integer
   Dim TimeArray() As Double
   
   Dim AIData() As Long
   Dim FltData() As Double
   Dim QData() As Long
   mlNumCSVLines = 0
   mnTimeStamped = False
   mlPreTrigCount = 0
   mlPTCountReturn = 0
   mlTotalCount = 0
   mlCount = 0
   mfRateReturned = 0: mfRate = 0
   FirstDataLine& = 0
   
   Open Filename For Input As #16
   Dim BinDat0 As Byte, BinDat1 As Byte, BinDat2 As Byte
   mlNumCSVLines = 0
   Do While Not EOF(16)
      Line Input #16, A1
      If A1 = "" Then A1 = "; "
      ReDim Preserve CSVLines(mlNumCSVLines)
      CSVLines(mlNumCSVLines) = A1
      mlNumCSVLines = mlNumCSVLines + 1
   Loop
   mlNumCSVLines = mlNumCSVLines - 1
   Close #16
   mlTotalCount = 0
   mfRateReturned = 0
   mnEngApp = True
   
   'evaluate file - find start and type of data
   If mlNumCSVLines > 0 Then CurLine$ = CSVLines(0)
   StartElement% = 0

   Me.txtData.Text = ""
   ThisLine& = 0
   Do While FirstDataLine& = 0
      TestLine$ = CSVLines(ThisLine&)
      'look for characters that would not occur in a
      'data line (any alpha that isn't hex)
      EndString& = Len(TestLine$) - 1
      For ChrPos& = 1 To EndString&
         CharVal& = Asc(Mid(TestLine$, ChrPos&, 1))
         TimeTest$ = Mid(TestLine$, ChrPos&, 1)
         If Not ((TimeTest$ = "A") Or (TimeTest$ = "P") Or (TimeTest$ = "M")) Then
            If (CharVal& > 58) And (CharVal& < 65) Then Exit For
            If (CharVal& > 70) And (CharVal& < 97) Then Exit For
            If (CharVal& > 102) And (CharVal& < 120) Then Exit For
            If (CharVal& > 120) And (CharVal& < 127) Then Exit For
         End If
      Next
      If ChrPos& > EndString& Then FirstDataLine& = ThisLine&
      ThisLine& = ThisLine& + 1
   Loop
   
   ReadingHeader% = True
   For LineNumber& = 0 To FirstDataLine& - 1
      CurLine$ = CSVLines(LineNumber&)
      If Mid(CurLine$, Len(CurLine$)) = "," Then
         CurLine$ = Left(CurLine$, Len(CurLine$) - 1)
      End If
      ParseDAQHeader CurLine$, LineNumber&
   Next
   mlTotalCount = mlTotalCount * mlQCount
   mlPreTrigCount = mlPreTrigCount * mlQCount
   mlPTCountReturn = mlPreTrigCount
   mlCount = mlPreTrigCount + mlTotalCount
   NumChans% = UBound(ChanDefs) + 1
   If FirstDataLine& < 4 Then
      CurLine$ = CSVLines(LineNumber&)
      DataElements = Split(CurLine$, ",")
      NumDE& = UBound(DataElements) + 1
      If DataElements(NumDE& - 1) = "" Then NumDE& = NumDE& - 1
      FirstDat& = NumDE& - NumChans%
      For Ch% = FirstDat& To NumChans%
         DatString$ = DataElements(Ch%)
         If (InStr(1, DatString$, ".")) Then ChanDefs(Ch% - 1).ChanUnits = "Volts"
      Next
   End If
   
   If mlCount = 0 Then
      'not specified in file text
      TotalCount& = (mlNumCSVLines - FirstDataLine&) + 1
      mlCount = TotalCount& * mlQCount
      mlTotalCount = mlCount
      If mfRateReturned = 0 Then CalcRate% = True
   End If
   
   mnResolution = 16
   
   'NumChans% = UBound(ChanDefs) + 1
   'determine the type of data in each channel
   For Ch% = 0 To NumChans% - 1
      TypeOfChan& = ChanDefs(Ch%).ChanType
      Units$ = ChanDefs(Ch%).ChanUnits
      If Units$ = "Volts" Or Units$ = "mVolts" Then UseFloat% = True
   Next
   
   SamplesPerChan& = (mlCount / mlQCount) - 1
   NumSamples& = mlCount - 1
   If UseFloat% Then
      ReDim FltData(NumChans% - 1, SamplesPerChan&)
   Else
      ReDim AIData(NumChans% - 1, SamplesPerChan&)
   End If
   CurLine$ = CSVLines(FirstDataLine&)
   If Mid(CurLine$, Len(CurLine$)) = "," Then
      CurLine$ = Left(CurLine$, Len(CurLine$) - 1)
   End If
   LineData = Split(CurLine$, ",")
   Elements% = UBound(LineData) + 1
   DataSlot% = 1
   If mnEngApp Then
      'StartElement% = StartElement% - 1
      DataSlot% = 0
   End If
   StartElement% = Elements% - (NumChans%)   ' - DataSlot%
   'TotalCount& = NumChans% * mlTotalCount
   
   If StartElement% > DataSlot% Then
      mnTimeStamped = True
      ReDim TimeArray(SamplesPerChan&)
   End If
   
   For DataLine& = FirstDataLine& To mlNumCSVLines
      CurLine$ = CSVLines(DataLine&)
      LineData = Split(CurLine$, ",")
      CurAChan% = 0
      CurQChan% = 0
      CurSample& = DataLine& - FirstDataLine&
      For CurElement% = StartElement% To StartElement% + (NumChans% - 1)
         If NumChans% > 0 Then
            ListElement$ = Trim(LineData(CurElement%))
            If Left(ListElement$, 2) = "0x" Then
               ListElement$ = "&H" & Mid(ListElement$, 3)
            End If
            If UseFloat% Then
               FloatData# = Val(ListElement$)
               FltData(CurAChan%, CurSample&) = FloatData#
            Else
               DataValue& = Val(ListElement$)
               If DataValue& < 0 Then DataValue& = DataValue& + 65536
               AIData(CurAChan%, CurSample&) = DataValue&
            End If
            CurAChan% = CurAChan% + 1
         End If
      Next
      If mnTimeStamped Then
         CurStamp$ = LineData(DataSlot%)
         Stamps = Split(CurStamp$, ":")
         TimeElements% = UBound(Stamps)
         If TimeElements% > 0 Then
            Secs$ = Stamps(TimeElements%)
            'Secs$ = Left(Secs$, Len(Secs$) - 1) 'remove end quote
            TimeArray(CurSample&) = Val(Secs$)
         End If
      End If
   Next
   
   If mnTimeStamped And CalcRate% Then
      StartDiff# = TimeArray(1) - TimeArray(0)
      If (StartDiff# < 0) Then StartDiff# = 60 - TimeArray(0)
      If StartDiff# < 0 Then StartDiff# = StartDiff# + 60
      Threshold# = StartDiff# * 0.1
      For i& = 1 To SamplesPerChan&
         CurDiff# = TimeArray(i&) - TimeArray(i& - 1)
         If CurDiff# < 0 Then CurDiff# = CurDiff# + 60
         If Abs(StartDiff# - CurDiff#) > Threshold# Then
            Exit For
         End If
         CumDiff# = CumDiff# + CurDiff#
      Next
      TimingChange& = i& - 1
      AvgTime# = CumDiff# / TimingChange&
      DeltaSample& = (i& * mlQCount) - 1
      
      If AvgTime# = 0 Then
         'external clock
      Else
         mfRate = 1 / AvgTime#
         mfRateReturned = mfRate
      End If
      SampleSet& = SamplesPerChan&
      If TimingChange& < SamplesPerChan& Then
         CumDiff# = 0
         SampleSet& = SamplesPerChan& - TimingChange&
         For i& = TimingChange& + 1 To SamplesPerChan&
            CurDiff# = TimeArray(i&) - TimeArray(i& - 1)
            If CurDiff# < 0 Then CurDiff# = CurDiff# + 60
            CumDiff# = CumDiff# + CurDiff#
         Next
      End If
      AvgTime# = CumDiff# / (SampleSet&)
      mlPreTrigCount = 0
      If TimingChange& < SampleSet& Then mlPreTrigCount = TimingChange&
      mlPTCountReturn = mlPreTrigCount
      If AvgTime# = 0 Then
         'external clock
      Else
         mfRateReturned = 1 / AvgTime#
      End If
   End If
   If UseFloat% Then
      'BufAlloc64
      DataHandle& = ScaledBufAlloc(Me, mlCount)
      mnAcqDataType = vbDouble
      If DataHandle& = 0 Then Exit Function
      If NumChans% > 0 Then
         ULStat = WDblArrayToBuf(Me, DataHandle&, FltData(), mlCount)
      End If
   Else
      DataHandle& = BufAlloc32(Me, mlCount)
      mnAcqDataType = vbLong
      If DataHandle& = 0 Then Exit Function
      If NumChans% > 0 Then
         ULStat = WArrayToBuf32(Me, DataHandle&, AIData(), mlCount, True)
      End If
   End If
   mlAcqPoints = mlCount
   Chans = Chans%
   ReadCSVtoMem = DataHandle&
   UpdateStatus
   
   Exit Function
   
csvError:
   MsgBox Error(Err), , "Error Opening CSV file " & Filename
   Exit Function
   Resume Next

End Function

Sub ParseDAQHeader(CurrentLine As String, LineNumber As Long)

   Elements = Split(CurrentLine, ":")
   If Not IsEmpty(Elements) Then
      NumColons& = UBound(Elements)
      If NumColons& = 0 Then
         If Left(CurrentLine, 5) = "Scan#" Then
            mlDataLine = LineNumber + 1
            mnEngApp = False
         Else
            Elements = Split(CurrentLine, ",")
            NumConfigs& = UBound(Elements)
            If NumConfigs& > 0 Then
               If Elements(0) = "" Then ParseChanDef NumConfigs&, Elements
            End If
         End If
         Exit Sub
      Else
         If NumColons& > 6 Then
            'likely bitwise digital
            Elements = Split(CurrentLine, ", ")
            NumConfigs& = UBound(Elements)
            ParseChanDef NumConfigs&, Elements
         End If
      End If
      EndLineDef& = Len(Elements(0)) + 1
      FirstElement$ = Elements(0)
      Select Case FirstElement$
         Case "Device"
            If NumColons& > 0 Then
               BoardName$ = Elements(1)
               msDisplayName = Trim(BoardName$)
               Me.txtData.Text = msDisplayName & vbCrLf
            End If
         Case "Serial No"
            If NumColons& > 0 Then
               SerialNum$ = Trim(Elements(1))
               Me.txtData.Text = Me.txtData.Text & SerialNum$ & vbCrLf
            End If
         Case "Start Time", "Log Enable Time", "Trigger Time"
            StartParams$ = Right(CurrentLine, Len(CurrentLine) - (EndLineDef& + 1))
            TimeElements = Split(StartParams$, " ")
            NumParams% = UBound(TimeElements)
            FirstParam$ = TimeElements(0) & vbCrLf
            If NumParams% > 0 Then TimeString$ = TimeElements(1)
            Me.txtData.Text = Me.txtData.Text & FirstParam$ & TimeString$ & vbCrLf
         Case "Post-Trigger Scan Count"
            CountLoc& = InStr(1, CurrentLine, ":")
            CountString$ = Mid(CurrentLine, CountLoc& + 1)
            PosttrigCount& = Val(CountString$)
            mlTotalCount = PosttrigCount& '* mlQCount
            mlCount = mlPreTrigCount + mlTotalCount
         Case "Pre-Trigger Scan Count"
            CountLoc& = InStr(1, CurrentLine, ":")
            CountString$ = Mid(CurrentLine, CountLoc& + 1)
            PretrigCount& = Val(CountString$)
            mlPreTrigCount = PretrigCount& '* mlQCount
            mlPTCountReturn = mlPreTrigCount
         Case "            Scan Rate(Hz)", "Pre-Trigger Scan Rate(Hz)"
            RateLoc& = InStr(1, CurrentLine, ":")
            RateString$ = Mid(CurrentLine, RateLoc& + 1)
            mfRate = Val(RateString$)
            If mlPreTrigCount = 0 Then mfRate = 0
         Case "             Scan Rate", "Post-Trigger Scan Rate"
            RateLoc& = InStr(1, CurrentLine, ":")
            RateString$ = Mid(CurrentLine, RateLoc& + 1)
            mfRateReturned = Val(RateString$)
            If mfRate = 0 Then mfRate = mfRateReturned
            'mlPreTrigCount = TotalCount& * mlQCount
            'mlPTCountReturn = mlPreTrigCount
         Case Else
            NameLoc& = InStr(1, CurrentLine, " (SN")
            If NameLoc& > 1 Then
               LineParams = Split(CurrentLine, " ")
               msDisplayName = LineParams(0)
               Me.txtData.Text = msDisplayName & vbCrLf
               If Len(CurrentLine) > NameLoc& + 5 Then
                  SerNumber$ = Mid(CurrentLine, NameLoc& + 5, 12)
                  Me.txtData.Text = Me.txtData.Text & SerNumber$ & vbCrLf
               End If
            End If
      End Select
   End If
   
End Sub
Sub ParseCfgAppHeader(CurrentLine As String)

   Elements = Split(CurrentLine, ":")
   If Not IsEmpty(Elements) Then
      FirstElement$ = Elements(0)
      Select Case FirstElement$
         Case "Start Time"
            TimeLoc& = InStr(1, CurrentLine, ":")
            TimeString$ = Mid(CurrentLine, TimeLoc& + 1)
         Case "Post-Trigger Scan Count"
            CountLoc& = InStr(1, CurrentLine, ":")
            CountString$ = Mid(CurrentLine, CountLoc& + 1)
            mlCount = Val(CountString$) * mlQCount
            mlTotalCount = mlPreTrigCount + mlCount
         Case "Pre-Trigger Scan Count"
            CountLoc& = InStr(1, CurrentLine, ":")
            CountString$ = Mid(CurrentLine, CountLoc& + 1)
            mlPreTrigCount = Val(CountString$) * mlQCount
            mlPTCountReturn = mlPreTrigCount
         Case "            Scan Rate(Hz)"
            RateLoc& = InStr(1, CurrentLine, ":")
            RateString$ = Mid(CurrentLine, RateLoc& + 1)
            mfRate = Val(RateString$)
         Case "             Scan Rate"
            RateLoc& = InStr(1, CurrentLine, ":")
            RateString$ = Mid(CurrentLine, RateLoc& + 1)
            mfRateReturned = Val(RateString$)
         Case Else
            NameLoc& = InStr(1, CurrentLine, " (SN")
            If NameLoc& > 1 Then msDisplayName = Left(CurrentLine, NameLoc& - 1)
      End Select
   End If
   
End Sub

Function ParseColumnLabels(LabelName As String, ByVal CurElement As Integer) As Integer

   Select Case LabelName
      Case "Scan Time", "Time"
         mnTimeStamped = True
         ParseColumnLabels = True
         Exit Function
   End Select
   Hint$ = UCase(Left(LabelName, 2))
   Select Case Hint$
      Case "CH", "AI"
         ChanParams = Split(LabelName, "_")
         NumParams% = UBound(ChanParams)
         If NumParams% > 0 Then
            For Param% = 1 To NumParams% '- 1
               CurParam$ = ChanParams(Param%)
               Select Case Param%
                  Case 1
                     manChanArray(CurElement) = Val(CurParam$)
                  Case 2
                     ChMode% = ANALOG
                     If CurParam$ = "SE" Then ChMode% = ANALOG_SE
                     manTypeArray(CurElement) = ChMode%
                  Case 3
                     GainCode% = GetLoggerRange(CurParam$)
                     manGainArray(CurElement) = GainCode%
               End Select
            Next
         Else
            Chan$ = Mid(LabelName, 3)
            If IsNumeric(Chan$) Then
               manChanArray(CurElement) = Val(Chan$)
            End If
         End If
      Case "DI"
         manTypeArray(CurElement) = DIGITAL16  '2
         manGainArray(CurElement) = -1
      Case "QU"
         ChanParams = Split(ListElement$, "_")
         NumParams% = UBound(ChanParams)
         If NumParams% > 0 Then
            For Param% = 1 To NumParams% '- 1
               CurParam$ = ChanParams(Param%)
               Select Case Param%
                  Case 1
                     'channel
                     manChanArray(CurElement) = Val(CurParam$)
                     manGainArray(CurElement) = -1
                  Case 2
                     'low or high
                     If CurParam$ = "Low" Then
                        manTypeArray(CurElement) = CTR32LOW
                     Else
                        manTypeArray(CurElement) = CTR32HIGH
                     End If
                  Case 3
                     'manGainArray(NumAChans%) = GainCode%
               End Select
            Next
         End If
         'NumQChans% = NumQChans% + 1
   End Select

End Function

Public Sub InitForm(FunctionInit As Integer)

   'mnuBoard_Click (FunctionInit)
   mnuFuncArray_Click (FunctionInit)
   
End Sub

Public Function GetQueueList(QChans As Variant, QGains As Variant, QTypes As Variant) As Long

   ChanElements% = UBound(ChanDefs)
   ReDim manChanArray(ChanElements%)
   ReDim manGainArray(ChanElements%)
   ReDim manTypeArray(ChanElements%)
   ChanCount% = ChanElements% + 1
   For Element% = 0 To ChanElements%
      manChanArray(Element%) = ChanDefs(Element%).ChanNum
      manGainArray(Element%) = ChanDefs(Element%).ChanRange
      manTypeArray(Element%) = ChanDefs(Element%).ChanType
   Next
   QChans = manChanArray()
   QGains = manGainArray()
   QTypes = manTypeArray()
   GetQueueList = ChanCount%
   
End Function

Public Function GetDataHandle(AcqOrGen As Integer, DataType As Long, NumSamples As Long) As Long

   Select Case AcqOrGen
      Case ACQUIREDDATA
         GetDataHandle = mlHandle
         If mnLibType = MSGLIB Then
            'If (Not (Me.mnuToEng.Checked)) And _
            (mnAcqDataType = vbDouble) Then Convert& = &H10
         End If
         'to do - set this at file read (ReadCSVtoMem)
         mnAcqDataType = vbLong
         mlAcqPoints = mlCount
         DataType = mnAcqDataType Or Convert&
         NumSamples = mlAcqPoints
      Case GENERATEDDATA
         GetDataHandle = mlGenHandle
         DataType = mnGenDataType
         NumSamples = mlGenPoints
   End Select
   
End Function

Private Sub ParseChanDef(NumDefs As Long, DefElements As Variant)

   For DefNum& = 0 To NumDefs
      DefString$ = Trim(DefElements(DefNum&))
      If Len(DefString$) > 0 Then
         TypeInd$ = Left(DefString$, 2)
         Select Case TypeInd$
            Case "AI"
               ReDim Preserve ChanDefs(NumChans%)
               UnitParse = Split(DefString$, "(")
               If UBound(UnitParse) > 0 Then Units$ = Left(UnitParse(1), Len(UnitParse(1)) - 1)
               ChanDefs(NumChans%).ChanUnits = Units$
               ChParms$ = UnitParse(0)
               If (InStr(1, ChParms$, "Diff") > 0) Then
                  Parms = Split(ChParms$, "Diff")
                  ChanDefs(NumChans%).ChanConfig = "Diff"
                  ChanDefs(NumChans%).ChanType = ANALOG_DIFF
               Else
                  Parms = Split(ChParms$, "Se")
                  ChanDefs(NumChans%).ChanConfig = "Se"
                  ChanDefs(NumChans%).ChanType = ANALOG_SE
               End If
               ChNum$ = Mid(ChParms$, 3, 2)
               Channel% = Val(ChNum$)
               ChanDefs(NumChans%).ChanNum = Channel%
               If UBound(Parms) > 0 Then
                  Range$ = Parms(1)
                  Select Case Range$
                     Case "30V"
                        ChRange& = BIP30VOLTS   'not a UL range
                     Case "10V"
                        ChRange& = BIP10VOLTS
                     Case "5V"
                        ChRange& = BIP5VOLTS
                     Case "1V"
                        ChRange& = BIP1VOLTS
                  End Select
                  ChanDefs(NumChans%).ChanRange = ChRange&
               End If
               NumChans% = NumChans% + 1
               mlQCount = mlQCount + 1
            Case "Ch"
               ReDim Preserve ChanDefs(NumChans%)
               ElementParse = Split(DefString$, "_")
               NumEls& = UBound(ElementParse)
               If NumEls& = 3 Then
                  ChNum$ = ElementParse(1)
                  Channel% = Val(ChNum$)
                  ChanDefs(NumChans%).ChanNum = Channel%
                  ChParms$ = ElementParse(2)
                  If ChParms$ = "SE" Then
                     ChanDefs(NumChans%).ChanConfig = "Se"
                     ChanDefs(NumChans%).ChanType = ANALOG_SE
                  Else
                     ChanDefs(NumChans%).ChanConfig = "Diff"
                     ChanDefs(NumChans%).ChanType = ANALOG_DIFF
                  End If
                  Range$ = ElementParse(3)
                  Select Case Range$
                     Case "30V"
                        ChRange& = BIP30VOLTS   'not a UL range
                     Case "10V"
                        ChRange& = BIP10VOLTS
                     Case "5V"
                        ChRange& = BIP5VOLTS
                     Case "1V"
                        ChRange& = BIP1VOLTS
                  End Select
               End If
               ChanDefs(NumChans%).ChanRange = ChRange&
               NumChans% = NumChans% + 1
               mlQCount = mlQCount + 1
            Case "CT"
               ReDim Preserve ChanDefs(NumChans%)
               CtType$ = Mid(DefString$, 5)
               Select Case CtType$
                  Case "Enc"
                     ChanDefs(NumChans%).ChanType = ENCDR
                  Case "Ctr", "Ud", "Prd", "Pw", "Tm"
                     ChanDefs(NumChans%).ChanType = CTR48
                  Case Else
                     ChanDefs(NumChans%).ChanType = -1
               End Select
               ChNum$ = Mid(DefString$, 4, 2)
               Channel% = Val(ChNum$)
               ChanDefs(NumChans%).ChanNum = Channel%
               NumChans% = NumChans% + 1
               mlQCount = mlQCount + 1
            Case "DI"
               ReDim Preserve ChanDefs(NumChans%)
               ChanDefs(NumChans%).ChanType = DIGITAL16
               BitChk = Split(DefString$, ":")
               If UBound(BitChk) > 0 Then
                  ChNum$ = BitChk(1)
               Else
                  ChNum$ = Mid(DefString$, 3, 2)
               End If
               Channel% = Val(ChNum$)
               ChanDefs(NumChans%).ChanNum = Channel%
               NumChans% = NumChans% + 1
               mlQCount = mlQCount + 1
            Case "Qu"
               ReDim Preserve ChanDefs(NumChans%)
               ChanDefs(NumChans%).ChanType = ENCDR
               ElementParse = Split(DefString$, "_")
               NumEls& = UBound(ElementParse)
               If NumEls& = 2 Then
                  ChNum$ = ElementParse(1)
                  Channel% = Val(ChNum$)
                  ChanDefs(NumChans%).ChanNum = Channel%
               End If
               NumChans% = NumChans% + 1
               mlQCount = mlQCount + 1
         End Select
      End If
   Next
   'mlQCount = NumChans% '- 1
   'mlTotalCount = mlTotalCount * mlQCount
   'mlPreTrigCount = mlPreTrigCount * mlQCount
   'mlPTCountReturn = mlPreTrigCount
   'mlCount = mlPreTrigCount + mlTotalCount
   
End Sub

Public Sub SetStaticOption(NewOption As Long)

   If NewOption = 0 Then
      mlStaticOpt = NewOption
   Else
      mlStaticOpt = mlStaticOpt Or NewOption
   End If
   msOpt = GetOptionsString(mlStaticOpt, ANALOG_IN)
   'mnRefreshProps = True
   
End Sub

Public Function GetStaticOption() As Long

   GetStaticOption = mlStaticOpt
   
End Function

