VERSION 5.00
Begin VB.Form frmScript 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scripting"
   ClientHeight    =   885
   ClientLeft      =   1470
   ClientTop       =   1875
   ClientWidth     =   8340
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   885
   ScaleWidth      =   8340
   Tag             =   "Scripter"
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&UniTest"
      Height          =   315
      Index           =   7
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Go to manual test app (left click) or view script header (right click) for the current script..."
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&>>"
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Move to next (left click) or specific (right click) line in the current script..."
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&<<"
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Move to previous (left click) or specific (right click) line in the current script..."
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&New"
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Record a new script file..."
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&Open"
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Open an existing script file..."
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&Run"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Run the current script..."
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer tmrScript 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6660
      Top             =   120
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "&Step"
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   3840
      TabIndex        =   5
      ToolTipText     =   "Step through the current script..."
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdScript 
      Appearance      =   0  'Flat
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Stop the current script..."
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblScriptRate 
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7080
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLoopStat 
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7140
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblScriptStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Script Status:  Idle"
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFileOpts 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Append"
         Index           =   2
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThread 
         Caption         =   "&Thread UL Calls"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuScripter 
         Caption         =   "Scripting"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuScriptDir 
         Caption         =   "Script Directory"
      End
      Begin VB.Menu mnuMasterDir 
         Caption         =   "Master Directory"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuTimer 
         Caption         =   "Timer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuScriptInf 
         Caption         =   "Show Comments"
      End
      Begin VB.Menu mnuTestOptions 
         Caption         =   "Select Tests..."
      End
      Begin VB.Menu mnuScriptEval 
         Caption         =   "Evaluation"
         Begin VB.Menu mnuPrintAll 
            Caption         =   "Print All"
         End
         Begin VB.Menu mnuEvalPrint 
            Caption         =   "Print Failure"
         End
         Begin VB.Menu mnuEvalPause 
            Caption         =   "Pause on Failure"
         End
         Begin VB.Menu mnuGuardbands 
            Caption         =   "Global Guardbands..."
         End
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Information"
      Begin VB.Menu mnuVariables 
         Caption         =   "Variables..."
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const mOPENFORM = 5001
Const mCLOSEFORM = 5002
Const mIDLE = 0
Const mRECORDING = 1
Const mRUNNING = 2
Const mSTEPPING = 4
Const mREADBACK = 5
Const mREADAHEAD = 6
Const mOPEN = 0
Const mNEW = 1
Const mRUN = 2
Const mSTOP = 3
Const mSTEP = 4
Const mPREVIOUS = 5
Const mNEXT = 6
Const mUNITEST = 7
Const mPAUSED = 8
Const mMASTER = 1
Const mSUBSCRIPT = 2
Const mNOOPLINE = 0
Const mCOMMENT = 1
Const mCOMMAND = 2
Const mINVALID = 3
Const mEVALUATE = 4
Const mLINETEXT = 5

Const WAITINGENDIF = 1
Const INFORLOOP = 2
Const INDOLOOP = 4
Const WAITINGENDDO = 8
Const INDOWHILELOOP = 16
Const WAITINGENDFOR = 32

Dim msScriptDir As String
Dim msBoardDir As String
Dim msScriptPath As String
Dim msMasterPath As String, msDevParmPath As String
Dim msScriptStatus As String, msMasterSelection As String
Dim mnMasterRun As Integer, mnStartupScript As Integer
Dim mnAbortScript As Integer
Dim mnCurrentMode As Integer
Dim mnStepping As Integer
Dim mnHeaderLines As Integer

Dim mnInstances(9) As Integer
Dim mnJustReading As Integer
Dim mlScriptTime As Long
Dim mlPauseTime As Long, mlDelayStart As Long
Dim mlScriptStartTime As Long
Dim mnCalMode As Integer, mlCalVal As Single
Dim msDefaultPath As String, msDefaultDrive As String
Dim mnScriptMode As Integer
Dim mnReadMaster As Integer
Dim msMasterFile As String, msScriptFile As String
Dim mnScriptClosed As Integer 'used to determine if ended locally or on error elsewhere
Dim msDevName As String, mnDevDupe As Integer
Dim masMaster() As String
Dim masScript() As String
Dim masInit() As String
Dim mnNumMasterLines As Integer, mnNumScriptLines As Integer
Dim mnNumInitLines As Integer
Dim mnMasterLine As Integer
Dim mnScriptLine As Integer
Dim msHeader As String
Dim mnMasterLoop As Integer, mnScriptLoop As Integer
Dim mnPaused As Integer

Dim mvVal1 As Variant, mvVal2 As Variant, mvVal3 As Variant
Dim mvVar1 As Variant, mvVar2 As Variant, mvVar3 As Variant
Dim msVerification As String
Dim mlAmplGB As Long, mlRateGB As Long
Dim mfSrcOffset As Single, mlMvgAvgGB As Long
Dim mnULErrHandling As Integer, mnULErrReporting As Integer
Dim mnLocalErrHandling As Integer, mnLocalErrReporting As Integer
Dim mfrmFormRef As Form, mnCheckingStatus As Integer
Dim mlWaitForCount As Long, mnWaitingEvent As Integer, mnWaitingStatVal As Integer
Dim mlEventType As Long, mlEventData As Long, mlTimeout As Long
Dim mlStopCount As Long
Dim mlStaticOpt As Long, mnCurConditional As Integer
Dim mlLowCtr As Long, mlHighCtr As Long, mlForCount As Long
Dim mnForStart As Integer, mnIfStarts As Integer, mnIfEnds As Integer
Dim mlStaticRange As Long
Dim mnSimIn As Integer, mnSimOut As Integer
Dim mlULError As Long
Dim mnDoStart As Integer, msDoCondition As String
Dim mnLogOpen As Integer, mlScriptRate As Long
Dim maStringList() As String, mnNumStrings As Integer
'Dim maCSVList() As String, mnNumCSVArgs As Integer

Dim ScriptVars() As Variant, mnNumScriptVars As Integer
Dim maForNest() As Long, mnForNest As Integer, mnInvalidEnd As Integer
Dim mParamList As New Collection
Dim msLatestRev As String, msTestString As String
Dim mnStartCommand As Integer
Dim mbSetLibToUL As Boolean

Private Sub AppendRecord()

   Caption = "Record Script"
   gnScriptSave = True
   mnCurrentMode = mRECORDING
   Open msScriptPath For Append As #2
   cmdScript(mSTOP).Caption = "Stop"
   cmdScript(mSTEP).ENABLED = False
   cmdScript(mSTOP).ENABLED = True
   lblScriptStatus.ForeColor = &HFF
   msScriptStatus = "Script Status:  Appending "
   lblScriptStatus.Caption = msScriptStatus & msScriptPath
   tmrScript.ENABLED = True

End Sub

Private Sub CloseScript()

   mnCurrentMode = mIDLE
   tmrScript.ENABLED = False
   cmdScript(mPREVIOUS).ENABLED = False
   cmdScript(mPREVIOUS).Caption = "&<<"
   If Not mnMasterRun Then
      gnScriptSave = False
      gnScriptRun = False
      mnStepping = False
      'mnRunning = False
      cmdScript(mSTEP).Caption = "&Step"
      cmdScript(mSTOP).ENABLED = False
      lblScriptStatus.ForeColor = &HFF0000
      If Len(msMasterFile) Then
         Preface$ = "Master Script " & msMasterFile & "idle" & vbCrLf & "("
         Terminator$ = ")"
      Else
         Preface$ = "Ready to load master script"
         Terminator$ = ""
         Me.cmdScript(mRUN).ENABLED = False
      End If
      msScriptStatus = "Script Status:  " & Preface$
      lblScriptStatus.Caption = msScriptStatus & msScriptPath & Terminator$
      cmdScript(mSTEP).ENABLED = True
   End If
   mnScriptLine = 0
   mlScriptTime = 0
   mlScriptStartTime = -1
   mnScriptClosed = True
   mnHeaderLines = -1
   mnNumInitLines = 0
   ReDim ScriptVars(1, 0)
   ReDim maForNest(5, 0)
   ReDim maStringList(0)
   mnNumStrings = -1
   mnNumScriptVars = 0
   mnReadMaster = False
   For Each ListObject In mParamList
      'remove
      mParamList.Remove (1)
   Next

End Sub

Private Sub cmdScript_Click(Index As Integer)

   If (mnScriptMode = mNEXT) Or (mnScriptMode = mPREVIOUS) Then
      If Not (mnScriptMode = Index) Then
         'the script line previously read (but not acted on)
         'will execute in this case
         If mnMasterLoop Then
            If mnMasterLine > 0 Then mnMasterLine = mnMasterLine - 1
         Else
            If mnScriptLine > 0 Then mnScriptLine = mnScriptLine - 1
         End If
      End If
   
   End If
   
   If Not (mnCurrentMode = mRUNNING) Or gnScriptPaused Then cmdScript(mRUN).ENABLED = True
   If (mnScriptMode = mSTOP) And (Index > mNEW) Then
      CloseScript
      If Not OpenMasterScript() Then Exit Sub
   End If
   DoEvents
   mnScriptMode = Index
   Select Case mnScriptMode
      Case mOPEN
         msMasterSelection = ""
         mnStartCommand = False
         mnuFile_Click (1)
      Case mNEW
         mnuFile_Click (0)
      Case mRUN
         frmScriptInfo.picMasterStat.Line (0, 0)-(mnNumMasterLines, 75), &H8000000F, BF
         PauseState% = gnScriptPaused
         gnScriptPaused = False
         cmdScript(mRUN).ENABLED = False
         If mnStepping Then
            mnStepping = False
            DoEvents
            If Not PauseState% Then ReadScript
            tmrScript.ENABLED = True
            'mnScriptLine = mnScriptLine - 1
            'ResetScript
         Else
            If mnMasterRun Then
               mnReadMaster = True
               mnCurrentMode = mRUNNING
               ReadMaster mnCurrentMode
            Else
               tmrScript.ENABLED = True
            End If
         End If
      Case mSTEP
         If mnStepping Then
            If mnCurrentMode = mREADAHEAD Then mnCurrentMode = mSTEPPING
            Select Case mnCurrentMode
               Case mIDLE
                  LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, Args)
                  If mnMasterLine > 0 Then MasterType% = ParseScript(mMASTER, mnMasterLine - 1, MasterArgs)
                  If MasterType% = mCOMMAND Then
                     If ((VarType(MasterArgs) And vbArray) = vbArray) And ((VarType(Args) And vbArray) = vbArray) Then
                        'get the board name from the master script and insert into subscript args
                        DevArg$ = Trim(MasterArgs(3))
                        If (Args(1) = mOPENFORM) And Not (DevArg$ = "0") Then
                           Args(3) = DevArg$
                        End If
                     End If
                     If LineType% = mCOMMAND Then
                        UpdateScriptStatus Args
                        mnCurrentMode = mRUNNING
                        mnStepping = True
                     End If
                  Else
                     ReadMaster mnCurrentMode
                  End If
                  frmScriptInfo.picMasterStat.Line (0, 0)-(mnNumMasterLines, 75), &H8000000F, BF
               Case mRECORDING
                  If gnScriptSave Then
                     DelayVal = InputBox("Number of seconds to delay:  ", "Add Delay to Script", "5")
                     If Not IsNull(DelayVal) Then Delay$ = Format$(DelayVal, "0")
                     Print #2, "0, 3000, 0, " & Delay$ & ",,,,,,,,,,,"
                  Else
                     AppendRecord
                  End If
               Case mRUNNING
                  ReadScript
               Case mSTEPPING
                  ReadScript
               Case mPAUSED
                     cmdScript(mSTEP).Caption = "Pause"
                     mnCurrentMode = mRECORDING
                     AdjustedStart& = Timer - mlPauseTime + mlScriptStartTime
                     mlScriptStartTime = AdjustedStart&
                     tmrScript.ENABLED = True
            End Select
         Else
            Select Case mnCurrentMode
               Case mIDLE
                  frmScriptInfo.picMasterStat.Line (0, 0)-(mnNumMasterLines, 75), &H8000000F, BF
                  If mnMasterRun Then   'And Not mnScriptLoop
                     mnCurrentMode = mSTEPPING
                     mnReadMaster = True
                     ReadMaster mnCurrentMode
                     LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, Args)
                     If mnMasterLine > 0 Then MasterType% = ParseScript(mMASTER, mnMasterLine - 1, MasterArgs)
                     If MasterType% = mCOMMAND Then
                        'get the board name from the master script and insert into subscript args
                        If ((VarType(MasterArgs) And vbArray) = vbArray) And ((VarType(Args) And vbArray) = vbArray) Then
                           DevArg$ = Trim(MasterArgs(3))
                           If (Args(1) = mOPENFORM) And Not (DevArg$ = "0") Then
                              Args(3) = DevArg$
                           End If
                        End If
                     End If
                     If LineType% = mCOMMAND Then UpdateScriptStatus Args
                  Else
                     If Not gnScriptRun Or (mnScriptLoop = 0) Then
                        If Not OpenScript() Then Exit Sub
                     End If
                     LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, Args)
                     If LineType% = mCOMMAND Then
                        UpdateScriptStatus Args
                     End If
                  End If
                  mnCurrentMode = mRUNNING
                  mnStepping = True
               Case mRECORDING
                  If gnScriptSave Then
                     DelayVal = InputBox("Number of seconds to delay:  ", "Add Delay to Script", "5")
                     If Not IsNull(DelayVal) Then Delay$ = Format$(DelayVal, "0")
                     'Print #2, "0" & Delay$ & ",3000,,,,,,,,,,,,,"
                     Print #2, "0, 3000, 0, " & Delay$ & ",,,,,,,,,,,"
                  Else
                     AppendRecord
                  End If
               Case mRUNNING
                  ReadScript
                  LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, Args)
                  If LineType% = mCOMMAND Then
                     UpdateScriptStatus Args
                  End If
               Case mPAUSED
                     cmdScript(mSTEP).Caption = "Pause"
                     mnCurrentMode = mRECORDING
                     AdjustedStart& = Timer - mlPauseTime + mlScriptStartTime
                     mlScriptStartTime = AdjustedStart&
                     tmrScript.ENABLED = True
            End Select
         End If
      Case mSTOP
         mnAbortScript = True
         gnScriptSave = False
         CloseScript
         For i% = Forms.Count - 1 To 0 Step -1
            If Not Forms(i%).Name = "frmScriptInfo" Then Unload Forms(i%)
         Next i%
         DoEvents
         mnAbortScript = False
         mnStepping = False
         mnMasterLine = 0
         mnCurConditional = 0
         'mlForCount = 0
         mnNumScriptVars = 0
         mnForNest = -1
         mnNumStrings = -1
         cmdScript(mRUN).ENABLED = True
         cmdScript(mSTOP).ENABLED = False
         Me.lblLoopStat.Visible = False
         Me.lblLoopStat.Caption = ""
         Me.lblScriptRate.Visible = False
         Me.lblScriptRate.Caption = ""
      Case mPREVIOUS
         If Not mnStepping Then
            tmrScript.ENABLED = False
            mnStepping = True
            cmdScript(mPREVIOUS).Caption = "&<<"
            cmdScript(mPREVIOUS).ENABLED = True
            mnScriptMode = mPAUSED
            cmdScript(mRUN).ENABLED = True
            gnScriptPaused = True
            gnScriptPaused = True
         Else
            mnScriptLine = mnScriptLine - 1
         End If
      Case mNEXT
         'If mnCurrentMode = mREADAHEAD Then
         mnJustReading = True
         If mnMasterLoop Then
            If Not (mnCurrentMode = mREADAHEAD) Then 'mnReadMaster
               ReadMaster mnCurrentMode
            ElseIf mnMasterRun Then
               cmdScript(mPREVIOUS).ENABLED = False
               gnScriptSave = False
               gnScriptRun = False
               mnStepping = False
               cmdScript(mSTEP).Caption = "&Step"
               cmdScript(mSTOP).ENABLED = False
               lblScriptStatus.ForeColor = &HFF0000
               msScriptStatus = "Script Status:  Idle ("
               lblScriptStatus.Caption = msScriptStatus & msMasterPath & ")"
               Caption = "Current Script: " & msMasterFile & msVerification
               cmdScript(mSTEP).ENABLED = True
            End If
         Else
            If Not gnScriptRun Then
               If Not OpenScript() Then Exit Sub
            End If
            mnCurrentMode = mREADAHEAD
            ReadScript
         End If
         'If Not mnReadMaster Then
         '   mnJustReading = True
         '   ReadScript
            mnJustReading = False
         'Else
            cmdScript(mPREVIOUS).Caption = "&<<"
            cmdScript(mPREVIOUS).ENABLED = True
         'End If
      Case mUNITEST
         mfmUniTest.picCommands.Visible = True
   End Select
   If mnStepping Then cmdScript(mPREVIOUS).ENABLED = (mnScriptLine > 0)

End Sub

Private Sub cmdScript_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   If (Button = 2) Then
      Select Case Index
         Case mPREVIOUS, mNEXT
            SLine$ = InputBox$("Enter script line to jump to:", "Jump To Script Line", Format$(mnScriptLine, "0"))
            If Len(SLine$) Then
               NewLine% = Val(SLine$)
               mnScriptLine = NewLine%
            End If
         Case mUNITEST
            If Not mnHeaderLines < 0 Then
               frmScriptFiles.txtHeader.Visible = True
               frmScriptFiles.cmdOK.ENABLED = True
               frmScriptFiles.cmdCancel.Visible = False
               frmScriptFiles.txtHeader.Text = msHeader
               frmScriptFiles.Show 1
               Unload frmScriptFiles
            Else
               MsgBox "No header currently exists. Is master script loaded?", , "No Header Available"
            End If
      End Select
   End If

End Sub

Private Sub ConfigAnalogIn(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   Select Case FuncID%
      Case AIn '1
         RunAIn FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AInScan, USBAInScan   '2
         mlStaticOpt = FormRef.GetStaticOption
         RunAInScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case ALoadQueue '3
         RunALoadQueue FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case APretrig '4
         mlStaticOpt = FormRef.GetStaticOption
         RunAPretrig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case ATrig '5
         RunATrig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case FileAInScan '6
         mlStaticOpt = FormRef.GetStaticOption
         RunFileAInScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case FileGetInfo '7
         'RunFileGetInfo FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case FilePretrig '8
         mlStaticOpt = FormRef.GetStaticOption
         RunFilePretrig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case TIn '9
         mlStaticOpt = FormRef.GetStaticOption
         RunTIn FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case TInScan   '10
         mlStaticOpt = FormRef.GetStaticOption
         RunTInScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AOut   '11
         RunAOut FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AConvertData '29
         'RunConvertData FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AConvertPretrigData '31
         RunConvertPretrigData FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SetTrigger '49
         RunSetTrigger FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetStatus '58
         Status% = RunGetStatus(FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
      Case StopBackground  '59
         RunStopBackground FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case EnableEvent  '71
         RunAIEnableEvent FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DisableEvent '72
         RunAIDisableEvent FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DaqInScan    '101
         mlStaticOpt = FormRef.GetStaticOption
         RunDaqInScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AIn32 '114
         mlStaticOpt = FormRef.GetStaticOption
         RunAIn32 FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case VIn32 '116
         mlStaticOpt = FormRef.GetStaticOption
         RunVIn32 FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AInputMode   '131
         RunAInputMode FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AChanInputMode   '132
         RunAChanInputMode FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
         
      Case SSetBoardName   '2000
         RunSetBoard FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SContPlot       '2002
         RunPlotContin FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SConvData       '2003
         RunConvert FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SConvPT         '2004
         RunConvertPT FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SCalCheck       '2008
         RunCalMode FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SCountSet       '2009
         RunCountSet FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SAddPTBuf       '2010
         RunSetPTBuf FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SGetTC          '2020
         RunGetTC FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SToEng          '2021
         RunUseEngUnits FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetPlotOpts    '2012
         RunSetPlotOpts FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SBufInfo        '2013
         RunBufInfo FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetPlotChan    '2014
         RunSetPlotChan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SNextBlock      '2015
         RunPlotBlock FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      'Case SSetBlock       '2016
      '   RunSetBlock FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetResolution       '2018
         RunSetRes FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SShowText       '2019
         RunShowText FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SPlotType       '2022
         RunSetPlotType FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SCalcNoise      '2023
         RunSetCalcNoise FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SGetStatus      '3002
         RunASetGetStatus FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SEvalEnable     '2031
         RunEvalEnable FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select

End Sub

Private Sub ConfigAnalogOut(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   Select Case FuncID%
      Case ALoadQueue '3
         RunALoadQueue FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AOut '11
         RunAOut FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case AOutScan   '12
         mlStaticOpt = FormRef.GetStaticOption
         RunAOutScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetStatus '58
         Status% = RunGetStatus(FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
      Case StopBackground  '59
         RunStopBackground FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case EnableEvent  '71
         RunAIEnableEvent FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DisableEvent '72
         RunAIDisableEvent FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DaqOutScan    '103
         mlStaticOpt = FormRef.GetStaticOption
         RunDaqOutScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case VOut '108
         mlStaticOpt = FormRef.GetStaticOption
         RunVOut FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBoardName   '2000
         RunSetBoard FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetData  '2005
         RunAOSetData FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetAmplitude   '2006
         RunAOSetAmpl FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetOffset   '2007
         RunAOSetOS FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetDevName  '2011
         RunAOSetDev FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SBufInfo        '2013
         RunBufInfo FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SGetStatus   '3002
         RunASetGetStatus FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select

End Sub

Private Sub ConfigConfig(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   Select Case FuncID%
      Case LoadConfig   '52
         'RunLoadConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetConfig    '54
         RunGetConfig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SetConfig    '55
         RunSetConfig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SelectSignal '77
         RunSelectSignal FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetSignal    '78
         'RunGetSignal FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DaqSetTrigger '102
         RunDaqSetTrig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DaqSetSetpoints    '109
         RunSetPoints FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      'Case SSetBoardName   '2000
      '   RunSetBoard FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select

End Sub

Private Sub ConfigCounter(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   Select Case FuncID%
      Case C8536Init    '13
         'RunC8536Init FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case C9513Init    '14
         RunC9513Init FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case C8254Config  '15
         RunC8254Config FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case C8536Config  '16
         'RunC8536Config FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case C9513Config  '17
         RunC9513Config FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CLoad        '18
         RunCLoad FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CIn          '19
         RunCIn FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CStoreOnInt  '20
         RunCStoreOnInt FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CFreqIn      '21
         RunCFreqIn FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetStatus    '58
         Status% = RunGetStatus(FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
      Case C7266Config  '67
         RunC7266Config FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CIn32        '68
         RunCIn32 FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CLoad32      '69
         RunCLoad32 FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CStatus      '70
         'RunCStatus FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CInScan      '94
         mlStaticOpt = FormRef.GetStaticOption
         RunCInScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CConfigScan  '95
         RunCConfigScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CClear       '96
         RunCClear FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case TimerOutStart  '97
         RunTimerOutStart FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case TimerOutStop  '98
         RunTimerOutStop FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      'Case SSetBoardName   '2000
      '   RunCtrSetBoard FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case CIn64     '122
         RunCIn64 FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SBufInfo        '2013
         RunBufInfo FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SShowText       '2019
         RunShowText FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select

End Sub

Private Sub ConfigDigitalIO(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   Select Case FuncID%
      Case DConfigPort  '22
         RunDConfig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DBitIn       '23
         RunDBitIn FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DIn          '24
         RunDIn FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DInScan      '25
         mlStaticOpt = FormRef.GetStaticOption
         RunDInScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DOut         '27
         RunDOut FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DOutScan     '28
         mlStaticOpt = FormRef.GetStaticOption
         RunDOutScan FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetStatus '58
         Status% = RunGetStatus(FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
      Case StopBackground  '59
         RunStopBackground FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case EnableEvent  '71
         RunDIEnableEvent FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DisableEvent '72
         RunDIDisableEvent FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DConfigBit   '76
         RunDConfigBit FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBoardName   '2000
         RunSetBoard FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SContPlot       '2002
         RunPlotContin FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetData  '2005
         RunDOSetData FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetAmplitude   '2006
         RunDOSetAmpl FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetOffset   '2007
         RunDOSetOS FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      'Case SSetBlock       '2016
      '   RunSetBlock FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SGetFormRef     '2041
         RunFormReference FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSelPortRange   '2042
         RunPortSelect FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetPortDirection  '2043
         RunPortConfig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SReadPortRange  '2044
         RunReadPortRange FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SWritePortRange '2045
         RunWritePortRange FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSelBitRange   '2046
         RunBitSelect FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SReadBitRange   '2047
         RunReadBitRange FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SWriteBitRange  '2048
         RunWriteBitRange FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBitDirection  '2052
         RunBitConfig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select

End Sub

Private Sub ConfigDigitalIn(FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   Dim FormRef As Form
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
   'replaced with ConfigDigitalIO
   On Error GoTo DINotOpen

   NumForms% = UBound(frmNewDigital)
   'configure per parameters
   Instance% = Val(Right$(FormID$, 2))
   If NumForms% < Instance% Then GoTo DINotOpen
   GotReference% = GetFormReference(FormID$, FormRef)
   If Not GotReference% Then Exit Sub

   Select Case FuncID%
      Case DConfigPort  '22
         'RunDConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DBitIn       '23
         'RunDBitIn FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DIn          '24
         'RunDIn FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DInScan      '25
         'RunDInScan FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DBitOut      '26
         'RunDBitOut FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DOut         '27
         'RunDOut FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DOutScan     '28
         'RunDOutScan FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case StopBackground  '59
         'RunStopBackground FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case EnableEvent  '71
         'RunDIEnableEvent FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DisableEvent '72
         'RunDIDisableEvent FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DisableEvent '72
         'RunDIDisableEvent FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DConfigBit   '76
         'RunDConfigBit FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBoardName   '2000
         'RunDIOSetBoard FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SContPlot       '2002
         RunPlotContin FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SGetFormRef     '2041
         'RunFormReference FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSelPortRange   '2042
         'RunPortSelect FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetPortDirection  '2043
         'RunPortConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SReadPortRange  '2044
         'RunReadPortRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SWritePortRange '2045
         'RunWritePortRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSelBitRange   '2046
         'RunBitSelect FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SReadBitRange   '2047
         'RunReadBitRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SWriteBitRange  '2048
         'RunWriteBitRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBitDirection  '2052
         'RunBitConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select
   Exit Sub

DINotOpen:
   If Not mnDigitalInError Then MsgBox "This script is calling a Digital Input form" & _
   " (" & FormID$ & ") " & "that isn't open. It may not work properly.", vbOKOnly, "Missing Form"
   mnDigitalInError = True
   Exit Sub

End Sub

Private Sub ConfigDigitalOut(FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   On Error GoTo DONotOpen
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
   'replaced with ConfigDigitalIO
   NumForms% = UBound(frmNewDigital)
   'configure per parameters
   Instance% = Val(Right$(FormID$, 2))
   If NumForms% < Instance% Then GoTo DONotOpen

   Select Case FuncID%
      Case DConfigPort  '22
         'RunDConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DBitOut      '26
         'RunDBitOut FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DOut         '27
         'RunDOut FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DOutScan     '28
         'RunDOutScan FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GetStatus '58
         'Status% = RunGetStatus(frmNewDigital(Instance%), A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
      Case StopBackground  '59
         'RunStopBackground frmNewDigital(Instance%), A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case DConfigBit   '76
         'RunDConfigBit FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBoardName   '2000
         'RunDIOSetBoard FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetData  '2005
         'RunDOSetData FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetAmplitude   '2006
         'RunDOSetAmpl FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetOffset   '2007
         'RunDOSetOS FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SGetFormRef     '2041
         'RunFormReference FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSelPortRange   '2042
         'RunPortSelect FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetPortDirection  '2043
         'RunPortConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SReadPortRange  '2044
         'RunReadPortRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SWritePortRange '2045
         'RunWritePortRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSelBitRange   '2046
         'RunBitSelect FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SReadBitRange   '2047
         'RunReadBitRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SWriteBitRange  '2048
         'RunWriteBitRange FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case SSetBitDirection  '2052
         'RunBitConfig FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select
   Exit Sub

DONotOpen:
   If Not mnDigitalOutError Then MsgBox "This script is calling a Digital Output form" & _
   " (" & FormID$ & ") " & "that isn't open. It may not work properly.", vbOKOnly, "Missing Form"
   mnDigitalOutError = True
   Exit Sub

End Sub

Private Sub ConfigGPIB(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   OptCondition = Trim(A5)
   If Not ((OptCondition = True) Or (Trim(OptCondition) = "")) Then Exit Sub
   Device$ = Trim(A2)
   If Not (Device$ = "All devices") Then
      FormRef.cmdConfigure.Caption = "D" & Device$
      FormRef.cmdConfigure = True
   End If

   If Not gnScriptRun Then Exit Sub
   Select Case FuncID%
      Case GPSend  '202
         RunGPIBWrite FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GPReceive      '203
         RunGPIBRead FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GPTrigger
         RunGPIBTrig FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GPDevClear         '205
         RunGPDevClear FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GPSelDevClear     '209
         RunGPSelDevClear FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GPIBSre     '210
         RunGPRen FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case GPIBReturn
         A3 = frmGPIBCtl.GetReturnVal
   End Select

End Sub

Private Sub ConfigUtils(FormRef As Form, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   Select Case FuncID%
      Case ToEngUnits   '32
         'RunToEngUnits FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case FromEngUnits '33
         'RunFromEngUnits FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case FileRead     '34
         'RunFileRead FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case MemRead      '37
         'RunMemRead FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case MemWrite     '38
         'RunMemWrite FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case InByte       '45
         RunInByte FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case OutByte      '46
         RunOutByte FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case InWord       '47
         RunInWord FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case OutWord      '48
         RunOutWord FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Case StopBackground  '59
         RunStopBackground FormRef, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      'Case SSetBoardName   '2000
      '   RunUtilSetBoard FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   End Select

End Sub

Private Sub Form_Load()

   'to do - re-enable when needed
   'mnuLibrary(NETLIB).Checked = gnThreading
   mnHeaderLines = -1
   gnUniScript = True
   Randomize

   'get scripting and master directories if stored
   'check registry first - then ini if not installed
   lpFileName$ = "UniTest.ini"

   CurGroupKey$ = "SOFTWARE\Measurement Computing\Universal Test Suite"
   KeyName$ = "ScriptPath"
   ProgExists% = GetRegGroup(HKEY_LOCAL_MACHINE, CurGroupKey$, hProgResult&)
   ScriptRegistered% = GetKeyValue(hProgResult&, KeyName$, KeyVal$)
   If ScriptRegistered% Then
      MasterDir$ = KeyVal$
      KeyName$ = "ScriptStorage"
      YN% = GetKeyValue(hProgResult&, KeyName$, StoreVal$)
      ScriptDir$ = MasterDir$ & StoreVal$
   Else
      lpApplicationName$ = "ScriptDir"
      lpKeyName$ = "ScriptStorage"
      nSize% = 256
      lpReturnedString$ = Space$(nSize%)
      lpDefault$ = Environ("UTSDir")
      x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
      ScriptDir$ = Left$(lpReturnedString$, x%)
      lpKeyName$ = "MasterStorage"
      nSize% = 256
      lpReturnedString$ = Space$(nSize%)
      lpDefault$ = Environ("UTBDir")
      x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
      MasterDir$ = Left$(lpReturnedString$, x%)
   End If
   
   If Len(ScriptDir$) Then msScriptDir = ScriptDir$
   If Len(MasterDir$) Then msBoardDir = MasterDir$

   If gnShowComments Then
      mnuScriptInf.Checked = True
      frmScriptInfo.Visible = True
   End If
   If gnPrintEval Then
      Me.mnuEvalPrint.Checked = True
      frmScriptInfo.Visible = True
   End If
   If gnPrintEvalAll Then
      Me.mnuPrintAll.Checked = True
      frmScriptInfo.Visible = True
   End If
   UpdateEvalStatus
   mnuEvalPause.Checked = gnPauseEval

   mfmUniTest.cmdUtils.Caption = "&UniScript"
   mfmUniTest.picCommands.Visible = False
   Me.Left = 0
   Me.Top = mfmUniTest.ScaleHeight - Me.Height
   Me.Width = mfmUniTest.ScaleWidth
   mlScriptStartTime = -1
   If (Not (Command = "")) And (msScriptPath = "") Then
      msScriptPath = Command
      'MsgBox Command, vbInformation, "Command"
      If (Left$(msScriptPath, 1) = Chr$(34)) Then msScriptPath = Mid$(msScriptPath, 2)
      If (Right$(msScriptPath, 1) = Chr$(34)) Then msScriptPath = Left$(msScriptPath, Len(msScriptPath) - 1)
      Extension$ = LCase$(Right$(msScriptPath, 4))
      If (Extension$ = ".utm") Or (Extension$ = ".uss") Then
         For CharPos% = Len(msScriptPath) To 1 Step -1
            If Mid$(msScriptPath, CharPos%, 1) = "\" Then Exit For
         Next CharPos%
         mnMasterRun = True
         msMasterPath = Left$(msScriptPath, CharPos% - 1)
         If (Extension$ = ".uss") Then
            msMasterFile = "startup.utm"
            PathArray = Split(msMasterPath, "\Master Scripts\")
            msDefaultPath = PathArray(0)
            msMasterSelection = PathArray(1)
            msMasterPath = msDefaultPath & "\" & msMasterFile
            mnStartCommand = True
         Else
            msMasterFile = Right$(msScriptPath, Len(msScriptPath) - CharPos%) & ":  "
            For CharPos% = 1 To Len(msScriptPath)
               If Mid$(msScriptPath, CharPos%, 1) = ":" Then Exit For
            Next CharPos%
            msDefaultPath = msMasterPath
            msMasterPath = msScriptPath
            mnStartCommand = False
         End If
         msDefaultDrive = Left$(msScriptPath, CharPos% - 1)
         'msScriptPath = msMasterPath & "\" & frmScriptFiles.txtPattern.Text
         If Not OpenMasterScript() Then Exit Sub
         If mnStartupScript Then
            CloseScript
            mnStartupScript = False
            If Not OpenMasterScript() Then Exit Sub
         End If
      Else
         If Not OpenScript() Then Exit Sub
         For CharPos% = Len(msScriptPath) To 1 Step -1
            If Mid$(msScriptPath, CharPos%, 1) = "\" Then Exit For
         Next CharPos%
         mnMasterRun = False
         msScriptFile = Right$(msScriptPath, Len(msScriptPath) - CharPos%)
      End If
      Caption = "Current Script: " & msMasterFile & msScriptFile
   ElseIf Not (msBoardDir = "") Then
      For CharPos% = Len(msBoardDir) To 1 Step -1
         If Mid$(msBoardDir, CharPos%, 1) = "\" Then Exit For
      Next CharPos%
      msMasterPath = Left$(msBoardDir, CharPos% - 1)
      For CharPos% = 1 To Len(msBoardDir)
         If Mid$(msBoardDir, CharPos%, 1) = ":" Then Exit For
      Next CharPos%
      msDefaultDrive = Left$(msBoardDir, CharPos% - 1)
      msDefaultPath = msMasterPath
   End If

   'get global guardbands
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "frmScriptInfo"
   nSize% = 256
   lpDefault$ = ""
   
   lpKeyName$ = "AmpGuard"
   lpReturnedString$ = Space$(nSize%)
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
      lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   Amplitude$ = Left$(lpReturnedString$, x%)
   mlAmplGB = Val(Amplitude$)
   SetAmpGB mlAmplGB

   lpKeyName$ = "RateGuard"
   lpReturnedString$ = Space$(nSize%)
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
      lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   RateGuard$ = Left$(lpReturnedString$, x%)
   mlRateGB = Val(RateGuard$)
   SetRateGB mlRateGB

   lpKeyName$ = "MvgAvg"
   lpReturnedString$ = Space$(nSize%)
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
      lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   MovingAvg$ = Left$(lpReturnedString$, x%)
   mlMvgAvgGB = Val(MovingAvg$)
   SetMvgAvgGB mlMvgAvgGB
   
   lpKeyName$ = "SourceOffset"
   nSize% = 256
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = Environ("0")
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
      lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   SourceOS$ = Left$(lpReturnedString$, x%)
   mfSrcOffset = Val(SourceOS$)
   SetOffsetTweak mfSrcOffset

   lpKeyName$ = "ParamRevision"
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "0.0.0"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
      lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   msLatestRev = Left$(lpReturnedString$, x%)
   
   mnForNest = -1
   mnNumStrings = -1
   ReDim maForNest(5, 0)
   'mnNumScriptVars = 0
   'ReDim ScriptVars(1, 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If mnAbortScript Then Cancel = True

End Sub

Private Sub Form_Resize()

   lblScriptStatus.Width = Me.Width
   PicDiv! = Me.ScaleWidth / 8
   For i% = 0 To 7
      cmdScript(i%).Width = PicDiv!
      cmdScript(i%).Left = PicDiv! * i%
      cmdScript(i%).Top = Me.ScaleHeight - cmdScript(i%).Height
   Next
   Me.lblLoopStat.Left = Me.lblScriptStatus.Width - Me.lblLoopStat.Width
   Me.lblScriptRate.Left = Me.lblScriptStatus.Width - Me.lblScriptRate.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

   gnScriptSave = False
   gnScriptRun = False
   Unload frmScriptInfo

   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "frmScriptInfo"
   lpKeyName$ = "Comments"
   lpString$ = "Disabled"
   If gnShowComments Then lpString$ = "Enabled"
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpKeyName$ = "PrintEval"
   lpString$ = "Disabled"
   If gnPrintEval Then lpString$ = "Enabled"
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpKeyName$ = "PauseEval"
   lpString$ = "Disabled"
   If gnPauseEval Then lpString$ = "Enabled"
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpKeyName$ = "PrintAllEval"
   lpString$ = "Disabled"
   If gnPrintEvalAll Then lpString$ = "Enabled"
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
   
   lpKeyName$ = "RateGuard"
   lpString$ = Format(mlRateGB, "0")
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
   
   lpKeyName$ = "AmpGuard"
   lpString$ = Format(mlAmplGB, "0")
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpKeyName$ = "MvgAvg"
   lpString$ = Format(mlMvgAvgGB, "0")
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
   
   lpKeyName$ = "SourceOffset"
   lpString$ = Format(mfSrcOffset, "0.0######")
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpKeyName$ = "ParamRevision"
   lpString$ = msLatestRev
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

End Sub

Private Sub mnuEvalPause_Click()

   Me.mnuEvalPause.Checked = Not mnuEvalPause.Checked
   gnPauseEval = mnuEvalPause.Checked
   UpdateEvalStatus
   
End Sub

Private Sub mnuEvalPrint_Click()

   Me.mnuEvalPrint.Checked = Not mnuEvalPrint.Checked
   gnPrintEval = Me.mnuEvalPrint.Checked
   frmScriptInfo.Visible = gnShowComments Or gnPrintEval
   UpdateEvalStatus
   
End Sub

Private Sub UpdateEvalStatus()

   If gnPauseEval Then PauseStat$ = "pause script on error"
   If gnPrintEval Then PrintStat$ = "print script errors"
   EvalStat$ = "     [ Evaluation status: On  ("
   TermStat$ = ") ]"
   If PrintStat$ & PauseStat$ = "" Then
      EvalStat$ = "   [ Evaluation status: Off ]"
      TermStat$ = ""
   End If
   If Not PrintStat$ = "" Then
      If Not PauseStat$ = "" Then SepStat$ = ",  "
   End If
   msVerification = EvalStat$ & PrintStat$ & SepStat$ & PauseStat$ & TermStat$
   If msScriptStatus = "" Then msScriptStatus = "Script Status:  Idle"
   Me.Caption = msScriptStatus & msVerification

End Sub
Private Sub mnuFile_Click(Index As Integer)

   On Error GoTo BadFile

   CloseScript
   mnCurrentMode = mIDLE
   mnHeaderLines = -1
   lblScriptStatus.ForeColor = &HFF0000
   msScriptStatus = "Script Status:  Idle "
   lblScriptStatus.Caption = msScriptStatus
   mlScriptStartTime = -1
   mlScriptTime = 0

   'get most recent directory if stored
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "ScriptDir"
   lpKeyName$ = "LastLoad"
   nSize% = 256
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = ""

   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   LastDir$ = Left$(lpReturnedString$, x%)
   If Len(LastDir$) Then
      For CharPos% = 1 To Len(LastDir$)
         If Mid$(LastDir$, CharPos%, 1) = ":" Then Exit For
      Next CharPos%
      SavedDrive$ = Left$(LastDir$, CharPos% - 1)
      SavedPath$ = LastDir$
   End If

   Select Case Index
      Case 0   'create new script file
         frmScriptFiles.Caption = "Open New Script"
         frmScriptFiles.cmbPattern.ENABLED = False
         frmScriptFiles.cmbPattern.AddItem "*.uts (Subscript Files)"
         frmScriptFiles.cmbPattern.ListIndex = 0
      Case 1   'run existing
         frmScriptFiles.Caption = "Open Existing Script"
         frmScriptFiles.cmbPattern.ENABLED = True
         frmScriptFiles.cmbPattern.AddItem "*.uts (Subscript Files)"
         frmScriptFiles.cmbPattern.AddItem "*.utm (Master Scripts)"
         frmScriptFiles.cmbPattern.AddItem "*.uts;*.utm (All Scripts)"
         frmScriptFiles.cmbPattern.ListIndex = 2
      Case 2   'append to existing
         frmScriptFiles.Caption = "Append Existing Script"
         frmScriptFiles.cmbPattern.AddItem "*.uts (Subscript Files)"
         frmScriptFiles.cmbPattern.ENABLED = True
         frmScriptFiles.cmbPattern.ListIndex = 0
   End Select
   
   If Len(LastDir$) Then
      frmScriptFiles.Drive1.Drive = SavedDrive$
      frmScriptFiles.Dir1.Path = SavedPath$
   ElseIf Len(msDefaultPath) Then
      frmScriptFiles.Drive1.Drive = msDefaultDrive
      frmScriptFiles.Dir1.Path = msDefaultPath
   ElseIf Len(msBoardDir) Then
      frmScriptFiles.Drive1.Drive = msBoardDir
   End If
   frmScriptFiles.Show 1

   If frmScriptFiles.txtPattern.Text <> "" Then
      msDefaultDrive = frmScriptFiles.Drive1.Drive
      msDefaultPath = frmScriptFiles.Dir1.Path
      msScriptPath = frmScriptFiles.File1.Path & "\" & frmScriptFiles.txtPattern.Text
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, msDefaultPath, lpFileName$)
      
      Select Case LCase$(Right$(msScriptPath, 4))
         Case ".uts"
            mnMasterRun = False
            msMasterFile = ""
            msScriptFile = frmScriptFiles.txtPattern.Text
         Case ".utm"
            msScriptFile = ""
            mnMasterRun = True
            msMasterPath = msScriptPath
            msMasterFile = frmScriptFiles.txtPattern.Text & ":  "
         Case Else
            Unload frmScriptFiles
            Exit Sub
      End Select
      Caption = "Current Script: " & msMasterFile & msScriptFile & msScriptStatus & msVerification
   Else
      Unload frmScriptFiles
      Exit Sub
   End If
   
   Select Case Index
      Case 0   'create new script file
         StartRecord
         mnStepping = True
      Case 1   'run existing
         If mnMasterRun Then
            frmScriptInfo.txtScriptInfo.Text = ""
            msMasterPath = msScriptPath
            If Not OpenMasterScript() Then Exit Sub
            If mnStartupScript Then
               CloseScript
               mnStartupScript = False
               If Not OpenMasterScript() Then
                  Unload frmScriptFiles
                  Exit Sub
               End If
            End If
         Else
            If Not OpenScript() Then Exit Sub
         End If
      Case 2   'append to existing
         AppendRecord
         mnStepping = True
   End Select
   Unload frmScriptFiles
   Exit Sub
   
BadFile:
   Resume Next

End Sub

Private Sub mnuGuardbands_Click()

   frmGuardBands.SetRateGB mlRateGB
   frmGuardBands.SetAmplGB mlAmplGB
   frmGuardBands.SetOffsetGB mfSrcOffset
   frmGuardBands.SetAvgGB mlMvgAvgGB
   frmGuardBands.Show 1
   
   mlRateGB = GetRateGB()
   mlAmplGB = GetAmpGB()
   mlMvgAvgGB = GetMvgAvgGB()
   mfSrcOffset = GetOffsetTweak()

   If 0 Then
      'moved this stuff to the Form "frmGuardBands"
      'so the following is no longer used
      If Not (frmGuardBands.cmdDone.ENABLED) Then
         Unload frmGuardBands
         Exit Sub
      End If
      If frmGuardBands.cmbApplyTo.ListIndex = 0 Then
         'if global, not board specific tweaks
         mlRateGB = Val(frmGuardBands.txtRate.Text)
         mlAmplGB = Val(frmGuardBands.txtAmpl.Text)
         mlMvgAvgGB = Val(frmGuardBands.txtAverage.Text)
         mfSrcOffset = Val(frmGuardBands.txtOffset.Text)
         
         SetAmpGB mlAmplGB
         SetRateGB mlRateGB
         SetMvgAvgGB mlMvgAvgGB
         SetOffsetTweak mfSrcOffset
      Else
         BoardAmpl& = Val(frmGuardBands.txtAmpl.Text)
         BoardRate& = Val(frmGuardBands.txtRate.Text)
         MovingAvg& = Val(frmGuardBands.txtAverage.Text)
         SimInput% = frmGuardBands.chkSimIn.value
         SimOutput% = frmGuardBands.chkSimOut.value
         ValuesSet% = Not ((BoardAmpl& = 0) And (BoardRate& = 0) And (MovingAvg& = 0))
         ValuesSet% = ValuesSet% Or Not ((SimInput% = 0) And (SimOutput% = 0))
         If Not frmGuardBands.EntryExists And Not ValuesSet% Then
            'if there's no ini entry for this device,
            'don't create one if parameters are zero
            If (BoardAmpl& = 0) And (BoardRate& = 0) And (MovingAvg& = 0) Then
               Unload frmGuardBands
               Exit Sub
            End If
         Else
            lpApplicationName$ = frmGuardBands.cmbApplyTo.Text
            lpFileName$ = "ScriptParams.ini"
            
            lpKeyName$ = "AmpGuard"
            lpString$ = Format(BoardAmpl&, "0")
            x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
            lpKeyName$ = "RateGuard"
            lpString$ = Format(BoardRate&, "0")
            x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
            lpKeyName$ = "MvgAvg"
            lpString$ = Format(MovingAvg&, "0")
            x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
            lpKeyName$ = "SimultaneousOut"
            lpString$ = Format(SimOutput%, "0")
            x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
            lpKeyName$ = "SimultaneousIn"
            lpString$ = Format(SimInput%, "0")
            x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
            
            ResetEvalBoard
         End If
      End If
   End If
   
   Unload frmGuardBands
   
End Sub

Private Sub mnuPrintAll_Click()

   Me.mnuPrintAll.Checked = Not mnuPrintAll.Checked
   gnPrintEvalAll = Me.mnuPrintAll.Checked
   frmScriptInfo.Visible = gnShowComments Or gnPrintEval Or gnPrintEvalAll
   UpdateEvalStatus

End Sub

Private Sub mnuTestOptions_Click()
   
    OpenTestOptionsForm
    
End Sub

Public Function GetTestOptValue(ByVal TestNum As Long) As Integer

   Dim Args(0)
   TestOptName$ = "test" & Format(TestNum, "0")
   Args(0) = TestOptName$
   VarFound% = CheckForVariables(Args)
   TestOptValue% = Val(Args(0))
   GetTestOptValue = TestOptValue%
   
End Function

Public Sub SetTestOptValue(ByVal TestNum As Long, ByVal value As Integer)

   TestOptName$ = "test" & Format(TestNum, "0")
   VarSet% = SetVariable(TestOptName$, value)

End Sub

Private Sub mnuTimer_Click()
      
   frmSetTimer.txtInterval.Text = tmrScript.Interval
   frmSetTimer.chkEnableTimer = Abs(mnuTimer.Checked)
   frmSetTimer.optTimerMode(0).Visible = False
   frmSetTimer.optTimerMode(1).Visible = False
   frmSetTimer.Show 1
   LoopRate& = Val(frmSetTimer.txtInterval.Text)
   mnuTimer.Checked = (frmSetTimer.chkEnableTimer.value = 1)
   Unload frmSetTimer
   If LoopRate& > 0 Then
      tmrScript.Interval = LoopRate&
   Else
      mnuTimer.Checked = False
   End If

End Sub

Private Sub mnuScriptInf_Click()

   mnuScriptInf.Checked = Not mnuScriptInf.Checked
   gnShowComments = mnuScriptInf.Checked
   frmScriptInfo.Visible = gnShowComments Or gnPrintEval Or gnPrintEvalAll
   If frmScriptInfo.Visible Then
      If Not msHeader = "" Then frmScriptInfo.txtScriptInfo.Text = msHeader
   Else
      Unload frmScriptInfo
   End If

End Sub

Private Function OpenMasterScript() As Integer

   On Error GoTo MSOpenError  'Resume Next
   gnScriptRun = True
   mnAbortScript = False
   Open msMasterPath For Input As #4

   If mnHeaderLines < 0 Then ReadHeader
   If mnAbortScript Then
      Close #4
      mnAbortScript = False
      Exit Function
   End If
   Do While Not EOF(4)
      Line Input #4, A1
      If A1 = "" Then A1 = "; "
      ReDim Preserve masMaster(mnNumMasterLines)
      masMaster(mnNumMasterLines) = A1
      mnNumMasterLines = mnNumMasterLines + 1
   Loop
   Close #4
   
   cmdScript(mOPEN).ENABLED = True
   cmdScript(mNEW).ENABLED = True
   cmdScript(mRUN).ENABLED = True
   cmdScript(mSTEP).ENABLED = True
   cmdScript(mSTOP).ENABLED = True
   cmdScript(mNEXT).ENABLED = True
   'cmdScript(mPREVIOUS).Enabled = False
   Caption = "Current Script: " & msMasterFile & msVerification
   msScriptStatus = "Script Status:  Master Script Loaded" & " ("
   lblScriptStatus.Caption = msScriptStatus & msScriptPath & ")"
   frmScriptInfo.picMasterStat.ScaleWidth = mnNumMasterLines
   frmScriptInfo.picMasterStat.Line (0, 0)-(mnNumMasterLines, 75), &H8000000F, BF
   OpenMasterScript = True
   Dim Args(0)
   Args(0) = "criticalwarning"
   VarFound% = CheckForVariables(Args)
   If VarFound% Then
      If Not (Args(0) = "") Then
         MsgBox "Critical warning: " & Args(0), vbCritical, "Scripted Warning"
      End If
   End If
   OpenTestOptionsForm
   Exit Function

MSOpenError:
   MsgBox Error(Err), , "Error Opening Master Script"
   Exit Function
   Resume 0

End Function

Private Function OpenScript() As Integer

   On Error GoTo FileErr   'Resume Next
   mnNumScriptLines = 0
   gnScriptRun = True
   ScriptToOpen$ = msScriptPath
   Open ScriptToOpen$ For Input As #2

   Do While Not EOF(2)
      Line Input #2, A1
      If A1 = "" Then A1 = "; "
      ReDim Preserve masScript(mnNumScriptLines)
      masScript(mnNumScriptLines) = A1
      mnNumScriptLines = mnNumScriptLines + 1
      'If LCase(A1) = "'end" Then Exit Do
   Loop
   Close #2
   If mnNumScriptLines = 0 Then
      MsgBox "The script '" & ScriptToOpen$ & "' has no data to read. (If no path " & _
      "is listed, check for rogue file in " & CurDir() & ".)", vbCritical, "Bad Script File"
      OpenScript = False
      Exit Function
   End If
   
   cmdScript(mOPEN).ENABLED = True
   cmdScript(mNEW).ENABLED = True
   If Not mnCurrentMode = mRUNNING Then cmdScript(mRUN).ENABLED = True
   cmdScript(mSTEP).ENABLED = True
   cmdScript(mSTOP).ENABLED = True
   cmdScript(mNEXT).ENABLED = True
   'cmdScript(mPREVIOUS).Enabled = False
   msScriptStatus = "Script Status:  SubScript Loaded" & " ("
   lblScriptStatus.Caption = msScriptStatus & ScriptToOpen$ & ")"
   mnScriptClosed = False
   OpenScript = True
   frmScriptInfo.picScriptStat.ScaleWidth = mnNumScriptLines
   frmScriptInfo.picScriptStat.Line (0, 0)-(mnNumScriptLines, 75), &H8000000F, BF
   Exit Function

FileErr:
   
   MasterPath$ = msBoardDir
   If (Attempts% = 3) Or (Left(ScriptToOpen$, 3) = "..\") Then
      MasterPath$ = msDefaultPath
   End If
   ScriptToOpen$ = LocateScriptFile(ScriptToOpen$, msScriptDir, MasterPath$, Attempts%)
   If ScriptToOpen$ = "" Then
      If msScriptDir = "" Or msMasterPath = "" Then Hint$ = _
      " Try setting the Script or Master Directories under the File menu."
      MsgBox "File not found in the following locations:" & vbCrLf & vbCrLf & _
      PathsChecked$ & vbCrLf & Hint$, , "Error Opening Script File"
      gnErrFlag = True
      Exit Function
   Else
      PathsChecked$ = PathsChecked$ & ScriptToOpen$ & vbCrLf
      Resume 0
   End If
   
   If (Err = 53) Or (Err = 76) Then
      'file not found or path not found
      'search alternative locations
      Select Case attempt%
         Case 0
            ScriptToOpen$ = msBoardDir & msScriptPath
         Case 1
            ScriptToOpen$ = msScriptDir & msScriptPath
         Case 2
            'check location of master script
            NumLocs& = FindInString(msMasterPath, "\", Locs)
            If NumLocs& > 0 Then ScriptToOpen$ = Left(msMasterPath, Locs(NumLocs&)) & msScriptPath
         Case Else
            MsgBox Error$(Err) & vbCrLf & "Make sure the paths listed below are valid." & _
            "Set paths using the 'File | Scrip Directory' and 'File | Master Directory' menus." & _
            vbCrLf & vbCrLf & msScriptPath & vbCrLf & msScriptDir & msScriptPath & vbCrLf & _
            msBoardDir & msScriptPath, , "Script File Not Found"
            mnScriptClosed = True
            Exit Function
      End Select
      attempt% = attempt% + 1
      Resume 0
   Else
      MsgBox Error$(Err), vbOKOnly, "Error Opening Subscript"
      Exit Function
   End If

End Function

Private Function OpenLogFile(Filename As String) As Integer

   On Error GoTo LogFileErr   'Resume Next
   If mnLogOpen Then Close #8
   Open Filename For Output As #8
   mnLogOpen = True

   Exit Function

LogFileErr:
   
   If (Err = 76) Then
      'path not found
      MsgBox Error$(Err) & vbCrLf & "Make sure the paths listed below are valid." & _
      "Set paths using the 'File | Scrip Directory' and 'File | Master Directory' menus." & _
      vbCrLf & vbCrLf & msScriptPath & vbCrLf & msScriptDir & msScriptPath & vbCrLf & _
      msBoardDir & msScriptPath, , "Script File Not Found"
      mnScriptClosed = True
   Else
      MsgBox Error$(Err), vbCritical, "Script File Not Found"
   End If
   Exit Function

End Function

Private Sub ReadHeader()

   Dim HeaderLines() As String
   
   Comment% = True
   msHeader = ""
   mnHeaderLines = 0
   Do
      Line Input #4, A1
      StartChar$ = Left$(A1, 1)
      If Not WaitingEndInit% Then
         If Not (StartChar$ = ";") Then
            If Not A1 = "'Initialize" Then
               ReDim masMaster(0)
               Comment% = False
               masMaster(0) = A1
               mnNumMasterLines = 1
               Exit Do
            Else
               If Not WaitingEndInit% Then
                  'load device specific variables if file exists
                  If Not (msDevParmPath = "") Then
                     InitPath$ = msDevParmPath
                     InitScript$ = ""
                     msDevParmPath = ""
                  Else
                     On Error GoTo CheckSubDirs
                     InitPath$ = msDefaultPath
                     InitScript$ = "\DeviceParams.uts"
                     MasterPath$ = LCase(msMasterPath)
                     If InStr(1, MasterPath$, "startup.utm") > 0 Then
                        InitScript$ = "\StartupParams.uts"
                     End If
                  End If
                  Open InitPath$ & InitScript$ For Input As #10
                  If Not NoInitFileFound% Then
                     Do While Not EOF(10)
                        Line Input #10, InitLine
                        If InitLine = "" Then InitLine = "; "
                        ReDim Preserve masInit(mnNumInitLines)
                        masInit(mnNumInitLines) = InitLine
                        mnNumInitLines = mnNumInitLines + 1
                        If LCase(InitLine) = "'end" Then Exit Do
                     Loop
                     Close #10
                     ReadInitFile
                  End If
               End If
               WaitingEndInit% = True
               InitStart% = NumHeaderLines%
            End If
         Else
            HeaderLine$ = A1
            ReDim Preserve HeaderLines(NumHeaderLines%)
            HeaderLines(NumHeaderLines%) = HeaderLine$
            mnHeaderLines = mnHeaderLines + 1
            NumHeaderLines% = NumHeaderLines% + 1
         End If
      Else
         If Not (StartChar$ = "'") Then
            HeaderLine$ = A1
            ReDim Preserve HeaderLines(NumHeaderLines%)
            HeaderLines(NumHeaderLines%) = HeaderLine$
            NumHeaderLines% = NumHeaderLines% + 1
         Else
            If A1 = "'End Initialize" Then
               'read the next line as start of master
               ReDim masMaster(0)
               Comment% = False
               Line Input #4, A1
               masMaster(0) = A1
               mnNumMasterLines = 1
               Exit Do
            Else
            HeaderLine$ = A1
            ReDim Preserve HeaderLines(NumHeaderLines%)
            HeaderLines(NumHeaderLines%) = HeaderLine$
            NumHeaderLines% = NumHeaderLines% + 1
            End If
         End If
      End If
   Loop While Comment%
   
   'All comment lines are stored,
   'now read init lines (if any)
   If WaitingEndInit% Then
      For LineNumber% = InitStart% To NumHeaderLines% - 1
         CurLine = HeaderLines(LineNumber%)
         If (mnScriptMode = mSTOP) Then Exit For
         ReadMaster 0, CurLine
      Next
   End If
   If Not (mnScriptMode = mSTOP) Then
   For LineNumber% = 0 To mnHeaderLines - 1
      CurLine = HeaderLines(LineNumber%)
      ReadMaster 0, CurLine
      'Reslt% = ParseScript(0, 0, Args)
   Next
   End If
   
   If mnHeaderLines = 0 Then
      msHeader = "No header exists for this script."
   'Else
   '   If gnShowComments Then frmScriptInfo.txtScriptInfo.Text = msHeader
   End If
   
   Exit Sub

CheckSubDirs:
   If DirsDown% = 0 Then
      DirArray = Split(msDefaultPath, "\")
      NumDirs% = UBound(DirArray)
      DirsDown% = NumDirs%
   End If
   SubDir$ = ""
   DirsDown% = DirsDown% - 1
   If DirsDown% > 0 Then
      For DirCount% = 0 To DirsDown%
         SubDir$ = SubDir$ & DirArray(DirCount%)
         If Not DirCount% = DirsDown% Then SubDir$ = SubDir$ & "\"
      Next
      InitPath$ = SubDir$
      Resume 0
   Else
      If CheckedForStartup% Then
         NoInitFileFound% = True
         Resume Next
      Else
         CheckedForStartup% = True
         InitPath$ = msDefaultPath
         InitScript$ = "\StartupParams.uts"
         Resume 0
      End If
   End If
   
End Sub

Private Sub ReadMaster(Index As Integer, Optional LineText As Variant)
   
   Do
      ReadAgain% = False
      TypeOfLine% = mLINETEXT
      If IsMissing(LineText) Then TypeOfLine% = mMASTER
      LineType% = ParseScript(TypeOfLine%, mnMasterLine, Args, LineText)
      StatusLine% = mnMasterLine
      If LineType% = mINVALID Then Exit Do
      
      If TypeOfLine% = mMASTER Then
         mnMasterLine = mnMasterLine + 1
         frmScriptInfo.picMasterStat.Line (0, 0)-(mnMasterLine, 75), &HFF0000, BF
         mnScriptLoop = False
         mnMasterLoop = True
         If mnStepping Then StatusLine% = mnMasterLine
         If mnMasterLine = mnNumMasterLines Then
            mnReadMaster = False
            mnMasterLine = 0
            mnMasterLoop = False
            mnCurConditional = 0
            mnNumScriptVars = 0
            mnForNest = -1
            mnNumStrings = -1
            Me.lblLoopStat.Visible = False
            Me.lblLoopStat.Caption = ""
            Me.lblScriptRate.Visible = False
            Me.lblScriptRate.Caption = ""
            ReDim maForNest(5, 0)
            ReDim ScriptVars(1, 0)
            ReDim maStringList(0)
            
            If Not mlScriptRate = 0 Then
               'in case timer rate has been changed
               'in a conditional and wasn't reset due
               'to user canceling script
               Me.tmrScript.Interval = mlScriptRate
            End If
            If mnLogOpen Then
               Close #8 'weekend
               mnLogOpen = 0
            End If
         End If
      End If
     
      If Not (mnCurConditional = 0) Then
         Select Case mnCurConditional
            Case WAITINGENDIF, WAITINGENDDO, WAITINGENDFOR
               lblScriptStatus.Caption = Format(mnScriptLine, "0") & " No-op"
               If mnReadMaster Then
                  'ReadMaster mnScriptMode
                  ReadAgain% = True
                  Me.tmrScript.ENABLED = True
               End If
               Exit Do
         End Select
      End If
      
      If (LineType% = mCOMMENT) Or (LineType% = mNOOPLINE) Then
         If mnAbortScript Then
            Exit Do
         End If
         If mnReadMaster Then
            'ReadMaster mnScriptMode
            ReadAgain% = True
         Else
            Exit Do
         End If
      End If
   
      If Not ReadAgain% Then
         If (mnScriptMode = mNEXT) Or (mnScriptMode = mPREVIOUS) Then
            UpdateScriptStatus Args
            Exit Do
         End If
         
         FuncID% = Val(Args(0))
         mnDevDupe = Val(Args(1))
         If mnDevDupe = SGetStringFromList Then VariableName$ = LCase(Trim(Args(4)))
         If mnDevDupe = SGetCSVsFromList Then VariableName$ = LCase(Trim(Args(5)))
         If mnDevDupe = SGetFormProps Then VariableName$ = LCase(Trim(Args(4)))
         If Not ((mnDevDupe = SSetVariable) Or (mnDevDupe = SSetVarDefault)) Then _
         FoundVar% = CheckForVariables(Args)
      
         A1 = Trim(Args(2))
         DevName$ = Trim(Args(3))
      
         Select Case FuncID%
            Case 0
               Select Case mnDevDupe
                  Case SShowDiag '2001
                     'user action message box
                     OptCondition = Trim(Args(8))
                     If (OptCondition = True) Or (Trim(OptCondition) = "") Then
                        DiagType% = Val(Args(4))
                        Select Case DiagType%
                           Case 0
                              Title$ = Trim(Args(6))
                              MsgBox Args(3), , Title$
                           Case 1
                              VarName$ = Trim(LCase(Args(5)))
                              Title$ = Trim(Args(6))
                              Default$ = Trim(Args(7))
                              Resp = InputBox(Args(3), Title$, Default$)
                              VarSet% = SetVariable(VarName$, Resp)
                        End Select
                     End If
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SSetBlock       '2016
                     OptCondition = Trim(Args(4))
                     If (OptCondition = True) Or (Trim(OptCondition) = "") Then
                        BlockSize& = Val(Args(3))
                        SetBlockSize BlockSize&, False
                        InitBlock True
                     End If
                     If mnReadMaster Then ReadAgain% = True
                  Case SLoadStringList '2025
                     Filename$ = Trim(Args(3))
                     VarName$ = LCase(Trim(Args(4)))
                     ReadStringList Filename$, VarName$
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SGetStringFromList '2026
                     StringIndex% = Val(Args(3))
                     If StringIndex% > mnNumStrings Then
                        MsgBox "Script requesting string at index " & _
                        Format(StringIndex%, "0") & ". Number of strings loaded = " & _
                        Format(mnNumStrings + 1, "0") & "."
                        gnErrFlag = True
                        Exit Do
                     Else
                        VarValue = maStringList(StringIndex%)
                        VarSet% = SetVariable(VariableName$, VarValue)
                        If mnReadMaster Then
                           'ReadMaster mnScriptMode
                           ReadAgain% = True
                        End If
                     End If
                  Case SSetPlotScaling    '2027
                     ScriptSet = Val(Args(3))
                     SetAutoScale ScriptSet
                     If mnReadMaster Then ReadAgain% = True
                  Case SLoadCSVList '2029
                     Filename$ = Trim(Args(3))
                     VarNameListLen$ = LCase(Trim(Args(4)))
                     VarNameNumVals$ = LCase(Trim(Args(5)))
                     ListName = LCase(Trim(Args(6)))
                     If Filename$ = "" Then
                        If Not (ListName = "") Then TargetName$ = vbCrLf & "Intended target name for list: " & ListName
                        MsgBox "Attempt to load a CSV list without a file name specified." & vbCrLf & "The number of arguments is targeted to '" & VarNameListLen$ & "'." & TargetName$
                        gnErrFlag = True
                        Exit Do
                     Else
                        ReadCSVList Filename$, VarNameListLen$, VarNameNumVals$, ListName
                     End If
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SGetCSVsFromList   '2033
                     Dim FirstList As CParamList
                     StringIndex% = Val(Args(3))
                     VarIndex% = Val(Args(4))
                     ListName = LCase(Trim(Args(6)))
                     If mParamList.Count = 0 Then
                        NumCSVArgs% = -1
                     Else
                        Set FirstList = mParamList.Item(1)
                        NamedList% = FirstList.NamedList
                        If NamedList% And (Not ListName = "") Then
                           Set FirstList = Nothing
                           Set FirstList = mParamList.Item(ListName)
                           NumCSVArgs% = FirstList.ListSize 'mParamList.Item(ListName).ListSize
                        Else
                           NumCSVArgs% = FirstList.ListSize
                        End If
                     End If
                     If StringIndex% > NumCSVArgs% Then
                        MsgBox "Script requesting comma separated values at index " & _
                        Format(StringIndex%, "0") & ". Number of CSVs loaded = " & _
                        Format(NumCSVArgs% + 1, "0") & "."
                        gnErrFlag = True
                        Exit Do
                     Else
                        VarValue = FirstList.GetListItem(StringIndex%, VarIndex%)
                        'VarValue = maCSVList(StringIndex%, VarIndex%)
                        VarSet% = SetVariable(VariableName$, VarValue)
                     End If
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                     Set FirstList = Nothing
                  Case SCopyFile '2036
                     Filename$ = Trim(Args(3))
                     FileDest$ = Trim(Args(4))
                     ScrCopyFile Filename$, FileDest$
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SRunApp '2037
                     CommandLine$ = Trim(Args(3))
                     WaitForClose$ = Trim(Args(4))
                     AppID = Shell(CommandLine$, vbNormalFocus)
                     WaitForAppClose AppID
                     DoEvents
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SResetConfig '2039
                     Filename$ = "cb.cfg"
                     ULStat& = cbSaveConfig(Filename$)
                     DoEvents
                     ULStat& = cbLoadConfig("cb.cfg")
                     DoEvents
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SScriptRate    '3008
                     Interval& = Val(Args(3))
                     frmScript.tmrScript.Interval = Interval&
                     Me.lblScriptRate.Visible = True
                     Me.lblScriptRate.Caption = "Rate: " & Format(Interval&, "0")
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SSetVariable '3009
                     OptCondition = Trim(Args(5))
                     If (OptCondition = True) Or (Trim(OptCondition) = "") Then
                        VariableName$ = Trim(LCase(Args(3)))
                        VarValue = Args(4)
                        'set variable to one of several values separated by ";"
                        NewValue = SelectValFromList(VarValue)
                        If NewValue = "" Then NewValue = VarValue
                        If Not NewValue = "_" Then
                           VarSet% = SetVariable(VariableName$, NewValue)
                        End If
                     End If
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SCloseApp    '3013
                     End
                  Case SSetVarDefault  '3014
                     ReDim VarToCheck(0)
                     VarToCheck(0) = Args(3)
                     FoundVar% = CheckForVariables(VarToCheck)
                     If Not FoundVar% Then
                        VariableName$ = Trim(LCase(Args(3)))
                        VarValue = Args(4)
                        Vars% = SetVariable(VariableName$, VarValue)
                     End If
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SPauseScript  '3015
                     OptCondition = Trim(Args(3))
                     If (OptCondition = True) Or (Trim(OptCondition) = "") Then
                        gnScriptPaused = True
                        'If (mnCurrentMode = mRUNNING) Then cmdScript_Click (5)
                        cmdScript_Click (5)
                     Else
                        If mnReadMaster Then
                           'ReadMaster mnScriptMode
                           ReadAgain% = True
                        End If
                     End If
                  Case SGetParameterString  '3016
                     'FoundVar% = CheckForVariables(Args)
                     FunctionID& = Val(Args(3))
                     ParamNum& = Val(Args(4))
                     ParamVal& = Val(Args(5))
                     VariableName$ = Trim(LCase(Args(6)))
                     VarValue = GetParamString(FunctionID&, ParamNum&, ParamVal&)
                     Vars% = SetVariable(VariableName$, VarValue)
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SPicklist    '3017
                     IndexStr$ = LCase(Trim(Args(3)))
                     IndexArray = Split(IndexStr$, " to ")
                     ListIndex& = Abs(Val(IndexArray(0)))
                     If UBound(IndexArray) > 0 Then
                        LastItem& = Val(IndexArray(1))
                        ListLength& = LastItem& - ListIndex&
                     Else
                        ListLength& = 0
                     End If
                     ParamList$ = Args(4)
                     VariableName$ = Trim(LCase(Args(5)))
                     SizeName$ = Trim(LCase(Args(6)))
                     DefaultValue$ = Trim(LCase(Args(7)))
                     VarValue = GetListItem(ListIndex&, ParamList$, ListSize&, ListLength&, DefaultValue$)
                     Vars% = SetVariable(VariableName$, VarValue)
                     If Not SizeName$ = "" Then Vars% = SetVariable(SizeName$, ListSize&)
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SGenRndVal      '3018
                     VariableName$ = Trim(LCase(Args(4)))
                     ArgString$ = Trim(Args(3))
                     SeedString$ = ParseUnits(ArgString$, UnitType%)
                     VarSeed! = Val(SeedString$)
                     RndSeed! = Rnd(1)
                     If UnitType% = UNITFLOAT Then
                        VarValue = Rnd(RndSeed!) * VarSeed!
                     Else
                        VarValue = CInt(Rnd(RndSeed!) * VarSeed!)
                     End If
                     Vars% = SetVariable(VariableName$, VarValue)
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SPickGroup   '3019
                     GroupIndicator$ = Trim(Args(3))
                     GroupIndex& = Val(Args(4))
                     GroupList$ = Trim(LCase(Args(5)))
                     VariableName$ = Trim(LCase(Args(6)))
                     SizeName$ = Trim(LCase(Args(7)))
                     GroupArray = Split(GroupList$, GroupIndicator$)
                     NumGroups& = UBound(GroupArray)
                     If GroupIndex& > NumGroups& Then GroupIndex& = NumGroups&
                     VarValue = GroupArray(GroupIndex&)
                     If Left$(VarValue, 1) = ";" Then VarValue = Mid$(VarValue, 2)
                     Vars% = SetVariable(VariableName$, VarValue)
                     If Not SizeName$ = "" Then Vars% = SetVariable(SizeName$, NumGroups&)
                     'Else
                     '   MsgBox "The group index specified (" & Format(GroupIndex&, "0") & _
                     '   ") is greater than the number of groups (" & Format(NumGroups&, "0") & _
                     '   ").", vbCritical, "Invalid Group in Script"
                     '   gnErrFlag = True
                     '   Exit Do
                     'End If
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SSetLibType  '3020
                     VarValue = Args(3)
                     If IsNumeric(VarValue) Then LibType% = Val(VarValue)
                     frmMain.SetDefaultLibType LibType%
                     If mnReadMaster Then ReadAgain% = True
                  Case SIsListed    '3022
                     ParamList$ = Args(3)
                     SearchParam$ = Trim(Args(4))
                     ListedBoolName$ = Trim(LCase(Args(5)))
                     IndexVarName$ = Trim(LCase(Args(6)))
                     ListSeparator$ = Trim(LCase(Args(7)))
                     If ListSeparator$ = "" Then ListSeparator$ = ";"
                     Listed% = False
                     ListArray = Split(ParamList$, ListSeparator$)
                     ListSize& = UBound(ListArray)
                     For ListedItem& = 0 To ListSize&
                        If ListArray(ListedItem&) = SearchParam$ Then
                           Listed% = True
                           If Not (IndexVarName$ = "") Then
                              Vars% = SetVariable(IndexVarName$, ListedItem&)
                           End If
                           Exit For
                        End If
                     Next
                     Vars% = SetVariable(ListedBoolName$, Listed%)
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                  Case SStopScript    '3026
                     OptCondition = Trim(Args(3))
                     If (OptCondition = True) Or (Trim(OptCondition) = "") Then
                        cmdScript_Click (mSTOP)
                        Exit Sub
                     End If
                     If mnReadMaster Then ReadAgain% = True
                  Case SSetMCCControl  '3027
                     VarValue = Trim(Args(3))
                     If IsNumeric(VarValue) Then UseMCCCtlr% = Val(VarValue)
                     frmMain.SetMCCControl UseMCCCtlr%
                     If mnReadMaster Then ReadAgain% = True
                  Case SEvalParamRev  '3029
                     Dim CheckFailed As Boolean
                     VarValue = Trim(Args(3))
                     RevRequired$ = VarValue
                     CompCondition$ = Trim(Args(4))
                     Args(0) = "dprevision"
                     VarFound% = CheckForVariables(Args)
                     If VarFound% Then
                        RevInUse$ = Args(0)
                        CheckFailed = CheckParamRevision(RevRequired$, RevInUse$, CompCondition$)
                     End If

                     If CheckFailed Then
                        cmdScript_Click (mSTOP)
                        mnAbortScript = True
                        MsgBox "The parameter file version (" & RevInUse$ & _
                        ") is older than the version required (" & RevRequired$ & _
                        ").", vbCritical, "Parameter File Update Required"
                        Exit Sub
                     End If
                     If mnReadMaster Then ReadAgain% = True
                  Case Else
                     BadLine% = StatusLine% + mnHeaderLines
                     FunctionDef$ = GetScriptString(mnDevDupe)
                     MFile$ = Left(msMasterFile, Len(msMasterFile) - 3)
                     MsgBox "The function " & FunctionDef$ & " in " & MFile$ & _
                     " line " & Format(BadLine%, "0") & " is not valid in a Master Script.", _
                     vbCritical, "Undefined Master Script Command"
               End Select
            Case SLoadSubScript  '4001
               mnMasterLoop = False
               OptCondition = Trim(Args(4))
               If (OptCondition = True) Or (OptCondition = "") Then
                  Dim DevNameArg(0)
                  DevNameArg(0) = "Device"
                  VarFound% = CheckForVariables(DevNameArg)
                  If VarFound% Then CurDev$ = DevNameArg(0)
                  Select Case DevName$
                     Case "AUXDIO", "PROGDIO", "AUXTRIG0", _
                     "AUXTRIG1", "AUXTRIG2", "AUXTRIG3"
                        msDevName = GetGPIBSurrogate(DevName$)
                        If msDevName = "" Then
                           MsgBox "No device assigned to " & DevName$ & " found.", vbCritical, "Control Device Not Found"
                           gnErrFlag = True
                           gnScriptRun = False
                           Exit Sub
                        Else
                           mbSetLibToUL = True
                        End If
                        If msDevName = CurDev$ Then
                           mnDevDupe = 1
                        End If
                     Case Else
                        If FuncID% = SLoadSubScript Then
                           If Not IsNumeric(DevName$) Then msDevName = DevName$
                        End If
                  End Select
                  Filename$ = A1
                  If InStr(1, Filename$, "~") = 0 Then
                     msScriptPath = A1
                     mnCurrentMode = mnScriptMode
                     If Not OpenScript() Then
                        'cmdScript_Click 3
                        gnScriptRun = False
                        Exit Do
                     End If
                     Caption = "Current Script: " & msMasterFile & msScriptPath & msVerification
                  Else
                     'no file is specified - read next master line
                     mnCurrentMode = mnScriptMode
                     If mnReadMaster Then
                        'ReadMaster mnScriptMode
                        ReadAgain% = True
                     End If
                     'Exit Do
                  End If
                  Select Case mnCurrentMode
                     Case mRUNNING
                        tmrScript.ENABLED = True
                     Case mSTEPPING
                        mnCurrentMode = mRUNNING
                        mnStepping = True
                        LineType% = ParseScript(mSUBSCRIPT, 0, Args)
                        If mnMasterLine > 0 Then MasterType% = ParseScript(mMASTER, mnMasterLine - 1, MasterArgs)
                        If MasterType% = mCOMMAND Then
                           If ((VarType(MasterArgs) And vbArray) = vbArray) And ((VarType(Args) And vbArray) = vbArray) Then
                              'get the board name from the master script and insert into subscript args
                              DevArg$ = Trim(MasterArgs(3))
                              If (Args(1) = mOPENFORM) And Not (DevArg$ = "0") Then
                                 Args(3) = DevArg$
                              End If
                           End If
                        End If
                        If LineType% = mCOMMAND Then UpdateScriptStatus Args
                  End Select
               Else
                  If mnReadMaster Then
                     'ReadMaster mnScriptMode
                     ReadAgain% = True
                  End If
               End If
            Case SCloseSubScript '4002
               CloseScript
            Case Else
               BadLine% = StatusLine% + mnHeaderLines
               MFile$ = Left(msMasterFile, Len(msMasterFile) - 3)
               MsgBox "The first argument (" & Format(FuncID%, "0") & ") in " & MFile$ & _
               " line " & Format(BadLine%, "0") & " is not valid in a Master Script.", _
               vbCritical, "Undefined Master Script Command"
         End Select
      End If
   Loop While ReadAgain%

End Sub

Private Sub ReadScript()
   
   Dim FormRef As Form
   
   LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, Args)
   StatusLine% = mnScriptLine
   mnScriptLine = mnScriptLine + 1
   frmScriptInfo.picScriptStat.Line (0, 0)-(mnScriptLine, 75), &HFF0000, BF
   mnScriptLoop = True
   mnMasterLoop = False
   If mnStepping Then StatusLine% = mnScriptLine
   If mnScriptLine = mnNumScriptLines Then
      mnCurrentMode = mIDLE
      tmrScript.ENABLED = False
      cmdScript(mPREVIOUS).ENABLED = False
      cmdScript(mPREVIOUS).Caption = "&<<"
      mnScriptLine = 0
      If Not mnMasterRun Then
         mnScriptLoop = False
         gnScriptSave = False
         gnScriptRun = False
         mnStepping = False
         cmdScript(mSTEP).Caption = "&Step"
         cmdScript(mSTOP).ENABLED = False
         lblScriptStatus.ForeColor = &HFF0000
         msScriptStatus = "Script Status:  Idle ("
         lblScriptStatus.Caption = msScriptStatus & msScriptPath & ")"
         cmdScript(mSTEP).ENABLED = True
      Else
         mnMasterLoop = True
      End If
   End If

   If gnErrFlag And gnPauseEval Then
      gnErrFlag = False
      gnScriptPaused = True
      If (mnCurrentMode = mRUNNING) And Not mnStepping Then cmdScript_Click (5)
      Exit Sub
   End If
   
   If Not (mnCurConditional = 0) Then
      Select Case mnCurConditional
         Case WAITINGENDIF, WAITINGENDDO, WAITINGENDFOR
            lblScriptStatus.Caption = Format(mnScriptLine, "0") & " No-op"
            Exit Sub
      End Select
   Else
      'if nested, set condition to outer loop
      If Not (mnForNest < 0) Then
         mnCurConditional = maForNest(5, mnForNest)
      End If
   End If
      
   'check for script variables
   If Not ((LineType% = mCOMMENT) Or (LineType% = mNOOPLINE)) Then
      FuncID% = Val(Args(1))
      If FuncID% = SLoadCSVList Then
         ListName = LCase(Trim(Args(6)))
         For Each ListObject In mParamList
            If ListObject.ListName = ListName Then
               'if the list already exists, convert only
               'the list file name
               ReDim ListFileName(0)
               ListFileName(0) = Args(3)
               FoundVar% = CheckForVariables(ListFileName)
               Args(3) = ListFileName(0)
               ListExists% = True
               Exit For
            End If
         Next
      End If
      If Not ListExists% Then
         If FuncID% = SGetStringFromList Then VariableName$ = LCase(Trim(Args(4)))
         If FuncID% = SGetCSVsFromList Then VariableName$ = LCase(Trim(Args(5)))
         If FuncID% = SGetFormProps Then VariableName$ = LCase(Trim(Args(4)))
         DontCheck% = ((FuncID% = SSetVariable) Or (FuncID% = SSetVarDefault))
         If Not DontCheck% Then _
         FoundVar% = CheckForVariables(Args)
      End If
   End If
   
   If LineType% = mEVALUATE Then
      If mnStepping And (mnScriptLine > 0) Then
         LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, NextArgs)
      Else
         NextArgs = Args
      End If
      If (LineType% = mEVALUATE) Or (LineType% = mCOMMAND) Then UpdateScriptStatus NextArgs
      CurrentInf$ = frmScriptInfo.txtScriptInfo.Text
      BlankLine$ = ""
      If Not CurrentInf$ = "" Then
         If Not (Right(CurrentInf$, 4) = (vbCrLf & vbCrLf)) Then
            BlankLine$ = vbCrLf
         End If
      End If
      EvalType = Args(1)
      Select Case EvalType
         Case ETimeStamp
            FailureInfo$ = Now()
            PrintTime% = gnShowComments Or gnPrintEvalAll
         Case EError
            EvalResult% = RunEval(Args, FailureInfo$)
         Case Else
            EvalResult% = RunEval(Args, FailureInfo$)
      End Select
      If EvalResult% Then
         FailFlag$ = "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯" & vbCrLf & _
         "*****   Data evaluation failure: " & msScriptPath & ", line " & _
         Format(mnScriptLine - 1, "0") & "   *****" & vbCrLf & vbCrLf
         EndFlag$ = "________________________________________________" & vbCrLf
         'chars 175 and 95
         ErrorLine$ = CurrentInf$ & vbCrLf & FailFlag$ & FailureInfo$ & EndFlag$
         If gnPrintEval Then frmScriptInfo.txtScriptInfo.Text = ErrorLine$
         If (gnLogEvalFail Or gnLogEvalAll) And mnLogOpen Then
            Print #8, FailFlag$ & FailureInfo$ & EndFlag$
         End If
         If gnPauseEval Then
            gnScriptPaused = True
            If (mnCurrentMode = mRUNNING) Then cmdScript_Click (5)
         End If
      Else
         ErrorLine$ = CurrentInf$ & BlankLine$ & FailureInfo$ & vbCrLf
         If gnPrintEvalAll Or PrintTime% Then
            If Not FailureInfo$ = "" Then
               frmScriptInfo.txtScriptInfo.Text = ErrorLine$
            End If
         End If
         If gnLogEvalAll And mnLogOpen Then
            If Not FailureInfo$ = "" Then Print #8, FailureInfo$ & vbCrLf
         End If
      End If
   ElseIf LineType% = mCOMMAND Then
      FormID$ = Trim(Args(0))
      FuncID% = Val(Args(1))
      FuncStat = Trim(Args(2))
      ThirdArg = Trim(Args(3))
      If Val(FormID$) > 2000 Then
         MsgBox "A Master Script command line cannot be run in a subscript.", _
         vbCritical, "Bad Script Line"
         gnErrFlag = True
         gnScriptRun = False
         Exit Sub
      End If

      If FormID$ = "8" Then FormID$ = "800"
      If FormID$ = "80" Then FormID$ = "800"
      If FormID$ = "808" Then FormID$ = "800"
      If (FuncID% = mOPENFORM) And (Not ((msDevName = "0") Or (msDevName = ""))) Then
         ThirdArg = msDevName
         NameIncluded% = True
      End If
      If LineType% = mCOMMAND Then
         If mnStepping And (mnScriptLine > 0) Then
            LineType% = ParseScript(mSUBSCRIPT, mnScriptLine, NextArgs)
         Else
            NextArgs = Args
         End If
         If (LineType% = mEVALUATE) Or (LineType% = mCOMMAND) Then
            If NameIncluded% Then NextArgs(3) = ThirdArg
            UpdateScriptStatus NextArgs
         End If
      End If

      cmdScript(mSTEP).Caption = "&Step"

      If mnJustReading Then
         If (mnCurrentMode = mIDLE) And mnMasterRun Then
            cmdScript(mPREVIOUS).ENABLED = False
            gnScriptSave = False
            gnScriptRun = False
            mnStepping = False
            cmdScript(mSTEP).Caption = "&Step"
            cmdScript(mSTOP).ENABLED = False
            lblScriptStatus.ForeColor = &HFF0000
            Caption = "Current Script: " & msMasterFile & msScriptStatus & msVerification
            cmdScript(mSTEP).ENABLED = True
         End If
         Exit Sub
      End If
      If Not Val(FormID$) = 0 Then
         CurFormType% = Val(Left$(FormID, 1))
         Instance% = Val(Right$(FormID, 2))
         Select Case CurFormType%
            Case ANALOG_IN, ANALOG_OUT
               'analog input and output are the same
               'so instances are added
               FormType% = ANALOG_IO
            Case DIGITAL_IN, DIGITAL_OUT
               'DIGITAL input and output are the same
               'so instances are added
               FormType% = DIGITAL_IO
            Case Else
               FormType% = CurFormType%
         End Select
         If Not FuncID% = mOPENFORM Then
            If (Instance% < (mnInstances(FormType%))) Or (FuncID% = mCLOSEFORM) Then
               GotFormReference% = GetFormReference(FormID$, FormRef)
               If Not GotFormReference% Then Exit Sub
            Else
               FormTypeName$ = GetFormTypeName(FormID$)
               If FormID$ = "800" Then
                  IEEEFunc$ = Trim(Args(1))
                  If IEEEFunc$ = "202" Then
                     OptCondition = Trim(Args(7))
                  Else
                     OptCondition = Trim(Args(5))
                  End If
                  If Not ((OptCondition = True) Or (Trim(OptCondition) = "")) Then Exit Sub
               End If
               gnErrFlag = FormNotOpen(FormTypeName$, FormID$)
               If gnErrFlag Then Exit Sub
            End If
         End If
      End If
      
      Select Case FuncID%
         Case SShowDiag '2001
            'user action message box
            OptCondition = Trim(Args(8))
            If (OptCondition = True) Or (Trim(OptCondition) = "") Then
               DiagType% = Val(Args(4))
               Select Case DiagType%
                  Case 0
                     Title$ = Trim(Args(6))
                     MsgBox Args(3), , Title$
                  Case 1
                     VarName$ = Trim(LCase(Args(5)))
                     Title$ = Trim(Args(6))
                     Default$ = Trim(Args(7))
                     Resp = InputBox(Args(3), Title$, Default$)
                     VarSet% = SetVariable(VarName$, Resp)
               End Select
            End If
         Case SSetBlock       '2016
            OptCondition = Trim(Args(8))
            If (OptCondition = True) Or (Trim(OptCondition) = "") Then
               BlockSize& = Val(Args(3))
               SetBlockSize BlockSize&, False
               InitBlock True
            End If
         Case SLogOutput      '2024
            ScreenOutput% = Val(Args(3))
            FileOutput% = Val(Args(4))
            Filename$ = Trim(Args(5))
            frmScriptInfo.Visible = Not (ScreenOutput% = 0)
            gnPauseEval = ((ScreenOutput% And 8) = 8)
            gnShowComments = ((ScreenOutput% And 4) = 4)
            gnPrintEval = ((ScreenOutput% And 2) = 2)
            gnPrintEvalAll = ((ScreenOutput% And 1) = 1)
            gnLogComments = ((FileOutput% And 4) = 4)
            gnLogEvalFail = ((FileOutput% And 2) = 2)
            gnLogEvalAll = ((FileOutput% And 1) = 1)
            If Not (FileOutput% = 0) Then OpenLogFile Filename$
            UpdateEvalStatus
         Case SLoadStringList '2025
            Filename$ = Trim(Args(3))
            VarName$ = LCase(Trim(Args(4)))
            ReadStringList Filename$, VarName$
         Case SGetStringFromList '2026
            StringIndex% = Val(Args(3))
            If StringIndex% > mnNumStrings Then
               MsgBox "Script requesting string at index " & _
               Format(StringIndex%, "0") & ". Number of strings loaded = " & _
               Format(mnNumStrings + 1, "0") & "."
               gnErrFlag = True
               Exit Sub
            Else
               VarValue = maStringList(StringIndex%)
               VarSet% = SetVariable(VariableName$, VarValue)
            End If
         Case SSetPlotScaling    '2027
            ScriptSet = Val(Args(3))
            SetAutoScale ScriptSet
         Case SSetFirstPlotPoint '2028
            FirstPoint& = Val(Args(3))
            SetFirstPoint FirstPoint&
            RePlot True
         Case SLoadCSVList '2029
            Filename$ = Trim(Args(3))
            VarNameListLen$ = LCase(Trim(Args(4)))
            VarNameNumVals$ = LCase(Trim(Args(5)))
            ListName = LCase(Trim(Args(6)))
            ReadCSVList Filename$, VarNameListLen$, VarNameNumVals$, ListName
         Case SGetCSVsFromList   '2033
            Dim FirstList As CParamList
            StringIndex% = Val(Args(3))
            VarIndex% = Val(Args(4))
            ListName = LCase(Trim(Args(6)))
            If mParamList.Count = 0 Then
               NumCSVArgs% = -1
            Else
               Set FirstList = mParamList.Item(1)
               NamedList% = FirstList.NamedList
               If NamedList% And (Not ListName = "") Then
                  Set FirstList = Nothing
                  Set FirstList = mParamList.Item(ListName)
                  NumCSVArgs% = FirstList.ListSize 'mParamList.Item(ListName).ListSize
               Else
                  NumCSVArgs% = FirstList.ListSize 'mParamList.Item(1).ListSize
               End If
            End If
            If StringIndex% > NumCSVArgs% Then
               MsgBox "Script requesting comma separated values at index " & _
               Format(StringIndex%, "0") & ". Number of CSVs loaded = " & _
               Format(NumCSVArgs% + 1, "0") & "."
               gnErrFlag = True
               Exit Sub
            Else
               'If Not ListName = "" Then
               VarValue = FirstList.GetListItem(StringIndex%, VarIndex%)
               'Else
               '   VarValue = mParamList.Item(1).GetListItem(StringIndex%, VarIndex%)
               'End If
               'VarValue = maCSVList(StringIndex%, VarIndex%)
               VarSet% = SetVariable(VariableName$, VarValue)
            End If
            Set FirstList = Nothing
         Case SCopyFile       '2036
            Filename$ = Trim(Args(3))
            FileDest$ = Trim(Args(4))
            ScrCopyFile Filename$, FileDest$
         Case SRunApp '2037
            CommandLine$ = Trim(Args(3))
            WaitForClose$ = Trim(Args(4))
            AppID = Shell(CommandLine$, vbNormalFocus)
         Case SResetConfig '2039
            Filename$ = "cb.cfg"
            ULStat& = cbSaveConfig(Filename$)
            DoEvents
            ULStat& = cbLoadConfig("cb.cfg")
            DoEvents
            FormID$ = "800"
            GotNewFormReference% = GetFormReference(FormID$, FormRef)
            If GotNewFormReference% Then
               'need to re-initialize GPIB devices
               ConfigCtrlBoard -1
            End If
         Case SGenerateData   '2049
            DataType% = Val(Args(3))
            Ampl$ = Args(7)
            OSet$ = Args(8)
            AmplitudeVal = FormRef.GetCountsFromUnits(Ampl$, False)
            OffsetVal = FormRef.GetCountsFromUnits(OSet$, True)
            RunGenerateData FormRef, DataType%, Args(4), Args(5), Args(6), _
            AmplitudeVal, OffsetVal, Args(9), Args(10), Args(11), _
            Args(12), Args(13), AuxHandle
         Case SPlotGenData    '2050
            RunPlotGenData FormRef
         Case SPlotAcqData    '2051
            RunPlotAcqData FormRef
         Case SWaitForIdle '2053
            StopCount& = Val(Args(3))
            IdleTimeout& = Val(Args(4))
            WaitForIdle FormRef, StopCount&, IdleTimeout&
         Case SWaitForEvent '2054
            EventType& = Val(Args(3))
            EventData& = Val(Args(4))
            EventTimeout& = Val(Args(5))
            WaitForEvent FormRef, EventType&, EventData&, EventTimeout&
         Case SWaitStatusChange '2055
            StopDelta& = Val(Args(3))
            WaitCondition& = Val(Args(4))
            StatusTimeout& = Val(Args(5))
            WaitStatusChange FormRef, StopDelta&, WaitCondition&, StatusTimeout&
         Case SStopOnCount '2056
            mlStopCount = Val(Args(3))
            mlTimeout = Val(Args(4))
            ReturnStat% = FormRef.StopOnCount(mlStopCount, mlTimeout)
            ScripEval.SetEvent 0, mlStopCount, ReturnStat%
         Case SPlotOnCount '2057
            PlotCount& = Val(Args(3))
            TimeLimit& = Val(Args(4))
            ReturnStat% = FormRef.PlotOnInterval(PlotCount&, TimeLimit&)
            'ScripEval.SetEvent 0, mlStopCount, ReturnStat%
         Case SSetBitsPerPort '2058
            PortNum& = Val(Args(3))
            NumBits& = Val(Args(4))
            BitsInPort& = Val(Args(5))
            ReturnStat% = FormRef.SetBitsPerPort(PortNum&, NumBits&, BitsInPort&)
         Case SCounterArm '2060
            OptCondition = Trim(Args(5))
            If (OptCondition = True) Or (OptCondition = "") Then
               CtrNum% = Val(Args(3))
               Enable% = Not (Val(Args(4)) = 0)
               ReturnStat% = FormRef.ArmCounter(CtrNum%, Enable%)
            End If
         Case SDelay    '3000
            'delay time
            DelayTimeRead& = Val(Args(3))
            mlScriptTime = DelayTimeRead&
            mlDelayStart = Timer
         Case SErrorPrint '3001
            ErrState% = Val(Args(3))
            If ErrState% = -1 Then
               If Not mnLocalErrReporting Then _
               gnLocalErrDisp = mnLocalErrReporting
            Else
               mnLocalErrReporting = gnLocalErrDisp
               gnLocalErrDisp = 0: If Not (ErrState% = 0) Then gnLocalErrDisp = 1
            End If
         Case SErrorFlow '3003
            ErrState% = Val(Args(3))
            If ErrState% = -1 Then
               If Not mnLocalErrHandling = 0 Then _
               geErrFlow = mnLocalErrHandling
            Else
               mnLocalErrHandling = geErrFlow
               geErrFlow = ErrState%
            End If
         Case SULErrFlow '3004
            ErrState% = Val(Args(3))
            If ErrState% = -1 Then
               If Not mnULErrHandling = 0 Then _
               gnErrHandling = mnULErrHandling
            Else
               mnULErrHandling = gnErrHandling
               gnErrHandling = ErrState%
            End If
            ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
         Case SULErrReport '3005
            ErrState% = Val(Args(3))
            If ErrState% = -1 Then
               If Not mnULErrReporting = 0 Then _
               gnErrReporting = mnULErrReporting
            Else
               mnULErrReporting = gnErrReporting
               gnErrReporting = ErrState%
            End If
            ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
         Case SSetStaticOption   '3006
            OptCondition = Trim(Args(4))
            If (OptCondition = True) Or (OptCondition = "") Then
               StaticOption& = Val(Args(3))
               FormRef.SetStaticOption StaticOption&
            End If
         Case SScriptRate    '3008
            Interval& = Val(Args(3))
            frmScript.tmrScript.Interval = Interval&
            Me.lblScriptRate.Visible = True
            Me.lblScriptRate.Caption = "Rate: " & Format(Interval&, "0")
         Case SSetVariable    '3009
            OptCondition = Trim(Args(5))
            If (OptCondition = True) Or (OptCondition = "") Then
               VariableName$ = Trim(LCase(Args(3)))
               VarValue = Args(4)
               'set variable to one of several values separated by ";"
               NewValue = SelectValFromList(VarValue)
               If NewValue = "" Then NewValue = VarValue
               If Not NewValue = "_" Then
                  VarSet% = SetVariable(VariableName$, NewValue)
               End If
            End If
         Case SGetFormProps   '3010
            PropName$ = Trim(LCase(Args(3)))
            'VariableName$ = Trim(LCase(Args(4)))
            PropVal = GetFormProps(FormRef, PropName$)
            VarSet% = SetVariable(VariableName$, PropVal)
         Case SGetStaticOptions  '3011
            mlStaticOpt = FormRef.GetStaticOption()
         Case SSetFormProps   '3012
            PropName$ = Trim(LCase(Args(3)))
            PropVal = Args(4)
            Success% = FormRef.SetFormProperty(PropName$, PropVal)
            gnErrFlag = Not Success%
         Case SCloseApp '3013
            End
         Case SSetVarDefault  '3014
            ReDim VarToCheck(0)
            VarToCheck(0) = Args(3)
            FoundVar% = CheckForVariables(VarToCheck)
            If Not FoundVar% Then
               VariableName$ = Trim(LCase(Args(3)))
               VarValue = Args(4)
               Vars% = SetVariable(VariableName$, VarValue)
            End If
         Case SPauseScript  '3015
               OptCondition = Trim(Args(3))
               If (OptCondition = True) Or (Trim(OptCondition) = "") Then
                  gnScriptPaused = True
                  If (mnCurrentMode = mRUNNING) Then cmdScript_Click (5)
               End If
         Case SGetParameterString  '3016
            'FoundVar% = CheckForVariables(Args)
            FunctionID& = Val(Args(3))
            ParamNum& = Val(Args(4))
            ParamVal& = Val(Args(5))
            VariableName$ = Trim(LCase(Args(6)))
            VarValue = GetParamString(FunctionID&, ParamNum&, ParamVal&)
            Vars% = SetVariable(VariableName$, VarValue)
         Case SPicklist    '3017
            IndexStr$ = LCase(Trim(Args(3)))
            IndexArray = Split(IndexStr$, " to ")
            ListIndex& = Abs(Val(IndexArray(0)))
            If UBound(IndexArray) > 0 Then
               LastItem& = Val(IndexArray(1))
               ListLength& = LastItem& - ListIndex&
            Else
               ListLength& = 0
            End If
            ParamList$ = Args(4)
            VariableName$ = Trim(LCase(Args(5)))
            SizeName$ = Trim(LCase(Args(6)))
            DefaultValue$ = Trim(LCase(Args(7)))
            VarValue = GetListItem(ListIndex&, ParamList$, ListSize&, ListLength&, DefaultValue$)
            Vars% = SetVariable(VariableName$, VarValue)
            If Not SizeName$ = "" Then Vars% = SetVariable(SizeName$, ListSize&)
         Case SGenRndVal      '3018
            VariableName$ = Trim(LCase(Args(4)))
            ArgString$ = Trim(Args(3))
            SeedString$ = ParseUnits(ArgString$, UnitType%)
            VarSeed! = Val(SeedString$)
            RndSeed! = Rnd(1)
            If UnitType% = UNITFLOAT Then
               VarValue = Rnd(RndSeed!) * VarSeed!
            Else
               VarValue = CInt(Rnd(RndSeed!) * VarSeed!)
            End If
            Vars% = SetVariable(VariableName$, VarValue)
         Case SPickGroup      '3019
            GroupIndicator$ = Trim(Args(3))
            GroupIndex& = Val(Args(4))
            GroupList$ = Trim(LCase(Args(5)))
            VariableName$ = Trim(LCase(Args(6)))
            SizeName$ = Trim(LCase(Args(7)))
            GroupArray = Split(GroupList$, GroupIndicator$)
            NumGroups& = UBound(GroupArray)
            If GroupIndex& > NumGroups& Then GroupIndex& = NumGroups&
            VarValue = GroupArray(GroupIndex&)
            If Left$(VarValue, 1) = ";" Then VarValue = Mid$(VarValue, 2)
            Vars% = SetVariable(VariableName$, VarValue)
            If Not SizeName$ = "" Then Vars% = SetVariable(SizeName$, NumGroups&)
            'Else
            '   MsgBox "The group index specified (" & Format(GroupIndex&, "0") & _
            '   ") is greater than the number of groups (" & Format(NumGroups&, "0") & _
            '   ").", vbCritical, "Invalid Group in Script"
            'End If
         Case SSetLibType     '3020
            LibString$ = Trim(Args(3))
            If IsNumeric(LibString$) Then LibType% = Val(LibString$)
            FormRef.SetLibType LibType%
         Case SCalcMaxSinDelta     '3021
            MaxDelta = GetSineMaxDelta(Args)
            VariableName$ = Trim(LCase(Args(6)))
            VarSet% = SetVariable(VariableName$, MaxDelta)
         Case SIsListed    '3022
            ParamList$ = Args(3)
            SearchParam$ = Trim(Args(4))
            ListedBoolName$ = Trim(LCase(Args(5)))
            IndexVarName$ = Trim(LCase(Args(6)))
            ListSeparator$ = Trim(LCase(Args(7)))
            If ListSeparator$ = "" Then ListSeparator$ = ";"
            Listed% = False
            ListArray = Split(ParamList$, ListSeparator$)
            ListSize& = UBound(ListArray)
            For ListedItem& = 0 To ListSize&
               If ListArray(ListedItem&) = SearchParam$ Then
                  Listed% = True
                  If Not (IndexVarName$ = "") Then
                     Vars% = SetVariable(IndexVarName$, ListedItem&)
                  End If
                  Exit For
               End If
            Next
            Vars% = SetVariable(ListedBoolName$, Listed%)
         Case SPeriodCalc    '3023
            RateString$ = Args(3)
            PeriodString$ = ConvertRateToPer(RateString$)
            VariableName$ = Trim(Args(4))
            Vars% = SetVariable(VariableName$, PeriodString$)
         Case SPulseWidthCalc    '3024
            TimeString$ = Args(3)
            WidthString$ = ConvertTimeToWidth(TimeString$)
            VariableName$ = Trim(Args(4))
            Vars% = SetVariable(VariableName$, WidthString$)
         Case SMapAISwitch    '3025
            MapString$ = Trim(Args(3))
            SetAiMap MapString$
         Case SStopScript    '3026
            OptCondition = Trim(Args(2))
            If (OptCondition = True) Or (Trim(OptCondition) = "") Then
               cmdScript_Click (mSTOP)
            End If
         Case SSetMCCControl  '3027
            VarValue = Trim(Args(3))
            If IsNumeric(VarValue) Then UseMCCCtlr% = Val(VarValue)
            frmMain.SetMCCControl UseMCCCtlr%
         Case SGetDP8200Cmd  '3028
            DPValue! = Val(Trim(Args(3)))
            DP8200$ = NumericToDP800Cmd(DPValue!)
            VariableName$ = Trim(Args(4))
            Vars% = SetVariable(VariableName$, DP8200$)
         Case SLoadSubScript  '4001
            'master script command to load subscript
            MsgBox "A Master Script command line cannot be run in a subscript.", _
            vbCritical, "Bad Script Line"
            gnErrFlag = True
            Exit Sub
         Case 4990 To 8000
            Select Case FuncID%
               Case mOPENFORM    '5001
                  If FormType% < 1 Then
                     'MsgBox "The script is attempting to open an invalid form type. ", vbOKOnly, "Bad Form"
                     FormTypeName$ = GetFormTypeName(FormID$)
                     gnErrFlag = FormNotOpen(FormTypeName$, FormID$)
                     If gnErrFlag Then Exit Sub
                  End If
                  mfmUniTest.cmdFormType(CurFormType% - 1) = True
                  If gnErrFlag Then Exit Sub
                  mnInstances(FormType%) = mnInstances(FormType%) + 1
                  GotNewFormReference% = GetFormReference(FormID$, FormRef)
                  If Not GotNewFormReference% Then Exit Sub
                  If Not ((msDevName = "0") Or (msDevName = "")) Then
                     Select Case FormType%
                        Case UTILITIES
                           RunUtilSetBoard FormRef, msDevName, mnDevDupe, Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14)
                        Case Else
                           If mbSetLibToUL Then
                              FormRef.SetLibType LibType%
                              mbSetLibToUL = False
                           End If
                           RunSetBoard FormRef, msDevName, mnDevDupe, Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14)
                     End Select
                  End If
               Case mCLOSEFORM   '5002
                  If mnInstances(FormType%) > 0 Then
                     mnInstances(FormType%) = mnInstances(FormType%) - 1
                     Unload FormRef
                  End If
            End Select
         Case Else
            FormType% = Val(Left$(FormID$, 1))
            Select Case FormType%
               Case 0
                  'read time logged
                  TimeRead& = Val(Mid$(FormID$, 2))
                  If mlScriptStartTime < 1 Then
                     If gnScriptRun Then mlScriptStartTime = Timer  'TimeRead&
                     mlScriptTime = 0
                  Else
                     mlScriptTime = TimeRead&
                  End If
               Case ANALOG_IN
                  ConfigAnalogIn FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
               Case ANALOG_OUT
                  ConfigAnalogOut FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
               Case DIGITAL_IN, DIGITAL_OUT
                  ConfigDigitalIO FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
               Case COUNTERS
                  ConfigCounter FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
               Case UTILITIES
                  ConfigUtils FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
               Case Config
                  ConfigConfig FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
               Case GPIB_CTL
                  ConfigGPIB FormRef, FuncID%, FuncStat, Args(3), _
                     Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), _
                     Args(10), Args(11), Args(12), Args(13), Args(14)
            End Select
      End Select
   End If

   If (Not gnScriptRun) And ((Not mnScriptClosed) Or mnReadMaster) Then
      'script cancelled by error elsewhere
      cmdScript_Click (mSTOP)
   End If
   If (mnCurrentMode = mIDLE) Then
      If mnReadMaster Then
         If gnScriptRun Then ReadMaster mnScriptMode
      ElseIf mnMasterRun Then
         cmdScript(mPREVIOUS).ENABLED = False
         gnScriptSave = False
         gnScriptRun = False
         mnStepping = False
         cmdScript(mSTEP).Caption = "&Step"
         cmdScript(mSTOP).ENABLED = False
         cmdScript(mRUN).ENABLED = True
         lblScriptStatus.ForeColor = &HFF0000
         msScriptStatus = "Script Status:  Subscript complete - read " & msMasterFile
         If mnMasterLine = 0 Then msScriptStatus = "Script Status:  Subscript and masterscript complete - END"
         lblScriptStatus.Caption = msScriptStatus
         Caption = "Current Script: " & msMasterFile & msVerification
         cmdScript(mSTEP).ENABLED = True
      End If
   End If

End Sub

Private Sub ReadInitFile()

   For InitLine% = 0 To mnNumInitLines - 1
      LineText = masInit(InitLine%)
      LineType% = ParseScript(mSUBSCRIPT, InitLine%, Args, LineText)
      If Not ((LineType% = mCOMMENT) Or (LineType% = mNOOPLINE)) Then
         FuncID% = Val(Args(1))
         Select Case FuncID%
            Case SSetVariable    '3009
               VariableName$ = Trim(LCase(Args(3)))
               VarValue = Args(4)
               'If VariableName$ = "dprevision" Then
               '   RevString$ = Trim(VarValue)
               '   RevVerified$ = CheckParamRevision(RevString$)
               '   VarSet% = SetVariable(VariableName$, RevVerified$)
               'Else
                  VarSet% = SetVariable(VariableName$, VarValue)
               'End If
               
            Case SPicklist    '3017
               IndexStr$ = LCase(Trim(Args(3)))
               IndexArray = Split(IndexStr$, " to ")
               ListIndex& = Abs(Val(IndexArray(0)))
               If UBound(IndexArray) > 0 Then
                  LastItem& = Val(IndexArray(1))
                  ListLength& = LastItem& - ListIndex&
               Else
                  ListLength& = 0
               End If
               ParamList$ = Args(4)
               VariableName$ = Trim(LCase(Args(5)))
               SizeName$ = Trim(LCase(Args(6)))
               DefaultValue$ = Trim(LCase(Args(7)))
               VarValue = GetListItem(ListIndex&, ParamList$, ListSize&, ListLength&, DefaultValue$)
               VarSet% = SetVariable(VariableName$, VarValue)
               If Not SizeName$ = "" Then VarSet% = SetVariable(SizeName$, ListSize&)
               'If mnReadMaster Then ReadMaster mnScriptMode
            Case Else
               Warning% = True
               If Not Warning% Then MsgBox "There is an invalid value in 'DeviceParams.uts'." & _
               vbCrLf & "Initialize files can only be used to set simple variables.", vbInformation, _
               "Invalid Line in Initializer Script"
         End Select
      End If
   Next
   
End Sub

Private Sub ResetScript()

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
   'this shouldn't be used anymore
   Close #2
   Close #3
   'Close #4
   'Open msScriptPath For Input As #2
   'Open msScriptPath For Input As #3
   If Not OpenScript() Then Exit Sub

   Input #3, FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   For SLine% = 1 To mnScriptLine - 1
      Input #2, FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
      Input #3, FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   Next
   Input #3, FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   mnScriptLine = mnScriptLine - 1
   'UpdateScriptStatus FormID$, FuncID%, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle
   mnScriptLine = mnScriptLine + 1

End Sub

Private Sub RunAIDisableEvent(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'disable the event
   FormRef.cmdConfigure.Caption = "d" & A2
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSetPoints(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   FormRef.cmdConfigure.Caption = "F7"
   FormRef.cmdConfigure = True

   'set queue size
   FormRef.cmdConfigure.Caption = "u" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   If Val(A2) > 0 Then
      'don't bother with these if queue is disabled

      'select a queue element
      FormRef.cmdConfigure.Caption = "v" & Format$(A3, "000")
      FormRef.cmdConfigure = True

      'set limit A
      FormRef.cmdConfigure.Caption = "a" & Format$(A4, "0.00")
      FormRef.cmdConfigure = True

      'set limit B
      FormRef.cmdConfigure.Caption = "b" & Format$(A5, "0.00")
      FormRef.cmdConfigure = True

      'set output 1
      FormRef.cmdConfigure.Caption = "c" & Format$(A6, "0.00")
      FormRef.cmdConfigure = True

      'set output 2
      FormRef.cmdConfigure.Caption = "d" & Format$(A7, "0.00")
      FormRef.cmdConfigure = True

      'set mask 1
      FormRef.cmdConfigure.Caption = "e" & Format$(A8, "0.00")
      FormRef.cmdConfigure = True

      'set mask 2
      FormRef.cmdConfigure.Caption = "f" & Format$(A9, "0.00")
      FormRef.cmdConfigure = True

      'select a flag type and latch (update on true and false)
      FormRef.cmdConfigure.Caption = "w" & Format$(A10, "000")
      FormRef.cmdConfigure = True

      'select an output type
      FormRef.cmdConfigure.Caption = "z" & Format$(A11, "000")
      FormRef.cmdConfigure = True

   End If

   'load the queue channel element
   FormRef.cmdConfigure.Caption = "i"
   FormRef.cmdConfigure = True

   'call the setpoint function
   FormRef.cmdConfigure.Caption = "j"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAIEnableEvent(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'set event size
   FormRef.cmdConfigure.Caption = "s" & A3
   FormRef.cmdConfigure = True

   EventType& = Val(A2)
   If EventType& = ALL_EVENT_TYPES Then
      'enable all events
      FormRef.cmdConfigure.Caption = "e" & A2
      FormRef.cmdConfigure = True
   Else
      'parse through and enable each type that is set
      For eType% = 0 To 6
         If (2 ^ eType% And EventType&) = 2 ^ eType% Then
            CurEventType& = Choose(eType% + 1, ON_SCAN_ERROR, _
            ON_EXTERNAL_INTERRUPT, ON_PRETRIGGER, ON_DATA_AVAILABLE, _
            ON_END_OF_AI_SCAN, ON_END_OF_AO_SCAN, ON_CHANGE_DI)
            FormRef.cmdConfigure.Caption = "e" & Format(CurEventType&, "0")
            FormRef.cmdConfigure = True
         End If
      Next
   End If

End Sub

Private Sub RunAIn(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set form to use AIn rather than AIn32
   FormRef.Set32Bit 1
   
   'set function to AIn
   FormRef.cmdConfigure.Caption = "F0"
   FormRef.cmdConfigure = True

   'set rate to zero
   FormRef.cmdConfigure.Caption = "T000"
   FormRef.cmdConfigure = True
   
   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A5, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A6, "000")
   FormRef.cmdConfigure = True
   
   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set total number of calls to cbAIn using auxilliary data
   If Not mnCalMode Then
      FormRef.cmdConfigure.Caption = "C" & Format$(A7, "000")
      FormRef.cmdConfigure = True
   End If

   If mnCalMode Then FormRef.cmdPlot.Caption = mlCalVal

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunAIn32(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set form to use AIn32 rather than AIn
   FormRef.Set32Bit 2
   
   'set function to AIn
   FormRef.cmdConfigure.Caption = "F0"
   FormRef.cmdConfigure = True

   'set rate to zero
   FormRef.cmdConfigure.Caption = "T000"
   FormRef.cmdConfigure = True
   
   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A7, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbAIn using auxilliary data
   If Not mnCalMode Then
      FormRef.cmdConfigure.Caption = "C" & Format$(A8, "000")
      FormRef.cmdConfigure = True
   End If

   'set options
   BaseOpts& = A5
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 21  'this number changes if new options are added
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 5 Then
            'if apparently SINGLEIO
            If ((2 ^ 6) And Opts&) = (2 ^ 6) Then
               'check if also DMAIO (if both SIO & DIO then
               MenuIndex% = 7 'its actually BLOCKIO)
               i% = 6
            End If
         End If
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%
   
   If mnCalMode Then FormRef.cmdPlot.Caption = mlCalVal

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunVIn32(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set form to use VIn32 rather than VIn
   FormRef.Set32Bit 2
   
   'set function to VIn
   FormRef.cmdConfigure.Caption = "F8"
   FormRef.cmdConfigure = True

   'set rate to zero
   FormRef.cmdConfigure.Caption = "T000"
   FormRef.cmdConfigure = True
   
   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A7, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbVIn32 using auxilliary data
   If Not mnCalMode Then
      FormRef.cmdConfigure.Caption = "C" & Format$(A8, "000")
      FormRef.cmdConfigure = True
   End If

   'set options
   BaseOpts& = A5
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 21  'this number changes if new options are added
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 5 Then
            'if apparently SINGLEIO
            If ((2 ^ 6) And Opts&) = (2 ^ 6) Then
               'check if also DMAIO (if both SIO & DIO then
               MenuIndex% = 7 'its actually BLOCKIO)
               i% = 6
            End If
         End If
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%
   
   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunAInScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)


   'set form to use native resolution
   FormRef.Set32Bit 0
   
   'set function
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True
   
   FormRef.cmdConfigure.Caption = "?"
   FormRef.cmdConfigure = True
   'check if Q set - if so, don't set chans
   ChanConfig$ = FormRef.cmdConfigure.Caption
   If Not Left(ChanConfig$, 1) = "Q" Then
      'set high chan
      FormRef.cmdConfigure.Caption = "H" & Format$(A3, "000")
      FormRef.cmdConfigure = True
      'set low chan
      FormRef.cmdConfigure.Caption = "L" & Format$(A2, "000")
      FormRef.cmdConfigure = True
   End If

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A8
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 21  'this number changes if new options are added
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 5 Then
            'if apparently SINGLEIO
            If ((2 ^ 6) And Opts&) = (2 ^ 6) Then
               'check if also DMAIO (if both SIO & DIO then
               MenuIndex% = 7 'its actually BLOCKIO)
               i% = 6
            End If
         End If
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%
   
   'set total count
   If Not mnCalMode Then
      FormRef.cmdConfigure.Caption = "C" & Format$(A4, "000")
      FormRef.cmdConfigure = True
   End If

   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A5, "000.0###")
   FormRef.cmdConfigure = True
   
   If mnCalMode Then FormRef.cmdPlot.Caption = mlCalVal

   'now run it
   DoEvents
   FormRef.cmdGo = True

End Sub

Private Sub RunSetBoard(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   BoardName$ = Trim(A1)
   DupeSelect$ = Trim(A2)
   FormRef.cmdConfigure.Caption = "B" & BoardName$ & "," & DupeSelect$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAOSetAmpl(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set amplitude to specified value
   Ampl$ = Trim(A3)
   If Trim(A3) = "" Then Ampl$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "M" & Ampl$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAOSetData(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set data menu to specified index
   FormRef.cmdConfigure.Caption = "D" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAOSetDev(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set GPIB device to device indicated
   FormRef.cmdConfigure.Caption = "U" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAOSetOS(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set offset to specified value
   Offset$ = Trim(A3)
   If Trim(A3) = "" Then Offset$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "I" & Offset$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAOut(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to AOut
   FormRef.cmdConfigure.Caption = "F0"
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set output value
   ATude$ = Trim(A4)
   If Not (InStr(1, ATude$, "%") = 0) Then
      Resolution% = GetResolution(FormRef)
      Range% = GetCurrentRange(FormRef)
      FSR& = 2 ^ Resolution%
      ATude$ = Left(ATude$, Len(ATude$) - 1)
      Ampl& = CLng(FSR& * (Val(ATude$) / 100))
      Amplitude$ = Format(Ampl&, "000")
   Else
      Amplitude$ = Format(ATude$, "000")
   End If
   FormRef.cmdConfigure.Caption = "V" & Amplitude$
   FormRef.cmdConfigure = True
   
   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A5, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A6, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbAOut using auxilliary data
   FormRef.cmdConfigure.Caption = "C" & Format$(A7, "000")
   FormRef.cmdConfigure = True

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunVOut(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to VOut
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set output value
   VoltVal$ = A4
   FormRef.SetVoltageValue Trim(VoltVal$)
   
   'to do - add options argument
   
   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A7, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbAOut using auxilliary data
   FormRef.cmdConfigure.Caption = "C" & Format$(A8, "000")
   FormRef.cmdConfigure = True

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunAOutScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to AOutScan
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'set total number of data values
   FormRef.cmdConfigure.Caption = "C" & Format$(A4, "000")
   FormRef.cmdConfigure = True

   'set low chan
   FormRef.cmdConfigure.Caption = "L" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   'set high chan
   FormRef.cmdConfigure.Caption = "H" & Format$(A3, "000")
   FormRef.cmdConfigure = True
   
   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A8
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 20
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 6 Then If ((2 ^ 5) And Opts&) = (2 ^ 5) Then MenuIndex% = 7 'special case for BLOCKIO
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%
   
   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A5, "000")
   FormRef.cmdConfigure = True
   
   'set output value    {This needs work.. A4 is handle to array}
   'FormRef.cmdConfigure.Caption = "D" & Format$(A4, "000")
   'FormRef.cmdConfigure = True
   
   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunAPretrig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set form to use native resolution
   FormRef.Set32Bit 0
   
   'set function
   FormRef.cmdConfigure.Caption = "F3"
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A7, "000")
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A9
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 20
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 6 Then If ((2 ^ 5) And Opts&) = (2 ^ 5) Then MenuIndex% = 7 'special case for BLOCKIO
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set low chan
   FormRef.cmdConfigure.Caption = "L" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   'set high chan
   FormRef.cmdConfigure.Caption = "H" & Format$(A3, "000")
   FormRef.cmdConfigure = True
   
   'set total count
   FormRef.cmdConfigure.Caption = "C" & Format$(A5, "000")
   FormRef.cmdConfigure = True

   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A6, "000")
   FormRef.cmdConfigure = True
   
   'set pretrig count
   FormRef.cmdConfigure.Caption = "P" & Format$(A4, "000")
   FormRef.cmdConfigure = True
   
   'now run it
   FormRef.cmdGo = True

End Sub

Private Sub RunASetGetStatus(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on or off GetStatus function
   StatusCheck$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "g" & StatusCheck$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunATrig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to ATrig
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'set trigtype
   FormRef.cmdConfigure.Caption = "G" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set low threshold
   FormRef.cmdConfigure.Caption = "Y" & Format$(A10, "000")
   FormRef.cmdConfigure = True

   'set high threshold
   FormRef.cmdConfigure.Caption = "Z" & Format$(A11, "000")
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A5, "000")
   FormRef.cmdConfigure = True

   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A7, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A8, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbAIn (after trigger) using auxilliary data
   FormRef.cmdConfigure.Caption = "C" & Format$(A9, "000")
   FormRef.cmdConfigure = True

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunBufInfo(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'update plot for current instance of form
   FormRef.cmdConfigure.Caption = "b"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunC8254Config(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to 8254Config
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'configure the counter
   FormRef.cmdConfigure.Caption = "C" & Format$(A3, "00")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunC9513Config(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to 9513Config
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'cbC9513Config(mnBoardNum, CounterNum%, GateControl%, CounterEdge%, CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)

   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "00")
   FormRef.cmdConfigure = True

   'set GateControl%
   FormRef.cmdConfigure.Caption = "G" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'set CounterEdge%
   FormRef.cmdConfigure.Caption = "E" & Format$(A4, "00")
   FormRef.cmdConfigure = True

   'set CountSource%
   FormRef.cmdConfigure.Caption = "S" & Format$(A5, "00")
   FormRef.cmdConfigure = True

   'set SpecialGate%
   FormRef.cmdConfigure.Caption = "P" & Format$(A6, "00")
   FormRef.cmdConfigure = True

   'set Reload%
   FormRef.cmdConfigure.Caption = "D" & Format$(A7, "00")
   FormRef.cmdConfigure = True

   'set RecycleMode%
   FormRef.cmdConfigure.Caption = "Y" & Format$(A8, "00")
   FormRef.cmdConfigure = True

   'set BCDMode%
   FormRef.cmdConfigure.Caption = "M" & Format$(A9, "00")
   FormRef.cmdConfigure = True

   'set CountDirec%
   FormRef.cmdConfigure.Caption = "T" & Format$(A10, "00")
   FormRef.cmdConfigure = True

   'set OutputCtrl% and configure the counter
   FormRef.cmdConfigure.Caption = "C" & Format$(A11, "00")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunC9513Init(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to 9513Init
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'cbC9513Init(BoardNum%, ChipNum%, FOutDivider%, FOutSource%, Compare1%, Compare2%, TimeOfDay%)

   'set chip number
   Chip% = A2
   FormRef.cmdConfigure.Caption = "H" & Format$(Chip%, "00")
   FormRef.cmdConfigure = True

   'set FOutDivider%
   FormRef.cmdConfigure.Caption = "W" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'set FOutSource%
   FormRef.cmdConfigure.Caption = "U" & Format$(A4, "00")
   FormRef.cmdConfigure = True

   'set Compare1%
   FormRef.cmdConfigure.Caption = "?" & Format$(A5, "00")
   FormRef.cmdConfigure = True

   'set Compare2%
   FormRef.cmdConfigure.Caption = "@" & Format$(A6, "00")
   FormRef.cmdConfigure = True

   'set TimeOfDay%
   FormRef.cmdConfigure.Caption = "A" & Format$(A7, "00")
   FormRef.cmdConfigure = True

   'init the counter
   FormRef.cmdOK = True

End Sub

Private Sub RunCalMode(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set unitest to calibration mode
   FormRef.cmdConfigure.Caption = "@" & A1
   FormRef.cmdConfigure = True
   mnCalMode = Val(A1)
   If A1 = "True" Then mnCalMode = True

End Sub

Private Sub RunCFreqIn(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CFreqIn
   FormRef.cmdConfigure.Caption = "F6"
   FormRef.cmdConfigure = True

   'cbCFreqIn(mnBoardNum, SigSource%, GateInterval%, CBCount%, Freq&)

   'set SigSource%
   Chip% = A2
   FormRef.cmdConfigure.Caption = "U" & Format$(Chip%, "00")
   FormRef.cmdConfigure = True

   'set GateInterval%
   FormRef.cmdConfigure.Caption = "V" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'run CFreqIn
   FormRef.cmdGo = True

End Sub

Private Sub RunCIn(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CIn
   FormRef.cmdConfigure.Caption = "F4"
   FormRef.cmdConfigure = True

   'clear the CLoad32 checkbox
   FormRef.cmdConfigure.Caption = "J0"
   FormRef.cmdConfigure = True
   
   'set counter number
   Counter% = Val(A2)
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True
   
   'set number of times to read counter
   '(negative values for separate reads preserving data)
   NumLoops& = Val(A4)
   If NumLoops& = 0 Then NumLoops& = 1
   FormRef.cmdConfigure.Caption = "K" & Format$(NumLoops&, "000")
   FormRef.cmdConfigure = True

   'read the counter
   FormRef.cmdConfigure.Caption = "I" & Format$(A3, "00")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCIn32(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CIn
   FormRef.cmdConfigure.Caption = "F4"
   FormRef.cmdConfigure = True

   'set the CLoad32 checkbox
   FormRef.cmdConfigure.Caption = "J1"
   FormRef.cmdConfigure = True
   
   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'set number of times to read counter
   '(negative values for separate reads preserving data)
   NumLoops& = Val(A4)
   If NumLoops& = 0 Then NumLoops& = 1
   FormRef.cmdConfigure.Caption = "K" & Format$(NumLoops&, "000")
   FormRef.cmdConfigure = True
   
   'read the counter
   FormRef.cmdConfigure.Caption = "I" & Format$(A3, "00")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCIn64(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CIn
   FormRef.cmdConfigure.Caption = "F4"
   FormRef.cmdConfigure = True

   'set the CLoad32 and CLoad64 checkboxes
   FormRef.cmdConfigure.Caption = "J3"
   FormRef.cmdConfigure = True
   
   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'set number of times to read counter
   '(negative values for separate reads preserving data)
   NumLoops& = Val(A4)
   If NumLoops& = 0 Then NumLoops& = 1
   FormRef.cmdConfigure.Caption = "K" & Format$(NumLoops&, "000")
   FormRef.cmdConfigure = True
   
   'read the counter
   FormRef.cmdConfigure.Caption = "I" & Format$(A3, "00")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCLoad(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CLoad
   FormRef.cmdConfigure.Caption = "F3"
   FormRef.cmdConfigure = True

   'clear the CLoad32 checkbox
   FormRef.cmdConfigure.Caption = "J0"
   FormRef.cmdConfigure = True
   
   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'load the counter
   FormRef.cmdConfigure.Caption = "L" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'uncheck any register checkboxes that might be set
   FormRef.cmdConfigure.Caption = "f"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCLoad32(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CLoad32
   FormRef.cmdConfigure.Caption = "F3"
   FormRef.cmdConfigure = True

   'set the CLoad32 checkbox
   FormRef.cmdConfigure.Caption = "J1"
   FormRef.cmdConfigure = True

   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'load the counter
   FormRef.cmdConfigure.Caption = "L" & Format$(A3, "00")
   FormRef.cmdConfigure = True
   
   'uncheck any register checkboxes that might be set
   FormRef.cmdConfigure.Caption = "f"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunConvert(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on cbAConvertData
   EnableConv$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "K" & Format$(EnableConv$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunConvertData(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on cbAConvertData
   EnableConv$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "E" & Format$(EnableConv$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunConvertPretrigData(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on ConvertPretrigData
   TrueFalseVal$ = Trim(A2)
   FormRef.cmdConfigure.Caption = "D" & TrueFalseVal$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunConvertPT(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on ConvertPretrigData
   EnablePT$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "J" & Format$(EnablePT$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCountSet(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'count set separately for calibration mode
   'set number of points to calibrate
   FormRef.cmdConfigure.Caption = "C" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCStoreOnInt(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to StoreOnInt
   FormRef.cmdConfigure.Caption = "F5"
   FormRef.cmdConfigure = True

   'cbCStoreOnInt(mnBoardNum, IntCount&, CntrControl%(0), mvHandle)

   'select multiple counters
   FormRef.cmdConfigure.Caption = "X" & Format$(A5, "000")
   FormRef.cmdConfigure = True

   'set IntCount&
   FormRef.cmdConfigure.Caption = "K" & Format$(A2, "000")
   FormRef.cmdConfigure = True
   
   'now run it
   FormRef.cmdGo = True

End Sub

Private Sub RunDBitIn(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DBitIn
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'read all bits
   FormRef.cmdConfigure.Caption = "I"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDBitOut(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DBitOut
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'set port to AUXPORT or FIRSTPORTA
   'and write bit value
   Port% = A2
   PortType$ = "Z"   'AUXPORT
   If A2 > 1 Then PortType$ = "9"   'FIRSTPORTA
   FormRef.cmdConfigure.Caption = PortType$ & Format$(A5, "00.00") & ":" & Format$(A6, "00.00") & ":" & Format$(A4, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDConfig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DConfigPort
   FormRef.cmdConfigure.Caption = "F3"
   FormRef.cmdConfigure = True

   'set port
   Port% = A2
   If A2 > 1 Then Port% = A2 - 8
   FormRef.cmdConfigure.Caption = "P" & Format$(Port%, "000")
   FormRef.cmdConfigure = True

   'set direction and write config
   FormRef.cmdConfigure.Caption = "E" & Format$(A3, "000")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDConfigBit(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DConfigBit
   FormRef.cmdConfigure.Caption = "F4"
   FormRef.cmdConfigure = True

   'set port to AUXPORT or FIRSTPORTA
   'and write bit value
   Port% = A2
   PortType$ = "X"   'AUXPORT
   If A2 > 1 Then PortType$ = "A"   'FIRSTPORTA
   FormRef.cmdConfigure.Caption = PortType$ & Format$(A5, "00.00") & ":" & Format$(A6, "00.00") & ":" & Format$(A4, "0")
   FormRef.cmdConfigure = True

   'set direction and write config
   'FormRef.cmdConfigure.Caption = "E" & Format$(A3, "000")
   'FormRef.cmdConfigure = True

End Sub

Private Sub RunDIDisableEvent(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'disable the event
   FormRef.cmdConfigure.Caption = "d" & A2
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDIEnableEvent(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'enable the event
   FormRef.cmdConfigure.Caption = "e" & A2
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDInScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DInScan
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'set port
   Port% = A2
   If A2 > 1 Then Port% = A2 - 8
   FormRef.cmdConfigure.Caption = "P" & Format$(Port%, "000")
   FormRef.cmdConfigure = True
   
   'set options
   BaseOpts& = A6
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   'to do - new implementation of BLOCK_IO
   For i% = 0 To 22  'this number changes if new options are added
      OptValue& = 2 ^ i%
      If (OptValue& And Opts&) = OptValue& Then
         Index = Switch(OptValue& = BACKGROUND, 0, OptValue& = CONTINUOUS, 1, _
         OptValue& = EXTCLOCK, 2, OptValue& = WORDXFER, 3, OptValue& = DWORDXFER, 4, _
         OptValue& = DMAIO, 5, OptValue& = BLOCKIO, 6, OptValue& = SIMULTANEOUS, 7, _
         OptValue& = EXTTRIGGER, 8, OptValue& = NONSTREAMEDIO, 9, OptValue& = ADCCLOCKTRIG, 10, _
         OptValue& = ADCCLOCK, 11, OptValue& = HIGHRESRATE, 12)
         If Not IsNull(Index) Then MenuIndex% = Index
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%
   
   NumSamples& = Val(A3)
   ScanRate& = Val(A4)
   FormRef.ScriptDInScan NumSamples&, ScanRate&

End Sub

Private Sub RunDIn(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DIn
   FormRef.cmdConfigure.Caption = "F0"
   FormRef.cmdConfigure = True

   'set port
   Port% = A2
   If A2 > 1 Then Port% = A2 - 8
   FormRef.cmdConfigure.Caption = "P" & Format$(Port%, "000")
   FormRef.cmdConfigure = True

   'read the port
   FormRef.cmdConfigure.Caption = "R"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDOut(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DOut
   FormRef.cmdConfigure.Caption = "F0"
   FormRef.cmdConfigure = True

   'set port
   Port% = A2
   If A2 > 1 Then Port% = A2 - 8
   FormRef.cmdConfigure.Caption = "P" & Format$(Port%, "000")
   FormRef.cmdConfigure = True

   'set the value to write
   FormRef.cmdConfigure.Caption = "V" & Format$(A3, "000")
   FormRef.cmdConfigure = True
   
   'write the port
   FormRef.cmdConfigure.Caption = "W"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunFileAInScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function
   FormRef.cmdConfigure.Caption = "F4"
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set filename
   NameOfFile$ = Trim(A7)
   FormRef.cmdConfigure.Caption = "N" & NameOfFile$
   FormRef.cmdConfigure = True
   
   'set options
   BaseOpts& = A8
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 20
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 6 Then If ((2 ^ 5) And Opts&) = (2 ^ 5) Then MenuIndex% = 7 'special case for BLOCKIO
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set low chan
   FormRef.cmdConfigure.Caption = "L" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   'set high chan
   FormRef.cmdConfigure.Caption = "H" & Format$(A3, "000")
   FormRef.cmdConfigure = True
   
   'set total count
   FormRef.cmdConfigure.Caption = "C" & Format$(A4, "000")
   FormRef.cmdConfigure = True

   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A5, "000")
   FormRef.cmdConfigure = True
   
   'now run it
   FormRef.cmdGo = True

End Sub

Private Sub RunFilePretrig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function
   FormRef.cmdConfigure.Caption = "F5"
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "R" & Format$(A7, "000")
   FormRef.cmdConfigure = True

   'set filename
   FormRef.cmdConfigure.Caption = "N" & A8
   FormRef.cmdConfigure = True
   
   'set options
   BaseOpts& = A9
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 20
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 6 Then If ((2 ^ 5) And Opts&) = (2 ^ 5) Then MenuIndex% = 7 'special case for BLOCKIO
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set low chan
   FormRef.cmdConfigure.Caption = "L" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   'set high chan
   FormRef.cmdConfigure.Caption = "H" & Format$(A3, "000")
   FormRef.cmdConfigure = True
   
   'set pretrig count
   FormRef.cmdConfigure.Caption = "P" & Format$(A4, "000")
   FormRef.cmdConfigure = True

   'set total count
   FormRef.cmdConfigure.Caption = "C" & Format$(A5, "000")
   FormRef.cmdConfigure = True

   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A6, "000")
   FormRef.cmdConfigure = True
   
   'now run it
   FormRef.cmdGo = True

End Sub

Private Function RunGetStatus(FormID As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle) As Integer

   FormID.cmdConfigure.Caption = "Q"
   FormID.cmdConfigure = True
   DoEvents
   Status$ = FormID.cmdConfigure.Caption
   NumArgs& = FindInString(Status$, ",", Locations)
   If NumArgs& = 1 Then
      StatVal& = Val(Left(Status$, Locations(0) - 1))
      CountVal& = Val(Mid(Status$, Locations(0) + 1, Locations(1) - (Locations(0) + 1)))
      IndexVal& = Val(Mid(Status$, Locations(1) + 1))
   End If
   ReturnVal% = StatVal&
   RunGetStatus = ReturnVal%

End Function

Private Sub RunGPDevClear(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   OptCondition = Trim(A3)
   If Not ((OptCondition = True) Or (Trim(OptCondition) = "")) Then Exit Sub
   'clear devices
   FormRef.cmdConfigure.Caption = "C"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunGPIBRead(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'read from the device
   FormRef.cmdConfigure.Caption = "R"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunGPIBTrig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'trigger the device (using GET)
   FormRef.cmdConfigure.Caption = "T"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunGPIBWrite(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   OptCondition = Trim(A5)
   If Not ((OptCondition = True) Or (Trim(OptCondition) = "")) Then Exit Sub
   
   'set the string to write
   FormRef.txtCommand.Text = Trim(A3)

   'write to the device
   FormRef.cmdConfigure.Caption = "W"
   FormRef.cmdConfigure = True
   If mnCalMode Then mlCalVal = Val(Mid$(A3, 3))
   If Not (Trim(A4) = "") Then
      ActualVal$ = FormRef.GetReturnVal()
      VarValue! = HP8112CmdToNumeric(ActualVal$)
      RateReturned! = 1 / VarValue!
      VarName$ = A4
      Vars% = SetVariable(VarName$, RateReturned!)
   End If

End Sub

Private Sub RunGPRen(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'toggle the REN line for local or remote
   FormRef.cmdConfigure.Caption = "L" & A3
   FormRef.cmdConfigure = True

End Sub

Private Sub RunGPSelDevClear(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'clear selected device
   FormRef.cmdConfigure.Caption = "S"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunInByte(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to InByte
   FormRef.cmdConfigure.Caption = "F5"
   FormRef.cmdConfigure = True

   'set byte mode (as opposed to word mode)
   FormRef.cmdConfigure.Caption = "Y"
   FormRef.cmdConfigure = True

   'set composite mode on or off
   FormRef.cmdConfigure.Caption = "C" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'set consecutive registers for composite on or off
   FormRef.cmdConfigure.Caption = "A" & Format$(A4, "00")
   FormRef.cmdConfigure = True

   'set mask value
   FormRef.cmdConfigure.Caption = "M" & Format$(A5, "00")
   FormRef.cmdConfigure = True

   'set where to apply mask
   MaskFirst% = A6
   MaskSecond% = A7
   MaskConfig% = Abs(MaskFirst%) + Abs(MaskSecond% * 2)
   FormRef.cmdConfigure.Caption = "N" & Format$(MaskConfig%, "00")
   FormRef.cmdConfigure = True

   'set surrogate enable / disable
   FormRef.cmdConfigure.Caption = "S" & Format$(A8, "00")
   FormRef.cmdConfigure = True

   'set surrogate board
   FormRef.cmdConfigure.Caption = "U" & Format$(A9, "00")
   FormRef.cmdConfigure = True

   'set devnum
   FormRef.cmdConfigure.Caption = "D" & Format$(A10, "00")
   FormRef.cmdConfigure = True
   
   'set register to read
   FormRef.cmdConfigure.Caption = "R" & Format$(A2, "00")
   FormRef.cmdConfigure = True
   
   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunInWord(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to InWord
   FormRef.cmdConfigure.Caption = "F5"
   FormRef.cmdConfigure = True

   'set word mode (as opposed to byte mode)
   FormRef.cmdConfigure.Caption = "W"
   FormRef.cmdConfigure = True

   'set composite mode on or off
   FormRef.cmdConfigure.Caption = "C" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'set consecutive registers for composite on or off
   FormRef.cmdConfigure.Caption = "A" & Format$(A4, "00")
   FormRef.cmdConfigure = True

   'set mask value
   FormRef.cmdConfigure.Caption = "M" & Format$(A5, "00")
   FormRef.cmdConfigure = True

   'set where to apply mask
   MaskFirst% = A6
   MaskSecond% = A7
   MaskConfig% = Abs(MaskFirst%) + Abs(MaskSecond% * 2)
   FormRef.cmdConfigure.Caption = "N" & Format$(MaskConfig%, "00")
   FormRef.cmdConfigure = True

   'set surrogate enable / disable
   FormRef.cmdConfigure.Caption = "S" & Format$(A8, "00")
   FormRef.cmdConfigure = True

   'set surrogate board
   FormRef.cmdConfigure.Caption = "U" & Format$(A9, "00")
   FormRef.cmdConfigure = True

   'set devnum
   FormRef.cmdConfigure.Caption = "D" & Format$(A10, "00")
   FormRef.cmdConfigure = True
   
   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunOutByte(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to OutByte
   FormRef.cmdConfigure.Caption = "F6"
   FormRef.cmdConfigure = True

   'set byte mode (as opposed to word mode)
   FormRef.cmdConfigure.Caption = "Y"
   FormRef.cmdConfigure = True

   'set composite mode on or off
   FormRef.cmdConfigure.Caption = "C" & Format$(A4, "00")
   FormRef.cmdConfigure = True

   'set consecutive registers for composite on or off
   FormRef.cmdConfigure.Caption = "A" & Format$(A5, "00")
   FormRef.cmdConfigure = True

   'set mask value
   FormRef.cmdConfigure.Caption = "M" & Format$(A6, "00")
   FormRef.cmdConfigure = True

   'set where to apply mask
   MaskFirst% = A7
   MaskSecond% = A8
   MaskConfig% = Abs(MaskFirst%) + Abs(MaskSecond% * 2)
   FormRef.cmdConfigure.Caption = "N" & Format$(MaskConfig%, "00")
   FormRef.cmdConfigure = True

   'set surrogate enable / disable
   FormRef.cmdConfigure.Caption = "S" & Format$(A9, "00")
   FormRef.cmdConfigure = True

   'set surrogate board
   FormRef.cmdConfigure.Caption = "U" & Format$(A10, "00")
   FormRef.cmdConfigure = True

   'set devnum
   FormRef.cmdConfigure.Caption = "D" & Format$(A11, "00")
   FormRef.cmdConfigure = True

   'set value to write
   FormRef.cmdConfigure.Caption = "V" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'set reg to write
   FormRef.cmdConfigure.Caption = "R" & Format$(A2, "00")
   FormRef.cmdConfigure = True
   
   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunOutWord(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to OutWord
   FormRef.cmdConfigure.Caption = "F6"
   FormRef.cmdConfigure = True

   'set word mode (as opposed to byte mode)
   FormRef.cmdConfigure.Caption = "W"
   FormRef.cmdConfigure = True

   'set composite mode on or off
   FormRef.cmdConfigure.Caption = "C" & Format$(A4, "00")
   FormRef.cmdConfigure = True

   'set consecutive registers for composite on or off
   FormRef.cmdConfigure.Caption = "A" & Format$(A5, "00")
   FormRef.cmdConfigure = True

   'set mask value
   FormRef.cmdConfigure.Caption = "M" & Format$(A6, "00")
   FormRef.cmdConfigure = True

   'set where to apply mask
   MaskFirst% = A7
   MaskSecond% = A8
   MaskConfig% = Abs(MaskFirst%) + Abs(MaskSecond% * 2)
   FormRef.cmdConfigure.Caption = "N" & Format$(MaskConfig%, "00")
   FormRef.cmdConfigure = True

   'set surrogate enable / disable
   FormRef.cmdConfigure.Caption = "S" & Format$(A9, "00")
   FormRef.cmdConfigure = True

   'set surrogate board
   FormRef.cmdConfigure.Caption = "U" & Format$(A10, "00")
   FormRef.cmdConfigure = True

   'set devnum
   FormRef.cmdConfigure.Caption = "D" & Format$(A11, "00")
   FormRef.cmdConfigure = True

   'set value to write
   FormRef.cmdConfigure.Caption = "V" & Format$(A3, "00")
   FormRef.cmdConfigure = True

   'set reg to write
   FormRef.cmdConfigure.Caption = "R" & Format$(A2, "00")
   FormRef.cmdConfigure = True

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunPlotBlock(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on ConvertPretrigData
   FormRef.cmdConfigure.Caption = "n" & Format$(A1, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunPlotContin(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn on continuous plotting
   PlotState$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "A" & Format$(PlotState$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSelectSignal(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to SelectSignal
   FormRef.cmdConfigure.Caption = "F3"
   FormRef.cmdConfigure = True

   'set direction
   FormRef.cmdConfigure.Caption = "D" & A2
   FormRef.cmdConfigure = True

   'clear the selection list boxes
   FormRef.cmdConfigure.Caption = "R"
   FormRef.cmdConfigure = True

   'set signal
   FormRef.cmdConfigure.Caption = "S" & A3
   FormRef.cmdConfigure = True

   'set connection
   FormRef.cmdConfigure.Caption = "C" & A4
   FormRef.cmdConfigure = True

   'set polarity
   FormRef.cmdConfigure.Caption = "P" & A5
   FormRef.cmdConfigure = True

   'execute the function
   FormRef.cmdConfigure.Caption = "X"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSetBlock(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'set block size
   FormRef.cmdConfigure.Caption = "m" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSetRes(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'set plot resolution
   FormRef.cmdConfigure.Caption = "r" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunGetConfig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

End Sub

Private Sub RunSetConfig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to SetConfig
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'set InfoType
   FormRef.cmdConfigure.Caption = "T" & A1 - 1
   FormRef.cmdConfigure = True
   
   'set DevNum
   FormRef.cmdConfigure.Caption = "E" & A2
   FormRef.cmdConfigure = True

   'set ConfigItem
   ItemVal$ = Trim(A4)
   If IsNumeric(ItemVal$) Then
      CfgItem% = A4
      ListItem% = FormRef.GetCfgItemIndex(CfgItem%)
      FormRef.cmdConfigure.Caption = "I" & ListItem%
      FormRef.cmdConfigure = True
   Else
      Component% = Val(A1) - 1
      ValSet$ = Trim(A5)
      FormRef.SetConfigValues Component%, ItemVal$, ValSet$
      MessageSet% = True
   End If

   'set ConfigVal
   If Not MessageSet% Then
      FormRef.cmdConfigure.Caption = "V" & A5
      FormRef.cmdConfigure = True
   End If
   
   'now run it
   FormRef.cmdOK = True

End Sub

Private Sub RunSetPlotChan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'set Plot channel
   FormRef.cmdConfigure.Caption = "c" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSetPlotOpts(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'set Retain Plot mode on or off
   FormRef.cmdConfigure.Caption = "x" & A1
   FormRef.cmdConfigure = True

   'set Plot Title mode to BoardName or Scale
   FormRef.cmdConfigure.Caption = "y" & A2
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSetPTBuf(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn 512 addition to buffer on or off
   AddBuf$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "a" & AddBuf$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunSetTrigger(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   If Not ((InStr(1, A4, ".") = 0) And (InStr(1, A3, ".") = 0)) Then
      TrigType% = Val(A2)
      LowThreshold! = Val(A3)
      HighThreshold! = Val(A4)
      TriggerOption% = Val(A5)
      
      FormRef.SetTrigVal TrigType%, LowThreshold!, HighThreshold!, TriggerOption%
   Else
      'set trigger type
      FormRef.cmdConfigure.Caption = "t" & Format$(A2, "000")
      FormRef.cmdConfigure = True
   
      'set low threshold
      FormRef.cmdConfigure.Caption = "l" & Format$(A3, "000")
      FormRef.cmdConfigure = True
   
      'set high threshold
      FormRef.cmdConfigure.Caption = "h" & Format$(A4, "000")
      FormRef.cmdConfigure = True
      
      'execute trigger function
      FormRef.cmdConfigure.Caption = "W"
      FormRef.cmdConfigure = True
   End If

End Sub

Private Sub RunStopBackground(FormID As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   FormID.cmdConfigure.Caption = "X"
   FormID.cmdConfigure = True

End Sub

Private Sub RunTIn(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to TIn
   FormRef.cmdConfigure.Caption = "F6"
   FormRef.cmdConfigure = True

   'set scale
   FormRef.cmdConfigure.Caption = "S" & Format$(A3, "000")
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A5
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 20
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 6 Then If ((2 ^ 5) And Opts&) = (2 ^ 5) Then MenuIndex% = 7 'special case for BLOCKIO
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set low chan using auxilliary data
   FormRef.cmdConfigure.Caption = "L" & Format$(A6, "000")
   FormRef.cmdConfigure = True

   'set high chan using auxilliary data
   FormRef.cmdConfigure.Caption = "H" & Format$(A7, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbTIn using auxilliary data
   FormRef.cmdConfigure.Caption = "C" & Format$(A8, "000")
   FormRef.cmdConfigure = True

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunTInScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to TInScan
   FormRef.cmdConfigure.Caption = "F7"
   FormRef.cmdConfigure = True

   'set scale
   FormRef.cmdConfigure.Caption = "S" & Format$(A4, "000")
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A6
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 20
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 6 Then If ((2 ^ 5) And Opts&) = (2 ^ 5) Then MenuIndex% = 7 'special case for BLOCKIO
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set low chan
   FormRef.cmdConfigure.Caption = "L" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   'set high chan
   FormRef.cmdConfigure.Caption = "H" & Format$(A3, "000")
   FormRef.cmdConfigure = True
   
   'set total number of calls to cbTInScan using auxilliary data
   FormRef.cmdConfigure.Caption = "C" & Format$(A7, "000")
   FormRef.cmdConfigure = True

   'start function
   FormRef.cmdGo = True

End Sub

Private Sub RunUtilSetBoard(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'set board to name specified
   BoardName$ = Trim(A2)
   DupeSelect$ = Trim(A3)
   FormRef.cmdConfigure.Caption = "B" & BoardName$ & "," & DupeSelect$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunALoadQueue(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   If Val(A3) > Val(A2) Then
      MsgBox "Script is configuring element " & A3 & _
      " but only " & A2 & " elements are in queue."
      gnErrFlag = True
      Exit Sub
   End If

   'set queue size
   FormRef.cmdConfigure.Caption = "u" & Format$(A2, "000")
   FormRef.cmdConfigure = True

   If Val(A2) > 0 Then
      'don't bother with these if queue is disabled

      'select a queue element
      FormRef.cmdConfigure.Caption = "v" & Format$(A3, "000")
      FormRef.cmdConfigure = True
      If FormRef.cmdConfigure.Caption = "Error" Then
         MsgBox "Script has not properly configured the number of elements in the queue.", _
         vbOKOnly, "Script Error"
         gnErrFlag = True
         Exit Sub
      End If

      'select a queue channel type and indicate if setpoint
      ChanType$ = Format$(A4, "00")
      If ChanType$ = "-2" Then ChanType$ = "0"
      FormRef.cmdConfigure.Caption = "w" & ChanType$
      FormRef.cmdConfigure = True

      'select a queue channel number
      FormRef.cmdConfigure.Caption = "z" & Format$(A5, "000")
      FormRef.cmdConfigure = True

      'select a queue channel gain
      FormRef.cmdConfigure.Caption = "R" & Format$(A6, "000")
      FormRef.cmdConfigure = True

      'set the channel mode value (if valid)
      ModeString$ = Trim(A8)
      If Not (ModeString$ = "") Then
         ModeIndex% = Val(A8)
         If (ModeIndex% = 0) Or (ModeIndex% = 1) Then
            FormRef.cmdConfigure.Caption = "[" & Format$(ModeIndex%, "0")
            FormRef.cmdConfigure = True
         End If
      End If
      
      'set the data rate value (if valid)
      DRateString$ = Trim(A9)
      DRateVal& = Val(DRateString$)
      If DRateVal& > 0 Then
         FormRef.cmdConfigure.Caption = "]" & DRateString$
         FormRef.cmdConfigure = True
      End If
      
      'load the queue channel element
      FormRef.cmdConfigure.Caption = "i"
      FormRef.cmdConfigure = True

      'set the enable setpoint value
      FormRef.cmdConfigure.Caption = "!" & Format$(A7, "000")
      FormRef.cmdConfigure = True
      
   End If

   'if not "load only", finish queue setup
   If Not (A4 = -2) Then
      FormRef.cmdConfigure.Caption = "j"
      FormRef.cmdConfigure = True
   End If

End Sub

Private Sub RunDOSetData(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'configure per parameters
   'set data menu to specified index
   FormRef.cmdConfigure.Caption = "D" & A1
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDOSetAmpl(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set amplitude to specified value
   FormRef.cmdConfigure.Caption = "M" & A3
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDOSetOS(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set offset to specified value
   FormRef.cmdConfigure.Caption = "J" & A3
   FormRef.cmdConfigure = True

End Sub

Private Sub RunDOutScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A6
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   'to do - new implementation of BLOCK_IO
   For i% = 0 To 22  'this number changes if new options are added
      OptValue& = 2 ^ i%
      If (OptValue& And Opts&) = OptValue& Then
         Index = Switch(OptValue& = BACKGROUND, 0, OptValue& = CONTINUOUS, 1, _
         OptValue& = EXTCLOCK, 2, OptValue& = WORDXFER, 3, OptValue& = DWORDXFER, 4, _
         OptValue& = DMAIO, 5, OptValue& = BLOCKIO, 6, OptValue& = SIMULTANEOUS, 7, _
         OptValue& = EXTTRIGGER, 8, OptValue& = NONSTREAMEDIO, 9, OptValue& = ADCCLOCKTRIG, 10, _
         OptValue& = ADCCLOCK, 11, OptValue& = HIGHRESRATE, 12)
         If Not IsNull(Index) Then MenuIndex% = Index
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set port
   'port is now ignored (replaced with select ports method)
   'Port% = A2
   'If A2 > 1 Then Port% = A2 - 8
   'FormRef.cmdConfigure.Caption = "P" & Format$(Port%, "000")
   'FormRef.cmdConfigure = True

   'set total count and rate
   NumSamples& = Val(A3)
   ScanRate& = Val(A4)
   
   'now run it
   FormRef.ScriptDOutScan NumSamples, ScanRate

End Sub

Private Sub StartRecord()

   gnScriptSave = True
   Open msScriptPath For Output As #2
   mlScriptStartTime = Timer
   mfmUniTest.picCommands.Visible = True
   mnCurrentMode = mRECORDING
   msScriptStatus = "Script Status:  Recording "
   lblScriptStatus.ForeColor = &HFF
   lblScriptStatus.Caption = msScriptStatus & msScriptPath
   cmdScript(mSTOP).ENABLED = True
   cmdScript(mSTEP).Caption = "Delay"
   tmrScript.ENABLED = True
   
End Sub

Private Sub StopRecord()

   gnScriptSave = False
   gnScriptRun = False
   Close #2
   lblScriptStatus.ForeColor = &HFF0000
   mlScriptStartTime = -1
   mlScriptTime = 0
   CloseScript
   tmrScript.ENABLED = False
   cmdScript(mSTEP).Caption = "&Step"

End Sub

Private Sub mnuVariables_Click()

   GridData = ScriptVars
   frmVariables.LoadGrid GridData
   frmVariables.Show 1
   Unload frmVariables
   
End Sub

Private Sub tmrScript_Timer()

   Static NumChecks As Long, FirstVal As Long
   If Not gnScriptRun Then Exit Sub
   If Not (mnCheckingStatus Or mnWaitingEvent Or mnWaitingStatVal) Then
      If tmrScript.ENABLED Then
         cmdScript(mPREVIOUS).ENABLED = True
         cmdScript(mPREVIOUS).Caption = "||"
         mnScriptMode = mRUN
         mnStepping = False
      End If
      If mnCurrentMode = mRECORDING Then
      ElseIf mlDelayStart <> 0 Then
         If (Timer - mlDelayStart) > mlScriptTime Then
            mlDelayStart = 0
            ReadScript
         Else
            CurETA& = mlScriptTime - (Timer - mlDelayStart)
            ReDim A(5) As Variant
            A(0) = 0: A(1) = 3000: A(2) = 0
            A(3) = Str(mlScriptTime)
            Args = A()
            UpdateScriptStatus Args
            lblScriptStatus.Caption = lblScriptStatus.Caption & _
            "  (Resuming script in " & Format(CurETA&, "0") & " seconds.)"
         End If
      Else 'read existing script
         If (Timer - mlScriptStartTime > mlScriptTime) Or (mlScriptTime = 0) Then
            'If mnReadMaster Then
            If mnMasterLoop Then
               ReadMaster mnScriptMode
            Else
               ReadScript
            End If
         End If
      End If
   Else
      DoEvents
      Select Case True
         Case mnWaitingEvent
            mfrmFormRef.GetEvent EventDetected&, EventData&, EventParam&
            NumChecks = NumChecks + 1
            If (mlTimeout > 0) Then
               CurETA& = mlTimeout - NumChecks
               ReDim A(5) As Variant
               A(0) = 0: A(1) = 2054: A(2) = 0
               A(3) = Str(mlEventType): A(4) = Str(mlEventData)
               A(5) = Str(mlTimeout): Args = A()
               UpdateScriptStatus Args
               lblScriptStatus.Caption = lblScriptStatus.Caption & _
               "  (Timeout in " & Format(CurETA&, "0") & " ticks.)"
               If (NumChecks > mlTimeout) Then
                  Abort% = True
               End If
            End If
            If (Not (EventDetected& And mlEventType) = 0) Or Abort% Then
               If (mlEventData > 0) And Not Abort% Then
                  'waiting for data of at least mlEventData
                  If EventData& < mlEventData Then Exit Sub
               End If
               ScripEval.SetEvent EventDetected&, EventData&, Abort%
               mnWaitingEvent = False
               NumChecks = 0
            End If
      Case mnWaitingStatVal
         'in this context, mlEventType is the type of compare to FirstVal
         'and mlEventData is the criteria for comparison
         StatError& = ReadStatus(mfrmFormRef, StatVal%, CurIndex&, CurCount&)
         If NumChecks = 0 Then FirstVal = CurCount&
         If (mlTimeout > 0) Then
            CurETA& = mlTimeout - NumChecks
            ReDim A(4) As Variant
            A(0) = 0: A(1) = 2053: A(2) = 0
            A(3) = Str(mlWaitForCount): A(4) = Str(mlTimeout)
            Args = A()
            UpdateScriptStatus Args
            lblScriptStatus.Caption = lblScriptStatus.Caption & _
            "  (Timeout in " & Format(CurETA&, "0") & " ticks.)"
            If (NumChecks > mlTimeout) Then Abort% = True
         End If
         If (NumChecks > mlTimeout) Or Not (StatError& = 0) Then Abort% = True 'at least one comparison
         NumChecks = NumChecks + 1
         If Not (NumChecks = 1) Then
            Select Case mlEventType
               Case 0
                  'keeps checking until no change between checks
                  If (CurCount& = FirstVal) Or Abort% Then
                     If NumChecks > mlEventData Then
                        ScripEval.SetEvent 0, CurCount&, Abort%
                        mnWaitingStatVal = False
                        NumChecks = 0
                     End If
                  Else
                     FirstVal = CurCount&
                  End If
               Case 1
                  'keeps checking until change between checks
                  If (CurCount& > FirstVal) Or Abort% Then
                     If NumChecks > mlEventData Then
                        ScripEval.SetEvent 0, CurCount&, Abort%
                        mnWaitingStatVal = False
                        NumChecks = 0
                     End If
                  End If
            End Select
            If ((StatVal% = IDLE) Or (StopCount%)) Or Abort% Then
               Set mfrmFormRef = Nothing
               mnCheckingStatus = False
               mnWaitingStatVal = False
               ScripEval.SetEvent 0, 0, Abort%
               NumChecks = 0
            End If
         End If
      Case mnCheckingStatus
         StatError& = ReadStatus(mfrmFormRef, StatVal%, CurIndex&, CurCount&)
         If mlWaitForCount > 0 Then StopCount% = CurCount& > mlWaitForCount
         NumChecks = NumChecks + 1
         If (mlTimeout > 0) Then
            CurETA& = mlTimeout - NumChecks
            ReDim A(4) As Variant
            A(0) = 0: A(1) = 2053: A(2) = 0
            A(3) = Str(mlWaitForCount): A(4) = Str(mlTimeout)
            Args = A()
            UpdateScriptStatus Args
            lblScriptStatus.Caption = lblScriptStatus.Caption & _
            "  (Timeout in " & Format(CurETA&, "0") & " ticks.)"
            If (NumChecks > mlTimeout) Or Not (StatError& = 0) Then Abort% = True
         End If
         If ((StatVal% = IDLE) Or (StopCount%)) Or Abort% Then
            Set mfrmFormRef = Nothing
            mnCheckingStatus = False
            ScripEval.SetEvent 0, 0, Abort%
            NumChecks = 0
         End If
      End Select
   End If

End Sub

Private Sub UpdateScriptStatus(Args As Variant)
   
   x% = VarType(Args)
   If x% = (vbArray Or vbVariant) Then
      ArraySize% = UBound(Args) - 1
      If (Val(Args(0)) = SLoadSubScript) Or (Val(Args(0)) = SCloseSubScript) Then
         FuncID% = Val(Args(0))
         ArgStart% = 2
      Else
         If ArraySize% > 0 Then FuncID% = Val(Args(1))
         If ArraySize% > 1 Then FuncStat& = Val(Args(2))
         ArgStart% = 3
         If FuncID% > 10000 Then ArgStart% = 2
      End If
   End If

   ArgString$ = "Arg Vals: "
   lblScriptStatus = mnScriptLine + 1 & ") " & GetFunctionString(FuncID%) & Chr$(13) & Chr$(10)
   Select Case FuncStat&
      'Case 0
      '   ArgString$ = ArgString$ & "  " & Val(FormID$)
      Case ANALOG_IN
         ArgString$ = ArgString$ & "Analog Input  "
      Case ANALOG_OUT
         ArgString$ = ArgString$ & "Analog Output  "
      Case DIGITAL_IN
         ArgString$ = ArgString$ & "Digital Input  "
      Case DIGITAL_OUT
         ArgString$ = ArgString$ & "Digital Output  "
      Case COUNTERS
         ArgString$ = ArgString$ & "Counters  "
      Case UTILITIES
         ArgString$ = ArgString$ & "Utilities  "
      Case Config
         ArgString$ = ArgString$ & "Configuration  "
      Case GPIB_CTL
         ArgString$ = ArgString$ & "GPIB Control  "
   End Select
   NumArgs% = GetNumArgs(FuncID%)
   For i% = 0 To NumArgs% - 1
      ArgVal = Trim(Args(i% + ArgStart%)) & " "
      If gnHexVals Then
         If (VarType(ArgVal) = 2) Or (VarType(ArgVal) = 3) Then
            ArgString$ = ArgString$ & Hex$(ArgVal) & ", "
         Else
            ArgString$ = ArgString$ & ArgVal & ", "
         End If
      Else
         ArgString$ = ArgString$ & ArgVal
      End If
   Next i%
   lblScriptStatus.Caption = lblScriptStatus.Caption & ArgString$

End Sub

Private Sub RunDaqInScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'low chan and high chan set by queue function (A2 is one element of channel array)

   'A3 is one element of channel type array

   'range is set by queue function (A4 is one element of gain array) and isn't passed here

   'A5 is the queue count already set in the queue function

   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A6, "000")
   FormRef.cmdConfigure = True
   
   'set total count
   FormRef.cmdConfigure.Caption = "C" & Format$(A8, "000")
   FormRef.cmdConfigure = True

   'A9 is handle

   'set options
   BaseOpts& = A10
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 21  'this number changes if new options are added
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 5 Then
            'if apparently SINGLEIO
            If ((2 ^ 6) And Opts&) = (2 ^ 6) Then
               'check if also DMAIO (if both SIO & DIO then
               MenuIndex% = 7 'its actually BLOCKIO)
               i% = 6
            End If
         End If
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'set pretrig value (A7 is value returned from function)
   FormRef.cmdConfigure.Caption = "M" & Format$(A11, "000")
   FormRef.cmdConfigure = True

   'now run it
   FormRef.cmdGo = True

End Sub

Private Sub RunDaqOutScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'low chan and high chan set by queue function (A2 is one element of channel array)

   'A3 is one element of channel type array

   'range is set by queue function (A4 is one element of gain array) and isn't passed here

   'A5 is the queue count already set in the queue function

   'set rate
   FormRef.cmdConfigure.Caption = "T" & Format$(A6, "000")
   FormRef.cmdConfigure = True
   
   'set total count
   FormRef.cmdConfigure.Caption = "C" & Format$(A7, "000")
   FormRef.cmdConfigure = True

   'A8 is handle

   'set options
   BaseOpts& = A9
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 21  'this number changes if new options are added
      If ((2 ^ i%) And Opts&) = (2 ^ i%) Then
         MenuIndex% = i%
         If i% = 5 Then
            'if apparently SINGLEIO
            If ((2 ^ 6) And Opts&) = (2 ^ 6) Then
               'check if also DMAIO (if both SIO & DIO then
               MenuIndex% = 7 'its actually BLOCKIO)
               i% = 6
            End If
         End If
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%

   'now run it
   FormRef.cmdGo = True

End Sub

Private Sub RunDaqSetTrig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to DaqSetTrigger
   FormRef.cmdConfigure.Caption = "F4"
   FormRef.cmdConfigure = True

   'set source
   FormRef.cmdConfigure.Caption = "m" & A2
   FormRef.cmdConfigure = True

   'set trigger type
   FormRef.cmdConfigure.Caption = "n" & A3
   FormRef.cmdConfigure = True

   'set trigger channel
   FormRef.cmdConfigure.Caption = "o" & A4
   FormRef.cmdConfigure = True

   'set channel type
   FormRef.cmdConfigure.Caption = "p" & A5
   FormRef.cmdConfigure = True

   'set range
   FormRef.cmdConfigure.Caption = "r" & A6
   FormRef.cmdConfigure = True
   
   'set trigger level
   FormRef.cmdConfigure.Caption = "t" & A7
   FormRef.cmdConfigure = True
   
   'set level variance
   FormRef.cmdConfigure.Caption = "y" & A8
   FormRef.cmdConfigure = True
   
   'set event to trigger (start or stop)
   FormRef.cmdConfigure.Caption = "k" & A9
   FormRef.cmdConfigure = True
   
   'execute the function
   FormRef.cmdConfigure.Caption = "x"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCConfigScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set counter number (stored in A9 as ctrnum + 1 - actual value is A2)
   Counter% = A9
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'set function to CConfigScan
   FormRef.cmdConfigure.Caption = "F1"
   FormRef.cmdConfigure = True

   'set mode parameter
   Mode$ = Trim(A3)
   FormRef.cmdConfigure.Caption = "m" & Format$(Mode$, "0")
   FormRef.cmdConfigure = True

   'set debounce parameter
   FormRef.cmdConfigure.Caption = "d" & Format$(A4, "0")
   FormRef.cmdConfigure = True

   'set debounce trigger parameter
   FormRef.cmdConfigure.Caption = "g" & Format$(A5, "0")
   FormRef.cmdConfigure = True

   'set edge detection parameter
   FormRef.cmdConfigure.Caption = "e" & Format$(A6, "0")
   FormRef.cmdConfigure = True

   'set map channel parameter
   FormRef.cmdConfigure.Caption = "p" & Format$(A7, "0")
   FormRef.cmdConfigure = True

   'execute function
   FormRef.cmdConfigure.Caption = "x"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunC7266Config(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'set function to CConfigScan
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'set quadrature parameter
   FormRef.cmdConfigure.Caption = "n" & Format$(A3, "0")
   FormRef.cmdConfigure = True
   
   'set counting mode parameter (normal, rangelimit, recycle, modulo)
   FormRef.cmdConfigure.Caption = "m" & Format$(A4, "0")
   FormRef.cmdConfigure = True

   'set decode parameter
   FormRef.cmdConfigure.Caption = "E" & Format$(A5, "0")
   FormRef.cmdConfigure = True

   'set index mode parameter
   FormRef.cmdConfigure.Caption = "G" & Format$(A6, "0")
   FormRef.cmdConfigure = True

   'set index invert parameter
   FormRef.cmdConfigure.Caption = "P" & Format$(A7, "0")
   FormRef.cmdConfigure = True

   'set flag pins parameter
   FormRef.cmdConfigure.Caption = "S" & Format$(A8, "0")
   FormRef.cmdConfigure = True

   'set gate parameter
   FormRef.cmdConfigure.Caption = "D" & Format$(A9, "0")
   FormRef.cmdConfigure = True
   
   DoEvents
   'execute function
   FormRef.cmdConfigure.Caption = "y"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCInScan(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set function to CInScan
   FormRef.cmdConfigure.Caption = "F3"
   FormRef.cmdConfigure = True

   'set low counter number
   Counter% = A2
   FormRef.cmdConfigure.Caption = "l" & Format$(Counter%, "0")
   FormRef.cmdConfigure = True

   'set high counter number
   Counter% = A3
   FormRef.cmdConfigure.Caption = "h" & Format$(Counter%, "0")
   FormRef.cmdConfigure = True

   'set scan count
   TotalCount& = A4
   FormRef.cmdConfigure.Caption = "i" & Format$(TotalCount&, "0")
   FormRef.cmdConfigure = True

   'set scan rate
   CBRate& = A5
   FormRef.cmdConfigure.Caption = "r" & Format$(CBRate&, "0")
   FormRef.cmdConfigure = True

   'set options
   BaseOpts& = A6
   Opts& = mlStaticOpt Or BaseOpts&
   FormRef.cmdConfigure.Caption = "O-1"  'clear options
   FormRef.cmdConfigure = True
   For i% = 0 To 5  'this number changes if new options are added
      MenuOptVal& = Choose(i% + 1, BACKGROUND, CONTINUOUS, _
      EXTTRIGGER, EXTCLOCK, CTR32BIT, CTR48BIT)
      If (MenuOptVal& And Opts&) = MenuOptVal& Then
         MenuIndex% = i%
         FormRef.cmdConfigure.Caption = "O" & Format$(MenuIndex%, "00")
         FormRef.cmdConfigure = True
      End If
   Next i%
   
   'execute CInScan function
   FormRef.cmdConfigure.Caption = "s"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunCClear(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set counter number (stored in A3 as ctrnum + 1 - actual value is A2)
   Counter% = A3
   FormRef.cmdConfigure.Caption = "N" & Format$(Counter%, "000")
   FormRef.cmdConfigure = True

   'set function to CClear
   FormRef.cmdConfigure.Caption = "F2"
   FormRef.cmdConfigure = True

   'execute function
   FormRef.cmdConfigure.Caption = "C"
   FormRef.cmdConfigure = True

End Sub

Private Sub RunTimerOutStop(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'call timer stop function
   FormRef.cmdConfigure.Caption = "a" & A2
   FormRef.cmdConfigure = True

End Sub

Private Sub RunTimerOutStart(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'set timer number
   FormRef.cmdConfigure.Caption = "b" & A2
   FormRef.cmdConfigure = True

   'call timer start function
   FormRef.cmdConfigure.Caption = "c" & A3
   FormRef.cmdConfigure = True

End Sub

Private Sub RunAInputMode(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'call AInputMode function (true sets SE)
   FormRef.SetInputMode A2

End Sub

Private Sub RunAChanInputMode(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'call AChanInputMode function (Channel, Mode)
   FormRef.SetChanInputMode A3, A4

End Sub

Private Sub RunSetCalcNoise(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   Calculate% = Val(A1)
   CalcNoise Calculate%
   
End Sub

Private Sub RunShowText(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   ShowText A2

End Sub

Private Sub RunSetPlotType(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   PlotType$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "{" & Format$(PlotType$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunGetTC(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'turn GetTCValues on or off
   GetTC$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "f" & Format$(GetTC$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunUseEngUnits(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'turn EngUnits on or off
   UseEng$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "}" & Format$(UseEng$, "0")
   FormRef.cmdConfigure = True

End Sub

Private Sub RunBitSelect(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'sets up a range of bits that remain in affect
   'for the given form until index of -1 is sent as first bit
   
   FirstBit& = Val(A1)
   LastBit& = Val(A2)
   FormRef.SelectBits FirstBit&, LastBit&

End Sub

Private Sub RunPortSelect(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'sets up a range of ports that remain in affect
   'for the given form until index of -1 is sent as first port
   
   FirstPort$ = Trim(A1)
   LastPort$ = Trim(A2)
   FormRef.cmdConfigure.Caption = "s" & FirstPort$ & "," & LastPort$
   FormRef.cmdConfigure = True

End Sub

Sub RunPlotAcqData(FormRef As Form)
   
   FormRef.PlotAcquiredData

End Sub

Sub RunPlotGenData(FormRef As Form)
   
   FormRef.PlotGenData

End Sub

Sub RunGenerateData(FormRef As Form, A1, A2, A3, A4, _
Ampl, OSet, A7, A8, A9, EvalOption, A11, AuxHandle)

   'creates data in memory and passes the handle to the
   'associated form along with required information such
   'as number of channels
   
   TypeOfData = VarType(Ampl)
   RangeVal% = Abs(GetCurrentRange(FormRef))
   Resolution% = Abs(GetResolution(FormRef))
   Select Case TypeOfData
      Case vbSingle
         SingleVal! = Ampl
         BipFactor% = 1
         If RangeVal% < 100 Then BipFactor% = 2
         Amplitude = (VoltsToCounts(Resolution%, RangeVal%, SingleVal!)) * BipFactor%
      Case vbString
         'set Amplitude to percentage of full scale
         Perc$ = Left(Ampl, Len(Ampl) - 1)
         FSFactor! = Val(Perc$) / 100
         Amplitude = (2 ^ Resolution%) * FSFactor!
      Case Else
         Amplitude = Ampl
   End Select
   TypeOfData = VarType(OSet)
   Select Case TypeOfData
      Case vbSingle
         SingleVal! = OSet
         Offset = GetCounts(Resolution%, RangeVal%, SingleVal!)
      Case vbString
         'set offset to percentage of full scale
         Perc$ = Left(OSet, Len(OSet) - 1)
         FSFactor! = Val(Perc$) / 100
         Offset = (2 ^ Resolution%) * FSFactor!
      Case Else
         Offset = OSet
   End Select
   
   DataType& = Val(A1)
   Cycles% = Val(A2)
   TotalCount& = Val(A3)
   NumChans% = Val(A4)
   SigType% = Val(A7)
   NewData% = Not (Val(A8) = 0)
   Channel% = Val(A9)
   
   If VarType(EvalOption) = vbString Then
      EmptyString% = (EvalOption = "")
   End If
   If IsEmpty(EvalOption) Or EmptyString% Then
      EvalParam& = 0
      FirstPoint& = 0
      EvaluationOption$ = ""
   Else
      ParseAll = Split(EvalOption, ";")
      NumOptions& = UBound(ParseAll)
      For CurOpt& = 0 To NumOptions&
         ThisOpt = Trim(ParseAll(CurOpt&))
         EvaluationType = Split(ThisOpt, " ")
         EvalParam& = UBound(EvaluationType)
         If EvalParam& = 0 Then
            FirstPoint& = Val(EvaluationType(0))
         Else
            EvaluationOption$ = LCase(EvaluationType(0))
            'following allows for "String = Value"
            'or "String Value" construct
            OptionValue& = EvaluationType(EvalParam&)
            Select Case EvaluationOption$
               Case "gain", "range"
                  Range& = OptionValue&
                  'set range
                  FormRef.cmdConfigure.Caption = "R" & Format$(Range&, "000")
                  FormRef.cmdConfigure = True
               Case "first", "firstpoint", "start"
                  FirstPoint& = OptionValue&
                  EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
            End Select
         End If
      Next
   End If
   DataHandle& = FormRef.ScriptGenData(DataType&, Cycles%, TotalCount&, NumChans%, _
   Amplitude, Offset, SigType%, NewData%, Channel%, FirstPoint&)
   'DataHandle& = GenerateData(DataType&, Cycles%, TotalCount&, NumChans%, _
   Amplitude, Offset, SigType%, NewData%, Channel%, FirstPoint&)
   
   'BufferSize& = GetCurBufferSize()
   'FormRef.SetupData DataHandle&, BufferSize&, NumChans%, DataType&
   'FormRef.PlotGenData
   
End Sub

Sub RunWriteBitRange(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'writes the data generated by 'GenerateData' to
   'the range of bits/ports selected using RunBitSelect/RunPortSelect
   
   NumBlocks& = Val(A1)
   FormRef.WriteSelectedBits NumBlocks&

End Sub

Sub RunWritePortRange(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'writes the data generated by 'GenerateData' to
   'the range of ports selected using RunPortSelect
   
   NumBlocks& = Val(A1)
   FormRef.WriteSelectedPorts NumBlocks&

End Sub

Sub RunReadBitRange(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'reads a range of bits selected using RunBitSelect
   
   NumberOfReads& = Val(A1)
   FormRef.ReadSelectedBits NumberOfReads&

End Sub

Sub RunReadPortRange(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'reads a range of ports selected using RunPortSelect
   
   NumberOfReads& = Val(A1)
   FormRef.ReadSelectedPorts NumberOfReads&
   'FormRef.cmdConfigure.Caption = "C" & NumberOfReads$
   'FormRef.cmdConfigure = True
   
   'FormRef.cmdConfigure.Caption = "i"
   'FormRef.cmdConfigure = True

End Sub

Sub RunBitConfig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'configures a range of ports selected using RunPortSelect
   
   BitDirection& = Val(A1)
   FormRef.ConfigSelectedBits BitDirection&

End Sub

Sub RunPortConfig(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'configures a range of ports selected using RunPortSelect
   
   PortDirection$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "c" & PortDirection$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunFormReference(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
   
   'gives current form a reference to another for interaction
   FormNumber$ = Trim(A1)
   FormRef.cmdConfigure.Caption = "6" & FormNumber$
   FormRef.cmdConfigure = True

End Sub

Private Sub RunChanSetup(FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   SetEvalChan Val(A1)

End Sub

Private Sub RunMaxMinSetup(FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)
         
   MinVal& = Val(A1)
   MaxVal& = Val(A2)
   SetVMinMaxStop MinVal&, MaxVal&

End Sub
Private Sub RunDeltaSetup(FormID$, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'configure the evaluation parameters
   MaxDelta& = Val(A1)
   SetDeltaStop MaxDelta&
   
End Sub

Private Sub RunEvalEnable(FormRef As Form, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, AuxHandle)

   'enable or disable data evaluation for a particular form
   EnableEval$ = Format(A1, "0")
   FormRef.cmdConfigure.Caption = "k" & EnableEval$
   FormRef.cmdConfigure = True

End Sub

Private Sub mnuScripter_Click()

   mnuScripter.Checked = Not mnuScripter.Checked
   If mnuScripter.Checked Then
      frmScript.Show
   Else
      gnUniScript = False
      mfmUniTest.Show
      mfmUniTest.cmdUtils.Caption = "Uti&lities"
      mfmUniTest.picCommands.Visible = True
      frmMain.mnuScripter.Checked = False
      Unload frmScript
   End If

End Sub

Private Function ParseScript(ByVal ScriptType As Integer, ByVal LineNumber As Integer, _
ByRef Args As Variant, Optional LineText As Variant) As Integer

   Dim ScriptItem() As Variant
   
   If IsMissing(LineText) Then
      If ScriptType = mMASTER Then
         ScriptLine$ = masMaster(LineNumber)
      Else
         ScriptLine$ = masScript(LineNumber)
      End If
   Else
      ScriptLine$ = LineText
   End If
   
   FirstChar$ = Left(ScriptLine$, 1)
   ConditionStatus% = (mnCurConditional = WAITINGENDIF) Or _
   (mnCurConditional = WAITINGENDDO) Or (mnCurConditional = WAITINGENDFOR)
   If FirstChar$ = ";" Then
      If ConditionStatus% Then
         Exit Function
      End If
      If Len(ScriptLine$) > 1 Then ScriptLine$ = Mid(ScriptLine$, 2)
      VarsInComment = Split(ScriptLine$, "{")
      NumCommentVars& = UBound(VarsInComment)
      For VarSeg& = 0 To NumCommentVars&
         'contains math expression
         LineSeg$ = VarsInComment(VarSeg&)
         MathEnd& = InStr(LineSeg$, "}")
         If MathEnd& > 0 Then
            'get portion of string containing math expression
            Expression$ = Left(LineSeg$, MathEnd& - 1)
            result = ParseMathExpr(Expression$)
            LineEnd$ = ""
            If Len(LineSeg$) > MathEnd& Then LineEnd$ = Mid(LineSeg$, MathEnd& + 1)
            If IsNumeric(result) Then
                LineSeg$ = Format(result, "General Number") & LineEnd$
            Else
                LineSeg$ = result & LineEnd$
            End If
         End If
         NewLine$ = NewLine$ & LineSeg$
      Next
      
      Args = NewLine$
      ReturnVal% = mCOMMENT
      If gnShowComments Then frmScriptInfo.txtScriptInfo.Text = frmScriptInfo.txtScriptInfo.Text & NewLine$ & vbCrLf
      If gnLogComments And mnLogOpen Then Print #8, NewLine$
   ElseIf FirstChar$ = "'" Then
      'Check if conditional, otherwise, do nothing
      NoOpLine$ = LCase(Mid(ScriptLine$, 2))
      If (mnInvalidEnd > 0) Then
         'ignore all lines except 'for' and 'next'
         TrimLine$ = Left(NoOpLine$, 4)
         If Not ((TrimLine$ = "for ") Or (TrimLine$ = "next")) Then
            Exit Function
         End If
      End If
      If Not (NoOpLine$ = "end") Then
         FirstChar$ = Left(NoOpLine$, 1)
         GetOut% = (FirstChar$ = ";") Or (NoOpLine$ = "")
         If GetOut% Then Exit Function
         mnCurConditional = SetCondition(NoOpLine$, ScriptType)
      Else
         If ScriptType = mMASTER Then
            'mnMasterLine = mnNumMasterLines - 1
            If mnCurConditional = 1 Then
               ReturnVal% = mCOMMENT
            Else
               cmdScript_Click mSTOP
               ReturnVal% = mINVALID
            End If
         Else
            If mnCurConditional = 1 Then
               ReturnVal% = mCOMMENT
            Else
               mnScriptLine = mnNumScriptLines - 1
               mnIfStarts = 0
            End If
         End If
      End If
   Else
      If ConditionStatus% Then
         Exit Function
      End If
      If mnCurConditional = WAITINGENDFOR Then
         If mnInvalidEnd > 0 Then Exit Function
      End If
      StartPos% = 1
      Do
         EndPos% = InStr(StartPos%, ScriptLine$, ",")
         ReDim Preserve ScriptItem(i%)
         If EndPos% > 0 Then
            ScriptArg = Mid$(ScriptLine$, StartPos%, EndPos% - StartPos%)
         Else
            ScriptArg = Mid$(ScriptLine$, StartPos%)
         End If
         MathStart& = InStr(ScriptArg, "{")
         If Not (MathStart& = 0) Then
            'contains math expression
            MathBegArray = Split(ScriptArg, "{")
            NumExpressions& = UBound(MathBegArray)
            InterrimRes$ = ""
            For Expr& = 0 To NumExpressions& '- 1
               Front$ = ""
               Back$ = ""
               MathResult$ = ""
               ThisSegment$ = MathBegArray(Expr&)
               If InStr(ThisSegment$, "}") > 0 Then
                  MathEndArray = Split(ThisSegment$, "}")
                  Expression$ = MathEndArray(0)
                  result = ParseMathExpr(Expression$)
                  MathResult$ = Format(result, "General Number")
                  If UBound(MathEndArray) > 0 Then
                     Back$ = MathEndArray(1)
                  End If
               Else
                  If (Expr& = 0) Then Front$ = ThisSegment$
                  If Expr& = NumExpressions& Then Back$ = ThisSegment$
               End If
               InterrimRes$ = InterrimRes$ & Front$ & MathResult$ & Back$
            Next
            If IsNumeric(InterrimRes$) Then
               ScriptItem(i%) = result
            Else
               ScriptItem(i%) = InterrimRes$
            End If
            If 0 Then
               'begin commented code
               MathEnd& = InStr(ScriptArg, "}")
               If MathStart > 1 Then
                  'get portion of string in front of math expression
                  Front$ = Left(ScriptArg, MathStart& - 1)
                  If Front$ = " " Then Front$ = ""
               End If
               If MathEnd& < Len(ScriptArg) Then
                  'get portion of string after math expression
                  Back$ = Mid(ScriptArg, MathEnd& + 1)
               End If
               Expression$ = Mid(ScriptArg, MathStart& + 1, _
               MathEnd& - (MathStart& + 1))
               result = ParseMathExpr(Expression$)
               If (Front$ = "") And (Back$ = "") Then
                  ScriptItem(i%) = result
               Else
                  ScriptItem(i%) = Front$ & Format(result, "General Number") & Back$
               End If
               'end commented code
            End If
         Else
            If InStr(1, ScriptArg, "|") Then
               ScriptItem(i%) = ""
            Else
               ScriptItem(i%) = ScriptArg
            End If
         End If
         If EndPos% > 0 Then StartPos% = EndPos% + 1
         i% = i% + 1
      Loop While EndPos% > 0
      If i% > 1 Then
         ReDim Preserve ScriptItem(i%)
         ScriptItem(i%) = ""  'Mid$(ScriptLine$, StartPos%)
         Args = ScriptItem
         ReturnVal% = mCOMMAND
      Else
         ReturnVal% = mINVALID
      End If
   End If
   If (ReturnVal% = mCOMMAND) And (i% > 1) Then
      If Val(ScriptItem(1)) > 9999 Then ReturnVal% = mEVALUATE
   End If
   ParseScript = ReturnVal%

End Function

Private Function ParseMathExpr(MathExpr As String) As Variant
   
   Dim GroupVal() As Double
   
   'following block is for info in case of error
   ScriptLine& = mnScriptLine
   FaultScript$ = msScriptPath
   If (ScriptLine& = 0) Then
      ScriptLine& = mnMasterLine + mnHeaderLines
      PathArray = Split(msMasterPath, "\")
      PathElements& = UBound(PathArray)
      FaultScript$ = PathArray(PathElements&)
   End If

   'separate and resolve grouped expressions
   NumSpecOps& = FindInString(MathExpr, "Abs(", OpsLocs)
   'to do - implement absolute value
   NumGroups& = FindInString(MathExpr, "(", Locations)
   EndGroups& = FindInString(MathExpr, ")", EndLocs)
   TempExpr$ = MathExpr
   If Not (NumGroups& = EndGroups&) Then
      MsgBox "Invalid math expression: '" & MathExpr & "'." & vbCrLf & _
      "Check script '" & FaultScript$ & "' line " & Format(ScriptLine&, "0") _
      & ".", vbCritical, "Script Error"
      gnErrFlag = True
      Exit Function
   End If
   
   Do While NumGroups > -1
      Group& = NumGroups
      For CloseParen& = 0 To EndGroups&
         If EndLocs(CloseParen&) > Locations(NumGroups) Then
            PairLoc& = CloseParen&
            Exit For
         End If
      Next
      IntExpr$ = Mid(TempExpr$, Locations(Group&) + 1, _
      EndLocs(PairLoc&) - (Locations(Group&) + 1))
      GoSub ResolveGroup
      ReDim Preserve GroupVal(SolvedGroups&)
      GroupVal(SolvedGroups&) = Resolution
      SolvedGroups& = SolvedGroups& + 1
      BeforeParens$ = ""
      If Locations(Group&) > 1 Then BeforeParens$ = Left(TempExpr$, Locations(Group&) - 1)
      AfterParens$ = ""
      If EndLocs(PairLoc&) < Len(TempExpr$) Then _
      AfterParens$ = Right(TempExpr$, Len(TempExpr$) - (EndLocs(PairLoc&)))
      If FmtString$ = "" Then FmtString$ = "General Number"
      TempExpr$ = BeforeParens$ & Format(Resolution, FmtString$) & AfterParens$
      NumGroups& = FindInString(TempExpr$, "(", Locations)
      EndGroups& = FindInString(TempExpr$, ")", EndLocs)
   Loop
   IntExpr$ = TempExpr$
   GoSub ResolveGroup
   ParseMathExpr = Resolution
   Exit Function
   
ResolveGroup:
   
   'get value of any variables
   Elements = Split(IntExpr$, " ")
   NumElements& = UBound(Elements)
   ReDim NewExpression(NumElements&)
   ReDim Operators(NumElements&) As Integer
   ReDim VarVal(0)
   For i& = 0 To NumElements&
      If IsNumeric(Elements(i&)) Then
         If InStr(Elements(i&), ".") = 0 Then
            NewExpression(i&) = CLng(Elements(i&))
         Else
            NewExpression(i&) = Val(Elements(i&))
         End If
      Else
         CurElement = LCase(Elements(i&))
         Select Case CurElement
            Case "*", "/"
               'second priority
               NewExpression(i&) = CurElement
               Operators(i&) = 2
            Case "\", "/\"
               'third priority
               NewExpression(i&) = CurElement
               Operators(i&) = 3
            Case "+", "-"
               'last priority (besides logic)
               NewExpression(i&) = CurElement
               Operators(i&) = 6
            Case "^"
               'first priority
               NewExpression(i&) = CurElement
               Operators(i&) = 1
            Case "div"
               'fourth priority
               NewExpression(i&) = CurElement
               Operators(i&) = 4
            Case "mod"
               'fifth priority
               NewExpression(i&) = CurElement
               Operators(i&) = 5
            Case "not"
               NegateBool% = True
            Case "="
               NewExpression(i&) = CurElement
               Operators(i&) = 7
            Case "<"
               NewExpression(i&) = CurElement
               Operators(i&) = 8
            Case ">"
               NewExpression(i&) = CurElement
               Operators(i&) = 9
            Case ">="
               NewExpression(i&) = CurElement
               Operators(i&) = 10
            Case "<="
               NewExpression(i&) = CurElement
               Operators(i&) = 11
            Case "or"
               NewExpression(i&) = CurElement
               Operators(i&) = 12
            Case "and"
               NewExpression(i&) = CurElement
               Operators(i&) = 13
            Case Else
               VarVal(0) = CurElement
               VarFound% = CheckForVariables(VarVal)
               If VarFound% Then
                  If IsNumeric(VarVal(0)) Then
                     If InStr(VarVal(0), ".") = 0 Then
                        'If FmtString$ = "" Then FmtString$ = "0"
                        FmtString$ = "0"
                     Else
                        FmtString$ = "0.0######"
                     End If
                     TempExpString$ = Format(VarVal(0), FmtString$)
                     VarVal(0) = Val(TempExpString$)
                  End If
                  If FmtString$ = "0" Then
                     NewExpression(i&) = CLng(VarVal(0))
                  Else
                     NewExpression(i&) = VarVal(0)
                  End If
               Else
                  VariableName$ = VarVal(0)
                  If Left(VariableName$, 8) = "fudgeadd" Then
                     'set the value to the default additive "0"
                     NewExpression(i&) = 0
                     FmtString$ = "0"
                     FudgeDefault% = True
                  End If
                  If Left(VariableName$, 9) = "fudgemult" Then
                     'set the value to the default multiplier "1"
                     NewExpression(i&) = 1
                     FmtString$ = "0"
                     FudgeDefault% = True
                  End If
                  If Left(VariableName$, 9) = "fudgebool" Then
                     'set the value to the default boolean "0"
                     NewExpression(i&) = 0
                     FmtString$ = "0"
                     FudgeDefault% = True
                  End If
                  If Left(VariableName$, 9) = "criticalw" Then
                     'set the value to the default string ""
                     NewExpression(i&) = ""
                     FmtString$ = ""
                     FudgeDefault% = True
                  End If
                  If Not FudgeDefault% Then
                     If VariableName$ = "~" Then
                        NewExpression(i&) = "~"
                     Else
                        MsgBox "Uninitialized variable or invalid " & _
                        "math expression in script (" & VariableName$ & ")." & vbCrLf & _
                        "Check script '" & FaultScript$ & "' line " & Format(ScriptLine&, "0") _
                        & ".", vbCritical, "Script Error"
                        gnErrFlag = True
                        Exit Function
                     End If
                  End If
               End If
         End Select
      End If
   Next
   
   'execute math operations in appropriate order
   If NumElements& = 0 Then
      'this shoud be a simple variable converted to a value
      result = NewExpression(0)
      Resolution = result
      Return
   End If
   ErrorExp& = -1
   If NumElements& = 1 Then
      'only one numeric - must be bool op
      If NegateBool% Then
         result = Not NewExpression(1)
      End If
   Else
      For Priority% = 1 To 13
         For Element& = 1 To NumElements& - 1
            StringCompare% = False
            CompareStrings% = False
            If Operators(Element&) = Priority% Then
               If Not IsNumeric(NewExpression(Element& + 1)) Then
                  ErrorExp& = (Element& + 1)
                  StringCompare% = True
               End If
               If Not IsNumeric(NewExpression(Element& - 1)) Then
                  ErrorExp& = (Element& - 1)
                  CompareStrings% = True
               End If
               If CompareStrings% Or StringCompare% Then
                  ValidOpString$ = NewExpression(Element&)
                  Select Case ValidOpString$
                     Case "=", "<", ">", ">=", "<="
                        ValidOperator% = True
                     Case Else
                        ValidOperator% = False
                  End Select
               End If
               If 1 Then
                  If (Not ErrorExp& < 0) Then
                     If Not ValidOperator% Then
                        MsgBox "Math operation attempted on non-numeric expression '" & _
                        NewExpression(ErrorExp&) & "'." & vbCrLf & "Try adding curly " & _
                        "braces (""{}"") to script '" & FaultScript$ & "' line " & _
                        Format(ScriptLine&, "0") & ".", vbCritical, "Invalid Script Expression"
                        gnErrFlag = True
                        Resolution = 0
                        Exit Function
                     End If
                  End If
                  FmtString$ = "0.0######"
                  NotFloat% = VarType(NewExpression(Element& - 1)) < vbSingle
                  NotFloat% = NotFloat% And VarType(NewExpression(Element& + 1)) < vbSingle
                  'NotFloat% = Not NotFloat%
               End If
               Select Case NewExpression(Element&)
                  Case "*"
                     result = NewExpression(Element& - 1) * NewExpression(Element& + 1)
                     If result = Fix(result) Then
                        If NotFloat% Then FmtString$ = "0"
                     End If
                  Case "+"
                     result = NewExpression(Element& - 1) + NewExpression(Element& + 1)
                     If result = Fix(result) Then
                        If NotFloat% Then FmtString$ = "0"
                     End If
                  Case "-"
                     result = NewExpression(Element& - 1) - NewExpression(Element& + 1)
                     If result = Fix(result) Then
                        If NotFloat% Then FmtString$ = "0"
                     End If
                  Case "/"
                     result = NewExpression(Element& - 1) / NewExpression(Element& + 1)
                     If result = Fix(result) Then
                        If NotFloat% Then FmtString$ = "0"
                     End If
                  Case "/\"
                     'divide and round up
                     dblValue# = NewExpression(Element& - 1) / NewExpression(Element& + 1)
                     Roundup& = Fix(dblValue# - (dblValue# \ 1 <> dblValue#))
                     FmtString$ = "0"
                     result = Roundup&
                  Case "^"
                     result = NewExpression(Element& - 1) ^ NewExpression(Element& + 1)
                     If result = Fix(result) Then
                        If NotFloat% Then FmtString$ = "0"
                     End If
                  Case "\"
                     If NewExpression(Element& + 1) = 1 Then
                        result = Int(NewExpression(Element& - 1) + 0.0001)
                     Else
                        result = NewExpression(Element& - 1) \ NewExpression(Element& + 1)
                     End If
                     FmtString$ = "0"
                  Case "div"
                     IntResult& = NewExpression(Element& - 1) / NewExpression(Element& + 1)
                     result = IntResult&
                     FmtString$ = "0"
                  Case "mod"
                     result = NewExpression(Element& - 1) Mod NewExpression(Element& + 1)
                     FmtString$ = "0"
                  Case "="
                     result = (NewExpression(Element& - 1) = NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
                  Case "<"
                     result = (NewExpression(Element& - 1) < NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
                  Case ">"
                     result = (NewExpression(Element& - 1) > NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
                  Case ">="
                     result = (NewExpression(Element& - 1) >= NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
                  Case "<="
                     result = (NewExpression(Element& - 1) <= NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
                  Case "or"
                     result = (NewExpression(Element& - 1) Or NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
                  Case "and"
                     result = (NewExpression(Element& - 1) And NewExpression(Element& + 1))
                     If NegateBool% Then result = Not result
                     FmtString$ = "0"
               End Select
               NewExpression(Element& - 1) = Format(result, FmtString$)
               NewExpression(Element& + 1) = Format(result, FmtString$)
            End If
         Next
      Next
   End If
   NegateBool% = 0
   Resolution = Format(result, FmtString$)
   Return
   
End Function

Private Sub mnuScriptDir_Click()

   Dim result As VbMsgBoxResult
   sDefault$ = msScriptDir
   ScriptDir$ = InputBox("Enter location of subscripts", "Subscript Directory", sDefault$)
   If ScriptDir$ = "" Then
      result = MsgBox("Are you sure you want to set the default directory to """"?", vbYesNo, "Remove Default Setting?")
      If result = vbNo Then
         MsgBox "Script directory unchanged. (" & msScriptDir & ").", vbOKOnly, "No Changes Made"
         Exit Sub
      End If
   End If
   If Not msScriptDir = ScriptDir$ Then
      If Not (ScriptDir$ = "") Then
         If Not Right(ScriptDir$, 1) = "\" Then ScriptDir$ = ScriptDir$ & "\"
      End If
      lpFileName$ = "UniTest.ini"
      lpApplicationName$ = "ScriptDir"
      lpKeyName$ = "ScriptStorage"
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, ScriptDir$, lpFileName$)
      msScriptDir = ScriptDir$
   End If

End Sub

Private Sub mnuMasterDir_Click()

   sDefault$ = msBoardDir
   MasterDir$ = InputBox("Enter location of master scripts", "Master Script Directory", sDefault$)
   'If MasterDir$ = "" Then Exit Sub
   If MasterDir$ = "" Then
      result = MsgBox("Are you sure you want to set the default directory to """"?", vbYesNo, "Remove Default Setting?")
      If result = vbNo Then
         MsgBox "Master directory unchanged. (" & msBoardDir & ").", vbOKOnly, "No Changes Made"
         Exit Sub
      End If
   End If
   If Not msBoardDir = MasterDir$ Then
      If Not (MasterDir$ = "") Then
         If Not Right(MasterDir$, 1) = "\" Then MasterDir$ = MasterDir$ & "\"
      End If
      lpFileName$ = "UniTest.ini"
      lpApplicationName$ = "ScriptDir"
      lpKeyName$ = "MasterStorage"
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, MasterDir$, lpFileName$)
      msBoardDir = MasterDir$
   End If

End Sub

Public Sub WaitForIdle(FormRef As Form, WaitForCount As Long, IdleTimeout As Long)

   If Not FormRef Is Nothing Then
      Set mfrmFormRef = FormRef
      mnCheckingStatus = True
      mlWaitForCount = WaitForCount
      mlTimeout = IdleTimeout
   Else
      MsgBox "Check for status not associated with valid form. ", vbOKOnly, "No Form Selected"
      gnScriptRun = False
   End If
   
End Sub

Public Sub WaitStatusChange(FormRef As Form, StopDelta As Long, WaitCondition As Long, StatusTimeout As Long)

   If Not FormRef Is Nothing Then
      Set mfrmFormRef = FormRef
      mnWaitingStatVal = True
      mlEventType = StopDelta
      mlEventData = WaitCondition
      mlTimeout = StatusTimeout
   Else
      MsgBox "Wait for status change not associated with valid form. ", vbOKOnly, "No Form Selected"
      gnScriptRun = False
   End If

End Sub

Public Sub WaitForEvent(FormRef As Form, EventType As Long, EventData As Long, EventTimeout As Long)

   If Not FormRef Is Nothing Then
      Set mfrmFormRef = FormRef
      mnWaitingEvent = True
      mlEventType = EventType
      mlEventData = EventData
      mlTimeout = EventTimeout
   Else
      MsgBox "Wait for event not associated with valid form. ", vbOKOnly, "No Form Selected"
      gnScriptRun = False
   End If
   
End Sub

Private Function SetCondition(NoOpLine As String, ScriptType As Integer) As Integer

   Dim FormRef As Form
   
   ParsedLine = Split(NoOpLine, " ", 6)
   TermsInExpression& = UBound(ParsedLine)
   Conditional$ = LCase(ParsedLine(0))
   If Len(Conditional$) > 1 Then
      Select Case Conditional$
         Case "do"
            If (mnCurConditional = WAITINGENDIF) Or (mnCurConditional = WAITINGENDDO) Then
               SetCondition = mnCurConditional
               Exit Function
            End If
            If TermsInExpression& > 0 Then
               LoopDef$ = LCase(ParsedLine(1))
               If LoopDef$ = "while" Then
                  'store loop condition
                  If msDoCondition = "" Then
                     CondStart& = InStr(1, NoOpLine, LoopDef$, vbTextCompare) + 6
                     LoopDef$ = Mid(NoOpLine, CondStart&)
                     msDoCondition = LCase(LoopDef$)
                  End If
                  Satisfied% = EvalCondition(msDoCondition)
                  If Negate% Then Satisfied% = (Not Satisfied%)
                  If Satisfied% Then
                     result% = INDOWHILELOOP
                     mnDoStart = mnScriptLine
                  Else
                     result% = WAITINGENDDO
                     mlScriptRate = Me.tmrScript.Interval
                     Me.tmrScript.Interval = 100
                  End If
               End If
            Else
               mnDoStart = mnScriptLine
               result% = INDOLOOP
            End If
         Case "end"
            If TermsInExpression& > 0 Then
               Terminator$ = LCase(ParsedLine(1))
               If Terminator$ = "if" Then
                  mnIfStarts = mnIfStarts - 1
                  If mnIfStarts < 0 Then
                     MsgBox "'End If' without a corresponding 'If' in script.", _
                     vbCritical, "Invalid Script Syntax"
                     gnErrFlag = True
                     gnScriptRun = False
                     Exit Function
                  End If
                  If mnCurConditional = WAITINGENDIF Then
                     If mnIfStarts < mnIfEnds Then
                        tmrScript.Interval = mlScriptRate
                        mnIfEnds = 0
                        result% = 0
                     Else
                        result% = WAITINGENDIF
                     End If
                  Else
                     If mnIfStarts = 0 Then mnIfEnds = 0
                     If (mnCurConditional = WAITINGENDDO) Then
                        SetCondition = mnCurConditional
                        Exit Function
                     End If
                  End If
                  'maForNest(5, mnForNest) = result%
               End If
            End If
         Case "for"
            If (mnCurConditional = WAITINGENDIF) Or (mnCurConditional = WAITINGENDDO) Then
               SetCondition = mnCurConditional
               Exit Function
            End If
            If mnCurConditional = WAITINGENDIF Then
               result% = WAITINGENDIF
            Else
               LoopDef$ = LCase(ParsedLine(1))
               LowCnt$ = ParsedLine(3)
               HighCnt$ = ParsedLine(5)
               If Not IsNumeric(LowCnt$) Or Not IsNumeric(HighCnt$) Then
                  'check for variables
                  ReDim Args(0)
                  If IsNumeric(LowCnt$) Then
                     mlLowCtr = Val(LowCnt$)
                  Else
                     Args(0) = LowCnt$
                     If Not CheckForVariables(Args) Then
                        NoValidArgs% = True
                     Else
                        mlLowCtr = Val(Args(0))
                     End If
                  End If
                  If IsNumeric(HighCnt$) Then
                     mlHighCtr = Val(HighCnt$)
                  Else
                     Args(0) = HighCnt$
                     If Not CheckForVariables(Args) Then
                        NoValidArgs% = True
                     Else
                        mlHighCtr = Val(Args(0))
                     End If
                  End If
                  If NoValidArgs% Then
                     result% = 0
                  Else
                     result% = INFORLOOP
                     mnForStart = mnScriptLine
                  End If
               Else
                  mlLowCtr = Val(LowCnt$)
                  mlHighCtr = Val(HighCnt$)
               End If
               If mlHighCtr < mlLowCtr Then
                  result% = WAITINGENDFOR
               Else
                  If mnInvalidEnd = 0 Then
                     result% = INFORLOOP
                  Else
                     result% = WAITINGENDFOR
                  End If
               End If
               If Not (result% = 0) Then
                  If ScriptType = mMASTER Then
                     mnForStart = mnMasterLine
                  Else
                     mnForStart = mnScriptLine
                  End If
                  mnForNest = mnForNest + 1
                  If UBound(maForNest, 2) < mnForNest Then
                     ReDim Preserve maForNest(5, mnForNest)
                  End If
                  maForNest(0, mnForNest) = ScriptType
                  maForNest(1, mnForNest) = mlLowCtr
                  maForNest(2, mnForNest) = mlHighCtr
                  maForNest(3, mnForNest) = mnForStart
                  If (mlHighCtr < mlLowCtr) Or (mnInvalidEnd > 0) Then
                     'Loop is invalid
                     mnInvalidEnd = mnInvalidEnd + 1
                  End If
               End If
               If Not (mnForNest < 0) Then
                  LoopVariable% = SetVariable(LoopDef$, mlLowCtr)
                  maForNest(4, mnForNest) = LoopVariable%
                  maForNest(5, mnForNest) = result%
                  Me.lblLoopStat.Visible = Not (result% = 0)
               End If
            End If
         Case "if"
            mnIfStarts = mnIfStarts + 1
            If (mnCurConditional = WAITINGENDIF) Or (mnCurConditional = WAITINGENDDO) Then
               SetCondition = mnCurConditional
               Exit Function
            End If
            Condition$ = LCase(ParsedLine(1))
            ExprStart& = InStr(1, NoOpLine, ParsedLine(1))
            If Condition$ = "not" Then
               Negate% = True
               Condition$ = LCase(ParsedLine(2))
               ExprStart& = InStr(1, NoOpLine, ParsedLine(2))
            End If
            If Condition$ = "staticoption" Then
               'line must contain the form ID as "(###)"
               If TermsInExpression& > 1 Then
                  FormIDString$ = ParsedLine(2)
                  If Not InStr(FormIDString$, "(") = 0 Then
                     FormID$ = Mid(FormIDString$, 2, 3)
                     GotForm% = GetFormReference(FormID$, FormRef)
                     If GotForm% Then mlStaticOpt = FormRef.GetStaticOption
                     ArgIndex% = TermsInExpression&
                     Condition$ = LCase(ParsedLine(ArgIndex%))
                  End If
               End If
            Else
               Condition$ = Mid(NoOpLine, ExprStart&)
            End If
            Satisfied% = EvalCondition(Condition$)
            If Negate% Then Satisfied% = (Not Satisfied%)
            If Not Satisfied% Then
               If mnIfEnds = 0 Then mnIfEnds = mnIfStarts
               result% = WAITINGENDIF
               mlScriptRate = Me.tmrScript.Interval
               Me.tmrScript.Interval = 100
            End If
         Case "loop"
            Select Case mnCurConditional
               Case INDOLOOP
                  'if in the loop, check condition to exit
                  If msDoCondition = "" Then
                     'if loopwhile condition is not set
                     CondStart& = InStr(1, NoOpLine, "while", vbTextCompare) + 6
                     LoopDef$ = Mid(NoOpLine, CondStart&)
                     msDoCondition = LCase(LoopDef$)
                  End If
                  Satisfied% = EvalCondition(msDoCondition)
                  If Negate% Then Satisfied% = (Not Satisfied%)
                  If Satisfied% Then
                     mnScriptLine = mnDoStart
                     result% = INDOLOOP
                  Else
                     result% = 0
                     msDoCondition = ""
                     mnDoStart = 0
                  End If
               Case INDOWHILELOOP
                  'loop again - condition evaluated at start
                  mnScriptLine = mnDoStart - 1
                  result% = INDOWHILELOOP
               Case WAITINGENDDO
                  Me.tmrScript.Interval = mlScriptRate
                  mnDoStart = 0
                  msDoCondition = ""
                  result% = 0
            End Select
         Case "next"
            If (mnCurConditional = WAITINGENDIF) Or (mnCurConditional = WAITINGENDDO) Then
               SetCondition = mnCurConditional
               Exit Function
            Else
               If mnCurConditional = WAITINGENDFOR Then
                  If mnInvalidEnd > 0 Then
                     result% = WAITINGENDFOR
                     mnInvalidEnd = mnInvalidEnd - 1
                     maForNest(5, mnForNest) = 0
                     mnForNest = mnForNest - 1
                     If mnInvalidEnd = 0 Then
                        result% = 0
                     End If
                  Else
                     result% = 0
                     maForNest(5, mnForNest) = result%
                  End If
               Else
                  If (mnForNest < 0) Then
                     MsgBox "'Next' without a corresponding 'For' in script.", _
                     vbCritical, "Invalid Script Syntax"
                     gnErrFlag = True
                     gnScriptRun = False
                     Exit Function
                  Else
                     HighCtr& = maForNest(2, mnForNest)
                     ForCountIndex% = maForNest(4, mnForNest)
                     ForCount& = ScriptVars(1, ForCountIndex%)
                     If ForCount& >= HighCtr& Then
                        result% = 0
                        mnForNest = mnForNest - 1
                        If Not (mnForNest < 0) Then
                           LowCtr& = maForNest(1, mnForNest)
                           ScriptVars(1, ForCountIndex%) = LowCtr&
                        End If
                        Me.lblLoopStat.Visible = False
                        Me.lblLoopStat.Caption = ""
                     Else
                        'mlForCount = mlForCount + 1
                        ForCount& = ForCount& + 1
                        ScriptVars(1, ForCountIndex%) = ForCount&
                        CurScriptType% = maForNest(0, mnForNest)
                        ForStart% = maForNest(3, mnForNest)
                        If CurScriptType% = mMASTER Then
                           mnMasterLine = ForStart%
                        Else
                           'mnScriptLine = mnForStart
                           mnScriptLine = ForStart%
                        End If
                        Me.lblLoopStat.Caption = Format(ForCount&, "0") & _
                        " of " & Format(HighCtr& + 1, "0")
                     End If
                  End If
               End If
            End If
         Case Else
            result% = mnCurConditional
      End Select
   End If
   SetCondition = result% 'Or mnCurConditional
   
End Function

Private Function EvalCondition(Condition As String) As Integer
   
   MathStart& = InStr(Condition, "{")
   If Not (MathStart& = 0) Then
      'contains math expression
      MathEnd& = InStr(Condition, "}")
      MathLength& = MathEnd& - MathStart&
      If MathStart > 1 Then
         'get portion of string in front of math expression
         Front$ = Left(Condition, MathStart& - 1)
      End If
      If MathEnd& < Len(Condition) Then
         'get portion of string after math expression
         Back$ = Mid(Condition, MathEnd& + 1)
      End If
      Expression$ = Mid(Condition, MathStart& + 1, MathLength& - 1)
      MathResult# = ParseMathExpr(Expression$)
      Condition = Front$ & Format(MathResult#, "General Number") & Back$
   End If
   ExprElements = Split(Condition, " ")
   ElementsInExpr& = UBound(ExprElements)
   If ElementsInExpr& = 0 Then
      'simple boolean value (not compared)
      FoundVars% = CheckForVariables(ExprElements)
      FirstExpr$ = ExprElements(0)
      If IsNumeric(FirstExpr$) Then
         ArgVal1 = Val(FirstExpr$)
         Satisfied% = Not (ArgVal1 = 0)
      Else
         Satisfied% = False
      End If
      EvalCondition = Satisfied%
      Exit Function
   End If
   Equal% = Not (InStr(1, Condition, "=") = 0)
   Greater% = Not (InStr(1, Condition, ">") = 0)
   Less% = Not (InStr(1, Condition, "<") = 0)
   If ElementsInExpr& = 2 Then
      FoundVars% = CheckForVariables(ExprElements)
      'parse out units
      FirstExpr$ = ExprElements(0)
      SecondExpr$ = ExprElements(2)
      ExprElements(0) = ParseUnits(FirstExpr$, UnitType%)
      ExprElements(2) = ParseUnits(SecondExpr$, UnitType%)
      'check if string is compared to numeric
      If Not (IsNumeric(ExprElements(0)) = IsNumeric(ExprElements(2))) Then
         Satisfied% = False
         EvalCondition = Satisfied%
         Exit Function
      Else
         'convert non-string data to numeric
         If IsNumeric(ExprElements(0)) Then
            ArgVal1 = Val(ExprElements(0))
         Else
            TempVal1 = LCase(ExprElements(0))
            ArgVal1 = Replace(TempVal1, """", "")
         End If
         If IsNumeric(ExprElements(2)) Then
            ArgVal2 = Val(ExprElements(2))
         Else
            TempVal2 = LCase(ExprElements(2))
            ArgVal2 = Replace(TempVal2, """", "")
         End If
      End If
   End If
   If Equal% Or Greater% Or Less% Then
      Select Case True
         Case Equal%
            Satisfied% = (ArgVal1 = ArgVal2)
         Case Greater%
            Satisfied% = (ArgVal1 > ArgVal2)
         Case Less%
            Satisfied% = (ArgVal1 < ArgVal2)
         Case Else
            Satisfied% = 0
      End Select
   Else
      'true / false tests
      Select Case Condition
         Case "adcclock"
            Satisfied% = ((mlStaticOpt And ADCCLOCK) = ADCCLOCK)
         Case "adcclocktrig"
            Satisfied% = ((mlStaticOpt And ADCCLOCKTRIG) = ADCCLOCKTRIG)
         Case "background"
            Satisfied% = ((mlStaticOpt And BACKGROUND) = BACKGROUND)
         Case "blockio"
            Satisfied% = ((mlStaticOpt And BLOCKIO) = BLOCKIO)
         Case "burstio"
            Satisfied% = ((mlStaticOpt And BURSTIO) = BURSTIO)
         Case "burstmode"
            Satisfied% = ((mlStaticOpt And BURSTMODE) = BURSTMODE)
         Case "continuous"
            Satisfied% = ((mlStaticOpt And CONTINUOUS) = CONTINUOUS)
         Case "convertdata"
            Satisfied% = ((mlStaticOpt And CONVERTDATA) = CONVERTDATA)
         Case "dmaio"
            Satisfied% = ((mlStaticOpt And DMAIO) = DMAIO)
         Case "extclock"
            Satisfied% = ((mlStaticOpt And EXTCLOCK) = EXTCLOCK)
         Case "exttrigger"
            Satisfied% = ((mlStaticOpt And EXTTRIGGER) = EXTTRIGGER)
         Case "nocalibratedata"
            Satisfied% = ((mlStaticOpt And NOCALIBRATEDATA) = NOCALIBRATEDATA)
         Case "nofilter"
            Satisfied% = ((mlStaticOpt And NOFILTER) = NOFILTER)
         Case "nonstreamedio"
            Satisfied% = ((mlStaticOpt And NONSTREAMEDIO) = NONSTREAMEDIO)
         Case "retrigmode"
            Satisfied% = ((mlStaticOpt And RETRIGMODE) = RETRIGMODE)
         Case "scaledata"
            Satisfied% = ((mlStaticOpt And SCALEDATA) = SCALEDATA)
         Case "siminput"
            Satisfied% = (mnSimIn = True)
         Case "simultaneous"
            Satisfied% = ((mlStaticOpt And SIMULTANEOUS) = SIMULTANEOUS)
         Case "singleio"
            Satisfied% = ((mlStaticOpt And SINGLEIO) = SINGLEIO)
         Case "wordxfer"
            Satisfied% = ((mlStaticOpt And WORDXFER) = WORDXFER)
         Case "dwordxfer"
            Satisfied% = ((mlStaticOpt And DWORDXFER) = DWORDXFER)
         Case "ulerror"
            Satisfied% = (Not mlULError = 0)
            mlULError = 0
         Case Else
            ReDim CondVal(0)
            CondVal(0) = Condition
            VarFound% = CheckForVariables(CondVal)
            If VarFound% Then Satisfied% = Not (CondVal(0) = 0)
      End Select
   End If
   EvalCondition = Satisfied%

End Function

Private Function GetFormProps(FormRef As Form, PropName As String) As Variant

   PropVal = FormRef.GetFormProperty(PropName)
   If PropVal = "Invalid" Then
      MsgBox PropName & " is not a known form property.", _
      vbOKOnly, "Invalid Form Property"
      gnErrFlag = True
   End If
   Select Case PropName
      Case "resolution"
         'PropVal = GetResolution(FormRef)
      Case "siminput"
         mnSimIn = PropVal 'FormRef.GetInScanConfig()
      Case Else
   End Select
   GetFormProps = PropVal
   
End Function

Public Function SetVariable(VariableName As String, VarValue As Variant) As Integer

   ReDim Preserve ScriptVars(1, mnNumScriptVars)
   
   If VariableName = "" Then
      MsgBox "Attempt to set a value to a variable with no name.", vbCritical, "Bad Variable Name"
      SetVariable = 0
      Exit Function
   End If
   If IsNumeric(VarValue) Then
      If InStr(VarValue, ".") = 0 Then
         'ValueToSet = CLng(VarValue)
         ValueToSet = Val(VarValue)
      Else
         ValueToSet = VarValue
      End If
   Else
      ValueToSet = Trim(VarValue)
   End If
   
   For VarNum% = 0 To mnNumScriptVars - 1
      Variable$ = ScriptVars(0, VarNum%)
      If VariableName = Variable$ Then
         ScriptVars(1, VarNum%) = ValueToSet
         SetVariable = VarNum%
         Exit Function
      End If
   Next
   If IsNumeric(VariableName) Then
      SetVariable = -1
   Else
      ScriptVars(0, mnNumScriptVars) = VariableName
      ScriptVars(1, mnNumScriptVars) = ValueToSet
      SetVariable = mnNumScriptVars
      mnNumScriptVars = mnNumScriptVars + 1
   End If
   
End Function

Public Function CheckForVariables(Args As Variant) As Integer

   'Dim NotMath As Boolean
   Dim TempArgs() As Variant
   NumTempArgs% = -1
   'NotMath = False
   
   For VarNum% = 0 To mnNumScriptVars - 1
      Variable$ = ScriptVars(0, VarNum%)
      If Not Variable$ = "" Then
         For Arg% = 0 To UBound(Args)
            Expression$ = ""
            Argument$ = LCase(Trim(Args(Arg%)))
            LiteralSpecStart& = InStr(1, Argument$, "[")
            If Not (LiteralSpecStart& > 0) Then
               MathStart& = InStr(Argument$, "{")
               MathEnd& = InStr(Argument$, "}")
               If Not (MathStart& = 0) Then
                  Expression$ = Mid(Argument$, MathStart& + 1, _
                  MathEnd& - (MathStart& + 1))
               End If
               ArgComponents = Split(Expression$, " ")
               If UBound(ArgComponents) > 1 Then
                  'check for math expression within argument
                  Argument$ = LCase(Trim(ArgComponents(0)))
                  MathFunc$ = ArgComponents(1)
                  MathArg# = Val(ArgComponents(2))
                  Select Case MathFunc$
                     Case "*"
                        MathOp% = 1
                     Case "/"
                        MathOp% = 2
                        If MathArg# = 0 Then
                           MsgBox "Invalid divisor used with script variable.", _
                           vbCritical, "Scripted Math Error"
                           MathOp% = 0
                        End If
                     Case "\"
                        MathOp% = 3
                        If MathArg# = 0 Then
                           MsgBox "Invalid divisor used with script variable.", _
                           vbCritical, "Scripted Math Error"
                           MathOp% = 0
                        End If
                     Case "+"
                        MathOp% = 4
                     Case "-"
                        MathOp% = 5
                     Case Else
                        MathOp% = 0
                  End Select
               Else
                  If Len(Expression$) > 0 Then
                     Argument$ = LCase(Trim(Expression$))
                  Else
                     Argument$ = LCase(Trim(Args(Arg%)))
                  End If
               End If
               If Variable$ = Argument$ Then
                  If LiteralSpecStart& = 1 Then
                     ArgVal = Mid(Argument$, LiteralSpecStart& + 1, _
                     Len(Argument$) - (LiteralSpecStart& + 1))
                  Else
                     If IsNumeric(ScriptVars(1, VarNum%)) Then
                        If InStr(ScriptVars(1, VarNum%), ".") = 0 Then
                           ArgVal = CLng(ScriptVars(1, VarNum%))
                        Else
                           ArgVal = ScriptVars(1, VarNum%)
                        End If
                     Else
                        ArgVal = ScriptVars(1, VarNum%)
                     End If
                     Select Case MathOp%
                        Case 0
                           Args(Arg%) = ArgVal
                           'NotMath = True
                        Case 1
                           Args(Arg%) = ArgVal * MathArg#
                        Case 2
                           Args(Arg%) = ArgVal / MathArg#
                        Case 3
                           Args(Arg%) = ArgVal \ MathArg#
                        Case 4
                           Args(Arg%) = ArgVal + MathArg#
                        Case 5
                           Args(Arg%) = ArgVal - MathArg#
                     End Select
                  End If
                  VarExists% = True
                  'If NotMath Then Exit For
               End If
               MathOp% = 0
            Else
               'Don't change the value, just pass back without braces
               Argument$ = Mid(Argument$, LiteralSpecStart& + 1, Len(Argument$) - (LiteralSpecStart& + 1))
               'Args(Arg%) = Argument$
               NumTempArgs% = NumTempArgs% + 1
               ReDim Preserve TempArgs(1, NumTempArgs%)
               TempArgs(0, NumTempArgs%) = Arg%
               TempArgs(1, NumTempArgs%) = Argument$
               Literal% = True
               'Exit For
            End If
         Next
      End If
   Next
   
   For CurTempArg% = 0 To NumTempArgs%
      ArgNum% = TempArgs(0, CurTempArg%)
      Argument$ = TempArgs(1, CurTempArg%)
      Args(ArgNum%) = Argument$
   Next
   
   CheckForVariables = VarExists%
   
End Function

Private Function SelectValFromList(SelectionList As Variant) As Variant
   
   SelStart& = InStr(SelectionList, "|")
   If Not SelStart& > 0 Then
      SelectValFromList = ""
      Exit Function
   End If
   ReDim WhichVar(0)
   WhichVar(0) = Left(SelectionList, SelStart& - 1)
   ValueList$ = Mid(SelectionList, SelStart& + 1)
   If IsNumeric(WhichVar(0)) Then
      Element& = WhichVar(0)
   Else
      FoundVariable% = CheckForVariables(WhichVar)
      If FoundVariable% Then Element& = WhichVar(0)
   End If
   Values = Split(ValueList$, ";")
   LastItem& = UBound(Values) + 1
   
   If (Element& < 1) Or (Element& > LastItem&) Then
      'don't set value to variable
      SelectValFromList = "_"
   Else
      SelectValFromList = Values(Element& - 1)
   End If

End Function

Public Sub SetULError(ErrorVal As Long)

   mlULError = ErrorVal
   
End Sub

Private Sub ReadStringList(Filename As String, VarName As String)

   On Error GoTo AltStringListPath
   
   TryFile$ = Filename
   If Not TryFile$ = "" Then
      Open TryFile$ For Input As #40
      ReDim Preserve maStringList(0)
      mnNumStrings = -1
      Do While Not EOF(40)
         Line Input #40, A1$
         If Not (A1$ = "") Then
            mnNumStrings = mnNumStrings + 1
            ReDim Preserve maStringList(mnNumStrings)
            maStringList(mnNumStrings) = A1$
         End If
      Loop
      Close #40
   End If
   VarSet% = SetVariable(VarName, mnNumStrings)
   Exit Sub
   
AltStringListPath:
   MasterPath$ = msBoardDir
   If (Attempts% = 3) Or (Attempts% = 0) Then MasterPath$ = msDefaultPath
   TryFile$ = LocateScriptFile(TryFile$, msScriptDir, MasterPath$, Attempts%)
   If TryFile$ = "" Then
      If msScriptDir = "" Or msMasterPath = "" Then Hint$ = _
      " Try setting the Script or Master Directories under the File menu."
      MsgBox "File not found in the following locations:" & vbCrLf & vbCrLf & _
      PathsChecked$ & vbCrLf & Hint$, , "Error Opening Script File"
      gnErrFlag = True
      Exit Sub
   Else
      PathsChecked$ = PathsChecked$ & TryFile$ & vbCrLf
      Resume 0
   End If

End Sub

Private Sub ReadCSVList(Filename As String, VarNameListSize As String, _
VarNameNumValues As String, Optional ListName As Variant)

   On Error GoTo AltCSVListPath
   Dim CSVList() As String, CSVParamArray() As String
   Dim NewList As New CParamList
   
   TryFile$ = Filename
   If Not TryFile$ = "" Then
      Open TryFile$ For Input As #40
      ReDim CSVList(0)
      NumCSVArgs% = -1
      Do While Not EOF(40)
         Line Input #40, A1$
         If Not (A1$ = "") Then
            NumCSVArgs% = NumCSVArgs% + 1
            ReDim Preserve CSVList(NumCSVArgs%)
            CSVList(NumCSVArgs%) = A1$
         End If
      Loop
      Close #40
   End If
   If Not NumCSVArgs% < 0 Then
      FirstLine$ = CSVList(0)
      ArgList = Split(FirstLine$, ",")
      ListSize% = UBound(ArgList)
      ReDim CSVParamArray(NumCSVArgs%, ListSize%)
      For CSVLine% = 0 To NumCSVArgs%
         CurLine$ = CSVList(CSVLine%)
         ArgList = Split(CurLine$, ",")
         For ArgVal% = 0 To ListSize%
            StringValue$ = ArgList(ArgVal%)
            CSVParamArray(CSVLine%, ArgVal%) = StringValue$
         Next
      Next
      For Each ListObject In mParamList
         If ListObject.ListName = ListName Then
            'remove
            mParamList.Remove (ListName)
            Exit For
         End If
      Next
      'If ListName = "" Then ListName = 1
      NewList.SetList CSVParamArray()
      NewList.ListName = ListName
      mParamList.Add NewList, ListName
      Set NewList = Nothing
      VarSet% = SetVariable(VarNameListSize, NumCSVArgs%)
      If VarSet% = -1 Then
         If Not ListName = "" Then NameOfList$ = vbCrLf & "List name = '" & ListName & "'. "
         MsgBox "Attempted to set the Numeric name '" & VarNameListSize & _
         "' to the NumValuesVariable argument by the SLoadCSVList function." & _
         NameOfList$ & bvCrLf & "You may have duplicate variable names or " & _
         "square brackets are needed around the variable name.", _
         vbExclamation, "Attempt to Set Numeric Variable Name"
      End If
      VarSet% = SetVariable(VarNameNumValues, ListSize%)
      If VarSet% = -1 Then
         If Not ListName = "" Then NameOfList$ = vbCrLf & "List name = '" & ListName & "'. "
         MsgBox "Attempted to set the Numeric name '" & VarNameNumValues & _
         "' to the NumValuesVariable argument by the SLoadCSVList function." & _
         NameOfList$ & bvCrLf & "You may have duplicate variable names or " & _
         "square brackets are needed around the variable name.", _
         vbExclamation, "Attempt to Set Numeric Variable Name"
      End If
   End If
   Exit Sub
   
AltCSVListPath:
   MasterPath$ = msBoardDir
   If (Attempts% = 3) Or (Attempts% = 0) Then MasterPath$ = msDefaultPath
   TryFile$ = LocateScriptFile(TryFile$, msScriptDir, MasterPath$, Attempts%)
   If TryFile$ = "" Then
      If msScriptDir = "" Or msMasterPath = "" Then Hint$ = _
      " Try setting the Script or Master Directories under the File menu."
      MsgBox "File not found in the following locations:" & vbCrLf & vbCrLf & _
      PathsChecked$ & vbCrLf & Hint$, , "Error Opening Script File"
      gnErrFlag = True
      Exit Sub
   Else
      PathsChecked$ = PathsChecked$ & TryFile$ & vbCrLf
      Resume 0
   End If

End Sub


Private Sub ScrCopyFile(FileSource As String, FileDestination As String)

   On Error GoTo AltSCFListPath
   
   TryFile$ = FileSource
   If Not TryFile$ = "" Then
      FileCopy TryFile$, FileDestination
   End If
   Exit Sub
   
AltSCFListPath:
   MasterPath$ = msBoardDir
   Dim MResult As VbMsgBoxResult
   If Err = 71 Then
      'drive not ready
      MResult = MsgBox("Cannot write " & FileDestination & " to destination.", vbRetryCancel, "Disk Not Ready")
      If MResult = vbCancel Then
         gnErrFlag = True
         Exit Sub
      Else
         Resume 0
      End If
   End If
   If Err = 76 Then
      'path not found
      If (InStr(1, TryFile$, "\") = 0) Then
         MResult = MsgBox("Cannot write " & FileDestination & " to destination.", vbRetryCancel, "Path Not Found")
         If MResult = vbCancel Then
            gnErrFlag = True
            Exit Sub
         Else
            Resume 0
         End If
      End If
   End If
   If (Attempts% = 3) Or (Attempts% = 0) Then MasterPath$ = msDefaultPath
   TryFile$ = LocateScriptFile(TryFile$, msScriptDir, MasterPath$, Attempts%)
   If TryFile$ = "" Then
      If msScriptDir = "" Or msMasterPath = "" Then Hint$ = _
      " Try setting the Script or Master Directories under the File menu."
      MsgBox "File not found in the following locations:" & vbCrLf & vbCrLf & _
      PathsChecked$ & vbCrLf & Hint$, , "Error Opening Script File"
      gnErrFlag = True
      Exit Sub
   Else
      PathsChecked$ = PathsChecked$ & TryFile$ & vbCrLf
      Resume 0
   End If

End Sub

Private Sub WaitForAppClose(AppID As Variant)

   Dim lngHandle As Long
   Dim lngExitCode As Long
   
   If AppID <> 0 Then
      lngHandle = OpenProcess(SYNCHRONIZE Or PROCESS_QUERY_INFORMATION, 0, AppID)
      If lngHandle <> 0 Then
         Screen.MousePointer = vbHourglass
         RetVal = WaitForSingleObject(lngHandle, INFINITE)
         CloseHandle lngHandle
         Screen.MousePointer = vbDefault
      End If
   End If
   
End Sub

Private Function GetListItem(ByVal ListIndex As Long, _
ByVal ParamList As String, ByRef ListSize As Long, _
ByRef ListLength As Long, ByVal DefaultValue As String) As Variant

   ListArray = Split(ParamList$, ";")
   If ParamList$ = "~" Then
      ListSize& = -1
      GetListItem = "~"
      Exit Function
   End If
   'ListLength&
   ListSize& = UBound(ListArray)
   If ListSize& = 0 Then
      If Not DefaultValue = "" Then
         ValidDefault% = Not IsNumeric(ListArray(0))
         ValidDefault% = ValidDefault% And IsNumeric(DefaultValue)
         If ValidDefault% Then
            GetListItem = Val(DefaultValue)
            Exit Function
         End If
      End If
   End If
   If (ListIndex& + ListLength) > ListSize& Then ListIndex = ListSize&
   For ItemInList& = 0 To ListLength&
      CurListItem = CurListItem & ListArray(ListIndex + ItemInList&)
      If Not ItemInList& = ListLength& Then CurListItem = CurListItem & ";"
   Next
   GetListItem = CurListItem
   
End Function

Function NumericToDP800Cmd(ByVal Command As Single) As String

   Select Case Command
      Case Is >= 10
         VRange$ = "V2+"
         VValue$ = Format(Command * 10000, "0000000")
      Case Is >= 0
         VRange$ = "V1+"
         VValue$ = Format(Command * 100000, "0000000")
      Case Is <= -10
         VRange$ = "V2"
         VValue$ = Format(Command * 10000, "0000000")
      Case Is < 0
         VRange$ = "V1"
         VValue$ = Format(Command * 100000, "0000000")
   End Select
   NumericToDP800Cmd = VRange$ & VValue$

End Function

Function HP8112CmdToNumeric(ByVal Command As String) As Single

   CmdArray = Split(Command, " ")
   ArraySize& = UBound(CmdArray)
   Cmnd$ = CmdArray(0)
   If ArraySize& > 0 Then NewValue$ = CmdArray(1)
   If ArraySize& > 1 Then NewUnit$ = CmdArray(2)
   Select Case Cmnd$
      Case "PER"
         Divisor! = 1000000000
         Select Case NewUnit$
            Case "MS"
               Divisor! = 1000
            Case "S"
               Divisor! = 1
            Case "US"
               Divisor! = 1000000
         End Select
         TempVal! = Val(NewValue$)
         NumericVal! = TempVal! / Divisor!
   End Select
   HP8112CmdToNumeric = NumericVal!
   
End Function

Function ConvertRateToPer(ByVal Rate As String) As String

   RateVal! = Val(Rate)
   If RateVal! = 0 Then
      MsgBox "Rate evaluates to 0. ", vbCritical, "Error In Script"
      gnErrFlag = True
      Exit Function
   End If
   PeriodVal! = 1 / RateVal!
   Index% = 1
   Do While PeriodVal! < 1
      PeriodVal! = PeriodVal! * 1000
      Index% = Index% + 1
   Loop
   Suffix$ = Choose(Index%, " S", " MS", " US", " NS")
   PerString$ = Format(PeriodVal!, "0.0##")
   ConvertRateToPer = "PER " & PerString$ & Suffix$
   
End Function

Function ConvertTimeToWidth(ByVal Time As String) As String

   TimeVal! = Val(Time)
   If TimeVal! = 0 Then
      MsgBox "Time evaluates to 0. ", vbCritical, "Error In Script"
      gnErrFlag = True
      Exit Function
   End If
   Index% = 1
   Do While TimeVal! < 1
      TimeVal! = TimeVal! * 1000
      Index% = Index% + 1
   Loop
   Suffix$ = Choose(Index%, " S", " MS", " US", " NS")
   WidthString$ = Format(TimeVal!, "0.0##")
   ConvertTimeToWidth = "WID " & WidthString$ & Suffix$

End Function

Public Sub PostTimeoutStatus(ByVal TimeLeft As Long)

   ReDim A(4) As Variant
   A(0) = 0: A(1) = 2056: A(2) = 0
   A(3) = Str(mlStopCount): A(4) = Str(mlTimeout)
   Args = A()
   UpdateScriptStatus Args
   lblScriptStatus.Caption = lblScriptStatus.Caption & _
   "  (Resuming script in " & Format(TimeLeft, "0") & " seconds.)"

End Sub

Private Sub OpenTestOptionsForm()

   Dim DiffMode As Boolean
   Dim Args(0)
   Dim ComboControl As Control
   
   DefaultWidth& = frmTestOptions.Width
   MFileName$ = LCase(msMasterFile)
   frmTestOptions.UseDropCommand False
   If InStr(1, MFileName$, "startup.utm") > 0 Then
      CurGroupKey$ = "SOFTWARE\Measurement Computing\Universal Test Suite"
      KeyName$ = "ScriptPath"
      ProgExists% = GetRegGroup(HKEY_LOCAL_MACHINE, CurGroupKey$, hProgResult&)
      YN% = GetKeyValue(hProgResult&, KeyName$, KeyVal$)
      If YN% Then
         KeyName$ = "ScriptStorage"
         YN% = GetKeyValue(hProgResult&, KeyName$, StoreVal$)
         KeyName$ = "TestDir"
         YN% = GetKeyValue(hProgResult&, KeyName$, TestDir$)
         VarFound% = YN%
         Root$ = KeyVal$ '& ScriptDir$
         msBoardDir = Root$
         msScriptDir = Root$ & StoreVal$
      Else
         Root$ = msBoardDir   'Args(0)
         Args(0) = "TestDir"
         VarFound% = CheckForVariables(Args)
         If VarFound% Then
            frmTestOptions.fraSetup.Visible = True
         End If
      End If
   End If
   
   If VarFound% Then
      frmTestOptions.fraSetup.Visible = True
      Args(0) = "ProdGroupList"
      VarFound% = CheckForVariables(Args)
      If VarFound% Then
         ProdGroups$ = Args(0)
         frmTestOptions.SetPaths Root$, ProdGroups$
      End If
      frmTestOptions.SetTestDir TestDir$
      If mnStartCommand Then
         SelArray = Split(msMasterSelection, "\")
         SubDirs% = UBound(SelArray)
         Set ComboControl = frmTestOptions.cmbTestCat
         StdParam$ = SelArray(0)
         frmTestOptions.SetSelection ComboControl, StdParam$
         If SubDirs% > 0 Then
            Set ComboControl = frmTestOptions.cmbTest
            StdParam$ = SelArray(1)
            frmTestOptions.SetSelection ComboControl, StdParam$
         End If
         frmTestOptions.UseDropCommand True
      End If
      DoEvents
      frmTestOptions.Show 1
      msTestString = frmTestOptions.cmbTestCat.Text
      DevParms$ = frmTestOptions.lblDParmPath.Caption
      TestFile$ = frmTestOptions.lblTestFile.Caption
      ValidStartup% = Not ((DevParms$ = "") Or (TestFile$ = ""))
      MasterFile$ = frmTestOptions.lblFileName.Caption
      Unload frmTestOptions
      If ValidStartup% Then
         mnStartupScript = True
         msMasterPath = TestFile$
         msDevParmPath = DevParms$
         msScriptPath = TestFile$
         msMasterFile = MasterFile$ & ":  "
      End If
      Exit Sub
   End If
   
   Args(0) = "Device"
   VarFound% = CheckForVariables(Args)
   If VarFound% Then
      frmTestOptions.SSTab1.TabVisible(1) = False
      frmTestOptions.SSTab1.TabVisible(2) = False
      frmTestOptions.lblDeviceName = Args(0)
   
      Args(0) = "TestBlocks"
      VarFound% = CheckForVariables(Args)
      If VarFound% = 0 Then
         frmTestOptions.chkTest(0).Visible = False
      Else
         ParamList$ = Args(0)
         frmTestOptions.SetParamList ParamList$
      End If
      
      Args(0) = "parameteropts"
      VarFound% = CheckForVariables(Args)
      If VarFound% = 0 Then
         frmTestOptions.txtParam(0).Visible = False
         frmTestOptions.lblParam(0).Visible = False
      Else
         ParamList$ = Args(0)
         frmTestOptions.SetValueList ParamList$
      End If
      Args(0) = "UseDAQFlex"
      VarFound% = CheckForVariables(Args)
      If VarFound% = 0 Then
         frmTestOptions.chkUseDF.ENABLED = False
      Else
         DFVal% = Val(Args(0))
         If Not (DFVal% = 0) Then frmTestOptions.chkUseDF.value = 1
      End If
      TopOffset& = frmTestOptions.chkUseDF.Top
      If InStr(1, msTestString, "Digital") Then
         Args(0) = "DioPortList"
         VarFound% = CheckForVariables(Args)
         If VarFound% Then
            PortList$ = Args(0)
            PortArray = Split(PortList$, ";")
            NumPorts& = UBound(PortArray)
            PortNum& = Val(PortArray(0))
            PortName$ = GetPortString(PortNum&)
            frmTestOptions.chkSE(0).Caption = PortName$
            HorizPos& = frmTestOptions.chkSE(0).Left
            'TopOffset& = frmTestOptions.chkSE(0).Top
            VertPos& = 2
            For ArrayItem& = 1 To NumPorts&
               Load frmTestOptions.chkSE(ArrayItem&)
               TopVal& = (260 * VertPos&) + TopOffset&
               frmTestOptions.chkSE(ArrayItem&).Top = TopVal&
               VertPos& = VertPos& + 1
               PortNum& = Val(PortArray(ArrayItem&))
               PortName$ = GetPortString(PortNum&)
               frmTestOptions.chkSE(ArrayItem&).Left = HorizPos&
               frmTestOptions.chkSE(ArrayItem&).Caption = PortName$
               frmTestOptions.chkSE(ArrayItem&).Visible = True
               frmTestOptions.chkSE(ArrayItem&).value = 1
               If VertPos& = 6 Then
                  VertPos& = 0
                  HorizPos& = HorizPos& + 3000
                  If HorizPos& > (DefaultWidth& - 1500) Then Wider& = Wider& + 2400
               End If
            Next
            frmTestOptions.chkSE(0).value = 1
         Else
            frmTestOptions.chkSE(0).Visible = False
         End If
         frmTestOptions.txtHighChan.Visible = False
         frmTestOptions.txtLowChan.Visible = False
         frmTestOptions.lblHighChannel.Visible = False
         frmTestOptions.lblLowChan.Visible = False
      ElseIf InStr(1, msTestString, "Counter") Then
         Args(0) = "CtrTypeList"
         VarFound% = CheckForVariables(Args)
         If VarFound% Then
            MaxChanArray = Split(Args(0), ";")
            NumCtrs% = UBound(MaxChanArray)
            CtrType% = MaxChanArray(0)
            CtrName$ = GetCtrTypeString(CtrType%)
            frmTestOptions.txtLowChan.ENABLED = True
            frmTestOptions.txtLowChan.Text = "0"
            frmTestOptions.lblLowChan.Caption = CtrName$
            frmTestOptions.txtHighChan.Visible = False
            frmTestOptions.lblHighChannel.Visible = False
            frmTestOptions.chkSE(0).Visible = False
         End If
      Else
         'check if SE only
         Args(0) = "MaxAiChanList"
         VarFound% = CheckForVariables(Args)
         If VarFound% Then
            MaxChanArray = Split(Args(0), ";")
            If UBound(MaxChanArray) = 0 Then
               SingleEnded% = True
            Else
               If MaxChanArray(0) = -1 Then
                  SingleEnded% = True
                  VarSet% = frmScript.SetVariable("highchan", MaxChanArray(1))
               End If
               If MaxChanArray(1) = -1 Then
                  DiffMode = True
                  VarSet% = frmScript.SetVariable("highchan", MaxChanArray(0))
               End If
            End If
         End If
         If SingleEnded% Then
            frmTestOptions.chkSE(0).value = 1
            frmTestOptions.chkSE(0).ENABLED = False
            VarSet% = frmScript.SetVariable("usesemode", 1)
         ElseIf DiffMode Then
            frmTestOptions.chkSE(0).value = 0
            frmTestOptions.chkSE(0).ENABLED = False
            VarSet% = frmScript.SetVariable("usesemode", 0)
         Else
            Args(0) = "UseSEMode"
            VarFound% = CheckForVariables(Args)
            If VarFound% = 0 Then
               frmTestOptions.chkSE(0).ENABLED = False
            Else
               SEVal% = Val(Args(0))
               If Not (SEVal% = 0) Then frmTestOptions.chkSE(0).value = 1
            End If
         End If
         Args(0) = "HighChan"
         VarFound% = CheckForVariables(Args)
         If VarFound% = 0 Then
            frmTestOptions.txtHighChan.ENABLED = False
            frmTestOptions.lblHighChannel.ENABLED = False
         Else
            HCVal% = Val(Args(0))
            frmTestOptions.txtHighChan.Text = Format(HCVal%, "0")
         End If
         Args(0) = "LowChan"
         VarFound% = CheckForVariables(Args)
         If VarFound% = 0 Then
            frmTestOptions.txtLowChan.ENABLED = False
            frmTestOptions.lblLowChan.ENABLED = False
         Else
            HCVal% = Val(Args(0))
            frmTestOptions.txtLowChan.Text = Format(HCVal%, "0")
         End If
      End If
      DoEvents
      If Wider& > 0 Then
         frmTestOptions.Width = DefaultWidth& + Wider&
         frmTestOptions.fraCommon.Width = frmTestOptions.Width - 500
      End If
      frmTestOptions.Show 1
      Unload frmTestOptions
   End If
   
End Sub

Private Function CheckParamRevision(ByVal RevRequired As String, ByVal _
   RevInUse As String, ByVal CompareCondition As String) As Boolean

   Dim CheckFailed As Boolean
   Dim MsgResponse As VbMsgBoxResult
   
   Warning$ = ""
   RevReqStripped$ = Replace(RevRequired, ".", "")
   RevInUseStripped$ = Replace(RevInUse, ".", "")
   FillSize% = 4 - Len(RevInUseStripped$)
   If FillSize% > 0 Then
      FillString$ = Left("000", FillSize%)
   End If
   NumericRevInUse& = Val(RevInUseStripped$ & FillString$)
   FillSize% = 4 - Len(RevReqStripped$)
   If FillSize% > 0 Then
      FillString$ = Left("000", FillSize%)
   End If
   NumericRevReq& = Val(RevReqStripped$ & FillString$)
   
   ComparisonString$ = Left(CompareCondition, 2)
   Select Case ComparisonString$
      Case ">"
         CheckFailed = Not (NumericRevInUse& > NumericRevReq&)
      Case ">="
         CheckFailed = (NumericRevInUse& < NumericRevReq&)
   End Select
   
   CompLoc& = InStr(1, CompareCondition, "* ")
   If (CompLoc& > 0) And CheckFailed Then
      ComparisonItem$ = Mid(CompareCondition, CompLoc& + 2)
      CheckFailed = False
      Warning$ = "Update from " & RevInUse & " to " & _
         RevRequired & " containing " & ComparisonItem$ _
         & " recommended for this test."
   End If
   If Not (Warning$ = "") Then
      MsgResponse = MsgBox(Warning$ & vbCrLf & "Continue test anyway?", vbYesNo, "Parameter File Update Recommended")
      If MsgResponse = vbNo Then CheckFailed = True
   End If
   
   CheckParamRevision = CheckFailed
      
End Function
