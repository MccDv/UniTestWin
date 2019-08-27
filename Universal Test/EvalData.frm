VERSION 5.00
Begin VB.Form frmEvalData 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Select Type of Data Evaluation"
   ClientHeight    =   5310
   ClientLeft      =   735
   ClientTop       =   5460
   ClientWidth     =   9030
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
   ScaleHeight     =   5310
   ScaleWidth      =   9030
   Tag             =   "900"
   Begin VB.CheckBox chkSaveData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Save Error Data to File"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4740
      TabIndex        =   44
      Top             =   480
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.TextBox txtNumMsgSamps 
      Height          =   285
      Left            =   7500
      TabIndex        =   42
      Text            =   "100"
      ToolTipText     =   "Use -1 for all data to file."
      Top             =   840
      Width           =   1155
   End
   Begin VB.CheckBox chkShowMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show Data Available Message Box"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4740
      TabIndex        =   41
      Top             =   180
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.TextBox txtFirstPoint 
      Height          =   285
      Left            =   3060
      TabIndex        =   38
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtNumSamps 
      Height          =   285
      Left            =   3060
      TabIndex        =   37
      Text            =   "1000"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtChan 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3300
      TabIndex        =   19
      Text            =   "0"
      Top             =   4890
      Width           =   375
   End
   Begin VB.Frame fraDeltaT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Evaluate Time"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4320
      TabIndex        =   24
      Top             =   1260
      Width           =   4575
      Begin VB.TextBox txtTriggerLevel 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3540
         TabIndex        =   15
         Text            =   "2048"
         Top             =   1020
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   900
         TabIndex        =   26
         Top             =   900
         Width           =   2595
         Begin VB.OptionButton optDetectCycle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Crossing at:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optDetectCycle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Peak"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.OptionButton optCycleUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Frequency"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton optCycleUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Period"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   1035
      End
      Begin VB.OptionButton optCycleUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "# Samples"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox txtMaxWinT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Text            =   "100"
         Top             =   1560
         Width           =   915
      End
      Begin VB.TextBox txtMinWinT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Text            =   "50"
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkTWindow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkShowPeriod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Input Waveform Cycles by:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " < Delta <"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   1620
         Width           =   915
      End
   End
   Begin VB.CheckBox chkEnableEval 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enable Data Evaluation"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3840
      TabIndex        =   20
      Top             =   4860
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fraDeltaV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Evaluate Amplitude"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   23
      Top             =   1260
      Width           =   4095
      Begin VB.CheckBox chkSamplePairs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pair check"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2700
         TabIndex        =   50
         Top             =   1200
         Width           =   1275
      End
      Begin VB.TextBox txtWindowPercentage 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2340
         TabIndex        =   48
         Text            =   "5"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chkInWindow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ignore values outside"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1500
         Width           =   2235
      End
      Begin VB.TextBox txtDeltaVMin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   46
         Text            =   "1"
         Top             =   1200
         Width           =   915
      End
      Begin VB.CheckBox chkMinDelta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if Delta <"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   33
         Text            =   "2056"
         Top             =   1920
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   32
         Text            =   "2040"
         Top             =   1920
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CheckBox chkMinMaxStop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if value <"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtVMin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   2
         Text            =   "2040"
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txtVMax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Text            =   "2056"
         Top             =   600
         Width           =   915
      End
      Begin VB.CheckBox chkShowMinMax 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Min/Max"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1635
      End
      Begin VB.CheckBox chkMaxDelta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if Delta >"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   1575
      End
      Begin VB.TextBox txtDeltaVMax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1740
         TabIndex        =   5
         Text            =   "50"
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblFSCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "% FS count"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "< value <"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   1980
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblOr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "or >"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2700
         TabIndex        =   27
         Top             =   660
         Width           =   375
      End
   End
   Begin VB.Frame fraDeltaCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Evaluate Count"
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      TabIndex        =   29
      Top             =   3720
      Width           =   8775
      Begin VB.TextBox txtCountIterations 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   59
         Text            =   "1"
         Top             =   660
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   3240
         TabIndex        =   55
         Top             =   540
         Width           =   1815
         Begin VB.OptionButton optCountLast 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   ">"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   58
            Top             =   180
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optCountLast 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "="
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   57
            Top             =   180
            Width           =   495
         End
         Begin VB.OptionButton optCountLast 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "<"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   56
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Frame fraCountFirst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   960
         TabIndex        =   51
         Top             =   540
         Width           =   1935
         Begin VB.OptionButton optCountFirst 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   ">"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   54
            Top             =   180
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optCountFirst 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "="
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   53
            Top             =   180
            Width           =   495
         End
         Begin VB.OptionButton optCountFirst 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "<"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.TextBox txtCountDelta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7140
         TabIndex        =   36
         Text            =   "1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkCountDelta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if delta <"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5220
         TabIndex        =   35
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkCountStop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Stop if"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtCountTrig 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   7
         Text            =   "5000"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCountThen 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         TabIndex        =   8
         Text            =   "-1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCntIterations 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "iterations"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   60
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblThen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Then"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6600
      TabIndex        =   22
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   7920
      TabIndex        =   21
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblNumMsgSamps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Number of Error Samples to Store"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3900
      TabIndex        =   43
      Top             =   900
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "First Sample to Evaluate"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   40
      Top             =   540
      Width           =   2895
   End
   Begin VB.Label lblNumSamps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Number of Samples to Evaluate"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   39
      Top             =   180
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel to evaluate:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   660
      TabIndex        =   28
      Top             =   4920
      Width           =   2535
   End
End
Attribute VB_Name = "frmEvalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnCycleUnits As Integer
Dim mnEvaluate As Integer

Private Sub chkSaveData_Click()

   ShowPoints = (chkShowMsg.value = 1) Or (Me.chkSaveData.value = 1)
   txtNumMsgSamps.Visible = ShowPoints
   lblNumMsgSamps.Visible = ShowPoints

End Sub

Private Sub chkShowMsg_Click()

   ShowPoints = (chkShowMsg.value = 1) Or (Me.chkSaveData.value = 1)
   txtNumMsgSamps.Visible = ShowPoints
   lblNumMsgSamps.Visible = ShowPoints
   
End Sub

Private Sub cmdCancel_Click()

   Me.Hide

End Sub

Private Sub cmdOK_Click()

   Dim ShowMsg As Boolean, SaveFile As Boolean
   
   mnEvaluate = False
   chkEnableEval.value = 0
   ECount = (chkCountStop.value = 1)
   EDelta = (chkMaxDelta.value = 1)
   EMinDelta = (chkMinDelta.value = 1)
   ERange = (chkMinMaxStop.value = 1)
   EInWindow = (chkInWindow.value = 1)
   'EOutWindow = (Me.chkMinMaxStop.value = 1)
   
   If EDelta Or ERange Or EMinDelta Then
      mnEvaluate = True
      'SetDeltaStop Val(txtDeltaVMax.Text)
      'A1 = txtDeltaVMax.Text
      StartSamp& = Val(txtFirstPoint.Text)
      NumSamps& = Val(txtNumSamps.Text)
      NumMsgSamps& = Val(txtNumMsgSamps.Text)
      EvalChannel& = Val(txtChan.Text)
      ShowMsg = (chkShowMsg.value = 1)
      SaveFile = (chkSaveData.value = 1)
      SetEvalParams geSTART, StartSamp&
      SetEvalParams geNUMPOINTS, NumSamps&
      SetEvalParams geEVALCHAN, EvalChannel&
      SetEvalParams geSHOWPASTEMSG, ShowMsg
      SetEvalParams geNUMMSGSAMPS, NumMsgSamps&
      SetEvalParams geWRITEFILEERRS, SaveFile
      If EDelta Or EMinDelta Then
         DeltaVal& = Val(txtDeltaVMax.Text)
         MinDeltaVal& = Val(txtDeltaVMin.Text)
         WinVal& = Val(txtWindowPercentage.Text)
         SetEvalParams geEVALDELTA, mnEvaluate
         SetEvalParams geEVALMINDELTA, EMinDelta
         SetEvalParams geEVALINWINDOW, EInWindow
         If (chkMaxDelta.value = 1) Then _
            SetEvalParams geDELTAVAL, DeltaVal&
         If (chkMinDelta.value = 1) Then _
            SetEvalParams geDELTAMIN, MinDeltaVal&
         If (Me.chkInWindow.value = 1) Then _
            SetEvalParams geINWINDOW, WinVal&
      Else
         SetEvalParams geEVALINWINDOW, False
      End If
      If ERange Then
         LowVal& = Val(txtVMin.Text)
         HighVal& = Val(txtVMax.Text)
         SetEvalParams geEVALRANGE, mnEvaluate
         SetEvalParams geMINVAL, LowVal&
         SetEvalParams geMAXVAL, HighVal&
      Else
         SetEvalParams geEVALRANGE, False
      End If
   Else
      SetEvalParams geEVALDELTA, False
      SetEvalParams geEVALINWINDOW, False
   End If
   SetEvalParams -1, (chkSamplePairs.value = 1)
   'If gnScriptSave Then SaveEvalParams SEvalDelta, A1, A2, A3
   SetEvalParams geENABLEEVAL, mnEvaluate

   Me.Hide
   Exit Sub
   
   If EOutWindow Then
      mnEvaluate = True
      If fraDeltaT.ENABLED Then
         SetVMinMaxStop Val(txtVMin.Text), Val(txtVMax.Text)
         A1 = txtVMin.Text
         A2 = txtVMax.Text
      Else
         SetVfMinMaxStop Val(txtVMin.Text), Val(txtVMax.Text)
      End If
   Else
      SetVMinMaxStop -1, -1
      A1 = -1
      A2 = -1
   End If
   If gnScriptSave Then SaveEvalParams SEvalMaxMin, A1, A2, A3

   If txtChan.Text = "A" Then
   Else
      SetEvalChan Val(txtChan.Text)
      A1 = txtChan.Text
   End If
   If gnScriptSave Then SaveEvalParams SEvalChannel, A1, A2, A3
   
   SetShowMinMax (chkShowMinMax.value = 1)
   SetShowCycles (chkShowPeriod.value = 1)
   SetCycleUnits mnCycleUnits
   
   SetPeakTrigType (optDetectCycle(0).value)
   SetTrigValue Val(txtTriggerLevel.Text)
   SetCountTrigValue (txtCountTrig.Text)
   SetCountThenValue (txtCountThen.Text)
   SetCountDelta (txtCountDelta.Text)
   SetDeltaT Val(txtMinWinT.Text), Val(txtMaxWinT.Text)
   mnEvaluate = mnEvaluate Or (chkShowMinMax.value = 1)
   mnEvaluate = mnEvaluate Or (chkShowPeriod.value = 1)
   mnEvaluate = mnEvaluate Or (chkCountStop.value = 1)
   mnEvaluate = mnEvaluate Or (chkCountDelta.value = 1)

   If mnEvaluate = True Then chkEnableEval.value = 1
   ClearDetails
   Me.Hide

End Sub

Private Sub optCycleUnits_Click(Index As Integer)

   mnCycleUnits = Index

End Sub

Sub SaveEvalParams(FunctionType As Long, A1 As Variant, A2 As Variant, A3 As Variant)
   
   FuncStat = 0
   For ArgNum% = 1 To 14
      ArgVar = Choose(ArgNum%, Me.Tag, FunctionType, FuncStat, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
      If IsNull(ArgVar) Or IsEmpty(ArgVar) Then
         PrintString$ = PrintString$ & ", "
      Else
         PrintString$ = PrintString$ & Format$(ArgVar, "0") & ", "
      End If
   Next
   Print #2, PrintString$; Format$(AuxHandle, "0")

End Sub

Private Sub txtNumMsgSamps_Change()

   If txtNumMsgSamps.Text = "-1" Then
      Me.chkShowMsg.value = 0
      Me.chkSaveData.value = 1
      Me.chkShowMsg.ENABLED = False
   Else
      Me.chkShowMsg.ENABLED = True
   End If
   
End Sub
