VERSION 5.00
Begin VB.Form frmGuardBands 
   Caption         =   "Global Tweaks"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   19
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4740
      TabIndex        =   13
      Top             =   3300
      Width           =   975
   End
   Begin VB.Frame fraMCCSource 
      Caption         =   "MCC Source Tweaks"
      Height          =   1515
      Left            =   120
      TabIndex        =   7
      Top             =   3180
      Width           =   4395
      Begin VB.TextBox txtOffset 
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblDescSource 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Values entered here will be added to any offset value specified in a script."
         ForeColor       =   &H80000007&
         Height          =   555
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   3915
      End
      Begin VB.Label lblOSUnits 
         Caption         =   "Volts"
         Height          =   195
         Left            =   1860
         TabIndex        =   10
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblOffset 
         Alignment       =   1  'Right Justify
         Caption         =   "Offset"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   795
      End
   End
   Begin VB.Frame fraEvalGuards 
      Caption         =   "Evaluation Guardbands"
      Height          =   2955
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5595
      Begin VB.CheckBox chkSimOut 
         Caption         =   "Simultaneous Output"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3420
         TabIndex        =   17
         Top             =   1620
         Width           =   1935
      End
      Begin VB.CheckBox chkSimIn 
         Caption         =   "Simultaneous Input"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3420
         TabIndex        =   16
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtAverage 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Text            =   "0"
         Top             =   1980
         Width           =   495
      End
      Begin VB.ComboBox cmbApplyTo 
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox txtRate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Negative number specifies use of requested rate rather than returned rate."
         Top             =   1260
         Width           =   675
      End
      Begin VB.TextBox txtAmpl 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   2
         Text            =   "0"
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label lblAverage 
         Alignment       =   1  'Right Justify
         Caption         =   "Moving Average (X script value)"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   2040
         Width           =   2475
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Apply to:"
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label lblDescEval 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Guardbands.frx":0000
         ForeColor       =   &H80000007&
         Height          =   855
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   5235
      End
      Begin VB.Label lblRate 
         Alignment       =   1  'Right Justify
         Caption         =   "Rate (S/~)"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblAmpl 
         Alignment       =   1  'Right Justify
         Caption         =   "Amplitude (12-bit LSBs)"
         Height          =   195
         Left            =   540
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   4740
      TabIndex        =   0
      Top             =   4260
      Width           =   975
   End
End
Attribute VB_Name = "frmGuardBands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnEntryExists As Integer
Dim mlRateGB As Long, mlAmplGB As Long
Dim mlMvgAvgGB As Long, mfSrcOffset As Single
Dim mnDevChange As Integer, mnGlobalChange As Integer

Private Sub chkSimIn_Click()

   mnDevChange = True
   cmdApply.Enabled = True
   
End Sub

Private Sub chkSimOut_Click()

   mnDevChange = True
   cmdApply.Enabled = True
   
End Sub

Private Sub cmbApplyTo_Click()

   If Not (cmbApplyTo.ListIndex < 0) Then
      Me.txtAmpl.Enabled = True
      Me.txtAverage.Enabled = True
      Me.txtRate.Enabled = True
      Me.chkSimIn.Enabled = True
      Me.chkSimOut.Enabled = True
      BoardName$ = cmbApplyTo.Text
      Tweaks$ = GetBoardTweaks(BoardName$)
      If Len(Tweaks$) Then
         mnEntryExists = True
         x = Split(Tweaks$, ",")
         NumTweaks& = UBound(x)
         For Tweak& = 0 To NumTweaks&
            ThisTweak$ = x(Tweak&)
            Parameter = Split(ThisTweak$, "=")
            TweakName$ = Parameter(0)
            value$ = Parameter(1)
            Select Case TweakName$
               Case "RateGuard"
                  Me.txtRate.Text = value$
               Case "AmpGuard"
                  Me.txtAmpl.Text = value$
               Case "MvgAvg"
                  Me.txtAverage.Text = value$
               Case "SimOut"
                  Me.chkSimOut.value = Val(value$)
               Case "SimIn"
                  Me.chkSimIn.value = Val(value$)
            End Select
         Next
      Else
         mnEntryExists = False
         Me.txtRate.Text = "0"
         Me.txtAmpl.Text = "0"
         txtAverage.Text = "0"
         Me.chkSimIn.value = 0
         Me.chkSimOut.value = 0
      End If
   Else
      frmGuardBands.txtRate.Text = Format(mlRateGB, "0")
      frmGuardBands.txtAmpl.Text = Format(mlAmplGB, "0")
      frmGuardBands.txtOffset.Text = Format(mfSrcOffset, "0.0#####")
      frmGuardBands.txtAverage.Text = Format(mlMvgAvgGB, "0")
      Me.txtAmpl.Enabled = False
      Me.txtAverage.Enabled = False
      Me.txtRate.Enabled = False
      Me.chkSimIn.Enabled = False
      Me.chkSimOut.Enabled = False
   End If

End Sub

Private Sub cmdApply_Click()

   SetParams

End Sub

Private Sub cmdCancel_Click()

   cmdDone.Enabled = False
   Me.Hide
   
End Sub

Private Sub cmdDone_Click()
   
   If Me.cmdApply.Enabled Then SetParams
   Me.Hide
   
End Sub

Private Sub Form_Load()

   FillAppliesTo
   
End Sub

Private Sub FillAppliesTo()

   cmbApplyTo.AddItem "All Devices"
   ListFile$ = GetBoardFile()
   If Not ListFile$ = "" Then
      Open ListFile$ For Input As #4
      Do While Not EOF(4)
         Line Input #4, A1$
         cmbApplyTo.AddItem A1$
      Loop
      Close #4
   End If
   
   'cmbApplyTo.ListIndex = 0

End Sub

Public Function EntryExists() As Integer

   EntryExists = mnEntryExists
   
End Function

Public Sub SetRateGB(RateGB As Long)

   mlRateGB = RateGB
   frmGuardBands.txtRate.Text = Format(mlRateGB, "0")

End Sub

Public Sub SetAmplGB(Ampl As Long)
   
   mlAmplGB = Ampl
   frmGuardBands.txtAmpl.Text = Format(mlAmplGB, "0")

End Sub

Public Sub SetAvgGB(AvgVal As Long)
   
   mlMvgAvgGB = AvgVal
   frmGuardBands.txtAverage.Text = Format(mlMvgAvgGB, "0")

End Sub

Public Sub SetOffsetGB(Offset As Single)

   mfSrcOffset = Offset
   frmGuardBands.txtOffset.Text = Format(mfSrcOffset, "0.0#####")

End Sub

Sub SetParams()
   
   If mnGlobalChange Then
      'if global tweaks were changed
      mlRateGB = Val(txtRate.Text)
      mlAmplGB = Val(txtAmpl.Text)
      mlMvgAvgGB = Val(txtAverage.Text)
      mfSrcOffset = Val(txtOffset.Text)
      
      SetAmpGB mlAmplGB
      SetRateGB mlRateGB
      SetMvgAvgGB mlMvgAvgGB
      SetOffsetTweak mfSrcOffset
   End If
   
   If mnDevChange Then
      BoardAmpl& = Val(txtAmpl.Text)
      BoardRate& = Val(txtRate.Text)
      MovingAvg& = Val(txtAverage.Text)
      SimInput% = chkSimIn.value
      SimOutput% = chkSimOut.value
      ValuesSet% = Not ((BoardAmpl& = 0) And (BoardRate& = 0) And (MovingAvg& = 0))
      ValuesSet% = ValuesSet% Or Not ((SimInput% = 0) And (SimOutput% = 0))
      If Not EntryExists And Not ValuesSet% Then
         'if there's no ini entry for this device,
         'don't create one if parameters are zero
         If (BoardAmpl& = 0) And (BoardRate& = 0) And (MovingAvg& = 0) Then
            Unload frmGuardBands
            Exit Sub
         End If
      Else
         lpApplicationName$ = cmbApplyTo.Text
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

End Sub

Private Sub txtAmpl_Change()

   mnDevChange = True
   cmdApply.Enabled = True
   
End Sub

Private Sub txtAverage_Change()

   mnDevChange = True
   cmdApply.Enabled = True
   
End Sub

Private Sub txtOffset_Change()

   mnGlobalChange = True
   cmdApply.Enabled = True
   
End Sub

Private Sub txtRate_Change()

   mnDevChange = True
   cmdApply.Enabled = True
   
End Sub
