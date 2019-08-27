VERSION 5.00
Begin VB.Form frmBuildData 
   Caption         =   "Data Builder"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPercent 
      Caption         =   "Percentage of FS"
      Height          =   255
      Left            =   300
      TabIndex        =   28
      Top             =   1500
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.OptionButton optType 
      Caption         =   "Modify existing data"
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Text            =   "1000"
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkModExisting 
      Caption         =   "Modify existing data"
      Enabled         =   0   'False
      Height          =   255
      Left            =   300
      TabIndex        =   24
      Top             =   3300
      Width           =   1815
   End
   Begin VB.OptionButton optType 
      Caption         =   "Double"
      Height          =   315
      Index           =   5
      Left            =   1680
      TabIndex        =   23
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   3180
      Width           =   975
   End
   Begin VB.Frame fraMods 
      Caption         =   "Modify Existing Data"
      Height          =   1455
      Left            =   3300
      TabIndex        =   17
      Top             =   1020
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtChannel 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Text            =   "-1"
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox txtStart 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Text            =   "0"
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblChannel 
         Caption         =   "Channel (-1 = All)"
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblStart 
         Caption         =   "Starting at"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1635
      End
   End
   Begin VB.TextBox txtCycles 
      Height          =   315
      Left            =   3300
      TabIndex        =   15
      Text            =   "1"
      Top             =   600
      Width           =   795
   End
   Begin VB.OptionButton optType 
      Caption         =   "Single"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton optType 
      Caption         =   "DWord (Dec)"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   13
      Top             =   3000
      Width           =   1995
   End
   Begin VB.TextBox txtNumChans 
      Height          =   315
      Left            =   3660
      TabIndex        =   11
      Text            =   "1"
      Top             =   180
      Width           =   435
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2580
      TabIndex        =   10
      Top             =   3180
      Width           =   975
   End
   Begin VB.OptionButton optType 
      Caption         =   "Long"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   9
      Top             =   2700
      Width           =   1215
   End
   Begin VB.OptionButton optType 
      Caption         =   "Integer"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   8
      Top             =   2400
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox cmbSignal 
      Height          =   315
      ItemData        =   "BuildData.frx":0000
      Left            =   300
      List            =   "BuildData.frx":0002
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1995
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   6
      Top             =   3180
      Width           =   975
   End
   Begin VB.TextBox txtOffset 
      Enabled         =   0   'False
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Text            =   "1000"
      ToolTipText     =   "Use decimal point to specify voltage."
      Top             =   1140
      Width           =   1635
   End
   Begin VB.TextBox txtAmplitude 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Text            =   "1000"
      ToolTipText     =   "Use decimal point to specify voltage."
      Top             =   720
      Width           =   1635
   End
   Begin VB.TextBox txtSamples 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Text            =   "1000"
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label lblTimeOut 
      Caption         =   "Timeout"
      Height          =   195
      Left            =   4080
      TabIndex        =   26
      Top             =   2700
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblCycles 
      Caption         =   "Number of Cycles"
      Height          =   195
      Left            =   4170
      TabIndex        =   16
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label lblNumChans 
      Caption         =   "Number of Channels"
      Height          =   195
      Left            =   4170
      TabIndex        =   12
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label lblSamples 
      Caption         =   "Samples"
      Height          =   195
      Left            =   2100
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Amplitude"
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lblOffset 
      Caption         =   "Offset"
      Height          =   195
      Left            =   2100
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmBuildData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlSamples As Long, mlDataType As Long
Dim mvAmplitude As Variant
Dim mvOffset As Variant
Dim mnSignal As Integer
Dim mnNumChans As Integer
Dim mnChannel As Integer, mnSampleMultiplier As Integer
Dim mlFirstPoint As Long, mnCountIsPerChan As Integer
Dim mnCycles As Integer, mnUseWinAPI As Integer
Dim mfrmCallingForm As Form, mnCurResolution As Integer
Dim mnBoardNum As Integer, mnRange As Integer
Dim mnInitialized As Integer, mnGenFloats As Integer
Dim mnPercFS As Integer

Private Sub chkModExisting_Click()

   For i% = 0 To 5
      Me.optType(i%).ENABLED = Not (chkModExisting.value = 1)
   Next
   fraMods.Visible = (chkModExisting.value = 1)
   If (chkModExisting.value = 1) Then
      txtNumChans.Text = mnNumChans
      txtSamples.Text = mlSamples
      Me.cmdApply.ENABLED = True
   End If
   txtNumChans.ENABLED = Not (chkModExisting.value = 1)
   
End Sub

Private Sub chkPercent_Click()

   FSPerc% = mnPercFS
   mnPercFS = (chkPercent.value = 1)
   Amplitude$ = Me.txtAmplitude.Text
   Offset$ = Me.txtOffset.Text
   AmpMax! = GetRangeVolts(mnRange)
   
   If mnPercFS Then
      If InStr(Amplitude$, "%") = 0 Then
         'If Not InStr(Amplitude$, ".") = 0 Then
         '   EngUnits! = Val(Amplitude$)
         '   If Not mnUseWinAPI Then
         '      ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, IntAmp%)
         '      LongAmp& = IntValToULong(IntAmp%)
         '   Else
         '   End If
         '   Amplitude$ = Format(LongAmp&, "0")
         'End If
         If mnGenFloats Then
            FSFactor! = Val(Amplitude$) / AmpMax!
         Else
            FS& = Val(Amplitude$)
            FSFactor! = FS& / (2 ^ mnCurResolution)
         End If
         txtAmplitude.Text = Format(FSFactor! * 100, "0") & "%"
      End If
      If InStr(Offset$, "%") = 0 Then
         'If Not InStr(Offset$, ".") = 0 Then
         '   EngUnits! = Val(Offset$)
         '   If Not mnUseWinAPI Then
         '      ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, IntAmp%)
         '      LongOS& = IntValToULong(IntAmp%)
         '   Else
         '   End If
         '   Offset$ = Format(LongOS&, "0")
         'End If
         If mnGenFloats Then
            lOffset& = Val(Offset$)
            If mnRange < 100 Then
               OSFactor! = ((AmpMax! / 2) - lOffset&) / AmpMax!
            Else
               OSFactor! = (AmpMax! * OSFactor!)
            End If
            'OSFactor! = Val(Amplitude$) / AmpMax!
         Else
            os& = Val(Offset$)
            OSFactor! = os& / (2 ^ mnCurResolution)
         End If
         txtOffset.Text = Format(OSFactor! * 100, "0") & "%"
      End If
   Else
      If Not InStr(Amplitude$, "%") = 0 Then
         Perc$ = Left(Amplitude$, Len(Amplitude$) - 1)
         FSFactor! = Val(Perc$) / 100
         If mnGenFloats Then
            'data as percentage of max volts based on range
            AmpMax! = GetRangeVolts(mnRange) '/ 2
            'If mnRange < 100 Then
            '   mvAmplitude = (AmpMax! * FSFactor!) '/ 2
            'Else
            '   mvAmplitude = AmpMax! * FSFactor!
            'End If
            Me.txtAmplitude.Text = Format(AmpMax! * FSFactor!, "0")
         Else
            'data as percentage of max counts
            FS& = 2 ^ mnCurResolution
            Me.txtAmplitude.Text = Format(FS& * FSFactor!, "0")
         End If
      End If
      If Not InStr(Offset$, "%") = 0 Then
         Perc$ = Left(Offset$, Len(Offset$) - 1)
         OSFactor! = Val(Perc$) / 100
         If mnGenFloats Then
            'data as percentage of max volts based on range
            OSMax! = GetRangeVolts(mnRange)
            If mnRange < 100 Then
               fOffset! = (OSMax! / 2) - (OSMax! * OSFactor!)
            Else
               fOffset! = OSMax! * OSFactor! '(mvAmplitude / 2) + (
               'mvOffset = (OSMax! / 2) * FSFactor!
            End If
            Me.txtOffset.Text = Format(fOffset!, "0")
         Else
            'data as percentage of max counts
            os& = 2 ^ mnCurResolution
            Me.txtOffset.Text = Format(os& * OSFactor!, "0")
         End If
      End If
   End If
   
End Sub

Private Sub cmbSignal_Click()

   cmdApply.ENABLED = True
   Me.txtOffset.ENABLED = cmbSignal.ListIndex > 0
   
End Sub

Private Sub cmdApply_Click()

   Dim DataReady As Boolean
   'If Me.chkModExisting.value = 0 Then cmdApply.Enabled = False
   DataReady = GenData()
   
End Sub

Private Sub cmdCancel_Click()

   Me.txtSamples.Text = "0"
   Me.Hide
   
End Sub

Private Sub Form_GotFocus()

   Me.txtSamples.Text = Val(mlSamples)
   
End Sub

Private Sub Form_Load()

   mnPercFS = True
   frmBuildData.txtAmplitude.Text = "50%"
   frmBuildData.txtOffset.Text = "50%"
   frmBuildData.txtSamples.Text = 1000
   mlSamples = 0
   mnCycles = 0
   mnInitialized = False
   'frmBuildData.txtSamples.Text = Me.txtCount.Text

End Sub

Private Sub optType_Click(Index As Integer)

   cmdApply.ENABLED = True
   
End Sub

Private Sub cmdDone_Click()
   
   Dim DataReady As Boolean
   
   If Me.cmdApply.ENABLED Then
      DataReady = GenData()
      If DataReady Then Me.Hide
   Else
      Me.Hide
   End If
   'Unload Me
   
End Sub

Private Sub Form_Activate()

   If Not mnInitialized Then
      Me.cmbSignal.AddItem "DC Level"
      Me.cmbSignal.AddItem "Square Wave"
      Me.cmbSignal.AddItem "Sine Wave"
      Me.cmbSignal.AddItem "Ramp"
      Me.cmbSignal.AddItem "Triangle Wave"
      Me.cmbSignal.AddItem "Random Signal"
      Me.cmbSignal.ListIndex = 2
      
      If Not mlSamples = 0 Then
         Me.txtSamples.Text = Format(mlSamples, "0")
         Me.txtNumChans.Text = Format(mnNumChans, "0")
         Me.txtChannel.Text = Format(mnChannel, "0")
         If mnCycles = 0 Then
            mnCycles = 1
            mnSignal = 2
         End If
         Me.cmbSignal.ListIndex = mnSignal
         Me.txtCycles.Text = Format(mnCycles, "0")
         Select Case mlDataType
            Case 1
               Me.optType(0).value = True
            Case 2
               Me.optType(1).value = True
            Case 4
               'Stop
               Me.optType(4).value = True
            Case 6
               Me.optType(5).value = True
         End Select
      End If
      
      If Left(mfrmCallingForm.Caption, 7) = "Digital" Then
         Dim tempMidScale As Double
         tempMidScale = (2 ^ mnCurResolution) / 2
         Me.txtAmplitude.Text = Format(tempMidScale, "0")
         Me.txtOffset.Text = Format(tempMidScale, "0")
         Me.chkPercent.value = 0
      End If
      mnInitialized = True
   End If

End Sub

Public Function GetDataType() As Long

   GetDataType = mlDataType
   
End Function

Public Function GetNumSamples() As Long

   GetNumSamples = mlSamples
   
End Function

Public Sub SetFormRef(CallingForm As Form)

   Set mfrmCallingForm = CallingForm
   
End Sub

Public Sub SetDefaults(BoardNum As Integer, BoardResolution As Integer, _
BoardRange As Integer, Samples As Long, Optional NumChans As Variant, _
Optional DataType As Variant, Optional UseWinAPI As Integer = False, _
Optional CountIsPerChan As Integer = False, Optional GenFloats As _
Integer = False, Optional MsgData As Integer = False)

   Handle& = mfrmCallingForm.GetDataHandle(GENERATEDDATA, CurDataType&, NumSamples&)
   DatExist% = (Handle& > 0) And (Not Samples > NumSamples&) And (NumChans = mnNumChans)
   chkModExisting.ENABLED = DatExist% And Not MsgData
   chkModExisting.value = 0
   
   mnBoardNum = BoardNum
   mnCurResolution = BoardResolution
   mnRange = BoardRange
   If Not IsMissing(NumChans) Then mnNumChans = NumChans
   If Not IsMissing(DataType) Then mlDataType = DataType
   mnSampleMultiplier = 1
   mnCountIsPerChan = CountIsPerChan
   If CountIsPerChan Then mnSampleMultiplier = mnNumChans
   Me.txtSamples.Text = Samples
   mlSamples = Samples
   mnUseWinAPI = UseWinAPI
   mnGenFloats = GenFloats
   
End Sub

Private Function GenData() As Boolean

   mnNumChans = Val(Me.txtNumChans.Text)
   mnCycles = Val(Me.txtCycles.Text)
   If mnCountIsPerChan Then mnSampleMultiplier = mnNumChans
   If Not (chkModExisting.value = 1) Then
      FirstPoint& = 0
      mlSamples = Val(Me.txtSamples.Text)
      mnChannel = -1
      Samples& = mlSamples * mnSampleMultiplier
      NewData% = True
   Else
      mlFirstPoint = Val(Me.txtStart.Text)
      FirstPoint& = mlFirstPoint * mnSampleMultiplier
      Samples& = Val(Me.txtSamples.Text) * mnSampleMultiplier
      If ((mlFirstPoint + Samples&) > (mlSamples * mnSampleMultiplier)) Then
         MsgBox "Cannot increase the number of samples when modifying existing data.", _
         vbOKOnly, "Too Many Samples"
         Exit Function
      End If
      DataChan% = Val(Me.txtChannel.Text)
      If Not DataChan% < mnNumChans Then
         MsgBox "Channel to modify must be less than " & Format(mnNumChans, "0") & ".", vbCritical, "Bad Channel Number"
         GenData = False
         Exit Function
      End If
      mnChannel = Val(Me.txtChannel.Text)
   End If
   mnSignal = Me.cmbSignal.ListIndex
   For i% = 0 To 5
      If Me.optType(i%).value Then
         'mlDataType = i% + 1
         DataType% = i% + 1
         Exit For
      End If
   Next i%
   Amplitude$ = Me.txtAmplitude.Text
   Offset$ = Me.txtOffset.Text
   If Not InStr(Amplitude$, "%") = 0 Then
      Perc$ = Left(Amplitude$, Len(Amplitude$) - 1)
      FSFactor! = Val(Perc$) / 100
      FSPerc% = True
   Else
      If Not InStr(Amplitude$, ".") = 0 Then
         FSFloat% = True
      Else
         FSFloat% = False
      End If
      FSFactor! = 1
   End If
   If Not InStr(Offset$, "%") = 0 Then
      Perc$ = Left(Offset$, Len(Offset$) - 1)
      OSFactor! = Val(Perc$) / 100
      OSPerc% = True
   Else
      If Not InStr(Amplitude$, ".") = 0 Then
         OSFloat% = True
      Else
         OSFloat% = False
      End If
      OSFactor! = 1
   End If
   
   If Not FSPerc% Then
      'data as entered
      mvAmplitude = Val(Amplitude$)
   Else
      'data as percentage of FS (either volts or counts)
      If mnGenFloats Then
         'data as percentage of max volts based on range
         AmpMax! = GetRangeVolts(mnRange) '/ 2
         If mnRange < 100 Then
            mvAmplitude = (AmpMax! * FSFactor!) '/ 2
         Else
            mvAmplitude = AmpMax! * FSFactor!
         End If
      Else
         'data as percentage of max counts
         mvAmplitude = (2 ^ mnCurResolution) * FSFactor!
      End If
   End If
   If Not OSPerc% Then
      'data as entered
      mvOffset = Val(Offset$)
   Else
      'data as percentage of FS (either volts or counts)
      If mnGenFloats Then
         'data as percentage of max volts based on range
         OSMax! = GetRangeVolts(mnRange)
         If mnRange < 100 Then
            mvOffset = (OSMax! / 2) - (OSMax! * OSFactor!)
         Else
            mvOffset = OSMax! * OSFactor! '(mvAmplitude / 2) + (
            'mvOffset = (OSMax! / 2) * FSFactor!
         End If
      Else
         'data as percentage of max counts
         mvOffset = (2 ^ mnCurResolution) * OSFactor!
      End If
   End If
   
   Select Case DataType%
      Case 1
         mlDataType = DataType%
         If (mnCurResolution = 0) Or FSFloat% Then
            EngUnits! = mvAmplitude
            If Not mnUseWinAPI Then
               ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, IntAmp%)
               LongAmp& = IntValToULong(IntAmp%)
            Else
            End If
            mvAmplitude = LongAmp&
         End If
         If (mnCurResolution = 0) Or OSFloat% Then
            EngUnits! = mvOffset
            ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, IntOffset%)
            LongOffset& = IntValToULong(IntOffset%)
            mvOffset = LongOffset&
         End If
      Case 2
         mlDataType = DataType%
         If mnCurResolution = 0 Then
            MsgBox "No conversion from engineering units to long exists in library. Use counts.", _
            vbOKOnly, "Use Counts to Generate Long Data"
            Exit Function
         End If
      Case 4, 6
         mlDataType = DataType%
         If mnCurResolution = 0 Then
         End If
      Case 5
         'don't change data type - existing buffer being modified
         Select Case mlDataType
            Case 1
               If mnCurResolution = 0 Then
                  EngUnits! = mvAmplitude
                  ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, IntAmp%)
                  LongAmp& = IntValToULong(IntAmp%)
                  mvAmplitude = LongAmp&
                  EngUnits! = mvOffset
                  ULStat = cbFromEngUnits(mnBoardNum, mnRange, EngUnits!, IntOffset%)
                  LongOffset& = IntValToULong(IntOffset%)
                  mvOffset = LongOffset&
               End If
            Case 2
               If mnCurResolution = 0 Then
                  MsgBox "No conversion from engineering units to long exists in library. Use counts.", _
                  vbOKOnly, "Use Counts to Generate Long Data"
                  Exit Function
               End If
            Case 4, 6
         End Select
   End Select
   
   If mnGenFloats Then FloatMask% = &H10
   SigParam% = mnSignal Or FloatMask%
   DataHandle& = GenerateData(mlDataType, mnCycles, Samples&, _
   mnNumChans, mvAmplitude, mvOffset, SigParam%, NewData%, _
   mnChannel, FirstPoint&, mnUseWinAPI)
   mfrmCallingForm.SetupData DataHandle&, mlSamples, mnNumChans, mlDataType
   mfrmCallingForm.PlotGenData
   Me.chkModExisting.ENABLED = Not (DataHandle& = 0)
   GenData = True

End Function

Private Sub txtAmplitude_Change()

   cmdApply.ENABLED = True

End Sub

Private Sub txtChannel_Change()

   cmdApply.ENABLED = True

End Sub

Private Sub txtCycles_Change()

   cmdApply.ENABLED = True

End Sub

Private Sub txtNumChans_Change()

   cmdApply.ENABLED = True

End Sub

Private Sub txtOffset_Change()

   cmdApply.ENABLED = True

End Sub

Private Sub txtSamples_Change()

   cmdApply.ENABLED = True
   
End Sub

Private Sub txtStart_Change()

   cmdApply.ENABLED = True

End Sub
