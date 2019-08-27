Attribute VB_Name = "ScripEval"
Const BELOWRANGE = 1
Const ABOVERANGE = 2

Dim mvLongChanData As Variant, mvLongOutData As Variant
Dim mnResolution As Integer, mnRange As Integer
Dim mnNumChans As Integer, mnChanIndex As Integer, mnFirstChan As Integer
Dim mnDataSet As Integer, msChansMeasured As String
'Dim mlRateRequested As Long, mlRateReturned As Long
Dim mfRateRequested As Single, mfRateReturned As Single
Dim mlAmplGB As Long, mlRateGB As Long
Dim mfBoardAmplGB As Single, mlBoardRateGB As Long
Dim mlMvgAvgGB As Long, mlBoardMvgAvgGB As Long
Dim mnSimOut As Integer, mnSimIn As Integer
Dim mnRateWarning As Integer
Dim mlQueueCount As Long, mlLastElement As Long
Dim mlQueueEnabled As Integer
Dim mnGainQueue() As Integer
Dim msBoardName As String
Dim mlTotalSamples As Long

Dim mnGPIBError As Integer, mnAnalogInError As Integer
Dim mnAnalogOutError As Integer, mnDigitalInError As Integer
Dim mnDigitalOutError As Integer, mnCounterError As Integer
Dim mnUtilError As Integer, mnDigitalIOError As Integer
Dim mnConfigError As Integer, mnAnalogIOError As Integer
Dim mDbl488 As Double, mvCustomRange As Variant
Dim mlEventType As Long, mlEventData As Long, mnTimeout As Integer
Dim msOptions As String

Public Function RunEval(Args As Variant, Description As String) As Integer

   Dim FormRef As Form
   msOptions = ""
   NumArgs% = UBound(Args)
   If NumArgs% > 1 Then
      EvalType% = Val(Args(1))
      FormID$ = Args(0)
      If FormID$ = "0" Then
         FormFound% = True
         MainForm% = True
      Else
         FormFound% = GetFormReference(FormID$, FormRef)
      End If
   End If
   If Not FormFound% Then Exit Function
   
   If Not MainForm% Then
      'get board specific tweaks
      FormRef.cmdConfigure.Caption = "$"
      FormRef.cmdConfigure = True
      BoardName$ = FormRef.cmdConfigure.Caption
      If Not BoardName$ = msBoardName Then
         msBoardName = BoardName$
         Dummy$ = GetBoardTweaks()
      End If
      'get board options
      FormRef.cmdConfigure.Caption = "8"
      FormRef.cmdConfigure = True
      CurOptions$ = FormRef.cmdConfigure.Caption
      If Not CurOptions$ = "" Then CurOptions$ = "  " & CurOptions$
   End If
   msOptions = CurOptions$
   
   Select Case EvalType%
      Case EStatus
         Condition% = Val(Args(2))
         Limit1& = Val(Args(3))
         Limit2& = Val(Args(4))
         FailTimeout& = Val(Args(5))
         result% = EvalStatus(FormRef, Condition%, Limit1&, Limit2&, FailTimeout&, Description)
      Case EEventType
         'limits (mlEventType, mlEventData, and mnTimeout)
         'are presumed set by the timer in frmScript
         EventType& = Val(Args(2))
         EventData& = Val(Args(3))
         FailIfData% = Val(Args(4))
         NoEvent% = Val(Args(5))
         EvalTimeout% = Val(Args(6))
         result% = EvalEvent(FormRef, EventType&, EventData&, FailIfData%, _
         NoEvent%, EvalTimeout%, Description)
      Case EDataDC To EData32Delta
         'get number of channels
         Chans% = GetNumberOfChannels(FormRef)
         ChanString$ = " channel"
         If mnNumChans > 1 Then ChanString$ = " channels"
         msChansMeasured = " measuring " & Format(mnNumChans, "0") & ChanString$
         
         mnChanIndex = Val(Args(3))
         If (mnChanIndex < mlQueueCount) And mlQueueEnabled Then
            ChanExists% = True
         Else
            ChanExists% = (mnChanIndex < mnNumChans)
         End If
         If Not ChanExists% Then
            Description = "Script requested evaluation of a channel that isn't included in the data." & _
            vbCrLf & "Requesting channel " & Format(mnChanIndex, "0") & " evaluation." & vbCrLf
            RunEval = True
            Exit Function
         End If
         
         'get total count returned
         FormRef.cmdConfigure.Caption = "="
         FormRef.cmdConfigure = True
         DoEvents
         SamplesAcquired& = Val(FormRef.cmdConfigure.Caption)
         If SamplesAcquired& < 0 Then
            'one sample per read, current samples is stored in last element
            mlLastElement = Abs(SamplesAcquired&)
            SamplesAcquired& = mlLastElement
         End If
         Samples& = Val(Args(2))
         If Not EvalType% = EDataOutVsIn Then
            If Samples& > SamplesAcquired& Then
               'If EvalType% = SEvalTrigPoint Then
               CompSamps& = Samples& - (Samples& Mod mnNumChans)
               If CompSamps& > SamplesAcquired& Then TooManySamples% = True
            End If
            If TooManySamples% Then
               Description = "Script requested evaluation of a greater number of samples than " & _
               "the number of samples acquired." & vbCrLf & "Acquired " & Format(SamplesAcquired&, "0") & _
               " samples, evaluating " & Format(Samples&, "0") & " samples." & vbCrLf
               RunEval = True
               Exit Function
            End If
         End If
         If SamplesAcquired& < Samples& Then
            'this compensates for pretrigger returns less than full buffer
            EvalNumber& = SamplesAcquired&
         Else
            EvalNumber& = Samples&
         End If
         result% = EvalScriptData(FormRef, EvalType%, EvalNumber&, Args, Description)
      Case EError
         ULFunction% = Val(Args(3))
         ExpectedError& = Val(Args(4))
         Alternate1& = Val(Args(5))
         Alternate2& = Val(Args(6))
         Alternate3& = Val(Args(7))
         Action& = Val(Args(8))
         result% = EvalError(FormRef, ULFunction%, ExpectedError&, Alternate1&, Alternate2&, Alternate3&, Action&, Description)
      Case EHistogram   '10051
         RateVerified% = GetRateParams(FormRef, Warning$)
         'get channel being evaluated
         LowChannel% = GetLowChannel(FormRef)
         'get resolution of data
         mnResolution = GetResolution(FormRef)
         Range% = GetCurrentRange(FormRef)
         BinSpread& = Val(Args(2))
         MaxRMS! = Val(Args(3))
         BinAvg$ = Trim(Args(4))
         If Not InStr(1, BinAvg$, ".") = 0 Then
            AverageVal = CSng(BinAvg$)
         Else
            AverageVal = CLng(BinAvg$)
         End If
         AvgValTol& = Val(Args(5))
         result% = EvalHistogram(BinSpread&, MaxRMS!, AverageVal, _
         AvgValTol&, Description)
      Case EStore488Value '10052
         Readback$ = FormRef.Get488ValueRead()
         mDbl488 = Val(Readback$)
         NoDescription% = True
      Case ECompareStoredValue '10053
         ExpectedValue# = Val(Args(2))
         ErrorUnits% = Val(Args(3))
         Tolerance = Val(Args(4))
         If ErrorUnits% = 1 Then
            'Error defined in LSBs
            'requires form reference with form set to
            'appropriate board and range
            FormRef.cmdConfigure.Caption = "2"
            FormRef.cmdConfigure = True
            mnResolution = Val(FormRef.cmdConfigure.Caption)
            Range% = GetCurrentRange(FormRef)
         End If
         result% = EvalStoredValue(ExpectedValue#, ErrorUnits%, Tolerance, _
         Description)
   End Select
   If NoDescription% Then
      Description = ""
   Else
      Description = msBoardName & CurOptions$ & vbCrLf & Description & Warning$
   End If
   RunEval = result%
   
End Function

Function EvalScriptData(ByVal FormRef As Form, ByVal Characteristic As Integer, _
ByVal Samples As Long, ByVal Args As Variant, result As String) As Integer

   Dim FullDataSet As Variant
   Dim TypeOfData As VbVarType
   
   ConvertToCounts% = True   'by default, float data gets changed to long data type
   'get handle to data
   DataHandle& = FormRef.GetDataHandle(ACQUIREDDATA, TypeOfData, BufferSize&)
   'TypeOfData may contain a 'don't convert doubles to long' indicator
   'in the second nibble - indicating counts are retrieved as double type
   If (TypeOfData And &H10) = &H10 Then ConvertToCounts% = False
   TypeOfData = (TypeOfData And &HF)
   If DataHandle& = 0 Then
      result = "Could not get a handle to the data. Data evaluation aborted." & vbCrLf
      EvalScriptData = True
      Exit Function
   End If
   'get resolution of data
   FormRef.cmdConfigure.Caption = "2"
   FormRef.cmdConfigure = True
   mnResolution = Val(FormRef.cmdConfigure.Caption)
   
   'get the rate requested and the rate returned
   RateVerified% = GetRateParams(FormRef, Warning$)
   'get channel being evaluated
   LowChannel% = GetLowChannel(FormRef)
   'get current range
   Range% = GetCurrentRange(FormRef)
   
   If Samples < 0 Then
      NumSamples& = Abs(Samples) + 1
      mnDataSet = True
   Else
      NumSamples& = Samples
      mnDataSet = False
   End If
   mlTotalSamples = NumSamples&
   
   NumChannels% = mnNumChans
   If Characteristic = EDataOutVsIn Then
      NumberOfBits% = Val(Args(6))
      'SampsPerChan& = BufferSize& \ NumChannels%
      If NumberOfBits% > 0 Then
         NumSamples& = NumberOfBits% * NumChannels%
      Else
         NumSamples& = NumChannels%
      End If
      'If NumberOfBits% > 1 Then NumChannels% = 1
   Else
      SampsPerChan& = NumSamples& \ NumChannels%
      NumSamples& = SampsPerChan& * NumChannels%
   End If

   If Not ((NumDimens& < 0) Or (SizePerChan& < 0)) Then
      ConvResult% = GetBytesFromWinBuf(DataHandle&, TypeOfData, _
      NumSamples&, NumChannels%, FullDataSet)
      If Not ConvResult% Then
         result = "There's a problem getting the data requested. " & _
         "It's possible that too many data points are being requested. " & _
         "Data evaluation aborted." & vbCrLf
         EvalScriptData = True
         Exit Function
      End If
      
      TimerState% = frmScript.tmrScript.ENABLED
      frmScript.tmrScript.ENABLED = False
      
      Select Case Characteristic
         Case EDataDC
            ATolVal$ = Args(5)
            AmpValString$ = Args(4)
            pATolVal$ = ParseUnits(ATolVal$, ATolUnitType%)
            ATol = ConvertStringToType(pATolVal$, ATolUnitType%)
            pAmpVal$ = ParseUnits(AmpValString$, AmpUnitType%)
            DCVal = ConvertStringToType(pAmpVal$, AmpUnitType%)
            UnitType% = (AmpUnitType% * &H10) + (ATolUnitType%)
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            DCOption = Args(6)
            EvalResult% = EvalDC(DCVal, ATol, DCOption, UnitType%, Description$)
         Case EDataPulse
            ATolVal$ = Args(6)
            HiValString$ = Args(4)
            LoValString$ = Args(5)
            pATolVal$ = ParseUnits(ATolVal$, ATolUnitType%)
            ATol = ConvertStringToType(pATolVal$, ATolUnitType%)
            pHiVal$ = ParseUnits(HiValString$, HiValUnitType%)
            HiVal = ConvertStringToType(pHiVal$, HiValUnitType%)
            pLoVal$ = ParseUnits(LoValString$, LoValUnitType%)
            LoVal = ConvertStringToType(pLoVal$, LoValUnitType%)
            'LSBs& = Val(Args(6))
            'ATol = ConvertLSBs(LSBs&)
            'HiVal! = Val(Args(4)): LoVal! = Val(Args(5)) ': ATol! = Val(Args(6))
            HiBy& = Val(Args(7)): LoBy& = Val(Args(8)): TTol& = Val(Args(9))
            Repeat = Trim(Args(10))
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            EvalResult% = EvalPulse(HiVal, LoVal, ATol, HiBy&, LoBy&, TTol&, Repeat, Description$)
         Case EDataDelta
            DeltaArg$ = Args(4)
            UnitsLoc& = InStr(DeltaArg$, "°")
            VoltsLoc& = InStr(DeltaArg$, "V")
            If Not (UnitsLoc& = 0) Then
               ConvertToCounts% = False
               UnitType% = 2
               DeltaMin! = Val(Left(DeltaArg$, UnitsLoc& - 1))
            ElseIf Not (VoltsLoc& = 0) Then
               ConvertToCounts% = False
               UnitType% = 1
               DeltaMin! = Val(Left(DeltaArg$, VoltsLoc& - 1))
            Else
               DeltaMin! = Val(Args(4))
            End If
            DeltaType& = Val(Args(5))
            EvalOption = Trim(Args(6))
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            EvalResult% = EvalDelta(DeltaMin!, DeltaType&, EvalOption, UnitType%, Description$)
         Case EDataAmplitude
            ATude$ = Trim(Args(4))
            ATolVal$ = Args(5)
            pATolVal$ = ParseUnits(ATolVal$, ATolUnitType%)
            ATolerance = ConvertStringToType(pATolVal$, ATolUnitType%)
            pATude$ = ParseUnits(ATude$, ATudeUnitType%)
            Amplitude = ConvertStringToType(pATude$, ATudeUnitType%)
            UnitType% = (ATudeUnitType% * &H10) + (ATolUnitType%)
            EvalOption = Trim(Args(6))
            TypeOfPlot = FormRef.GetFormProperty("plottype")
            If TypeOfPlot = 6 Then
               CalcDerivative FullDataSet
               ConvertToCounts% = False
            End If
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            If Not DataRetrieved% Then Exit Function
            mvLongChanData = LongChanData
            EvalResult% = EvalAmplitude(Amplitude, ATolerance, EvalOption, UnitType%, Description$)
         Case SEvalTrigPoint  '10015
            'get pretrig count returned
            FormRef.cmdConfigure.Caption = "7"
            FormRef.cmdConfigure = True
            DoEvents
            ATol = Trim(Args(6))
            PreTrigPointsReturned& = Val(FormRef.cmdConfigure.Caption)
            'If mnSimIn Then
            '   TLimit& = PreTrigPointsReturned&
            'Else
               TLimit& = PreTrigPointsReturned& \ mnNumChans
            'End If
            TrigPolarity% = Val(Args(4)): Threshold! = Val(Args(5))
            TTol& = Val(Args(7))
            EvalOption = Trim(Args(8))
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            EvalResult% = EvalTrigPoint(TrigPolarity%, Threshold!, ATol, _
            TLimit&, TTol&, EvalOption, Description$)
         Case EDataTime
            ThreshArg$ = Args(4): GuardArg$ = Args(5)
            ThreshVal$ = ParseUnits(ThreshArg$, ThreshUnitType%)
            Threshold = ConvertStringToType(ThreshVal$, ThreshUnitType%)
            GuardVal$ = ParseUnits(GuardArg$, GuardUnitType%)
            Guardband = ConvertStringToType(GuardVal$, GuardUnitType%)
            DataType% = (ThreshUnitType% * &H10) + (GuardUnitType%)
            GuardCounts = Guardband
            If GuardUnitType% = UNITFLOAT Then
               dATude# = Guardband
               'TH! = Guardband
               GuardCounts = VoltsToHiResCounts(mnResolution, mnRange, dATude#, mvCustomRange)
               'GuardCounts = GetCounts(mnResolution, mnRange, TH!)
            End If
            TimeTol$ = Args(8)
            If IsNumeric(TimeTol$) Then
               TolVal! = Val(TimeTol$)
               Multiplier% = (Abs(TolVal!) < 1)
            End If
            PercLoc& = InStr(1, TimeTol$, "%")
            DecLoc& = InStr(1, TimeTol$, ".")
            If (PercLoc& > 1) Then
               TTol& = Val(Left(TimeTol$, PercLoc& - 1)) * 1000000
            ElseIf (DecLoc& > 0) And Multiplier% Then
               TTol& = TolVal! * 100000000
            Else
               TTol& = TolVal!
            End If
            SourceFreq! = Val(Args(6)): TimeLimit = Val(Args(7))
            EvaluationOption$ = Trim(Args(9))
            'EvalOption$ = ""
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            EvalResult% = EvalTime(Threshold, GuardCounts, DataType%, SourceFreq!, _
            TimeLimit, TTol&, EvaluationOption$, Description$)
         Case EDataOutVsIn
            FormID$ = Trim(Args(5))
            NumberOfBits% = Val(Args(6)) '/ mnNumChans
            CompareType& = Val(Args(4))
            NumSamplesPer& = NumSamples& / mnNumChans
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, NumSamplesPer&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            NumPorts& = Samples * 1
            EvalResult% = EvalOutVsIn(FormID$, NumberOfBits%, NumSamples&, CompareType&, Description$)
         Case EData32Delta
            Delta32Min& = Val(Args(4)): DeltaType& = Val(Args(5))
            EvalOption = Trim(Args(6))
            DataRetrieved% = GetLongChanData(FullDataSet, mnChanIndex, SampsPerChan&, _
            FirstPoint&, LongChanData, ConvertToCounts%)
            mvLongChanData = LongChanData
            EvalResult% = Eval32Delta(Delta32Min&, DeltaType&, EvalOption, Description$)
      End Select
   Else
      EvalResult% = True
      If (NumDimens& < 0) Then
         BadNumbers$ = "Invalid number of channels (" & Format(mnNumChans, "0") & ") for array."
      End If
      If (SizePerChan& < 0) Then
         BadNumbers$ = "Invalid number of samples (" & Format(Samples, "0") & ") for array."
      End If
      Description$ = BadNumbers$
   End If
   mvLongChanData = Null
   frmScript.tmrScript.ENABLED = TimerState%
   EvalScriptData = EvalResult% Or WarnRate%
   result = Description$ & Warning$ & vbCrLf

End Function

Function EvalStoredValue(ByVal ExpectedValue#, ByVal ErrorUnits%, _
ByVal Tolerance As Variant, Message As String) As Integer

   ErrorValue# = mDbl488 - ExpectedValue#

   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   If ErrorUnits% = 1 Then
      'need to convert values into LSBs for comparison
      ErrorVar = VoltsToHiResCounts(mnResolution, RangeVal%, ErrorValue#)
      Units$ = " LSBs"
   Else
      ErrorVar = ErrorValue#
      Units$ = " Volts"
   End If
   If Abs(ErrorVar) > Tolerance Then
      Failure$ = "GPIB value check failure." & vbCrLf & _
      "Value read = " & Format(mDbl488, "0.00######") & vbCrLf & _
      "Scripted limits = " & Format(ExpectedValue#, "0.00######") & " ±" & _
      Format(Tolerance, "0.######") & Units$ & "." & vbCrLf & _
      "Error = " & Format(ErrorVar, "0.00####") & Units$ & "." & vbCrLf
      EvalStoredValue = True
   Else
      Failure$ = "GPIB confirmation value = " & Format(ErrorVar, "0.0####") & Units$ & "."
   End If
   Message = Failure$
   
End Function

Function EvalHistogram(ByVal BinSpread&, ByVal MaxRMS!, ByVal AverageVal As Variant, _
ByVal AvgValTol As Long, Message As String) As Integer
   
   Dim DataType As VbVarType
   
   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   If mnResolution = 0 Then mnResolution = 16
   DataType = VarType(AverageVal)
   If DataType = vbSingle Then
      SingleVal! = AverageVal
      AvgDCCounts& = GetCounts(mnResolution, RangeVal%, SingleVal!)
   Else
      AvgDCCounts& = AverageVal
   End If

   If Not mfBoardAmplGB = 0 Then LSBs& = mfBoardAmplGB
   'BoardGB! = ConvertLSBs(LSBs&)
   'AvgTolCounts& = AvgValTol! + mlAmplGB + BoardGB!
   AvgValMin& = AvgDCCounts& - AvgValTol: AvgValMax& = AvgDCCounts& + AvgValTol
   CurRange$ = " using range " & GetRangeString(RangeVal%)
   
   NumberOfBins& = GetHistogram(Average, RMSVal, NumSamples&)
   
   If NumberOfBins& > BinSpread& Then
      Failure$ = "Noise level failure evaluating " & Format(NumSamples&, "0") & _
      " samples on channel " & Format(mnFirstChan, "0") & "." & vbCrLf & _
      "Acquisition rate = " & Format(mfRateReturned, "0.0###") & _
      "S/s." & vbCrLf & "Bin spread = " & Format(NumberOfBins&, "0") & " bins " & _
      CurRange$ & "." & vbCrLf & _
      "Scripted limits = " & Format(BinSpread&, "0") & " bins maximum." & _
      vbCrLf & "Error = " & Format(NumberOfBins& - BinSpread&, "0") & " bins." & vbCrLf
      FailHist% = True
   End If

   If Not FailHist% Then
      AvgValError! = Average - AvgDCCounts&
      If Abs(AvgValError!) > AvgValTol Then
         Failure$ = "Average histogram value failure evaluating " & Format(NumSamples&, "0") & _
         " samples on channel " & Format(mnFirstChan, "0") & "." & vbCrLf & _
         "Acquisition rate = " & Format(mfRateReturned, "0.0###") & _
         "S/s." & vbCrLf & "Average value = " & Format(Average, "0.00") & " counts" & _
         CurRange$ & "." & vbCrLf & _
         "Scripted limit = " & Format(AvgDCCounts&, "0") & " ±" & Format(AvgValTol, "0") & " max " & _
         " LSBs (" & Format(AvgValMin&, "0") & " to " & Format(AvgValMax&, "0") & " counts)." & _
         vbCrLf & "Error = " & Format(AvgValError!, "0.00") & " LSBs." & vbCrLf
         FailHist% = True
      End If
   End If
   
   If Not FailHist% Then
      If RMSVal > MaxRMS! Then
         Failure$ = "RMS noise value failure evaluating " & Format(NumSamples&, "0") & _
         " samples on channel " & Format(mnFirstChan, "0") & "." & vbCrLf & _
         "Acquisition rate = " & Format(mfRateReturned, "0.0###") & _
         "S/s." & vbCrLf & "RMS value = " & Format(RMSVal, "0.00") & " RMS LSBs " & _
         CurRange$ & "." & vbCrLf & _
         "Scripted limits = " & Format(MaxRMS!, "0.00") & " RMS LSBs max " & _
         vbCrLf & "Error = " & Format(MaxRMS! - RMSVal, "0.00") & " RMS LSBs." & vbCrLf
         FailHist% = True
      End If
   End If
   
   Message = Failure$
   If FailHist% Then
      EvalHistogram = True
   Else
      AvgValError! = Average - AvgDCCounts&
      Message = "Histogram values verified" & Format(mfRateReturned, "0") & CurRange$ & vbCrLf & _
      "evaluating " & Format(NumSamples&, "0") & " samples on channel " & _
      Format(mnFirstChan, "0") & " @ " & Format(mfRateReturned, "0.0###") & "S/s." & _
      vbCrLf & "Bin spread = " & Format(NumberOfBins&, "0") & " bins (" & _
      "scripted limit = " & Format(BinSpread&, "0") & " bins max)." & vbCrLf & _
      "RMS noise = " & Format(RMSVal, "0.00") & " RMS LSBs (" & _
      "scripted limit = " & Format(MaxRMS!, "0.00") & " RMS LSBs max)." & vbCrLf & _
      "Average error = " & Format(AvgValError!, "0.00") & " counts (" & _
      "Scripted limit = " & Format(AvgDCCounts&, "0") & " ±" & Format(AvgValTol, "0") & _
      " counts max.)" & vbCrLf
   End If
   
End Function

Function EvalError(ByVal FormRef As Form, ByVal ULFunction As Integer, _
ByVal ExpectedError As Long, ByVal Alternate1 As Long, ByVal Alternate2 As Long, _
ByVal Alternate3 As Long, ByVal Action As Long, result As String) As Integer

   Dim ExpectedErrors(3) As Long 'weekend
   Dim Alt(3) As String
   ExpectedErrors(0) = ExpectedError: ExpectedErrors(1) = Alternate1
   ExpectedErrors(2) = Alternate2: ExpectedErrors(3) = Alternate3
   'FuncName$ = GetFunctionName(ULFunction)
   FuncName$ = GetFunctionString(ULFunction)
   Details$ = GetErrorParams(FormRef, ExpectedError)
   ErrorFound& = FindFunctionError(ULFunction, Arguments$)
   ErrorFoundConst$ = GetErrorConst(ErrorFound&)
   If (Action And 8) = 8 Then
      'include "no error" in evaluation
      CheckNoError% = True
   End If
   If ErrorFound& = -1 Then
      Reason$ = " the script references a function (" & vbCrLf & _
      FuncName$ & ") that was not found in the most recent functions called"
      Fail% = True
   End If
   If Not Fail% Then
      For ErrOpt% = 0 To 3
         If (ErrorFound& = 0) And (Not CheckNoError%) Then
            'no error doesn't count as a match
            Reason$ = " no error was detected"
            Fail% = True
            Exit For
         End If
         If ErrorFound& = ExpectedErrors(ErrOpt%) Then
            MatchFound% = True
            ErrText$ = GetErrorText(ExpectedErrors(ErrOpt%))
            ErrorExpConst$ = GetErrorConst(ExpectedErrors(ErrOpt%))
            Exit For
         Else
            Alt(ErrOpt%) = " or " & Format(ExpectedErrors(ErrOpt%), "0")
         End If
      Next
      Fail% = Not MatchFound%
      If Fail% Then Reason$ = " error code " & Format(ErrorFound&, "0") & _
      " (" & ErrorFoundConst$ & ") was detected"
   End If
   
   If Fail% Then
      Descrip$ = "Expected error code " & Format(ExpectedError, "0") & _
      " (" & ErrorExpConst$ & ")" & Alt(0) & Alt(1) & Alt(2) & " but" & Reason$ & "." & vbCrLf & _
      "Function called = " & FuncName$ & "." & vbCrLf & _
      "Details: " & Arguments$ & "." & vbCrLf
      EvalError = True
   Else
      Descrip$ = "Verified error code " & Format(ExpectedError, "0") & _
      " (" & ErrorExpConst$ & " - " & ErrText$ & ") returned from " & _
      FuncName$ & "." & vbCrLf & "Details: " & Arguments$ & "."
      If Not Details$ = "" Then Descrip$ = Descrip$ & vbCrLf & Details$ & "."
   End If
   result = Descrip$
   
End Function

Function EvalStatus(ByVal FormRef As Form, ByVal Condition As Integer, ByVal Limit1 As Long, _
ByVal Limit2 As Long, ByVal FailTimeout As Long, result As String) As Integer

   StatError& = ReadStatus(FormRef, StatVal%, IndexVal&, CountVal&)
   If mnTimeout And Not (FailTimeout = 0) Then
      Fail1% = True
      If Condition = IDLE Then
         Suffix$ = "to go IDLE."
         TimeoutRunning% = True
      Else
         Suffix$ = "to return RUNNING."
      End If
      Descrip$ = "Timeout occurred while waiting for status " & Suffix$
   End If
   If (Not TimeoutRunning%) And (Not Fail1%) Then
      Select Case Condition
         Case 0   'idle
            If (StatVal% = RUNNING) Then
               Fail1% = True
               Descrip$ = "Expected device to be IDLE but detected RUNNING."
               NotIdle% = True
            Else
               Descrip$ = "Verified device to be IDLE with CurCount = " & _
               Format(CountVal&, "0") & "."
            End If
            LowCount% = (CountVal& < Limit1) 'Fail2%
            HighCount% = (CountVal& > Limit2)   'Fail3%
            If LowCount% Or HighCount% Then
               Descrip$ = "Expected IDLE with count between " & Format(Limit1, "0") & " and " & _
               Format(Limit2, "0") & " but detected " & Format(CountVal&, "0") & "."
               Fail1% = True
            End If
         Case 1, 2   'active, running unless count is totalized (used in retrigger, etc)
            If (StatVal% = IDLE) Then
               If (CountVal& > Limit2) Then
                  Fail1% = True
                  'to do - does this make sense?
                  PreDescrip$ = "Device is IDLE as expected but expected "
                  PDHighCount% = True
               Else
                  If Condition = 1 Then
                     Fail1% = True
                     Descrip$ = "Expected device to be RUNNING but detected IDLE."
                     NotRunning% = True
                  Else
                     If (CountVal& < Limit2) Then
                        Fail1% = True
                        Descrip$ = "Expected device to be RUNNING or completed " & _
                        "with a count of " & Format(Limit2, "0") & "," & vbCrLf & "but detected IDLE with count of " & Format(CountVal&, "0") & "."
                        NotRunning% = True
                     End If
                  End If
               End If
            Else
               PreDescrip$ = "Verified device to be RUNNING with "
            End If
            If Limit1 = -100 Then
               'check after Limit2 milliseconds that status doesn't change
               t0! = Timer()
               Lag! = Limit2 / 1000
               Do
                  t1! = Timer()
                  Diff! = t1! - t0!
                  DoEvents
               Loop While Diff! < Lag!
               CompareVal& = CountVal&
               StatError& = ReadStatus(FormRef, StatVal%, IndexVal&, CountVal&)
               If Not CompareVal& = CountVal& Then
                  Fail1% = True
                  Descrip$ = "Expected no change in CurCount& but detected change " & _
                  "from " & Format(CompareVal&, "0") & " to " & Format(CountVal&, "0") & "."
                  CountChange% = True  'Fail2% = True
               Else
                  Descrip$ = PreDescrip$ & "with no change in CurCount&."
               End If
            Else
               LowCount% = (CountVal& < Limit1) 'Fail2%
               HighCount% = (CountVal& > Limit2)   'Fail3%
               If LowCount% Or HighCount% Then
                  Fail1% = True
                  Descrip$ = "Expected RUNNING with count between " & Format(Limit1, "0") & " and " & _
                  Format(Limit2, "0") & " but detected " & Format(CountVal&, "0") & "."
               Else
                  If Not NotRunning% Then
                     Descrip$ = PreDescrip$ & "count between " & _
                     Format(Limit1, "0") & " and " & Format(Limit2, "0") & _
                     " (CurCount = " & Format(CountVal&, "0") & ")."
                  End If
               End If
            End If
      End Select
   End If
   Failure% = Fail1%
   'If Fail1% Then
      'possibly a script error - checking for running but scan completed
   '   Descrip$ = PreDescrip$ & vbCrLf & "Possible script error - checking for running but scan completed."
   '   result = Descrip$ & vbCrLf
   'End If
   If Failure% Then
      result = "***  GetStatus failure   ****" & vbCrLf & Descrip$ & vbCrLf
   Else
      result = Descrip$ & vbCrLf
   End If
   EvalStatus = Failure%
   
End Function

Function EvalEvent(ByVal FormRef As Form, ByVal EventType As Long, ByVal EventData As Long, _
ByVal FailIfData As Integer, ByVal NoEvent As Integer, ByVal EvalTimeout As Integer, _
result As String) As Integer

   If mnTimeout Or (mlEventType = 0) Then FormRef.GetEvent mlEventType, mlEventData, EventParam&
   
   ExpEvent% = (NoEvent = 0)
   If Not ExpEvent% Then Prefix$ = "no "
   If mlEventType = 0 Then
      TypeString$ = "no"
   Else
      TypeString$ = GetEventTypeString(mlEventType)
   End If
   DataString$ = Format(mlEventData, "0")
   ExpTypeString$ = Prefix$ & GetEventTypeString(EventType)
   ExpDataString$ = Format(EventData, "0")
   FailTimeout% = Not (EvalTimeout = 0)
   If mnTimeout Then WaitTimo$ = " after timeout"
      
   If Not ExpEvent% Then
      If Not (mlEventType = 0) Then
         Descrip$ = "Expected " & ExpTypeString$ & " but " & TypeString$ & vbCrLf & _
         "event occurred" & WaitTimo$ & " (data = " & DataString$ & ")."
         Fail1% = True
      Else
         Descrip$ = "Verified " & ExpTypeString$ & _
         " event occurred, data value = " & _
         ExpDataString$ & WaitTimo$ & "."
      End If
   Else
      If mlEventType = 0 Then
         Descrip$ = "Expected " & ExpTypeString$ & " event but " & _
         TypeString$ & " event occurred" & WaitTimo$ & "."
         Fail1% = True
      End If
   End If
   If FailTimeout% Then
      If mnTimeout Then
         Descrip$ = "Expected " & ExpTypeString$ & " event but timeout occurred with " & _
         TypeString$ & " event."
         Fail1% = True
      End If
   End If
   If Not (Fail1% Or TestComplete%) Then
      If ExpEvent% Then
         If Not ((EventType And mlEventType) = mlEventType) Then
            Fail1% = True
            Descrip$ = "Expected " & ExpTypeString$ & " event but " & _
            vbCrLf & TypeString$ & WaitTimo$ & " event occurred."
         End If
      End If
   End If
   If ExpEvent% Then
      DataDiff& = mlEventData - EventData
      Select Case FailIfData
         Case 1   'less than
            DataWindow$ = ", fail if less than " & ExpDataString$
            If Not DataDiff& > 0 Then
               FailData% = True
            End If
         Case 2   'not equal
            DataWindow$ = ", fail if not equal to " & ExpDataString$
            If Not DataDiff& = 0 Then
               FailData% = True
            End If
         Case 3   'greater than
            DataWindow$ = ", fail if greater than " & ExpDataString$
            If Not DataDiff& < 0 Then
               FailData% = True
            End If
      End Select
      If Not (Fail1% Or TestComplete%) Then
         If FailData% Then
            Descrip$ = "Expected " & ExpTypeString$ & DataWindow$ & _
            vbCrLf & ExpDataString$ & " but " & TypeString$ & _
            " event occurred, data = " & DataString$ & WaitTimo$ & "."
         End If
      End If
   Else
      DataWindow$ = "."
   End If
   
   If Fail1% Then
      result = "***  EvalEvent failure   ****" & vbCrLf & Descrip$ & vbCrLf
   Else
      Descrip$ = "Verified " & TypeString$ & " event" & WaitTimo$ & _
      ", data value = " & DataString$ & "." & vbCrLf & _
      "Script limits = " & ExpTypeString$ & " event" & _
      DataWindow$
      result = Descrip$ & vbCrLf
   End If
   mnTimeout = False
   mlEventType = 0
   mlEventData = 0
   FormRef.SetEvent mlEventType, mlEventData
   EvalEvent = Fail1%
   
End Function

Function EvalPulse(ByVal HiVal As Variant, ByVal LoVal As Variant, _
ByVal ATol As Variant, ByVal HiBy As Long, ByVal LoBy As Long, _
ByVal TTol As Long, ByVal EvalOption As Variant, Message As String) As Integer

   FirstPoint& = 0
   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   ParseAll = Split(EvalOption, ";")
   NumOptions& = UBound(ParseAll)
   For CurOpt& = 0 To NumOptions&
      ThisOpt = Trim(ParseAll(CurOpt&))
      EvaluationType = Split(ThisOpt, " ")
      EvalParam& = UBound(EvaluationType)
      EOptionDesc$ = "."
      If EvalParam& = 0 Then
         Repeat% = False
      Else
         EvaluationOption$ = LCase(EvaluationType(0))
         'following allows for "String = Value"
         'or "String Value" construct
         EvaluationValue$ = EvaluationType(EvalParam&)
         If IsNumeric(EvaluationValue$) Then
            OptionValue& = EvaluationType(EvalParam&)
         Else
            OptionValue& = 0
            Repeat% = False
         End If
         Select Case EvaluationOption$
            Case "repeat"
               Repeat% = OptionValue&
            Case "first", "firstpoint", "start"
               FirstPoint& = OptionValue& \ mnNumChans
               EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
         End Select
      End If
   Next
   NumSamples& = UBound(mvLongChanData)
   HiVolts! = HiVal
   LoVolts! = LoVal
   HiCounts& = GetCounts(mnResolution, RangeVal%, HiVolts!)
   LoCounts& = GetCounts(mnResolution, RangeVal%, LoVolts!)
   ChanCount% = 1
   If TTol < 0 Then
      ChanCount% = mnNumChans
      TTol = Abs(TTol)
   End If
   'AmplGBCounts& = VoltsToCounts(mnResolution, RangeVal%, mfAmplGB)
   'AmplGBCountsBd& = VoltsToCounts(mnResolution, RangeVal%, mfBoardAmplGB)
   TolVolts! = ATol
   AmpTolCounts& = VoltsToCounts(mnResolution, RangeVal%, TolVolts!) '+ AmplGBCounts& + AmplGBCountsBd&
   ATolCounts& = AmpTolCounts& + AmplGBCounts& + AmplGBCountsBd&
   HiMin& = HiCounts& - ATolCounts&: HiMax& = HiCounts& + ATolCounts&
   LoMin& = LoCounts& - ATolCounts&: LoMax& = LoCounts& + ATolCounts&
   
   TransUp& = (HiBy \ ChanCount%) + FirstPoint&
   TransDown& = (LoBy \ ChanCount%) + FirstPoint&
   If Not (TransUp& > NumSamples&) Then CheckPosTrans% = True
   If Not (TransDown& > NumSamples&) Then CheckNegTrans% = True
   HiByMax& = TransUp& + TTol: HiByMin& = TransUp& - TTol
   LoByMax& = TransDown& + TTol: LoByMin& = TransDown& - TTol
   StartPoint& = FirstPoint&
   If TransUp& < TransDown& Then
      'checking for rising edge first
      limits$ = "Scripted limits = Transition to " & Format(HiCounts&, "0") & _
      ", ±" & Format(ATolCounts&, "0") & " LSBs by sample " & Format(TransUp& * mnNumChans, "0") & _
      ", ±" & Format(TTol * mnNumChans, "0") & " samples" & vbCrLf & "followed by transition to " & _
      Format(LoCounts&, "0") & ", ±" & Format(ATolCounts&, "0") & " LSBs by sample " & _
      Format(TransDown& * mnNumChans, "0") & ", ±" & Format(TTol * mnNumChans, "0") & " samples."
      TransMin& = HiMin&: TransMax& = HiMax&: FindLowPulse% = False
      GoSub FindTransition
      If (TransPoint& = StartPoint&) Then
         FailText$ = "No transition was detected on channel " & Format(Channel%, "0") & msChansMeasured & "."
         EvalPulseFail% = True
      Else
         ErrorValue& = TransPoint& - TransUp&
         AErrorValue& = CurValue& - HiCounts&
         If Abs(AErrorValue&) > ATolCounts& Then
            FailText$ = "Pulse amplitude measured at " & Format(CurValue&, "0") & " counts" & _
            " at sample " & Format(FirstPoint& + (DataPoint& * mnNumChans), "0") & "." & vbCrLf & _
            "Scripted limits = " & Format(LoCounts&, "0") & " low and " & Format(HiCounts&, "0") & _
            " high, ±" & Format(ATolCounts&, "0") & " counts."
            EvalPulseFail% = True
         Else
            If (TransPoint& > HiByMax&) Or (TransPoint& < HiByMin&) Then
               State$ = "high": TimeSpec$ = Format(TransUp& * mnNumChans, "0")
               TimeMin$ = Format(HiByMin& * mnNumChans, "0")
               TimeMax$ = Format(HiByMax& * mnNumChans, "0")
               GoSub BuildErrString
               EvalPulseFail% = True
            Else
               TError1$ = Format(ErrorValue& * mnNumChans, "0") & " samples, " & _
               Format(AErrorValue&, "0") & " LSBs, high, "
               FirstTrans$ = "Transition to " & Format(CurValue&, "0") & " occurred at " & _
               Format(FirstPoint& + (TransPoint& * mnNumChans), "0") & " samples"
            End If
         End If
      End If
      'if previous test passed, then check for falling edge
      If (Not EvalPulseFail%) And CheckNegTrans% Then
         TransMin& = LoMin&: TransMax& = LoMax&: FindLowPulse% = True
         GoSub FindTransition
         If TransPoint& = StartPoint& Then
            FailText$ = "No transition was detected on channel " & _
            Format(Channel%, "0") & msChansMeasured & "."
            EvalPulseFail% = True
         Else
            ErrorValue& = TransPoint& - TransDown&
            AErrorValue& = CurValue& - LoCounts&
            If Abs(AErrorValue&) > ATolCounts& Then
               FailText$ = "Pulse amplitude measured at " & Format(CurValue&, "0") & " counts" & _
               " at sample " & Format(FirstPoint& + (DataPoint& * mnNumChans), "0") & "." & vbCrLf & _
               "Scripted limits = " & Format(LoCounts&, "0") & " and " & Format(HiCounts&, "0") & _
               " counts, ±" & Format(ATolCounts&, "0") & "."
               EvalPulseFail% = True
            Else
               If (TransPoint& > LoByMax&) Or (TransPoint& < LoByMin&) Then
                  State$ = "low": TimeSpec$ = Format(TransDown& * mnNumChans, "0")
                  TimeMin$ = Format(LoByMin& * mnNumChans, "0")
                  TimeMax$ = Format(LoByMax& * mnNumChans, "0")
                  GoSub BuildErrString
                  EvalPulseFail% = True
               Else
                  TError2$ = Format(ErrorValue& * mnNumChans, "0") & " samples, " & _
                  Format(AErrorValue&, "0") & " LSBs, low."
                  SecondTrans$ = " and to " & Format(CurValue&, "0") & " at " & _
                  Format(FirstPoint& + (TransPoint& * mnNumChans), "0") & " samples."
               End If
            End If
         End If
      End If
      If Not EvalPulseFail% Then
         PrevTransPoint& = TransPoint&
         If Repeat% Then
            'at least one more transition to high expected
            RepeatLimit$ = "Initial transition should be followed by at least one additional pulse."
            TransMin& = HiMin&: TransMax& = HiMax&: FindLowPulse% = False
            GoSub FindTransition
            If Not (TransPoint& > PrevTransPoint&) Then
               FailText$ = "Pulse check failure on channel " & Format(Channel%, "0") & "." & vbCrLf & _
               "Script specifies repeating pulses, but only one pulse was detected in data."
               EvalPulseFail% = True
            Else
               ThirdTrans$ = "Confirmed additional transitions after initial pulse."
            End If
         ElseIf CheckNegTrans% Then
            'should be low for remainder of data points unless
            'only high transition specified (TransDown& > NumSamples)
            RepeatLimit$ = "Single pulse only specified."
            TransMin& = HiMin&: TransMax& = HiMax&: FindLowPulse% = False
            GoSub FindTransition
            If Not (TransPoint& = 0) Then
               FailText$ = "Pulse check failure on channel " & Format(Channel%, "0") & "." & vbCrLf & _
               "Second transition to high after " & Format(PrevTransPoint& * mnNumChans, "0") & _
               " samples." & vbCrLf & _
               "Script specifies a single pulse, but a second pulse was detected in data at " & _
               Format(FirstPoint& + (TransPoint& * mnNumChans), "0") & " samples."
               EvalPulseFail% = True
            Else
               ThirdTrans$ = "Confirmed steady state after initial pulse."
            End If
         End If
      End If
   Else
      'check for falling edge first
      limits$ = "Scripted limits = Transition to " & Format(LoCounts&, "0") & _
      ", ±" & Format(ATolCounts&, "0") & " LSBs by " & Format(TransDown& * mnNumChans, "0") & " samples" & _
      ", ±" & Format(TTol * mnNumChans, "0") & " samples followed by transition to " & _
      Format(HiCounts&, "0") & ", ±" & Format(ATolCounts&, "0") & " LSBs by sample " & _
      Format(TransUp& * mnNumChans, "0") & ", ±" & Format(TTol * mnNumChans, "0") & " samples."
      TransMin& = LoMin&: TransMax& = LoMax&: FindLowPulse% = True
      GoSub FindTransition
      If TransPoint& = StartPoint& Then
         FailText$ = "No transition was detected on channel " & Format(Channel%, "0") & msChansMeasured & "."
         EvalPulseFail% = True
      Else
         ErrorValue& = TransPoint& - TransDown&
         AErrorValue& = CurValue& - LoCounts&
         If Abs(AErrorValue&) > ATolCounts& Then
            FailText$ = "Pulse amplitude measured at " & Format(CurValue&, "0") & " counts" & _
            " at sample " & Format(FirstPoint& + (DataPoint& * mnNumChans), "0") & "." & vbCrLf & _
            "Scripted limits = " & Format(LoCounts&, "0") & " and " & Format(HiCounts&, "0") & _
            " counts, ±" & Format(ATolCounts&, "0") & "."
            EvalPulseFail% = True
         Else
            If (TransPoint& > LoByMax&) Or (TransPoint& < LoByMin&) Then
               State$ = "low": TimeSpec$ = Format(TransDown& * mnNumChans, "0")
               TimeMin$ = Format(LoByMin& * mnNumChans, "0")
               TimeMax$ = Format(LoByMax& * mnNumChans, "0")
               GoSub BuildErrString
               EvalPulseFail% = True
            Else
               TError1$ = Format(ErrorValue& * mnNumChans, "0") & " samples, " & _
               Format(AErrorValue&, "0") & " LSBs, low, "
               FirstTrans$ = "Transition to " & Format(CurValue&, "0") & _
               " occurred at " & Format(FirstPoint& + (TransPoint& * mnNumChans), "0") & " samples"
            End If
         End If
      End If
      'then check for rising edge
      If (Not EvalPulseFail%) And CheckPosTrans% Then
         TransMin& = HiMin&: TransMax& = HiMax&: FindLowPulse% = False
         GoSub FindTransition
         If TransPoint& = StartPoint& Then
            FailText$ = "No transition was detected on channel " & Format(Channel%, "0") & msChansMeasured & "."
            EvalPulseFail% = True
         Else
            ErrorValue& = TransPoint& - TransUp&
            If Abs(AErrorValue&) > ATolCounts& Then
               FailText$ = "Pulse amplitude measured at " & Format(CurValue&, "0") & " counts" & _
               " at sample " & Format(FirstPoint& + (DataPoint& * mnNumChans), "0") & "." & vbCrLf & _
               "Scripted limits = " & Format(LoCounts&, "0") & " and " & Format(HiCounts&, "0") & _
               " counts, ±" & Format(ATolCounts&, "0") & "."
               EvalPulseFail% = True
            Else
               If (TransPoint& > HiByMax&) Or (TransPoint& < HiByMin&) Then
                  State$ = "high": TimeSpec$ = Format(TransUp& * mnNumChans, "0")
                  TimeMin$ = Format(HiByMin& * mnNumChans, "0")
                  TimeMax$ = Format(HiByMax& * mnNumChans, "0")
                  GoSub BuildErrString
                  EvalPulseFail% = True
               Else
                  TError2$ = Format(ErrorValue& * mnNumChans, "0") & " samples, " & _
                  Format(AErrorValue&, "0") & " LSBs, high."
                  SecondTrans$ = " and to " & Format(CurValue&, "0") & " at " & _
                  Format(FirstPoint& + (TransPoint& * mnNumChans), "0") & " samples."
               End If
            End If
         End If
      End If
      PrevTransPoint& = TransPoint&
      If Not EvalPulseFail% Then
         If Repeat% Then
            'at least one more transition to low expected
            RepeatLimit$ = "Initial transition should be followed by at least one additional pulse."
            TransMin& = LoMin&: TransMax& = LoMax&: FindLowPulse% = True
            GoSub FindTransition
            If Not (TransPoint& > PrevTransPoint&) Then
               'FailText$ = "Expected transition to low after " & Format(PrevTransPoint&, "0") & _
               '". Second pulse did not occur or wasn't detected."
               FailText$ = "Pulse check failure on channel " & Format(Channel%, "0") & "." & vbCrLf & _
               "Script specifies repeating pulses, but only one pulse was detected in data."
               EvalPulseFail% = True
            Else
               ThirdTrans$ = "Confirmed additional transitions after initial pulse."
            End If
         ElseIf CheckPosTrans% Then
            'should be high for remainder of data points unless
            'only low transition specified (TransUp& > NumSamples)
            RepeatLimit$ = "Single pulse only specified."
            TransMin& = LoMin&: TransMax& = LoMax&: FindLowPulse% = True
            GoSub FindTransition
            If Not (TransPoint& = 0) Then
               'FailText$ = "Transition to low occurred at " & Format(TransPoint&, "0") & _
               '". Expected only one pulse according to the script."
               FailText$ = "Pulse check failure on channel " & Format(Channel%, "0") & "." & vbCrLf & _
               "Second transition to low after " & Format(PrevTransPoint& * mnNumChans, "0") & _
               " samples." & vbCrLf & _
               "Script specifies a single pulse, but a second pulse was detected in data at " & _
               Format(FirstPoint& + (TransPoint& * mnNumChans), "0") & " samples."
               EvalPulseFail% = True
            Else
               ThirdTrans$ = "Confirmed steady state after initial pulse."
            End If
         End If
      End If
   End If
   If Not EvalPulseFail% Then
      Message = "Pulse verified on channel " & Format(Channel%, "0") & msChansMeasured & "." & _
      vbCrLf & FirstTrans$ & SecondTrans$ & vbCrLf & ThirdTrans$ & vbCrLf & _
      limits$ & vbCrLf & RepeatLimit$ & vbCrLf & "Error = " & TError1$ & TError2$
   Else
      Message = FailText$
   End If
   EvalPulse = EvalPulseFail%
   
   Exit Function
   
FindTransition:
   TransPoint& = 0
   CurState% = 0
   For DataPoint& = StartPoint& To (NumSamples&)
      CurValue& = mvLongChanData(DataPoint&)
      If (CurValue& > TransMin&) And (CurValue& < TransMax&) Then
         TransPoint& = DataPoint&
'         If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
         StartPoint& = DataPoint& + 1
         Exit For
      End If
      If (CurValue& > HiMax&) Or (CurValue& < LoMin&) Then
         TransPoint& = DataPoint&
         '         If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
         StartPoint& = DataPoint& + 1
         Exit For
      End If
      'If (CurValue& > LoMax&) And (CurValue& < HiMin&) Then
      '   TransPoint& = DataPoint&
      '   StartPoint& = DataPoint& + 1
      '   Exit For
      'End If
      If CurValue& < TransMin& Then
         If CurState% = ABOVERANGE Then
            TransPoint& = DataPoint&
            '         If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
            StartPoint& = DataPoint& + 1
            Exit For
         Else
            If FindLowPulse% Then
               'if looking for falling edge and already below
               TransPoint& = DataPoint&
               '         If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
               StartPoint& = DataPoint& + 1
               Exit For
            End If
         End If
         CurState% = BELOWRANGE
      Else
         If CurState% = BELOWRANGE Then
            TransPoint& = DataPoint&
            '         If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
            StartPoint& = DataPoint& + 1
            Exit For
         Else
            If Not FindLowPulse% Then
               'if looking for rising edge and already above
               TransPoint& = DataPoint&
               '         If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
               StartPoint& = DataPoint& + 1
               Exit For
            End If
         End If
         CurState% = ABOVERANGE
      End If
      'CurState% = 0
'      If CurValue& > TransMax& Then
'         If CurState% = BELOWRANGE Then
'            Exit For
'         Else
'            If Not FindLowPulse% Then
'               'if looking for falling edge and already below
'               Exit For
'            End If
'         End If
'         CurState% = ABOVERANGE
'      Else
'         If CurState% = ABOVERANGE Then
'            Exit For
'         Else
'            If FindLowPulse% Then
'               'if looking for rising edge and already above
'               Exit For
'            End If
'         End If
'         CurState% = BELOWRANGE
'      End If
   Next DataPoint&
   'If mnSimIn = 0 Then TransPoint& = DataPoint& * mnNumChans
   'StartPoint& = DataPoint& + 1
   Return
   
BuildErrString:
   FailText$ = "Pulse check failure on channel " & Format(Channel%, "0") & "." & vbCrLf & _
   "Transition to " & State$ & " occurred at " & Format(FirstPoint& + (TransPoint& * _
   mnNumChans), "0") & " samples." & vbCrLf & _
   "Scripted limits = sample " & TimeSpec$ & ", ±" & Format(TTol * mnNumChans, "0") & _
   " samples (between samples " & TimeMin$ & " and " & TimeMax$ & ")  " & _
   vbCrLf & "Error = " & Format(ErrorValue& * mnNumChans, "0") & " samples."
   Return

End Function

Function EvalOutVsIn(ByVal FormID As String, ByVal NumBits As Integer, ByVal Samples As Long, _
ByVal CompareType As Long, Description As String) As Integer

   Dim FormRef As Form, TypeOfData As VbVarType
   
   FormFound% = GetFormReference(FormID$, FormRef)
   If Not FormFound% Then Exit Function
   
   'get current sample (stored in mlCount in FormRef)
   FormRef.cmdConfigure.Caption = "="
   FormRef.cmdConfigure = True
   CurSample& = Val(FormRef.cmdConfigure.Caption)
   NumRead& = UBound(mvLongChanData)
   
   DataHandle& = FormRef.GetDataHandle(GENERATEDDATA, TypeOfData, BufferSize&)
   If DataHandle& = 0 Then
      Description = "Could not get a handle to the generated" & _
      " data. Data evaluation aborted." & vbCrLf
      EvalOutVsIn = True
      Exit Function
   End If
   NumChans% = mnNumChans
   If NumBits = 0 Then
      NumChans% = 0
      CurSample& = (CurSample& - 1) * mnNumChans
   End If
   ConvResult% = GetBytesFromWinBuf(DataHandle&, TypeOfData, _
   Samples, mnNumChans, vDataArray, CurSample& - NumChans%)
   If Abs(NumBits > 1) Then
      BitEval% = True
      NumOfBits% = Abs(NumBits)
      'NumRead& = ((NumRead& + 1) * (NumOfBits%)) - 1
      ConvResult% = GetBitsFromArray(vDataArray, TypeOfData, _
      1, NumOfBits%, OutputData, mnChanIndex)
      'ConvResult% = GetBitsByPortFromArray(vDataArray, _
      TypeOfData, NumChans%, NumBits * mnNumChans, OutputData)
   Else
      'evaluate as port value
      SampsPerChan& = BufferSize& \ mnNumChans
      DataRetrieved% = GetLongChanData(vDataArray, mnChanIndex, 1, _
      FirstPoint&, LongChanData, True)
      OutputData = LongChanData
   End If
   If BitEval% Then
      NumOutSamples& = (NumBits * 1) - 1  'mnNumChans
      SampleIndex& = CurSample& - 1 'NumBits
   Else
      NumOutSamples& = UBound(OutputData) 'Samples - 1
      'SampleIndex& = CurSample& - 1
   End If
   If NumRead& < NumOutSamples& Then
      Description = "Not all output samples were read. There are " & _
      Format(NumOutSamples& + 1, "0") & " samples available, but only " & _
      Format(NumRead& + 1, "0") & " samples were read."
      EvalOutVsIn = True
      Exit Function
   End If
   If SampleIndex& < 0 Then
      Description = "Input vs ouput values cannot be evaluated. No output data index is available."
      EvalOutVsIn = True
      Exit Function
   End If
   
   Select Case CompareType
      Case 0
         'exactly equal
         BitMask& = &HFFFFFFFF
      Case Is > 0
         'number of bits to compare
         Bits% = Abs(CompareType)
         BitMask& = 2 ^ Bits% - 1
         Invert% = False
      Case Is < 0
         'invert value, number of bits to compare
         Bits% = Abs(CompareType)
         BitMask& = 2 ^ Bits% - 1
         Invert% = True
   End Select
   If BitEval% Then BitMask& = 1
   
   For DataPoint& = 0 To NumOutSamples&
      CompareValue& = OutputData(DataPoint&)   'SampleIndex& +
      If Invert% Then
         CompareValue& = (Not CompareValue&) And BitMask&
      Else
         CompareValue& = (CompareValue& And BitMask&)
      End If
      ReadMask& = (mvLongChanData(DataPoint&) And BitMask&)
      ErrorCounts& = ReadMask& - CompareValue&
      If Not (ErrorCounts& = 0) Then
         Message$ = "Output vs Input failure on channel " & Format(mnChanIndex, "0") & _
         msChansMeasured & ". " & vbCrLf & "Difference = " & Format(ErrorCounts&, "0") & _
         " counts at sample " & Format(DataPoint&, "0") & "." & vbCrLf
         FailCompareIO% = True
         EvalOutVsIn = True
         Exit For
      End If
   Next DataPoint&
   
   If Not FailCompareIO% Then
      Message$ = "Verified output vs input on channel " & Format(mnChanIndex, "0") & _
      msChansMeasured & ". " & vbCrLf
   End If
   Description = Message$ & "Comparing " & Format(NumOutSamples& + 1, "0") & _
   " output samples to samples read." & vbCrLf & _
   "Scripted limits = No difference between output and input."
   EvalOutVsIn = FailCompareIO%
   
End Function

Function EvalDC(ByVal DCVal As Variant, ByVal ATol As Variant, ByVal EvalOption As Variant, _
ByVal UnitType As Integer, Message As String) As Integer

   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   NumSamples& = UBound(mvLongChanData)
   If mnResolution = 0 Then mnResolution = 16
   TolDataType% = UnitType And &HF
   AmpDataType% = (UnitType And &HF0) / &H10
   
   Select Case AmpDataType%
      Case 0
         DCCounts& = DCVal
      Case UNITPERCENT  '1
      Case UNITVOLTS, UNITFLOAT '2
         SingleVal! = DCVal
         DCCounts& = GetCounts(mnResolution, RangeVal%, SingleVal!)
      Case UNITDEGREES '3
      Case UNITCOUNTS   '4
         DCCounts& = DCVal
   End Select

   If Not mfBoardAmplGB = 0 Then LSBs& = mfBoardAmplGB
   Select Case UnitType
      Case 0
         BoardGB! = ConvertLSBs(LSBs&)
      Case UNITPERCENT  '1
      Case UNITVOLTS '2
         SingleVal! = DCVal
         'DCCounts& = GetCounts(mnResolution, RangeVal%, SingleVal!)
      Case UNITDEGREES '3
      Case UNITCOUNTS   '4
         BoardGB! = LSBs&
   End Select
   ATolCounts& = ATol + mlAmplGB + BoardGB!
   DCValMin& = DCCounts& - ATolCounts&: DCValMax& = DCCounts& + ATolCounts&
   RangeString$ = GetRangeString(RangeVal%)
   If Not RangeString$ = "" Then
      CurRange$ = " using range " & GetRangeString(RangeVal%)
   Else
      CurRange$ = ""
   End If
   
   FirstPoint& = 0
   EOptionDesc$ = "."
   LimitString$ = Format(DCCounts&, "0") & " ±" & Format(ATolCounts&, "0")
   LimitSpread$ = " (" & Format(DCValMin&, "0") & " to " & Format(DCValMax&, "0") & " counts)."
   If VarType(EvalOption) = vbString Then
      EmptyString% = (EvalOption = "")
   End If
   If IsEmpty(EvalOption) Or EmptyString% Then
      EvalParam& = 0
      Span& = 0
      EvaluationOption$ = ""
   Else
      If IsNumeric(EvalOption) Then
         Span& = Val(EvalOption)
      Else
         OptString$ = Trim(EvalOption)
         If OptString$ = "" Then
            Span& = 0
         Else
            EvaluationType = Split(OptString$, " ")
            EvalParam& = UBound(EvaluationType)
            If EvalParam& = 0 Then
               Span& = Val(EvaluationType(0))
            Else
               EvaluationOption$ = LCase(EvaluationType(0))
               'following allows for "String = Value"
               'or "String Value" construct
               StringVal$ = EvaluationType(EvalParam&)
               If IsNumeric(StringVal$) Then OptionValue& = EvaluationType(EvalParam&)
               Select Case EvaluationOption$
                  Case "movingaverage", "moving"
                     Span& = OptionValue&
                     'EOptionDesc$ = ", moving average = " & Format(Span&, "0")
                  Case "first", "firstpoint", "start"
                     FirstPoint& = OptionValue&
                     EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
                  Case "greater", ">", "greaterthan"
                     LimitString$ = " > " & Format(DCCounts&, "0")
                     LimitSpread$ = "."
                     CompType% = 1
                  Case "less", "<", "lessthan"
                     LimitString$ = " < " & Format(DCCounts&, "0")
                     CompType% = 2
                     LimitSpread$ = "."
                  Case "not", "<>", "notequal"
                     LimitString$ = " <> " & Format(DCCounts&, "0")
                     LimitSpread$ = "."
                     CompType% = 3
               End Select
            End If
         End If
      End If
   End If
   If Span& > 0 Then
      'use moving average
      EOptionDesc$ = ", moving average = " & Format(Span&, "0") & "."
   End If
   
   SampleSet& = (NumSamples& - FirstPoint&) + 1
   For DataPoint& = FirstPoint& To NumSamples&
      If Span& > 0 Then
         'use moving average
         Span1& = DataPoint& + Span& - 1
         If (DataPoint& + Span&) > (NumSamples& - 2) Then Exit For
         CumVal1& = 0: CumVal2& = 0
         SpanCum& = 0
         For Element& = DataPoint& To Span1&
            CumVal1& = CumVal1& + (mvLongChanData(Element&) - DCCounts&)
            SpanCum& = SpanCum& + mvLongChanData(Element&)
         Next Element&
         ErrorCounts& = CumVal1& / Span&
         SpanCounts& = SpanCum& / Span&
      Else
         ErrorCounts& = mvLongChanData(DataPoint&) - DCCounts&
      End If
      Select Case CompType%
         Case 1
            'greater than
            If Not (ErrorCounts& > 0) Then
               MaxErr& = ErrorCounts&
               MaxSample& = DataPoint&
            Else
               ErrorCounts& = 0
            End If
         Case 2
            'less than
            If Not (ErrorCounts& = 0) Then
               MaxErr& = ErrorCounts&
               MaxSample& = DataPoint&
            End If
         Case Else
            If Abs(ErrorCounts&) > Abs(MaxErr&) Then
               MaxErr& = ErrorCounts&
               MaxSample& = DataPoint&
            End If
      End Select
      If Abs(ErrorCounts&) > ATolCounts& Then
         TripValue& = Format(mvLongChanData(DataPoint&), "0")
         If Span& > 0 Then TripValue& = SpanCounts&
         Failure$ = "DC level failure on channel " & Format(Channel%, "0") & _
         msChansMeasured & "." & vbCrLf & Format(SampleSet&, "0") & _
         " samples evaluated" & EOptionDesc$ & vbCrLf & "DC level = " & _
         Format(TripValue&, "0") & " counts at sample " & _
         Format(DataPoint& * mnNumChans, "0") & CurRange$ & "." & vbCrLf & _
         "Scripted limits = " & LimitString$ & _
         " LSBs" & LimitSpread$ & _
         vbCrLf & "Error = " & Format(ErrorCounts&, "0") & " LSBs."
         If NumSamples& - DataPoint& < 5 Then
            If Not ((mlTotalSamples Mod mnNumChans) = 0) Then
               Failure$ = Failure$ & vbCrLf & "Possible cause for failure: " & _
               "The number of samples specified in the script (" & Format(mlTotalSamples, "0") & _
               ") doesn't appear to be divisible by the number of channels specified (" & _
               Format(mnNumChans, "0") & ")."
            End If
         End If
         Message = Failure$
         FailEvalDC% = True
         Exit For
      End If
   Next DataPoint&
   If FailEvalDC% Then
      EvalDC = True
   Else
      Message = "DC level verified on channel " & Format(Channel%, "0") & _
      msChansMeasured & "." & vbCrLf & Format(SampleSet&, "0") & _
      " samples evaluated" & CurRange$ & EOptionDesc$ & vbCrLf & _
      "Scripted limits = " & LimitString$ & _
      " LSBs" & LimitSpread$ & _
      vbCrLf & "Error = " & Format(MaxErr&, "0") & " LSBs at sample " & Format(MaxSample&, "0") & "."
   End If

End Function

Function EvalAmplitude(Amplitude As Variant, ATol As Variant, EvalOption As Variant, DataType As Integer, Message As String) As Integer

   If (mlQueueCount > 0) And (mlQueueCount = mnNumChans) Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   FirstPoint& = 0
   
   If VarType(EvalOption) = vbString Then
      EmptyString% = (EvalOption = "")
   End If
   If IsEmpty(EvalOption) Or EmptyString% Then
      EvalParam& = 0
      Span& = 0
      EvaluationOption$ = ""
   Else
      ParseAll = Split(EvalOption, ";")
      NumOptions& = UBound(ParseAll)
      For CurOpt& = 0 To NumOptions&
         ThisOpt = Trim(ParseAll(CurOpt&))
         EvaluationType = Split(ThisOpt, " ")
         EvalParam& = UBound(EvaluationType)
         If EvalParam& = 0 Then
            Span& = Val(EvaluationType(0))
         Else
            EvaluationOption$ = LCase(EvaluationType(0))
            'following allows for "String = Value"
            'or "String Value" construct
            OptionValue& = EvaluationType(EvalParam&)
            Select Case EvaluationOption$
               Case "movingaverage", "moving"
                  Span& = OptionValue&
                  'EOptionDesc$ set below with AvgPoints& parameter
               Case "first", "firstpoint", "start"
                  FirstPoint& = OptionValue& / mnNumChans
                  EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
            End Select
         End If
      Next
   End If
   
   If ATol < 0 Then
      If ATol = -1 Then
         'Amplitude specifies maximum amplitude
         MaxOnly% = True
      Else
         'Amplitude specifies minimum amplitude
         MinOnly% = True
      End If
   End If
   'RangeVal% = mnRange
   NumSamples& = UBound(mvLongChanData)
   TolDataType% = DataType And &HF
   d& = (DataType And &HF0) / &H10
   AmpDataType% = d&
   For Item% = 1 To 2
      Parameter% = Choose(Item%, AmpDataType%, TolDataType%)
      ParameterVal = Choose(Item%, Amplitude, ATol)
      Select Case Parameter%
         Case 0, UNITPERCENT, UNITFLOAT  '1, 5
            If mnResolution > 16 Then
               dATude# = ParameterVal
               ParamCounts = VoltsToHiResCounts(mnResolution, RangeVal%, dATude#, mvCustomRange)
            Else
               ATude! = ParameterVal
               ParamCounts = VoltsToCounts(mnResolution, RangeVal%, ATude!, mvCustomRange)
            End If
            Units$ = " LSBs"
            ParamFormatString$ = "0"
            TFormatString$ = "0"
         Case UNITVOLTS '2
            ParamCounts = ParameterVal
            Units$ = " volts"
            ParamFormatString$ = "0.0#####"
            If Abs(Amplitude) < 0.001 Then ParamFormatString$ = "0.0##E+00"
         Case UNITDEGREES '3
            ParamCounts = ParameterVal
            Units$ = " °"
            'Units$ = " degrees"
            ParamFormatString$ = "0.0##"
            If Abs(ParameterVal) < 0.001 Then ParamFormatString$ = "0.0##E+00"
         Case UNITCOUNTS   '4
            ParamCounts = ParameterVal
            Units$ = " counts"
            ParamFormatString$ = "0"
         Case UNITLSBS   '6
            ParamCounts = ParameterVal
            Units$ = " LSBs"
            ParamFormatString$ = "0"
         Case Else
            Stop
            'not sure why this is necessary
            If mnResolution > 16 Then
               dATude# = ParameterVal
               ParamCounts = VoltsToHiResCounts(mnResolution, RangeVal%, dATude#, mvCustomRange)
            Else
               ATude! = ParameterVal
               ParamCounts = VoltsToCounts(mnResolution, RangeVal%, ATude!, mvCustomRange)
            End If
            Units$ = " V"
      End Select
      If Item% = 1 Then
         AmpCounts = ParamCounts
         AmpCountsString$ = Format(ParamCounts, ParamFormatString$)
         AUnits$ = Units$
      Else
         If Not mfBoardAmplGB = 0 Then
            LSBs& = mfBoardAmplGB
            BoardGB! = LSBs&
            If Parameter% = UNITLSBS Then BoardGB! = ConvertLSBs(LSBs&)
         End If
         TolVal& = ParamCounts
         Tolerance = TolVal&
         If Parameter% = UNITLSBS Then Tolerance = ConvertLSBs(TolVal&)
         ATolCounts = Tolerance + mlAmplGB + BoardGB!
         'ATolCountsString$ = Format(ATolCounts, "0")
         ATolCountsString$ = Format(ATolCounts, ParamFormatString$)
         TUnits$ = Units$
      End If
   Next Item%
   GlobalMvgAvg& = Span& * mlMvgAvgGB
   BoardMvgAvg& = Span& * mlBoardMvgAvgGB
   TotalMvgAvgComp& = GlobalMvgAvg& + BoardMvgAvg&
   If TotalMvgAvgComp& = 0 Then
      If Span& > 1 Then AvgPoints& = Span&
   Else
      AvgPoints& = TotalMvgAvgComp&
   End If
   
   AmplValMin = AmpCounts - ATolCounts: AmplValMax = AmpCounts + ATolCounts
   AmplValMinString$ = Format(AmplValMin, "0.0########")
   If Abs(AmplValMin) < 0.001 Then AmplValMinString$ = Format(AmplValMin, "0.0####E+00")
   AmplValMaxString$ = Format(AmplValMax, "0.0########")
   If Abs(AmplValMax) < 0.001 Then AmplValMaxString$ = Format(AmplValMax, "0.0####E+00")
   
   If RangeVal% = -2 Then
      CurRange$ = ""
   Else
      CurRange$ = " using range " & GetRangeString(RangeVal%)
   End If
   If AvgPoints& > 0 Then
      'use moving average
      EOptionDesc$ = ", moving average = " & Format(AvgPoints&, "0") & "."
   End If
   
   MaxAmpl = -16777216: MinAmpl = 16777216
   SampleSet& = (NumSamples& - FirstPoint&) + 1
   If SampleSet& < 2 Then
      MsgBox "Script is requesting evaluation of " & _
      Format(SampleSet&, "0") & " data points (from " & _
      Format(FirstPoint&, "0") & " to " & Format(NumSamples& + 1, "0") _
      & ").", vbOKOnly, "Invalid Script Configuration"
      EvalAmplitude = True
      Message = "Invalid script configuration error."
      Exit Function
   End If
   For DataPoint& = FirstPoint& To NumSamples& '- 1  'SampleSet& - 1
      If AvgPoints& > 0 Then
         'use moving average
         Span1& = DataPoint& + AvgPoints& - 1
         If (DataPoint& + AvgPoints&) > (NumSamples - 2) Then Exit For
         CumVal1& = 0: CumVal2& = 0
         For Element& = DataPoint& To Span1&
            CumVal1& = CumVal1& + mvLongChanData(Element&)
         Next Element&
         AmpVal = CumVal1& / AvgPoints&
      Else
         AmpVal = mvLongChanData(DataPoint&)
      End If
      If AmpVal > MaxAmpl Then
         MaxAmpl = AmpVal
         MaxSample& = DataPoint&
      End If
      If AmpVal < MinAmpl Then
         MinAmpl = AmpVal
         MinSample& = DataPoint&
      End If
   Next DataPoint&
   AmplPeak = MaxAmpl - MinAmpl
   AmplPeakString$ = Format(AmplPeak, "0.0######")
   If Abs(AmplPeak) < 0.001 Then AmplPeakString$ = Format(AmplPeak, "0.0##E+00")
   MaxAmplString$ = Format(MaxAmpl, "0.0##")
   If Abs(MaxAmpl) < 0.001 Then MaxAmplString$ = Format(MaxAmpl, "0.0##E+00")
   MinAmplString$ = Format(MinAmpl, "0.0##")
   If Abs(MinAmpl) < 0.001 Then MinAmplString$ = Format(MinAmpl, "0.0##E+00")

   If MinOnly% Or MaxOnly% Then
      If MaxOnly% Then
         Info$ = "Scripted limits = " & AmpCountsString$ & AUnits$ & " max."
         If AmplPeak > AmpCounts Then
            EvalAmplitude = True
         End If
      Else
         Info$ = "Scripted limits = " & AmpCountsString$ & AUnits$ & " min."
         If AmplPeak < AmpCounts Then
            EvalAmplitude = True
         End If
      End If
      result$ = "Amplitude read = " & AmplPeakString$ & _
      CurRange$ & msChansMeasured & "." & vbCrLf
      If EvalAmplitude Then
         Message = "Amplitude failure on channel " & Format(Channel%, "0") & _
         " comparing sample " & Format(MinSample& * mnNumChans, "0") & " and sample " & _
         Format(MaxSample& * mnNumChans, "0") & "." & vbCrLf & Format(SampleSet&, "0") & _
         " samples evaluated" & EOptionDesc$ & vbCrLf & result$ & Info$
      Else
         Message = "Amplitude verified on channel " & Format(Channel%, "0") & _
         msChansMeasured & "." & vbCrLf & "Value read = " & AmplPeakString$ & _
         CurRange$ & "." & vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & _
         EOptionDesc$ & vbCrLf & result$ & Info$
      End If
   Else
      ErrorValue = AmplPeak - AmpCounts
      ErrorValueString$ = Format(ErrorValue, "0.0##")
      If Abs(ErrorValue) < 0.001 Then ErrorValueString$ = Format(ErrorValue, "0.0##E+00")
      If (AmplPeak > AmplValMax) Or (AmplPeak < AmplValMin) Then
         If Not AmplPeak = 0 Then
            Prefix$ = "Amplitude failure on channel " & Format(Channel%, "0") & _
            " comparing sample " & Format(MinSample& * mnNumChans, "0") & " and sample " & _
            Format(MaxSample& * mnNumChans, "0") & "." & vbCrLf
         Else
            Prefix$ = "Amplitude failure on channel " & Format(Channel%, "0") & _
            " (no signal detected - " & MaxAmplString$ & " counts in " & _
            Format(SampleSet&, "0") & " samples evaluated.)" & vbCrLf
         End If
         If (NumSamples& - MinSample& < 5) Or (NumSamples& - MaxSample& < 5) Then
            If Not ((mlTotalSamples Mod mnNumChans) = 0) Then
               Failure$ = Failure$ & vbCrLf & "Possible cause for failure: " & _
               "The number of samples specified in the script (" & Format(mlTotalSamples, "0") & _
               ") doesn't appear to be divisible by the number of channels specified (" & _
               Format(mnNumChans, "0") & ")."
            End If
         End If
         Message = Prefix$ & Format(SampleSet&, "0") & " samples evaluated" & _
         EOptionDesc$ & vbCrLf & "Amplitude read = " & AmplPeakString$ & _
         CurRange$ & msChansMeasured & "." & vbCrLf & _
         "Scripted limits = " & AmpCountsString$ & " ±" & ATolCountsString$ & _
         TUnits$ & " (" & AmplValMinString$ & " to " & AmplValMaxString$ & ")." & _
         vbCrLf & "Error = " & ErrorValueString$ & AUnits$ & "." & Failure$
         EvalAmplitude = True
      Else
         Message = "Amplitude verified on channel " & Format(Channel%, "0") & _
         msChansMeasured & "." & vbCrLf & Format(SampleSet&, "0") & _
         " samples evaluated" & EOptionDesc$ & vbCrLf & "Value read = " & _
         AmplPeakString$ & CurRange$ & "." & vbCrLf & "Scripted limits = " & _
         AmpCountsString$ & " ±" & ATolCountsString$ & TUnits$ & " (" & _
         AmplValMinString$ & " to " & AmplValMaxString$ & ")  " & vbCrLf & _
         "Error = " & ErrorValueString$ & AUnits$ & "."
      End If
   End If
   
End Function

Function Eval32Delta(DeltaMin As Long, DeltaType As Long, _
EvalOption As Variant, Message As String) As Integer

   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
   Else
      Channel% = mnChanIndex
   End If
   NumSamples& = UBound(mvLongChanData)
   If NumSamples& < 1 Then
      'no sample to compare to
      Message = "Delta failure due to insufficient number of samples."
      DeltaFail% = True
   End If
   
   FirstPoint& = 0
   
   EOptionDesc$ = "."
   If VarType(EvalOption) = vbString Then
      EmptyString% = (EvalOption = "")
   End If
   If IsEmpty(EvalOption) Or EmptyString% Then
      EvalParam& = 0
      Span& = 0
      EvaluationOption$ = ""
   Else
      ParseAll = Split(EvalOption, ";")
      NumOptions& = UBound(ParseAll)
      For CurOpt& = 0 To NumOptions&
         ThisOpt = Trim(ParseAll(CurOpt&))
         EvaluationType = Split(ThisOpt, " ")
         EvalParam& = UBound(EvaluationType)
         If EvalParam& = 0 Then
            Span& = Val(EvaluationType(0))
         Else
            EvaluationOption$ = LCase(EvaluationType(0))
            'following allows for "String = Value"
            'or "String Value" construct
            OptionValue& = EvaluationType(EvalParam&)
            Select Case EvaluationOption$
               Case "movingaverage", "moving"
                  Span& = OptionValue&
                  EOptionDesc$ = ", moving average = " & Format(Span&, "0")
               Case "first", "firstpoint", "start"
                  FirstPoint& = OptionValue&
                  EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
               Case "ignore"
                  Span& = OptionValue&
                  IgnoreVal& = OptionValue&
               Case "tolerance", "tol"
                  MinDiff& = 65535
                  TolVal& = OptionValue&
                  EOptionDesc$ = ", tolerance set at " & Format(TolVal&, "0") & "."
            End Select
         End If
      Next
   End If
   
   If mnDataSet Then
      If Not (mvLongChanData(NumSamples&) > 1) Then
         'no data to compare to - exit without error
         Exit Function
      Else
         LastDataSample& = mvLongChanData(NumSamples&) - 1
      End If
   Else
      LastDataSample& = NumSamples&
   End If
   
   MinDelta& = 1000000
   SampleSet& = (LastDataSample& - FirstPoint&) + 1
   For DataPoint& = FirstPoint& To LastDataSample& - 1
      If (Span& > 0) And (DeltaType < 3) Then
         'use moving average
         Span1& = DataPoint& + (Span& - 1)
         If (DataPoint& + Span&) > (NumSamples - 2) Then Exit For
         CumVal1& = 0: CumVal2& = 0
         For Element& = DataPoint& To Span1&
            CumVal1& = CumVal1& + mvLongChanData(Element&)
         Next Element&
         For Element& = DataPoint& + 1 To Span1& + 1
            CumVal2& = CumVal2& + mvLongChanData(Element&)
         Next Element&
         AvgVal1& = CumVal1& / Span&
         AvgVal2& = CumVal2& / Span&
         DeltaVal& = AvgVal2& - AvgVal1&
      Else
         DeltaVal& = mvLongChanData(DataPoint& + 1) - mvLongChanData(DataPoint&)
      End If
      Select Case DeltaType
         Case 0
            If Abs(DeltaVal&) > DeltaMin Then
               If DeltaType = 0 Then   'ChangeExceeds
                  If (FailCount& < IgnoreVal&) Then
                     FailCount& = FailCount& + 1
                  Else
                     'find first exceeding DeltaMin only
                     Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
                     vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
                     vbCrLf & "Scripted limits = " & Format(DeltaMin, "0") & _
                     " counts change between samples maximum." & vbCrLf & _
                     "Difference = " & Format(DeltaVal&, "0") & " counts comparing sample " & _
                     Format(DataPoint&, "0") & " and sample " & Format(DataPoint& + 1, "0") & "."
                     DeltaFail% = True
                     Exit For
                  End If
               End If
            End If
            If Abs(DeltaVal&) > MaxDelta& Then
               MaxDelta& = DeltaVal&
               MDSample& = DataPoint&
               NumExceeded& = NumExceeded& + 1
            End If
         Case 1, 2
            If Abs(DeltaVal&) > MaxDelta& Then
               MaxDelta& = DeltaVal&
               MDSample& = DataPoint&
               NumExceeded& = NumExceeded& + 1
            End If
         Case 3
            'change must be non-positive or specified rollover minimum
            If DeltaVal& > 0 Then
               If Abs(DeltaVal&) < (DeltaMin - TolVal&) Then
                  'find first value only that's positive and not at least DeltaMin
                  If (Span& > Failures&) Then
                     Failures& = Failures& + 1
                  Else
                     Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
                     vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
                     vbCrLf & "Scripted limits = negative values with rollover at " & _
                     Format(DeltaMin, "0") & "." & vbCrLf & _
                     "Difference = " & Format(DeltaVal&, "0") & " counts comparing sample " & _
                     Format(DataPoint&, "0") & " and sample " & Format(DataPoint& + 1, "0") & "."
                     DeltaFail% = True
                     Exit For
                  End If
               End If
            Else
               If Abs(DeltaVal&) < MinDelta& Then
                  MinDelta& = Abs(DeltaVal&)
                  MDSample& = DataPoint&
               End If
            End If
         Case 4
            'change must be non-negative or specified rollover minimum
            If DeltaVal& < 0 Then
               If Abs(DeltaVal&) < (DeltaMin - TolVal&) Then
                  'find first value only that's not negative and not at least DeltaMin
                  If (Span& > Failures&) Then
                     Failures& = Failures& + 1
                  Else
                     Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
                     vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
                     vbCrLf & "Scripted limits = positive values with rollover at " & _
                     Format(DeltaMin, "0") & "." & vbCrLf & _
                     "Difference = " & Format(DeltaVal&, "0") & " counts comparing sample " & _
                     Format(DataPoint&, "0") & " and sample " & Format(DataPoint& + 1, "0") & "."
                     DeltaFail% = True
                     Exit For
                  End If
               End If
            Else
               If Abs(DeltaVal&) < MinDelta& Then
                  MinDelta& = Abs(DeltaVal&)
                  MDSample& = DataPoint&
               End If
            End If
         Case 5
            'must contain intervals of no change of at least DeltaMin length
            DeltaFail% = True
            If (DeltaVal& = 0) Then
               NumExceeded& = NumExceeded& + 1
               If (NumExceeded& + 1) > DeltaMin Then
                  DeltaFail% = False
                  MaxNumExceeded& = NumExceeded&
                  Exit For
               End If
               If NumExceeded& = 1 Then MDSample& = DataPoint&
            Else
               If NumExceeded& > MaxNumExceeded& Then
                  MaxNumExceeded& = NumExceeded&
                  MaxMDSample& = MDSample&
               End If
               NumExceeded& = 0
            End If
         Case 6
            'must contain at least one change of DeltaMin ±TolVal&
            TotalDiff& = DeltaVal& - DeltaMin
            If Abs(TotalDiff&) <= Abs(MinDiff&) Then
               MinDiff& = TotalDiff&
               ClosestDelta& = DeltaVal&
               MDSample& = DataPoint&
            End If
            AboveLow% = (DeltaVal& >= (DeltaMin - TolVal&))
            BelowHigh% = (DeltaVal& <= (DeltaMin + TolVal&))
            DeltaFail% = True
            If AboveLow% And BelowHigh% Then
               NumExceeded& = NumExceeded& + 1
               DeltaFail% = False
               Exit For
            End If
      End Select
   Next DataPoint&
   
   If Not DeltaFail% Then
      Select Case DeltaType
         Case 0   'fail if ChangeExceeds
            'if this failed, it would be trapped above
            Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
            vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
            vbCrLf & "Scripted limits = " & Format(DeltaMin, "0") & _
            " counts change maximum." & vbCrLf & "Max change in data = " & _
            Format(MaxDelta&, "0") & " counts at sample " & Format(MDSample&, "0") & "."
         Case 1   'fail if NoChangeExceeds
            If MaxDelta& < DeltaMin Then
               Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
               vbCrLf & "Scripted limits = " & Format(DeltaMin, "0") & " counts change " & _
               "between samples minimum." & vbCrLf & "Maximum difference in samples = " & _
               Format(MaxDelta&, "0") & " counts at sample " & Format(MDSample&, "0") & "."
               DeltaFail% = True
            Else
               Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & vbCrLf & _
               "Scripted limits = " & Format(DeltaMin, "0") & " counts change minimum." & _
               vbCrLf & "Max change in data = " & Format(MaxDelta&, "0") & _
               " counts at sample " & Format(MDSample&, "0") & "."
            End If
         Case 2   'Specific delta
            If Not (MaxDelta& = DeltaMin) Then
               Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & vbCrLf & _
               "Scripted limits = " & Format(DeltaMin, "0") & " counts change " & _
               "between samples exactly." & vbCrLf & "Difference in samples = " & _
               Format(MaxDelta&, "0") & " counts at sample " & Format(MDSample&, "0") & "."
               DeltaFail% = True
            Else
               Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & vbCrLf & _
               "Scripted limits = " & Format(DeltaMin, "0") & " counts change exactly." & _
               vbCrLf & "Change in data = " & Format(MaxDelta&, "0") & _
               " counts for all samples."
            End If
         Case 3
            'change must be negative or specified rollover minimum
            Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
            vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
            vbCrLf & "Scripted limits = negative values with rollover at " & _
            Format(DeltaMin, "0") & "." & vbCrLf & _
            "Minimum negative delta detected = " & Format(MinDelta&, "0") & _
            " counts at sample " & Format(DataPoint&, "0") & "."
         Case 4
            'change must be positive or specified rollover minimum
            Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
            vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
            vbCrLf & "Scripted limits = positive values with rollover at " & _
            Format(DeltaMin, "0") & "." & vbCrLf & _
            "Minimum positive delta detected = " & Format(MinDelta&, "0") & _
            " counts at sample " & Format(DataPoint&, "0") & "."
         Case 5
            'contains intervals of no change of at least DeltaMin length
            If MaxMDSample& = 0 Then MaxMDSample& = MDSample&
            Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
            vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
            vbCrLf & "Scripted limits = interval of no change of at least " & _
            Format(DeltaMin, "0") & " samples." & vbCrLf & _
            "Minimum interval detected from sample " & Format(MaxMDSample&, "0") & _
            " to at least sample " & Format(MaxMDSample& + MaxNumExceeded&, "0") & "."
         Case 6
            'at least one change of DeltaMin ±TolVal&
            Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
            vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
            vbCrLf & "Scripted limits = at least one change of " & _
            Format(DeltaMin, "0") & " ±" & Format(TolVal&, "0") & "." & vbCrLf & _
            "Closest change detected = " & Format(ClosestDelta&, "0") & _
            " at sample " & Format(MDSample&, "0") & "."
      End Select
   Else
      If DeltaType = 5 Then
         'contains no intervals of no change of at least DeltaMin length
         Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
         vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
         vbCrLf & "Scripted limits = interval of no change of at least " & _
         Format(DeltaMin, "0") & " samples." & vbCrLf
         If MaxNumExceeded& > 0 Then
            Message = Message & "Maximum interval detected = " & Format(MaxNumExceeded&, "0") & _
            " from sample " & Format(MaxMDSample&, "0") & " to sample " & _
            Format(MaxMDSample& + MaxNumExceeded&, "0") & "."
         Else
            Message = Message & "No interval of no change detected."
         End If
      End If
      If DeltaType = 6 Then
         Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
         vbCrLf & Format(SampleSet&, "0") & " samples evaluated" & EOptionDesc$ & _
         vbCrLf & "Scripted limits = at least one change of " & _
         Format(DeltaMin, "0") & " ±" & Format(TolVal&, "0") & "." & vbCrLf & _
         "Closest change detected = " & Format(ClosestDelta&, "0") & _
         " at sample " & Format(MDSample&, "0") & "."
      End If
   End If
   Eval32Delta = DeltaFail%

End Function

Function EvalDelta(DeltaMin As Single, DeltaType As Long, _
EvalOption As Variant, Convert As Integer, Message As String) As Integer

   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   NumSamples& = UBound(mvLongChanData)
   FirstPoint& = 0
   
   EOptionDesc$ = "."
   If VarType(EvalOption) = vbString Then
      EmptyString% = (EvalOption = "")
   End If
   If IsEmpty(EvalOption) Or EmptyString% Then
      EvalParam& = 0
      Span& = 0
      EvaluationOption$ = ""
   Else
      EvaluationType = Split(EvalOption, " ")
      EvalParam& = UBound(EvaluationType)
      If EvalParam& = 0 Then
         Span& = Val(EvaluationType(0))
      Else
         
         'following allows for "String = Value"
         'or "String Value" construct
         ParseAll = Split(EvalOption, ";")
         NumOptions& = UBound(ParseAll)
         For CurOpt& = 0 To NumOptions&
            ThisOpt = Trim(ParseAll(CurOpt&))
            EvaluationType = Split(ThisOpt, " ")
            EvalParam& = UBound(EvaluationType)
            EvaluationOption$ = LCase(EvaluationType(0))
            OptionValue& = EvaluationType(EvalParam&)
            Select Case EvaluationOption$
               Case "movingaverage", "moving"
                  Span& = OptionValue&
                  EOptionDesc$ = ", moving average = " & Format(Span&, "0")
               Case "first", "firstpoint", "start"
                  FirstPoint& = OptionValue&
                  EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
               Case "ignore"
                  Span& = OptionValue&
                  IgnoreVal& = OptionValue&
               Case "tolerance", "tol"
                  MinDiff& = 65535
                  TolVal& = OptionValue&
                  EOptionDesc$ = ", tolerance set at " & Format(TolVal&, "0") & "."
            End Select
         Next
      End If
   End If
   Select Case Convert
      Case 0
         DeltaCounts = VoltsToCounts(mnResolution, RangeVal%, DeltaMin) + TolVal&
         Units$ = " counts "
         DeltaCountsString$ = Format(DeltaCounts, "0")
      Case 1
         DeltaCounts = DeltaMin
         Units$ = " volts "
         DeltaCountsString$ = Format(DeltaCounts, "0.0##")
         If Abs(DeltaCounts) < 0.001 Then DeltaCountsString$ = Format(DeltaCounts, "0.0##E+00")
      Case 2
         DeltaCounts = DeltaMin
         Units$ = " degrees "
         DeltaCountsString$ = Format(DeltaCounts, "0.0##")
         If Abs(DeltaCounts) < 0.001 Then DeltaCountsString$ = Format(DeltaCounts, "0.0##E+00")
   End Select
   If mfRateReturned > 0 Then
      TestRate$ = " at " & Format(mfRateReturned, "0.0###") & " S/s"
   ElseIf mfRateRequested > 0 Then
      TestRate$ = " at " & Format(mfRateRequested, "0.0###") & " S/s"
   End If
   
   GlobalMvgAvg& = Span * mlMvgAvgGB
   BoardMvgAvg& = Span * mlBoardMvgAvgGB
   TotalMvgAvgComp& = GlobalMvgAvg& + BoardMvgAvg&
   If TotalMvgAvgComp& = 0 Then
      If Span > 1 Then AvgPoints& = Span
   Else
      AvgPoints& = TotalMvgAvgComp&
   End If
   
   For DataPoint& = FirstPoint& To NumSamples& - 1
      SampleIndex& = (DataPoint& * mnNumChans) + mnChanIndex
      If AvgPoints& > 0 Then
         'use moving average
         Span1& = DataPoint& + (AvgPoints& - 1)
         If (DataPoint& + AvgPoints&) > (NumSamples - 2) Then Exit For
         CumVal1& = 0: CumVal2& = 0
         For Element& = DataPoint& To Span1&
            CumVal1& = CumVal1& + mvLongChanData(Element&)
         Next Element&
         For Element& = DataPoint& + 1 To Span1& + 1
            CumVal2& = CumVal2& + mvLongChanData(Element&)
         Next Element&
         AvgVal1& = CumVal1& / AvgPoints&
         AvgVal2& = CumVal2& / AvgPoints&
         DeltaVal = AvgVal2& - AvgVal1&
      Else
         DeltaVal = mvLongChanData(DataPoint& + 1) - mvLongChanData(DataPoint&)
      End If
      If Abs(DeltaVal) > DeltaCounts Then
         If DeltaType = 0 Then   'ChangeExceeds
            If (FailCount& < IgnoreVal&) Then
               FailCount& = FailCount& + 1
            Else
                'find first exceeding DeltaMin only
                Message = "Delta failure on channel " & Format(Channel%, "0") & _
                " comparing sample " & Format(SampleIndex&, "0") & " and sample " & _
                Format(SampleIndex& + mnNumChans, "0") & "." & vbCrLf & "Difference = " & _
                Format(DeltaVal, "0.0##") & Units$ & msChansMeasured & TestRate$ & _
                "." & vbCrLf & "Moving average = " & Format(AvgPoints&, "0") & "." & _
                vbCrLf & "Scripted limits = " & _
                DeltaCountsString$ & Units$ & " change between samples maximum." & _
                vbCrLf & "Error = " & Format(Abs(DeltaVal) - DeltaCounts, "0") & " counts."
                'FirstPoint& = SampleIndex& - 50
                'LastPoint& = SampleIndex& + 50
                'If FirstPoint& < 0 Then FirstPoint& = 0
                'If LastPoint& > NumSamples& Then LastPoint& = NumSamples&
                'For i& = FirstPoint& To LastPoint&
                '   DataString$ = DataString$ & Format(i&, "0") & _
                '     "; " & Format(mvLongChanData(i&), "0") & vbCrLf
                'Next
                'Clipboard.Clear
                'Clipboard.SetText (DataString$)
                'MsgBox "Data available to paste.", vbInformation, "Clipboard Loaded"
                DeltaFail% = True
                Exit For
            End If
         End If
      End If
      If Abs(DeltaVal) > Abs(MaxDelta) Then
         MaxDelta = DeltaVal
         MDSample& = DataPoint&
         NumExceeded& = NumExceeded& + 1
         MaxDeltaString$ = Format(MaxDelta, "0.0##")
         If Abs(MaxDelta) < 0.001 Then MaxDeltaString$ = Format(MaxDelta, "0.0##E+00")
      End If
   Next DataPoint&
   
   If Not DeltaFail% Then
      Select Case DeltaType
         Case 0   'fail if ChangeExceeds
            'if this failed, it would be trapped above
            Message = "Delta verified on channel " & Format(Channel%, "0") & _
            vbCrLf & "Max change in data = " & MaxDeltaString$ & _
            Units$ & " at sample " & Format(MDSample& * mnNumChans, "0") & _
            ", moving average = " & Format(AvgPoints&, "0") & "." & vbCrLf & _
            "Data acquired" & msChansMeasured & TestRate$ & vbCrLf & _
            "Scripted limits = " & DeltaCountsString$ & Units$ & "change maximum."
         Case 1   'fail if NoChangeExceeds
            If Abs(MaxDelta) < DeltaCounts Then
               Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & "Maximum difference in samples = " & Format(MaxDelta, "0.0##") & _
               Units$ & " at sample " & Format(MDSample& * mnNumChans, "0") & _
               ", moving average = " & Format(AvgPoints&, "0") & "." & vbCrLf & _
               "Data acquired" & msChansMeasured & TestRate$ & vbCrLf & _
               "Scripted limits = " & DeltaCountsString$ & Units$ & "change " & _
               "between samples minimum."
               DeltaFail% = True
            Else
               Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & "Max change in data = " & Format(MaxDelta, "0.0##") & _
               Units$ & " at sample " & Format(MDSample& * mnNumChans, "0") & _
               ", moving average = " & Format(AvgPoints&, "0") & "." & vbCrLf & _
               "Data acquired" & msChansMeasured & TestRate$ & vbCrLf & _
               "Scripted limits = " & DeltaCountsString$ & Units$ & "change minimum."
            End If
         Case 2   'Specific delta
            If Not (MaxDelta = DeltaCounts) Then
               Message = "Delta failure on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & "Difference in samples = " & Format(MaxDelta, "0.0##") & _
               Units$ & " at sample " & Format(MDSample& * mnNumChans, "0") & _
               ", moving average = " & Format(AvgPoints&, "0") & "." & vbCrLf & _
               "Data acquired" & msChansMeasured & TestRate$ & vbCrLf & _
               "Scripted limits = " & DeltaCountsString$ & Units$ & "change " & _
               "between samples exactly."
               DeltaFail% = True
            Else
               Message = "Delta verified on channel " & Format(Channel%, "0") & "." & _
               vbCrLf & "Change in data = " & Format(MaxDelta, "0.0##") & _
               Units$ & " at sample " & Format(MDSample& * mnNumChans, "0") & _
               ", moving average = " & Format(AvgPoints&, "0") & "." & vbCrLf & _
               "Data acquired" & msChansMeasured & TestRate$ & vbCrLf & _
               "Scripted limits = " & DeltaCountsString$ & Units$ & "change exactly."
            End If
      End Select
   End If
   EvalDelta = DeltaFail%

End Function

Function EvalTrigPoint(ByVal Polarity As Integer, ByVal Threshold As Single, _
ByVal Guardband As Variant, ByVal TLimit As Long, ByVal TTol As Long, _
EvalOption As Variant, Message As String) As Integer

   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex + mnFirstChan
      RangeVal% = mnRange
   End If
   
   AmpTol$ = Guardband  'value may be changed by EvalOption below
   pATolVal$ = ParseUnits(AmpTol$, ATolUnitType%)
   If ATolUnitType% = UNITCOUNTS Then
      LSBs& = Val(AmpTol$)
      GuardCounts& = ConvertLSBs(LSBs&)
   Else
      Gband! = ConvertStringToType(pATolVal$, ATolUnitType%)
      GuardCounts& = VoltsToCounts(mnResolution, mnRange, Gband!)
   End If

   FirstPoint& = 0
   If VarType(EvalOption) = vbString Then
      EmptyString% = (EvalOption = "")
   End If
   If IsEmpty(EvalOption) Or EmptyString% Then
      EvalParam& = 0
      Span& = 0
      EvaluationOption$ = ""
   Else
      EvaluationType = Split(EvalOption, " ")
      EvalParam& = UBound(EvaluationType)
      If EvalParam& = 0 Then
         Span& = Val(EvaluationType(0))
      Else
         EvaluationOption$ = LCase(EvaluationType(0))
         'following allows for "String = Value"
         'or "String Value" construct
         OptionValue& = EvaluationType(EvalParam&)
         Select Case EvaluationOption$
            Case "first", "firstpoint", "start"
               FirstPoint& = OptionValue&
               EOptionDesc$ = ", starting at sample " & Format(FirstPoint&, "0") & "."
            Case "leveltol", "leveltolerance"
               If OptionValue& > 0 Then GuardCounts& = OptionValue&
         End Select
      End If
   End If
   
   NumSamples& = UBound(mvLongChanData)
   ThrshCounts& = GetCounts(mnResolution, RangeVal%, Threshold)
   
   If Not (TLimit Mod mnNumChans) = 0 Then
      'the trigger point should be divisible by number of channels
      Warning$ = vbCrLf & vbCrLf & "Better results may be obtained if the PretrigCount " & _
      "is evenly divisible by number of channels."
   End If
   TrigSample& = TLimit '/ mnNumChans
   If TLimit = 0 Then
      'no pretrigger, so limit will be 0 to TTol + TTol
      TTarget& = TTol
      TLimitMin& = TTarget& - TTol: TLimitMax& = TTarget& + TTol
   Else
      TTarget& = TLimit
      TLimitMax& = TrigSample& + TTol: TLimitMin& = TrigSample& - TTol
   End If
   LowTrig& = ThrshCounts& - GuardCounts&: HighTrig& = ThrshCounts& + GuardCounts&
   If (Polarity% = GATENEGHYS) Or (Polarity% = GATEOUTWINDOW) Then
      Gband! = ConvertStringToType(pATolVal$, ATolUnitType%)
      LowTrig& = GetCounts(mnResolution, mnRange, Gband!)
      TLimitMax& = NumSamples&
      TTarget& = TTol
   End If
   If (Polarity% = GATEPOSHYS) Or (Polarity% = GATEINWINDOW) Then
      Gband! = ConvertStringToType(pATolVal$, ATolUnitType%)
      HighTrig& = GetCounts(mnResolution, mnRange, Gband!)
      TLimitMax& = NumSamples&
      TTarget& = TTol
   End If
   
   EvalInfoText$ = vbCrLf & "Evaluated " & Format(NumSamples& + 1, "0") & _
   " data points" & " on channel " & Format(Channel%, "0") & "."
   TrigLimitText$ = Format(ThrshCounts&, "0") & " ±" & Format(GuardCounts&, "0") & _
   " counts at " & Format(TTarget& * mnNumChans, "0") & " ±" & Format(TTol, "0") & _
   " samples."
   TargetValue& = ThrshCounts&
   Select Case Polarity
      Case TRIGABOVE
         TypeText$ = "Trigger"
         TrigText$ = " above "
      Case TRIGBELOW
         TypeText$ = "Trigger"
         TrigText$ = " below "
      Case GATEABOVE
         TypeText$ = "Gate"
         TrigText$ = " above "
         ScriptLimitText$ = Format(ThrshCounts&, "0") & " - " & _
         Format(GuardCounts&, "0") & " counts"
         TLimitMax& = NumSamples&
      Case GATEBELOW
         TypeText$ = "Gate"
         TrigText$ = " below "
         ScriptLimitText$ = Format(ThrshCounts&, "0") & " + " & _
         Format(GuardCounts&, "0") & " counts"
         TLimitMax& = NumSamples&
      Case GATENEGHYS, GATEPOSHYS
         TypeText$ = "Gate"
         TrigText$ = " from "
         TargetValue& = ThrshCounts&
         RangeText$ = Format(ThrshCounts&, "0") & _
         " to " & Format(HighTrig&, "0")
         ScriptLimitText$ = Format(HighTrig&, "0") & " - " & _
         Format(GuardCounts&, "0") & " falling to " & Format(ThrshCounts&, "0") & _
         " within " & Format(TTol * mnNumChans, "0") & " samples"
         If Polarity = GATENEGHYS Then
            TargetValue& = LowTrig&
            RangeText$ = Format(LowTrig&, "0") & _
            " to " & Format(ThrshCounts&, "0")
            ScriptLimitText$ = Format(LowTrig&, "0") & " + " & _
            Format(GuardCounts&, "0") & " rising to " & Format(ThrshCounts&, "0") & _
            " within " & Format(TTol * mnNumChans, "0") & " samples"
         End If
         TrigLimitText$ = RangeText$ & " counts within " & Format(TTarget& * mnNumChans, "0") & _
         " samples"
      Case GATEHIGH
         TypeText$ = "Gate"
         TrigText$ = " above "
      Case GATELOW
         TypeText$ = "Gate"
         TrigText$ = " below "
      Case TRIGPOSEDGE
         TypeText$ = "Trigger"
         TrigText$ = " rising to "
      Case TRIGNEGEDGE
         TypeText$ = "Trigger"
         TrigText$ = " falling from "
      Case GATEINWINDOW
         TypeText$ = "Gate"
         TrigText$ = " inside "
         TargetValue& = ThrshCounts&
         RangeText$ = Format(ThrshCounts&, "0") & _
         " and " & Format(HighTrig&, "0")
         ScriptLimitText$ = Format(HighTrig&, "0") & _
         " and " & Format(ThrshCounts&, "0") & " ± " & Format(GuardCounts&, "0")
         TrigLimitText$ = RangeText$ & " counts"
      Case GATEOUTWINDOW
         TypeText$ = "Gate"
         TrigText$ = " outside "
         TargetValue& = LowTrig&
         RangeText$ = Format(LowTrig&, "0") & _
         " and " & Format(ThrshCounts&, "0")
         ScriptLimitText$ = Format(LowTrig&, "0") & _
         " and " & Format(ThrshCounts&, "0") & " ± " & Format(GuardCounts&, "0")
         TrigLimitText$ = RangeText$ & " counts"
   End Select
   
   DataOffset& = FirstPoint& / mnNumChans
   For CurPoint& = TLimitMin& To TLimitMax&
      CurValue& = mvLongChanData(CurPoint& + DataOffset&)
      Select Case Polarity
         Case TRIGABOVE
            If (CurValue& < LowTrig&) And (CurPoint& > TLimit) Then
               PolarityFail% = True
               PolFail$ = " should be TRIGABOVE but TRIGBELOW data characteristics detected"
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurPoint& > TLimit) Then
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurValue& > ThrshCounts&) Then
               Triggered% = True
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case TRIGBELOW
            If (CurValue& > HighTrig&) And (CurPoint& > TLimit) Then
               PolarityFail% = True
               PolFail$ = " should be TRIGBELOW but TRIGABOVE data characteristics detected"
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurPoint& > TLimit) Then
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If CurValue& < ThrshCounts& Then
               Triggered% = True
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case GATENEGHYS
            If (CurValue& < (LowTrig& - GuardCounts&)) Then
               PolarityFail% = True
               PolFail$ = " should be GATENEGHYS but values below " & Format(LowTrig&, "0") & " detected"
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurValue& < (LowTrig& + GuardCounts&)) Then
               Triggered% = True
               TrigReference& = CurPoint&
               TrigValue& = CurValue&
            End If
            If (CurValue& > ThrshCounts&) And Triggered% Then
               TrigPoint& = CurPoint& - TrigReference&
               Exit For
            End If
            If Not Triggered% Then
               If (CurValue& < TrigValue&) Or (CurPoint& = 0) Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
               End If
            End If
         Case GATEPOSHYS
            If (CurValue& > (HighTrig& + GuardCounts&)) Then
               PolarityFail% = True
               PolFail$ = " should be GATEPOSHYS but values above " & Format(HighTrig&, "0") & " detected"
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurValue& > (HighTrig& - GuardCounts&)) Then
               Triggered% = True
               TrigReference& = CurPoint&
               TrigValue& = CurValue&
            End If
            If (CurValue& < ThrshCounts&) And Triggered% Then
               TrigPoint& = CurPoint& - TrigReference&
               'TrigValue& = CurValue&
               Exit For
            End If
            If Not Triggered% Then
               If CurValue& > TrigValue& Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
               End If
            End If
         Case GATEABOVE
            If (CurValue& < (ThrshCounts& - GuardCounts&)) Then
               PolarityFail% = True
               PolFail$ = " should be GATEABOVE but values below " & Format(ThrshCounts&, "0") & " detected"
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurValue& > ThrshCounts&) Then
               Triggered% = True
               If (CurValue& < TrigReference&) Or (TrigPoint& = 0) Then
                  TrigPoint& = CurPoint&
                  TrigReference& = CurValue&
               End If
            End If
         Case GATEBELOW
            If (CurValue& > (ThrshCounts& + GuardCounts&)) Then
               PolarityFail% = True
               PolFail$ = " should be GATEBELOW but values above " & Format(ThrshCounts&, "0") & " detected"
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
            If (CurValue& < ThrshCounts&) Then
               Triggered% = True
               If (CurValue& > TrigReference&) Or (TrigPoint& = 0) Then
                  TrigPoint& = CurPoint&
                  TrigReference& = CurValue&
               End If
            End If
         Case TRIGHIGH, GATEHIGH
            If (CurValue& < LowTrig&) And (CurPoint& > TLimit) Then
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               PolarityFail% = True
               Exit For
            End If
            If (CurValue& > ThrshCounts&) Then
               Triggered% = True
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case TRIGLOW, GATELOW
            If (CurValue& > HighTrig&) And (CurPoint& > TLimit) Then
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               PolarityFail% = True
               Exit For
            End If
            If (CurValue& < ThrshCounts&) Then
               Triggered% = True
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case GATEINWINDOW
            If CurValue& > (HighTrig& + GuardCounts&) Then
               TargetValue& = HighTrig&
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               PolarityFail% = True
               Exit For
            End If
            If CurValue& < (LowTrig& - GuardCounts&) Then
               TargetValue& = LowTrig&
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               PolarityFail% = True
               Exit For
            End If
            If (CurPoint& = TLimitMax&) Then
               Triggered% = True
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case GATEOUTWINDOW
            If ((CurValue& < (LowTrig& - GuardCounts&)) And (CurValue& > (ThrshCounts& + GuardCounts&))) Then
               If (CurValue& - ThrshCounts&) > (LowTrig& - CurValue&) Then
                  TargetValue& = LowTrig&
               Else
                  TargetValue& = ThrshCounts&
               End If
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               PolarityFail% = True
               Exit For
            End If
            If (CurPoint& = TLimitMax&) Then
               Triggered% = True
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case TRIGPOSEDGE
            If Triggered% Then
               If (CurValue& < TrigValue&) Then
                  PolarityFail% = True
                  PolFail$ = " should be rising edge, but trigger value is " & _
                  Format(TrigValue&, "0.0###") & " and next value is " & _
                  Format(CurValue&, "0.0###") & "."
               End If
               Exit For
            Else
               If Not (TLimitMax& > CurPoint&) Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
                  TrigFail% = True
                  Exit For
               End If
            End If
            If TrigInit% Then
               If CurValue& > ThrshCounts& Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
                  Triggered% = True
               End If
            End If
            If (CurValue& < LowTrig&) Then TrigInit% = True
            If (TLimit = 0) Then
               TrigInit% = False
               If CurValue& > LowTrig& Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
                  Triggered% = True
               End If
            End If
            If Triggered% Then
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
         Case TRIGNEGEDGE
            If Triggered% Then
               If (CurValue& > TrigValue&) Then
                  PolarityFail% = True
                  PolFail$ = " should be falling edge, but trigger value is " & _
                  Format(TrigValue&, "0.0###") & " and next value is " & _
                  Format(CurValue&, "0.0###") & "."
               End If
               Exit For
            Else
               If Not (TLimitMax& > CurPoint&) Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
                  TrigFail% = True
                  Exit For
               End If
            End If
            If TrigInit% Then
               If CurValue& < ThrshCounts& Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
                  Triggered% = True
                  'If (CurValue& < LowTrig&) Or (CurValue& > HighTrig&) Then
                  '   LevelFail% = True
                  '   PolFail$ = " should occur between " & Format(LowTrig&, "0.0###") & _
                  '   " and " & Format(HighTrig&, "0.0###") & _
                  '   " but trigger value is " & Format(TrigValue&, "0.0###") & "."
                  Exit For
                  'End If
               End If
            End If
            If (CurValue& > HighTrig&) Then TrigInit% = True
            If (TLimit = 0) Then
               TrigInit% = False
               If CurValue& < HighTrig& Then
                  TrigPoint& = CurPoint&
                  TrigValue& = CurValue&
                  Triggered% = True
               End If
            End If
            If Triggered% Then
               TrigPoint& = CurPoint&
               TrigValue& = CurValue&
               Exit For
            End If
      End Select
   Next CurPoint&
   
   'ActualPoint& = (TrigPoint& * mnNumChans)
   'no longer compensate for chans here
   'Pretrig samples is divided by chans if not mnSimIn
   TimeError& = TrigPoint& - TLimit
   ValueError& = TrigValue& - TargetValue&
   If Not Triggered% Then
      Description$ = TypeText$ & " parameter not met. Value at sample " & Format((TrigPoint& + FirstPoint&) * mnNumChans, "0") & " on channel " _
      & Format(Channel%, "0") & " = " & Format(TrigValue&, "0") & " counts." & vbCrLf & _
      TypeText$ & TrigText$ & TrigLimitText$ & " expected," & msChansMeasured & _
      "." & EvalInfoText$ & vbCrLf & "Scripted limits = " & TypeText$ & TrigText$ & ScriptLimitText$ & _
      ". " & vbCrLf & "Error = " & Format(ValueError&, "0") & " counts."
      FailTrig% = True
   End If
   If (TrigPoint& < TLimitMin&) Or (TrigPoint& > TLimitMax&) Then
      Description$ = TypeText$ & " point failure on channel " & Format(Channel%, "0") & _
      ". " & TypeText$ & TrigText$ & Format(ThrshCounts&, "0") & " detected at sample " & _
      Format(TrigPoint& * mnNumChans, "0") & "," & msChansMeasured & _
      "." & EvalInfoText$ & vbCrLf & _
      "Scripted limits = " & TypeText$ & TrigText$ & ScriptLimitText$ & _
      "." & vbCrLf & "Error = " & Format(ValueError&, "0") & " counts."
      FailTrig% = True
   End If
   If Not FailTrig% Then
      If PolarityFail% Then
         Description$ = TypeText$ & " failure on channel " & _
         Format(Channel%, "0") & "." & vbCrLf & "Polarity" & _
         PolFail$ & vbCrLf & "Value at sample " & Format((TrigPoint& _
         + FirstPoint&) * mnNumChans, "0") & " on channel " _
         & Format(Channel%, "0") & " = " & Format(TrigValue&, "0") & _
         " counts." & vbCrLf & "Scripted limits = " & TypeText$ & _
         TrigText$ & ScriptLimitText$ & "." & vbCrLf & _
         "Error = " & Format(ValueError&, "0") & " counts."
         FailTrig% = True
      End If
      If LevelFail% Then
         Description$ = TypeText$ & "level failure on channel " & _
         Format(Channel%, "0") & " at sample " & Format(TrigPoint& * _
         mnNumChans, "0") & "." & vbCrLf & "Trigger level" & PolFail$ & _
         EvalInfoText$ & vbCrLf & "Scripted limits = " & TypeText$ & _
         TrigText$ & ScriptLimitText$ & "." & vbCrLf & "Error = " & _
         Format(TimeError&, "0") & " samples, " & Format(ValueError&, "0") & " counts."
         FailTrig% = True
      End If
   End If
   
   If Not FailTrig% Then
      If Not TimeError& < 0 Then NumSign$ = "+"
      Description$ = TypeText$ & " verified on channel " & _
      Format(Channel%, "0") & msChansMeasured & "." & EvalInfoText$ & _
      vbCrLf & TypeText$ & " value " & Format(TrigValue&, "0") & _
      " occurred at sample " & Format((TrigReference& + TrigPoint&) * _
      mnNumChans, "0") & "." & vbCrLf & "Scripted limits = " & _
      TypeText$ & TrigText$ & ScriptLimitText$ & "."
   End If
   EvalTrigPoint = FailTrig%
   Message = Description$ & Warning$
   
End Function

Function EvalTime(ByVal Threshold As Variant, ByVal Guardband As Variant, _
ByVal DataType As Integer, ByVal SourceFreq As Single, ByVal TLimit As Variant, _
ByVal TTol As Long, EvalOption As String, Message As String) As Integer

   Dim OptionValue As Variant
   
   RateFactor% = 1
   FreqOfSource! = SourceFreq
   FirstPoint& = 0
   
   If Not EvalOption = "" Then
      EvaluationType = Split(EvalOption, " ")
      EvalParam& = UBound(EvaluationType)
      EOptionDesc$ = "."
      If EvalParam& = 0 Then
         Span& = Val(EvaluationType(0))
      Else
         EvaluationOption$ = LCase(EvaluationType(0))
         'following allows for "String = Value"
         'or "String Value" construct
         EvaluationValue$ = EvaluationType(EvalParam&)
         If IsNumeric(EvaluationValue$) Then
            OptionValue = Val(EvaluationType(EvalParam&))
         Else
            OptionValue = 0
         End If
         Select Case EvaluationOption$
            Case "movingaverage", "moving"
               Span& = OptionValue
               'EOptionDesc$ set below with AvgPoints& parameter
            Case "first", "firstpoint", "start"
               CurPoint& = OptionValue
               EOptionDesc$ = ", starting at sample " & Format(CurPoint&, "0") & "."
            Case "percenttol"
               Percentage! = OptionValue / 100
         End Select
      End If
   End If
   If (mlRateGB < 0) Then
      RateGuardBand& = Abs(mlRateGB) - 1
      UseReqRate% = True
   Else
      RateGuardBand& = mlRateGB
   End If
   If (mlBoardRateGB < 0) Then
      BoardRateGB& = Abs(mlBoardRateGB) - 1
      UseReqRate% = True
   Else
      BoardRateGB& = mlBoardRateGB
   End If
   If Not InStr(1, msOptions, "EXTCLOCK") = 0 Then XClock% = True
   'If Not InStr(1, msOptions, "HIGHRESRATE") = 0 Then RateDivide% = True
   Divisor! = 1
   'If RateDivide% Then Divisor! = 1000
   If (TTol > 10000) Then
      PercTol! = TTol / 100000000
      TTol = 0
      UsePercentage% = True
   End If
   If (TTol < 0) Then   'Or ((mnSimIn = 0) And XClock%)
      'rate is dependant on number of channels
      RateFactor% = mnNumChans
   End If
   If mlQueueCount > 0 Then
      Channel% = mnGainQueue(0, mnChanIndex)
      RangeVal% = mnGainQueue(1, mnChanIndex)
   Else
      Channel% = mnChanIndex
      RangeVal% = mnRange
   End If
   'RangeVal% = mnRange
   NumSamples& = UBound(mvLongChanData)
   ThreshDataType% = (DataType And &HF0) / &H10
   GuardDataType% = DataType And &HF
   If (ThreshDataType% = 0) Or (ThreshDataType% = UNITFLOAT) Or (ThreshDataType% = UNITVOLTS) Then
      TH! = Threshold
      If Not IsEmpty(mvCustomRange) Then
         ThrshCounts = GetCounts(mnResolution, RangeVal%, TH!, mvCustomRange)
      Else
         ThrshCounts = GetCounts(mnResolution, RangeVal%, TH!)
      End If
      If Not mfBoardAmplGB = 0 Then LSBs& = mfBoardAmplGB
      BoardGB! = ConvertLSBs(LSBs&)
      LowTrig = ThrshCounts - Guardband - mlAmplGB - BoardGB!
      HighTrig = ThrshCounts + Guardband + mlAmplGB + BoardGB!
   Else
      ThrshCounts = Threshold
      LowTrig = ThrshCounts - Guardband
      HighTrig = ThrshCounts + Guardband
   End If
   TimeTolerance& = Abs(TTol) + RateGuardBand& + BoardRateGB&
   Dim PosTrigSamples() As Long
   Dim NegTrigSamples() As Long

   Searching% = True
   CurPoint& = FirstPoint&

   Do
      'find the first point not within the guardband
      If DataPoint& >= NumSamples& Then
         Searching% = False
      Else
         For DataPoint& = CurPoint& To NumSamples&
            CurValue& = mvLongChanData(DataPoint&)
            If (CurValue& < LowTrig) And (Not RiseFound%) Then
               rising% = True
               RiseFound% = True
               FallFound% = False
               CurPoint& = DataPoint&
               Exit For
            End If
            If (CurValue& > HighTrig) And (Not FallFound%) Then
               Falling% = True
               FallFound% = True
               RiseFound% = False
               CurPoint& = DataPoint&
               Exit For
            End If
         Next DataPoint&
      End If
      Select Case True
         Case rising%
            rising% = False
            TrgFound% = FindATrigger(TRIGABOVE, CurPoint&, NumSamples&, ThrshCounts, DataPoint&)
            CurPoint& = DataPoint&
            If TrgFound% Then
               ReDim Preserve PosTrigSamples(NumPosTrigs&)
               PosTrigSamples(NumPosTrigs&) = DataPoint&
               NumPosTrigs& = NumPosTrigs& + 1
            End If
         Case Falling%
            Falling% = False
            TrgFound% = FindATrigger(TRIGBELOW, CurPoint&, NumSamples&, ThrshCounts, DataPoint&)
            CurPoint& = DataPoint&
            If TrgFound% Then
               ReDim Preserve NegTrigSamples(NumNegTrigs&)
               NegTrigSamples(NumNegTrigs&) = DataPoint&
               NumNegTrigs& = NumNegTrigs& + 1
            End If
         Case Else
            Searching% = False
      End Select
   Loop While Searching% And Not FailTime%
   
   If (NumPosTrigs& < 2) Or (NumNegTrigs& < 2) Then
      WarnSlowSource% = True
      Warning$ = vbCrLf & vbCrLf & _
      "________________________________________________" & vbCrLf & _
      "Warning! - The rate of the signal source " & _
      "may be too low. Check script values for signal setup." & _
      vbCrLf & "May also be caused by evaluating too few samples." & _
      vbCrLf & "Add samples if more are available than specified in the script." & _
      vbCrLf & "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯"
      'chars 175 and 95
   End If
   
   If (NumPosTrigs& > 1) Or (NumNegTrigs& > 1) Then
      TotalTrigs& = (NumNegTrigs& - 1) + (NumPosTrigs& - 1)
      ReDim LongVarDiffs(TotalTrigs& - 1) As Variant
      For VarSample& = 0 To NumNegTrigs& - 2
         Diff& = NegTrigSamples(VarSample& + 1) - NegTrigSamples(VarSample&)
         If (VarSample& > 0) And Not SamplesDefined% Then
            If Diff& = LongVarDiffs(VarSample& - 1) Then
               Sample1& = NegTrigSamples(VarSample& - 1) * mnNumChans
               Sample2& = NegTrigSamples(VarSample&) * mnNumChans
               SamplesDefined% = True
            End If
         End If
         LongVarDiffs(VarSample&) = Diff&
      Next
      For VarSample& = 0 To NumPosTrigs& - 2
         Diff& = PosTrigSamples(VarSample& + 1) - PosTrigSamples(VarSample&)
         If (VarSample& > 0) And Not SamplesDefined% Then
            If Diff& = LongVarDiffs(((NumNegTrigs& - 1) + VarSample&) - 1) Then
               Sample1& = PosTrigSamples(VarSample& - 1) * mnNumChans
               Sample2& = PosTrigSamples(VarSample&) * mnNumChans
               SamplesDefined% = True
            End If
         End If
         LongVarDiffs((NumNegTrigs& - 1) + VarSample&) = Diff&
      Next
      QuickSortVariants LongVarDiffs, 0, TotalTrigs& - 1
      For AllSamps& = 0 To TotalTrigs& - 2
         'in any two agree, use that value
         'otherwise, use the median
         If LongVarDiffs(AllSamps& + 1) = LongVarDiffs(AllSamps&) Then
            Diff& = LongVarDiffs(AllSamps&)
            FoundPair% = True
            Match& = Match& + 1
         Else
            If Match& > MatchCount& Then
               ConvergeVal& = Diff&
               MatchCount& = Match&
            End If
            Match& = 0
         End If
      Next
      If Not FoundPair% Then
         Median& = TotalTrigs& \ 2
         Diff& = LongVarDiffs(Median&)
         Convergence$ = "Median result from " & Format(TotalTrigs&, "0") & " data points."
      Else
         If Match& > MatchCount& Then
            ConvergeVal& = Diff&
            MatchCount& = Match&
         End If
         'added " * RateFactor%" to following line 4/1/2010
         Diff& = ConvergeVal& '* RateFactor%
         Convergence$ = "Converged to " & Format(Diff&, "0") & " on " & Format(MatchCount&, "0") & _
         " of " & Format(TotalTrigs&, "0") & " data point pairs."
      End If
   Else
      FailSlowSource% = True
   End If
   If Not FailSlowSource% Then
      'Spread& = Sample2& - Sample1&
      CalcRate! = Diff& * FreqOfSource! * RateFactor%
      If (TLimit = 0) Then
         'not external clock so use return value
         RateTarget! = mfRateReturned / Divisor!
         'following for troubleshooting when incorrect rate returned
         If UseReqRate% Then RateTarget! = mfRateRequested / Divisor!
      Else
         'rate is unknown so must be supplied by script
         RateTarget! = TLimit / Divisor!
      End If
      If FreqOfSource! = 0 Then
         MsgBox "Source frequency is set to zero. Check script for " & _
         "uninitialized variables.", vbOKOnly, "Bad Source Frequency"
         gnErrFlag = True
         Exit Function
      End If
      Divisor! = 1
      If RateDivide% Then Divisor! = 1000
      IdealSamplesPer! = (RateTarget! / FreqOfSource!) / RateFactor%
      SamplesPerError! = (Diff&) - IdealSamplesPer!   '* RateFactor%
      If UsePercentage% Then
         TimeTolerance& = IdealSamplesPer! * PercTol!
         If TimeTolerance& = 0 Then TimeTolerance& = 1
      End If
      If Percentage! > 0 Then
         AddTimeTol& = IdealSamplesPer! * Percentage!
      End If
      RateTolerance! = FreqOfSource! * (TimeTolerance& + AddTimeTol&)
      MaxRate! = FreqOfSource! * (IdealSamplesPer! + TimeTolerance& + AddTimeTol&)
      MinRate! = FreqOfSource! * (IdealSamplesPer! - TimeTolerance& - AddTimeTol&)
      RateErr! = CalcRate! - RateTarget!
      If Abs(SamplesPerError!) > (TimeTolerance& + AddTimeTol&) Then
         CalcFail$ = vbCrLf & "Rate calculated = " & Format(CalcRate!, "0.0#") & _
         " S/s (" & Format(Diff&, "0") & " S/~). Source freq = " & _
         Format(FreqOfSource!, "0.0####") & " Hz. " & vbCrLf & _
         "Error = " & Format(RateErr!, "0.0#") & " S/s (" & Format(SamplesPerError!, "0.0##") & _
         " Samples per cycle)."
         If Abs(FreqOfSource!) > 0.5 Then
            If Not (RateTarget! Mod FreqOfSource!) = 0 Then
               FailTip$ = vbCrLf & vbCrLf & _
                  "Better results may be obtained if the sample rate " & _
                  "is evenly divisible by the source frequency."
            End If
         End If
         FailTime% = True
      End If
   Else
      Description$ = "Rate could not be calculated on channel " _
         & Format(Channel%, "0") & " evaluating " & _
         Format(NumSamples& + 1, "0") & " data points." & vbCrLf & _
         "Check that the desired frequency is available at the device input " & _
         vbCrLf & "and that enough samples are specified."
   End If
   
   If FailTime% Then
      Description$ = "Rate failure on channel " & Format(Channel%, "0") & _
      msChansMeasured & "." & vbCrLf & "Evaluated " & Format(NumSamples& + 1, "0") & _
      " data points, " & Format(NumPosTrigs&, "0") & " positive, " & Format(NumNegTrigs&, "0") & _
      " negative triggers found." & vbCrLf & "Trigger window = " & Format(LowTrig, "0") & _
      " to " & Format(HighTrig, "0") & " counts." & vbCrLf & Convergence$ & vbCrLf & _
      "Requested " & mfRateRequested & " S/s, Returned " & mfRateReturned & " S/s." & _
      vbCrLf & "Scripted limits = " & _
      Format(RateTarget!, "0.0###") & "  ±" & Format(RateTolerance!, "0.0#") & " S/s (" & _
      Format(MinRate!, "0.0#") & " to " & Format(MaxRate!, "0.0#") & " S/s)," & vbCrLf & _
      Format(IdealSamplesPer!, "0.0##") & " ±" & Format(TimeTolerance& + AddTimeTol&, "0") & _
      " samples per cycle." & _
      vbCrLf & CalcFail$ & FailTip$
   Else
      If Not FailSlowSource% Then
         Description$ = "Rate verified on channel " & Format(Channel%, "0") & msChansMeasured & "." & _
         vbCrLf & "Evaluated " & Format(NumSamples& + 1, "0") & " data points." & _
         vbCrLf & "Calculated rate = " & Format(CalcRate!, "0.0#") & " S/s (" & Format(Diff&, "0") & " S/~)." & _
         " Requested " & mfRateRequested & " S/s, Returned " & mfRateReturned & " S/s." & _
         vbCrLf & Convergence$ & vbCrLf & "Scripted limits = " & _
         Format(RateTarget!, "0.0####") & "  ±" & Format(RateTolerance!, "0.0##") & " S/s (" & _
         Format(MinRate!, "0.0#") & " to " & Format(MaxRate!, "0.0#") & " S/s)," & vbCrLf & _
         Format(IdealSamplesPer!, "0.0##") & " ±" & Format(TimeTolerance& + AddTimeTol&, "0") & _
         " samples per cycle." & vbCrLf & "Error = " & Format(RateErr!, "0.0#") & " S/s (" & _
         Format(SamplesPerError!, "0.0##") & " Samples per cycle)."
      End If
   End If
   EvalTime = FailTime% Or FailSlowSource%
   Message = Description$ & Warning$

End Function

Function GetLongChanData(ByVal AllIntegers As Variant, ByVal ChanOfInterest As Integer, _
ByVal NumSamples As Long, ByVal FirstSample As Long, LongChanData As Variant, _
ByVal Convert As Integer) As Integer

   If ChanOfInterest < 0 Then
      MsgBox "Invalid channel (" & Format(ChanOfInterest, "0") & ") passed for evaluation.", vbCritical, "Bad Channel"
      Exit Function
   End If
   ReDim LongArray(NumSamples - 1)
   TypeOfData = VarType(AllIntegers)
   
   For l& = FirstSample To FirstSample + (NumSamples - 1)
      Select Case TypeOfData
         Case vbArray Or vbInteger
            If Convert Then
               LongArray(Sample&) = (AllIntegers(ChanOfInterest, l&) Xor &H8000) + 32768
            Else
               LongArray(Sample&) = AllIntegers(ChanOfInterest, l&)
            End If
         Case vbArray Or vbLong
            LongArray(Sample&) = AllIntegers(ChanOfInterest, l&)
         Case vbArray Or vbSingle, vbArray Or vbDouble, vbArray Or vbVariant
            ATude# = AllIntegers(ChanOfInterest, l&)
            If Convert Then
               AmpCounts& = GetHiResCounts(mnResolution, mnRange, ATude#, mvCustomRange)
               LongArray(Sample&) = AmpCounts&
            Else
               LongArray(Sample&) = ATude#
            End If
         Case vbArray Or vbCurrency
            LongArray(Sample&) = AllIntegers(ChanOfInterest, l&) * 10000
      End Select
      Sample& = Sample& + 1
   Next
      
   LongChanData = LongArray()
   GetLongChanData = True

End Function

Private Function ConvertPTrig(BoardNum As Long, PretrigCount As Long, _
TotalCount As Long, DataFromBuf As Variant) As Long

   ULStat& = cbAConvertPretrigData(BoardNum, PretrigCount, TotalCount, ADData%, ChanTags%)
   
End Function

Public Sub SetAmpGB(ByVal AmpGB As Long)

   mlAmplGB = AmpGB
   
End Sub

Public Function GetAmpGB() As Long

   GetAmpGB = mlAmplGB
   
End Function

Public Sub SetRateGB(ByVal RateGB As Long)

   mlRateGB = RateGB
   
End Sub

Public Function GetRateGB() As Long

   GetRateGB = mlRateGB
   
End Function

Public Sub SetMvgAvgGB(ByVal MvgAvg As Long)

   mlMvgAvgGB = MvgAvg
   
End Sub

Public Function GetMvgAvgGB() As Long

   GetMvgAvgGB = mlMvgAvgGB
   
End Function

Function FindATrigger(Direction As Integer, FirstPoint As Long, LastPoint As Long, Threshold As Variant, TrigPoint As Long) As Integer

   'to do - handle gating and hysteresis
   If (Direction = TRIGABOVE) Or (Direction = TRIGPOSEDGE) Then
      For DataPoint& = FirstPoint To LastPoint
         CurValue& = mvLongChanData(DataPoint&)
         If CurValue& > Threshold Then
            AboveThreshold% = True
            Exit For
         End If
      Next DataPoint&
   Else
      For DataPoint& = FirstPoint To LastPoint
         CurValue& = mvLongChanData(DataPoint&)
         If CurValue& < Threshold Then
            BelowThreshold% = True
            Exit For
         End If
      Next DataPoint&
   End If
   
   'If (Direction = TRIGABOVE) Or (Direction = TRIGBELOW) Then
      FindATrigger = BelowThreshold% Or AboveThreshold%
      TrigPoint = DataPoint&
   'Else
      
   'End If
   
End Function

Function GetBoardTweaks(Optional BoardName As String = "") As String
   
   'get board guardbands
   lpFileName$ = "ScriptParams.ini"
   If BoardName = "" Then
      lpApplicationName$ = msBoardName
      SetLocals% = True
   Else
      lpApplicationName$ = BoardName
   End If
   nSize% = 256
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = ""
   
   lpKeyName$ = "AmpGuard"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
   lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   Amplitude$ = Left$(lpReturnedString$, x%)
   StringSize% = StringSize% + x%
   lpKeyName$ = "RateGuard"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
   lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   StringSize% = StringSize% + x%
   RateTweak$ = Left$(lpReturnedString$, x%)
   lpKeyName$ = "MvgAvg"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
   lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   StringSize% = StringSize% + x%
   AvgTweak$ = Left$(lpReturnedString$, x%)
   lpKeyName$ = "SimultaneousOut"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
   lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   StringSize% = StringSize% + x%
   SOutTweak$ = Left$(lpReturnedString$, x%)
   lpKeyName$ = "SimultaneousIn"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
   lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   StringSize% = StringSize% + x%
   SInTweak$ = Left$(lpReturnedString$, x%)
   
   If SetLocals% Then
      mfBoardAmplGB = Val(Amplitude$)
      mlBoardRateGB = Val(RateTweak$)
      mlBoardMvgAvgGB = Val(AvgTweak$)
      mnSimOut = Val(SOutTweak$)
      mnSimIn = Val(SInTweak$)
   Else
      If Not (StringSize% = 0) Then TweakVals$ = "RateGuard=" & _
      RateTweak$ & ",AmpGuard=" & Amplitude$ & ",MvgAvg=" & AvgTweak$ & _
      ",SimOut=" & SOutTweak$ & ",SimIn=" & SInTweak$
   End If
   GetBoardTweaks = TweakVals$
   
End Function

Public Sub ResetEvalBoard()

   'resets board name so the board tweaks will be reloaded
   msBoardName = ""
   
End Sub

Function GetErrorParams(FormRef As Form, ErrorCode As Long) As String

   Select Case ErrorCode
      Case BADRANGE
         Range% = GetCurrentRange(FormRef)
         Chans% = GetNumberOfChannels(FormRef)
         If mlQueueCount > 0 Then
            RangeVal% = mnGainQueue(1, mnChanIndex)
            Channel% = mnGainQueue(0, mnChanIndex)
            QueueChan$ = " on channel " & Format(Channel%, "0")
         Else
            RangeVal% = mnRange
         End If
         RangeTested$ = GetRangeString(RangeVal%)
         ErrParams$ = "Error code BADRANGE returned for " & RangeTested$
   End Select
   GetErrorParams = ErrParams$
   
End Function


Function GetLowChannel(FormRef As Form) As Integer

   'get low channel
   FormRef.cmdConfigure.Caption = "#"
   FormRef.cmdConfigure = True
   DoEvents
   If FormRef.cmdConfigure.Caption = "-1" Then
      'channel queue is being used - get the queue in GetNumberOfChannels function
      mnFirstChan = -1
   Else
      mnFirstChan = Val(FormRef.cmdConfigure.Caption)
   End If
   GetLowChannel = mnFirstChan

End Function

Function GetNumberOfChannels(FormRef As Form) As Integer

   'get number of channels
   FormRef.cmdConfigure.Caption = "?"
   FormRef.cmdConfigure = True
   DoEvents
   InfoReturned$ = FormRef.cmdConfigure.Caption
   QType$ = Left(InfoReturned$, 1)
   Select Case QType$
      Case "I"
         'chans may be set to individual ranges
         'chan count not queued, but range may be
         'set queue size to total # of chans but
         'numchans to the number of chans in scan
         RetrieveQueue% = True
         mlQueueEnabled = False
         If Len(InfoReturned$) > 1 Then
            Channels$ = Mid(InfoReturned$, 2)
            ChanQ = Split(Channels$, "/")
            ChansInScan$ = ChanQ(0)
            If IsNumeric(ChansInScan$) Then mnNumChans = Val(ChansInScan$)
            If UBound(ChanQ) > 0 Then
               ChansInQ$ = ChanQ(1)
               If IsNumeric(ChansInQ$) Then mlQueueCount = Val(ChansInQ$)
            End If
         End If
      Case "Q"
         RetrieveQueue% = True
         RetrieveCount% = True
         mlQueueEnabled = True
      Case Else
         mlQueueCount = 0
         mlQueueEnabled = False
         mnNumChans = Val(FormRef.cmdConfigure.Caption)
   End Select
   
   If RetrieveQueue% Then
      mlQueueCount = FormRef.GetQueueList(QChans, QGains, QTypes)
      If mlQueueCount > 0 Then ReDim mnGainQueue(1, mlQueueCount - 1)
      For QElement& = 0 To mlQueueCount - 1
         mnGainQueue(0, QElement&) = QChans(QElement&)
         mnGainQueue(1, QElement&) = QGains(QElement&)
      Next QElement&
      If RetrieveCount% Then mnNumChans = mlQueueCount
   End If
   GetNumberOfChannels = mnNumChans

End Function

Function GetResolution(FormRef As Form) As Integer
   
   'get resolution of data
   FormRef.cmdConfigure.Caption = "2"
   FormRef.cmdConfigure = True
   GetResolution = Val(FormRef.cmdConfigure.Caption)

End Function

Function GetCurrentRange(FormRef As Form) As Integer

   'get range at which data was collected
   FormRef.cmdConfigure.Caption = "3"
   FormRef.cmdConfigure = True
   RangeCode$ = FormRef.cmdConfigure.Caption
   If Left(RangeCode$, 1) = "C" Then
      'non-standard full scale value
      RangeInf = Split(RangeCode$, ",")
      mnRange = Val(RangeInf(1))
      mvCustomRange = Val(RangeInf(2))
      GetCurrentRange = mnRange
      Exit Function
   ElseIf Left(RangeCode$, 1) = "I" Then
      mnRange = mnGainQueue(1, mnFirstChan + mnChanIndex)
      Exit Function
   Else
      mvCustomRange = Null
   End If
   If RangeCode$ = "Q" Then
      If mlQueueCount > 0 Then
         mnRange = mnGainQueue(1, mnChanIndex)
      Else
         mnRange = Val(RangeCode$)
      End If
   Else
      mnRange = Val(RangeCode$)
   End If
   GetCurrentRange = mnRange

End Function

Function GetRateParams(FormRef As Form, Warning As String) As Integer

   'get the rate requested and the rate returned
   FormRef.cmdConfigure.Caption = "4"
   FormRef.cmdConfigure = True
   mfRateRequested = Val(FormRef.cmdConfigure.Caption)
   FormRef.cmdConfigure.Caption = "5"
   FormRef.cmdConfigure = True
   mfRateReturned = Val(FormRef.cmdConfigure.Caption)
   If Not mfRateRequested = 0 Then
      If Abs(mfRateRequested - mfRateReturned) > (mfRateRequested * 0.01) Then
         If Not mnRateWarning Then
            WarnRate% = True
         End If
      End If
   Else
      If Not mfRateRequested = mfRateReturned Then
         If Not mnRateWarning Then
            WarnRate% = True
         End If
      End If
   End If
   If WarnRate% Then
      Warning = vbCrLf & vbCrLf & "Warning - Possible problem with rate returned." & _
      vbCrLf & "Rate requested was " & Format(mfRateRequested, "0.0###") & _
      " but rate returned was " & Format(mfRateReturned, "0.0###") & "."
      mnRateWarning = True
   End If

End Function

Function ConvertLSBs(LSBs As Long) As Single

   'scripted LSB tolerance is in terms of 12-bit LSBs
   'here, it is converted to a voltage value
   ATol! = 2 ^ (mnResolution - 12) * LSBs
   If ATol! = 0 Then ATol! = 0.5
   ConvertLSBs = ATol!

End Function

Function ReadStatus(FormRef As Form, StatVal%, CurIndex As Long, _
CurCount As Long) As Long

   FormRef.cmdConfigure.Caption = "Q"
   FormRef.cmdConfigure = True
   DoEvents
   Status$ = FormRef.cmdConfigure.Caption
   StatArgs = Split(Status$, ",")
   NumArgs& = UBound(StatArgs)
   If NumArgs& = 3 Then
      StatError& = StatArgs(0)
      StatVal% = StatArgs(1)
      CurCount = StatArgs(2)
      CurIndex = StatArgs(3)
   End If
   ReadStatus = StatError&
   
End Function

Public Sub SetEvent(ByVal EventType As Long, ByVal EventData As Long, ByVal Abort As Integer)

   mlEventType = EventType
   mlEventData = EventData
   mnTimeout = Abort
   
End Sub

Private Function SetVarToType(ByVal Argument As String, ArgType As Integer, Convert As Integer) As Variant
   
   UnitsStripped$ = ParseUnits(Argument, UnitType%)
   ArgType = UnitType%
   Select Case UnitType%
      Case UNITPERCENT  '1
         FSR& = 2 ^ mnResolution
         SetVarToType = CLng(FSR& * (Val(UnitsStripped$) / 100))
      Case UNITVOLTS '2
         Convert = False
         SetVarToType = Val(UnitsStripped$)
      Case UNITDEGREES '3
         Convert = False
         SetVarToType = Val(UnitsStripped$)
      Case Else
         If (Not InStr(1, UnitsStripped$, ".") = 0) Then
            SetVarToType = CSng(UnitsStripped$)
         Else
            Amp = Val(UnitsStripped$)
            SetVarToType = CLng(Amp)
         End If
   End Select

End Function

Private Function ConvertStringToType(ValueString As String, ValueType As Integer) As Variant

   Select Case ValueType
      Case UNITPERCENT  '1
         FSR& = 2 ^ mnResolution
         ConvertedValue = CLng(FSR& * (Val(ValueString) / 100))
      Case UNITVOLTS    '2
         ConvertedValue = Val(ValueString)
      Case UNITDEGREES  '3
         ConvertToCounts% = False
         ConvertedValue = Val(ValueString)
      Case UNITCOUNTS '4
         ConvertToCounts% = False
         ConvertedValue = Val(ValueString)
      Case UNITFLOAT '5
         ConvertedValue = CSng(ValueString)
      Case Else
         Amp = Val(ValueString)
         ConvertedValue = CLng(Amp)
   End Select
   ConvertStringToType = ConvertedValue

End Function

Public Function GetSineMaxDelta(Args As Variant) As Single

   Dim FormRef As Form
   msOptions = ""
   NumArgs% = UBound(Args)
   If NumArgs% > 1 Then
      EvalType% = Val(Args(1))
      FormID$ = Args(0)
      If FormID$ = "0" Then
         FormFound% = True
         MainForm% = True
      Else
         FormFound% = GetFormReference(FormID$, FormRef)
      End If
      SigAmp! = Val(Trim(Args(3)))
      SourceRate! = Val(Trim(Args(4)))
      SamplesPerChan% = Val(Trim(Args(5)))
      ExtClockRate! = Val(Args(7))
   End If
   
   If Not FormFound% Then Exit Function
   
   If ExtClockRate! = 0 Then
      FormRef.cmdConfigure.Caption = "5"
      FormRef.cmdConfigure = True
      mfRateReturned = Val(FormRef.cmdConfigure.Caption)
   End If
   
   mnResolution = GetResolution(FormRef)
   NumChans% = 1
   If Not (SamplesPerChan% = 0) Then
      Chans% = GetNumberOfChannels(FormRef)
      NumChans% = Chans%
   End If
   SamplesPerCycle& = ((mfRateReturned / SourceRate!) / NumChans%) '/ 4
   If SamplesPerCycle& = 0 Then SamplesPerCycle& = 1
   DeltaRads! = 6.283185 / SamplesPerCycle&
   MaxDeltaV! = (Sin(DeltaRads!)) * SigAmp!
   intZeroScale% = 1
   If mnRange < 100 Then
      ZeroScale& = (2 ^ mnResolution / 2) + 1
      intZeroScale% = ULongValToInt(ZeroScale&)
   End If
   MinV! = GetVolts(mnResolution, mnRange, intZeroScale%)
   If MaxDeltaV! < MinV! Then MaxDeltaV! = MinV!
   'mnResolution
   GetSineMaxDelta = MaxDeltaV! * 4

End Function

