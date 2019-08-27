Attribute VB_Name = "EvalData"
Const FROMARRAY = 0
Const BUFFER = 1
Const FILE = 2
Const XMEM = 3
Const RealArray = 4
Const PARTBUFFER = 5
Const MEM_PRETRIG = 6
Const PRETRIGBUFFER = &H11

Dim mnAnalyzeV As Integer, mnAnalyzeT As Integer
Dim mlEndBuffer As Long, mlBlockSize As Long
Dim mlArraySize As Long, mlStartBlock As Long
Dim mnTransferType As Integer, mvMemHandle As Variant

Dim mnConvertData As Integer, mnConvToEng As Integer
Dim mHistBins() As Long, mDatArray() As Integer, mlLongArray() As Long
Dim mvPlotArray() As Variant, mPointArray() As Integer
Dim mnStopChan As Integer, mnStartChan As Integer
Dim mnLastChan As Integer, mnAllChans As Integer

Dim mfNoForm As Form
Dim mnErrorDetected As Integer

Dim malTransitionCount() As Long

Dim mlStopDVMax As Long, mlStopDTMax As Long
Dim mlStopVMin As Long, mlStopVMax As Long
Dim mfStopVMin As Single, mfStopVMax As Single
Dim mlStopWinLow As Long, mlStopWinHigh As Long

Dim mnShowMinMax As Integer
Dim mnFreqShow As Integer
Dim mfRate As Single, mnCycleUnits As Integer

Dim mlTriggerValue As Long, mnByPeak As Integer
Dim msCountTrig As Single, mlCountThen As Long, msDeltaCount As Single
Dim msTotalCount As Single, msTempCount As Single
Dim mlTMin As Long, mlTMax As Long

Function EvalBuffer(BufHandle As Long, BufCount As Long, Chans As Integer) As Integer

   frmPlot.picDetails.Line (30, 0)-(80, 3.6), &HFFFFFF, BF
   mnErrorDetected = False
   mlEndBuffer = BufCount
   mnLastChan = Chans
   mlBlockSize = (mlEndBuffer) \ (mnLastChan + 1)
   'mlBlockSize = mlEndBuffer
   If mnAllChans Then mnStopChan = mnLastChan
   
   mlArraySize = mlEndBuffer - 1
   mnTransferType = BUFFER
   mvMemHandle = BufHandle
   XferBlock
   EvalBuffer = mnErrorDetected

End Function

Function EvalDeltaV(Min&, Max&, MinBin&, MaxBin&, DeltaExceedBin&, MaxDeltaBin&) As Long

   Min& = 65536
   Max& = 0
   For Chan% = mnStartChan To mnStopChan
      For Element& = 1 To mlEndBuffer - 1 'mlBlockSize - 1
         If mvPlotArray(Chan%, Element& - 1) < Min& Then
            Min& = mvPlotArray(Chan%, Element& - 1)
            MinBin& = ((Element& - 1) * (UBound(mvPlotArray) + 1)) + Chan% 'Element& - 1
            If Not mnShowMinMax Then
               If (Not (mlStopVMin < 0)) And (Min& < mlStopVMin) Then
                  minExceed& = Min&
                  If Not mnShowMinMax Then Exit For
               End If
            End If
         End If
         If mvPlotArray(Chan%, Element& - 1) > Max& Then
            Max& = mvPlotArray(Chan%, Element& - 1)
            MaxBin& = (Element& * (UBound(mvPlotArray) + 1)) + Chan%
            If Not mnShowMinMax Then
               If (Not (mlStopVMax < 0)) And (Max& > mlStopVMax) Then Exit For
            End If
         End If
         Delta& = mvPlotArray(Chan%, Element&) - mvPlotArray(Chan%, Element& - 1)
         
         'following line used for verification of 4020 counter in test mode
         'If Delta& = -255 Then Delta& = 1
         
         If (mlStopDVMax > 0) And (Abs(Delta&) > mlStopDVMax) And (DeltaExceedBin& = 0) Then
            DeltaExceedBin& = (Element& * (UBound(mvPlotArray) + 1)) + Chan%
         End If
         If Abs(Delta&) > MaxDelta& Then
            MaxDelta& = Abs(Delta&)
            MaxDeltaBin& = (Element& * (UBound(mvPlotArray) + 1)) + Chan%
         End If
      Next Element&
   Next Chan%
   EvalDeltaV = MaxDelta&
         
End Function

Function EvalIntArray(DatArray() As Integer, TotalCount&, FirstPoint&)
   
   mnErrorDetected = False
   Chans% = UBound(DatArray, 1)
   If ((mnStartChan > Chans%) Or (mnStopChan > Chans%)) Then
      MsgBox "The data evaluator is set up for a channel that isn't included in the scan.", , "Invalid Evaluation Channel"
      mnErrorDetected = 999
      gnCancel = True
      Exit Function
   End If

   mnLastChan = Chans%
   If TotalCount& = 0 Then
      'If mnTransferType = FROMARRAY Then
        mlArraySize = UBound(DatArray, 2)
      mlEndBuffer = mlArraySize \ (mnLastChan + 1)
      mlStartBlock = 0
      'InitBlock
   Else
      mlArraySize = TotalCount&
      mlEndBuffer = mlArraySize \ (mnLastChan + 1)
      mlStartBlock = FirstPoint& \ (mnLastChan + 1)
   End If
   
   ReDim mvPlotArray(mnLastChan, mlArraySize)
   mnDataInit = True
   x& = IntArrayToULong(DatArray(), mvPlotArray(), mlEndBuffer - 1, mlStartBlock)
   
   If mnAnalyzeV Then
      Delta& = EvalDeltaV(Min&, Max&, MinBin&, MaxBin&, DeltaExceedBin&, MaxDeltaBin&)
      If mlStopDVMax > 0 Then TrapDeltaV Delta&, DeltaExceedBin&
      If Not ((mlStopVMin = -1) And (mlStopVMax = -1)) Then TrapVMinMax Min&, Max&, MinBin&, MaxBin&
      If mnShowMinMax Then ShowMinMax Min&, Max&, MinBin&, MaxBin&
   End If
   If mnAnalyzeT Then
      lngPeriod& = EvalPeriod(Chan%)
      If Not (lngPeriod& = 0) Then
         If mlStopDTMax > 0 Then TrapDeltaT lngPeriod&
         If mnFreqShow Then ShowCycles lngPeriod&
      End If
   End If

End Function

Function EvalPartialBuf(BufHandle As Long, BufCount As Long, BufStart As Long, Chans As Integer) As Integer

   mnErrorDetected = False
   mlEndBuffer = BufCount
   'mlBlockSize = mlEndBuffer
   'InitPlot Chans
   mnLastChan = Chans
   'InitBlock
   mlBlockSize = mlEndBuffer \ (mnLastChan + 1)
   mlArraySize = mlBlockSize
   mlStartBlock = BufStart
   
   mnTransferType = PARTBUFFER
   mvMemHandle = BufHandle
   XferBlock
   EvalPartialBuf = mnErrorDetected
   'ReturnToOwner

End Function

Function EvalPeriod(Chan%) As Long

   Incrementing% = True
   mnGaurdBand = 3
   TriggerValue& = mlTriggerValue
   Do
      If Incrementing% Then
         If mnByPeak Then TriggerValue& = mvPlotArray(Chan%, CurPoint&)
         If Not (TriggerValue& < mvPlotArray(Chan%, CurPoint& + 1)) Then
            'check until all points within guardband are less than the trigger point
            TestPoint& = CurPoint& - 1
            Do
               Test% = 0
               TestPoint& = TestPoint& + 1
               For i% = 1 To mnGaurdBand
                  If mlArraySize < (i% + TestPoint&) Then Exit Do
                  If mnByPeak Then TriggerValue& = mvPlotArray(Chan%, TestPoint&)
                  If Not (mvPlotArray(Chan%, TestPoint&) < mvPlotArray(Chan%, TestPoint& + i%)) Then Test% = Test% + 1
               Next i%
            Loop While Test% < mnGaurdBand
            Incrementing% = False
            Decrementing% = True
            CurPoint& = TestPoint&
            If CurPoint& > 0 Then
               ReDim Preserve malTransitionCount(Transition%)
               malTransitionCount(Transition%) = CurPoint&
               Transition% = Transition% + 1
            End If
         End If
      End If
      If Decrementing% Then
         If mnByPeak Then TriggerValue& = mvPlotArray(Chan%, CurPoint&)
         If Not (TriggerValue& > mvPlotArray(Chan%, CurPoint& + 1)) Then
            'check until all points within guardband are less than the trigger point
            TestPoint& = CurPoint& - 1
            Do
               Test% = 0
               TestPoint& = TestPoint& + 1
               For i% = 1 To mnGaurdBand
                  If mlArraySize < (i% + TestPoint&) Then Exit Do
                  If mnByPeak Then TriggerValue& = mvPlotArray(Chan%, TestPoint&)
                  If Not (mvPlotArray(Chan%, TestPoint&) > mvPlotArray(Chan%, TestPoint& + i%)) Then Test% = Test% + 1
               Next i%
            Loop While Test% < mnGaurdBand
            Incrementing% = True
            Decrementing% = False
            CurPoint& = TestPoint&
            If CurPoint& > 0 Then
               ReDim Preserve malTransitionCount(Transition%)
               malTransitionCount(Transition%) = CurPoint&
               Transition% = Transition% + 1
            End If
         End If
      End If
      CurPoint& = CurPoint& + 1
   Loop While CurPoint& < mlArraySize
   Diff& = 0
   LoopCount% = 0
   If Transition% > 5 Then
      For x% = 1 To Transition% - 5 Step 2
         Diff& = Diff& + (malTransitionCount(x% + 2) - malTransitionCount(x%))
         LoopCount% = LoopCount% + 1
      Next x%
      EvalPeriod = Diff& / LoopCount%
   End If

End Function

Sub SetConvertEval(BoardNum As Integer, CONVERTDATA As Integer)

   mnConvertData = CONVERTDATA

End Sub

Sub SetCycleUnits(Units%)

   mnCycleUnits = Units%

End Sub

Sub SetDeltaStop(DeltaVal&)

   mlStopDVMax = DeltaVal&
   mnAnalyzeV = Not (mlStopDVMax = 0)

End Sub

Sub SetValueWindow(LowVal As Long, HighVal As Long)

   mlStopWinLow = LowVal
   mlStopWinHigh = HighVal
   
End Sub

Sub SetEvalChan(Chan%)

   If Chan% < 0 Then
      mnStartChan = 0
      mnAllChans = True
   Else
      mnStartChan = Chan%
      mnStopChan = Chan%
      mnAllChans = False
   End If

End Sub

Sub SetPeakTrigType(ByPeak%)

   mnByPeak = ByPeak%

End Sub

Sub SetRate(RateVal As Single)

   mfRate = RateVal

End Sub

Sub SetShowCycles(ShowVal%)

   mnFreqShow = ShowVal%
   If ShowVal% Then mnAnalyzeT = True
   ShowCycles 0

End Sub

Sub SetShowMinMax(ShowVal%)

   mnShowMinMax = ShowVal%
   If ShowVal% Then mnAnalyzeMinMax = True
   ShowMinMax -1, -1, 0, 0

End Sub

Sub SetTrigValue(TrigVal&)

   mlTriggerValue = TrigVal&

End Sub

Sub SetCountTrigValue(TrigVal!)

   msCountTrig = TrigVal!

End Sub

Sub SetCountDelta(DeltaVal!)

   msDeltaCount = DeltaVal!

End Sub

Sub SetCountThenValue(ThenVal&)

   mlCountThen = ThenVal&

End Sub

Sub SetVMinMaxStop(MinStop&, MaxStop&)

   mlStopVMin = MinStop&
   mlStopVMax = MaxStop&
   If Not ((mlStopVMin = -1) And (mlStopVMax = -1)) Then mnAnalyzeV = True

End Sub

Sub SetVfMinMaxStop(MinStop!, MaxStop!)

   mfStopVMin = MinStop!
   mfStopVMax = MaxStop!
   If Not (mfStopVMin = -9999) Then mnAnalyzeV = True

End Sub

Sub SetDeltaT(MinT&, MaxT&)

   mlTMin = MinT&
   mlTMax = MaxT&
   mlStopDTMax = MaxT&
   
End Sub

Sub ClearDetails()

   frmPlot.picDetails.Line (32, 0)-(80, 3.6), &HFFFFFF, BF

End Sub

Sub ShowCycles(PERIOD&)

   frmPlot.picDetails.Line (48, 0.7)-(64, 1.7), &HFFFFFF, BF
   If PERIOD& = 0 Then Exit Sub
   frmPlot.picDetails.CurrentX = 48
   frmPlot.picDetails.CurrentY = 0.8
   Select Case mnCycleUnits
      Case 0   'samples
         frmPlot.picDetails.Print "Samples / cycle: " & PERIOD&
      Case 1   'period
         CycleTime! = PERIOD& * (1 / mfRate) * 1000
         frmPlot.picDetails.Print "Period: " & Format$(CycleTime!, "0.000") & " ms"
      Case 2   'frequency
         Freq! = 1 / (PERIOD& * (1 / mfRate)) / 1000
         frmPlot.picDetails.Print "Frequency: " & Format$(Freq!, "0.000") & " kHz"
   End Select

End Sub

Sub ShowFrequency(PERIOD&)

   frmPlot.picDetails.Line (48, 0.8)-(64, 1.6), &HFFFFFF, BF
   If PERIOD& = 0 Then Exit Sub
   Freq! = 1 / (PERIOD& * (1 / mfRate)) / 1000
   frmPlot.picDetails.CurrentX = 48
   frmPlot.picDetails.CurrentY = 0.8
   frmPlot.picDetails.Print "Frequency: " & Format$(Freq!, "0.000") & " kHz"

End Sub

Sub ShowMinMax(Min&, Max&, MinBin&, MaxBin&)

   CurState% = frmEvalData.Visible
   frmEvalData.Visible = False
   frmPlot.picDetails.Line (30, 1.8)-(47, 3.6), &HFFFFFF, BF
   DrawLine MinBin&
   DrawLine2 MaxBin&
   If Min& = -1 And Max& = -1 Then Exit Sub
   frmPlot.picDetails.CurrentX = 30
   frmPlot.picDetails.CurrentY = 1.8
   frmPlot.picDetails.Print "Min Value: " & Min& & " @ " & MinBin&
   frmPlot.picDetails.CurrentX = 30
   frmPlot.picDetails.CurrentY = 2.8
   frmPlot.picDetails.Print "Max Value: " & Max& & " @ " & MaxBin&
   frmEvalData.Visible = CurState%

End Sub

Sub TrapDeltaT(Delta&)

   If Delta& > mlTMax Then
      frmPlot.picDetails.Line (48, 0)-(64, 0.8), &HFFFFFF, BF
      frmPlot.picDetails.CurrentX = 48
      frmPlot.picDetails.CurrentY = 0
      frmPlot.picDetails.Print "Delta T exceeded: " & Format$(mlTMax, "0")
      'DrawLine Element&
      gnCancel = True
      mnErrorDetected = True
   End If

End Sub

Sub TrapDeltaV(Delta&, Element&)

   frmPlot.picDetails.Line (32, 1)-(48, 1.8), &HFFFFFF, BF
   frmPlot.picDetails.CurrentX = 32
   frmPlot.picDetails.CurrentY = 1.8
   If Delta& > mlStopDVMax Then
      DrawLine Element&
      frmPlot.picDetails.CurrentX = 32
      frmPlot.picDetails.CurrentY = 1
      frmPlot.picDetails.Print "Delta Exceeded: " & Delta&
      gnCancel = True
      mnErrorDetected = True
   End If

End Sub

Sub TrapVMinMax(VMin&, VMax&, MinBin&, MaxBin&)
   
   If VMax& > mlStopVMax Then
      DrawLine MaxBin&
      frmPlot.picDetails.CurrentX = 48
      frmPlot.picDetails.CurrentY = 2.8
      frmPlot.picDetails.Print "Max Value Exceeded: " & VMax& & " @ " & MaxBin&
      gnCancel = True
      mnErrorDetected = True
   End If
   If VMin& < mlStopVMin Then
      DrawLine MinBin&
      frmPlot.picDetails.CurrentX = 48
      frmPlot.picDetails.CurrentY = 1.7
      frmPlot.picDetails.Print "Min Value Exceeded: " & VMin& & " @ " & MinBin&
      gnCancel = True
      mnErrorDetected = True
   End If

End Sub

Sub XferBlock()
   
   If mlArraySize < 0 Then mlArraySize = 0
   ReDim DatArray(mnLastChan, mlArraySize) As Integer
   If mnTransferType = PARTBUFFER Then
      FirstPoint& = mlStartBlock
   Else
      FirstPoint& = mlStartBlock * (mnLastChan + 1)
   End If
   NumPoints& = (mlBlockSize * (mnLastChan + 1))
   
   Select Case mnTransferType
      Case BUFFER, PARTBUFFER
         ULStat = cbWinBufToArray(mvMemHandle, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, WinBufToArray, ULStat, mvMemHandle, DatArray(0, 0), FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case FILE
         ULStat = cbFileRead(msFileName, FirstPoint&, NumPoints&, DatArray(0, 0))
         If SaveFunc(mfNoForm, FileRead, ULStat, msFileName, FirstPoint&, NumPoints&, DatArray(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case XMEM
         ULStat = cbMemRead(mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, MemRead, ULStat, mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case MEM_PRETRIG
         ULStat = cbMemReadPretrig(mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, MemReadPretrig, ULStat, mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case Else
         Exit Sub
   End Select
   If mnConvertData Then
      ReDim ChanTags(mnLastChan, mlBlockSize) As Integer
      ULStat = cbAConvertData(mnBoardNum, NumPoints&, DatArray(0, 0), ChanTags(0, 0))
      If SaveFunc(mfNoForm, AConvertData, ULStat, mnBoardNum, NumPoints&, DatArray(0, 0), ChanTags(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
   If mnConvToEng Then
      ReDim ReelArray!(mnLastChan, mlBlockSize)
      For Chan% = 0 To mnLastChan
         For Sample% = 0 To mlArraySize
            ULStat = cbToEngUnits(mnBoardNum, mnRange, DatArray(Chan%, Sample%), EngUnits!)
            If SaveFunc(mfNoForm, ToEngUnits, ULStat, mnBoardNum, mnRange, DatArray(Chan%, Sample%), EngUnits!, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            ReelArray!(Chan%, Sample%) = EngUnits!
         Next Sample%
      Next Chan%
      x% = EvalRealArray(ReelArray!(), 0, 0)
   Else
      x% = EvalIntArray(DatArray(), 0, 0)
   End If

End Sub

Function EvalRealArray(FltArray() As Single, TotalCount&, FirstPoint&) As Integer
   
   mnErrorDetected = False
   Chans% = UBound(FltArray, 1)
   mnStopChan = Chans%
   'If ((mnStartChan > Chans%) Or (mnStopChan > Chans%)) Then
   '   MsgBox "The data evaluator is set up for a channel that isn't included in the scan.", , "Invalid Evaluation Channel"
   '   mnErrorDetected = 999
   '   gnCancel = True
   '   Exit Function
   'End If

   mnLastChan = Chans%
   If TotalCount& = 0 Then
      mlArraySize = UBound(FltArray, 2)
      mlEndBuffer = mlArraySize \ (mnLastChan + 1)
      mlStartBlock = 0
   Else
      mlArraySize = TotalCount&
      mlEndBuffer = mlArraySize \ (mnLastChan + 1)
      mlStartBlock = FirstPoint& \ (mnLastChan + 1)
   End If
   
   ReDim mvPlotArray(mnLastChan, mlArraySize)
   For Ch% = 0 To mnLastChan
      For Samp& = 0 To mlArraySize
         mvPlotArray(Ch%, Samp&) = FltArray(Ch%, Samp&)
      Next
   Next
   mnDataInit = True
   
   If mnAnalyzeV Then
      Min! = mfStopVMin
      Max! = mfStopVMax
      result% = EvalFloatV(Min!, Max!)
   End If

End Function

Function EvalFloatV(Min!, Max!) As Integer

   For Chan% = mnStartChan To mnStopChan
      For Element& = 0 To mlEndBuffer
         If mvPlotArray(Chan%, Element&) < Min! Then
            Min! = mvPlotArray(Chan%, Element&)
            gnCancel = True
            mnErrorDetected = True
            If Not (mfStopVMin < 0) And (Min! < mfStopVMin) Then Exit For
         End If
         If mvPlotArray(Chan%, Element&) > Max! Then
            Max! = mvPlotArray(Chan%, Element&)
            gnCancel = True
            mnErrorDetected = True
            If Not (mfStopVMin < 0) And (Max! > mfStopVMax) Then Exit For
         End If
         
         If (mlStopDVMax > 0) And (Abs(Delta&) > mlStopDVMax) And (DeltaExceedBin& = 0) Then
            DeltaExceedBin& = (Element& * (UBound(mvPlotArray) + 1)) + Chan%
         End If
         If Abs(Delta&) > MaxDelta& Then
            MaxDelta& = Abs(Delta&)
            MaxDeltaBin& = (Element& * (UBound(mvPlotArray) + 1)) + Chan%
         End If
      Next Element&
   Next Chan%

End Function

Function EvalCount(CBCount&) As Integer

   Static SecondHalf%
   If msDeltaCount > 0 Then
      If SecondHalf% = True Then
         If Abs(CBCount& - msTempCount) < msDeltaCount Then
            frmPlot.picDetails.Line (32, 1)-(48, 1.8), &HFFFFFF, BF
            frmPlot.picDetails.CurrentX = 32
            frmPlot.picDetails.CurrentY = 1.8
            frmPlot.picDetails.CurrentX = 32
            frmPlot.picDetails.CurrentY = 1
            frmPlot.picDetails.Print "Count change < : " & Format(msDeltaCount, "0") & _
            " (" & Format(Abs(CBCount& - msTempCount), "0") & ")"
            gnCancel = True
            mnErrorDetected = True
            SecondHalf% = False
         Else
            frmPlot.picDetails.Line (32, 1)-(48, 1.8), &HFFFFFF, BF
            msTempCount = CBCount&
         End If
      Else
         SecondHalf% = True
         msTempCount = CBCount&
      End If
      Exit Function
   End If
   ConvCount! = CBCount&
   If CBCount& < 0 Then
      ConvCount! = 4294967296# + CBCount&
      SecondHalf% = True
   Else
      If SecondHalf% Then msTempCount = msTempCount + 4294967296#
      SecondHalf% = False
   End If
   SngCount! = msTempCount + ConvCount!
   If SngCount! > msCountTrig Then
      TrapOnCount SngCount!, 1
   End If

End Function

Sub TrapOnCount(SngCount!, Iteration As Integer)

   frmPlot.picDetails.Line (32, 1)-(48, 1.8), &HFFFFFF, BF
   frmPlot.picDetails.CurrentX = 32
   frmPlot.picDetails.CurrentY = 1.8
   frmPlot.picDetails.CurrentX = 32
   frmPlot.picDetails.CurrentY = 1
   frmPlot.picDetails.Print "Count Exceeded: " & _
      Format(SngCount!, "0") & " (X" & _
      Format(Iteration, "0") & ")"
   gnCancel = True
   mnErrorDetected = True

End Sub
