Attribute VB_Name = "PlotInterface"

'Description
'  interface between applications that need to plot
'  data and the plotting form (plot.frm)

'Public interfaces:
'  Plot (Array() As Integer)
'     pass a two dimensional integer array, the first
'     dimension containing the number of channels (zero
'     based), the second dimension containing data
'  PlotBuffer (BufHandle As Long, BufCount As Long, Chans As Integer)
'     pass the handle to an existing buffer containing data
'     to be plotted, the number of data points to be plotted
'     and the number of channels (zero based)
'  SetPlotType (PlotType As Integer)
'     sets plot for VvsT, Histogram or Text (use global
'     constants defined below)
'  SetPlotResolution (ResVal As Integer)
'     sets plot resolution to handle data at 12 or 16 bit
'     resolution (ResVal = 16 for 16 bit, anything else = 12)
'  SetAutoScale ()
'     cycles through three modes of scaling ... plot starts
'     at Fixed Full Scale .. calling this routine changes
'     scaling to Autoscale, then to Fixed Current Scale, then
'     back again to Fixed Full Scale
'  FastPlot (Value As Integer)
'     sets plot mode to pixel when Value is non-zero
'     pixel mode is faster but shows less horizontal resolution
'     so spikes in data may not be shown on plot
'  PlotTCValues (mnBoardNum, manChanArray(0), manTypeArray(0), mlQCount, mvHandle, FirstPoint&, NumPoints&, TempScale&)
'     used with new cbDaqInScan function to convert Thermocouple raw data into temperature

'Dependencies
'     cbw.bas (uses cbWinBufToArray())
'     plot.frm
'     math.bas
'     plot16.bas (if using VB3) or plot32.bas
'____________________________________________________

'values for mnPlotType
#If NETOPS Then
   Dim NetLibrary As ULNet.ULNetClass
#Else
   Dim NetLibrary As Object
#End If

Global Const VOLTS_VS_TIME = 0
Global Const HISTOGRAM = 1
Global Const PRINT_TEXT = 2
Global Const SINGLE_POINT = 3
Global Const PLOT_CHUNKS = 4
Global Const PRINT_LIST = 5
Global Const DERIVATIVE = 6
Global Const TO_PRINTER = &H10

Global Const geENABLEEVAL As Long = 1
Global Const geEVALDELTA As Long = 2
Global Const geEVALRANGE As Long = 4
Global Const geEVALMINDELTA As Long = 8
Global Const geEVALINWINDOW As Long = 16
Global Const geSTART As Long = 32
Global Const geNUMPOINTS As Long = 64
Global Const geDELTAVAL As Long = 128
Global Const geEVALCHAN As Long = 256
Global Const geMINVAL As Long = 512
Global Const geMAXVAL As Long = 1024
Global Const geDELTAMIN As Long = 2048
Global Const geINWINDOW As Long = 4096
Global Const geSHOWPASTEMSG As Long = 65536
Global Const geNUMMSGSAMPS As Long = 131072
Global Const geWRITEFILEERRS As Long = 262144

Const FROMARRAY = 0
Const BUFFER = 1
Const FILE = 2
Const XMEM = 3
Const RealArray = 4
Const PARTBUFFER = 5
Const MEM_PRETRIG = 6
Const PRETRIGBUFFER = &H11
Const BUFFER32 = 7
Const PARTBUFFER32 = 8
Const FROMCALLBACK = 9
Const FROMMESSAGE = 10

Global gfOwner As Form, glaPlotColor(7) As Long
Global gnRunDefault As Integer, gnCancel As Integer
Global gnErrorBoard As Integer

Dim mfOwner As Form, mnPlotTitle As Integer
Dim mvMemHandle As Variant, mnHardCopy As Integer
Dim mnRetainPlot As Integer, mnDefaultWidth As Integer
Dim mnPlotFormsLoaded As Integer, mnDataInit As Integer

Dim mlArraySize As Long, mn16Bit As Integer, mnAutoScale As Integer
Dim mnResolution As Integer, mnScaleChange As Integer
Dim mbManualSet As Boolean

Dim mHistBins() As Long, mDatArray() As Integer, mlLongArray() As Long
Dim mPointArray() As Integer  'mvPlotArray() As Variant,
Dim mvPlot As Variant

Dim mnNumBins As Integer, mnFirstBin As Integer, mvLargestBin As Variant
Dim mlNumBins As Long, mvFirstBin As Variant

Dim mnFirstChan As Integer, mnLastChan As Integer, mnChanSet As Integer
Dim mfHRes As Single, mnPlotType As Integer, mlPerChan As Long
Dim mnTransferType As Integer, msMore As String
Dim msListString As String, mlPrinterOS As Long

Dim mlStartBlock As Long, mlEndBlock As Long
Dim mlBlockSize As Long, mlEndBuffer As Long
Dim mlFirstPoint As Long
Dim mlStaticBlockSize As Long ', mlDefaultBlockSize As Long
Dim mlHoldBlockSize As Long, mnResidual As Integer
Dim mnBlocksEnabled As Integer, mnPointsPlotted() As Integer
Dim mnChansLogged() As Integer, mvDivisor As Variant

Dim msFileName As String, msPreamble As String

Dim mlRefLine As Long, mlRefLine2 As Long
Dim mlPreTrigCount As Long, mlRate As Long
Dim mnBoardNum As Integer, mnRange As Integer, mnMemBoard As Integer
Dim mnConvertData As Integer, mnConvToEng As Integer
Dim mnConvBufToEng As Integer
Dim mnCalData As Integer

Dim msDetail As String, msRange As String
Dim mfNoForm As Form

Dim mnUseText As Integer
Dim mnCalcRMS As Integer
Dim mnPlotFixedBlock As Integer
Dim mlSettleStep, mlSettleRate As Long
Dim mnUse32 As Integer, mnUse64 As Integer
Dim mnContinuous As Integer, mnFloat As Integer
Dim mlNumChunks As Long, mnScaled As Integer
Dim mvLsbRms As Variant, mvAverage As Variant
Dim sHexArray() As String

Dim mlBlockSave As Long

Dim manChans() As Integer, mnChansInitialized As Integer
Dim mnLibType As Integer, mnChansSelected As Integer
Dim mlStartText As Long, mlTextCount As Long, mlEvalChan As Long
Dim mbEval As Boolean, mbDeltaEval As Boolean, mbRangeEval As Boolean
Dim mbDeltaMinEval As Boolean, mbDeltaMaxEval As Boolean
Dim mbInWindow As Boolean, mbOutWindow As Boolean
Dim mbFileWrite As Boolean, mbMsgBoxEval As Boolean
Dim mbPairCheck As Boolean
Dim mlEvalStart As Long, mlEvalPoints As Long, mlEvalDeltaVal As Long
Dim mlEvalMinVal As Long, mlEvalMaxVal As Long, mlEvalDeltaMinVal As Long
Dim mlMsgBoxSamps As Long, mlWindowDef As Long
Dim mfTickSize As Single

Sub CalcNoise(TrueFalse As Integer)

   mnCalcRMS = TrueFalse
   RePlot False

End Sub

Sub ClearPlot()

   frmPlot.picPlot.Cls

End Sub

Sub ConvertPTBuffer(BoardNum As Integer, BufHandle As Long, PretrigCount As Long, _
TotalCount As Long, BufferSize As Long, Chans As Integer)

   mvMemHandle = BufHandle
   If mvMemHandle = 0 Then Exit Sub
   mnTransferType = PRETRIGBUFFER
   mlEndBuffer = BufferSize
   'NumChans% = (mnLastChan - mnFirstChan) + 1
   mlPerChan = (BufferSize / (Chans + 1))
   mlArraySize = TotalCount
   mlPreTrigCount = PretrigCount
   mnLastChan = Chans
   
   mnBoardNum = BoardNum
   
   TransferBlock
   
End Sub
   
Sub CreatePTBuffer()

   ReDim DatArray(mnLastChan, mlPerChan) As Integer
   ReDim ChanTags(mnLastChan, mlPerChan) As Integer
   FirstPoint& = 0
   NumPoints& = mlEndBuffer
   'mvMemHandle = BufHandle
   
   ULStat = cbWinBufToArray(mvMemHandle, DatArray(0, 0), FirstPoint&, NumPoints&)
   If SaveFunc(mfNoForm, WinBufToArray, ULStat, mvMemHandle, DatArray(0, 0), FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   
   'NumPoints& = TotalCount
   'DataDimension& = NumPoints& \ (mnLastChan + 1)
   ULStat = cbAConvertPretrigData(mnBoardNum, mlPreTrigCount, mlArraySize, DatArray(0, 0), ChanTags(0, 0))
   If SaveFunc(mfNoForm, AConvertPretrigData, ULStat, mnBoardNum, mlPreTrigCount, mlArraySize, DatArray(0, 0), ChanTags(0, 0), A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

   'following 5 lines added 10/25/02 to allow incremental plotting for converted data
   NumPoints& = mlArraySize
   'mlPerChan = (NumPoints& / (mnLastChan + 1)) - 1
   'ReDim DatArray(mnLastChan, mlPerChan) As Integer
   ULStat = cbWinArrayToBuf(DatArray(0, 0), mvMemHandle, FirstPoint&, NumPoints&)
   If SaveFunc(mfNoForm, WinArrayToBuf, ULStat, DatArray(0, 0), mvMemHandle, FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   mlEndBuffer = NumPoints&
   'mnTransferType = BUFFER
   'TransferBlock
   Exit Sub

   
   ReDim Preserve DatArray(mnLastChan, DataDimension&)
   ReDim Preserve ChanTags(mnLastChan, DataDimension&)
   If mnConvToEng Then
      ReDim ReelArray!(mnLastChan, mlBlockSize + 1)
      For Chan% = 0 To mnLastChan
         For Sample% = 0 To mlBlockSize
            ULStat = cbToEngUnits(mnBoardNum, mnRange, DatArray(Chan%, Sample%), EngUnits!)
            If SaveFunc(mfNoForm, ToEngUnits, ULStat, mnBoardNum, mnRange, DatArray(Chan%, Sample%), EngUnits!, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            ReelArray!(Chan%, Sample%) = EngUnits!
         Next Sample%
      Next Chan%
      x% = PlotRealArray(ReelArray!(), mlEndBuffer)
   Else
      x% = PlotIntArray(DatArray())
   End If

End Sub

Sub ConvertToEng(BoardNum As Integer, ConvToEng As Integer, Optional TickSize As Single)

   If IsMissing(TickSize) Then
      mfTickSize = 0
   Else
      mfTickSize = TickSize
   End If
   mnConvToEng = ConvToEng
   mnBoardNum = BoardNum

End Sub

Sub DrawLine(LinePos As Long)

   mlRefLine = LinePos

End Sub

Sub DrawLine2(LinePos As Long)

   mlRefLine2 = LinePos

End Sub

Sub ExecutePrint()

   Printer.EndDoc
   mlPrinterOS = 0

End Sub

Sub FastPlot(value As Integer)

   'If value Then
   '   frmPlot.picPlot.ScaleMode = 3
   'Else
   '   frmPlot.picPlot.ScaleMode = 1
   'End If

End Sub

Sub FormLoading()

   mnPlotFormsLoaded = mnPlotFormsLoaded + 1

End Sub

Sub FormUnLoading()

   mnPlotFormsLoaded = mnPlotFormsLoaded - 1
   mnLastChan = 0
   mnDataInit = False

End Sub

Function GetBlockSize() As Long

   lplFileName$ = "PlotForm.ini"
   lpApplicationName$ = "PlotParams"
   lpKeyName$ = "NumPoints"
   lpDefault$ = "1000"
   nSize% = 9
   lpReturnedString$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lplFileName$)
   If StringSize% > 0 Then sDefault$ = Left$(lpReturnedString$, StringSize%)
   BlockSize& = Val(sDefault$)
   mlStaticBlockSize = BlockSize&
   If mlBlockSize = 0 Then
      mlBlockSize = BlockSize&
   Else
      BlockSize& = mlBlockSize
   End If
   GetBlockSize = BlockSize&

End Function

Function GetFirstPoint() As Long

   GetFirstPoint = mlFirstPoint  'mlStartBlock
   
End Function

Function GetBufHandle() As Long

   GetBufHandle = mvMemHandle

End Function

Function GetLastChan() As Integer

   GetLastChan = mnLastChan

End Function

Function GetSampleDiff(mRefPoint As Integer, PlotPos As Integer, SampleVal As Single) As Long

   'for use only by 'Plot.frm'
   If mnPlotType = HISTOGRAM Then
      If mn16Bit Then
         If (PlotPos < 1) Or (PlotPos > mlNumBins + 1) Then
            SampleVal = 0
         Else
            SampleVal = mHistBins(PlotPos - 1) - mHistBins(mRefPoint - 1)
         End If
         GetSampleDiff = mvFirstBin + PlotPos - 1
      Else
         If (PlotPos < 1) Or (PlotPos > mnNumBins + 1) Then
            SampleVal = 0
         Else
            SampleVal = mHistBins(PlotPos - 1) - mHistBins(mRefPoint - 1)
         End If
         GetSampleDiff = PlotPos - mRefPoint
      End If
   Else
      Samp& = Int(mfHRes * PlotPos)
      Ref% = Int(mfHRes * mRefPoint)
      SampleVal = mvPlot(mnChan, Samp&) - mvPlot(mnChan, Ref%)
      GetSampleDiff = (Samp& - Ref%) * (mnLastChan + 1)
   End If

End Function

Function GetSampleNum(PlotPos As Integer, SampleVal As Variant, Chan As Integer) As Long

   'for use only by 'Plot.frm'
   If Not mnDataInit Then
      GetSampleNum = -1
      Exit Function
   End If
   ReDim mHistBins(0)
   
   NumChans% = (mnLastChan + 1)
   If mnPlotType = HISTOGRAM Then
      mlNumBins = BinData(Chan, mvPlot, FirstBin, _
         mvLargestBin, LargestBinCode, mHistBins(), Divisor%, Average)
      If (PlotPos < 1) Or (PlotPos > mlNumBins + 1) Then
         SampleVal = 0
      Else
         SampleVal = mHistBins(PlotPos - 1)
      End If
      GetSampleNum = mvFirstBin + PlotPos - 1
   Else
      FirstPoint& = mlStartBlock \ NumChans%
      LastSample& = UBound(mvPlot, 2)
      If mlStartBlock = 0 Then FirstPoint& = mlFirstPoint
      Samp& = (Int(mfHRes * PlotPos)) + FirstPoint& '\ NumChans%
      If Samp& > UBound(mvPlot, 2) Then Samp& = LastSample&
      If gnHexVals Then
         If mnUse64 Then
            SampleVal = sHexArray(Chan, Samp&)
         Else
            x = mvPlot(Chan, Samp&)
            If x > 2147483647 Then x = mvPlot(Chan, Samp&) - 2 ^ 32
            SampleVal = "0x" & Hex$(x)
            'SampleVal = Hex(mvPlot(Chan, Samp&))
         End If
      Else
         SampleVal = mvPlot(Chan, Samp&)
      End If
      GetSampleNum = Samp& * NumChans% + Chan
   End If

End Function

Sub GetSettleStep(SettleStep As Long, SettleRate As Long)

   SettleStep = mlSettleStep
   SettleRate = mlSettleRate

End Sub

Sub InitBlock(IncrementBlock As Integer)
   
   NumChans% = mnLastChan + 1
   ChanCount# = NumChans%
   BlockSize# = mlBlockSize
   DataSize# = ChanCount# * BlockSize#
   msMore = ""
   ChanGroups& = mlBlockSize Mod (mnLastChan + 1)
   BlockIncrement& = (mlBlockSize - 1)
   If mlEndBlock = 0 Then
      If mlEndBuffer > DataSize# Then
         mlEndBlock = mlFirstPoint + BlockIncrement&
         msMore = "More >>"
      Else
         mlEndBlock = BlockIncrement&
      End If
   Else
      If mlEndBuffer > (mlBlockSize * NumChans%) Then
         If IncrementBlock Then mlStartBlock = ((mlEndBlock + 1) * NumChans%)
         If Not mlStartBlock < (mlEndBuffer - ChanGroups&) Then
            'reset to start
            mlStartBlock = 0
            mlEndBlock = BlockIncrement&
            If IncrementBlock Then msMore = "More >>"
         Else
            If IncrementBlock Then msMore = "More >>"
            If (((mlEndBlock + 1) + (BlockIncrement& + 1)) * NumChans%) > mlEndBuffer Then _
            BlockIncrement& = ((mlEndBuffer - (mlStartBlock)) / NumChans%) - 1
            If mlStartText = 0 Then mlEndBlock = ((mlStartBlock / NumChans%) + BlockIncrement&)
            If ((mlEndBlock) >= mlPerChan) Then
               If IncrementBlock Then msMore = "<< More"
            End If
         End If
      Else
         If IncrementBlock Then
            mlStartBlock = 0
            mlEndBlock = mlEndBuffer - 1
         End If
      End If
   End If
   
End Sub

Private Sub InitPlot(Chans As Integer)

   MenusLoaded% = frmPlot.mnuCursChan.Count - 1
   Chans = mnLastChan
   If mnChansSelected = 0 Then mnChansSelected = 1
   If (Not (Chans < 0)) And (Chans < MenusLoaded%) Then
      For i% = MenusLoaded% To Chans + 1 Step -1
         Unload frmPlot.mnuCursChan(i%)
         Unload frmPlot.mnuOneChan(i%)
      Next
      frmPlot.mnuCursChan(0).Checked = True
      mnChanSet = -1
      SetChannel mnChanSet
   End If
   If Chans > MenusLoaded% Then
      For i% = MenusLoaded% + 1 To Chans
         Load frmPlot.mnuCursChan(i%)
         Load frmPlot.mnuOneChan(i%)
         frmPlot.mnuCursChan(i%).Caption = "Ch" & i% & " (Ctl-Alt " & i% & ")"
         frmPlot.mnuOneChan(i%).Caption = "Ch" & i% & " (Ctl " & i% & ")"
         frmPlot.mnuCursChan(i%).Checked = False
         frmPlot.mnuOneChan(i%).Checked = False
      Next
      SetChannel -1
   End If
   
   glaPlotColor(0) = &HFF0000
   glaPlotColor(1) = &HFF&
   glaPlotColor(2) = &H8000&
   glaPlotColor(3) = &HC000C0
   glaPlotColor(4) = &HFFFF00
   glaPlotColor(5) = &HC0C0C0
   glaPlotColor(6) = &H800080
   glaPlotColor(7) = &H0&
   
   If Not mlStartBlock = 0 Then FirstPoint& = mlStartBlock
   If Not (mlFirstPoint = 0) And (FirstPoint& = 0) Then FirstPoint = mlFirstPoint
   If Not mnTransferType = FROMCALLBACK Then
      If mlFirstPoint > 0 Then
         NumPlotPoints& = (mlEndBuffer - mlFirstPoint) - 1
         If NumPlotPoints& > mlBlockSize Then
            NumPlotPoints& = mlBlockSize
            WidthRequired& = mlBlockSize
         End If
      Else
         NumPlotPoints& = (mlEndBuffer \ (mnLastChan + 1)) - 1
      End If
   Else
      NumPlotPoints& = (mlEndBuffer - mlFirstPoint) - 1
   End If
   'If mlBlockSize > mnDefaultWidth Then
   If mlBlockSize < NumPlotPoints& Then
      WidthRequired& = mlBlockSize - 1
   Else
      WidthRequired& = NumPlotPoints&
   End If
   'Else
   '   WidthRequired& = mlBlockSize - 1
   'End If
   If WidthRequired& > mnDefaultWidth Then
      frmPlot.picPlot.ScaleWidth = mnDefaultWidth
      mfHRes = (WidthRequired& - 1) / frmPlot.picPlot.ScaleWidth
   Else
      If mlBlockSize > 0 Then frmPlot.picPlot.ScaleWidth = mlBlockSize
      If WidthRequired& > 1 Then
         frmPlot.picPlot.ScaleWidth = WidthRequired&
      Else
         frmPlot.picPlot.ScaleWidth = 1
      End If
      mfHRes = 1
   End If

End Sub

Sub NextBlock(Chan%)
   
   BufferHandle& = GetBufHandle()
   If BufferHandle& = 0 Then Exit Sub
   'SetBlock
   IncrementBlock% = True
   Plot IncrementBlock%
   If Not frmPlot.txtShow.Visible Then
      If mnShowDiff Then
         'lblSampNumLabel.Caption = "D " & mRefPoint & ":"
         'lblSampNum.Caption = GetSampleNum(0, SampValue#, Chan%) - mvRefSample
         'SampValue# = SampValue# - mvRefVal
         'ValString$ = "D " & mvRefVal & ":"
      Else
         frmPlot.lblSampNumLabel.Caption = "Sample #: "
         frmPlot.lblSampNum.Caption = GetSampleNum(0, SampValue, Chan%)
         ValString$ = "Ch " & Chan% & " value: "
      End If
      frmPlot.lblAmplLabel.Caption = ValString$
      frmPlot.lblAmpl.Caption = Format(SampValue, "0.0####")
   End If

End Sub

Private Sub Plot(IncrementBlock As Integer)

   If Not mnDataInit Then Exit Sub
   InitBlock IncrementBlock
   InitPlot Chans%
   If mnLastChan < 0 Then Exit Sub
   'frmPlot.Visible = True
   SetChannel -2
   frmPlot.linCursor.Visible = False
   Select Case mnPlotType
      Case HISTOGRAM
         'frmPlot.txtShow.Visible = False
         ScalePlot
         PlotHist
      Case VOLTS_VS_TIME, DERIVATIVE
         'frmPlot.txtShow.Visible = False
         If mnUseText Then
            PrintText
         Else
            ScalePlot
            PlotWave
         End If
      Case PRINT_TEXT
         'frmPlot.txtShow.Visible = True
         PrintText
      Case SINGLE_POINT
         'frmPlot.txtShow.Visible = False
         'If mnPointsPlotted(Channel) > 0 Then mnPointsPlotted(Channel) = mnPointsPlotted(Channel) + 1
         ScalePlot
         PlotWave
      Case PLOT_CHUNKS
         If mnUseText Then
            PrintText
         Else
            ScalePlot
            PlotWave
         End If
   End Select
   NumChans% = (mnLastChan + 1)
   If Not mnTransferType = FROMCALLBACK Then
      If mlBlockSize < ((mlArraySize + 1) / NumChans%) Then
         If mlFirstPoint > 0 Then
            PointsPlotting& = mlBlockSize
         Else
            PointsPlotting& = mlBlockSize - mlFirstPoint   '* (mnLastChan + 1)
         End If
      Else
         PointsPlotting& = (mlEndBuffer / NumChans%) - mlFirstPoint  'mlPerChan + 1   'mlArraySize + 1
      End If
   Else
      PointsPlotting& = mlEndBuffer
   End If
   frmPlot.picDetails.Line (0, 2.5)-(15, 4), &HFFFFFF, BF
   frmPlot.picDetails.CurrentX = 0
   frmPlot.picDetails.CurrentY = 2.7
   'PlotPerChan& = mlPerChan + 1
   'If mlBlockSize < mlPerChan Then PlotPerChan& = mlBlockSize / (mlLastChan + 1)
   frmPlot.picDetails.Print "(" & PointsPlotting& & " /chan Plotted)"
   
End Sub

Sub PlotFromMessage(MsgLibrary As Object, DeviceID As String, _
BufCount As Long, Chans As Integer, Timeout As Long)

   'to do - remove this when implemented through buffer
   ReDim MultiDatArray(Chans, BufCount - 1) As Double
   MsgResult$ = MsgLibrary.ReadInScanMultiChan(MultiDatArray(), BufCount, Timeout)
   If SaveMsg(mfNoForm, "ReadInScanMultiChan(" & _
   Format(BufCount, "0") & ")", MsgResult$) Then
      Exit Sub
   End If

   mnLastChan = Chans
   'If (BufCount Mod (Chans + 1) = 0) Then Residual% = 1
   mlPerChan = BufCount '/ (Chans + 1)) - Residual%
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   If Not mlEndBuffer = 0 Then
      If (Not mlEndBuffer = BufCount) Or (Not Chans = mnLastChan) _
      Then RestorePlot False
   End If
   mlEndBuffer = BufCount
   mlArraySize = mlEndBuffer - 1
   mnTransferType = FROMCALLBACK
   mlEndBlock = 0
   mlStartBlock = 0
   'ReDim mvPlotArray(mnLastChan, mlArraySize)
   NumChans% = UBound(MultiDatArray, 1)
   If UBound(MultiDatArray, 1) < Chans Then
      MsgBox "Although " & Chans + 1 & " channels of data were expected, only " & _
      NumChans% + 1 & " channels were returned.", vbInformation, "Incomplete Data"
   End If
   mvPlot = MultiDatArray(Ch%, l&)
   'For Ch% = 0 To NumChans%
   '   For l& = 0 To mlArraySize
   '      mvPlot(Ch%, l&) = MultiDatArray(Ch%, l&)
   '   Next l&
   'Next Ch%
   mnDataInit = True
   Plot False

End Sub

Sub PlotFromCallback(CBArray As Variant, BufCount As Long, Chans As Integer)

   mnLastChan = Chans
   mlPerChan = BufCount
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   If Not mlEndBuffer = 0 Then
      If (Not mlEndBuffer = BufCount) Or (Not Chans = mnLastChan) _
      Then RestorePlot False
   End If
   mnTransferType = FROMCALLBACK
   mlEndBuffer = BufCount + 1
   mlArraySize = BufCount
   mlEndBlock = 0
   mlStartBlock = 0
   'ReDim mvPlotArray(mnLastChan, mlPerChan)
   mvPlot = CBArray()
   'For Ch% = 0 To Chans
   '   For l& = 0 To mlPerChan
   '      mvPlot(Ch%, l&) = CBArray(Ch%, l&)
   '   Next l&
   'Next Ch%
   mnDataInit = True
   Plot False

End Sub

Sub PlotBuffer(BufHandle As Long, BufCount As Long, Chans As Integer)

   mvMemHandle = BufHandle
   If mvMemHandle = 0 Then Exit Sub
   If Chans < 0 Then Exit Sub
   mnLastChan = Chans
   If (BufCount Mod (Chans + 1) = 0) Then Residual% = 1
   mlPerChan = (BufCount / (Chans + 1)) - Residual%
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   If (Not mlEndBuffer = 0) And (Not mnPlotFixedBlock) Then
      If (Not mlEndBuffer = BufCount) Or (Not Chans = mnLastChan) _
      Then RestorePlot False
   End If
   mlEndBuffer = BufCount
   mlArraySize = mlEndBuffer - 1
   If Not mnLibType = MSGLIB Then mnTransferType = BUFFER
   mlEndBlock = 0
   mlStartBlock = 0
   TransferBlock

End Sub

Sub PlotFile(Filename$)

   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   frmPlot.Cls
   frmPlot.linCursor.Visible = False
   mnTransferType = FILE
   msFileName = Filename$
   
   Dim NoForm As Form
   ULStat = cbFileGetInfo(msFileName, LowChan%, HighChan%, PretrigCount&, TotalCount&, CBRate&, Gain&)
   If SaveFunc(mfNoForm, FileGetInfo, ULStat, msFileName, LowChan%, HighChan%, PretrigCount&, TotalCount&, CBRate&, Gain&, A8, A9, A10, A11, 0) Then Exit Sub
   Chans% = HighChan% - LowChan%
   
   If TotalCount& = 0 Then Exit Sub
   mlEndBuffer = TotalCount& - 1
   mlPreTrigCount = PretrigCount&
   RefLine& = mlPreTrigCount
   DrawLine RefLine&
   InitPlot Chans%
   mnLastChan = Chans%
   If mnPlotType = PRINT_TEXT Then
      CrLf$ = Chr$(13) & Chr$(10)
      msPreamble = "File name: " & msFileName & CrLf$
      msPreamble = msPreamble & "Low channel: " & LowChan% & CrLf$
      msPreamble = msPreamble & "High channel: " & HighChan% & CrLf$
      msPreamble = msPreamble & "Pretrigger points: " & PretrigCount& & CrLf$
      msPreamble = msPreamble & "Total points: " & TotalCount& & CrLf$
      msPreamble = msPreamble & "Rate per channel: " & CBRate& & CrLf$
      msPreamble = msPreamble & "Range: " & GetRangeString(CInt(Gain&)) & CrLf$ & CrLf$
   End If
   InitBlock True
   TransferBlock

End Sub

Private Sub PlotHist()

   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   Chan% = mnChanSet
   If mnChanSet < 0 Then Chan% = 0
   ReDim mHistBins(0)

   'mvFirstBin = code of the first bin with data in it
   'mvLargestBin = number of hits in the bin with the most data in it (primary bin)
   'LargestBinCode = code of mvLargestBin (primary bin)

   mlNumBins = BinData(Chan%, mvPlot, FirstBin, mvLargestBin, LargestBinCode, mHistBins(), Divisor%, Average)
   mvAverage = Average
   If mnCalcRMS Then
      ArraySize = UBound(mvPlot, 2)
      For i& = 0 To ArraySize
         If Not (IsEmpty(mvPlot(Chan%, i&))) Then
           noise = Average - mvPlot(Chan%, i&)
           square_total = square_total + (noise ^ 2)
         End If
      Next i&
      mvLsbRms = (square_total / ArraySize) ^ 0.5
   End If
   
   CrLf$ = Chr$(13) & Chr$(10)
   frmPlot.picPlot.ScaleWidth = (mlNumBins + 2) / Divisor%
   frmPlot.picPlot.ScaleTop = mvLargestBin + 2
   frmPlot.picPlot.ScaleHeight = -1 * (mvLargestBin + 2)
   If mlArraySize Then
      If Not mnContinuous Then frmPlot.ZOrder
      If Not mnUseText Then
         frmPlot.picPlot.ScaleWidth = (mlNumBins + 2) / Divisor%
         frmPlot.CurrentX = 50
         frmPlot.CurrentY = frmPlot.picPlot.Height - 100
         frmPlot.Print "Plotting..."
         DoEvents
         
         If Not mnRetainPlot Then frmPlot.picPlot.Cls
         frmPlot.picPlot.ForeColor = glaPlotColor(Chan% Mod 8)
         frmPlot.picPlot.CurrentX = 1
         frmPlot.picPlot.CurrentY = 0
      Else
         frmPlot.txtShow.Text = ""
      End If
      If mlNumBins < 32767 Then
         For d& = 0 To (mlNumBins / Divisor%) - 1
            If Not mnUseText Then
               frmPlot.picPlot.Line (d& + 1, 0)-(d& + 1, mHistBins(d&))
            Else
               frmPlot.txtShow.Text = frmPlot.txtShow.Text _
                  & Format$(mvFirstBin + d&, "0") & Chr$(9) _
                  & Format$(mHistBins(d&), "0") & CrLf$
            End If
            If mHistBins(d&) = mvLargestBin Then PrimElement& = d&
         Next d&
         TotalElements& = mlArraySize + 1
         If Not (PrimElement& - 1) < 0 Then Totalm1& = mHistBins(PrimElement& - 1)
         If Not (PrimElement& - 2) < 0 Then Totalm2& = mHistBins(PrimElement& - 2)
         If Not ((PrimElement& + 1) > (mlNumBins - 1)) Then Totalp1& = mHistBins(PrimElement& + 1)
         If Not ((PrimElement& + 2) > (mlNumBins - 1)) Then Totalp2& = mHistBins(PrimElement& + 2)
         Prim! = ((mvLargestBin) / TotalElements&) * 100
         Mid3! = ((mvLargestBin + Totalm1& + Totalp1&) / TotalElements&) * 100
         Mid5! = ((mvLargestBin + Totalm2& + Totalm1& + Totalp2& + Totalp1&) / TotalElements&) * 100
         'frmPlot.picDetails.Cls
         Indent% = 32
         frmPlot.picDetails.Line (Indent% + 4, 1)-(Indent% + 32, 4), &HFFFFFF, BF
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.Print "Primary: " & Prim! & "%"
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Primary ±1: " & Mid3! & "%"
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Primary ±2: " & Mid5! & "%"
         frmPlot.picDetails.CurrentX = Indent% + 16
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.Print "Average: " & Format$(Average, "#.##")
         If mnCalcRMS Then
            frmPlot.picDetails.CurrentX = Indent% + 16
            frmPlot.picDetails.Print "RMS Noise: " & Format$(mvLsbRms, "0.000") & " LSBs"
         End If
      Else
         Indent% = 32
         frmPlot.picDetails.Line (Indent% + 4, 1)-(Indent% + 32, 4), &HFFFFFF, BF
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.Print "Histogram overflow (too many bins)"
      End If
   Else
      If Not mnRetainPlot Then frmPlot.picPlot.Cls
   End If
   frmPlot.Cls

End Sub

Function PlotIntArray(DatArray() As Integer)

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
   'On Error GoTo PIAHandler
   If mnTransferType = RealArray Then mnTransferType = FROMARRAY
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
      'mnPlotFormsLoaded = True
   End If
   If mnTransferType = FROMARRAY Then mlPerChan = UBound(DatArray, 2)
   
   ReDim PlotArray(mnLastChan, mlPerChan) As Variant
   mnDataInit = True
   x& = IntArrayToULong(DatArray(), PlotArray(), mlArraySize, 0)
   
   Exit Function

AbortPIA:
   Exit Function
   
PIAHandler:
   If Err = 9 Then
      mlArraySize = mLastChan%
      If mChanSet% Then
         mnLastChan = mChanSet%
      Else
         mnLastChan = 0
      End If
      Resume Next
   Else
      MsgBox Error$(Err)
      'Resume 0
      Resume AbortPIA
   End If

End Function

Sub PlotLongArray(DatArray() As Integer, NumPoints As Long, FirstPoint As Long)

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
   'On Error GoTo PIAHandler
   If mnTransferType = RealArray Then mnTransferType = FROMARRAY
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
      'mnPlotFormsLoaded = True
   End If
   WholeArray& = UBound(DatArray, 2)
   
   Chans% = UBound(DatArray, 1)
   If mnTransferType = FROMARRAY Then mlArraySize = NumPoints \ (Chans% + 1)
   InitPlot Chans%
   mnLastChan = Chans%
   If mnTransferType = FROMARRAY Then
      mlEndBuffer = mlArraySize * (mnLastChan + 1)
      InitBlock True
   End If
   
   mvPlot = DatArray()
   'ReDim mvPlotArray(mnLastChan, mlArraySize)
   'For Ch% = 0 To mnLastChan
   '   For l& = 0 To mlArraySize - 1
   '      mvPlot(Ch%, l&) = DatArray(Ch%, l& + FirstPoint \ (mnLastChan + 1))
   '   Next l&
   'Next Ch%
   mnDataInit = True
   'x& = IntArrayToULong(DatArray(), mvPlotArray())
   ScalePlot
'   Plot
   Exit Sub

AbortPLA:
   Exit Sub
   
PLAHandler:
   If Err = 9 Then
      mlArraySize = mLastChan%
      If mChanSet% Then
         mnLastChan = mChanSet%
      Else
         mnLastChan = 0
      End If
      Resume Next
   Else
      MsgBox Error$(Err)
      'Resume 0
      Resume AbortPLA
   End If

End Sub

Sub PlotMemory(BoardNum As Integer, MemBoard As Integer, NumPoints As Long, FirstPoint As Long, Chans As Integer)

   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
      'mnPlotFormsLoaded = True
   End If
   frmPlot.Cls
   frmPlot.linCursor.Visible = False
   mnTransferType = XMEM
   ParentBoard% = mnBoardNum
   mnBoardNum = BoardNum
   mnMemBoard = MemBoard
   ULStat = cbMemReset(mnMemBoard)
   If SaveFunc(mfNoForm, MemReset, ULStat, mnMemBoard, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

   mlEndBuffer = 128000000 'NumPoints - 1
   InitPlot Chans
   mnLastChan = Chans
'   InitBlock
   mlStartBlock = FirstPoint
   TransferBlock

End Sub

Sub PlotMemPretrig(BoardNum As Integer, MemBoard As Integer, NumPoints As Long, Chans As Integer)

   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
      'mnPlotFormsLoaded = True
   End If
   frmPlot.Cls
   frmPlot.linCursor.Visible = False
   mnTransferType = MEM_PRETRIG
   mnBoardNum = BoardNum
   mnMemBoard = MemBoard
   ULStat = cbMemReset(mnMemBoard)
   If SaveFunc(mfNoForm, MemReset, ULStat, mnBoardNum, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

   mlEndBuffer = NumPoints - 1
   InitPlot Chans
   mnLastChan = Chans
'   InitBlock
   TransferBlock

End Sub

Sub PlotPartialBuf(BufHandle As Long, BufCount As Long, BufStart As Long, Chans As Integer)

   mvMemHandle = BufHandle
   If mvMemHandle = 0 Then Exit Sub
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   mlEndBuffer = BufCount / 2
   mlArraySize = BufCount
   mnLastChan = Chans
   mlStartBlock = BufStart
   mnTransferType = PARTBUFFER
   mlPerChan = (BufCount / (Chans + 1)) - 1
   
   TransferBlock

End Sub

Function PlotPoint(PointVal As Variant, Channel As Integer)

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
   Static stPlotBegun%, stBlockSize%, stCurrentChan%
   mnPointsPlotted(Channel) = mnPointsPlotted(Channel) + 1
   If VarType(PointVal) > 3 Then mnTransferType = RealArray
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
      'mnPlotFormsLoaded = True
   End If
   mlArraySize = mlBlockSize
   
   frmPlot.linCursor.Visible = False
   For i% = 0 To mnLastChan
      If mnChansLogged(i%) = Channel Then
         Match% = True
         If stCurrentChan% <> i% Then
            'mnPointsPlotted(stCurrentChan%) = mnPointsPlotted(stCurrentChan%) - 1
            stCurrentChan% = i%
         End If
      End If
   Next i%

   If (Not Match%) Or (Not stPlotBegun%) Or (stBlockSize% <> mlBlockSize) Then
      If Not stPlotBegun% Then mnLastChan = 0
      If (Not Match%) Then
         mnLastChan = mnLastChan + 1
         mnPointsPlotted(stCurrentChan%) = mnPointsPlotted(stCurrentChan%) - 1
         stCurrentChan% = mnLastChan
      End If
'      InitBlock
      InitPlot mnLastChan
      If Not mnRetainPlot Then frmPlot.picPlot.Cls
      'ReDim Preserve mvPlotArray(8, mlBlockSize)
      ReDim Preserve mnPointsPlotted(mnLastChan)
      ReDim Preserve mnChansLogged(mnLastChan)
      frmPlot.picPlot.ScaleWidth = mlBlockSize
      stBlockSize% = mlBlockSize
      mnChansLogged(mnLastChan) = Channel
      frmPlot.Cls
   End If
   DoEvents

   frmPlot.Visible = True
   'frmPlot.SetFocus
   If Not mnContinuous Then frmPlot.ZOrder
   If mnPointsPlotted(stCurrentChan%) > mlBlockSize - 1 Then
      mnPointsPlotted(stCurrentChan%) = 0
      If Not mnRetainPlot Then frmPlot.picPlot.Cls
   End If
   If mnTransferType = RealArray Then
      mvPlot(stCurrentChan%, mnPointsPlotted(stCurrentChan%)) = PointVal
   Else
      mvPlot(stCurrentChan%, mnPointsPlotted(stCurrentChan%)) = IntValToULong(CInt(PointVal))
   End If
   
'   InitBlock
   ScalePlot
'   Plot
   'mnPointsPlotted(Channel) = mnPointsPlotted(Channel) + 1
   stPlotBegun% = True

End Function

Function PlotRealArray(DatArray() As Single, NumPoints As Long)

   'If gnIDERunning Then
   '   Stop
   'Else
   '   Dim Resp As VbMsgBoxResult
   '   Resp = MsgBox("This path is a Stop statement " & _
   '   "in the IDE. Check Local Error Handling options. " _
   '   & vbCrLf & vbCrLf & "          Click Yes to attempt " & _
   '   "to continue, No to exit application.", _
   '   vbYesNo, "Attempt To Continue?")
   '   If Resp = vbNo Then End
   'End If
   On Error GoTo PRAHandler
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
   End If
   mlEndBuffer = NumPoints
   
   If UBound(DatArray, 2) = 0 Then RePlot False
   mnTransferType = RealArray
   mlArraySize = UBound(DatArray, 2)
   Chans% = UBound(DatArray, 1)
   InitPlot Chans%
   mnLastChan = Chans%
   mlBlockSize = (mlArraySize + 1) * (mnLastChan + 1)
   'If mlStartBlock = 0 Then InitBlock
   'ReDim mvPlotArray(mnLastChan, mlArraySize)
   mnDataInit = True
   
   mvPlot = DatArray()
   'For Ch% = 0 To mnLastChan
   '   For l& = 0 To mlArraySize
   '      mvPlot(Ch%, l&) = DatArray(Ch%, l&)
   '   Next l&
   'Next Ch%
   If mnUseText Then
      Plot False
   Else
      ScalePlot
   End If
'   Plot
   Exit Function

AbortPRA:
   Exit Function

PRAHandler:
   If Err = 9 Then
      mlArraySize = mLastChan%
      If mChanSet% Then
         mnLastChan = mChanSet%
      Else
         mnLastChan = 0
      End If
      Resume Next
   Else
      MsgBox Error$(Err)
      Resume AbortPRA
   End If

End Function

Private Sub PlotWave()

   CrLf$ = Chr$(13) & Chr$(10)
   LastChan% = UBound(manChans)
   FirstChan% = 0
   NumChans% = LastChan% + 1
   
   'If mlArraySize = 0 Then Exit Sub
   If mlEndBuffer = 0 Then Exit Sub
   If Not mnContinuous Then frmPlot.ZOrder
   DoEvents
   If (mlBlockSize - 1) < mlArraySize Then
      PointsPlotting& = mlBlockSize * (mnLastChan + 1)
   Else
      PointsPlotting& = mlArraySize + 1
   End If
   frmPlot.Cls
   frmPlot.CurrentX = 50
   frmPlot.CurrentY = frmPlot.picPlot.Height - 500
   frmPlot.Print "Plotting"
   frmPlot.CurrentX = 50
   frmPlot.CurrentY = frmPlot.picPlot.Height - 200
   If mnTransferType = FROMCALLBACK Then PointsPlotting& = PointsPlotting& * NumChans%
   frmPlot.Print PointsPlotting& & " points..."
   PointsPlotted% = 0
   DoEvents
      
   If Not mnRetainPlot Then
      frmPlot.picPlot.Cls
      'frmPlot.txtShow.Text = ""
   End If

   If 0 Then   'mnPlotType = SINGLE_POINT
      For CurChan% = FirstChan% To LastChan%
         Chan% = manChans(CurChan%)
         PointsToPlot& = mnPointsPlotted(CurChan%)
         If PointsToPlot& = 0 Then
            'just position start point and don't
            'draw line (set PlotStart > PointsToPlot)
            frmPlot.picPlot.Cls
            If (CurChan% = 0) And (Not mnRetainPlot) Then
               frmPlot.picPlot.Cls
               If mnUseText Then frmPlot.txtShow.Text = ""
            End If
            PlotStart% = 1
            CurPoint = mvPlot(Chan%, 0) - mvFirstBin
            If mnUseText Then
               frmPlot.txtShow.Text = "0" & _
                  Chr$(9) & Format$(mvPlot(Chan%, 0), "0.0000") & CrLf$
            Else
               frmPlot.picPlot.CurrentX = 0
               frmPlot.picPlot.CurrentY = CurPoint / mvDivisor
               If mnHardCopy Then
                  Printer.CurrentX = 0
                  Printer.CurrentY = PlotRef& - (CurPoint / mvDivisor) + mlPrinterOS
               End If
            End If
         Else
            PlotStart% = PointsToPlot&
            If (PointsToPlot& >= frmPlot.picPlot.ScaleWidth) Or mnScaleChange Then
               frmPlot.picPlot.Cls
               PlotStart% = 1
               If mnUseText Then
                  frmPlot.txtShow.Text = "0" & Chr$(9) _
                     & Format$(mvPlot(Chan%, 0), "0.0000") & CrLf$
               Else
                  CurPoint = mvPlot(Chan%, 0) - mvFirstBin
                  frmPlot.picPlot.CurrentX = 0
                  frmPlot.picPlot.CurrentY = CurPoint / mvDivisor
               End If
            End If
         End If
      Next CurChan%
   Else
      PointsToPlot& = frmPlot.picPlot.ScaleWidth '- 1
      If mnPlotType = PLOT_CHUNKS Then PointsToPlot& = (mlNumChunks) / mfHRes '- 1  \ NumChans%
      PlotStart% = 1
      DataStart& = (mlStartBlock \ NumChans%) / mfHRes  '(mnLastChan + 1)
      If mlStartBlock = 0 Then DataStart& = (mlFirstPoint \ NumChans%) / mfHRes
      For CurChan% = FirstChan% To LastChan%
         Chan% = manChans(CurChan%)
         frmPlot.picPlot.ForeColor = glaPlotColor(Chan% Mod 8)
         CurPoint = (mvPlot(Chan%, Int(DataStart& * mfHRes)) - mvFirstBin)
         frmPlot.picPlot.CurrentX = 0
         If mvDivisor = 0 Then mvDivisor = &H1
         frmPlot.picPlot.CurrentY = CurPoint / mvDivisor
         If PointsToPlot& > 0 Then
            For i& = PlotStart% To PointsToPlot&
               Sample& = Int((DataStart& + i&) * mfHRes)
               If ((Sample& - mlFirstPoint) < mlEndBlock) Then
                  If Not ((Sample& - DataStart&) > mlPerChan) Then
                     If mnUseText Then
                        frmPlot.txtShow.Text = frmPlot.txtShow.Text _
                           & Format$(i&, "0") & Chr$(9) _
                           & Format$(mvPlot(Chan%, i&), "0.0000") & CrLf$
                     Else
                        CurPoint = mvPlot(Chan%, Sample&) - mvFirstBin
                        frmPlot.picPlot.Line -(i&, CurPoint / mvDivisor)
                     End If
                  End If
               End If
            Next i&
         End If
      Next CurChan%
   End If
   'frmPlot.picDetails.Line (32, 1.9)-(46, 3.5), &HFFFFFF, BF
   If (mlRefLine > 0) Or (mlRefLine2 > 0) Then
      SaveColor& = frmPlot.picPlot.ForeColor
      frmPlot.picPlot.ForeColor = &HFF00&
      frmPlot.picPlot.DrawStyle = 2
      If mlRefLine > 0 Then
         ViewStart& = mlStartBlock * (mnLastChan + 1)
         ViewSize& = mlBlockSize * (mnLastChan + 1)
         If (mlRefLine > mlBlockSize) And (mlRefLine > ViewStart&) Then
            If (mlRefLine > ViewStart&) And ((ViewStart& + ViewSize&) - mlRefLine > 0) Then
               NewLine& = mlRefLine - ViewStart&
               LinePosition& = (Int(NewLine& \ (mnLastChan + 1)) / mfHRes)
               frmPlot.picPlot.Line (LinePosition&, frmPlot.picPlot.ScaleTop)-(LinePosition&, frmPlot.picPlot.ScaleHeight)
            End If
         Else
            If mlRefLine > ViewStart& Then
               LinePosition& = (Int(mlRefLine \ (mnLastChan + 1)) / mfHRes)
               frmPlot.picPlot.Line (LinePosition&, frmPlot.picPlot.ScaleTop)-(LinePosition&, frmPlot.picPlot.ScaleHeight)
            End If
         End If
         frmPlot.picDetails.Line (30, 0.799999970197678)-(46, 1.7), &HFFFFFF, BF
         frmPlot.picDetails.CurrentX = 30
         frmPlot.picDetails.CurrentY = 0.799999970197678 '  2.7
         frmPlot.picDetails.Print "Ref Line @ " & mlRefLine
      End If
      If mlRefLine2 > 0 Then
         LinePosition& = (Int(mlRefLine2 \ (mnLastChan + 1)) / mfHRes)
         frmPlot.picPlot.Line (LinePosition&, frmPlot.picPlot.ScaleTop)-(LinePosition&, frmPlot.picPlot.ScaleHeight)
         frmPlot.picDetails.CurrentX = 30
         frmPlot.picDetails.CurrentY = 0.8 '  2.7
         frmPlot.picDetails.Print "Ref Line @ " & mlRefLine2
      End If
      DoEvents
      frmPlot.picPlot.ForeColor = SaveColor&
      frmPlot.picPlot.DrawStyle = 0
   End If
   'If Not mnRetainPlot Then frmPlot.picPlot.Cls
   
   If (mnPlotType <> SINGLE_POINT) And (mnPlotType <> PLOT_CHUNKS) Then
      frmPlot.Cls
      frmPlot.CurrentX = 50
      frmPlot.CurrentY = frmPlot.picPlot.Height - 200
      frmPlot.Print msMore
   End If
   mnScaleChange = False

End Sub

Sub PrintList()
   
   Static stCurIndex&
   PrintBlock% = mlBlockSize
   If stCurIndex& = 0 Then stCurIndex& = 1
   If Len(msListString) > PrintBlock% Then
      msMore = "More >>"
      If Len(msListString) - (stCurIndex& + PrintBlock%) < 80 Then
         PrintBlock% = Len(msListString) - stCurIndex&
         Search% = False
      Else
         SearchBlock% = 80
         Search% = True
      End If
      If Search% Then
         For i% = 0 To SearchBlock%
            If Asc(Mid$(msListString, stCurIndex& + PrintBlock% + i%, 1)) = 13 Then
               PrintBlock% = PrintBlock% + i%
               Exit For
            End If
         Next i%
      End If
      frmPlot.txtShow.Text = Mid$(msListString, stCurIndex&, PrintBlock%)
      stCurIndex& = stCurIndex& + PrintBlock% + 2
      If stCurIndex& >= Len(msListString) Then
         msMore = "<< More"
         stCurIndex& = 0
      End If
      frmPlot.Cls
      frmPlot.CurrentX = 50
      frmPlot.CurrentY = frmPlot.picPlot.Height - 200
      frmPlot.Print msMore
   Else
      frmPlot.txtShow.Text = msListString
   End If
   If Not mnContinuous Then frmPlot.ZOrder

End Sub

Private Sub PrintText()

   Dim InfoRetrieved As Boolean
   
   NumChans% = mnLastChan + 1
   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      frmPlot.Visible = True
      'mnPlotFormsLoaded = True
   End If
   frmPlot.Visible = True
   CrLf$ = Chr$(13) & Chr$(10)
   TextString$ = "Sample" & Chr$(9)
   For Ch% = 0 To mnLastChan
      If mfTickSize = 0 Then
         If (mnTransferType = RealArray) Or mnFloat Or mnConvToEng Then
            TextString$ = TextString$ & "   ch" & Ch% + mnFirstChan & Chr$(9) & Chr$(9)
         Else
            TextString$ = TextString$ & "ch" & Ch% + mnFirstChan & Chr$(9)
         End If
      Else
         TextString$ = TextString$ & "  ch" & Ch% + mnFirstChan & "(ms)" & Chr$(9) ' & Chr$(9)
      End If
   Next Ch%
   TextString$ = TextString$ & CrLf$ & CrLf$
   If Not mnContinuous Then frmPlot.ZOrder
   
   LastPoint& = mlEndBlock
   If LastPoint& > mlPerChan Then LastPoint& = mlPerChan
   If mlTextCount = 0 Then
      mlTextCount = 1000
   End If
   
   Residual% = mlFirstPoint Mod NumChans%
   mlFirstPoint = mlFirstPoint - Residual%
   Residual% = mlStartText Mod NumChans%
   mlStartText = mlStartText - Residual%
   FirstPoint& = (mlStartBlock / NumChans%) + (mlStartText / NumChans%) + mlFirstPoint
   Residual% = mlTextCount Mod NumChans%
   mlTextCount = mlTextCount - Residual%
   If (FirstPoint& + mlTextCount) < LastPoint& Then LastPoint& = (FirstPoint& + mlTextCount) - 1
   
   'Resolution% = GetFormProperty("resolution")
   For i& = FirstPoint& To LastPoint&
      TextString$ = TextString$ & (i&) * (mnLastChan + 1) & Chr$(9)
      If (mnTransferType = RealArray) Or mnFloat Or mnConvToEng Then
         For Ch% = 0 To mnLastChan
            CurReading = mvPlot(Ch%, i&)
            If CurReading < -999 Then
               If (mfOwner.Tag > 99) And (mfOwner.Tag < 200) Then
                  TCInfo$ = mfOwner.GetTCInfo(CurReading)
                  InfoRetrieved = Not (TCInfo$ = "")
               End If
            Else
               InfoRetrieved = False
            End If
            If InfoRetrieved Then
               TextString$ = TextString$ & TCInfo$ & Chr$(9)
            Else
               If mfTickSize = 0 Then
                  If Abs(mvPlot(Ch%, i&)) > 999 Then
                     TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "+.000E+00;-.000E+00") & Chr$(9)
                  ElseIf Abs(mvPlot(Ch%, i&)) < 0.0001 Then
                     TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "+0.000E+00;-0.000E+00") & Chr$(9)
                  Else
                     TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "+##0.0000000;-##0.0000000") & Chr$(9)
                  End If
               Else
                  TextString$ = TextString$ & Format$(mvPlot(Ch%, i&) * 1000, "+0.0000000;-0.0000000") & Chr$(9)
               End If
            End If
         Next Ch%
      Else
         For Ch% = 0 To mnLastChan
            If gnHexVals Then
               'If mnUse64 And Not mnFloat Then
                  'TextString$ = TextString$ & "0x" & sHexArray(Ch%, i&) & Chr$(9)
               'Else
               '   x = mvPlot(Ch%, i&)
               '   If x > 2147483647 Then x = mvPlot(Ch%, i&) - 2 ^ 32
               '   TextString$ = TextString$ & "0x" & Hex$(x) & Chr$(9)
               'End If
               'Temp$ = GetHexValue(mvPlot(Ch%, i&), mnResolution)
               TextString$ = TextString$ & GetHexValue(mvPlot(Ch%, i&), mnResolution)
            Else
               TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "0") & Chr$(9)
            End If
         Next Ch%
      End If
      TextString$ = TextString$ & CrLf$
   Next i&
   frmPlot.txtShow.Text = msPreamble & TextString$
   frmPlot.Cls
   frmPlot.CurrentX = 50
   frmPlot.CurrentY = frmPlot.picPlot.Height - 200
   frmPlot.Print msMore

End Sub

Sub RePlot(ScaleChanged As Integer)

   'for use only by 'Plot.frm'
   If Not mnDataInit Then Exit Sub
   If mnPlotFormsLoaded = 0 Then Exit Sub
   If ScaleChanged Then
      InitPlot True
      ScalePlot
   End If
   Plot False
   Exit Sub
   ScalePlot
   frmPlot.Cls
   frmPlot.linCursor.Visible = False
   Select Case mnPlotType
      Case HISTOGRAM
         'frmPlot.txtShow.Visible = False
         PlotHist
      Case VOLTS_VS_TIME
         'frmPlot.txtShow.Visible = False
         If mnUseText Then
            PrintText
         Else
            PlotWave
         End If
      Case PRINT_TEXT
         'frmPlot.txtShow.Visible = True
         PrintText
      Case SINGLE_POINT
         'frmPlot.txtShow.Visible = False
         If mnPointsPlotted(0) > 0 Then PlotWave
      Case PLOT_CHUNKS
         PlotWave
   End Select
   If mnBlocksEnabled Then
      If mlEndBlock < mlEndBuffer Then
         msMore = "More >>"
      Else
         msMore = "<< More"
      End If
   End If

End Sub

Sub RestorePlot(RefreshPlot As Integer)

   SetFirstPoint 0
   SetBlockSize -1, False
   'BufferHandle& = GetBufHandle()
   IncrementBlock% = False
   InitBlock IncrementBlock%
   If RefreshPlot Then TransferBlock
   frmPlot.mnuRestore.ENABLED = False

End Sub

Sub ReturnToOwner()

   DoEvents
   Set gfOwner = mfOwner
   mfOwner.SetFocus

End Sub

Sub SaveBlockSize()
      
   If mlStaticBlockSize = 0 Then Exit Sub
   
   lplFileName$ = "PlotForm.ini"
   lpApplicationName$ = "PlotParams"
   lpKeyName$ = "NumPoints"
   lpString$ = Str$(mlStaticBlockSize)
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)

End Sub

Private Sub ScalePlot()
      
   Static stLastBin
   
   OriginalH! = frmPlot.picPlot.ScaleHeight
   OriginalT! = frmPlot.picPlot.ScaleTop

   If (mnPlotType = PRINT_TEXT) Or (mnPlotType = PRINT_LIST) Then Exit Sub
   ScaleMode% = mnAutoScale
   If mnPlotType = HISTOGRAM Then ScaleMode% = 1
   If mnPlotType = SINGLE_POINT Then
      If (mnPointsPlotted(Chan%) > 0) And (mnPointsPlotted(Chan%) _
         < frmPlot.picPlot.ScaleWidth) Then ScaleMode% = 0
   End If
   Select Case ScaleMode%
      Case -1  'manual scaling
         mvFirstBin = 0
         PlotMax! = frmPlot.picPlot.ScaleTop
         PlotMin! = frmPlot.picPlot.ScaleTop + frmPlot.picPlot.ScaleHeight
         PlotMode$ = "Manual scale ("
      Case 0   'fixed full scale
         If (mnTransferType = RealArray) Or mnScaled Or mnConvToEng Then
            GoSub EvalData
            If Abs(MaxBin) > Abs(mvFirstBin) Then
               TestBin = Abs(MaxBin)
            Else
               TestBin = Abs(mvFirstBin)
            End If
            Scaler% = 0
            result! = 2
            Do
               Scaler% = Scaler% + 1
               CurScale = SD
               SD = Choose(Scaler%, 500000, 100000, 50000, 10000, _
                  5000, 1000, 500, 200, 100, 50, 20, 10, 5, 4, 2, 1)
               If IsNull(SD) Then SD = Choose(Scaler% - 13, 1, 0.5, _
                  0.2, 0.1, 0.05, 0.02, 0.01, 0.005, 0.002, 0.001)
               If Not IsNull(SD) Then
                  ScaleDiv! = Val(SD)
                  result! = TestBin / ScaleDiv!
               End If
            Loop While result! < 0.9 And Scaler% < 23
            If Scaler% > 1 Then Scaler% = Scaler% - 1
            SD = CurScale 'Choose(Scaler%, 10000, 5000, 1000, 500, 200, 100, 50, 20, 10, 5, 4, 2, 1)
            If IsNull(SD) Then SD = Choose(Scaler% - 13, 1, 0.5, 0.2, 0.1, 0.05, 0.02, 0.01, 0.005, 0.002, 0.001)
            ScaleDiv! = Val(SD)
            LastBin = ScaleDiv!
            mvFirstBin = -1 * ScaleDiv!
            mn16Bit = False
            
         Else
            If mnResolution = 0 Then
               GoSub EvalData
               mvFirstBin = 0
               Select Case LastBin
                  Case Is < 17
                     LastBin = 15
                     mn16Bit = False
                  Case Is < 65
                     LastBin = 63
                     mn16Bit = False
                  Case Is < 257
                     LastBin = 255
                     mn16Bit = False
                  Case Is < 1025
                     LastBin = 1023
                     mn16Bit = False
                  Case Is < 4097
                     LastBin = 4095
                     mn16Bit = False
                  Case Is < 8192
                     LastBin = 8191
                     mn16Bit = False
                  Case Is < 16385
                     LastBin = 16383
                     mn16Bit = False
                  Case Is < 65537
                     LastBin = 65535
                     mn16Bit = True
                  Case Is < 1048577
                     LastBin = 1048575
                     mn16Bit = True
                  Case Is < 16777217
                     LastBin = 16777215
                     mn16Bit = True
                     mvDivisor = &H100
                  Case Is < 4294967296#
                     LastBin = 4294967295#
                     mn16Bit = True
                     mvDivisor = &H10000
                  Case Else
                     LastBin = 66535
                     mvFirstBin = -1000
                     mn16Bit = True
               End Select
            Else
               Resolution% = mnResolution
               'If mnUse32 Then Resolution% = 32
               LastBin = 2 ^ Resolution% - 1
               mvFirstBin = 0
            End If
         End If
         PlotMax! = LastBin
         PlotMin! = mvFirstBin
         PlotMode$ = "Fixed full scale ("
         mvDivisor = &H1
         RealData% = mnScaled Or mnConvToEng Or (mnTransferType = RealArray)
         If (LastBin > 4095) And (Not RealData%) Then mvDivisor = &H10
         If mnResolution = 64 Then mvDivisor = 281474976710656# '2 ^ 32
         If mnResolution = 48 Then mvDivisor = &H1000000
         If mnResolution = 32 Then mvDivisor = &H10000
         If mnResolution = 24 Then mvDivisor = &H100
         stScaleFactor! = (LastBin - mvFirstBin) / mvDivisor
         If Not (stScaleFactor! = 0) Then
            frmPlot.picPlot.ScaleHeight = -1 * stScaleFactor!
            frmPlot.picPlot.ScaleTop = stScaleFactor! '/ mvDivisor
         End If
         'frmPlot.Caption = PlotMode$ & PlotMin! & " to " & PlotMax! & ")"
      Case 1   'auto scale
         GoSub EvalData
         If stLastBin = mvFirstBin Then
            PlotMax! = mvFirstBin + 2
            PlotMin! = mvFirstBin - 2
         Else
            stLastBin = MaxBin 'LastBin
            PlotMax! = MaxBin 'LastBin
            PlotMin! = mvFirstBin
         End If
         If mnPlotType = HISTOGRAM Then
            PlotMode$ = "Bins: "
         Else
            PlotMode$ = "Auto scale ("
         End If
         mvDivisor = 1
         If (mnTransferType = RealArray) Or mnConvToEng Or mnScaled Then
            If stLastBin = mvFirstBin Then
               'Margin = stLastBin + 2
               frmPlot.picPlot.ScaleHeight = -1 * (PlotMax! - PlotMin!) '-4
               frmPlot.picPlot.ScaleTop = PlotMax! 'mvFirstBin + (mvFirstBin * 0.05)   'stLastBin + 2
            Else
               frmPlot.picPlot.ScaleHeight = -1 * (stLastBin - mvFirstBin) 'LastBin
               frmPlot.picPlot.ScaleTop = (stLastBin - mvFirstBin)  'stLastBin replaced LastBin
            End If
         Else
            Select Case stLastBin
               Case Is < 65537
                  mvDivisor = &H1
               Case Is < 1048577
                  mvDivisor = &H10
               Case Is < 16777217
                  mvDivisor = &H100
               'Case Is < 268435456
               '   mvDivisor = &H10000
               'Case Is < 2 ^ 31
               '   mvDivisor = &H100000
               Case Else
                  mvDivisor = &H100000
            End Select
            If NumBins& > 1 Then CompVal% = 1
            stScaleFactor! = (MaxBin - mvFirstBin) / mvDivisor
            'ToDo - add a scale buffer to evenly distribute data vertically
            If stScaleFactor! = 0 Then stScaleFactor! = 1
            'stScaleFactor! = (NumBins& - CompVal%) / mvDivisor
            frmPlot.picPlot.ScaleHeight = -1 * stScaleFactor!
            frmPlot.picPlot.ScaleTop = stScaleFactor! 'LastBin / mvDivisor
         End If
         'frmPlot.Caption = PlotMode$ & PlotMin! & " to " & PlotMax! & ")"
      Case 2   'fixed scale
         If stLastBin = mvFirstBin Then
            PlotMax! = frmPlot.picPlot.ScaleTop
            PlotMin! = PlotMax! + frmPlot.picPlot.ScaleHeight
         Else
            PlotMax! = stLastBin
            PlotMin! = mvFirstBin
         End If
         PlotMode$ = "Fixed scale ("
         'frmPlot.Caption = PlotMode$ & PlotMin! & " to " & PlotMax! & ")"
         mvDivisor = &H1
         Select Case stLastBin
            Case Is < 65537
               mvDivisor = &H1
            Case Is < 1048577
               mvDivisor = &H10
            Case Is < 16777217
               mvDivisor = &H100
            Case Else
               mvDivisor = &H10000
         End Select
   End Select
   Select Case mnPlotTitle
      Case 0
         If mnPlotType = HISTOGRAM Then
            frmPlot.Caption = PlotMode$ & BinCount& & "  Max bin: " & LargestBin & " @ " & LargestBinCode
         Else
            frmPlot.Caption = PlotMode$ & PlotMin! & " to " & PlotMax! & ")"
         End If
      Case 1
         frmPlot.Caption = mfOwner.Caption
   End Select
   
   'If (OriginalH! <> frmPlot.picPlot.ScaleHeight) Or (OriginalT! <> frmPlot.picPlot.ScaleTop) Then mnScaleChange = True
   mnScaleChange = True
   Exit Sub

EvalData:

   If mnChanSet < 0 Then
      If mnPlotType = HISTOGRAM Then
         FirstChan% = 0
         LastChan% = 0
      Else
         FirstChan% = 0
         LastChan% = mnLastChan
      End If
   Else
      FirstChan% = mnChanSet
      LastChan% = mnChanSet
   End If
   NumChans% = mnChansSelected '- 1   'mnLastChan + 1
   ReDim HistBins(0) As Long
   MinBin = 2 ^ 32 '65535
   MaxBin = -1 * 2 ^ 24
   NumBins& = 0
   BinCount& = 65535
   For CurChan% = 0 To NumChans% - 1
      Chan% = manChans(CurChan%)
      If mnPlotType = PLOT_CHUNKS Then
         NumberOfPoints = (mlNumChunks) '- 1  \ NumChans%
      Else
         Samps& = UBound(mvPlot, 2) - mlFirstPoint
         If mlEndBlock < Samps& Then Samps& = mlEndBlock
         NumberOfPoints = Samps& - (Samps& Mod NumChans%)
      End If
      FirstBin = mlFirstPoint
      'If mlStartBlock = 0 Then FirstBin = mlFirstPoint
      If ((NumberOfPoints + FirstBin) > mlPerChan) Then
      'If (mlEndBlock > NumberOfPoints) Then
         NumberOfPoints = NumberOfPoints - FirstBin
      End If
      'If (mlBlockSize < NumberOfPoints) Then
      '   NumberOfPoints = mlBlockSize
      'End If
      If NumberOfPoints > 0 Then BinCount& = BinData(Chan%, mvPlot, FirstBin, _
      LargestBin, LargestBinCode, HistBins(), Divisor%, Average, NumberOfPoints)
      If (mnTransferType = RealArray) Or mnConvToEng Or mnScaled Then
         'If Abs(LargestBin) > Abs(LastBin) Then LastBin = LargestBin
         'Spread = Abs(Average - LargestBinCode)
         'Temp = Average - Spread
         If (MinBin >= 2 ^ 32) Or (FirstBin < MinBin) Then MinBin = FirstBin
         'Temp = Average + Spread
         If (MaxBin <= (-1 * (2 ^ 24))) Or (LargestBin > MaxBin) Then MaxBin = LargestBin
      Else
         CurLast = BinCount& + FirstBin - 1
         If CurLast > LastBin Then LastBin = CurLast
         If (FirstBin < MinBin) Or (IsEmpty(MinBin)) Then MinBin = FirstBin
         If (LastBin > MaxBin) Or (IsEmpty(MaxBin)) Then MaxBin = LastBin
      End If
   Next CurChan%
   PlotBuf = 0.005
   NewBuf = (MaxBin - MinBin) * 0.2
   If (NewBuf > PlotBuf) Or (BinCount& = 1) Then PlotBuf = NewBuf
   MaxBin = MaxBin + PlotBuf
   MinBin = MinBin - PlotBuf
   If BinCount& > 65534 Then
      NumBins& = 65535
   Else
      NumBins& = MaxBin - MinBin
   End If
   If Not (NumBins& > 0) Then NumBins& = 1
   mvFirstBin = MinBin
   'LastBin = MaxBin
   Return
      
End Sub

Sub SetAutoScale(Optional ScriptSet)

   If IsMissing(ScriptSet) Then
      mnAutoScale = mnAutoScale + 1
      If mnAutoScale > 2 Then mnAutoScale = 0
   Else
      mnAutoScale = ScriptSet
   End If
   RePlot False

End Sub

Sub SetEndBlock(EndBlock As Long)

   If EndBlock = -1 Then
      mlEndBlock = mlBlockSize
   Else
      mlEndBlock = Abs(EndBlock)
   End If

End Sub

Sub SetBlockSize(BlockSize As Long, Store As Integer)

   If Not (mnPlotFormsLoaded > 0) Then frmPlot.Show
   DoEvents
   If Not mlFirstPoint = 0 Then Exit Sub
   If BlockSize = -1 Then
      'set back to default
      mlBlockSize = mlStaticBlockSize
      mlStartBlock = 0  'mlFirstPoint
      mlEndBlock = 0
      mnPlotFixedBlock = False
   ElseIf BlockSize < 0 Then
      If Not mnPlotFixedBlock Then mlStaticBlockSize = mlBlockSize
      mnPlotFixedBlock = True
      mlBlockSize = Abs(BlockSize)
   Else
      mlBlockSize = Abs(BlockSize)
      If Store Then
         mlStaticBlockSize = mlBlockSize
      Else
         If Not (mlStaticBlockSize = mlBlockSize) Then _
         frmPlot.mnuRestore.ENABLED = True
         If gnScriptRun Then mnPlotFixedBlock = True
      End If
      mlEndBlock = 0
      mlStartBlock = 0  'mlFirstPoint
   End If
   
   If (mnPlotType = SINGLE_POINT) Or (mnPlotType = PRINT_TEXT) Then
      'If UBound(mvPlot, 2) <> mlBlockSize Then ReDim mvPlotArray(8, mlBlockSize)
      mnDataInit = True
   End If
   If Store Then SaveBlockSize

End Sub

Sub SetCalData(BoardNum As Integer, Range As Integer, SetVal%)
   
   mnCalData = SetVal
   mnBoardNum = BoardNum
   mnRange = Range
   'SetConvertEval mnBoardNum, mnConvertData

End Sub

Sub SetChannel(Chan As Integer)

   PlotChans% = frmPlot.mnuOneChan.Count - 1
   If Chan > mnLastChan Then Exit Sub
   If mnDefaultWidth = 0 Then mnDefaultWidth = frmPlot.picPlot.ScaleWidth
   '-1 is all channels from menu
   If Chan = -2 Then
      If Not mnChansInitialized Then
         Chan = mnChanSet
         mnChansInitialized = True
      Else
         Exit Sub
      End If
   End If
   
   If (Chan = -1) Then  'Or (PlotChans% >= ChansSelected%)
      ReDim manChans(mnLastChan)
      For ChanMenu% = 0 To PlotChans%
         frmPlot.mnuOneChan(ChanMenu%).Checked = False
         manChans(ChanMenu%) = ChanMenu%
      Next
      Chan = -1
      ChansSelected% = mnLastChan + 1
   Else
      For ChanMenu% = 0 To PlotChans%
         If frmPlot.mnuOneChan(ChanMenu%).Checked Then
            If Chan = ChanMenu% Then
               If mnChansSelected > 1 Then frmPlot.mnuOneChan(ChanMenu%).Checked = False
            Else
               ReDim Preserve manChans(ChansSelected%)
               manChans(ChansSelected%) = ChanMenu%
               ChansSelected% = ChansSelected% + 1
            End If
         Else
            If ChanMenu% = Chan Then
               ReDim Preserve manChans(ChansSelected%)
               manChans(ChansSelected%) = ChanMenu%
               ChansSelected% = ChansSelected% + 1
               frmPlot.mnuOneChan(ChanMenu%).Checked = True
            End If
         End If
      Next
   End If
   frmPlot.mnuAllChans.Checked = False
   
   mnChanSet = Chan
   If mnChanSet = -1 Then
      frmPlot.mnuAllChans.Checked = True
   End If
   mnChansSelected = ChansSelected%

End Sub

Sub SetConvertData(BoardNum As Integer, SetVal As Integer)
   
   mnConvertData = SetVal
   mnBoardNum = BoardNum
   SetConvertEval mnBoardNum, mnConvertData

End Sub

Sub SetDetails(FunctionType As Integer, Param1 As Variant, Param2 As Variant, Param3 As Variant, Param4 As Variant, Param5 As Variant, Param6 As Variant, Param7 As Variant)

   Indent% = 16
   msDetail = mfOwner.lblStatus.Caption
   Select Case FunctionType
      Case AIn, AInScan, APretrig, ATrig, FileAInScan, FilePretrig, AIn32
         mnFirstChan = Val(Param1)
         If Param2 = "Q" Then
            ShowQ% = True
         Else
            LastChan% = Val(Param2)
         End If
         If Not ShowQ% Then mnRange = Val(Param5)
         If Not mfOwner.Caption = "Logger Functions" Then
            Range$ = Mid$(mfOwner.mnuRange.Caption, 9)
            Range$ = Left$(Range$, Len(Range$) - 1)
         End If
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         If mlSettleStep > 0 Then
            frmPlot.picDetails.Print "Points Avgd " & Param3
         Else
            frmPlot.picDetails.Print "TotalCount " & Param3
         End If
         
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         If ShowQ% Then
            frmPlot.picDetails.Print "ChanQ " & Param1
            frmPlot.picDetails.CurrentX = Indent%
            frmPlot.picDetails.Print "GainQ " & Param5
         Else
            frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
            frmPlot.picDetails.CurrentX = Indent%
            frmPlot.picDetails.Print "Range " & Range$
         End If
         If Not (FunctionType = AIn) Then
            frmPlot.picDetails.CurrentX = Indent%
            If Not mlSettleStep = 0 Then
               frmPlot.picDetails.Print Param4
            Else
               frmPlot.picDetails.Print "Rate " & Param4
            End If
         End If
         If (FunctionType = APretrig) Or (FunctionType = FilePretrig) Or (mlSettleStep > 0) Then
            frmPlot.picDetails.CurrentY = 1.8
            frmPlot.picDetails.CurrentX = Indent% + 14
            If mlSettleStep > 0 Then
               frmPlot.picDetails.Print Param7
            Else
               frmPlot.picDetails.Print "Pretrig Points " & Param7
            End If
         End If
      'Case ALoadQueue
      'Case APretrig
      'Case ATrig
      'Case FileAInScan
      'Case FilePretrig
      Case CIn, CIn32, CIn64
         msDetail = Switch(FunctionType = CIn, "cbCIn()", _
         FunctionType = CIn32, "cbCIn32()", FunctionType = CIn64, "cbCIn64()")
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param1
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Ctr List: " & Param2
      Case CInScan
         mnFirstChan = Val(Param1)
         LastChan% = Val(Param2)
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Format(Param7, "0")
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Rate " & Param4
      Case TIn, TInScan
         mnFirstChan = Val(Param1)
         If Param2 = "Q" Then
            ShowQ% = True
         Else
            LastChan% = Val(Param2)
         End If
         If Not ShowQ% Then mnRange = Val(Param5)
         If Not mfOwner.Caption = "Logger Functions" Then
            Range$ = Mid$(mfOwner.mnuRange.Caption, 9)
            Range$ = Left$(Range$, Len(Range$) - 1)
         End If
         'mnFirstChan = Val(Param1)
         'LastChan% = Val(Param2)
         If IsNumeric(Param5) Then
            RangeVal% = Val(Param5)
            ScaleString$ = Choose(RangeVal% + 1, "Celsius", "Fahrenheit", "Kelvin", "", "Volts", "No Scale")
         Else
            ScaleString$ = Param5
         End If
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param3
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         If ShowQ% Then
            frmPlot.picDetails.Print "ChanQ " & Param1
            frmPlot.picDetails.CurrentX = Indent%
            'frmPlot.picDetails.Print "GainQ " & Param5
         Else
            frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
            frmPlot.picDetails.CurrentX = Indent%
            'frmPlot.picDetails.Print "Range " & Range$
         End If
         'frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Scale " & ScaleString$
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Every " & Param4
         If VarType(Param7) = vbString Then
            frmPlot.picDetails.CurrentY = 1
            frmPlot.picDetails.CurrentX = Indent% + Indent%
            frmPlot.picDetails.Print Param7
         End If
      Case TInScan
      Case AOut
         mnFirstChan = Param1
         LastChan% = Param2
         mnRange = Param5
         Range$ = Mid$(mfOwner.mnuRange.Caption, 9)
         Range$ = Left$(Range$, Len(Range$) - 1)
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param3
         
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Range " & Range$
      Case AOutScan
         mnFirstChan = Val(Param1)
         If Param2 = "Q" Then
            ShowQ% = True
         Else
            LastChan% = Val(Param2)
         End If
         mnRange = Val(Param5)
         OwnerType$ = mfOwner.Tag
         If Not OwnerType$ = "" Then
            If Not Left(OwnerType$, 1) = "3" Then
               Range$ = Mid$(mfOwner.mnuRange.Caption, 9)
               Range$ = Left$(Range$, Len(Range$) - 1)
            End If
         End If
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param3
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         'frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
         'frmPlot.picDetails.CurrentX = Indent%
         'frmPlot.picDetails.Print "Range " & Range$
         If ShowQ% Then
            frmPlot.picDetails.Print "ChanQ " & Param1
            frmPlot.picDetails.CurrentX = Indent%
            frmPlot.picDetails.Print "GainQ " & Param5
         Else
            frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
            frmPlot.picDetails.CurrentX = Indent%
            frmPlot.picDetails.Print "Range " & Range$
         End If
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Rate " & Param4 '/ 1000 & "kHz"
      Case CIn
      Case CStoreOnInt
         'PortName$ = GetPortString(CInt(Param1))
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param1
      Case CFreqIn
      Case DInScan, DOutScan
         PortName$ = GetPortString(CInt(Param1))
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param2
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Port: " & PortName$
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Rate " & Param4 '/ 1000 & "kHz"
      Case VIn
         mnFirstChan = Val(Param1)
         LastChan% = Val(Param2)
         mnRange = Val(Param5)
         Range$ = Mid$(mfOwner.mnuRange.Caption, 9)
         Range$ = Left$(Range$, Len(Range$) - 1)
         frmPlot.picDetails.Cls
         frmPlot.picDetails.Print msDetail
         frmPlot.picDetails.Print Param6  'board name and number
         frmPlot.picDetails.Print "TotalCount " & Param3
         
         frmPlot.picDetails.CurrentY = 1
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Channels " & mnFirstChan & " to " & LastChan% + mnFirstChan
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print "Range " & Range$
         If Param4 = 0 Then
            RateString$ = "Rate = software loop"
         Else
            RateString$ = "Rate = " & Param4
            'TimerRate! = Param4 / 1000
            'ET! = Param3 * TimerRate!
            'RateString$ = "Every " & TimerRate! & " secs (" & ET! & " secs total)"
         End If
         frmPlot.picDetails.CurrentX = Indent%
         frmPlot.picDetails.Print RateString$

   End Select

End Sub

Sub SetFirstPoint(FirstPointVal As Long)

   If (Not gnScriptRun) And (FirstPointVal = 0) _
   Then mnPlotFixedBlock = False
   mlFirstPoint = FirstPointVal \ (mnLastChan + 1)
   'frmPlot.mnuRestore.Enabled = Not (mlFirstPoint = 0)
   'mlStartBlock = FirstPointVal '- mlBlockSize - 1
'   mlEndBlock = mlStartBlock + mlBlockSize
   'mlEndBlock = FirstPointVal + mlBlockSize - 1
   'SetBlock

End Sub

Sub SetManualScale(ScaleMin As Single, ScaleMax As Single)

   mnAutoScale = -1
   frmPlot.picPlot.ScaleHeight = -1 * (ScaleMax - ScaleMin)
   frmPlot.picPlot.ScaleTop = ScaleMax
   RePlot False

End Sub

Sub SetPlotOptions(PlotOptions As Integer, TitleType As Integer)

   mnRetainPlot = PlotOptions
   mnPlotTitle = TitleType
   'following 4 lines exist in case this function is
   'not called from the plot form
   frmPlot.mnuTitleType(0).Checked = False
   frmPlot.mnuTitleType(1).Checked = False
   frmPlot.mnuTitleType(mnPlotTitle).Checked = True
   frmPlot.mnuRetain.Checked = mnRetainPlot
   
   ScalePlot

End Sub

Sub SetPlotType(PlotType As Integer, Owner As Form)

   If Not (mfOwner Is Owner) Then ChangeDetails% = True
   Set mfOwner = Owner
   If mnLastChan < 0 Then Exit Sub
   If Not (mnPlotType = (PlotType And &HF)) Then mnRetainPlot = False
   mnPlotType = PlotType And &HF
   Select Case mnPlotType
      Case PRINT_LIST
         SetBlockSize 5000, False
         frmPlot.Show
         frmPlot.ShowText True
         'frmPlot.txtShow.Visible = True
         mnRetainPlot = False
      Case SINGLE_POINT
         mnRetainPlot = True
         'ReDim mvPlotArray(8, mlBlockSize)
         mnDataInit = True
         ReDim mnPointsPlotted(0)
         mnTransferType = -1
         mnLastChan = 0
         mnPointsPlotted(0) = -1
         ReDim mnChansLogged(0)
      Case Else
         If Not mnPlotFixedBlock Then
            If (mlStaticBlockSize > 0) And (mlBlockSize <> _
            (mlEndBuffer \ (mnLastChan + 1))) Then _
            SetBlockSize mlStaticBlockSize, False
            mlStartBlock = 0  'mlFirstPoint
         End If
         frmPlot.ShowText False
   End Select
   mnHardCopy = (PlotType And &HF0) = &H10
   If ChangeDetails% Then Set Owner = Nothing

End Sub

Sub SetPlotResolution(ResVal As Integer, Optional ManualSet As Variant, Optional DoReplot As Variant)

   Dim PlotNow As Boolean
   
   If IsMissing(ManualSet) Then
      If mbManualSet Then Exit Sub
   Else
      mbManualSet = ManualSet
   End If
   PlotNow = False
   If IsMissing(DoReplot) Then PlotNow = True
   
   mnResolution = ResVal
   If ResVal = 16 Then
      mn16Bit = True
   Else
      mn16Bit = False
   End If
   Index% = Switch(ResVal = 0, 1, ResVal = 4, 2, _
   ResVal = 8, 3, ResVal = 12, 4, ResVal = 13, 5, _
   ResVal = 14, 6, ResVal = 16, 7, ResVal = 24, 8, _
   ResVal = 32, 9, ResVal = 48, 10, ResVal = 64, 11)
   For FormNum% = 0 To Forms.Count - 1
      If Left$(Forms(FormNum%).Caption, 12) = "Analog Input" Then
         AInFormOpen% = True
         Exit For
      End If
   Next
   For i% = 0 To frmPlot.mnuResolution.Count - 1
      frmPlot.mnuResolution(i%).Checked = False
      If AInFormOpen% Then Forms(FormNum%).mnuResolution(i%).Checked = False
   Next i%
   frmPlot.mnuResolution(Index%).Checked = True
   If AInFormOpen% Then Forms(FormNum%).mnuResolution(Index%).Checked = True
   If ResVal = 0 Then mnAutoScale = 1
   If PlotNow And (Not mnUseText) Then RePlot False

End Sub

Sub SetSettleStep(SettleStep As Long, SettleRate As Double)

   mlSettleStep = SettleStep
   mlSettleRate = SettleRate

End Sub

Sub ShowText(ByVal TF As Integer)

   mnUseText = TF
   frmPlot.txtShow.Visible = TF
   If mnUseText And (mnPlotType = VOLTS_VS_TIME) Then
      If mlBlockSize > 500 Then SetBlockSize 500, False  'SetEndBlock 500 '
   Else
      mlFirstPoint = 0
   End If

End Sub

Sub TextList(ListString As String)

   If mnPlotFormsLoaded = 0 Then
      Load frmPlot
      'mnPlotFormsLoaded = True
   End If
   frmPlot.Visible = True
   msListString = ListString
   mnPlotType = PRINT_LIST
   IncrementBlock% = True
   InitBlock IncrementBlock%
   PrintList
   
End Sub

Sub TransferBlock()

   'for use only by 'Plot.frm'
   If mnPlotType = PRINT_LIST Then Exit Sub
   'If Not (mlBlockSize > 0) Then InitBlock
   DoEvents
   If (mnTransferType = FROMARRAY) Then
      IncrementBlock% = True
      Plot IncrementBlock%
      Exit Sub
   End If
   'to do - move following two lines to init?
   'If (mlBlockSize - 1) < mlArraySize Then mlArraySize = mlBlockSize - 1
   'If mlArraySize < 0 Then mlArraySize = 0
   FormType$ = Left(mfOwner.Caption, 7)
   FormPropsAvailable% = ((FormType$ = "Analog ") _
      Or (FormType$ = "Digital") _
      Or (FormType$ = "Counter") Or (FormType$ = "cbCInSc") Or (FormType$ = "cbCIn32"))
   If FormPropsAvailable% Then mnResolution = mfOwner.GetFormProperty("resolution")
   If (mnUse64 And mnUse32) Then
      If mnFloat Then
         ReDim DblArray(mnLastChan, mlPerChan) As Double
      Else
         If mnScaled Then
            ReDim sHexArray(mnLastChan, mlPerChan)
            DimensionTwo& = ((mlPerChan + 1) * 2) - 1
            ReDim LongArray(mnLastChan, DimensionTwo&) As Long
            ReDim PlotArray(mnLastChan, mlPerChan) As Variant
         Else
         End If
      End If
   ElseIf mnUse64 Then
      ReDim VarArray(mnLastChan, mlPerChan) As Currency
   ElseIf mnUse32 Then
      If mnFloat Then
         ReDim SnglArray(mnLastChan, mlPerChan) As Single
      Else
         ReDim LongArray(mnLastChan, mlPerChan) As Long
         ReDim DoubleArray(mnLastChan, mlPerChan) As Double
      End If
   Else
      If gnLibType = NETLIB Then
         DimSize& = ((mnLastChan + 1) * (mlPerChan + 1)) - 1
         ReDim NetArray(DimSize&) As Integer
      Else
         ReDim DatArray(mnLastChan, mlPerChan) As Integer
         If mnResolution > 12 Then _
            ReDim PlotArray(mnLastChan, mlPerChan) As Variant
      End If
   End If
   If mnTransferType = PARTBUFFER Then
      IncrementBlock% = False
      FirstPoint& = 0
      NumPoints& = mlArraySize
   Else
      FirstPoint& = mlStartBlock * (mnLastChan + 1)
      NumPoints& = mlEndBuffer
   End If
   Handle& = mvMemHandle
   
   Select Case mnTransferType
      Case BUFFER, PARTBUFFER, RealArray, PRETRIGBUFFER, FROMMESSAGE
         If Not ((mnTransferType = PRETRIGBUFFER) And mnConvertData) Then
            If (mnUse64 And mnUse32) Then
               If mnScaled Or mnConvBufToEng Or mnFloat Then
                  ULStat = LoadDblArrayFromWinBuf64(mfNoForm, mvMemHandle, DblArray(), FirstPoint&, NumPoints&)
               Else
                  ULStat = WBufToArray32(mvMemHandle, LongArray(0, 0), FirstPoint&, NumPoints& * 2)
                  If SaveFunc(mfNoForm, WinBufToArray32, ULStat, mvMemHandle, LongArray(0, 0), FirstPoint&, NumPoints& * 2, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               End If
            ElseIf mnUse64 Then
               ULStat = LoadArrayFromWinBuf64(mfNoForm, mvMemHandle, VarArray(), FirstPoint&, NumPoints&)
               If SaveFunc(mfNoForm, WinBufToArray64, ULStat, mvMemHandle, VarArray(0, 0), FirstPoint&, NumPoints& * 2, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            ElseIf mnUse32 Then
               msPreamble = ""
               If mnFloat Then
                  'changed the count on this to channel count for TInScan
                  'apparently not used by other functions - fix if not the case
                  CopyMemory SnglArray(0, 0), ByVal Handle&, (NumPoints&) * 4
                  'ULStat = WSngArrayToBuf(mfNoForm, Handle&, SnglArray(), NumPoints&)
               Else
                  
                  ULStat = LoadArrayFromWinBuf32(mfNoForm, Handle&, LongArray(), FirstPoint&, NumPoints&)
                  If Not (ULStat = 0) Then Exit Sub
               End If
            Else
               msPreamble = ""
               If gnLibType = NETLIB Then
                  ULStat = NetLibrary.LoadArrayFromWinBuffer(Handle&, NetArray(), FirstPoint&, NumPoints&)
               Else
                  ULStat = LoadArrayFromWinBuf(mfNoForm, Handle&, DatArray(), FirstPoint&, NumPoints&)
               End If
               If Not (ULStat = 0) Then Exit Sub
            End If
         End If
      Case FILE
         ULStat = cbFileRead(msFileName, FirstPoint&, NumPoints&, DatArray(0, 0))
         If SaveFunc(mfNoForm, FileRead, ULStat, msFileName, FirstPoint&, NumPoints&, DatArray(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case XMEM
         msPreamble = ""
         ULStat = cbMemRead(mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, MemRead, ULStat, mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case MEM_PRETRIG
         msPreamble = ""
         ULStat = cbMemReadPretrig(mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, MemReadPretrig, ULStat, mnMemBoard, DatArray(0, 0), FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case Else
         Exit Sub
   End Select
   If Not ULStat = 0 Then Exit Sub
   
   If mnConvertData Then
      ReDim ChanTags(mnLastChan, mlPerChan) As Integer
      If mnTransferType = PRETRIGBUFFER Then
         CreatePTBuffer
         NumPoints& = (mlArraySize - 1)
         mlPerChan = (mlArraySize / (mnLastChan + 1)) - 1
         ReDim DatArray(mnLastChan, mlPerChan) As Integer
         ULStat = LoadArrayFromWinBuf(mfNoForm, Handle&, DatArray(), FirstPoint&, mlArraySize)
      Else
         ULStat = cbAConvertData(mnBoardNum, NumPoints&, DatArray(0, 0), ChanTags(0, 0))
         If SaveFunc(mfNoForm, AConvertData, ULStat, mnBoardNum, NumPoints&, DatArray(0, 0), ChanTags(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      End If
   End If
   If mnCalData Then
      ULStat = cbACalibrateData(mnBoardNum, NumPoints&, mnRange, DatArray(0, 0))
      If SaveFunc(mfNoForm, ACalibrateData, ULStat, mnBoardNum, NumPoints&, mnRange, DatArray(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   End If
      
   'If Resolution% > 12 Then
   '   ReDim mvPlotArray(mnLastChan, mlPerChan)
   'End If

   If mnConvToEng Then
      If mnUse32 Then
         ReDim DblArray(mnLastChan, mlPerChan) As Double
         For Chan% = 0 To mnLastChan
            For Sample& = 0 To mlPerChan
               If mfTickSize = 0 Then
                  ULStat = cbToEngUnits32(mnBoardNum, mnRange, LongArray(Chan%, Sample&), dEngUnits#)
                  If SaveFunc(mfNoForm, ToEngUnits, ULStat, mnBoardNum, _
                     mnRange, LongArray(Chan%, Sample&), dEngUnits#, _
                     A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               Else
                  dEngUnits# = mfTickSize * LongArray(Chan%, Sample&)
               End If
               DblArray(Chan%, Sample&) = dEngUnits#
            Next Sample&
         Next Chan%
         mvPlot = DblArray()
      Else
         ReDim SnglArray(mnLastChan, mlPerChan) As Single
         For Chan% = 0 To mnLastChan
            For Sample& = 0 To mlPerChan
               If mfTickSize = 0 Then
                  ULStat = cbToEngUnits(mnBoardNum, mnRange, DatArray(Chan%, Sample&), EngUnits!)
                  If SaveFunc(mfNoForm, ToEngUnits, ULStat, mnBoardNum, _
                     mnRange, DatArray(Chan%, Sample&), EngUnits!, _
                     A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
               Else
                  EngUnits! = mfTickSize * DatArray(Chan%, Sample&)
               End If
               SnglArray(Chan%, Sample&) = EngUnits!
            Next Sample&
         Next Chan%
         mvPlot = SnglArray()
      End If
   ElseIf mnConvBufToEng Then
      'this is used in the .net library only
      NumChannels% = mnLastChan + 1
   Else
      If (mnUse64 And mnUse32) Then
         For Chan% = 0 To mnLastChan
            If mnFloat Then
               mvPlot = DblArray()
            Else
               For Sample& = 0 To (mlPerChan * 2) Step 2
                  sHexArray(Chan%, CombinedSample&) = _
                     Hex(LongArray(Chan%, Sample& + 1)) & Hex(LongArray(Chan%, Sample&))
                  LWord = LongArray(Chan%, Sample&)
                  If LWord < 0 Then LWord = LWord + 2 ^ 32
                  PlotArray(Chan%, CombinedSample&) = _
                     LWord + (LongArray(Chan%, Sample& + 1) * 2 ^ 32)
                  CombinedSample& = CombinedSample& + 1
               Next Sample&
               CombinedSample& = 0
            End If
         Next Chan%
         If mnScaled Then mnTransferType = RealArray
      ElseIf mnUse64 Then
         Dim CMult As Double
         ReDim TempVar(mnLastChan, mlPerChan) As Variant
         CMult = 10000
         For Sample& = 0 To mlPerChan
            TempVar(Chan%, Sample&) = VarArray(Chan%, Sample&) * CMult
         Next
         mvPlot = TempVar()
      ElseIf mnUse32 Then
         If mnFloat Then
            mvPlot = SnglArray()
            mnTransferType = RealArray
         Else
            For Ch% = 0 To mnLastChan
               For l& = 0 To mlPerChan
                  If LongArray(Ch%, l&) < 0 Then
                     DoubleArray(Ch%, l&) = 2 ^ 32 + LongArray(Ch%, l&)
                  Else
                     DoubleArray(Ch%, l&) = LongArray(Ch%, l&)
                  End If
               Next l&
            Next Ch%
            mvPlot = DoubleArray()
            'mnTransferType = FROMARRAY
         End If
      Else
         If gnLibType = NETLIB Then
            x& = NetIntArrayToULong(NetArray(), PlotArray(), mlPerChan, 0, mnLastChan)
         Else
            If mnResolution < 13 Then
               mvPlot = DatArray()
            Else
               x& = IntArrayToULong(DatArray(), PlotArray(), mlPerChan, 0)
               mvPlot = PlotArray()
            End If
         End If
      End If
   End If
   
   If mnPlotType = DERIVATIVE Then
      CalcDerivative
   End If
   
   mnDataInit = True
   Plot IncrementBlock%
   If mbEval Then RunDataEval mlPerChan, mlEvalChan
   
End Sub

Sub SetBufferType(BufType As Integer)

   BufRes% = BufType And &HF
   mnScaled = (BufType And &H10) = &H10
   mnFloat = (BufType And &H20) = &H20
   mnLibType = UNILIB
   If (BufType And &H100) = &H100 Then mnLibType = MSGLIB
   If (BufType And &H40) = &H40 Then mnTransferType = FROMMESSAGE
   If (BufType And &H80) = &H80 Then mnTransferType = BUFFER
   frmPlot.mnuHex.ENABLED = True
   Select Case BufRes%
      Case 1
         mnUse32 = False
         mnUse64 = False
      Case 2, 4
         mnUse32 = True
         mnUse64 = False
      Case 3
         mnUse64 = True
         mnUse32 = True
         gnHexVals = False
         frmPlot.mnuHex.ENABLED = False
      Case 5
         mnUse64 = True
         mnUse32 = False
   End Select

End Sub

Sub PlotTCValues(ByVal BoardNum As Integer, ChanArray() As Integer, TypeArray() As Integer, ByVal QCount As Long, ByVal mvHandle, ByVal FirstPoint&, ByVal NumPerChan&, ByVal TempScale&, GainArray() As Integer)

   If IsEmpty(mvHandle) Then Exit Sub
   If mvHandle = 0 Then Exit Sub
   If QCount < 1 Then Exit Sub
   For QElement% = 0 To QCount - 1
      If TypeArray(QElement%) = TC Then
         NumTCChans& = NumTCChans& + 1
      Else
         If Not (TypeArray(QElement%) = CJC) Then NumOther& = NumOther& + 1
      End If
   Next QElement%
   TotalPoints& = NumPerChan& * QCount
   NumPoints& = NumPerChan& - 1
   Chans& = NumTCChans& + NumOther&
   ReDim DatArray(QCount - 1, NumPoints) As Integer
   ReDim TempValArray!(NumTCChans& - 1, NumPoints)    '(should be bigger than it has to be) huh?
   ReDim AllValArray!(Chans& - 1, NumPoints)
   ULStat = IOTGetTCValues(BoardNum, ChanArray(0), TypeArray(0), QCount, mvHandle, FirstPoint&, NumPerChan&, TempScale&, TempValArray!(0, 0))
   If SaveFunc(mfNoForm, GetTCValues, ULStat, BoardNum, ChanArray(0), TypeArray(0), QCount, mvHandle, FirstPoint&, NumPerChan&, TempScale&, TempValArray!(0, 0), A10, A11, 0) Then Exit Sub
   ULStat = cbWinBufToArray(mvHandle, DatArray(0, 0), FirstPoint&, TotalPoints&)
   For SampNum& = 0 To NumPoints&
      For QElement% = 0 To QCount - 1
         If Not (TypeArray(QElement%) = CJC) Then
            ChanType% = TypeArray(QElement%)
            If ChanType% = TC Then
               AllValArray!(ArrayCount&, SampNum&) = TempValArray!(TCCount&, SampNum&)
               ArrayCount& = ArrayCount& + 1
               TCCount& = TCCount& + 1
            Else
               If QElement% > 0 Then
                  If (TypeArray(QElement% - 1) = CJC) Then OtherCount& = OtherCount& + 1
               End If
               Select Case ChanType%
                  Case ANALOG, ANALOG_SE, ANALOG_DIFF
                     CurRange& = GainArray(QElement%)
                     ULStat = cbToEngUnits(BoardNum, CurRange&, DatArray(QElement%, SampNum&), EngUnits!)
                     AllValArray!(ArrayCount&, SampNum&) = EngUnits!
                  Case Else
                     AllValArray!(ArrayCount&, SampNum&) = DatArray(QElement%, SampNum&)
               End Select
               ArrayCount& = ArrayCount& + 1
               OtherCount& = OtherCount& + 1
            End If
         End If
      Next QElement%
      ArrayCount& = 0
      OtherCount& = 0
      TCCount& = 0
   Next SampNum&
   x% = PlotRealArray(AllValArray!(), NumPoints)

End Sub

Sub SetContActive(ContState As Integer)

   mnContinuous = ContState

End Sub

Function PlotRealChunk(NumPoints As Long)

'   If mnPlotFormsLoaded = 0 Then
'      Load frmPlot
'      frmPlot.Visible = True
'   End If
   
'   If NumPoints = 0 Then RePlot
'   mnTransferType = RealArray
'   mlArraySize = UBound(DatArray, 2)
'   Chans% = UBound(DatArray, 1)
'   InitPlot Chans%
'   mnLastChan = Chans%
'   mlEndBuffer = (mlArraySize + 1) * (mnLastChan + 1)
'   InitBlock True
'   ReDim mvPlotArray(mnLastChan, mlPerChan)
'   mnDataInit = True
   
   mlNumChunks = NumPoints
'   For Ch% = 0 To mnLastChan
'      For l& = 0 To NumPoints
'         mvPlotArray(Ch%, l&) = DatArray(Ch%, l&)
'      Next l&
'   Next Ch%
'   ScalePlot
'   Plot

End Function

Public Function GetHistogram(Average As Variant, RMSVal As Variant, Samples As Long) As Long

   GetHistogram = mlNumBins
   Average = mvAverage
   RMSVal = mvLsbRms
   Samples = mlArraySize + 1
   
End Function

Public Function SaveDataAsText(Filename As String) As Integer

   On Error GoTo NoFile
   
   Open Filename For Output As #100
   TextString$ = "Sample" & ", "
   For Ch% = 0 To mnLastChan
      TextString$ = TextString$ & "ch" & Ch% & ", "
   Next Ch%
   Print #100, TextString$
   For i& = 0 To mlPerChan
      TextString$ = (i&) * (mnLastChan + 1) & ", "
      If mnTransferType = RealArray Then
         For Ch% = 0 To mnLastChan
            If Abs(mvPlot(Ch%, i&)) > 999 Then
               TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "+0.000000E+00;-0.000000E+00") & ", "
            Else
               TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "+000.00000000;-000.00000000") & ", "
            End If
         Next Ch%
      Else
         For Ch% = 0 To mnLastChan
            If gnHexVals Then
               TextString$ = TextString$ & "0x" & Hex$(mvPlot(Ch%, i&)) & ", "
            Else
               TextString$ = TextString$ & Format$(mvPlot(Ch%, i&), "0") & ", "
            End If
         Next Ch%
      End If
      Print #100, TextString$
   Next i&
   Close #100
   SaveDataAsText = True
   Exit Function
   
NoFile:
   MsgBox Error(Err), vbCritical, "Can't Create File"
   Exit Function
   
End Function

Sub SetBoardReference(BoardRef As Object)

   Set NetLibrary = BoardRef
   
End Sub

Public Sub GetTextRange(ByRef StartText As Long, ByRef TextCount As Long, ByRef MaxEnd As Long)

   StartText = mlStartText
   TextCount = mlTextCount
   MaxEnd = mlEndBuffer
   
End Sub

Public Sub SetTextRange(ByVal StartText As Long, ByVal TextCount As Long)

   mlStartText = StartText
   mlTextCount = TextCount
   mlEndBlock = mlStartText + mlTextCount

End Sub

Private Sub RunDataEval(ByVal EvalPoints As Long, Channel As Long)

   Dim OutWindow As Boolean
   Dim SamplePairCheck As Boolean
   
   If (mlEvalPoints > 0) And (mlEvalPoints < _
      (EvalPoints + 2)) Then EvalPoints = mlEvalPoints
   If mlEvalStart > mlArraySize Then
      MsgBox "Evaluation starting at " & _
         Format(mlEvalStart, "0") & " is not possible " & _
         "since only " & Format(mlArraySize, "0") & " samples " & _
         "are available.", vbInformation, "Evaluation Skipped"
      Exit Sub
   End If
   StartSamp& = mlEvalStart
   StopSamp& = StartSamp& + EvalPoints
   If StopSamp& > mlArraySize Then StopSamp& = mlArraySize
   
   ErrList$ = ""
   NumErrs& = 0
   frmPlot.lblAnalyzing.Visible = True
   DoEvents
   
   If mbDeltaMinEval Or mbDeltaMaxEval Then
      For DPoint& = StartSamp& To StopSamp& - 3 'Step 2
         'Diff& = Abs(mvPlot(Channel, DPoint& + 1) - mvPlot(Channel, DPoint&))
         OutWindow = False
         If (mbOutWindow Or mbInWindow) Then
            MaxCount = (2 ^ mnResolution)
            Window = MaxCount * (mlWindowDef / 100)
            OutWindow = False
            CurVal = mvPlot(Channel, DPoint&)
            OutWindow = (CurVal < Window) Or (CurVal > (MaxCount - Window))
         End If
         SamplePairCheck = False
         If Not OutWindow Then
            vDiff = Abs(mvPlot(Channel, DPoint& + 1) - mvPlot(Channel, DPoint&))
            vDiff2 = Abs(mvPlot(Channel, DPoint& + 2) - mvPlot(Channel, DPoint&))
            If mbPairCheck Then SamplePairCheck = (vDiff2 < mlEvalDeltaMinVal)
            If mbDeltaMaxEval Then
               If (vDiff > mlEvalDeltaVal) Then
                  TestCondition$ = " for delta greater than " & Format(mlEvalDeltaVal, "0")
                  ErrPoint$ = "  Diff = " & Format(vDiff, "0") & vbCrLf
                  ErrList$ = ErrList$ & Format(DPoint&, "0") & "; " & _
                     mvPlot(Channel, DPoint&) & vbCrLf & Format(DPoint& + 1, "0") & _
                     "; " & mvPlot(Channel, DPoint& + 1) & ErrPoint$ & vbCrLf
                  NumErrs& = NumErrs& + 1
                  LastErrSamp& = DPoint&
                  DoEvents
                  If NumErrs& > mlMsgBoxSamps Then EndIt% = True
               End If
            End If
            If mbDeltaMinEval Then
               If (vDiff < mlEvalDeltaMinVal) Or SamplePairCheck Then
                  TestCondition$ = " for delta less than " & Format(mlEvalDeltaMinVal, "0")
                  'If (vDiff2 < vDiff) Then vDiff = vDiff2
                  If (vDiff < mlEvalDeltaMinVal) Then ErrPoint$ = "  Diff = " & Format(vDiff, "0") & vbCrLf
                  If SamplePairCheck Then ErrPoint$ = "  Diff = " & Format(vDiff2, "0") & vbCrLf
                  ErrList$ = ErrList$ & Format(DPoint&, "0") & "; " & _
                     mvPlot(Channel, DPoint&) & vbCrLf & Format(DPoint& + 1, "0") & _
                     "; " & mvPlot(Channel, DPoint& + 1) '& ErrPoint$ & vbCrLf
                  If Not SamplePairCheck Then ErrList$ = ErrList$ & ErrPoint$ & vbCrLf
                  If SamplePairCheck Then ErrList$ = ErrList$ & vbCrLf & Format(DPoint& + 2, "0") & "; " & mvPlot(Channel, DPoint& + 2) & ErrPoint$ & vbCrLf
                  NumErrs& = NumErrs& + 1
                  LastErrSamp& = DPoint&
                  DoEvents
                  If NumErrs& > mlMsgBoxSamps Then EndIt% = True
               End If
            End If
            If EndIt% Then Exit For
         End If
      Next
   End If
   If mbRangeEval Then
      TestCondition$ = " for value > " & Format(mlEvalMaxVal, "0") & " or < " & Format(mlEvalMinVal, "0")
      For DPoint& = StartSamp& To StopSamp& - 1
         If mvPlot(Channel, DPoint&) > mlEvalMaxVal Then
            ErrPoint$ = "  Max = " & mvPlot(Channel, DPoint&) & vbCrLf
            ErrList$ = ErrList$ & Format(DPoint&, "0") _
               & "; " & ErrPoint$ & vbCrLf
            NumErrs& = NumErrs& + 1
            LastErrSamp& = DPoint&
            DoEvents
            If NumErrs& > mlMsgBoxSamps Then EndIt% = True
         End If
         If EndIt% Then Exit For
         If mvPlot(Channel, DPoint&) < mlEvalMinVal Then
            ErrPoint$ = "  Min = " & mvPlot(Channel, DPoint&) & vbCrLf
            ErrList$ = ErrList$ & Format(DPoint&, "0") _
               & "; " & ErrPoint$ & vbCrLf
            NumErrs& = NumErrs& + 1
            LastErrSamp& = DPoint&
            DoEvents
            If NumErrs& > mlMsgBoxSamps Then EndIt% = True
         End If
         If EndIt% Then Exit For
      Next
   End If
   If NumErrs& > 0 Then
      ErrTime$ = Now() & vbCrLf
      AnalogForm% = InStr(1, mfOwner.Caption, "Analog") > 0
      If AnalogForm% Then
         RateReturned! = mfOwner.GetFormProperty("rate")
         CurRate$ = Format(RateReturned! / 1000, "0.0###") & "kHz"
         AcqParams$ = "Acquired " & Format(mlArraySize + 1, "0") & _
            " samples at " & CurRate$ & " on channel " & Channel & vbCrLf
      End If
      EvalDesc$ = "Evaluating " & Format(EvalPoints, "0") & _
         " starting at sample " & Format(mlEvalStart, "0") & _
         TestCondition$ & "." & vbCrLf & vbCrLf
      DataStart& = LastErrSamp& - (mlMsgBoxSamps \ 2)
      If DataStart& < 0 Then DataStart& = 0
      DataEnd& = LastErrSamp& + (mlMsgBoxSamps \ 2)
      If DataEnd& > mlArraySize Then DataEnd& = mlArraySize
      If mbMsgBoxEval Then
         'gnCancel = True
         Success% = mfOwner.SetFormProperty("cancel", True)
         For DSamp& = DataStart& To DataEnd&
            SampList$ = SampList$ & Format(DSamp&, "0") _
               & "; " & mvPlot(Channel, DSamp&) & vbCrLf
         Next
         Clipboard.Clear
         Clipboard.SetText ErrTime$ & AcqParams$ & EvalDesc$ & ErrList$ & SampList$
         MsgBox Format(NumErrs&, "0") & _
         " data points available to paste.", _
         vbInformation, "Data Available"
         EndIt% = False
      End If
      If mbFileWrite Then
         frmPlot.MousePointer = vbHourglass
         frmPlot.picPlot.ENABLED = False
         frmScriptInfo.Show
         DataDir$ = Environ("TEMP")
         TimeCode$ = Format(Hour(Now), "00") & _
            Format(Minute(Now), "00") & _
            Format(Second(Now), "00")
         DataPath$ = DataDir$ & "\ErrorData" & TimeCode$ & ".txt"
         DoEvents
         Open DataPath$ For Output As #128
         If mlMsgBoxSamps = -1 Then
            FirstPoint& = mlEvalStart
            DataEnd& = mlArraySize
            LastPoint& = 10999
            If LastPoint& > DataEnd& Then LastPoint& = DataEnd&
            FullBar& = frmScriptInfo.picMasterStat.ScaleWidth
            Segment& = FullBar& / (mlArraySize / LastPoint&)
            frmScriptInfo.picMasterStat.Line (0, 0)-(FullBar&, 75), &H8000000F, BF
         Else
            FirstPoint& = DataStart&
            LastPoint& = FirstPoint& + 999
            If LastPoint& > DataEnd& Then LastPoint& = DataEnd&
         End If
         CurSamp& = 0
         Print #128, ErrTime$ & AcqParams$ & EvalDesc$ & ErrList$ & SampList$
         Do While CurSamp& < DataEnd&
            SampList$ = ""
            For CurSamp& = FirstPoint& To LastPoint&
               'SampList$ &
               SampList$ = Format(CurSamp&, "0") _
                  & ", " & mvPlot(Channel, CurSamp&) '& vbCrLf
               Print #128, SampList$
            Next
            'Print #128, SampList$
            FirstPoint& = CurSamp&
            LastPoint& = FirstPoint& + 10999
            If LastPoint& > DataEnd& Then LastPoint& = DataEnd&
            'frmScriptInfo.picMasterStat
            SegCnt& = SegCnt& + 1
            frmScriptInfo.picMasterStat.Line (0, 0)-(Segment& * SegCnt&, 75), &HFF0000, BF
            'frmPlot.picPlot.Line (0, 2040)-(Segment& * SegCnt&, 2048)
            DoEvents
         Loop
         Close #128
         If Dir(DataPath$, vbNormal) = "" Then
            frmScriptInfo.txtScriptInfo = "Attempt to write file to " _
               & DataPath$ & " failed."
         Else
            frmScriptInfo.txtScriptInfo = frmScriptInfo.txtScriptInfo _
               & DataEnd& & " samples saved to " & DataPath$ & "." & vbCrLf
         End If
         frmPlot.MousePointer = vbDefault
         frmPlot.picPlot.ENABLED = True
      End If
   End If
   frmPlot.lblAnalyzing.Visible = False
   DoEvents
   
End Sub

Public Sub SetEvalParams(ByVal Parameter As Long, ByVal ParamValue As Variant)

   If Parameter < 0 Then
      Select Case Parameter
         Case -1
            mbPairCheck = ParamValue
      End Select
   Else
      For i& = 0 To 18
         CurParam& = (Parameter And (2 ^ i&))
         Select Case CurParam&
            Case geENABLEEVAL
               mbEval = ParamValue
            Case geEVALDELTA
               mbDeltaEval = ParamValue
               'mbDeltaMinEval = False
               mbDeltaMaxEval = False
            Case geEVALRANGE
               mbRangeEval = ParamValue
            Case geEVALMINDELTA
               mbDeltaMinEval = ParamValue
            Case geEVALINWINDOW
               mbInWindow = ParamValue
            Case geEVALCHAN
               mlEvalChan = ParamValue
            Case geSTART
               mlEvalStart = ParamValue
            Case geNUMPOINTS
               mlEvalPoints = ParamValue
               frmPlot.lblAnalyzing.Caption = "Analyzing " _
                  & Format(mlEvalPoints, "0") & " samples..."
            Case geEVALCHAN
               mlEvalChan = ParamValue
            Case geDELTAVAL
               mlEvalDeltaVal = ParamValue
               mbDeltaMaxEval = True
            Case geMINVAL
               mlEvalMinVal = ParamValue
            Case geMAXVAL
               mlEvalMaxVal = ParamValue
            Case geDELTAMIN
               mlEvalDeltaMinVal = ParamValue
               'mbDeltaMinEval = True
            Case geINWINDOW
               mlWindowDef = ParamValue
            Case geSHOWPASTEMSG
               mbMsgBoxEval = ParamValue
            Case geNUMMSGSAMPS
               mlMsgBoxSamps = ParamValue
            Case geWRITEFILEERRS
               mbFileWrite = ParamValue
         End Select
      Next i&
   End If
   
End Sub

Public Function GetEvalParams(ByVal Parameter As Long) As Variant

   Dim EvalResult As Variant

   If Parameter < 0 Then
      Select Case Parameter
         Case -1
            EvalResult = mbPairCheck
      End Select
   Else
      Select Case Parameter
         Case geENABLEEVAL
            EvalResult = mbEval
         Case geEVALDELTA
            EvalResult = mbDeltaMaxEval   'mbDeltaEval
         Case geEVALRANGE
            EvalResult = mbRangeEval
         Case geEVALMINDELTA
            EvalResult = mbDeltaMinEval
         Case geEVALINWINDOW
            EvalResult = mbInWindow
         Case geSTART
            EvalResult = mlEvalStart
         Case geNUMPOINTS
            EvalResult = mlEvalPoints
         Case geEVALCHAN
            EvalResult = mlEvalChan
         Case geDELTAVAL
            EvalResult = mlEvalDeltaVal
         Case geMINVAL
            EvalResult = mlEvalMinVal
         Case geMAXVAL
            EvalResult = mlEvalMaxVal
         Case geDELTAMIN
            EvalResult = mlEvalDeltaMinVal
         Case geINWINDOW
            EvalResult = mlWindowDef
         Case geSHOWPASTEMSG
            EvalResult = mbMsgBoxEval
         Case geNUMMSGSAMPS
            EvalResult = mlMsgBoxSamps
         Case geWRITEFILEERRS
            EvalResult = mbFileWrite
      End Select
   End If
   GetEvalParams = EvalResult
   
End Function

Public Sub CalcDerivative(Optional DerivArray As Variant)

   If IsMissing(DerivArray) Then
      For Ch% = 0 To mnLastChan
         For l& = 1 To mlPerChan
            mvPlot(Ch%, l& - 1) = mvPlot(Ch%, l&) - mvPlot(Ch%, l& - 1)
         Next l&
         mvPlot(Ch%, l& - 1) = mvPlot(Ch%, l& - 2)
      Next Ch%
   Else
      DerivArray = mvPlot
   End If
   
End Sub
