Attribute VB_Name = "DataManager"
'general data generation and manipulation

'dependencies
'  cbw.dll, cbw.bas

'________________________________________

Global Const INVALIDBUFFER = 1
Global Const BUFFERTOOSMALL = 2
Global Const NOBUFFER = 3
Global Const ACQUIREDDATA = 1
Global Const GENERATEDDATA = 2

'Global MCCData%()

Dim mvHandle As Variant
Dim mlGenData As Long
Dim mlBufSize As Long, mlDataType As Long

Dim mfNoForm As Form
Dim mlNumCSVLines As Long

Function CheckBuffer(Handle As Variant, Size As Long) As Integer
   
   CheckBuffer = 0
   If WIN32APP Then
      If (VarType(Handle) = V_INTEGER) Then
         Choice = MsgBox("Cannot verify Windows buffer size because the " & _
            "handle is not 32 bit.  Read anyway?", 4, "Incompatible Handle")
         If Choice = 7 Then
            CheckBuffer = INVALIDBUFFER
            gnErrFlag = True
            Exit Function
         End If
      Else
         '32 bit lib returns a virtual memory address
         'BufferSize& = VirtualQuery(lpAddress, lpBuffer, dwLength&)
         Exit Function
      End If
   End If
   If Handle Then
      If Not ((GlobalFlags(Handle) And GMEM_DISCARDED) = 0) Then
         CheckBuffer = INVALIDBUFFER
      Else
         BufSize& = GlobalSize(Handle)
         If BufSize& < (Size * 2) Then CheckBuffer = BUFFERTOOSMALL
         Size = BufSize& / 2   'return size in integers
      End If
   Else
      CheckBuffer = NOBUFFER
   End If

End Function

Sub InitOutputBuffer(Handle As Variant, DataType As Integer, CBCount _
    As Long, Chans As Integer, Amplitude As Long, Offset As Long, UseLibrary As Integer)
   
   'UseLibrary has dual purpose - if 0, don't use cbWinBufAlloc to establish memory
   'Use the AllocateMemory() function instead -
   'if non-zero, use value to determine if DataList will be used
   'If UseLibrary is 1, use the list
   If CBCount < 1 Then Exit Sub
   DataChans = Array(0, 0, 0, 0, 0, 0, 0, 0)
   If UseLibrary = 1 Then UseAssignedChans% = GetDataList(DataChans)
   'set up a two dimensional local array
   If DataType = 6 Then 'use data from a previous A/D conversion
      BufSize& = CBCount
      BufferStat% = CheckBuffer(mvHandle, BufSize&)
      If (BufferStat% = INVALIDBUFFER) Or (BufferStat% = NOBUFFER) Then
         'invalid buffer, search for a valid one
         FindBuffer% = True
      Else
         Handle = mvHandle
         CBCount = BufSize&
      End If
      If FindBuffer% Then
         NumFuncsSaved% = GetHistory() - 1
         ReDim FuncsCalled(mnFunctionHistory - 1, 14)
         For FuncList% = 0 To NumFuncsSaved%
            If FuncsCalled(FuncList%, 0) = AInScan Then
               mvHandle = FuncsCalled(FuncList%, 8)
               Exit For
            End If
         Next FuncList%
         BufferStat% = CheckBuffer(mvHandle, BufSize&)
         If (BufferStat% = INVALIDBUFFER) Or (BufferStat% = NOBUFFER) Then
            'invalid buffer, search for a valid one
            MsgBox "Could not find a valid A/D data buffer.  " & _
                "Acquire data with A/D and try again.", , "No Buffer Found"
            Handle = 0
         Else
            Handle = mvHandle
            CBCount = BufSize&
         End If
      End If
      Exit Sub
   End If

   NumChans% = Chans + 1
   If NumChans% < 1 Then Exit Sub
   If (CBCount < 32768) Or WIN32APP Then
      CBCount = CBCount - (CBCount Mod NumChans%)
      TotalPerBlock& = CBCount
      If TotalPerBlock& = 0 Then TotalPerBlock& = 1
      PerChan& = CBCount / NumChans%
      BufferLoads% = 0
      ReDim LocArray(Chans, PerChan& - 1) As Integer
   Else
      Block% = CBCount \ 16384
      Extra% = (CBCount - (Block% * 16384&)) \ 6
      TotalPerBlock& = 16384 + Extra%
      CBCount = TotalPerBlock& * CLng(Block%)
      BufferLoads% = Block% - 1
      'CBCount = CBCount - (CBCount Mod NumChans%)
      PerChan& = TotalPerBlock& / NumChans%
      ReDim LocArray(Chans, PerChan& - 1) As Integer
   End If
   If CBCount = 0 Then CBCount = 1
   
   If Not (Handle = 0) Then
      If Not (UseLibrary = 0) Then
         ULStat = cbWinBufFree(Handle)
         If SaveFunc(mfNoForm, WinBufFree, ULStat, Handle, A2, _
            A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
         Handle = 0
         LongHandle& = mvHandle
         PlotBuffer LongHandle&, 0, mnLastChan - mnFirstChan
      Else
         If FreeMemory(Handle) Then Handle = 0
      End If
   End If
   
   If Not (UseLibrary = 0) Then
      Handle = cbWinBufAlloc(CBCount)
      If SaveFunc(mfNoForm, WinBufAlloc, Handle, CBCount, A2, _
        A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   Else
      'use this rather than cbWinBufAlloc() because the
      'latter does not initialize the buffer to zero
      'bug #408
      Handle = AllocateMemory(CBCount)
   End If
   If Handle = 0 Then Stop
   
   For Chan% = 0 To Chans
      ThisDataType% = DataType
      ThisChan% = Chan%
      If UseAssignedChans% Then
         If ThisChan% > 7 Then ThisChan% = Chan% - 8
         ThisDataType% = DataChans(ThisChan%)
      End If
      Select Case ThisDataType%
         Case 0   'send single values
            LocArray(Chan%, Sample&) = ULongValToInt(Amplitude)
         Case 1   'square wave
            'x% = 1
            For Sample& = 0 To PerChan& - 1 Step 2
               LongVal& = Offset + (Amplitude / 2) '* x% / 2
               LocArray(Chan%, Sample&) = ULongValToInt(LongVal&)
               LongVal& = Offset - (Amplitude / 2)
               If PerChan& < 2 Then Exit For
               LocArray(Chan%, Sample& + 1) = ULongValToInt(LongVal&)
               'x% = x% * -1
            Next
         Case 2   'sine wave
            'Base& = (Amplitude / 2) + (Offset / 2)
            For Sample& = 0 To PerChan& - 1
               If Chan% Mod 2 Then
                  Modulation& = Cos(Sample& / (PerChan& / 6.28)) * Amplitude / 2
                  If ((Chan% + 1) Mod 4) = 0 Then
                     LongVal& = Offset + Modulation&
                  Else
                     LongVal& = Offset - Modulation&
                  End If
               Else
                  Modulation& = Sin(Sample& / (PerChan& / 6.28)) * Amplitude / 2
                  If (Chan% Mod 4) = 0 Then
                     LongVal& = Offset + Modulation&
                  Else
                     LongVal& = Offset - Modulation&
                  End If
               End If
               LocArray(Chan%, Sample&) = ULongValToInt(LongVal&)
            Next
         Case 3   'ramp
            For Sample& = 0 To PerChan& - 1
               LongVal& = Offset - (Amplitude / 2) + Sample& / PerChan& * Amplitude
               LocArray(Chan%, Sample&) = ULongValToInt(LongVal&)
            Next
         Case 4   'triangle
            Element& = (PerChan& / 2)
            Direction% = 1
            For Sample& = 0 To PerChan& - 1 'Step 2
               LongVal& = Offset - (Amplitude / 2) + Element& / PerChan& * Amplitude
               If Element& + (Cycles * 2 * Direction%) > (PerChan&) - 1 Then Direction% = -1
               If Element& + (Cycles * 2 * Direction%) < 0 Then Direction% = 1
               Element& = Element& + (2 * Direction%) 'Cycles *
               LocArray(Chan%, Sample&) = ULongValToInt(LongVal&)
            Next
         Case 5   'cal voltages (-FS, x.., Midscale, x.., +FS)
            Increment& = (Amplitude / PerChan&)
            For Sample& = 0 To PerChan& - 1
               LongVal& = (Offset - Amplitude / 2 - Increment& / 2) _
                  + Increment& * (Sample& + 1) '/ PerChan& * Amplitude
               LocArray(Chan%, Sample&) = ULongValToInt(LongVal&)
            Next
         Case 6   'data from A/D conversion
      End Select
   Next Chan%
   For CurBlock& = 0 To BufferLoads%
      ULStat = cbWinArrayToBuf(LocArray(0, 0), Handle, _
         CurBlock& * TotalPerBlock&, TotalPerBlock&)
      If SaveFunc(mfNoForm, WinArrayToBuf, ULStat, LocArray(0, 0), _
         Handle, CurBlock& * TotalPerBlock&, TotalPerBlock&, _
         A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   Next CurBlock&

End Sub

Sub TransferData(mvMemHandle, MemBoard%, Filename$, FirstPoint&, NumPoints&, mnNumChans)

   PointsPerChan& = NumPoints& \ mnNumChans
   Chans% = mnNumChans - 1
   ReDim DatArray(Chans%, PointsPerChan&) As Integer
   If Not (Filename$ = "") Then
      mnTransferType = FILE
   ElseIf Not (mvMemHandle = 0) Then
      mnTransferType = BUFFER
   Else
      mnTransferType = XMEM
   End If
   
   Select Case mnTransferType
      Case BUFFER
         msPreamble = ""
         ULStat = cbWinBufToArray(mvMemHandle, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, WinBufToArray, ULStat, mvMemHandle, DatArray(0, 0), _
            FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case FILE
         ULStat = cbFileRead(Filename$, FirstPoint&, NumPoints&, DatArray(0, 0))
         If SaveFunc(mfNoForm, FileRead, ULStat, Filename$, FirstPoint&, NumPoints&, _
            DatArray(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case XMEM
         ULStat = cbMemRead(MemBoard%, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, MemRead, ULStat, MemBoard%, DatArray(0, 0), FirstPoint&, _
            NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case MEM_PRETRIG
         ULStat = cbMemReadPretrig(MemBoard%, DatArray(0, 0), FirstPoint&, NumPoints&)
         If SaveFunc(mfNoForm, MemReadPretrig, ULStat, MemBoard%, DatArray(0, 0), _
            FirstPoint&, NumPoints&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      Case Else
         Exit Sub
   End Select
   If mnConvertData Then
      ReDim ChanTags(Chans%, PointsPerChan&) As Integer
      ULStat = cbAConvertData(mnBoardNum, NumPoints&, DatArray(0, 0), ChanTags(0, 0))
      If SaveFunc(mfNoForm, AConvertData, ULStat, mnBoardNum, NumPoints&, _
         DatArray(0, 0), ChanTags(0, 0), A5, A6, A7, A8, A9, A10, A11, 0) _
         Then Exit Sub
   End If
   If mnConvToEng Then
      ReDim ReelArray!(Chans%, PointsPerChan&)
      For Chan% = 0 To Chans%
         For Sample% = 0 To PointsPerChan&
            ULStat = cbToEngUnits(mnBoardNum, mnRange, _
               DatArray(Chan%, Sample%), EngUnits!)
            If SaveFunc(mfNoForm, ToEngUnits, ULStat, mnBoardNum, mnRange, _
               DatArray(Chan%, Sample%), EngUnits!, A5, A6, A7, A8, _
               A9, A10, A11, 0) Then Exit Sub
            ReelArray!(Chan%, Sample%) = EngUnits!
         Next Sample%
      Next Chan%
   End If
   
End Sub

Function GetDataList(ByRef DataChans As Variant) As Integer

   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "ChanData"
   nSize% = 8
   lpKeyName$ = "Chan0"
   lpDefault$ = Format(DataChans(0))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(0) = Val(DataType$)
   
   lpKeyName$ = "Chan1"
   lpDefault$ = Format(DataChans(1))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(1) = Val(DataType$)
   
   lpKeyName$ = "Chan2"
   lpDefault$ = Format(DataChans(2))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(2) = Val(DataType$)
   
   lpKeyName$ = "Chan3"
   lpDefault$ = Format(DataChans(3))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(3) = Val(DataType$)
   
   lpKeyName$ = "Chan4"
   lpDefault$ = Format(DataChans(4))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(4) = Val(DataType$)
   
   lpKeyName$ = "Chan5"
   lpDefault$ = Format(DataChans(5))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(5) = Val(DataType$)
   
   lpKeyName$ = "Chan6"
   lpDefault$ = Format(DataChans(6))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(6) = Val(DataType$)
   
   lpKeyName$ = "Chan7"
   lpDefault$ = Format(DataChans(7))
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   DataChans(7) = Val(DataType$)
   
   lpKeyName$ = "UseChanList"
   lpDefault$ = "0"
   DataType$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
      lpKeyName$, lpDefault$, DataType$, nSize%, lpFileName$)
   DataType$ = Left$(DataType$, StringSize%)
   GetDataList = Val(DataType$)

End Function

Function GetBytesFromWinBuf(ByVal Handle As Long, ByVal DataType As Long, _
ByVal NumPoints As Long, NumChans As Integer, Bytes As Variant, _
Optional FirstPoint As Long) As Integer

   Dim IntSourceData() As Integer, LngSourceData() As Long
   Dim TempInt() As Integer
   
   If (NumPoints < 1) Or (NumChans < 1) Then Exit Function
   PerChan& = NumPoints \ NumChans
   'FirstPoint& = 0

   Select Case DataType
      Case vbInteger
         ReDim IntSourceData(NumChans - 1, PerChan& - 1)
         If (Not gbULLoaded) Then
            XFerPoints& = (NumPoints + FirstPoint&)
            ReDim TempInt(NumChans - 1, XFerPoints& - 1)
            'CopyMemory IntSourceData(0, 0), ByVal Handle, NumPoints * 2
            CopyMemory TempInt(0, 0), ByVal Handle, XFerPoints& * 2
            IntSourceData(0, 0) = TempInt(0, FirstPoint&)
            ErrCode& = 0
         Else
            'ULStat = cbWinBufToArray(Handle, nDatArray(0, 0), FirstPoint, Samples)
            ErrCode& = cbWinBufToArray(Handle, IntSourceData(0, 0), FirstPoint&, NumPoints)
         End If
         Bytes = IntSourceData()
      Case vbLong
         ReDim LngSourceData(NumChans - 1, PerChan& - 1) As Long
         If (Not gbULLoaded) Or ForceWinAPI Then
            CopyMemory LngSourceData(0, 0), ByVal Handle, NumPoints * 4
            ErrCode& = 0
         Else
            'ULStat = WBufToArray32(Handle, lDatArray(0, 0), FirstPoint, Samples)
            ErrCode& = cbWinBufToArray32(Handle, LngSourceData(0, 0), FirstPoint&, NumPoints)
         End If
         Bytes = LngSourceData()
      Case vbSingle
         ReDim SngSourceData(NumChans - 1, PerChan& - 1) As Single
         CopyMemory SngSourceData(0, 0), ByVal Handle, NumPoints * 4
         Bytes = SngSourceData()
      Case vbDouble
         ReDim DblSourceData(NumChans - 1, PerChan& - 1) As Double
         CopyMemory DblSourceData(0, 0), ByVal Handle, NumPoints * 8
         Bytes = DblSourceData()
      Case vbDecimal
         ReDim CurSourceData(NumChans - 1, PerChan& - 1) As Currency
         CopyMemory CurSourceData(0, 0), ByVal Handle, NumPoints * 8
         Bytes = CurSourceData()
   End Select
   If Not ErrCode& = 0 Then Exit Function
   GetBytesFromWinBuf = True

End Function

Public Function GenerateData(DataType As Long, Cycles As Integer, CBCount As Long, _
NumberOfChans As Integer, Amplitude As Variant, Offset As Variant, SignalParams As Integer, _
Optional NewData As Integer = True, Optional Channel As Integer = -1, _
Optional FirstPoint As Long = 0, Optional UseWinAPI As Integer = 0) As Long

   'Returns:   handle to data
   'DataType:  1 = Integer (signed 16-bit), 2 = Long (signed 32-bit),
   '           3 = Decimal (signed 96-bit), 4 = Single
   '           5 = handle to existing buffer for modification
   '           6 = Double
   
   Dim DataVal As Variant
   Dim IntegerArray() As Integer
   Dim LongArray() As Long
   Dim SnglArray() As Single
   Dim DblArray() As Double
   Dim NoForm As Form
   
   SignalType% = SignalParams And &HF
   DataFormat% = SignalParams And &HF0
   If Not (CBCount > 0) Then
      'get rid of the handle to the existing data buffer
      If Not (mlGenData = 0) Then
         
         If Not BufFree(NoForm, mlGenData, UseWinAPI) Then Exit Function
         mlGenData = 0: mlBufSize = 0
      End If
      GenerateData = 0
      Exit Function
   End If
   
   If NewData Then
      BufferSize& = CBCount
      mlDataType = DataType
   Else
      If (mlBufSize = CBCount) Then CBCount = CBCount - 1
      If ((CBCount + FirstPoint) > mlBufSize) Then
         MsgBox "Number of data points and starting point specified " & _
         "results in points beyond the end of the existing buffer.", _
         vbCritical, "Invalid Modification of Existing Data"
         Exit Function
      End If
      BufferSize& = mlBufSize
      DataType = mlDataType
   End If
   PerChan& = CBCount / NumberOfChans
   If PerChan& = 0 Then PerChan& = 1
   BufPerChan& = BufferSize& / NumberOfChans
   If BufferSize& < NumberOfChans Then
      MsgBox "Insufficient samples to fill channels. Requesting " & _
      Format(NumberOfChans, "0") & " channels but only " & _
      Format(BufferSize&, "0") & " samples.", vbOKOnly, "Data Generation Error"
      Exit Function
   End If

   Select Case DataType
      Case 1
         ReDim IntegerArray(NumberOfChans - 1, BufPerChan& - 1)
         Use32% = False
      Case 2
         ReDim LongArray(NumberOfChans - 1, BufPerChan& - 1)
         Use32% = True
      Case 4
         ReDim SnglArray(NumberOfChans - 1, BufPerChan& - 1)
      Case 6
         ReDim DblArray(NumberOfChans - 1, BufPerChan& - 1)
         Use64% = True
   End Select
   
   If NewData Then
      'get rid of the handle to the existing data buffer
      If Not (mlGenData = 0) Then
         If Not BufFree(NoForm, mlGenData, UseWinAPI) Then Exit Function
         mlGenData = 0: mlBufSize = 0
      End If
   Else
      'use existing buffer for modification
      Select Case DataType
         Case 1
            MemResult& = LoadArrayFromWinBuf(NoForm, mlGenData, _
               IntegerArray(), 0, BufferSize&, UseWinAPI)
            If Not MemResult& = 0 Then Exit Function
         Case 2
            ULStat = cbWinBufToArray32(mlGenData, LongArray(0, 0), 0, BufferSize&)
            If SaveFunc(mfNoForm, WinBufToArray32, ULStat, mlGenData, LongArray(0, 0), 0, _
               BufferSize&, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
         Case 4
            CopyMemory SnglArray(0, 0), ByVal mlGenData, BufferSize& * 4
         Case 6
            CopyMemory DblArray(0, 0), ByVal mlGenData, BufferSize& * 8
      End Select
   End If
   
   If (Channel = -1) Then
      FirstChan% = 0
      LastChan% = NumberOfChans - 1
   Else
      FirstChan% = Channel
      LastChan% = (Channel + CBCount) - 1
   End If
   Direction% = 1
   If (Cycles = 0) Then
      TogglePoint& = 0
   Else
      TogglePoint& = ((PerChan& / Cycles) / 2) - 1
   End If
   
   For Chan% = FirstChan% To LastChan%
      For Sample& = FirstPoint To (FirstPoint + PerChan&) - 1
         Select Case SignalType
            Case 0   'send single values
               DataVal = Amplitude
            Case 1   'square wave
               DataVal = Offset + ((Amplitude / 2) * Direction%)
               Element& = Element& + 1
               If Element& > TogglePoint& Then
                  Direction% = Direction% * -1
                  Element& = 0
               End If
            Case 2   'sine wave
               If Chan% Mod 2 Then
                  Modulation = Cos(Sample& * Cycles / (PerChan& / 6.28)) * Amplitude / 2
                  If ((Chan% + 1) Mod 4) = 0 Then
                     DataVal = Offset + Modulation
                  Else
                     DataVal = Offset - Modulation
                  End If
               Else
                  Modulation = Sin(Sample& * Cycles / (PerChan& / 6.28318530718)) * Amplitude / 2
                  If (Chan% Mod 4) = 0 Then
                     DataVal = Offset + Modulation
                  Else
                     DataVal = Offset - Modulation
                  End If
               End If
            Case 3   'ramp
               ShiftVal = (Amplitude / (FirstPoint + PerChan&) - FirstPoint) / 2
               DataVal = Offset - (Amplitude / 2) + Element& * Cycles / PerChan& * Amplitude
               DataVal = DataVal + ShiftVal
               Element& = Element& + 1
               If Element& * Cycles > (FirstPoint + PerChan&) - 1 Then
                  Element& = 0
               End If
            Case 4   'triangle
               If Sample& = 0 Then Element& = (PerChan& / 2)
               DataVal = Offset - (Amplitude / 2) + Element& / PerChan& * Amplitude
               If Element& + (Cycles * 2 * Direction%) > (PerChan&) - 1 Then Direction% = -1
               If Element& + (Cycles * 2 * Direction%) < 0 Then Direction% = 1
               Element& = Element& + (Cycles * 2 * Direction%)
            Case 5   'random
               Randomize
               If Amplitude = 0 Then
                  MsgBox "Random signal cannot have 0 amplitude.", vbCritical, "Cannot Generate Data"
                  Exit Function
               End If
               DataVal = Offset + (Amplitude / 2 - (Amplitude * Rnd(Timer)))
               If DataVal = Amplitude Then DataVal = DataVal - 1
            Case 6   'cal voltages (-FS, x.., Midscale, x.., +FS)
               Increment& = (Amplitude / PerChan&)
               DataVal = (Offset - Amplitude / 2 - Increment& / 2) + Increment& * (Sample& + 1) '/ PerChan& * Amplitude
            Case 7   'data from A/D conversion
         End Select
         If Not (SignalType = 7) Then
            Select Case DataType
               Case 1
                  IntegerArray(Chan%, Sample&) = ConvertDataType(DataType, DataVal)
               Case 2
                  LongArray(Chan%, Sample&) = ConvertDataType(DataType, DataVal)
               Case 4
                  SnglArray(Chan%, Sample&) = DataVal
               Case 6
                  DblArray(Chan%, Sample&) = DataVal
            End Select
         End If
      Next Sample&
      Element& = 0
   Next Chan%
   
   Select Case DataType
      Case 1
         If mlGenData = 0 Then
            mlGenData = BufAlloc16(NoForm, BufferSize&, UseWinAPI)
            If mlGenData = 0 Then Exit Function
         End If
         MemResult& = WArrayToBuf(NoForm, mlGenData, IntegerArray(), BufferSize&, UseWinAPI)
      Case 2
         If mlGenData = 0 Then
            mlGenData = BufAlloc32(NoForm, BufferSize&, UseWinAPI)
            If mlGenData = 0 Then Exit Function
         End If
         Dummy& = WArrayToBuf32(NoForm, mlGenData, LongArray(), BufferSize&, UseWinAPI)
      Case 4
         If mlGenData = 0 Then
            mlGenData = BufAlloc32(NoForm, BufferSize&, UseWinAPI)
            If mlGenData = 0 Then Exit Function
         End If
         Dummy& = WArrayToBuf32(NoForm, mlGenData, LongArray(), BufferSize&, UseWinAPI)
      Case 6
         If mlGenData = 0 Then
            mlGenData = ScaledBufAlloc(NoForm, BufferSize&, UseWinAPI)
            If mlGenData = 0 Then Exit Function
         End If
         ULStat = WDblArrayToBuf(NoForm, mlGenData, DblArray(), BufferSize&, UseWinAPI)
   End Select
   GenerateData = mlGenData
   If NewData Then mlBufSize = BufferSize&

End Function

Function ConvertDataType(ToType As Long, DataVal As Variant) As Variant

   'ToType:  1 = Integer (signed 16-bit), 2 = Long (signed 32-bit),
   '         3 = Decimal (signed 96-bit), 4 = Single

   Select Case ToType
      Case 1
         Select Case DataVal
            'remove decimal points
            Case Is > 65535
               ConvertResult = -1
            Case Is < -32768
               ConvertResult = 0
            Case Else
               IntData& = Fix(DataVal)
               If IntData& > 32767 Then
                  ConvertResult = IntData& - 65536
               Else
                  ConvertResult = CInt(IntData&)
               End If
         End Select
         ConvertDataType = CInt(ConvertResult)
      Case 2
         Select Case DataVal
            Case Is > 4294967295#
               ConvertResult = -1
            Case Is < -2147483648#
               ConvertResult = 0
            Case Else
               If Fix(DataVal) > 2147483647 Then
                  ConvertResult = Fix(DataVal) - 4294967296#
               Else
                  ConvertResult = Fix(DataVal)
               End If
         End Select
         ConvertDataType = CLng(ConvertResult)
      Case 3
         ConvertResult = DataVal
         ConvertDataType = CDec(Fix(ConvertResult))
      Case 4
         'not implemented
   End Select

End Function

Public Function GetStringValueAsUType(ByVal DataValue As String, ByRef Resolution As Integer) As String

   Dim NumericValue As Variant
   Dim LNumericValue As Variant
   Dim DataSize As Integer
   Dim StringValue As String
   
   StringValue = DataValue
   If Not (InStr(1, DataValue, "^") = 0) Then
      Notation = Split(DataValue, "^")
      Mantissa$ = Notation(0)
      SecondPart$ = Notation(1)
      ExponentString$ = SecondPart$
      AdjustBy& = 0
      MathPart = Split(SecondPart$, "+")
      If Not (MathPart(0) = SecondPart$) Then
         ExponentString$ = MathPart(0)
         AdjustBy& = Val(MathPart(1))
      Else
         MathPart = Split(SecondPart$, "-")
         If Not (MathPart(0) = SecondPart$) Then
            AdjustBy& = Val(MathPart(1)) * -1
            ExponentString$ = MathPart(0)
         End If
      End If
      Mant& = Val(Mantissa$)
      Expnt& = Val(ExponentString$)
      NumericValue = Mant& ^ Expnt& + AdjustBy&
      StringValue = Format(NumericValue, "0")
   End If
   
   StringSize% = Len(StringValue)
   DataSize = Resolution
   Select Case StringSize%
      Case Is < 5
         If Not (Resolution > 16) Then DataSize = 16
      Case Is < 10
         If Not (Resolution > 32) Then DataSize = 32
      Case Is <= 19
         DataSize = 64
      Case Is > 19
         If (StringSize% = 20) Then
            'If (Left(StringValue, 1) = "-") Then
            'Else
            'End If
            DataSize = 64
         Else
            DataSize = 0
         End If
   End Select
   
   Select Case DataSize
      Case 0
         Resolution = 0
         DType$ = "all data types."
      Case 16
         NumericValue = Val(StringValue)
         DataAsTypeString$ = Format(NumericValue, "0")
      Case 32
         NumericValue = Val(StringValue)
         If Resolution < 17 Then
            If NumericValue > 65535 Then
               Resolution = 0
               DType$ = "16-bit."
            Else
               If NumericValue > 32767 Then
                  NumericValue = Val(StringValue) - 65536
               End If
            End If
         End If
         DataAsTypeString$ = Format(NumericValue, "0")
      Case 64
         NumericValue = CDec(StringValue)
         SignString$ = ""
         If Resolution < 33 Then
            If NumericValue > 4294967295# Then
               Resolution = 0
               DType$ = "32-bit."
            Else
               If NumericValue > 2147483647 Then
                  NumericValue = Val(StringValue) - 4294967296#
               End If
            End If
            DataAsTypeString$ = Format(NumericValue, "0")
         Else
            SignChange = CDec(2 ^ 32) * CDec(2 ^ 31) - 1
            If NumericValue > SignChange Then
               NumericValue = CDec(NumericValue) - ((SignChange + 1) * 2) '- 1
               StringValue = Format(Abs(NumericValue), "0")
               StringSize% = Len(StringValue)
               SignString$ = "-"
            End If
            If StringSize% > 4 Then
               UCurrencyString$ = Left(StringValue, Len(StringValue) - 4)
               LCurrencyString$ = Right(StringValue, 4)
               NumericValue = Val(SignString$ & UCurrencyString$)
               LNumericValue = Val(LCurrencyString$)
            Else
               UCurrencyString$ = "0"
               LCurrencyString$ = StringValue
               NumericValue = 0
            End If
            If Abs(NumericValue) = 922337203685477# Then
               'check lower value fits
               If NumericValue < 0 Then
                  If LNumericValue > 5808 Then
                     Resolution = 0
                     DType$ = "64-bit."
                  End If
               Else
                  If LNumericValue > 5807 Then
                     Resolution = 0
                     DType$ = "64-bit."
                  End If
               End If
            End If
            DataAsTypeString$ = SignString$ & UCurrencyString$ & "." & LCurrencyString$
         End If
   End Select
   If Resolution = 0 Then DataAsTypeString$ = "Value outside the range of " & DType$
   GetStringValueAsUType = DataAsTypeString$
   
End Function

Public Function GetStringValueFromCur(ByVal DataValue As Currency) As String

   LeftOfDecimal = Fix(DataValue)
   RightOfDecimal = DataValue - LeftOfDecimal
   LoDString$ = Format(LeftOfDecimal, "0")
   RoDString$ = Format(RightOfDecimal, "0.0000")
   CurString$ = LoDString$ & Right(RoDString$, 4)
   GetStringValueFromCur = CurString$

End Function

Public Function GetHexValue(ByVal DataValue As Variant, ByVal Resolution As Integer) As String
   
   Dim Segment() As Variant
   Dim Bit64 As Boolean
   
   NewData = DataValue
   If DataValue < 0 Then
      Select Case Resolution
         Case 64
            NewData = CDec(2 ^ 32) * CDec(2 ^ 32) + DataValue
            TotalLen% = 16
            Segment() = Array(6, 2, 6, 2)
         Case 48
            TotalLen% = 12
         Case 32
            NewData = 2 ^ 32 + DataValue
         Case 16
            NewData = 2 ^ 16 + DataValue
      End Select
   End If
   
   LoopVal = 0
   Val1% = 1
   Val2% = 0
   Val3% = 0
   EvalSegment = NewData
   If Fix(EvalSegment / (CDec(2 ^ 32) * CDec(2 ^ 31))) > 0 Then
      ShiftData = EvalSegment - (CDec(2 ^ 32) * CDec(2 ^ 31))
      HalfData = NewData - (Fix(NewData / CDec(2 ^ 32)) * CDec(2 ^ 32))
      EvalSegment = ShiftData
      Bit64 = True
   End If
   
   If Fix(EvalSegment / (CDec(2 ^ 31) * CDec(2 ^ 29))) > 0 Then
      If Bit64 Then
         If (HalfData > 0) Then  'And (Fix(HalfData / 2 ^ 40) > 0)
            Val1% = 6
            Val2% = 10
            Val3% = 0
         Else
            Val1% = 6
            Val2% = 6 '10
            Val3% = 4
         End If
      Else
         Val1% = 8
         Val2% = 6
         Val3% = 2
      End If
   ElseIf Fix(EvalSegment / (CDec(2 ^ 31) * CDec(2 ^ 25))) > 0 Then
      Val1% = 7
      Val2% = 6
      Val3% = 2
   ElseIf Fix(EvalSegment / (CDec(2 ^ 31) * CDec(2 ^ 24))) > 0 Then
      Val1% = 6
      Val2% = 6
      Val3% = 2
   ElseIf Fix(EvalSegment / (CDec(2 ^ 31) * CDec(2 ^ 21))) > 0 Then
      Val1% = 8
      Val2% = 6
   ElseIf Fix(EvalSegment / 2 ^ 48) > 0 Then
      Val1% = 7
      Val2% = 6
   ElseIf Fix(EvalSegment / 2 ^ 47) > 0 Then
      Val1% = 6
      Val2% = 6
   ElseIf Fix(EvalSegment / 2 ^ 44) > 0 Then
      Val1% = 8
      Val2% = 4
   ElseIf Fix(EvalSegment / 2 ^ 40) > 0 Then
      If Bit64 Then
         Val1% = 6
         Val2% = 6 '10
         Val3% = 0
      Else
         Val1% = 7
         Val2% = 4
      End If
   ElseIf Fix(EvalSegment / 2 ^ 39) > 0 Then
      If Bit64 Then
         Val1% = 6
         Val2% = 6
         Val3% = 4
      Else
         Val1% = 6
         Val2% = 4
      End If
   ElseIf Fix(EvalSegment / 2 ^ 36) > 0 Then
      Val1% = 8
      Val2% = 2
   ElseIf Fix(EvalSegment / 2 ^ 32) > 0 Then
      Val1% = 7
      Val2% = 2
   ElseIf Fix(EvalSegment / 2 ^ 31) > 0 Then
      Val1% = 6
      Val2% = 2
   End If
   Segment = Array(Val1%, Val2%, Val3%)
   TotalLen% = Val1% + Val2% + Val3%
   
   If (NewData < 0) And (Resolution > 16) Then LoopVal = 1
   FlipLimit% = Resolution
   If Resolution > 32 Then FlipLimit% = 32
   SignFlip = CDec((2 ^ FlipLimit%) / 2) - 1
   
   For Iteration& = 0 To LoopVal
      VarValue = NewData
      If Resolution > 16 Then
         If (Iteration& = 0) And (VarValue < 0) Then
            TempVal = CDec(2 ^ (Resolution / 2) + VarValue)
            If TempVal < 0 Then
               VarValue = NewData
            Else
               VarValue = 2 ^ (Resolution / 2) - 1
            End If
         ElseIf VarValue < 0 Then
            VarValue = 2 ^ (Resolution / 2) + VarValue
         End If
      Else
         If VarValue < 0 Then
            VarValue = (2 ^ Resolution) + VarValue
         End If
      End If
      Do
         ShiftExp& = 0
         Shifter = CDec(2 ^ ShiftExp&)
         Do
            Shifter = CDec(2 ^ ShiftExp&)
            TempVal = Fix(VarValue / Shifter)
            ShiftExp& = ShiftExp& + 8
         Loop While TempVal > SignFlip
         PartialVal& = Fix(TempVal)
         StringVal$ = Hex(PartialVal&)
         SegLen% = Segment(CurSegment%)
         Prepend$ = ""
         VarValue = VarValue - (PartialVal& * Shifter)
         ExpectedLen% = SegLen% - Len(StringVal$)
         If (ExpectedLen% > 0) Then Prepend$ = String(ExpectedLen%, "0")
         StringVal$ = Prepend$ & StringVal$
         TextValue$ = TextValue$ & StringVal$
         CurSegment% = CurSegment% + 1
      Loop While VarValue > 0
   Next
   ExpectedLen% = TotalLen% - Len(TextValue$)
   If ExpectedLen% > 0 Then
      Append$ = String(ExpectedLen%, "0")
      TextValue$ = TextValue$ & Append$
   End If
   GetHexValue = TextValue$
   If Len(TextValue$) < TotalLen% Then Stop

End Function

Public Function GetCurBufferSize() As Long

   GetCurBufferSize = mlBufSize
   
End Function

Public Function BufFree(ByVal CallingForm As Form, _
ByVal BufHandle As Long, Optional ForceWinAPI As Integer) As Integer

   If (Not gbULLoaded) Or ForceWinAPI Then
      Handle = BufHandle
      result% = FreeMemory(BufHandle)
      BufFree = result%
   Else
      ULStat = cbWinBufFree(BufHandle)
      x% = SaveFunc(CallingForm, WinBufFree, ULStat, BufHandle, _
         A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      BufFree = (ULStat = 0)
   End If
   
End Function

Public Function BufAlloc16(CallingForm As Form, ByVal BufferSize As Long, _
Optional ByVal ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      Handle& = AllocateMemory(BufferSize)
   Else
      Handle& = cbWinBufAlloc(BufferSize)
      x% = SaveFunc(CallingForm, WinBufAlloc, Handle&, BufferSize, _
         A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   End If
   BufAlloc16 = Handle&
   
End Function

Public Function BufAlloc32(CallingForm As Form, ByVal BufferSize As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      'AllocateMemory doubles BufferSize again, so * 4 is resulting size
      Handle& = AllocateMemory(BufferSize * 2)
   Else
      Handle& = cbWinBufAlloc32(BufferSize)
      x% = SaveFunc(CallingForm, WinBufAlloc32, Handle&, BufferSize, _
         A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   End If
   BufAlloc32 = Handle&

End Function

Public Function BufAlloc64(CallingForm As Form, ByVal BufferSize As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      'AllocateMemory doubles BufferSize again, so * 8 is resulting size
      Handle& = AllocateMemory(BufferSize * 4)
   Else
      Handle& = cbWinBufAlloc64(BufferSize)
      x% = SaveFunc(CallingForm, WinBufAlloc64, Handle&, BufferSize, _
         A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   End If
   BufAlloc64 = Handle&

End Function

Public Function ScaledBufAlloc(CallingForm As Form, ByVal BufferSize As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      'AllocateMemory doubles BufferSize again, so * 8 is resulting size
      Handle& = AllocateMemory(BufferSize * 4)
   Else
      Handle& = ScaledWBufAlloc(BufferSize&)
      x% = SaveFunc(CallingForm, ScaledWinBufAlloc, Handle&, BufferSize, _
         A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   End If
   ScaledBufAlloc = Handle&

End Function

Function WSngArrayToBuf(ByVal CallingForm As Form, ByVal Handle As Long, _
sDatArray() As Single, ByVal Samples As Long, Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or 1 Or ForceWinAPI Then
      CopyMemory ByVal Handle, sDatArray(0, 0), Samples * 4
      ErrCode& = 0
   End If
   WSngArrayToBuf = ErrCode&

End Function

Function WArrayToBuf(ByVal CallingForm As Form, ByVal Handle As Long, nDatArray() As Integer, _
ByVal Samples As Long, Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory ByVal Handle, nDatArray(0, 0), Samples * 2
      ErrCode& = 0
   Else
      ULStat& = cbWinArrayToBuf(nDatArray(0, 0), Handle, 0, Samples)
      x% = SaveFunc(CallingForm, WinArrayToBuf, ULStat&, nDatArray(0, 0), _
      Handle, 0, Samples, A5, A6, A7, A8, A9, A10, A11, 0)
      ErrCode& = ULStat
   End If
   WArrayToBuf = ErrCode&

End Function

Function WArrayToBuf32(CallingForm As Form, MemHandle As Long, lDatArray() _
   As Long, Samples As Long, Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory ByVal MemHandle, lDatArray(0, 0), Samples * 4
      ErrCode& = 0
   Else
      If Not LibSupportsFunction(WinArrayToBuf32) Then Exit Function
      ULStat& = cbWinArrayToBuf32(lDatArray(0, 0), MemHandle, 0, Samples)
      x% = SaveFunc(CallingForm, WinArrayToBuf32, ULStat&, lDatArray(0, 0), _
      MemHandle, 0, Samples, A5, A6, A7, A8, A9, A10, A11, 0)
      ErrCode& = ULStat
   End If
   WArrayToBuf32 = ErrCode&
   
End Function


Function WDblArrayToBuf(CallingForm As Form, MemHandle As Long, _
dDatArray() As Double, ByVal Samples As Long, Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory ByVal MemHandle, dDatArray(0, 0), Samples * 8
      ErrCode& = 0
   Else
      ErrCode& = cbScaledWinArrayToBuf(dDatArray(0, 0), MemHandle, 0, Samples)
      x% = SaveFunc(CallingForm, ScaledWinArrayToBuf, ErrCode&, dDatArray(0, 0), _
      MemHandle, 0, Samples, A5, A6, A7, A8, A9, A10, A11, 0)
   End If
   WDblArrayToBuf = ErrCode&
   
End Function

Function LoadArrayFromWinBuf(CallingForm As Form, ByVal Handle As Long, _
nDatArray() As Integer, ByVal FirstPoint As Long, ByVal Samples As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory nDatArray(0, 0), ByVal Handle, Samples * 2
      ErrCode& = 0
   Else
      ULStat = cbWinBufToArray(Handle, nDatArray(0, 0), FirstPoint, Samples)
      x% = SaveFunc(CallingForm, WinBufToArray, ULStat, Handle, nDatArray(0, 0), _
         FirstPoint, Samples, A5, A6, A7, A8, A9, A10, A11, 0)
      ErrCode& = ULStat
   End If
   WBufToArray = ErrCode&

End Function

Function LoadArrayFromWinBuf32(CallingForm As Form, ByVal Handle As Long, _
lDatArray() As Long, ByVal FirstPoint As Long, ByVal Samples As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory lDatArray(0, 0), ByVal Handle, Samples * 4
      ErrCode& = 0
   Else
      ULStat = WBufToArray32(Handle, lDatArray(0, 0), FirstPoint, Samples)
      If SaveFunc(CallingForm, WinBufToArray32, ULStat, Handle, lDatArray(0, 0), _
         FirstPoint, Samples, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
      ErrCode& = ULStat
   End If
   LoadArrayFromWinBuf32 = ErrCode&
   
End Function

Function LoadArrayFromWinBuf64(CallingForm As Form, ByVal Handle As Long, _
lDatArray() As Currency, ByVal FirstPoint As Long, ByVal Samples As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory lDatArray(0, 0), ByVal Handle, Samples * 8
      ErrCode& = 0
   Else
      ULStat = WBufToArray64(Handle, lDatArray(0, 0), FirstPoint, Samples)
      If SaveFunc(CallingForm, WinBufToArray64, ULStat, Handle, lDatArray(0, 0), _
         FirstPoint, Samples, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   End If

End Function

Function LoadDblArrayFromWinBuf64(ByVal CallingForm As Form, ByVal Handle As Long, _
dDatArray() As Double, ByVal FirstPoint As Long, ByVal Samples As Long, _
Optional ForceWinAPI As Integer) As Long

   If (Not gbULLoaded) Or ForceWinAPI Then
      CopyMemory dDatArray(0, 0), ByVal Handle, Samples * 8
      ErrCode& = 0
   Else
      ULStat = ScaledWBufToArray(Handle, dDatArray(0, 0), FirstPoint, Samples)
      If SaveFunc(CallingForm, ScaledWinBufToArray, ULStat, Handle, dDatArray(0, 0), _
         FirstPoint, Samples, A5, A6, A7, A8, A9, A10, A11, 0) Then ULStat = -1
   End If
   LoadDblArrayFromWinBuf64 = ULStat

End Function

Sub ClearHandle()

   mlGenData = 0
   
End Sub
