Attribute VB_Name = "RevTrack"
' includes Universal Library functions up to
' and including

Const DISCOVERYREV! = 6.35
Const DIOARRAYREV! = 6.5
Const DALARMCLEARREV! = 6.51

Dim mabIgnoreRevFlags(2) As Boolean 'increase array size
                                 'as rev constants are added

Dim mnFireGo As Integer
Dim mfNoForm As Form
Dim mnIntCount As Integer
Dim mnEventNum As Long
Dim mlNumBoardsActive As Long
Dim ActiveBoards() As Long
Dim ActiveForms() As Form

Function BitConfig(ByVal BoardNum&, ByVal PortNum&, ByVal BitNum&, ByVal Direction&)

   #If MAKEREV >= 521 Then
      BitConfig = cbDConfigBit(ByVal BoardNum&, ByVal PortNum&, ByVal BitNum&, ByVal Direction&)
   #Else
      MsgBox "Function not available before Revision 5.21.", , "UL Update Required"
   #End If

End Function

Function LEDFlash(mnBoardNum) As Long

   #If MAKEREV >= 540 Then
      'LEDFlash = cbLocateUSBDevice(mnBoardNum)
      LEDFlash = cbFlashLED(mnBoardNum)
   #Else
      MsgBox "Function not available before Revision 5.40.", , "UL Update Required"
   #End If

End Function

Function GetSignal524(ByVal BoardNum&, ByVal Direction&, ByVal Signal&, ByVal Index&, ByRef Connection&, ByRef Polarity&) As Long

   #If MAKEREV >= 524 Then
      ULStat = cbGetSignal(BoardNum&, Direction&, Signal&, Index&, Connection&, Polarity&)
   #Else
      MsgBox "Function not available before Revision 5.24.", , "UL Update Required"
   #End If
   GetSignal524 = ULStat

End Function

Function SelectSignal524(ByVal BoardNum&, ByVal Direction&, ByVal Signal&, ByVal Connection&, ByVal Polarity&) As Long

   #If MAKEREV >= 524 Then
      ULStat = cbSelectSignal(BoardNum&, Direction&, Signal&, Connection&, Polarity&)
   #Else
      MsgBox "Function not available before Revision 5.30.", , "UL Update Required"
   #End If
   SelectSignal524 = ULStat

End Function


Function SetConfig520(ByVal InfoType%, ByVal BoardNum, ByVal DevNum%, ByVal ConfigItem%, ConfigVal&) As Long

   SetConfig520 = cbSetConfig(InfoType%, BoardNum, DevNum%, ConfigItem%, ConfigVal&)

End Function

Sub FireGo(BoardNum As Integer, FireVal As Integer)

   mnFireGo = FireVal

End Sub

Sub InstallBoards()
   
   MsgBox "The 'AddBoard' function is not available in this revision."

End Sub

Function EventEnable(ByVal FormInstance As Integer, ByVal BoardNum As Integer, EventType As Long, EventSize As Long, UserData As Form)

   Dim AddBoard As Boolean
   
   #If MAKEREV >= 520 Then
      ULStat = cbEnableEvent(BoardNum, EventType, EventSize, AddressOf MyCallbackFunc, UserData)
      If ULStat = 0 Then
         AddBoard = True
         If mlNumBoardsActive = 0 Then ReDim ActiveBoards(mlNumBoardsActive)
         ArraySize& = UBound(ActiveBoards)
         For i% = 0 To ArraySize& - 1
            If ActiveBoards(i%) = BoardNum Then
               AddBoard = False
               Exit For
            End If
         Next
         If AddBoard Then
            ReDim Preserve ActiveBoards(mlNumBoardsActive)
            ReDim Preserve ActiveForms(mlNumBoardsActive)
            ActiveBoards(mlNumBoardsActive) = BoardNum
            Set ActiveForms(mlNumBoardsActive) = UserData
            mlNumBoardsActive = mlNumBoardsActive + 1
         End If
         'InputEventBoard As Long, OutputEventBoard As Long
         'InputEventType As Long, OutputEventType As Long
         'InputEventForm As Form, OutputEventForm As Form
      End If
   #Else
      MsgBox "Function not available before Revision 5.20.", , "UL Update Required"
   #End If
   EventEnable = ULStat
   
End Function

Public Function CheckActiveEvents(ByVal BoardNum As Long, ByRef ActiveOnForm As Form) As Boolean

   Dim BoardIsActive As Boolean
   
   If mlNumBoardsActive = 0 Then
      BoardIsActive = False
   Else
   ArraySize& = UBound(ActiveBoards)
      For i% = 0 To ArraySize&
         If ActiveBoards(i%) = BoardNum Then
            BoardIsActive = True
            Set ActiveOnForm = ActiveForms(i%)
            Exit For
         End If
      Next
   End If
   CheckActiveEvents = BoardIsActive
   
End Function

Sub MyCallbackFunc(ByVal BoardNum As Integer, ByVal EventType As Long, ByVal EventData As Long, ByRef FormRef As Form)

   Static EventString$
   mnEventNum = mnEventNum + 1
   x% = SaveFunc(mfNoForm, CallbackFunc, 0, BoardNum, EventType, _
      EventData, FormRef.Name, A5, A6, A7, A8, A9, A10, A11, 0)
   FormRef.SetEvent EventType, EventData
   Select Case EventType
      Case ON_SCAN_ERROR
         FormRef.Controls("lblStatus").Caption = "Event OnScanError " & _
         Format$(EventData, "0") & " (" & GetErrorConst(EventData) & ")"
      Case ON_EXTERNAL_INTERRUPT
         'FormRef.Controls("lblStatus").Caption = "Number of Interrupts = " & Str(EventData)
         FormRef.Controls("lblStatus").Caption = "Event OnExternalInterrupt - " & _
         Format$(EventData, "0") & " interrupts received."
         If mnFireGo Then
            FormRef.Controls("cmdGo") = True
            FormRef.Controls("cmdGo").FontBold = Not FormRef.Controls("cmdGo").FontBold
         End If
      Case ON_PRETRIGGER
         FormRef.Controls("lblStatus").Caption = "Trigger occurred after " & _
         Format$(EventData, "0") & " samples.."
         FormRef.Controls("lblPTCount").Caption = Format$(EventData, "0")
      Case ON_DATA_AVAILABLE
         FormRef.eventPlot
         FormRef.Controls("lblStatus").Caption = " " & _
         Format$(EventData, "0") & " samples acquired."
      Case ON_END_OF_AI_SCAN
         If FormRef.Controls("mnuStopBG").Checked Then
            'FormRef.Controls("cmdStop") = True
            'ULStat = StopBackground520(BoardNum, AIFUNCTION)
            'If (Not gnScriptSave) Or (ULStat <> 0) Then
            '   If SaveFunc(FormRef, StopBackground, ULStat, BoardNum, _
            '      AIFUNCTION, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            'End If
            'If Not FormRef.Controls("tmrGoLoop").Enabled _
               Then FormRef.Controls("cmdStop").Visible = False
         End If
         FormRef.Controls("tmrCheckStatus").ENABLED = False
         FormRef.eventPlot
         FormRef.Controls("lblStatus").Caption = "Event EndOfInputScan at " & _
         Format$(EventData, "0") & " samples."
         DoEvents
         'FormRef.Controls("cmdPlot") = True
      Case ON_END_OF_AO_SCAN
         If FormRef.Controls("mnuStopBG").Checked Then
            'ULStat = cbStopBackground(BoardNum)
            ULStat = StopBackground520(BoardNum, AOFUNCTION)
            If (Not gnScriptSave) Or (ULStat <> 0) Then
               If SaveFunc(FormRef, StopBackground, ULStat, BoardNum, AOFUNCTION, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
            End If
            If Not FormRef.Controls("tmrGoLoop").ENABLED Then FormRef.Controls("cmdStop").Visible = False
         End If
         FormRef.Controls("tmrCheckStatus").ENABLED = False
         DoEvents
         FormRef.Controls("cmdPlot") = True
         FormRef.Controls("lblStatus").Caption = "Event EndOfOutputScan at " & _
         Format$(EventData, "0") & " samples."
      Case ON_CHANGE_DI
         ShowText True
         TypeOfPort& = (EventData And &HF0000) / 65536
         ValueOfPort& = EventData And &HFF
         If mnEventNum = 1 Then EventString$ = ""
         EventString$ = EventString$ & mnEventNum & ") Port Type = " & Str(TypeOfPort&) & _
         "  Port Value = " & Str(ValueOfPort&) & Chr$(13) & Chr$(10)
         'FormRef.Controls("lblStatus").Caption = EventString$
         TextList EventString$
         FormRef.Controls("lblStatus").Caption = "Event OnChangeDI at " & GetPortStringEx(TypeOfPort&) & _
         " to value of " & Format$(ValueOfPort&, "0") & "."
   End Select

End Sub

Sub ShowIntStatus(IntForm As Form, BoardNum As Integer)

   'Set faFormRefs(0) = IntForm

End Sub

Function UninstallEvent(ByVal BoardNum As Long, ByVal EventType As Long) As Integer
   
   Dim BoardRemoved As Boolean
   
   mnEventNum = 0
   If EventType = ON_EXTERNAL_INTERRUPT Then
      mnFireGo = False
      mnIntCount = 0
   End If
   #If MAKEREV >= 520 Then
      UninstallEvent = cbDisableEvent(BoardNum, EventType)
      If UninstallEvent = 0 Then
         BoardRemoved = False
         ArraySize& = UBound(ActiveBoards)
         For i% = 0 To ArraySize&
            If BoardRemoved Then
               If Not (i% > ArraySize&) Then
                  ActiveBoards(i% - 1) = ActiveBoards(i%)
                  Set ActiveForms(i% - 1) = ActiveForms(i%)
               End If
            Else
               If ActiveBoards(i%) = BoardNum Then
                  BoardRemoved = True
                  mlNumBoardsActive = mlNumBoardsActive - 1
               End If
            End If
         Next
         If mlNumBoardsActive > 0 Then
            mlNumBoardsActive = mlNumBoardsActive - 1
            ReDim Preserve ActiveBoards(mlNumBoardsActive)
            ReDim Preserve ActiveForms(mlNumBoardsActive)
         Else
            ActiveBoards(0) = -1
            Set ActiveForms(0) = Nothing
         End If
      End If
   #Else
      UpdateWarning "cbDisableEvent"
      UninstallEvent = NOTWINDOWSFUNC
   #End If

End Function

Function StopBackground520(ByVal board&, ByVal BGFunction&) As Long

   #If MAKEREV >= 520 Then
      StopBackground520 = cbStopBackground(board&, BGFunction&)
   #Else
      StopBackground520 = cbStopBackground(board&)
   #End If

End Function

Function GetStatus520(ByVal BoardNum&, Status%, CurCount&, CurIndex&, ByVal BGFunction&) As Long

   #If MAKEREV >= 520 Then
      GetStatus520 = cbGetStatus(BoardNum&, Status%, CurCount&, CurIndex&, BGFunction&)
   #Else
      GetStatus520 = cbGetStatus(BoardNum&, Status%, CurCount&, CurIndex&)
   #End If
    
End Function

Function GetConfig520(ByVal InfoType%, ByVal BoardNum, ByVal DevNum%, ByVal ConfigItem%, ConfigVal&)

    #If MAKEREV >= 520 Then
        GetConfig520 = cbGetConfig(InfoType%, BoardNum, DevNum%, ConfigItem%, ConfigVal&)
    #Else
        GetConfig520 = cbGetConfig(InfoType%, BoardNum, DevNum%, ConfigItem%, ValConfig%)
        ConfigVal& = ValConfig%
    #End If
    
End Function

Function CIn520(ByVal board%, ByVal Counter%, CBCount&) As Long

    #If MAKEREV >= 520 Then
        CIn520 = cbCIn(board%, Counter%, ValCount%)
        CBCount& = ValCount%
    #Else
        CIn520 = cbCIn(board%, Counter%, CBCount&)
    #End If

End Function

Function CLoad64Bit(ByVal BoardNum&, ByVal RegNum&, ByVal LoadValue As Currency) As Long

   #If MAKEREV >= 587 Then
       'TranslatedValue@ = LoadValue / 10000
       CLoad64Bit = cbCLoad64(BoardNum&, RegNum&, LoadValue)
   #Else
      MsgBox "Function not available before Revision 5.88.", , "UL Update Required"
      CLoad64Bit = NOTWINDOWSFUNC
   #End If

End Function

Sub UpdateWarning(FuncString As String)

   If geErrFlow = 0 Then Exit Sub
   RevString$ = Format$(CURRENTREVNUM, "0.00")
   MsgBox "Function " & FuncString & " not available in Revision " & RevString$ & ".", , "UL Update Required"
   
End Sub

Function InByte1632(board%, Register&)

   InByte1632 = cbInByte(board%, Register&)

End Function

Function OutWord1632(board%, RegToWrite&, ValToWrite%) As Long

   OutWord1623 = cbOutWord(board%, RegToWrite&, ValToWrite%)

End Function

Function OutByte1632(board%, RegToWrite&, ValToWrite%) As Long

   OutByte1623 = cbOutByte(board%, RegToWrite&, ValToWrite%)

End Function

Sub InitVB()

   'vbCrLf = Chr(13) & Chr(10)
   'used only in VB3 (integral to others)
   gn540 = False
#If MAKEREV >= 540 Then
   gn540 = True
#End If

End Sub

Function CtrInScan(ByVal BoardNum&, ByVal LowChan&, ByVal HighChan&, ByVal CBCount&, CBRate&, ByVal MemHandle&, ByVal Options&) As Long

#If MAKEREV > 569 Then
   CtrInScan = cbCInScan(BoardNum&, LowChan&, HighChan&, CBCount&, CBRate&, MemHandle&, Options&)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
#End If

End Function

Function CtrConfigScan(ByVal BoardNum&, ByVal Chan&, ByVal Mode&, ByVal DebounceTime&, ByVal DebounceTrigger&, ByVal EdgeDetection&, ByVal TickSize&, ByVal MapChannel&) As Long

#If MAKEREV > 569 Then
   CtrConfigScan = cbCConfigScan(BoardNum&, Chan&, Mode&, DebounceTime&, DebounceTrigger&, EdgeDetection&, TickSize&, MapChannel&)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
#End If

End Function

Function CtrClear(ByVal BoardNum&, ByVal CounterNum&) As Long

#If MAKEREV > 569 Then
   CtrClear = cbCClear(ByVal BoardNum&, ByVal CounterNum&)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
#End If

End Function

Function PlsOutStart(ByVal BoardNum&, ByVal TimerNum&, Frequency As Double, DutyCycle As Double, _
PulseCount As Long, InitialDelay As Double, IdleState As Long, Options As Long) As Long

#If MAKEREV > 587 Then
   'added Options at 588
   PlsOutStart = cbPulseOutStart(BoardNum&, TimerNum&, Frequency, DutyCycle, _
   PulseCount, InitialDelay, IdleState, Options)
#Else
   MsgBox "Function not available before Revision 5.88.", , "UL Update Required"
   PlsOutStart = NOTWINDOWSFUNC
#End If

End Function

Function PlsOutStop(ByVal BoardNum&, ByVal TimerNum&) As Long

#If MAKEREV > 587 Then
   PlsOutStop = cbPulseOutStop(BoardNum&, TimerNum&)
#Else
   MsgBox "Function not available before Revision 5.85.", , "UL Update Required"
   PlsOutStop = NOTWINDOWSFUNC
#End If

End Function

Function TmrOutStart(ByVal BoardNum&, ByVal TimerNum&, Frequency As Double) As Long

#If MAKEREV > 569 Then
   TmrOutStart = cbTimerOutStart(BoardNum&, TimerNum&, Frequency)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
   TmrOutStart = NOTWINDOWSFUNC
#End If

End Function

Function TmrOutStop(ByVal BoardNum&, ByVal TimerNum&) As Long

#If MAKEREV > 569 Then
   TmrOutStop = cbTimerOutStop(BoardNum&, TimerNum&)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
   TmrOutStop = NOTWINDOWSFUNC
#End If

End Function

Function WBufAlloc32(ByVal NumPoints&) As Long

#If MAKEREV > 569 Then
   WBufAlloc32 = cbWinBufAlloc32(ByVal NumPoints&)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
   WBufAlloc32 = NOTWINDOWSFUNC
#End If

End Function

Function WBufAlloc64(ByVal NumPoints&) As Long

#If MAKEREV > 588 Then
   WBufAlloc64 = cbWinBufAlloc64(ByVal NumPoints&)
#Else
   MsgBox "Function not available before Revision 5.89.", , "UL Update Required"
   WBufAlloc64 = NOTWINDOWSFUNC
#End If

End Function

Function WBufToArray32(ByVal MemHandle&, DataBuffer&, ByVal FirstPoint&, ByVal CBCount&) As Long

#If MAKEREV > 569 Then
   WBufToArray32 = cbWinBufToArray32(MemHandle&, DataBuffer&, FirstPoint&, CBCount&)
#Else
   MsgBox "Function not available before Revision 5.70.", , "UL Update Required"
   WBufToArray32 = NOTWINDOWSFUNC
#End If

End Function

Function WBufToArray64(ByVal MemHandle&, DataBuffer As Currency, ByVal FirstPoint&, ByVal CBCount&) As Long

#If MAKEREV > 588 Then
   WBufToArray64 = cbWinBufToArray64(MemHandle&, DataBuffer, FirstPoint&, CBCount&)
#Else
   MsgBox "Function not available before Revision 5.89.", , "UL Update Required"
   WBufToArray64 = NOTWINDOWSFUNC
#End If

End Function

Function ScaledWBufToArray(ByVal MemHandle&, DblArray#, ByVal FirstPoint&, ByVal CBCount&) As Long

#If MAKEREV > 586 Then
   ScaledWBufToArray = cbScaledWinBufToArray(MemHandle&, DblArray#, FirstPoint&, CBCount&)
#Else
   MsgBox "Function not available before Revision 5.87.", , "UL Update Required"
   ScaledWBufToArray = NOTWINDOWSFUNC
#End If

End Function

Function ScaledWBufAlloc(ByVal BufferSize As Long) As Long

#If MAKEREV > 587 Then
   ScaledWBufAlloc = cbScaledWinBufAlloc(BufferSize)
#Else
   MsgBox "Function not available before Revision 5.88.", , "UL Update Required"
   ScaledWBufAlloc = NOTWINDOWSFUNC
#End If

End Function

Function ScaledWArrayToBuf(DataArray#, ByVal MemHandle&, ByVal FirstPoint&, ByVal CBCount&) As Long

#If MAKEREV > 588 Then
   ScaledWArrayToBuf = cbScaledWinArrayToBuf(DataArray#, MemHandle&, FirstPoint&, CBCount&)
#Else
   MsgBox "Function not available before Revision 5.89.", , "UL Update Required"
   ScaledWArrayToBuf = NOTWINDOWSFUNC
#End If
   
End Function
Function IOTDaqInScan(ByVal BoardNum&, ChanArray%, ChanTypeArray%, GainArray%, ByVal ChanCount&, CBRate&, PretrigCount&, CBCount&, ByVal MemHandle&, ByVal Options&) As Long

#If MAKEREV > 570 Then
   IOTDaqInScan = cbDaqInScan(BoardNum&, ChanArray%, ChanTypeArray%, GainArray%, ChanCount&, CBRate&, PretrigCount&, CBCount&, MemHandle&, Options&)
#Else
   MsgBox "Function not available before Revision 5.71.", , "UL Update Required"
   IOTDaqInScan = NOTWINDOWSFUNC
#End If

End Function

Function IOTDaqSetTrigger(ByVal BoardNum&, ByVal TrigSource&, ByVal TrigSense&, ByVal TrigChan&, ByVal ChanType&, ByVal Gain&, ByVal Level!, ByVal Variance!, ByVal TrigEvent&) As Long

#If MAKEREV > 570 Then
   IOTDaqSetTrigger = cbDaqSetTrigger(BoardNum&, TrigSource&, TrigSense&, TrigChan&, ChanType&, Gain&, Level!, Variance!, TrigEvent&)
#Else
   MsgBox "Function not available before Revision 5.71.", , "UL Update Required"
   IOTDaqSetTrigger = NOTWINDOWSFUNC
#End If

End Function

Function IOTDaqOutScan(ByVal BoardNum&, ChanArray%, ChanTypeArray%, GainArray%, ByVal ChanCount&, CBRate&, CBCount&, ByVal MemHandle&, ByVal Options&) As Long

#If MAKEREV > 570 Then
   IOTDaqOutScan = cbDaqOutScan(BoardNum&, ChanArray%, ChanTypeArray%, GainArray%, ChanCount&, CBRate&, CBCount&, MemHandle&, Options&)
#Else
   MsgBox "Function not available before Revision 5.71.", , "UL Update Required"
   IOTDaqOutScan = NOTWINDOWSFUNC
#End If

End Function

Function IOTGetTCValues(ByVal BoardNum&, ChanArray%, ChanTypeArray%, ByVal ChanCount&, ByVal MemHandle&, ByVal FirstPoint&, ByVal Count&, ByVal CBScale&, TempValArray!) As Long

#If MAKEREV > 570 Then
   IOTGetTCValues = cbGetTCValues(BoardNum&, ChanArray%, ChanTypeArray%, ChanCount&, MemHandle&, FirstPoint&, Count&, CBScale&, TempValArray!)
#Else
   MsgBox "Function not available before Revision 5.71.", , "UL Update Required"
   IOTGetTCValues = NOTWINDOWSFUNC
#End If

End Function

Function SetConfigString573(ByVal InfoType&, ByVal BoardNum&, ByVal DevNum&, ByVal ConfigItem&, ByVal ConfigVal$, ByRef ConfigLen&) As Long

#If MAKEREV > 572 Then
   SetConfigString573 = cbSetConfigString(InfoType&, BoardNum&, DevNum&, ConfigItem&, ConfigVal$, ConfigLen&)
#Else
   MsgBox "Function not available before Revision 5.73.", , "UL Update Required"
   SetConfigString573 = NOTWINDOWSFUNC
#End If

End Function

Function GetConfigString573(ByVal InfoType&, ByVal BoardNum&, ByVal DevNum&, ByVal ConfigItem&, ByRef ConfigVal$, ByRef ConfigLen&) As Long

#If MAKEREV > 572 Then
   GetConfigString573 = cbGetConfigString(InfoType&, BoardNum&, DevNum&, ConfigItem&, ConfigVal$, ConfigLen&)
#Else
   MsgBox "Function not available before Revision 5.73.", , "UL Update Required"
   GetConfigString573 = NOTWINDOWSFUNC
#End If

End Function

Function ReadTEDS(BoardNum&, Chan&, DataBuffer As Byte, CBCount&, Options&) As Long

#If MAKEREV > 588 Then
   ReadTEDS = cbTEDSRead(mnBoardNum, Chan&, DataBuffer, CBCount&, Options&)
#Else
   MsgBox "Function not available before Revision 5.89.", , "UL Update Required"
   ReadTEDS = NOTWINDOWSFUNC
#End If

End Function

Public Function LibSupportsDiscovery() As Boolean

   Dim RevOK As Boolean
   
   RevOK = True
   If gfDLLRev < DISCOVERYREV! Then
      Dim response As VbMsgBoxResult
      If gfDLLRev = 0 Then
         response = MsgBox("Cannot determine UL version. " & _
            "Re-install Universal Test " & vbCrLf & "or move this " _
            & "executable to the installation directory" & _
            vbCrLf & "(Measurement Computing\Universal Test by default)." & _
            vbCrLf & "UL version must be " & Format(DISCOVERYREV!, "0.00") _
            & " or greater for this function." & vbCrLf & _
            "Call function anyway?", vbYesNo, "Cannot Verify UL Version")
      Else
         response = MsgBox("UL must be " & Format(DISCOVERYREV!, "0.00") & _
            " or later for these functions. " & vbCrLf & _
            "Application will fail otherwise. " & vbCrLf & "Current UL version " & _
            "appears to be " & Format(gfDLLRev, "0.00") & "." & vbCrLf _
            & "Call function anyway?", vbYesNo, "Verify UL Version")
      End If
      If response = vbNo Then RevOK = False
   End If
   LibSupportsDiscovery = RevOK

End Function

Public Function LibSupportsFunction(ByVal FunctionVal As Integer) As Boolean

   Dim RevOK As Boolean
   Dim FlagIndex As Integer
   
   Select Case FunctionVal
      Case IgnoreInstaCal To GetNetDeviceDescriptor
         RevImplemented! = DISCOVERYREV!
         FlagIndex = 0
      Case AInputMode To DOut32
         RevImplemented! = DIOARRAYREV!
         FlagIndex = 1
      Case DClearAlarm
         RevImplemented! = DALARMCLEARREV!
         FlagIndex = 2
   End Select
   
   RevOK = True
   If Not mabIgnoreRevFlags(FlagIndex) Then
      If gfDLLRev < RevImplemented! Then
         Dim response As VbMsgBoxResult
         FuncString$ = GetFunctionString(FunctionVal)
         NameEnd& = InStr(1, FuncString$, "(") - 1
         FuncName$ = Left(FuncString$, NameEnd&)
         If gfDLLRev <= 0 Then
            Prob$ = ""
            Loca$ = CurDir()
            If gfDLLRev = -2 Then Prob$ = " (missing 'VB5STKIT.DLL' at " & Loca$ & ")"
            response = MsgBox("Cannot determine UL version" & Prob$ & ". " & vbCrLf & _
               "Re-install Universal Test " & "or move " _
               & "this executable to the installation directory " & _
               "(Measurement Computing\Universal Test" & vbCrLf & "by default).  " & _
               FuncName$ & " is supported in UL version " & _
               Format(RevImplemented!, "0.00") & " or greater." & vbCrLf & vbCrLf & _
               "Call function anyway?", vbYesNo, "Cannot Verify UL Version ")
         Else
            response = MsgBox(FuncName$ & " is supported in UL version " & _
               Format(RevImplemented!, "0.00") & " or later. " & vbCrLf & _
               "Current UL version " & "appears to be " & Format(gfDLLRev, _
               "0.00") & "." & vbCrLf & "Application will fail if " & _
               FuncName$ & " is not implemented " & vbCrLf & "in the DLL in use." _
               & vbCrLf & vbCrLf & "Call function anyway?", vbYesNo, _
               "Verify UL Version for " & FuncName$ & " Support")
         End If
         If response = vbNo Then
            RevOK = False
         Else
            mabIgnoreRevFlags(FlagIndex) = True
         End If
      End If
   End If
   LibSupportsFunction = RevOK

End Function


