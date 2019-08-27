Attribute VB_Name = "ControlBoards"
Const A_In = 1
Const A_INSCAN = 2
Const A_Out = 3
Const A_OUTSCAN = 4
Const D_OUT = 5
Const D_OUTSCAN = 6
Const PULSE_OUT = 7
Const CLOCK_OUT = 8
Const TRIG_OUT = 9
Const ENC_POS = 10
Const ENC_NEG = 11
Const ENC_RATE = 12
Const ENC_INIT = 13
Const ENC_INDEX = 14
Const COMPL_OUT = 15
Const SWITCH_OUT = 16
Const EXTCTL_OUT = 17
Const LOOPBACK_OUT = 18
Const DIG_TRIG = 19
Const IDX_DELAY = 20
Const IDX_LENGTH = 21
Const IDX_EDGE = 22

Const BC2500AO = 1
Const BC2500AI = 2
Const BCALTAO = 3
Const BCADVPULSE = 8
Const BC9513 = 16
Const BCTRIGGER = 32
Const BCAUXSWITCH = 64
Const BCXPORTSWITCH = 128
Const BCAUXSELECT = 256
Const BCPORTSELECT = 512
Const BCAUXDIO = 1024
Const BCPORTLOOP = 2048
Const BCAUXTRIG = 4096
Const BCPORTDIO = 8192

'signal sources (mnSigSource)
Const PULSEGEN = 1
Const FUNCGEN = 2
Const TRIGGEN = 3
Const DIGGEN = 4
Const DIGINIT = 5

'digital output triggers
Const DIGCHGSTATE = 1
Const DIGTOGGLE = 2

Type MCCContrl
   GPIBName As String
   BoardName As String
   BoardNum As Integer
   BoardClass As Integer
   DevBase As Integer
   DevIndex As Integer
End Type

Dim mauMCtl() As MCCContrl
Dim mnNumCtlrs As Integer
Dim mnDataType As Integer
Dim mfFreq As Double, mfAmpl As Single, mfOffset As Single
Dim mlWidth As Long, mfRate As Single
Dim mfEncRate As Single, mfRate1 As Single, mfRate2 As Single
Dim mnOutCtl As Integer, mnEncOutCtl As Integer
Dim mnGenFunc As Integer
Dim mnCntSrc As Integer 'one of the divider sources (FREQ1,2, etc)
                        'for main source ctr (base + 4)
Dim mnAltSrc As Integer 'one of the divider sources (FREQ1,2, etc)
                        'for alt source ctr (base + 3)
Dim mnEncSrc As Integer 'divider source when ENCODER mode is used
Dim mnUseAlt As Integer 'semaphore to configure alt source ctr (base + 3)
Dim mlCtrSrc As Long    'the value of BICLOCK as set in Instacal
Dim mnUseTmr As Integer 'for MCC devices - use the timer output of 2500 series
                        'as a pacer for external clock scripts
Dim mnLoadReg As Integer, mlDuty As Long
Dim mlEncDuty As Long
Dim mlIdxBoard As Long, mnIdxDelay As Integer
Dim mnIdxLength As Integer, mnIdxEdge As Integer
Dim mnPulseFunc As Integer, mnSigSource As Integer
Dim mnSaveOutCtl As Integer, mnSaveFuncType As Integer
Dim mfOffSetTweak As Single
Dim mlTriggerDelay As Long
Dim mnNoNonStream As Integer, mvCustomRange As Variant
Dim mlPulseCount As Long, mlIdleState As Long
Dim mfTrigFreq As Double, mfDuty As Double, mfDelay As Double
Dim mfRateDivisor As Single
Dim mfrmNoForm As Form
Dim mlSwitchPort As Long, mlSwitchPortEnd As Long
Dim mlXCtlPort As Long, mlXCtlPortEnd As Long
Dim mlLoopPort As Long, mlLoopPortEnd As Long
Dim mlXPortDIOEnd As Long, mlXPortDIO As Long
Dim AiRelayMap() As Long, mnDefaultMap As Integer
Dim mlDTrigLine As Long, mlAuxTrigState As Long, mlAuxPrevState As Long
Dim mlDTrigDir As Long, mlAuxTrigType As Long

Sub SetInitialValues()
   
   mlIdxBoard = -1
   mnIdxLength = -1

End Sub

Function InitControlBoards(Device As Integer, BdDevName As String, _
Optional SignalRoutingOnly As Boolean) As Integer
   
   If SignalRoutingOnly Then
      Select Case BdDevName
         Case "SWITCH", "XSELECT", "LOOPBACK"
         Case Else
            InitControlBoards = 0
            Exit Function
      End Select
   End If
   'find if there's a control board associated with
   'the GPIB board name passed in BdDevName
   lpFileName$ = "Cfg488crh.ini"
   lpApplicationName$ = "ControlBoard"
   lpKeyName$ = BdDevName
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "-1"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   Device = Val(Left$(lpReturnedString$, x%))

   'if found, find out if there's a board installed at that BoardNum
   InfoType% = BOARDINFO
   DevNum% = 0
   ConfigItem% = BIBOARDTYPE
   If gbULLoaded Then
      ULStat& = GetConfig520(InfoType%, Device, DevNum%, ConfigItem%, ValConfig&)
      If Not (ValConfig& = 0) Then
         'there is a legit control board installed
         InitControlBoards = ValConfig&
         BoardName$ = GetNameOfBoard(Device)
         frmNewGPIBCtl(0).cmbDevice.AddItem BdDevName
         frmNewGPIBCtl(0).cmbDevice.ListIndex = mnNumCtlrs
         ReDim Preserve mauMCtl(mnNumCtlrs)
         mauMCtl(mnNumCtlrs).GPIBName = BdDevName
         mauMCtl(mnNumCtlrs).BoardName = BoardName$
         mauMCtl(mnNumCtlrs).BoardNum = Device
         mauMCtl(mnNumCtlrs).DevBase = 0
         'set the board class
         Select Case BoardName$
            Case "PCI-2517", "PCI-2515", "USB-2527", "USB-2537"
               mauMCtl(mnNumCtlrs).BoardClass = BC2500AO
            Case "USB-1616HS-4", "USB-1616HS-2"
               mauMCtl(mnNumCtlrs).BoardClass = BC2500AO
            Case "PCI-2511", "PCI-2513", "USB-2523", "USB-2533"
               mauMCtl(mnNumCtlrs).BoardClass = BC2500AI
            Case "USB-3101FS"
               mauMCtl(mnNumCtlrs).BoardClass = BCALTAO
            Case "USB-4301", "USB-4302", "USB-4303", "USB-4304"
               mauMCtl(mnNumCtlrs).BoardClass = BC9513
               mauMCtl(mnNumCtlrs).DevBase = 1
            Case "USB-1808", "USB-1808X"
               mauMCtl(mnNumCtlrs).BoardClass = BCADVPULSE
               mlIdxBoard = Device
            Case "PCI-CTR05", "PCI-CTR10", "PCI-CTR20HD"
               mauMCtl(mnNumCtlrs).BoardClass = BC9513
               mauMCtl(mnNumCtlrs).DevBase = 1
            Case "USB-1208HS", "USB-1208HS-2AO", "USB-1208HS-4AO"
               Select Case BdDevName
                  Case "AUXDIO"
                     mauMCtl(mnNumCtlrs).BoardClass = BCAUXDIO
                  Case "AUXTRIG0", "AUXTRIG1", "AUXTRIG2", "AUXTRIG3"
                     mauMCtl(mnNumCtlrs).BoardClass = BCAUXTRIG
                     'mlDTrigLine = Val(Right(BdDevName, 1))
                  Case Else
                     mauMCtl(mnNumCtlrs).BoardClass = BCTRIGGER
               End Select
               'If BdDevName = "AUXDIO" Then
               '   mauMCtl(mnNumCtlrs).BoardClass = BCAUXDIO
               'Else
               '   mauMCtl(mnNumCtlrs).BoardClass = BCTRIGGER
               'End If
            Case "USB-2553", "USB-2557", "USB-2559"
               mauMCtl(mnNumCtlrs).BoardClass = BCTRIGGER
            Case "USB-1608G", "USB-1608GX", "USB-1608GX-2AO"
               If BdDevName = "AUXDIO" Then
                  mauMCtl(mnNumCtlrs).BoardClass = BCAUXDIO
               Else
                  mauMCtl(mnNumCtlrs).BoardClass = BCTRIGGER
               End If
            Case "PCI-DIO48", "PCI-DIO48H", "PCI-DIO96", "PCI-DIO96H", _
            "USB-DIO24H/37", "miniLAB 1008", "miniLAB 1008-001", _
            "USB-1408FS", "USB-1208FS", "USB-7204", "USB-ERB24", "USB-ERB08"
               'programmable devices (currently ignores AUXPORT, if any)
               ULStat& = cbGetConfig(BOARDINFO, Device, 0, BIDINUMDEVS, NumDigDevs&)
               ULStat& = cbGetConfig(DIGITALINFO, Device, 0, DIDEVTYPE, PortType&)
               LastDev& = NumDigDevs& - 1
               If PortType& = AUXPORT Then LastDev& = NumDigDevs& - 2
               Select Case BdDevName
                  Case "PROGDIO"
                     mlXPortDIO = 10
                     If PortType& > 10 Then mlXPortDIO = PortType&
                     mlXPortDIOEnd = mlXPortDIO + LastDev&
                     mauMCtl(mnNumCtlrs).BoardClass = BCPORTDIO
                     FirstPort& = mlXPortDIO
                     LastPort& = mlXPortDIOEnd
                  Case "SWITCH"
                     mlSwitchPort = 10
                     If PortType& > 10 Then mlSwitchPort = PortType&
                     mlSwitchPortEnd = mlSwitchPort + LastDev&
                     mauMCtl(mnNumCtlrs).BoardClass = BCXPORTSWITCH
                     FirstPort& = mlSwitchPort
                     LastPort& = mlSwitchPortEnd
                  Case "XSELECT"
                     mlXCtlPort = 10
                     If PortType& > 10 Then mlXCtlPort = PortType&
                     mlXCtlPortEnd = mlXCtlPort + LastDev&
                     mauMCtl(mnNumCtlrs).BoardClass = BCPORTSELECT
                     FirstPort& = mlXCtlPort
                     LastPort& = mlXCtlPortEnd
                  Case "LOOPBACK"
                     mlLoopPort = 10
                     If PortType& > 10 Then mlLoopPort = PortType&
                     mlLoopPortEnd = mlLoopPort + LastDev&
                     mauMCtl(mnNumCtlrs).BoardClass = BCPORTLOOP
                     FirstPort& = mlLoopPort
                     LastPort& = mlLoopPortEnd
               End Select
               For CurPort& = FirstPort& To LastPort&
                  ULStat& = cbDConfigPort(Device, CurPort&, DIGITALOUT)
               Next
            Case "PCI-PDISO16", "USB-7112", "USB-7110", "PCI-PDISO8", "USB-PDISO8/40"
               'non-programmable devices (currently AUXPORT only)
               Select Case BdDevName
                  Case "SWITCH"
                     mlSwitchPort = 1
                     mlSwitchPortEnd = 1
                     mauMCtl(mnNumCtlrs).BoardClass = BCAUXSWITCH
                  Case "XSELECT"
                     mlXCtlPort = 1
                     mlXCtlPortEnd = 1
                     mauMCtl(mnNumCtlrs).BoardClass = BCAUXSELECT
                  Case "LOOPBACK"
                     mlLoopPort = 1
                     mlLoopPortEnd = 1
                     mauMCtl(mnNumCtlrs).BoardClass = BCPORTLOOP
               End Select
         End Select
         mauMCtl(mnNumCtlrs).DevIndex = 0
         Select Case BdDevName
            'if there are multiple outputs, index each
            Case "PULSEGEN1"
               mauMCtl(mnNumCtlrs).DevIndex = 1
            Case "PULSEGEN2"
               mauMCtl(mnNumCtlrs).DevIndex = 2
         End Select
         mnNumCtlrs = mnNumCtlrs + 1
      End If
   End If
   If Not mnDefaultMap Then SetAiMap "-1"

End Function

Sub WriteMCC(Device As Integer, Cmd As String)

   If Not Device < 0 Then BoardNum% = mauMCtl(Device).BoardNum
   FunctionType% = InterpretCmds(Device, Cmd)
   Select Case FunctionType%
      Case A_OUTSCAN
         AnalogOut BoardNum
      Case PULSE_OUT
         Select Case mnSigSource
            Case DIGGEN
               ToggleAuxTrig BoardNum, Device
            Case DIGINIT
               InitAuxTrig BoardNum, Device
            Case Else
               PulseOut BoardNum, Device
         End Select
      Case CLOCK_OUT
         Select Case mnSigSource
            Case PULSEGEN
               OrgRate! = mfRate
               ClockOut BoardNum, Device
               If mfRateDivisor = 0 Then mfRateDivisor = 1
               If Not (OrgRate! = mfRate / mfRateDivisor) _
               Then Cmd$ = mfRate / mfRateDivisor
            Case FUNCGEN
               TimerOut BoardNum
            Case TRIGGEN
               TriggerCtl BoardNum
         End Select
      Case TRIG_OUT
         TriggerOut BoardNum, Device
      Case ENC_POS
         SetEncDirection BoardNum, Device, 0
      Case ENC_NEG
         SetEncDirection BoardNum, Device, 1
      Case ENC_RATE
         SetEncRate BoardNum, Device
      Case ENC_INIT
         EncInit BoardNum, Device
      Case ENC_INDEX
         EncIndex BoardNum, Device
      Case IDX_DELAY, IDX_LENGTH, IDX_EDGE
         SetIdxParameters
      Case COMPL_OUT
         SetCompl BoardNum, Device
      Case SWITCH_OUT
         SetSwitch BoardNum, Cmd
      Case EXTCTL_OUT
         SetExtCtl BoardNum, Cmd
      Case LOOPBACK_OUT
         SetLoopBack BoardNum, Cmd
   End Select

End Sub

Sub AnalogOut(BoardNum As Integer)
         
   If (mfFreq = 0) And (mnDataType > 0) Then Exit Sub
   ULStat = StopBackground520(BoardNum, AOFUNCTION)
   SaveFlow% = geErrFlow
   geErrFlow = 0  'don't bother reporting or trapping error here
   DispSave% = gnLocalErrDisp
   gnLocalErrDisp = False
   ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
   DataType% = mnDataType
   Select Case DataType%
      Case 0
         CBCount& = 1
      Case 1
         'square wave
         CBCount& = 1000
         CBRate& = mfFreq * 2
      Case Else
         CBCount& = 10000
         CBCount& = 10000
         CBRate& = CBCount& * mfFreq
         Do While CBRate& > 1000000
            CBRate& = CBRate& / 10
            CBCount& = CBCount& / 10
         Loop
         Do While CBRate& < 1
            CBRate& = CBRate& * 10
            CBCount& = CBCount& * 10
         Loop
   End Select
   nRange% = UNI10VOLTS
   If DataType% = 0 Then nRange% = BIP10VOLTS
   If Not IsEmpty(mvCustomRange) Then
      Amplitude& = GetCounts(16, nRange%, mfAmpl, mvCustomRange)
   Else
      Amplitude& = GetCounts(16, nRange%, mfAmpl)
   End If
   AdjOffset& = GetCounts(16, nRange%, Abs(mfOffset))
   If DataType% > 0 Then
      AdjOffset& = AdjOffset& / 2
      If Sgn(mfOffset) = -1 Then
         Offset& = 32768 - AdjOffset&
      Else
         Offset& = 32768 + AdjOffset&
      End If
   Else
      Offset& = 32768 - AdjOffset&
      If Sgn(mfOffset) = -1 Then Offset& = Offset& * -1
   End If
   Chans% = 0
   If DataType% = 0 Then
      DataValue% = ULongValToInt(Amplitude + Offset&)
      ULStat = cbAOut(BoardNum, FirstChan%, nRange%, DataValue%)
   Else
      If CBRate& = 0 Then CBRate& = 1
      Options& = BACKGROUND + CONTINUOUS + NONSTREAMEDIO
      InitOutputBuffer Handle, DataType%, CBCount&, Chans%, Amplitude&, Offset&, True
      If Not mnNoNonStream Then ULStat = cbAOutScan(BoardNum, FirstChan%, _
      LastChan%, CBCount&, CBRate&, nRange%, Handle, Options&)
      If (ULStat = BADOPTION) Or mnNoNonStream Then
         mnNoNonStream = True
         If (ULStat = BADOPTION) Then MsgBox "Non-streaming output not supported. " & _
         "Signal generation may be limited.", vbInformation, "Limited Rate Control Device"
         Options& = BACKGROUND + CONTINUOUS
         ULStat = cbAOutScan(BoardNum, FirstChan%, LastChan%, CBCount&, CBRate&, nRange%, Handle, Options&)
      End If
   End If
   geErrFlow = SaveFlow%
   gnLocalErrDisp = DispSave%
   ULStat = cbErrHandling(gnLocalErrDisp, DONTSTOP)
   If Not ULStat = 0 Then
      If SaveFunc(mfrmNoForm, AOutScan, ULStat, BoardNum, FirstChan%, LastChan%, CBCount&, CBRate&, nRange%, Handle, Options&, A9, A10, A11, 0) Then Exit Sub
   End If
   
End Sub

Sub InitAuxTrig(BoardNum, Device)

   ULStat = cbDConfigBit(BoardNum, AUXPORT, mlDTrigLine, mlDTrigDir)
   ULStat = cbDBitOut(BoardNum, AUXPORT, mlDTrigLine, mlAuxTrigState)

End Sub

Sub ToggleAuxTrig(BoardNum As Integer, Device As Integer)

   If mlAuxTrigType = DIGCHGSTATE Then
      ActiveState& = 1
      If mlAuxPrevState = 1 Then ActiveState& = 0
      PassiveState& = ActiveState&
      mlAuxPrevState = ActiveState&
   Else
      PassiveState& = mlAuxTrigState
      ActiveState& = 1
      If PassiveState& = 1 Then ActiveState& = 0
   End If
   ULStat = cbDBitOut(BoardNum, AUXPORT, mlDTrigLine, ActiveState&)
   ULStat = cbDBitOut(BoardNum, AUXPORT, mlDTrigLine, PassiveState&)
   
End Sub

Sub PulseOut(BoardNum As Integer, Device As Integer)

   BoardClass = mauMCtl(Device).BoardClass
   If BoardClass = BCAUXTRIG Then
   Else
      'load the source counter to establish pulse width
      'default configuration is 10us
      'configre base + 4 unless using alternate source (for pulsegen 1 and 2)
      LoadCount& = mlWidth
      If Not mnUseAlt Then
         CounterNum% = mauMCtl(Device).DevBase + 4
         Source% = mnCntSrc
      Else
         CounterNum% = mauMCtl(Device).DevBase + 3
         Source% = mnAltSrc
      End If
      If ((mlCtrSrc / 2) * mlWidth) > 50000 Then
         AdjustSource% = True
         LoadCount& = (mlWidth / 1000) * 2
         CountNum% = mauMCtl(Device).DevBase + 3
         ULStat = cbC9513Config(BoardNum, CountNum%, NOGATE, POSITIVEEDGE, _
         FREQ2, CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
         LoadVal& = (mlCtrSrc * 10000) / 2
         ULStat = cbCLoad32(BoardNum, CountNum%, LoadVal&)
         Source% = CounterNum%
      Else
         CountNum% = mauMCtl(Device).DevBase + 3
         ULStat = cbC9513Config(BoardNum, CountNum%, NOGATE, POSITIVEEDGE, _
         FREQ1, CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
         LoadVal& = (mlCtrSrc * 10) / 2
         ULStat = cbCLoad32(BoardNum, CountNum%, LoadVal&)
      End If
      ULStat = cbC9513Config(BoardNum, CounterNum%, NOGATE, POSITIVEEDGE, Source%, _
      CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
      LoadValue& = ((mlCtrSrc) / 2) * LoadCount&
      ULStat = cbCLoad32(BoardNum, CounterNum%, LoadValue&)
      
      'configure the output counter for single pulse
      CounterNum% = mauMCtl(Device).DevIndex + mauMCtl(Device).DevBase
      GateControl% = NOGATE
      SpecialGate% = CBDISABLED
      RecycleMode% = ONETIME
      If Not mnUseAlt Then
         CountSource% = mauMCtl(Device).DevBase       'CTRINPUT1
      Else
         CountSource% = mauMCtl(Device).DevBase + 4   'CTRINPUT5
      End If
      If mlTriggerDelay > 0 Then
         GateControl% = AHEGATE
         SpecialGate% = True
         RecycleMode% = RECYCLE
         CountSource% = mnAltSrc + 1
         'mnAltSrc = FREQ2
      End If
      CounterEdge% = POSITIVEEDGE
      Reload% = LOADREG
      BCDMode% = CBDISABLED
      CountDirec% = COUNTDOWN
      OutputCtrl% = mnOutCtl  'HIGHPULSEONTC
      ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
      CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
      If mlTriggerDelay > 0 Then
         Delay& = mlTriggerDelay
         If Delay& < 2 Then Delay& = 2
         ULStat = cbCLoad(BoardNum, CounterNum%, Delay&)
      End If
      InitCounter Device, 0
      SetCompl BoardNum, Device
   End If
   
End Sub

Sub SetCompl(BoardNum As Integer, Device As Integer)

   'configure the output counter for single pulse
   CounterNum% = mauMCtl(Device).DevIndex + mauMCtl(Device).DevBase
   GateControl% = NOGATE
   CounterEdge% = POSITIVEEDGE
   If Not mnUseAlt Then
      CountSource% = mauMCtl(Device).DevBase       'CTRINPUT1
   Else
      CountSource% = mauMCtl(Device).DevBase + 4   'CTRINPUT5
   End If
   SpecialGate% = CBDISABLED
   Reload% = LOADREG
   RecycleMode% = ONETIME
   BCDMode% = CBDISABLED
   CountDirec% = COUNTDOWN
   OutputCtrl% = mnOutCtl  'HIGHPULSEONTC
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
   CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)

End Sub

Sub TriggerOut(BoardNum As Integer, Device As Integer)

   CounterNum% = mauMCtl(Device).DevIndex + mauMCtl(Device).DevBase
   ULStat = cbCLoad32(BoardNum, CounterNum%, mlTriggerDelay)
   
End Sub

Sub TimerOut(BoardNum As Integer)

   If mfFreq < 15 Then Exit Sub
   ULStat = cbTimerOutStart(BoardNum, 0, mfFreq)
   
End Sub

Sub TriggerCtl(BoardNum As Integer)

   If mnOutCtl = DISCONNECTED Then
      ULStat = cbPulseOutStop(BoardNum, 0)
      Exit Sub
   End If
   If mfTrigFreq = 0 Then Exit Sub
   DutyCycle# = mfDuty
   If mfDuty = 0 Then DutyCycle# = 0.5
   ULStat = cbPulseOutStart(BoardNum, 0, mfTrigFreq, DutyCycle#, _
   mlPulseCount, mfDelay, mlIdleState, Options&)
   
End Sub

Sub EncInit(BoardNum As Integer, Device As Integer)

   'set up source counters to establish source for encoder counters
   BaseCounter% = mauMCtl(Device).DevBase
   CounterNum% = BaseCounter% + 4
   ULStat = cbC9513Config(BoardNum, CounterNum%, NOGATE, POSITIVEEDGE, mnEncSrc, _
   CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
   
   'this counter will be used as source for 5 if rate is low
   'CounterNum% = BaseCounter% + 3
   'ULStat = cbC9513Config(BoardNum, CounterNum%, NOGATE, POSITIVEEDGE, FREQ1, _
   'CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, HIGHPULSEONTC)
   
   'configure the counters 1 and 2 to count on opposite edges of source counter
   Ctr0Dir = NEGATIVEEDGE
   Ctr1Dir = POSITIVEEDGE
   
   CounterNum% = BaseCounter%
   GateControl% = NOGATE
   CounterEdge% = Ctr0Dir
   CountSource% = CTRINPUT1
   SpecialGate% = CBDISABLED
   Reload% = LOADREG
   RecycleMode% = RECYCLE
   BCDMode% = CBDISABLED
   CountDirec% = COUNTDOWN
   OutputCtrl% = mnEncOutCtl
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
   CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
   
   CounterNum% = BaseCounter% + 1
   CounterEdge% = Ctr1Dir
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
   CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
   
   CounterNum% = BaseCounter% + 2   'index counter
   CounterEdge% = Ctr1Dir
   OutputCtrl% = ALWAYSLOW
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, NEGATIVEEDGE, _
   CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
   
   LoadValue& = 1
   ULStat = cbCLoad32(BoardNum, BaseCounter%, LoadValue&)
   ULStat = cbCLoad32(BoardNum, BaseCounter% + 1, LoadValue&)

End Sub

Sub SetEncRate(BoardNum As Integer, Device As Integer)
   
   BaseCounter% = mauMCtl(Device).DevBase
   If mfEncRate = 0 Then
      ULStat = cbC9513Config(BoardNum%, BaseCounter% + 4, NOGATE, POSITIVEEDGE, mnEncSrc, _
      CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, ALWAYSLOW)
   Else
      ULStat = cbC9513Config(BoardNum%, BaseCounter% + 4, NOGATE, POSITIVEEDGE, mnEncSrc, _
      CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
      LoadValue& = (10000 * mlCtrSrc / 4) / mfEncRate
      ULStat = cbCLoad32(BoardNum, BaseCounter% + 4, LoadValue&)
   End If
   SetIdxParameters

End Sub

Sub SetEncDirection(BoardNum As Integer, Device As Integer, Direction As Integer)
   
   BaseCounter% = mauMCtl(Device).DevBase
   ULStat = cbC9513Config(BoardNum%, BaseCounter% + 4, NOGATE, POSITIVEEDGE, mnEncSrc, _
   CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, ALWAYSLOW)
   
   'configure the counters 1 and 2 to count on opposite edges of source counter
   If Direction = 1 Then
      Ctr0Dir = POSITIVEEDGE
      Ctr1Dir = NEGATIVEEDGE
      LowCtr% = BaseCounter%
      HighCtr% = BaseCounter% + 1
   Else
      Ctr0Dir = NEGATIVEEDGE
      Ctr1Dir = POSITIVEEDGE
      LowCtr% = BaseCounter% + 1
      HighCtr% = BaseCounter%
   End If
   
   CounterNum% = BaseCounter%
   GateControl% = NOGATE
   CounterEdge% = Ctr0Dir
   CountSource% = CTRINPUT1
   SpecialGate% = CBDISABLED
   Reload% = LOADREG
   RecycleMode% = RECYCLE
   BCDMode% = CBDISABLED
   CountDirec% = COUNTDOWN
   OutputCtrl% = mnEncOutCtl
   
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
   CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
   
   CounterNum% = BaseCounter% + 1
   CounterEdge% = Ctr1Dir
   
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
   CountSource%, SpecialGate%, Reload%, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
   
   ULStat = cbCLoad(BoardNum%, LowCtr%, 1)
   PortVal% = &HE0 + LowCtr%
   PortNum% = 1
   ULStat = cbOutByte(BoardNum%, PortNum%, PortVal%)
   ULStat = cbCLoad(BoardNum%, HighCtr%, 1)
   PortVal% = &HE0 + HighCtr%
   PortNum% = 1
   ULStat = cbOutByte(BoardNum%, PortNum%, PortVal%)
   ULStat = cbCLoad(BoardNum%, HighCtr%, 1)
   DoEvents

   ULStat = cbC9513Config(BoardNum%, BaseCounter% + 4, NOGATE, POSITIVEEDGE, mnEncSrc, _
   CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)

End Sub

Sub EncIndex(BoardNum As Integer, Device As Integer)

   BaseCounter% = mauMCtl(Device).DevBase
   CounterNum% = BaseCounter% + 2   'index counter
   CounterEdge% = NEGATIVEEDGE
   If mlEncDuty = 0 Then
      OutputCtrl% = ALWAYSLOW
   Else
      OutputCtrl% = HIGHPULSEONTC
   End If
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
   CounterNum%, CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, OutputCtrl%)
   ULStat = cbCLoad(BoardNum%, CounterNum%, mlEncDuty)
   
End Sub

Sub SetIdxParameters()

   TimerNum& = 0
   PulseCount& = 1
   CycleDiv% = 4
   If mlIdxBoard = -1 Then Exit Sub
   If mnIdxLength = -1 Then
      ULStat = cbPulseOutStop(mlIdxBoard, 0)
      Exit Sub
   End If
   Frequency# = mfEncRate / 4
   InitialDelay# = (1 / mfEncRate) * (mnIdxDelay / 100)
   DutyCycle# = mnIdxLength * 0.0025
   ULStat = cbPulseOutStart(mlIdxBoard, TimerNum&, Frequency#, DutyCycle#, _
      PulseCount&, InitialDelay#, IdleState&, RETRIGMODE)
   
End Sub

Sub ClockOut(BoardNum As Integer, Device As Integer)

   'set configuration for continuous pulse from output counter
   'load the source counter to establish clock rate

   CounterNum% = mauMCtl(Device).DevBase + mauMCtl(Device).DevIndex
   GateControl% = NOGATE
   CounterEdge% = POSITIVEEDGE
   CountSource% = mnCntSrc
   SpecialGate% = CBDISABLED
   
   RecycleMode% = RECYCLE
   BCDMode% = CBDISABLED
   CountDirec% = COUNTDOWN
   If mnLoadReg = LOADREG Then
      OutputCtrl% = HIGHPULSEONTC
   Else
      OutputCtrl% = TOGGLEONTC
      RecycleMode% = ONETIME
      ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, _
      FREQ1, SpecialGate%, mnLoadReg, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
      ULStat = cbCLoad32(BoardNum, CounterNum%, 2)
      'wait for counter to hit terminal count
      t0! = Timer()
      Do
         tDiff! = Timer() - t0!
      Loop While tDiff! < 0.1
      DoEvents
      
      'initialize output to 0
      PortVal% = &HE0 + CounterNum%
      PortNum% = 1
      ULStat = cbOutByte(BoardNum%, PortNum%, PortVal%)
      RecycleMode% = RECYCLE
   End If
   If mnUseAlt Then
      CountSource% = mnAltSrc 'mauMCtl(Device).DevBase + 4
   Else
      CountSource% = mnCntSrc
   End If
   ULStat = cbC9513Config(BoardNum, CounterNum%, GateControl%, CounterEdge%, CountSource%, SpecialGate%, mnLoadReg, RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
   
   If mnLoadReg = LOADREG Then
      LoadValue& = mlCtrSrc * mfRate
      ULStat = cbCLoad32(BoardNum, CounterNum%, LoadValue&)
   Else
      LoadValue& = mlCtrSrc * mfRate
      HoldLoad& = LoadValue& * (mlDuty / 100)
      CtrLoad& = LoadValue& - HoldLoad&
      ULStat = cbCLoad32(BoardNum, CounterNum%, CtrLoad&)
      HoldRegister% = (HOLDREG1 + CounterNum%) - 1
      ULStat = cbCLoad32(BoardNum, HoldRegister%, HoldLoad&)
   End If
   If LoadValue& > 0 Then mfRate = 1# / (mlCtrSrc / LoadValue&)
   
End Sub

Function InterpretCmds(DevIndex As Integer, CommandStr As String) As Integer

   If Not DevIndex < 0 Then
      DevName = mauMCtl(DevIndex).GPIBName
      BoardNum = mauMCtl(DevIndex).BoardNum
   End If
   Select Case CommandStr
      Case "STOPBG"
         ConfigCtrlBoard DevIndex
         Exit Function
      Case "TRIG"
         Select Case DevName
            Case "AUXTRIG0", "AUXTRIG1", "AUXTRIG2", "AUXTRIG3"
               FuncType% = PULSE_OUT
               mnSigSource = DIGGEN
            Case "TRIGGER"
               FuncType% = CLOCK_OUT
               mnSigSource = TRIGGEN
            Case Else
               FuncType% = TRIG_OUT
         End Select
         mnSaveFuncType = FuncType%
         InterpretCmds = FuncType%
         Exit Function
   End Select

   Select Case DevName
      Case "GPIB0", "GPIB1"
      Case "8112A", "8112", "HP8112A", "HP8112", _
      "PULSEGEN0", "PULSEGEN1", "PULSEGEN2", "TRIGGER"
         If Len(CommandStr) > 5 Then
            Func$ = Left$(CommandStr, 3)
            Suffix$ = Trim(Right$(CommandStr, 2))
            value$ = Mid$(CommandStr, 4, Len(CommandStr) - 5)
            mnUseAlt = False
            If (DevName = "PULSEGEN1") Or (DevName = "PULSEGEN2") Then mnUseAlt = True
            Select Case Func$
               Case "WID"
                  If DevName = "TRIGGER" Then
                     Select Case Suffix$
                        Case "S"
                           DivVal# = 1
                        Case "MS"
                           DivVal# = 1000
                        Case "US"
                           DivVal# = 1000000
                        Case "NS"
                           DivVal# = 1000000000
                     End Select
                     WidVal# = Val(value$) / DivVal#
                     mfTrigFreq = 1 / (WidVal# * 3)
                     mfDuty = 0.333
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                  Else
                     'following based on the default 10us rate set up in ConfigCtrlBoard
                     Mult& = 1
                     If mnUseAlt Then
                        mnAltSrc = FREQ1
                     Else
                        mnCntSrc = FREQ1
                     End If
                     If Not mnOutCtl = LOWPULSEONTC Then mnOutCtl = HIGHPULSEONTC
                     If Suffix$ = "MS" Then
                        If mnUseAlt Then
                           mnAltSrc = FREQ3
                        Else
                           mnCntSrc = FREQ3
                        End If
                        Mult& = 10
                     End If
                     If Suffix$ = "S" Then
                        If mnUseAlt Then
                           mnAltSrc = FREQ4
                        Else
                           mnCntSrc = FREQ4
                        End If
                        Mult& = 1000
                     End If
                     mlWidth = Val(value$) * Mult&
                     FuncType% = PULSE_OUT
                     mnLoadReg = LOADREG
                  End If
               Case "PER"
                  If DevName = "TRIGGER" Then
                     Select Case Suffix$
                        Case "S"
                           DivVal# = 1
                        Case "MS"
                           DivVal# = 1000
                        Case "US"
                           DivVal# = 1000000
                        Case "NS"
                           DivVal# = 1000000000
                     End Select
                     PerVal# = Val(value$)
                     mfTrigFreq = 1 / (PerVal# / DivVal#)
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                  Else
                     'assume microseconds
                     mfRateDivisor = 1
                     If mnUseAlt Then
                        mnAltSrc = FREQ1
                     Else
                        mnCntSrc = FREQ1
                     End If
                     If Suffix$ = "MS" Then
                        mfRateDivisor = 10
                        If mnUseAlt Then
                           mnAltSrc = FREQ3
                        Else
                           mnCntSrc = FREQ3
                        End If
                     End If
                     If Suffix$ = "S" Then
                        mfRateDivisor = 1000
                        If mnUseAlt Then
                           mnAltSrc = FREQ4
                        Else
                           mnCntSrc = FREQ4
                        End If
                     End If
                     mfRate = Val(value$) * mfRateDivisor
                     FuncType% = CLOCK_OUT
                     mnSigSource = PULSEGEN
                  End If
               Case "DTY"
                  Suffix$ = Mid$(CommandStr, 5)
                  PerSign& = InStr(1, Suffix$, "%")
                  NumVal$ = Left(Suffix$, PerSign& - 1)
                  If DevName = "TRIGGER" Then
                     mfDuty = Val(NumVal$) / 100
                     mnSigSource = TRIGGEN
                  Else
                     mlDuty = Val(Suffix$)
                     mnLoadReg = LOADANDHOLDREG
                     If mlDuty < 1 Then mnLoadReg = LOADREG
                     mnSigSource = PULSEGEN
                  End If
                  FuncType% = CLOCK_OUT
               Case "DEL"
                  If DevName = "TRIGGER" Then
                     Select Case Suffix$
                        Case "S"
                           DivVal# = 1
                        Case "MS"
                           DivVal# = 1000
                        Case "US"
                           DivVal# = 1000000
                        Case "NS"
                           DivVal# = 1000000000
                     End Select
                     mfDelay = Val(value$) / DivVal#
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                  Else
                     DelayVal& = Val(value$)
                     Factor& = Fix(Log(DelayVal&) / Log(10))
                     Mult& = 10 ^ (Factor& - 1)
                     If Mult& < 1 Then Mult& = 1
                     If mnUseAlt Then
                        mnAltSrc = FREQ3
                     Else
                        mnCntSrc = FREQ3
                     End If
                     Divider& = 10
                     If Suffix$ = "US" Then
                        If mnUseAlt Then
                           mnAltSrc = FREQ2
                        Else
                           mnCntSrc = FREQ2
                        End If
                        Mult& = Mult& / 10
                        If Mult& < 1 Then Mult& = 1
                        Divider& = 1
                     End If
                     If Suffix$ = "S" Then
                        Mult& = 10
                        If mnUseAlt Then
                           mnAltSrc = FREQ4
                        Else
                           mnCntSrc = FREQ4
                        End If
                        Divider& = 1000
                     End If
                     mlWidth = 5 * Mult&
                     FuncType% = PULSE_OUT
                     mnOutCtl = HIGHPULSEONTC
                     mnLoadReg = LOADREG
                     mlTriggerDelay = DelayVal& '/ (mlWidth / Divider&)
                     'If mlTriggerDelay = 0 Then mlTriggerDelay = DelayVal&
                  End If
            End Select
         Else
            Select Case CommandStr
               Case "C0"
                  If DevName = "TRIGGER" Then
                     mnOutCtl = TOGGLEONTC
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                     If (mlPulseCount = 0) And (mlIdleState = 1) Then
                        'if continuous pulse, invert the duty cycle
                        'mfDuty = 1 - mfDuty
                        ULStat = cbPulseOutStop(BoardNum, 0)
                     End If
                     mlIdleState = 0
                  Else
                     InitCounter DevIndex, 0
                     mnOutCtl = HIGHPULSEONTC
                     FuncType% = COMPL_OUT
                  End If
               Case "C1"
                  If DevName = "TRIGGER" Then
                     mnOutCtl = TOGGLEONTC
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                     If (mlPulseCount = 0) And (mlIdleState = 0) Then
                        'if continuous pulse, invert the duty cycle
                        'mfDuty = 1 - mfDuty
                        ULStat = cbPulseOutStop(BoardNum, 0)
                     End If
                     mlIdleState = 1
                  Else
                     InitCounter DevIndex, 1
                     mnOutCtl = LOWPULSEONTC
                     FuncType% = COMPL_OUT
                  End If
               Case "D0"
                  If DevName = "TRIGGER" Then
                     mnOutCtl = TOGGLEONTC
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                  Else
                     mnOutCtl = mnSaveOutCtl
                     If Not mnSaveFuncType = 0 Then
                        FuncType% = mnSaveFuncType
                     Else
                        FuncType% = PULSE_OUT
                     End If
                  End If
               Case "D1"
                  OutputOff% = True
                  mnOutCtl = DISCONNECTED
                  FuncType% = PULSE_OUT
                  If DevName = "TRIGGER" Then
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                  End If
                  InterpretCmds = FuncType%
                  Exit Function
               Case "M1"
                  'continuous output
                  mnPulseFunc = CLOCK_OUT
                  mnSigSource = FUNCGEN
                  If DevName = "TRIGGER" Then
                     mnSigSource = TRIGGEN
                     mlPulseCount = 0
                     mlIdleState = 0
                  End If
               Case "M2"
                  'triggered output
                  If DevName = "TRIGGER" Then
                     FuncType% = CLOCK_OUT
                     mnSigSource = TRIGGEN
                  Else
                     mnPulseFunc = TRIG_OUT
                  End If
                  mlPulseCount = 1
            End Select
            If Not OutputOff% Then mnSaveOutCtl = mnOutCtl
         End If
      Case "FLUKE45", "F45"
         'GetCommands = GetFluke45Cmds(Instance%, CommandNum)
      Case "F8840", "8840", "FLUKE8840", "F8840A", "8840A", "FLUKE8840A"
         'GetCommands = GetF8840Cmds(Instance%, CommandNum)
      Case "HP34401", "34401", "HP34401A", "34401A"
         'GetCommands = GetHP34401Cmds(Instance%, CommandNum)
      Case "DP8200", "8200", "DP8200N", "8200N"
         'GetCommands = GetDP8200Cmds(Instance%, CommandNum)
      Case "HP3325", "3325", "HP3325A", "3325A"
         If Len(CommandStr) > 4 Then
            Func$ = Left$(CommandStr, 2)
            Suffix$ = Right$(CommandStr, 2)
            value$ = Mid$(CommandStr, 3, Len(CommandStr) - 4)
            Select Case Func$
               Case "FR"
                  Mult& = 1
                  If Suffix$ = "KH" Then Mult& = 1000
                  If Suffix$ = "MH" Then Mult& = 1000000
                  mfFreq = Val(value$) * Mult&
                  FuncType% = mnGenFunc  'A_OUTSCAN (could apply to TimerOut or AOut)
               Case "AM"
                  Div! = 1
                  If Suffix$ = "MV" Then Div! = 1000
                  If Suffix$ = "VR" Then Div! = 1.414
                  If Suffix$ = "MR" Then Div! = 1414
                  mfAmpl = Val(value$) / Div!
                  FuncType% = A_OUTSCAN
               Case "OF"
                  Div! = 1
                  If Suffix$ = "MV" Then Div! = 1000
                  If Suffix$ = "VR" Then Div! = 1.414
                  If Suffix$ = "MR" Then Div! = 1414
                  mfOffset = (Val(value$) / Div!) + mfOffSetTweak
                  FuncType% = A_OUTSCAN
            End Select
         Else
            Func$ = Left$(CommandStr, 2)
            value$ = Right$(CommandStr, 1)
            Select Case Func$
               Case "FU"
                  mnSigSource = FUNCGEN
                  mnGenFunc = A_OUTSCAN
                  Select Case value$
                     Case "0" 'DC level
                        mnDataType = 0
                        FuncType% = A_OUTSCAN
                     Case "1" 'sine
                        mnDataType = 2
                        FuncType% = A_OUTSCAN
                     Case "2" 'square
                        mnDataType = 1
                        FuncType% = A_OUTSCAN
                     Case "3" 'triangle
                        mnDataType = 4
                        FuncType% = A_OUTSCAN
                     Case "4" 'pos ramp
                        mnDataType = 3
                        FuncType% = A_OUTSCAN
                     Case "5" 'neg ramp
                        mnDataType = 3 'not yet conf for neg
                        FuncType% = A_OUTSCAN
                     Case "6" 'not valid for 3325 - used for MCC Ctl
                        FuncType% = CLOCK_OUT
                        mnGenFunc = CLOCK_OUT '(applies to TimerOut)
                  End Select
            End Select
         End If
      Case "ENCODER"
         mlWidth = 1
         SplitCmd = Split(CommandStr, " ")
         EncString$ = SplitCmd(0)
         Select Case EncString$
            Case "Dir"
               If SplitCmd(1) = "+" Then
                  FuncType% = ENC_POS
               Else
                  FuncType% = ENC_NEG
               End If
            Case "Index"
               IndexString$ = SplitCmd(1)
               mlEncDuty = Val(IndexString$)
               FuncType% = ENC_INDEX
            Case "Init"
               BaseCounter% = mauMCtl(DevIndex).DevBase + 4
               InitCounter DevIndex, 2, BaseCounter%
               mnEncOutCtl = TOGGLEONTC
               FuncType% = ENC_INIT
            Case "Rate"
               SpeedString$ = SplitCmd(1)
               Suffix$ = SplitCmd(2)
               Mult& = 1
               mnEncSrc = FREQ3
               If Suffix$ = "kHz" Then
                  Mult& = 10
                  mnEncSrc = FREQ1
               End If
               mfEncRate = Val(SpeedString$) * Mult&
               FuncType% = ENC_RATE
         End Select
      Case "INDEXER"
         mlWidth = 1
         SplitCmd = Split(CommandStr, " ")
         IdxString$ = SplitCmd(0)
         Select Case IdxString$
            Case "Delay"
               IdxDlyString$ = SplitCmd(1)
               mnIdxDelay = Val(IdxDlyString$)
               FuncType% = IDX_DELAY
            Case "Length"
               IdxLenString$ = SplitCmd(1)
               mnIdxLength = Val(IdxLenString$)
               FuncType% = IDX_LENGTH
            Case "Edge"
               IdxEdgeString$ = SplitCmd(1)
               mnIdxEdge = Val(IdxEdgeString$)
               FuncType% = IDX_EDGE
         End Select
      Case "AUXTRIG0", "AUXTRIG1", "AUXTRIG2", "AUXTRIG3"
         CmdArray = Split(CommandStr, " ")
         BaseCmd$ = CmdArray(0)
         Select Case BaseCmd$
            Case "C0", "C1"
               FuncType% = PULSE_OUT
               mnSigSource = DIGINIT
               mlDTrigDir = DIGITALOUT
               mlAuxTrigState = 0
               If BaseCmd$ = "C1" Then mlAuxTrigState = 1
               If UBound(CmdArray) > 0 Then mlDTrigLine = Val(CmdArray(1))
            Case "D0", "D1"
               FuncType% = PULSE_OUT
               mnSigSource = DIGINIT
               mlDTrigDir = DIGITALOUT
               If BaseCmd$ = "D1" Then mlDTrigDir = DIGITALIN
               If UBound(CmdArray) > 0 Then mlDTrigLine = Val(CmdArray(1))
            Case "M1", "M2"
               FuncType% = PULSE_OUT
               mnSigSource = DIGINIT
               mlDTrigDir = DIGITALOUT
               mlAuxTrigType = DIGTOGGLE
               If BaseCmd$ = "M1" Then mlAuxTrigType = DIGCHGSTATE
               If UBound(CmdArray) > 0 Then mlDTrigLine = Val(CmdArray(1))
            'Case "ToggleState"
               'FuncType% = PULSE_OUT
               'mnSigSource = DIGGEN
         End Select
      Case "SWITCH"
         FuncType% = SWITCH_OUT
      Case "XSELECT"
         FuncType% = EXTCTL_OUT
      Case "LOOPBACK"
         FuncType% = LOOPBACK_OUT
   End Select
   mnSaveFuncType = FuncType%
   InterpretCmds = FuncType%

End Function

Sub InitCounter(Device As Integer, CounterState As Integer, Optional SpecifiedCounter As Variant)

   If IsMissing(SpecifiedCounter) Then
      CounterNum% = mauMCtl(Device).DevIndex + mauMCtl(Device).DevBase
   Else
      CounterNum% = SpecifiedCounter
   End If
   
   BoardNum% = mauMCtl(Device).BoardNum
   Select Case CounterState
      Case 0
         'initialize output to 0
         PortVal% = &HE0 + CounterNum%
         PortNum% = 1
         ULStat = cbOutByte(BoardNum%, PortNum%, PortVal%)
      Case 1
         'initialize to 0 then toggle high
         PortVal% = &HE0 + CounterNum%
         PortNum% = 1
         ULStat = cbOutByte(BoardNum%, PortNum%, PortVal%)
         GateControl% = NOGATE
         CounterEdge% = POSITIVEEDGE
         CountSource% = FREQ1
         SpecialGate% = CBDISABLED
         Reload% = LOADREG
         RecycleMode% = ONETIME
         BCDMode% = CBDISABLED
         CountDirec% = COUNTDOWN
         OutputCtrl% = TOGGLEONTC
         ULStat = cbC9513Config(BoardNum%, CounterNum%, GateControl%, _
         CounterEdge%, CountSource%, SpecialGate%, Reload%, _
         RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
         ULStat = cbCLoad(BoardNum%, CounterNum%, 4)
      Case 2
         'initialize to 0 then set low
         PortVal% = &HE0 + CounterNum%
         PortNum% = 1
         ULStat = cbOutByte(BoardNum%, PortNum%, PortVal%)
         GateControl% = NOGATE
         CounterEdge% = POSITIVEEDGE
         CountSource% = FREQ1
         SpecialGate% = CBDISABLED
         Reload% = LOADREG
         RecycleMode% = ONETIME
         BCDMode% = CBDISABLED
         CountDirec% = COUNTDOWN
         OutputCtrl% = HIGHPULSEONTC
         ULStat = cbC9513Config(BoardNum%, CounterNum%, GateControl%, _
         CounterEdge%, CountSource%, SpecialGate%, Reload%, _
         RecycleMode%, BCDMode%, CountDirec%, OutputCtrl%)
         ULStat = cbCLoad(BoardNum%, CounterNum%, 4)
   End Select
   
End Sub

Sub SetSwitch(BoardNum As Integer, Cmd As String)

   CmdArray = Split(Cmd, " ")
   SwitchCmd$ = CmdArray(0)
   NumRelaysConfigured& = UBound(AiRelayMap)
   Select Case SwitchCmd$
      Case "CH"
         BitNum& = Val(CmdArray(1))
         If BitNum& > NumRelaysConfigured& Then
            MsgBox "Relays are only configured for " & _
               NumRelaysConfigured& + 1 & " bits.", _
               vbInformation, "Use Direct Control"
            Exit Sub
         End If
         RelayNum& = AiRelayMap(BitNum&)
         ULStat& = cbDBitOut(BoardNum, mlSwitchPort, RelayNum&, 1)
      Case "CHS"
         For SwitchPort& = mlSwitchPort To mlSwitchPortEnd
            ULStat& = cbDOut(BoardNum, SwitchPort&, 0)
         Next
         BitRange$ = CmdArray(1)
         RangeArray = Split(BitRange$, "-")
         LowBit& = RangeArray(0)
         HighBit& = RangeArray(1)
         If HighBit& > NumRelaysConfigured& Then
            MsgBox "Relays are only configured for " & _
               NumRelaysConfigured& + 1 & " bits.", _
               vbInformation, "Use Direct Control"
            Exit Sub
         End If
         For BitNum& = LowBit& To HighBit&
            RelayNum& = AiRelayMap(BitNum&)
            ULStat& = cbDBitOut(BoardNum, mlSwitchPort, RelayNum&, 1)
         Next
      Case "CLEAR"
         For SwitchPort& = mlSwitchPort To mlSwitchPortEnd
            ULStat& = cbDOut(BoardNum, SwitchPort&, 0)
         Next
      Case "REMAP"
         MapString$ = CmdArray(1)
         SetAiMap MapString$
   End Select

End Sub

Sub SetExtCtl(BoardNum As Integer, Cmd As String)

   'to do - handle FIRSTPORTA types
   BitPort& = mlXCtlPort
   BitOffset& = 0
   If mlXCtlPort > 10 Then
      BitPort& = 10
      BitOffset& = 8 * (mlXCtlPort - 10)
   End If
   CmdArray = Split(Cmd, " ")
   ExtCtlCmd$ = CmdArray(0)
   Select Case ExtCtlCmd$
      Case "XAIALTSIGS"
         For CtlPort& = mlXCtlPort To mlXCtlPortEnd
            ULStat& = cbDOut(BoardNum, CtlPort&, 0)
         Next
         LoopChans = Split(CmdArray(1), "-")
         LowAoLoop& = Val(LoopChans(0))
         HighAoLoop& = Val(LoopChans(1))
         For AOCh& = LowAoLoop& To HighAoLoop&
            ULStat& = cbDBitOut(BoardNum, BitPort&, AOCh& + BitOffset&, 1)
            'ULStat& = cbDBitOut(BoardNum, mlLoopPort, AOCh&, 1)
         Next
      Case "XAIALTSIG"
         AoLoop& = Val(CmdArray(1))
         ULStat& = cbDBitOut(BoardNum, BitPort&, AoLoop& + BitOffset&, 1)
      Case "XAICLOCK"
         ULStat& = cbDBitOut(BoardNum, BitPort&, 5 + BitOffset&, 1)
      Case "XAOCLOCK"
         
         ULStat& = cbDBitOut(BoardNum, BitPort&, 6 + BitOffset&, 1)
      Case "XCTLOFF"
         For XCtlPort& = mlXCtlPort To mlXCtlPortEnd
            ULStat& = cbDOut(BoardNum, XCtlPort&, 0)
         Next
      Case "XTRIG"
         ULStat& = cbDBitOut(BoardNum, BitPort&, 7 + BitOffset&, 1)
      Case "TRG8112SRC"
         ULStat& = cbDBitOut(BoardNum, BitPort&, 4 + BitOffset&, 1)
      Case "TRIGLOOP", "AUXTRIGLOOP"
         ULStat& = cbDBitOut(BoardNum, BitPort&, 0 + BitOffset&, 1)
         'ULStat& = cbDBitOut(BoardNum, mlLoopPort, 4, 1)
   End Select

End Sub

Sub SetLoopBack(BoardNum As Integer, Cmd As String)

   'to do - handle FIRSTPORTA types
   BitPort& = mlLoopPort
   If mlLoopPort > 10 Then
      BitPort& = 10
      BitOffset& = 8 * (mlLoopPort - 10)
   End If
   CmdArray = Split(Cmd, " ")
   ExtCtlCmd$ = CmdArray(0)
   Select Case ExtCtlCmd$
      Case "AOLOOP"
         LoopChans = Split(CmdArray(1), "-")
         LowAoLoop& = Val(LoopChans(0))
         HighAoLoop& = Val(LoopChans(1))
         For AOCh& = LowAoLoop& To HighAoLoop&
            ULStat& = cbDBitOut(BoardNum, BitPort&, AOCh& + BitOffset&, 1)
         Next
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find XSELECT device and switch to alternate source
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCPORTSELECT) Or (BClass% = BCAUXSELECT) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "XAIALTSIGS " & CmdArray(1)
               SetExtCtl SelBoard%, XCmd$
               Exit For
            End If
         Next
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find SWITCH device and enable relevant channels
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCXPORTSWITCH) Or (BClass% = BCAUXSWITCH) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "CHS " & CmdArray(1)
               SetSwitch SelBoard%, XCmd$
               Exit For
            End If
         Next
      Case "LOOPOFF"
         For LoopPort& = mlLoopPort To mlLoopPortEnd
            ULStat& = cbDOut(BoardNum, LoopPort&, 0)
         Next
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find XSELECT device and switch to alternate source
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCPORTSELECT) Or (BClass% = BCAUXSELECT) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "XCTLOFF"
               SetExtCtl SelBoard%, XCmd$
               Exit For
            End If
         Next
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find SWITCH device and enable relevant channels
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCXPORTSWITCH) Or (BClass% = BCAUXSWITCH) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "CLEAR"
               SetSwitch SelBoard%, XCmd$
               Exit For
            End If
         Next
      Case "TRIGLOOP"
         ULStat& = cbDBitOut(BoardNum, BitPort&, 4 + BitOffset&, 1)
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find XSELECT device and switch CH0 to alternate source
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCPORTSELECT) Or (BClass% = BCAUXSELECT) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "XAIALTSIG 0"
               SetExtCtl SelBoard%, XCmd$
               Exit For
            End If
         Next
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find SWITCH device and enable channel 0
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCXPORTSWITCH) Or (BClass% = BCAUXSWITCH) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "CH 0"
               SetSwitch SelBoard%, XCmd$
               Exit For
            End If
         Next
      Case "AUXTRIGLOOP"
         ULStat& = cbDBitOut(BoardNum, BitPort&, 6 + BitOffset&, 1)
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find XSELECT device and switch CH0 to alternate source
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCPORTSELECT) Or (BClass% = BCAUXSELECT) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "XAIALTSIG 0"
               SetExtCtl SelBoard%, XCmd$
               Exit For
            End If
         Next
         For Ctlr% = 0 To mnNumCtlrs - 1
            'find SWITCH device and enable channel 0
            BClass% = mauMCtl(Ctlr%).BoardClass
            If (BClass% = BCXPORTSWITCH) Or (BClass% = BCAUXSWITCH) Then
               SelBoard% = mauMCtl(Ctlr%).BoardNum
               XCmd$ = "CH 0"
               SetSwitch SelBoard%, XCmd$
               Exit For
            End If
         Next
   End Select

End Sub

Sub ConfigCtrlBoard(Device As Integer)

   If Device = -1 Then
      FirstDevice% = 0
      LastDevice% = mnNumCtlrs - 1
   Else
      FirstDevice% = Device
      LastDevice% = Device
   End If
   
   For DevIndex% = FirstDevice% To LastDevice%
      BdClass% = mauMCtl(DevIndex%).BoardClass
      BdNum% = mauMCtl(DevIndex%).BoardNum
      ULStat = StopBackground520(BdNum%, AOFUNCTION)
      If Not (ULStat = 0) Then Exit Sub
      ULStat = StopBackground520(BdNum%, AIFUNCTION)
      ULStat = StopBackground520(BdNum%, DIFUNCTION)
      ULStat = StopBackground520(BdNum%, DOFUNCTION)
      ULStat = StopBackground520(BdNum%, CTRFUNCTION)
      ULStat = StopBackground520(BdNum%, DAQIFUNCTION)
      ULStat = StopBackground520(BdNum%, DAQOFUNCTION)
      Select Case BdClass%
         Case BC2500AO
            mnDataType = 2 'sine
            ULStat = cbAOut(BdNum%, 0, BIP10VOLTS, -32768)
            ULStat = cbAOut(BdNum%, 1, BIP10VOLTS, -32768)
            mnGenFunc = A_OUTSCAN
            ULStat = cbTimerOutStop(BdNum%, 0)
         Case BC2500AI
            ULStat = cbTimerOutStop(BdNum%, 0)
         Case BCALTAO
            mnDataType = 2 'sine
            ULStat = cbAOut(BdNum%, 0, BIP10VOLTS, -32768)
            ULStat = cbAOut(BdNum%, 1, BIP10VOLTS, -32768)
            mnGenFunc = A_OUTSCAN
            BoardName$ = mauMCtl(DevIndex%).BoardName
            mvCustomRange = GetCustomRange(BoardName$)
         Case BC9513
            'set ctr5 as source for ctr1 and establish clock rate
            BaseCounter% = mauMCtl(DevIndex%).DevBase
            InfoType% = BOARDINFO
            DevNum% = 0
            ConfigItem% = BICLOCK
            ULStat = GetConfig520(InfoType%, BdNum%, DevNum%, ConfigItem%, ValConfig&)
            If Not (ValConfig& = 0) Then mlCtrSrc = ValConfig&
            ULStat = cbC9513Init(BdNum%, 1, 1, 1, 0, 0, 1)
            'main source counter - used for input for first three ctrs in some cases
            'set clock for 100kHz initially (resulting in 10us resolution on ctr1)
            CounterNum% = BaseCounter% + 4
            ULStat = cbC9513Config(BdNum%, CounterNum%, NOGATE, POSITIVEEDGE, _
            FREQ1, CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
            LoadVal& = (mlCtrSrc * 10) / 2
            ULStat = cbCLoad32(BdNum%, CounterNum%, LoadVal&)
            'alternate source counter - used for input for 2nd & 3rd ctrs in some cases
            'set clock for 100kHz initially (resulting in 10us resolution on ctr1)
            CounterNum% = BaseCounter% + 3
            ULStat = cbC9513Config(BdNum%, CounterNum%, NOGATE, POSITIVEEDGE, _
            FREQ1, CBDISABLED, LOADREG, RECYCLE, CBDISABLED, COUNTDOWN, TOGGLEONTC)
            LoadVal& = (mlCtrSrc * 10) / 2
            ULStat = cbCLoad32(BdNum%, CounterNum%, LoadVal&)
            'set ctr1, ctr2, and ctr3 output to high impedance
            ULStat = cbC9513Config(BdNum%, BaseCounter%, NOGATE, POSITIVEEDGE, _
            CTRINPUT1, CBDISABLED, LOADREG, ONETIME, CBDISABLED, COUNTDOWN, DISCONNECTED)
            ULStat = cbC9513Config(BdNum%, BaseCounter% + 1, NOGATE, NEGATIVEEDGE, _
            CTRINPUT1, CBDISABLED, LOADREG, ONETIME, CBDISABLED, COUNTDOWN, DISCONNECTED)
            ULStat = cbC9513Config(BdNum%, BaseCounter% + 2, NOGATE, NEGATIVEEDGE, _
            CTRINPUT1, CBDISABLED, LOADREG, ONETIME, CBDISABLED, COUNTDOWN, DISCONNECTED)
            ULStat = cbC9513Config(BdNum%, BaseCounter% + 3, NOGATE, NEGATIVEEDGE, _
            CTRINPUT1, CBDISABLED, LOADREG, ONETIME, CBDISABLED, COUNTDOWN, DISCONNECTED)
            mnOutCtl = HIGHPULSEONTC
            mlDuty = 0
            mnLoadReg = LOADREG
            mlTriggerDelay = 2
         Case BCTRIGGER
            ULStat = cbPulseOutStop(BdNum%, 0)
         Case BCAUXSWITCH, BCXPORTSWITCH
            If Not Device = -1 Then
               For SwitchPort& = mlSwitchPort To mlSwitchPortEnd
                  ULStat = cbDOut(BdNum%, mlSwitchPort, 0)
               Next
            End If
         Case BCAUXSELECT, BCPORTSELECT
            If Not Device = -1 Then
               For XCtlPort& = mlXCtlPort To mlXCtlPortEnd
                  ULStat = cbDOut(BdNum%, XCtlPort&, 0)
               Next
            End If
         Case BCPORTLOOP
            If Not Device = -1 Then
               For LoopPort& = mlLoopPort To mlLoopPortEnd
                  ULStat = cbDOut(BdNum%, LoopPort&, 0)
               Next
            End If
      End Select
   Next
   
End Sub

Sub DestroyCtrlBoards()

   mnNumCtlrs = 0

End Sub

Sub SetOffsetTweak(Offset As Single)

   mfOffSetTweak = Offset
   
End Sub

Public Function GetOffsetTweak() As Single

   GetOffsetTweak = mfOffSetTweak
   
End Function

Public Sub SetAiMap(MapString As String)

   If Left$(MapString, 2) = "-1" Then
      DefaultMap% = True
      ChanSpec& = InStr(1, MapString, ";")
      If Not ChanSpec& = 0 Then
         'default mapping with high channel count
         MapString = Mid$(MapString, ChanSpec& + 1)
         LastChan& = Val(MapString)
      Else
         'default mapping with default channel count
         LastChan& = 15
      End If
   Else
      If (MapString = "") Then
         'set default relay map
         LastChan& = 15
         DefaultMap% = True
      Else
         MapArray = Split(MapString, ";")
         NumChans& = UBound(MapArray)
         If NumChans& = 0 Then
            LastChan& = Val(MapString)
            DefaultMap% = True
         Else
            LastChan& = NumChans&
            DefaultMap% = False
         End If
      End If
   End If
   
   ReDim AiRelayMap(LastChan&)
   If DefaultMap% Then
      For RChan& = 0 To LastChan&
         AiRelayMap(RChan&) = RChan&
      Next
   Else
      ReDim AiRelayMap(NumChans&)
      For RChan& = 0 To NumChans&
         MapChan& = Val(MapArray(RChan&))
         AiRelayMap(RChan&) = MapChan&
      Next
   End If
   mnDefaultMap = True
   
End Sub

Public Function GetGPIBSurrogate(ByVal DevType As String) As String

   NumTypes& = mnNumCtlrs - 1 'UBound(mauMCtl)
   For CurType& = 0 To NumTypes&
      If mauMCtl(CurType&).GPIBName = DevType Then
         BoardName$ = mauMCtl(CurType&).BoardName
         Exit For
      End If
   Next
   GetGPIBSurrogate = BoardName$
   
End Function
