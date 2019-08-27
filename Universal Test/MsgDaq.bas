Attribute VB_Name = "MsgDaq"
Global Const NAME_ONLY = 0
Global Const NAME_SERNO = 1
Global Const NAME_ID = 2
Global Const NAME_SERNO_ID = 3

Dim mnNumBoards As Integer
Dim msBoardString As String
Dim mnInitialized As Integer, mnNameDef As Integer
Dim moNoForm As Form

Function InitMsgDaq() As Integer
   
   On Error GoTo NoMsgAPI
   If mnInitialized Then
      InitMsgDaq = mnInitialized
      Exit Function
   End If

#If MSGOPS Then
   Dim MsgInt As MBDClass.MBDComClass
   If MsgInt Is Nothing Then Set MsgInt = New MBDClass.MBDComClass
#Else
   Preface$ = "The current configuration for Universal Test is for the MBD Library, "
   MsgBox Preface$ & "but this version of Universal Test is not MBD compatible. " & _
   "Select a different library using the File menu.", _
   vbCritical, "MBD Library Not Available"
   gnLibType = INVALIDLIB
   gnErrFlag = True
   Exit Function
#End If
   
   msBoardString = MsgInt.FindDevices(mnNameDef)
   NameFormat$ = GetNameFormatString(mnNameDef)
   If SaveMsg(moNoForm, "GetDeviceNames(" & NameFormat$ & ")", msBoardString) Then Exit Function

   BoardArray = Split(msBoardString, "|")
   NumBoards% = UBound(BoardArray)
   If NumBoards% < 0 Then NumBoards% = 0
   mnNumBoards = NumBoards%
   mnInitialized = True
   InitMsgDaq = mnInitialized
   Set MsgInt = Nothing
   gnLibType = MSGLIB
   Exit Function
   
NoMsgAPI:
   If gnLibType = MSGLIB Then
      Preface$ = "The current configuration for Universal Test is for the MBD Library, "
   Else
      Preface$ = "You attempted to set the current Universal Test libary to the MBD Library, "
   End If
   gnLibType = INVALIDLIB
   MsgBox Preface$ & _
   "but the MBD Library was not detected. " & vbCrLf & "Select a different library using the File menu.", _
   vbCritical, "MBD Library Not Available"
   Exit Function
   
End Function

Function GetNumMsgBoards() As Integer
   
   'Get number of message-based boards installed
   If mnInitialized Then
      GetNumMsgBoards = mnNumBoards
   Else
      If InitMsgDaq Then
         GetNumMsgBoards = mnNumBoards
      Else
         GetNumMsgBoards = 0
      End If
   End If

End Function

Function GetNameOfMsgBoard(BoardNum As Integer) As String

   If Not mnInitialized Then
      If Not InitMsgDaq() Then
         GetNameOfMsgBoard = "Could not initialize MBD library.|"
         Exit Function
      End If
   End If
   
   If Len(msBoardString) > 1 Then
      BoardArray = Split(msBoardString, "|")
      GetNameOfMsgBoard = BoardArray(BoardNum)
   Else
      GetNameOfMsgBoard = "None Installed"
   End If

End Function

Function GetNameFormatString(ByVal NameFormat As Integer) As String

   Select Case NameFormat
      Case NAME_ONLY
         FormatString$ = "NameOnly"
      Case NAME_SERNO
         FormatString$ = "NameAndSerno"
      Case NAME_ID
         FormatString$ = "NameAndID"
      Case NAME_SERNO_ID
         FormatString$ = "NameSernoAndID"
   End Select
   GetNameFormatString = FormatString$
   
End Function

Function GetMsgChMode(ModeCode As Integer) As String

   ModeString$ = Choose(ModeCode + 1, "DIFF", "SE")
   GetMsgChMode = ModeString$
   
End Function

Function GetMsgRange(RangeCode As Integer) As String

   Select Case RangeCode
      Case BIP5VOLTS
         RangeString$ = "BIP5V"
      Case BIP10VOLTS
         RangeString$ = "BIP10V"
      Case BIP2PT5VOLTS
         RangeString$ = "BIP2.5V"
      Case BIP1PT25VOLTS
         RangeString$ = "BIP1.25V"
      Case BIP1VOLTS
         RangeString$ = "BIP1V"
      Case BIP2VOLTS
         RangeString$ = "BIP2V"
      Case BIP20VOLTS
         RangeString$ = "BIP20V"
      Case BIP4VOLTS
         RangeString$ = "BIP4V"
      Case BIPPT625VOLTS
         RangeString$ = "BIP625.0E-3V"
      Case BIPPT312VOLTS
         RangeString$ = "BIP312.5E-3V"
      Case BIPPT156VOLTS
         RangeString$ = "BIP156.25E-3V"
      Case BIPPT078VOLTS
         RangeString$ = "BIP78.125E-3V"
      Case BIPPT073125VOLTS
         RangeString$ = "BIP73.125E-3V"
      Case UNI5VOLTS
         RangeString$ = "UNI5V"
      Case UNI4VOLTS
         RangeString$ = "UNI4.096V"
      Case Else
         RangeString$ = "Unsupported"
   End Select
   GetMsgRange = RangeString$
   
End Function

Function GetOptionCodeFromMsg(OptionMessage As String, OptionIndex As Integer) As Long
   
   Select Case OptionMessage
      Case "BACKGROUND"
         OptionCode& = BACKGROUND
         OptionIndex = 0
      Case "BLOCKIO"
         OptionCode& = BLOCKIO
         OptionIndex = 7
      Case "SINGLEIO"
         OptionCode& = SINGLEIO
         OptionIndex = 5
      Case "BURSTIO"
         OptionCode& = BURSTIO
         OptionIndex = 16
      Case "EXTPACER" 'EXTSYNC
         OptionCode& = EXTCLOCK
         OptionIndex = 2
      Case "TRIG"
         OptionCode& = EXTTRIGGER
         OptionIndex = 14
      Case "CAL"
         OptionCode& = NOCALIBRATEDATA
         OptionIndex = 15
      Case Else
         OptionCode& = 0
         OptionIndex = -1
   End Select
   GetOptionCodeFromMsg = OptionCode&

End Function

Function GetRangeCodeFromMsg(RangeMessage As String) As Integer

   Select Case RangeMessage
      Case "BIP5V"
         RangeCode% = BIP5VOLTS
      Case "BIP10V"
         RangeCode% = BIP10VOLTS
      Case "BIP2.5V"
         RangeCode% = BIP2PT5VOLTS
      Case "BIP1.25V"
         RangeCode% = BIP1PT25VOLTS
      Case "BIP1V"
         RangeCode% = BIP1VOLTS
      Case "BIP2V"
         RangeCode% = BIP2VOLTS
      Case "BIP20V"
         RangeCode% = BIP20VOLTS
      Case "BIP4V"
         RangeCode% = BIP4VOLTS
      Case "BIP625.0E-3V"
         RangeCode% = BIPPT625VOLTS
      Case "BIP312.5E-3V"
         RangeCode% = BIPPT312VOLTS
      Case "BIP156.25E-3V"
         RangeCode% = BIPPT156VOLTS
      Case "BIP78.125E-3V"
         RangeCode% = BIPPT078VOLTS
      Case "BIP73.125E-3V"
         RangeCode% = BIPPT073125VOLTS
      Case "UNI5V"
         RangeCode% = UNI5VOLTS
      Case "UNI4.096V"
         RangeCode% = UNI4VOLTS
      Case Else
         RangeCode% = NOTUSED
   End Select
   GetRangeCodeFromMsg = RangeCode%
   
End Function

Function GetTrigTypeFromMsg(TrigMessage As String) As Integer

   Select Case TrigMessage
      Case "ABOVE"
         TrigType% = TRIGABOVE
      Case "BELOW"
         TrigType% = TRIGBELOW
      Case "LOWHYST"
         TrigType% = GATENEGHYS
      Case "HIGHHYST"
         TrigType% = GATEPOSHYS
      Case "GATEABOVE"
         TrigType% = GATEABOVE
      Case "GATEBELOW"
         TrigType% = GATEBELOW
      Case "GATEINWINDOW"
         TrigType% = GATEINWINDOW
      Case "GATEOUTWINDOW"
         TrigType% = GATEOUTWINDOW
      Case "GATEHIGH"
         TrigType% = GATEHIGH
      Case "GATELOW"
         TrigType% = GATELOW
      Case "HIGH"
         TrigType% = TRIGHIGH
      Case "LOW"
         TrigType% = TRIGLOW
      Case "EDGE/RISING"
         TrigType% = TRIGPOSEDGE
      Case "EDGE/FALLING"
         TrigType% = TRIGNEGEDGE
   End Select
   GetTrigTypeFromMsg = TrigType%

End Function

Function GetMsgTrigTypeString(ULTrigTypeValue) As String

   Select Case ULTrigTypeValue
      Case TRIGABOVE
         TrigTypeString$ = "ABOVE"
      Case TRIGBELOW
         TrigTypeString$ = "BELOW"
      Case GATENEGHYS
         TrigTypeString$ = "LOWHYST"
      Case GATEPOSHYS
         TrigTypeString$ = "HIGHHYST"
      Case GATEABOVE
         TrigTypeString$ = "GATEABOVE"
      Case GATEBELOW
         TrigTypeString$ = "GATEBELOW"
      Case GATEINWINDOW
         TrigTypeString$ = "GATEINWINDOW"
      Case GATEOUTWINDOW
         TrigTypeString$ = "GATEOUTWINDOW"
      Case GATEHIGH
         TrigTypeString$ = "GATEHIGH"
      Case GATELOW
         TrigTypeString$ = "GATELOW"
      Case TRIGHIGH
         TrigTypeString$ = "LEVEL/HIGH"
      Case TRIGLOW
         TrigTypeString$ = "LEVEL/LOW"
      Case TRIGPOSEDGE
         TrigTypeString$ = "EDGE/RISING"
      Case TRIGNEGEDGE
         TrigTypeString$ = "EDGE/FALLING"
   End Select
   GetMsgTrigTypeString = TrigTypeString$

End Function

Function GetAIScanProps(DeviceID As String, Device As Object, PropsList As Variant) As Integer

   Dim PropArray() As String
   Dim MBResponse As VbMsgBoxResult
   Component$ = "AISCAN"
   MsgSupport$ = Device.GetSupportedMessages(Component$)
   ScanSupport% = Not (MsgSupport$ = "")
   'don't trap warnings here
   SaveWarnFlow% = geWarnFlow
   SaveWarnDisp% = gnLocalWarnDisp
   geWarnFlow = 0
   gnLocalWarnDisp = 0
   If ScanSupport% Then
      ScanParams = Split(MsgSupport$, "|")
      If IsEmpty(ScanParams) Then
         NumParams& = -1
      Else
         NumParams& = UBound(ScanParams)
      End If
      For i% = 0 To NumParams&
         If InStr(1, ScanParams(i%), "?") = 1 Then
            XfrModeMsg$ = ScanParams(i%)
            MsgResult$ = Device.SendMessage(XfrModeMsg$)
            If SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) Then
               MBResponse = MsgBox(MsgResult$ & vbCrLf & "( " & XfrModeMsg$ & _
               " )" & vbCrLf & "Continue checking properties?", vbYesNo, "Message DAQ Error")
               NoSupportLoc& = InStr(1, MsgResult$, "An unknown error has occurred")
                  If Not (NoSupportLoc& = 0) Then
                  ReDim Preserve PropArray(NumProps%)
                  PropArray(NumProps%) = PropMsg$ & "=NOTSUPPORTED"
                  NumProps% = NumProps% + 1
                  End If
               If MBResponse = vbNo Then Exit For
            Else
               ValueLoc& = InStr(1, MsgResult$, ":")
               NoSupportLoc& = InStr(1, MsgResult$, "does not support the command")
               NoSupportLoc& = NoSupportLoc& + InStr(1, MsgResult$, "message sent is not supported")
               If (ValueLoc& > 0) And (ValueLoc& < 18) Then
                  ReDim Preserve PropArray(NumProps%)
                  PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + 1)
                  NumProps% = NumProps% + 1
               Else
                  If Not (NoSupportLoc& = 0) Then
                     ReDim Preserve PropArray(NumProps%)
                     PropArray(NumProps%) = PropMsg$ & "=NOTSUPPORTED"
                     NumProps% = NumProps% + 1
                  End If
               End If
            End If
         End If
      Next
      
      Component$ = "AITRIG"
      MsgSupport$ = Device.GetSupportedMessages(Component$)
      TrigSupport% = Not (MsgSupport$ = "")
      If TrigSupport% Then
         TrigParams = Split(MsgSupport$, "|")
         If IsEmpty(TrigParams) Then
            NumParams& = -1
         Else
            NumParams& = UBound(TrigParams)
         End If
         For i% = 0 To NumParams&
            If InStr(1, TrigParams(i%), "?") = 1 Then
               TrigString$ = TrigParams(i%)
               MsgResult$ = Device.SendMessage(TrigString$)
               x% = SaveMsg(moNoForm, "SendMessage(" & TrigString$ & ")", MsgResult$)
               ValueLoc& = InStr(1, MsgResult$, ":")
               If (ValueLoc& > 0) And (ValueLoc& < 18) Then
                  ReDim Preserve PropArray(NumProps%)
                  PropArray(NumProps%) = MsgResult$
                  NumProps% = NumProps% + 1
               End If
            End If
         Next
      End If
   End If
   PropsList = PropArray()
   GetAIScanProps = NumProps% - 1
   geWarnFlow = SaveWarnFlow%
   gnLocalWarnDisp = SaveWarnDisp%
   
End Function

Function GetAIProps(DeviceID As String, Device As Object, PropsList As Variant) As Integer

   Dim PropArray() As String
   Component$ = "AI"
   MsgSupport$ = Device.GetSupportedMessages(Component$)
   AiSupport% = Not (MsgSupport$ = "")
   'don't trap warnings here
   SaveWarnFlow% = geWarnFlow
   SaveWarnDisp% = gnLocalWarnDisp
   geWarnFlow = 0
   gnLocalWarnDisp = 0
   If AiSupport% Then
      AiParams = Split(MsgSupport$, "|")
      If IsEmpty(AiParams) Then
         NumParams& = -1
      Else
         NumParams& = UBound(AiParams)
      End If
      For i% = 0 To NumParams&
         If InStr(1, AiParams(i%), "?") = 1 Then
            XfrModeMsg$ = AiParams(i%)
            XfrModeMsg$ = Replace(XfrModeMsg$, "{*}", "{0}")
            XfrModeMsg$ = Replace(XfrModeMsg$, "/*", "/VOLTS")
            MsgResult$ = Device.SendMessage(XfrModeMsg$)
            If SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) Then
            Else
               ValueLoc& = InStr(1, MsgResult$, ":")
               If ValueLoc& = 0 Then
                  ValueLoc& = InStr(1, MsgResult$, "=")
                  NoPropReturned% = True
               End If
               NoSupportLoc& = InStr(1, MsgResult$, "does not support the command")
               If (ValueLoc& > 0) And (ValueLoc& < 18) Then
                  ReDim Preserve PropArray(NumProps%)
                  If NoPropReturned% Then ValueLoc& = 0
                  PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + 1)
                  NumProps% = NumProps% + 1
               Else
                  If Not (NoSupportLoc& = 0) Then
                     ReDim Preserve PropArray(NumProps%)
                     PropArray(NumProps%) = PropMsg$ & "=NOTSUPPORTED"
                     NumProps% = NumProps% + 1
                  End If
               End If
            End If
         End If
      Next
   End If
   PropsList = PropArray()
   GetAIProps = NumProps% - 1
   geWarnFlow = SaveWarnFlow%
   gnLocalWarnDisp = SaveWarnDisp%
   
End Function

Function GetAOProps(DeviceID As String, Device As Object, PropsList As Variant) As Integer

   Dim PropArray() As String
   'don't trap warnings here
   SaveWarnFlow% = geWarnFlow
   SaveWarnDisp% = gnLocalWarnDisp
   geWarnFlow = 0
   gnLocalWarnDisp = 0
   CompMsg$ = "?AO"
   For i% = 1 To 3
      PropMsg$ = Choose(i%, "", "{0}:SCALE", ":RANGE{0}")
      XfrModeMsg$ = CompMsg$ & PropMsg$
      MsgResult$ = Device.SendMessage(XfrModeMsg$)
      x% = SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) 'Then
      LocOffset& = 1
      ValueLoc& = InStr(1, MsgResult$, ":")
      If ValueLoc& = 0 Then
         ValueLoc& = InStr(1, MsgResult$, "=")
         If ValueLoc& > 0 Then
            ValueLoc& = 1
            LocOffset& = 0
         End If
      End If
      If ValueLoc& > 0 Then
         ReDim Preserve PropArray(NumProps%)
         PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
         NumProps% = NumProps% + 1
      End If
   Next
   PropsList = PropArray()
   GetAOProps = NumProps% - 1
   geWarnFlow = SaveWarnFlow%
   gnLocalWarnDisp = SaveWarnDisp%
   
End Function

Function GetTmrProps(DeviceID As String, Device As Object, PropsList As Variant) As Integer
   
   Dim PropArray() As String
   'don't trap warnings here
   SaveWarnFlow% = geWarnFlow
   SaveWarnDisp% = gnLocalWarnDisp
   geWarnFlow = 0
   gnLocalWarnDisp = 0
   CompMsg$ = "?TMR"
   MsgResult$ = Device.SendMessage(CompMsg$)
   x% = SaveMsg(moNoForm, "SendMessage(" & CompMsg$ & ")", MsgResult$)
   LocOffset& = 1
   ValueLoc& = InStr(1, MsgResult$, ":")
   If ValueLoc& = 0 Then
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         ValueLoc& = 1
         LocOffset& = 0
      End If
   End If
   If ValueLoc& > 0 Then
      ReDim Preserve PropArray(NumProps%)
      PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
      NumProps% = NumProps% + 1
      NumLoc& = InStr(1, MsgResult$, "=")
      If NumLoc& > 0 Then
         PortString$ = Mid(MsgResult$, NumLoc& + 1)
         NumCtrs% = Val(PortString$)
      End If
   End If
   For i% = 1 To NumCtrs%
      For PropItem% = 1 To 5
         PropName$ = Choose(PropItem%, "PERIOD", "DUTYCYCLE", "DELAY", "PULSECOUNT", "IDLESTATE")
         PropMsg$ = "{" & i% - 1 & "}:" & PropName$
         XfrModeMsg$ = CompMsg$ & PropMsg$
         MsgResult$ = Device.SendMessage(XfrModeMsg$)
         x% = SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) 'Then
         LocOffset& = 1
         ValueLoc& = InStr(1, MsgResult$, ":")
         If ValueLoc& = 0 Then
            ValueLoc& = InStr(1, MsgResult$, "=")
            If ValueLoc& > 0 Then
               ValueLoc& = 1
               LocOffset& = 0
            End If
         End If
         If ValueLoc& > 0 Then
            ReDim Preserve PropArray(NumProps%)
            PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
            NumProps% = NumProps% + 1
         End If
      Next
   Next
   PropsList = PropArray()
   GetTmrProps = NumProps% - 1
   geWarnFlow = SaveWarnFlow%
   gnLocalWarnDisp = SaveWarnDisp%

End Function

Function GetCtrProps(DeviceID As String, Device As Object, PropsList As Variant) As Integer

   Dim PropArray() As String
   'don't trap warnings here
   SaveWarnFlow% = geWarnFlow
   SaveWarnDisp% = gnLocalWarnDisp
   geWarnFlow = 0
   gnLocalWarnDisp = 0
   CompMsg$ = "?CTR"
   MsgResult$ = Device.SendMessage(CompMsg$)
   x% = SaveMsg(moNoForm, "SendMessage(" & CompMsg$ & ")", MsgResult$)
   LocOffset& = 1
   ValueLoc& = InStr(1, MsgResult$, ":")
   If ValueLoc& = 0 Then
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         ValueLoc& = 1
         LocOffset& = 0
      End If
   End If
   If ValueLoc& > 0 Then
      ReDim Preserve PropArray(NumProps%)
      PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
      NumProps% = NumProps% + 1
      NumLoc& = InStr(1, MsgResult$, "=")
      If NumLoc& > 0 Then
         PortString$ = Mid(MsgResult$, NumLoc& + 1)
         NumCtrs% = Val(PortString$)
      End If
   End If
   If 0 Then
         For i% = 1 To NumCtrs%
            PropMsg$ = "{" & i% - 1 & "}"
            XfrModeMsg$ = CompMsg$ & PropMsg$
            MsgResult$ = Device.SendMessage(XfrModeMsg$)
            x% = SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) 'Then
            LocOffset& = 1
            ValueLoc& = InStr(1, MsgResult$, ":")
            If ValueLoc& = 0 Then
               ValueLoc& = InStr(1, MsgResult$, "=")
               If ValueLoc& > 0 Then
                  ValueLoc& = 1
                  LocOffset& = 0
               End If
            End If
            If ValueLoc& > 0 Then
               ReDim Preserve PropArray(NumProps%)
               PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
               NumProps% = NumProps% + 1
            End If
         Next
   End If
   PropsList = PropArray()
   GetCtrProps = NumProps% - 1
   geWarnFlow = SaveWarnFlow%
   gnLocalWarnDisp = SaveWarnDisp%
   
End Function

Function GetDIOProps(DeviceID As String, Device As Object, PropsList As Variant) As Integer

   Dim PropArray() As String
   'don't trap warnings here
   SaveWarnFlow% = geWarnFlow
   SaveWarnDisp% = gnLocalWarnDisp
   geWarnFlow = 0
   gnLocalWarnDisp = 0
   CompMsg$ = "?DIO"
   MsgResult$ = Device.SendMessage(CompMsg$)
   x% = SaveMsg(moNoForm, "SendMessage(" & CompMsg$ & ")", MsgResult$)
   LocOffset& = 1
   ValueLoc& = InStr(1, MsgResult$, ":")
   If ValueLoc& = 0 Then
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         ValueLoc& = 1
         LocOffset& = 0
      End If
   End If
   If ValueLoc& > 0 Then
      ReDim Preserve PropArray(NumProps%)
      PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
      NumProps% = NumProps% + 1
      NumLoc& = InStr(1, MsgResult$, "=")
      If NumLoc& > 0 Then
         PortString$ = Mid(MsgResult$, NumLoc& + 1)
         NumPorts% = Val(PortString$)
      End If
   End If
   For i% = 1 To NumPorts%
      PropMsg$ = "{" & i% - 1 & "}" 'number of bits in port
      XfrModeMsg$ = CompMsg$ & PropMsg$
      MsgResult$ = Device.SendMessage(XfrModeMsg$)
      x% = SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) 'Then
      LocOffset& = 1
      ValueLoc& = InStr(1, MsgResult$, ":")
      If ValueLoc& = 0 Then
         ValueLoc& = InStr(1, MsgResult$, "=")
         If ValueLoc& > 0 Then
            ValueLoc& = 1
            LocOffset& = 0
         End If
      End If
      If ValueLoc& > 0 Then
         ReDim Preserve PropArray(NumProps%)
         PropArray(NumProps%) = Mid(MsgResult$, ValueLoc& + LocOffset&)
         NumProps% = NumProps% + 1
      End If
      PortString$ = "{" & i% - 1 & "}"
      PropMsg$ = PortString$ & ":DIR" 'port configuration (IN, OUT, or bit field)
      XfrModeMsg$ = CompMsg$ & PropMsg$
      MsgResult$ = Device.SendMessage(XfrModeMsg$)
      x% = SaveMsg(moNoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$) 'Then
      LocOffset& = 1
      ValueLoc& = InStr(1, MsgResult$, ":")
      If ValueLoc& = 0 Then
         ValueLoc& = InStr(1, MsgResult$, "=")
         If ValueLoc& > 0 Then
            ValueLoc& = 1
            LocOffset& = 0
         End If
      End If
      If ValueLoc& > 0 Then
         ReDim Preserve PropArray(NumProps%)
         PropArray(NumProps%) = MsgResult$   'Mid(MsgResult$, ValueLoc& + LocOffset&)
         NumProps% = NumProps% + 1
      End If
   Next
   PropsList = PropArray()
   GetDIOProps = NumProps% - 1
   geWarnFlow = SaveWarnFlow%
   gnLocalWarnDisp = SaveWarnDisp%
   
End Function

Function MsgCheckStatus(ByVal CallingForm As Form, ByVal DeviceID As String, _
MsgLibrary As Object, ByVal CompType As String, _
CurrentCount As Long, CurrentIndex As Long) As Integer
     
   Dim Resp As VbMsgBoxResult
   If MsgLibrary Is Nothing Then Exit Function
   StatMsg$ = "?" & CompType & ":STATUS"
   MsgResult$ = MsgLibrary.SendMessage(StatMsg$)
   MsgAck& = InStr(1, MsgResult$, "STATUS")
   If Not (MsgAck& > 0) Then
      MsgQualifier$ = "Error: "
   Else
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         StatString$ = Mid(MsgResult$, ValueLoc& + 1)
         Status% = Switch(StatString$ = "IDLE", IDLE, _
         StatString$ = "RUNNING", RUNNING, StatString$ = "OVERRUN", OVERRUN, _
         StatString$ = "UNDERRUN", UNDERRUN, StatString$ = "INTERRUPTED", _
         INTERRUPTED, StatString$ <> "", -1)
      Else
         Status% = -2
      End If
      If Status% = OVERRUN Then
         ULStat = OVERRUN
         MsgQualifier$ = "Error: "
      End If
      If Status% = UNDERRUN Then
         ULStat = UNDERRUN
         MsgQualifier$ = "Error: "
      End If
      If Status% = INTERRUPTED Then
         ULStat = INTERRUPTED
         MsgQualifier$ = "Error: "
      End If
   End If
   If SaveMsg(CallingForm, "SendMessage(" & StatMsg$ & ")", _
   MsgQualifier$ & MsgResult$) Then Exit Function
   If Status% = -1 Then
      ChrsReturned& = Len(StatString$)
      For Character& = 1 To ChrsReturned&
         AscValue% = Asc(Mid(StatString$, Character&, 1))
         ChrList$ = ChrList$ & Format(AscValue%, "0") & " "
         If AscValue% < 32 Then
            BadString$ = BadString$ & "_"
         Else
            BadString$ = BadString$ & Chr(AscValue%)
         End If
      Next
      MsgBox "Corrupt message returned." & vbCrLf & BadString$ & vbCrLf & ChrList$
   End If
   'If Status% = RUNNING Then
      CountMsg$ = "?" & CompType & ":COUNT"
      IndexMsg$ = "?" & CompType & ":INDEX"
      MsgResult$ = MsgLibrary.SendMessage(CountMsg$)
      If SaveMsg(CallingForm, "SendMessage(" & CountMsg$ & ")", MsgResult$) Then Exit Function
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         CountString$ = Mid(MsgResult$, ValueLoc& + 1)
         If IsNumeric(CountString$) Then
            CurCount& = Val(CountString$)
         Else
            If gnIDERunning Then
               Stop
            Else
               Resp = MsgBox("This path is a Stop statement " & _
               "in the IDE. Check Local Error Handling options. " _
               & vbCrLf & vbCrLf & "          Click Yes to attempt " & _
               "to continue, No to exit application.", _
               vbYesNo, "Attempt To Continue?")
               If Resp = vbNo Then End
            End If
         End If
      Else
         CurCount& = 0
      End If
      MsgResult$ = MsgLibrary.SendMessage(IndexMsg$)
      If SaveMsg(CallingForm, "SendMessage(" & IndexMsg$ & ")", MsgResult$) Then Exit Function
      ValueLoc& = InStr(1, MsgResult$, "=")
      If ValueLoc& > 0 Then
         IndexString$ = Mid(MsgResult$, ValueLoc& + 1)
         If IsNumeric(IndexString$) Then
            CurIndex& = Val(IndexString$)
         Else
            If gnIDERunning Then
               Stop
            Else
               Resp = MsgBox("This path is a Stop statement " & _
               "in the IDE. Check Local Error Handling options. " _
               & vbCrLf & vbCrLf & "          Click Yes to attempt " & _
               "to continue, No to exit application.", _
               vbYesNo, "Attempt To Continue?")
               If Resp = vbNo Then End
            End If
         End If
      Else
         CurIndex& = 0
      End If
   'End If
   
   CurrentCount = CurCount&
   CurrentIndex = CurIndex&
   MsgCheckStatus = Status%

End Function

Function GetAllMsgBoardNames() As String

   If Not mnInitialized Then
      If Not InitMsgDaq() Then
         GetAllMsgBoardNames = "Could not initialize MBD library.|"
         Exit Function
      End If
   End If
   
   If Len(msBoardString) > 1 Then
      BoardNames$ = msBoardString
   Else
      BoardNames$ = "None installed.|"
   End If
   GetAllMsgBoardNames = BoardNames$

End Function

Sub SetNameFormat(ByVal NameFormat As Integer)
   
   mnNameDef = NameFormat

End Sub

Function GetNameFormat() As Integer
   
   GetNameFormat = mnNameDef

End Function

Function GetPacerStrings(ByVal Device As Object, ByVal Component As String, _
ByVal PacerState As Integer) As String
   
   Dim NoForm As Form
   
   Select Case PacerState
      Case 0
         Default$ = "DISABLE"
         Alternate$ = "MASTER"
      Case 1
         Default$ = "SLAVE"
         Alternate$ = "ENABLE"
   End Select
   
   XfrModeMsg$ = "@" & Component & ":EXTPACER"
   MsgResult$ = Device.SendMessage(XfrModeMsg$)
   x% = SaveMsg(NoForm, "SendMessage(" & XfrModeMsg$ & ")", MsgResult$)
   ValidArray = Split(MsgResult$, "%")
   If InStr(1, MsgResult$, "NOT_SUPPORTED") > 0 Then
      ClockOff$ = Alternate$
   Else
      If UBound(ValidArray) > 0 Then
         ChoiceList$ = ValidArray(1)
         ChoiceArray = Split(ChoiceList$, ",")
         Choices& = UBound(ChoiceArray)
         If InStr(1, ChoiceList$, Default$) > 0 Then
            For Choice& = 0 To Choices&
               CurChoice$ = ChoiceArray(Choice&)
               If InStr(1, CurChoice$, Default$) > 0 Then
                  ClockOff$ = CurChoice$
                  Exit For
               End If
            Next
         Else
            For Choice& = 0 To Choices&
               CurChoice$ = ChoiceArray(Choice&)
               If InStr(1, CurChoice$, Alternate$) > 0 Then
                  ClockOff$ = CurChoice$
                  Exit For
               End If
            Next
         End If
      End If
   End If
   GetPacerStrings = ClockOff$

End Function

Function GetMsgADResolution(ByVal Device As Object) As Integer
   
   Dim NoForm As Form
   ADResMsg$ = "@AI:MAXCOUNT"
   MsgResult$ = Device.SendMessage(ADResMsg$)
   x% = SaveMsg(NoForm, "SendMessage(" & ADResMsg$ & ")", MsgResult$)
   ValidArray = Split(MsgResult$, "%")
   If UBound(ValidArray) > 0 Then
      CountString$ = ValidArray(1)
      Select Case CountString$
         Case "4095"
            Resolution% = 12
         Case "16383"
            Resolution% = 14
         Case "65535"
            Resolution% = 16
         Case "16777215"
            Resolution% = 24
      End Select
   End If
   GetMsgADResolution = Resolution%

End Function


Function GetMsgDAResolution(ByVal Device As Object) As Integer
   
   Dim NoForm As Form
   DAResMsg$ = "@AO:MAXCOUNT"
   MsgResult$ = Device.SendMessage(DAResMsg$)
   x% = SaveMsg(NoForm, "SendMessage(" & DAResMsg$ & ")", MsgResult$)
   ValidArray = Split(MsgResult$, "%")
   If UBound(ValidArray) > 0 Then
      CountString$ = ValidArray(1)
      Select Case CountString$
         Case "4095"
            Resolution% = 12
         Case "65535"
            Resolution% = 16
         Case "16777215"
            Resolution% = 24
      End Select
   End If
   GetMsgDAResolution = Resolution%

End Function

