Attribute VB_Name = "GPIBInterface"
Global gnLevel As Integer
Global ganStatBits%(18), gasStatBits$(18)
Global ganErrCodes%(24), gasErrCode$(24)

Dim mnInitGPIB As Integer

Type GPIBCfg
   Name As String
   address As Integer
   Handle As Integer
End Type

Dim mauCfg() As GPIBCfg

Sub DoIBASK(Handle%, Index%, ItemName$, Param%, value%)

   If Handle% < 0 Then Handle% = gnLevel
   'If Not CheckInit() Then Exit Sub
   If Param% < 0 Then
      'frmIBConfig.lblCurValue.Visible = False
      'frmIBConfig.lblVal.Visible = False
      'frmIBConfig.Show 1
      'Param% = Val(frmIBConfig.lblCurOption.Caption)
      'Index% = Val(frmIBConfig.cmbOption.ListIndex)
      'ItemName$ = frmIBConfig.cmbOption.Text
      'Unload frmIBConfig
      If Param% = 0 Then Exit Sub
   End If

   'DoEvents
   If gnUseFunctions Then
      'If gbGPIBActive% Then
      '   Call WaitTillDone: Exit Sub
      'Else
      '   gbGPIBActive% = True
      'End If
      x% = ilask(Handle%, Param%, value%)
      gbGPIBActive% = False
      Dev$ = mauCfg(Device).Name
      PrintStatus GPIBAsk, Dev$, "DoIBASK (" & ItemName$ & "); ilask(" & "0x" & Hex$(Handle%) & ", 0x" & Hex$(Param%) & ", 0x" & Hex$(value%) & ") = 0x" & Hex$(x%), 0
   Else
      'If gbGPIBActive% Then
      '   Call WaitTillDone: Exit Sub
      'Else
      '   gbGPIBActive% = True
      'End If
      ibask Handle%, Param%, value%
      gbGPIBActive% = False
      PrintStatus GPIBAsk, Dev$, "DoIBASK (" & ItemName$ & "); ibask(" & "0x" & Hex$(Handle%) & ", 0x" & Hex$(Param%) & ", 0x" & Hex$(value%) & ")", 0
   End If
   
End Sub

Sub DoIBFIND(BDname$, BoardDev%)

   BoardDev% = -1
   If BDname$ = "" Then
      Func$ = "a name"
      WarnParam$ = "Init"
      'If Not PrintWarning(Func$, WarnParam$) Then Exit Sub
   End If
   
   'gnCICOnly = True
   'If gbGPIBActive% Then
   '   Call WaitTillDone: Exit Sub
   'Else
   '   gbGPIBActive% = True
   'End If
   If gnUseFunctions Then
      BoardDev% = ilfind(BDname$)
      If BoardDev% = -1 Then BoardDev% = ilfind(BDname$)
      gbGPIBActive% = False
      TestIt% = 1
      If ibsta% And EERR Then
         If iberr% = ECIC Then TestIt% = 0
         If iberr% = ENEB Then TestIt% = 0
         If iberr% = EDVR Then TestIt% = 0 'workaround for NI difference
      End If
      Dev$ = mauCfg(Device).Name
      PrintStatus GPFind, Dev$, "DoIBFIND; ilfind('" & BDname$ & "') = 0x" & Hex$(ud%), 1
   Else
      ibfind BDname$, BoardDev%
      'first call doesn't work in 16 bit library for some reason
      If BoardDev% = -1 Then ibfind BDname$, BoardDev%
      gbGPIBActive% = False
      TestIt% = 1
      If ibsta% And EERR Then
         If iberr% = ECIC Then TestIt% = 0
         If iberr% = ENEB Then TestIt% = 0
         If iberr% = EDVR Then TestIt% = 0 'workaround for NI difference
      End If
      Dev$ = mauCfg(Device).Name
      PrintStatus GPFind, Devs$, "DoIBFIND; ibfind('" & BDname$ & "', 0x" & Hex$(BoardDev%) & ")", TestIt%
   End If

End Sub

Sub DoIBSre(BoardDev%, ByVal Ren%)

   board% = mauCfg(BoardDev%).Handle
   Dev$ = mauCfg(BoardDev%).Name
   Ren% = Abs(Ren%) '-1 won't work in 32 bit
   ibsre board%, Ren%
   PrintStatus GPIBSre, Dev$, "ibsre (" & Format$(board%, "0") & ", " & Format$(Ren%, "0") & ")", True
   
End Sub

Sub FillDevList(Instance%)

   'frmNewGPIBCtl(Instance%).cmbDevice.Clear
   For BdDevIndex% = 0 To 33
      DeviceName$ = mauCfg(BdDevIndex%).Name
      HasName% = Not (DeviceName$ = "" Or DeviceName$ = "NOADDR")
      If HasName% Then
         frmNewGPIBCtl(Instance%).cmbDevice.AddItem mauCfg(BdDevIndex%).Name, BdDevIndex%
         GPIBCtlrs% = GPIBCtlrs% + 1
      End If
   Next BdDevIndex%
   If frmNewGPIBCtl(Instance%).cmbDevice.ListCount > 0 _
      Then frmNewGPIBCtl(Instance%).cmbDevice.ListIndex = 0
   frmNewGPIBCtl(Instance%).SetNumGPIBCtlrs GPIBCtlrs%

End Sub

Function GetEncoderCmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Init", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Direction", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Speed", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Index", 3
   Else
      Select Case CommandNum
         Case 0
            GetEncoderCmds = "Init "
         Case 1
            GetEncoderCmds = "Dir " & NUL$ & "+/-"
         Case 2
            GetEncoderCmds = "Rate " & NUL$ & "0/0.1/1/10/50/100/500" & NUL$ & " Hz/ kHz"
         Case 3
            GetEncoderCmds = "Index " & NUL$ & "0/8/16/32/64/128/256/512/1024/2048"
      End Select
   End If

End Function

Function GetIndexerCmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Delay (% of Cycle)", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Length (% of Cycle)", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "A-Cycle Edge", 2
   Else
      Select Case CommandNum
         Case 0
            GetIndexerCmds = "Delay " & NUL$ & "1/10/30/50/80"
         Case 1
            GetIndexerCmds = "Length " & NUL$ & "1/10/50/100/200/300"
         Case 2
            GetIndexerCmds = "Edge " & NUL$ & "+/-"
      End Select
   End If

End Function

Function GetTriggerCmds(Instance As Integer, CommandNum As Integer) As String
   
   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Normal", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Trigger", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable: Off", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable: On", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Period:", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Width:", 5
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Duty Cycle:", 6
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Complement On:", 7
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Complement Off:", 8
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Delay:", 9
   Else
      Select Case CommandNum
         Case 0
            GetTriggerCmds = "M1"
         Case 1
            GetTriggerCmds = "M2"
         Case 2
            GetTriggerCmds = "D0"
         Case 3
            GetTriggerCmds = "D1"
         Case 4
            GetTriggerCmds = "PER" & NUL$ & " 1/ 5/ 10/ 50/ 100/ 500" & NUL$ & " NS/ US/ MS/ S"
         Case 5
            GetTriggerCmds = "WID" & NUL$ & " 1/ 5/ 10/ 50/ 100/ 500" & NUL$ & " NS/ US/ MS/ S"
         Case 6
            GetTriggerCmds = "DTY" & NUL$ & " 5%/ 10%/ 50%/ 80%"
         Case 7
            GetTriggerCmds = "C1"
         Case 8
            GetTriggerCmds = "C0"
         Case 9
            GetTriggerCmds = "DEL" & NUL$ & " 1/ 5/ 10/ 50/ 100/ 500" & NUL$ & " NS/ US/ MS/ S"
      End Select
   End If

End Function


Function GetAuxTrigCmds(Instance As Integer, CommandNum As Integer) As String
   
   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Normal", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Toggle", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable: Off", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable: On", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Complement On:", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Complement Off:", 5
   Else
      Select Case CommandNum
         Case 0
            GetAuxTrigCmds = "M1" & NUL$ & " 0/ 1/ 2/ 3"
         Case 1
            GetAuxTrigCmds = "M2" & NUL$ & " 0/ 1/ 2/ 3"
         Case 2
            GetAuxTrigCmds = "D0" & NUL$ & " 0/ 1/ 2/ 3"
         Case 3
            GetAuxTrigCmds = "D1" & NUL$ & " 0/ 1/ 2/ 3"
         Case 4
            GetAuxTrigCmds = "C1" & NUL$ & " 0/ 1/ 2/ 3"
         Case 5
            GetAuxTrigCmds = "C0" & NUL$ & " 0/ 1/ 2/ 3"
      End Select
   End If

End Function

Function Get8112Cmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Normal", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Trigger", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Gate", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Ext Width", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mode: Ext Burst", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable: Off", 5
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable: On", 6
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Period:", 7
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Width:", 8
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Duty Cycle:", 9
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "High Level:", 10
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Low Level:", 11
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Complement On:", 12
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Complement Off:", 13
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Recall Saved Settings:", 14
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Store Current Settings:", 15
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Delay:", 16
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Error", 17
   Else
      Select Case CommandNum
         Case 0
            Get8112Cmds = "M1"
         Case 1
            Get8112Cmds = "M2"
         Case 2
            Get8112Cmds = "M3"
         Case 3
            Get8112Cmds = "M4"
         Case 4
            Get8112Cmds = "M5"
         Case 5
            Get8112Cmds = "D0"
         Case 6
            Get8112Cmds = "D1"
         Case 7
            Get8112Cmds = "PER" & NUL$ & " 1/ 5/ 10/ 50/ 100/ 500" & NUL$ & " NS/ US/ MS/ S"
         Case 8
            Get8112Cmds = "WID" & NUL$ & " 1/ 5/ 10/ 50/ 100/ 500" & NUL$ & " NS/ US/ MS/ S"
         Case 9
            Get8112Cmds = "DTY" & NUL$ & " 5%/ 10%/ 50%/ 80%"
         Case 10
            Get8112Cmds = "HIL" & NUL$ & " 1V/ 1.5V/ 2V/ 5V"
         Case 11
            Get8112Cmds = "LOL" & NUL$ & " 1V/ 0V/ -2V/ -5V"
         Case 12
            Get8112Cmds = "C1"
         Case 13
            Get8112Cmds = "C0"
         Case 14
            Get8112Cmds = "RCL" & NUL$ & " 1/ 2/ 3/ 4"
         Case 15
            Get8112Cmds = "STO" & NUL$ & " 1/ 2/ 3/ 4"
         Case 16
            Get8112Cmds = "DEL" & NUL$ & " 1/ 5/ 10/ 50/ 100/ 500" & NUL$ & " NS/ US/ MS/ S"
         Case 17
            Get8112Cmds = "IERR"
      End Select
   End If

End Function

Function GetSwitchCmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Set", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range Set", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Clear", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Mapping", 3
   Else
      Select Case CommandNum
         Case 0
            GetSwitchCmds = "CH" & NUL$ & " 0/ 1/ 2/ 3/ 4/ 5/ 6/ 7/ 8/ 9/ 10/ 11/ 12/ 13/ 14/ 15"
         Case 1
            GetSwitchCmds = "CHS" & NUL$ & " 0-7/ 0-15"
         Case 2
            GetSwitchCmds = "CLEAR"
         Case 3
            GetSwitchCmds = "REMAP 0;8;1;9;2;10;3;11"
      End Select
   End If

End Function

Function GetExtCtlCmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "XControl Off", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Dig Trigger", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "HP8112 TrigSource", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "AI Clock", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "AO Clock", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Alt Signal", 5
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Alt Signal Range", 6
   Else
      Select Case CommandNum
         Case 0
            GetExtCtlCmds = "XCTLOFF"
         Case 1
            GetExtCtlCmds = "XTRIG"
         Case 2
            GetExtCtlCmds = "TRG8112SRC"
         Case 3
            GetExtCtlCmds = "XAICLOCK"
         Case 4
            GetExtCtlCmds = "XAOCLOCK"
         Case 5
            GetExtCtlCmds = "XAIALTSIG 0"
         Case 6
            GetExtCtlCmds = "XAIALTSIGS" & NUL$ & " 0-1/ 0-3"
      End Select
   End If

End Function

Function GetLoopCmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Loop Off", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Trig Loopback", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "AO Loopback", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Digital Loopback", 3
   Else
      Select Case CommandNum
         Case 0
            GetLoopCmds = "LOOPOFF"
         Case 1
            GetLoopCmds = "TRIGLOOP"
         Case 2
            GetLoopCmds = "AOLOOP" & NUL$ & " 0-1/ 0-3"
         Case 3
            GetLoopCmds = "AUXTRIGLOOP" & NUL$ & " 0/ 1/ 2/ 3"
      End Select
   End If

End Function

Function GetBoardCmds(Instance As Integer, CommandNum As Integer) As String

   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Command 1", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Command 2", 1
   Else
      Select Case CommandNum
         Case 0
            GetBoardCmds = "Cmd1"
         Case 1
            GetBoardCmds = "Cmd2"
      End Select
   End If

End Function

Function GetCommands(Instance As Integer, DevName As String, CommandNum As Integer) As String

   If CommandNum < 0 Then frmNewGPIBCtl(Instance%).cmbFuncList.Clear
   Select Case DevName 'UCase(mauCfg(DevIndex).Name)
      Case "8112A", "8112", "HP8112A", "HP8112", "PULSEGEN0", "PULSEGEN1", "PULSEGEN2"
         GetCommands = Get8112Cmds(Instance%, CommandNum)
      Case "FLUKE45", "F45"
         GetCommands = GetFluke45Cmds(Instance%, CommandNum)
      Case "F8840", "8840", "FLUKE8840", "F8840A", "8840A", "FLUKE8840A"
         GetCommands = GetF8840Cmds(Instance%, CommandNum)
      Case "HP34401", "34401", "HP34401A", "34401A"
         GetCommands = GetHP34401Cmds(Instance%, CommandNum)
      Case "DP8200", "8200", "DP8200N", "8200N"
         GetCommands = GetDP8200Cmds(Instance%, CommandNum)
      Case "HP3325", "3325", "HP3325A", "3325A"
         GetCommands = GetHP3325Cmds(Instance%, CommandNum)
      Case "ENCODER"
         GetCommands = GetEncoderCmds(Instance%, CommandNum)
      Case "INDEXER"
         GetCommands = GetIndexerCmds(Instance%, CommandNum)
      Case "TRIGGER"
         GetCommands = GetTriggerCmds(Instance%, CommandNum)
      Case "AUXTRIG0", "AUXTRIG1", "AUXTRIG2", "AUXTRIG3"
         GetCommands = GetAuxTrigCmds(Instance%, CommandNum)
      Case "SWITCH"
         GetCommands = GetSwitchCmds(Instance%, CommandNum)
      Case "XSELECT"
         GetCommands = GetExtCtlCmds(Instance%, CommandNum)
      Case "LOOPBACK"
         GetCommands = GetLoopCmds(Instance%, CommandNum)
      Case Else
         'assume it's a GPIB board
         GetCommands = GetBoardCmds(Instance%, CommandNum)
   End Select

End Function

Function GetDP8200Cmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Go To Local", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Set Voltage:", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Set Hi Voltage:", 2
   Else
      Select Case CommandNum
         Case 0
            GetDP8200Cmds = "L"
         Case 1
            GetDP8200Cmds = "V1" & NUL$ & "+/-" & NUL$ & "1000000/0500000/0200000/0100000"
         Case 2
            GetDP8200Cmds = "V2" & NUL$ & "+/-" & NUL$ & "0500000/0200000/0100000/0050000"
      End Select
   End If

End Function

Function GetF8840Cmds(Instance As Integer, CommandNum As Integer) As String

   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reset", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 200mV", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 2V", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 20V", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 200V", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: AUTO", 5
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Volts AC", 6
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Volts DC", 7
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Ohms", 8
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Identify", 9
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "External Trigger", 10
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Internal Trigger", 11
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Send Trigger", 12
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Self Test", 13
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Error Status", 14
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Enable Front Panel SRQ", 15
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Enable Output Suffix", 16
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable Output Suffix", 17
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reading Rate Fast", 18
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reading Rate Slow", 19
   Else
      Select Case CommandNum
         Case 0
            GetF8840Cmds = "*"
         Case 1
            GetF8840Cmds = "R1"
         Case 2
            GetF8840Cmds = "R2"
         Case 3
            GetF8840Cmds = "R3"
         Case 4
            GetF8840Cmds = "R4"
         Case 5
            GetF8840Cmds = "R0"
         Case 6
            GetF8840Cmds = "F2"
         Case 7
            GetF8840Cmds = "F1"
         Case 8
            GetF8840Cmds = "F3"
         Case 9
            GetF8840Cmds = "G8"
         Case 10
            GetF8840Cmds = "T2"
         Case 11
            GetF8840Cmds = "T0"
         Case 12
            GetF8840Cmds = "?"
         Case 13
            GetF8840Cmds = "Z0"
         Case 14
            GetF8840Cmds = "G7"
         Case 15
            GetF8840Cmds = "N45P1"
         Case 16
            GetF8840Cmds = "Y1"
         Case 17
            GetF8840Cmds = "Y0"
         Case 18
            GetF8840Cmds = "S2"
         Case 19
            GetF8840Cmds = "S0"
      End Select
   End If

End Function

Function GetFluke45Cmds(Instance As Integer, CommandNum As Integer) As String

   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reset", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 300mV", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 3V", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 30V", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: 300V", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Range: AUTO", 5
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Volts AC", 6
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Volts DC", 7
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Ohms", 8
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Identify", 9
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "External Trigger", 10
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Internal Trigger", 11
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Send Trigger", 12
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Self Test", 13
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Error?", 14
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Measurement", 15
   Else
      Select Case CommandNum
         Case 0
            GetFluke45Cmds = "*RST"
         Case 1
            GetFluke45Cmds = "RANGE 1"
         Case 2
            GetFluke45Cmds = "RANGE 2"
         Case 3
            GetFluke45Cmds = "RANGE 3"
         Case 4
            GetFluke45Cmds = "RANGE 4"
         Case 5
            GetFluke45Cmds = "AUTO"
         Case 6
            GetFluke45Cmds = "VAC"
         Case 7
            GetFluke45Cmds = "VDC"
         Case 8
            GetFluke45Cmds = "OHMS"
         Case 9
            GetFluke45Cmds = "*IDN?"
         Case 10
            GetFluke45Cmds = "TRIGGER 2"
         Case 11
            GetFluke45Cmds = "TRIGGER 1"
         Case 12
            GetFluke45Cmds = "*TRG"
         Case 13
            GetFluke45Cmds = "*TST?"
         Case 14
            GetFluke45Cmds = "*ESR?"
         Case 15
            GetFluke45Cmds = "VAL?"
      End Select
   End If

End Function

Function GetHP3325Cmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Select Function:", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Select Frequency:", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Set Amplitude:", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Set Offset:", 3
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure Resistance:", 3
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure Frequency:", 4
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure DC Current:", 5
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure AC Current", 6
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Measurement", 7
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Ohms", 8
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Identify", 9
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "External Trigger", 10
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Internal Trigger", 11
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Send Trigger", 12
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Self Test", 13
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Error Status", 14
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Clear Status", 15
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Enable Output Suffix", 16
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Disable Output Suffix", 17
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reading Rate Fast", 18
      'frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reading Rate Slow", 19
   Else
      Select Case CommandNum
         Case 0
            GetHP3325Cmds = "FU" & NUL$ & "0/1/2/3/4/5"
         Case 1
            GetHP3325Cmds = "FR" & NUL$ & "1/2/5/10/100/" & NUL$ & "HZ/KH/MH"
         Case 2
            GetHP3325Cmds = "AM" & NUL$ & "10/5/2/1/0.1/0.01" & NUL$ & "VO/MV/VR/MR"
         Case 3
            GetHP3325Cmds = "OF" & NUL$ & "01/0.5/0.1/00/-0.1/-01" & NUL$ & "VO"
         'Case 4
         '   GetHP3325Cmds = "CONF:FREQ" & NUL$ & "/0.01/0.1/1/10" & NUL$ & "/, RES MIN/, RES MAX"
         'Case 5
         '   GetHP3325Cmds = "CONF:CURR:DC" & NUL$ & "/ AUTO/ 0.01/ 0.1/ 1/ 10" & NUL$ & "/, RES MIN/, RES MAX"
         'Case 6
         '   GetHP3325Cmds = "CONF:CURR:AC" & NUL$ & "/ AUTO/ 0.01/ 0.1/ 1/ 10" & NUL$ & "/, RES MIN/, RES MAX"
         'Case 7
         '   GetHP3325Cmds = "Read?"
         'Case 8
         '   GetHP3325Cmds = "F3"
         'Case 9
         '   GetHP3325Cmds = "G8"
         'Case 10
         '   GetHP3325Cmds = "T2"
         'Case 11
         '   GetHP3325Cmds = "T0"
         'Case 12
         '   GetHP3325Cmds = "?"
         'Case 13
         '   GetHP3325Cmds = "Z0"
         'Case 14
         '   GetHP3325Cmds = "SYST:ERR?"
         'Case 15
         '   GetHP3325Cmds = "*CLS"
         'Case 16
         '   GetHP3325Cmds = "Y1"
         'Case 17
         '   GetHP3325Cmds = "Y0"
         'Case 18
         '   GetHP3325Cmds = "S2"
         'Case 19
         '   GetHP3325Cmds = "S0"
      End Select
   End If '

End Function

Function GetHP34401Cmds(Instance As Integer, CommandNum As Integer) As String

   NUL$ = Chr$(0)
   If CommandNum < 0 Then
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Reset", 0
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure DC Voltage:", 1
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure AC Voltage:", 2
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure Resistance:", 3
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure Frequency:", 4
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure DC Current:", 5
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Configure AC Current", 6
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Measurement", 7
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Get Error Status", 8
      frmNewGPIBCtl(Instance%).cmbFuncList.AddItem "Clear Status", 9
   Else
      Select Case CommandNum
         Case 0
            GetHP34401Cmds = "*RST"
         Case 1
            GetHP34401Cmds = "CONF:VOLT:DC" & NUL$ & "/ AUTO/ MAX/ MIN/ 0.01/ 0.1/ 1/ 10/:RANGE" & NUL$ & "/ MAX/ MIN/, RES MIN/, RES MAX"
         Case 2
            GetHP34401Cmds = "CONF:VOLT:AC" & NUL$ & "/ AUTO/ 0.01/ 0.1/ 1/ 10" & NUL$ & "/, RES MIN/, RES MAX"
         Case 3
            GetHP34401Cmds = "CONF:RES" & NUL$ & "/ AUTO/ 1/ 1000/ 10000/ 100000" & NUL$ & "/, RES MIN/, RES MAX"
         Case 4
            GetHP34401Cmds = "CONF:FREQ" & NUL$ & "/ AUTO/ 0.01/ 0.1/ 1/ 10" & NUL$ & "/, RES MIN/, RES MAX"
         Case 5
            GetHP34401Cmds = "CONF:CURR:DC" & NUL$ & "/ AUTO/ 0.01/ 0.1/ 1/ 10" & NUL$ & "/, RES MIN/, RES MAX"
         Case 6
            GetHP34401Cmds = "CONF:CURR:AC" & NUL$ & "/ AUTO/ 0.01/ 0.1/ 1/ 10" & NUL$ & "/, RES MIN/, RES MAX"
         Case 7
            GetHP34401Cmds = "Read?"
         Case 8
            GetHP34401Cmds = "SYST:ERR?"
         Case 9
            GetHP34401Cmds = "*CLS"
      End Select
   End If

End Function

Function GetStatString() As String
   
   Stat$ = "Status: 0x" & Hex$(ibsta%) & " ("
   i% = 0
   Do While gasStatBits$(i%) <> ""        ' Print names for status bits
      If ibsta% And ganStatBits%(i%) Then
         Stat$ = Stat$ & gasStatBits$(i%) & " "
      End If
      i% = i% + 1
   Loop
   If Left$(Stat$, Len(Stat$) - 1) = " " Then Stat$ = Left$(Stat$, Len(Stat$) - 1)
   Stat$ = Stat$ & ")"
   Stat$ = Stat$ & " ibcntl& =" & Str$(ibcntl&)
   GetStatString = Stat$

End Function

Function GPIBSetDevice(Instance As Integer, DevName As String) As Integer

   For SearchIndex% = 0 To frmNewGPIBCtl(Instance%).cmbDevice.ListCount - 1
      If Not (InStr(frmNewGPIBCtl(Instance%).cmbDevice.List(SearchIndex%), DevName) = 0) Then
         frmNewGPIBCtl(Instance%).cmbDevice.ListIndex = SearchIndex%
         Complete% = True
         Exit For
      End If
   Next SearchIndex%
   GPIBSetDevice = Complete%

End Function
Function GPIBRead() As String

   frmNewGPIBCtl(0).cmdRead = True
   GPIBRead = frmNewGPIBCtl(0).lblResult.Caption
   
End Function

Sub GPIBWrite(Instance As Integer, CmdString As String)

   frmNewGPIBCtl(Instance%).txtCommand.Text = CmdString
   frmNewGPIBCtl(Instance%).cmdWrite = True

End Sub

Sub InitArrays()

   ganStatBits%(0) = DCAS:  gasStatBits$(0) = "DCAS"
   ganStatBits%(1) = DTAS:  gasStatBits$(1) = "DTAS"
   ganStatBits%(2) = LACS:  gasStatBits$(2) = "LACS"
   ganStatBits%(3) = TACS:  gasStatBits$(3) = "TACS"
   ganStatBits%(4) = AATN:   gasStatBits$(4) = "AATN"
   ganStatBits%(5) = CIC:   gasStatBits$(5) = "CIC"
   ganStatBits%(6) = RREM:   gasStatBits$(6) = "RREM"
   ganStatBits%(7) = LOK:   gasStatBits$(7) = "LOK"
   ganStatBits%(8) = CMPL:  gasStatBits$(8) = "CMPL"
   ganStatBits%(9) = EEVENT:    gasStatBits$(9) = "EEVENT"
   ganStatBits%(10) = SPoll:   gasStatBits$(10) = "SPOLL"
   ganStatBits%(11) = RQS:     gasStatBits$(11) = "RQS"
   ganStatBits%(12) = SRQI: gasStatBits$(12) = "SRQI"
   ganStatBits%(13) = EEND:  gasStatBits$(13) = "EEND"
   ganStatBits%(14) = TIMO: gasStatBits$(14) = "TIMO"
   ganStatBits%(15) = EERR:  gasStatBits$(15) = "EERR"
   ganStatBits%(16) = 0:       gasStatBits$(16) = ""
   
   ' Initialize error code array
   ganErrCodes%(0) = EDVR:   gasErrCode$(0) = "EDVR"
   ganErrCodes%(1) = ECIC:   gasErrCode$(1) = "ECIC"
   ganErrCodes%(2) = ENOL:   gasErrCode$(2) = "ENOL"
   ganErrCodes%(3) = EADR:   gasErrCode$(3) = "EADR"
   ganErrCodes%(4) = EARG:   gasErrCode$(4) = "EARG"
   ganErrCodes%(5) = ESAC:   gasErrCode$(5) = "ESAC"
   ganErrCodes%(6) = EABO:   gasErrCode$(6) = "EABO"
   ganErrCodes%(7) = ENEB:   gasErrCode$(7) = "ENEB"
   ganErrCodes%(8) = EOIP:   gasErrCode$(8) = "EOIP"
   ganErrCodes%(9) = ECAP:   gasErrCode$(9) = "ECAP"
   ganErrCodes%(10) = EFSO:  gasErrCode$(10) = "EFSO"
   ganErrCodes%(11) = EBUS:  gasErrCode$(11) = "EBUS"
   ganErrCodes%(12) = ESTB:  gasErrCode$(12) = "ESTB"
   ganErrCodes%(13) = ESRQ:  gasErrCode$(13) = "ESRQ"
   ganErrCodes%(14) = ETAB:  gasErrCode$(14) = "ETAB"

   ganErrCodes%(15) = EBRK:  gasErrCode$(15) = "EBRK"
   ganErrCodes%(16) = ESLC:  gasErrCode$(16) = "ESLC"
   ganErrCodes%(17) = ETMR:  gasErrCode$(17) = "ETMR"
   ganErrCodes%(18) = ECFG:  gasErrCode$(18) = "ECFG"
   ganErrCodes%(19) = ESML:  gasErrCode$(19) = "ESML"
   ganErrCodes%(20) = EOVR:  gasErrCode$(20) = "EOVR"
   ganErrCodes%(21) = EVDD:  gasErrCode$(21) = "EVDD"
   ganErrCodes%(22) = EWMD:  gasErrCode$(22) = "EWMD"
   ganErrCodes%(23) = EINT:  gasErrCode$(23) = "EINT"

End Sub

Function InitDevice(DevIndex As Integer) As Integer

   BdDevName$ = mauCfg(DevIndex).Name
   If mauCfg(DevIndex).Handle = -1 Then
      DoIBFIND BdDevName$, ud%
      If Not (ud% > -1) Then Exit Function
      mauCfg(DevIndex).Handle = ud%
   Else
      ud% = mauCfg(DevIndex).Handle
   End If
   
   If gnDeviceLevel Then gnLevel = ud%
   gbGPIBActive% = True
   ibask ud%, IbcPAD, value%
   gbGPIBActive% = False
   Dev$ = mauCfg(Device).Name
   PrintStatus GPIBAsk, Dev$, "IBAsk (IbcPAD); ibask(" & "0x" & Hex$(ud%) & ", 0x" & Hex$(IbcPAD) & ", 0x" & Hex$(value%) & ")", 0
   mauCfg(DevIndex).address = value%
   'ibask UD%, IbcREADDR, Value%
   gbGPIBActive% = False
   'PrintStatus "IBAsk (IbcREADDR); ibask(" & "0x" & Hex$(UD%) & ", 0x" & Hex$(IbcREADDR) & ", " & Value% & ")", 0
   'mauCfg(DevIndex).ForcedAddressing = Value%
   InitDevice = True

End Function

Function InitGPIB() As Integer

   If mnInitGPIB Then
      InitGPIB = mnInitGPIB
      Exit Function
   End If

   On Error GoTo MissingGPIB

   lpFileName$ = "Cfg488crh.ini"

   SendIFC (0)   'make sure there's a library installed
   If Not (ibsta < 0) Then
      'library exists, but no board installed or configured
      'see if MCC signal controllers are installed
      mnInitGPIB = True
   Else
      InitGPIB = False
      Exit Function
   End If

   InitArrays  'arrays of GPIB status and error strings and codes

   ReDim mauCfg(34)
   boardindex% = 34
   mauCfg(boardindex%).Handle = -1
   mauCfg(boardindex%).address = -1
   mauCfg(boardindex%).Name = "NOADDR"
      
   boardindex% = 0
   BdDevName$ = Space$(32)
   StrLen& = Len(BdDevName$)
   GetGPIBBoardName boardindex%, BdDevName$, UsingIni%

   If UsingIni% Then
      'the ini file is a workaround to using board names
      'the NI library no longer uses board names, except for GPIBx
      'so to be compatible with this program, the name is picked up
      'from the ini file by the GetGPIBBoardName subroutine
      For CheckAddress% = 0 To 33
         mauCfg(CheckAddress%).Handle = -1
         mauCfg(CheckAddress%).address = -1
         GetGPIBBoardName CheckAddress%, BdDevName$, UsingIni%
         If Len(BdDevName$) Then
            mauCfg(boardindex%).Name = BdDevName$
            'create descriptor if GPIB board (not device)
            If InitDevice(boardindex%) Then
               'this calls ibfind - it will only work for a board
               'since the devices are not named in the NI library
               mnInitGPIB = True
               InitGPIB = mnInitGPIB
               If gnBoardIndex = -1 Then gnBoardIndex = boardindex%
               ud% = mauCfg(boardindex%).Handle
               gnBoardsInstalled = gnBoardsInstalled + 1
               ibsic ud%
            Else
               'must be a device - in this case the address is CheckAddress%
               mauCfg(boardindex%).address = CheckAddress%
               mauCfg(boardindex%).Name = BdDevName$
            End If
            boardindex% = boardindex% + 1
         End If
      Next CheckAddress%
   Else
      For boardindex% = 0 To 1
         mauCfg(boardindex%).Handle = -1
         mauCfg(boardindex%).address = -1
         BdDevName$ = Space$(32)
         StrLen& = Len(BdDevName$)
      
         GetGPIBBoardName boardindex%, BdDevName$, UsingIni%
         BdDevName$ = Trim$(BdDevName$)
         i% = 0
         If Len(BdDevName$) Then
            Do
               i% = i% + 1
               CurChar% = Asc(Mid$(BdDevName$, i%, 1))
            Loop While (CurChar% <> 0) And (i% < Len(BdDevName$))
            If CurChar% = 0 Then BdDevName$ = Left$(BdDevName$, i% - 1)
            mauCfg(boardindex%).Name = BdDevName$
            'frmMain.optSelectBoard(boardindex%).Caption = mauCfg(boardindex%).Name
            If Len(BdDevName$) Then
               'create descriptor if GPIB board (not device)
               If InitDevice(boardindex%) Then
                  mnInitGPIB = True
                  InitGPIB = mnInitGPIB
                  If gnBoardIndex = -1 Then gnBoardIndex = boardindex% - 1
                  ud% = mauCfg(boardindex%).Handle
                  gnBoardsInstalled = gnBoardsInstalled + 1
                  If IbaBoardType% = 0 Then
                     'using NI header
                     NIHeader% = True
                     DoIBASK ud%, 0, "IbaBaseAddr", IbaBaseAddr%, value%
                     'NI returns 0 for base address with error
                     If value% = 0 Then
                        'its likely this board
                        BoardType$ = "AT-GPIB/TNT (Plug and Play)"
                        value% = -1
                     Else
                        'using the CBI library but NI header
                        DoIBASK ud%, 0, ItemName$, &H300, value%
                     End If
                  Else
                     DoIBASK ud%, 0, ItemName$, IbaBoardType%, value%
                  End If
                  Select Case value%
                     Case 0
                        BoardType$ = "Unknown board type"
                     Case 1
                        BoardType$ = "Unknown board type"
                     Case 2
                        DoIBASK ud%, 0, ItemName$, IbaChipType%, value%
                        If value% = 0 Then BoardType$ = "CIO-PC2A"
                        If value% = 1 Then BoardType$ = "ISA-GPIB-PC2A"
                     Case 3
                        BoardType$ = "ISA-GPIB/LC"
                     Case 4
                        BoardType$ = "ISA-GPIB"
                     Case 5
                        BoardType$ = "PCM-GPIB"
                     Case 6
                        BoardType$ = "PCI-GPIB"
                     Case 7
                        BoardType$ = "PC104-GPIB"
                     Case 14
                        BoardType$ = "CPCI-GPIB"
                  End Select
                  If NIHeader% Then HeadString$ = " (Using NI header)"
                  frmMain.Caption = frmMain.Caption & " " & BoardType$ & HeadString$
                  'x% = InitDevice(BoardIndex%)
                  'mauCfg(boardindex%).NameOfType = BoardType$
               End If
               'frmNewGPIBCtl(Instance%).cmbDevice.AddItem mauCfg(BoardIndex%).Name, BoardIndex%
            End If
         End If
      Next boardindex%
      'If Not (gnBoardsInstalled = 0) Then frmMain.optSelectBoard(gnBoardIndex).Value = True

      'Get the rest of the device names in the cfg file
      'If gnBoardsInstalled = 2 Then frmMain.optSelectBoard(1).Visible = True
      For DeviceIndex% = 2 To 33
         mauCfg(DeviceIndex%).Handle = -1
         mauCfg(DeviceIndex%).address = -1
         BdDevName$ = Space$(32)
         StrLen& = Len(BdDevName$)
         If Not (UCase$(Command$) = "/C") Then GetGPIBBoardName DeviceIndex%, BdDevName$, UsingIni%
         BdDevName$ = Trim$(BdDevName$)
         i% = 0
         If Len(BdDevName$) Then
            Do
               i% = i% + 1
               CurChar% = Asc(Mid$(BdDevName$, i%, 1))
            Loop While (CurChar% <> 0) And (i% < Len(BdDevName$))
            If CurChar% = 0 Then BdDevName$ = Left$(BdDevName$, i% - 1)
            mauCfg(DeviceIndex%).Name = BdDevName$
            'frmNewGPIBCtl(Instance%).cmbDevice.AddItem mauCfg(DeviceIndex%).Name, DeviceIndex%
            x% = InitDevice(DeviceIndex%)
        End If
      Next DeviceIndex%
      Screen.MousePointer = vbDefault
   End If
   Exit Function

MissingGPIB:
   Device% = -1
   BdDevName$ = "No Library"
   Exit Function

End Function

Sub LinkCommand(Instance As Integer, CmdString As String)

   If CmdString = "Read" Then
      frmNewGPIBCtl(Instance%).cmdRead = True
   Else
      frmNewGPIBCtl(Instance%).txtCommand.Text = CmdString
   End If

End Sub

Function LinkRetrieve(Instance As Integer) As Single

   LinkRetrieve = frmNewGPIBCtl(Instance%).lblResult

End Function

Sub LinkStart(Instance As Integer)

   frmNewGPIBCtl(Instance%).cmdWrite = True

End Sub

Public Function OpenGPIB() As Integer
   
   ReDim Preserve frmNewGPIBCtl(gnGPIBCtlForms)
   Set frmNewGPIBCtl(gnGPIBCtlForms) = New frmGPIBCtl
   frmNewGPIBCtl(gnGPIBCtlForms).Show
   If Not gnErrFlag Then
      frmNewGPIBCtl(gnGPIBCtlForms).Left = mfmUniTest.ScaleWidth - frmNewGPIBCtl(gnGPIBCtlForms).Width
      frmNewGPIBCtl(gnGPIBCtlForms).Top = frmNewGPIBCtl(gnGPIBCtlForms).Height * gnGPIBCtlForms
      frmNewGPIBCtl(gnGPIBCtlForms).Tag = Hex$(GPIB_CTL * &H100 + gnGPIBCtlForms)
      gnGPIBCtlForms = gnGPIBCtlForms + 1
      If mnInitGPIB Then FillDevList gnGPIBCtlForms - 1
      OpenGPIB = True
   End If

End Function

Sub PrintStatus(FuncID As Integer, Device$, ErrStr$, CheckError%)
   
   ' Name:        PrintStatus
   ' Arguments:   ---
   
   ' Description: Prints the global GPIB status and error codes

   Stat$ = GetStatString$()
   If 55 - Len(ErrStr$) > 0 Then
      Fill$ = Space$(70 - Len(ErrStr$))
   Else
      Fill$ = CrLf$ & Space$(55)
   End If
   
   If ibsta% And EERR Then
      i% = 0
      e$ = ""
      Do While gasErrCode$(i%) <> ""
         If iberr% = ganErrCodes%(i%) Then
            e$ = e$ & gasErrCode$(i%)
            Exit Do
         End If
         i% = i% + 1
      Loop
      If CheckError% Then
         gnGlobalFlag = True
         CrLf$ = Chr$(13) & Chr$(10)
         MsgBox ErrStr$ & CrLf$ & CrLf$ & Stat$ & CrLf$ & CrLf$ & "Error:   " & e$, 48, "GPIB Error"
      End If
   End If

   DoEvents

   StatNum% = StatNum% + 1
   gnClearStat% = 0
   If gnScriptSave Then
      A2 = Device$
      A3 = frmNewGPIBCtl(0).txtCommand.Text
      Select Case FuncID
         Case GPSend To GPDevClear, GPSelDevClear, GPIBSre
            Print #2, "8, " & Format$(FuncID, "0") & ", " & Format$(ibsta%, "0") & ", " & Format$(A1, "") & ", " & _
            Format$(A2, "") & ", "; A3; ","; A4; ","; A5; ","; A6; ","; A7; ","; A8; ","; A9; ","; A10; ","; A11; ","; AuxHandle
            If Not InStr(Device$, "8200") = 0 Then
               'add a delay for 8200 to settle
               Print #2, "0, 3000, 0, 1,,,,,,,,,,,"
            End If
      End Select
      'not logged:  GPFind, GPIBAsk, GPInit, GPPtrs
   End If

End Sub

Sub ReadGPIB(Device As Integer, GPIBReturn As String)

   address% = mauCfg(Device).address
   Receive 0, address%, GPIBReturn, STOPend
   GPIBReturn = Trim$(GPIBReturn)
   Dev$ = mauCfg(Device).Name
   PrintStatus GPReceive, Dev$, "Receive (0, " & Format$(address%, "0") & ", " & GPIBReturn & ", 0x" & Hex$(STOPend) & ")", True

End Sub

Sub TriggerGPIB(Device As Integer)

   address% = mauCfg(Device).address
   Trigger 0, address%
   Dev$ = mauCfg(Device).Name
   PrintStatus GPTrigger, Dev$, "Trigger (0, " & Format$(address%, "0") & ")", True

End Sub

Sub CloseGPIB(Instance As Integer)

   On Error GoTo CloseError
   Unload frmNewGPIBCtl(Instance%)
   Exit Sub

CloseError:
   If Err = 9 Then
      'Instance doesn't exist (subscript out of range)
      Exit Sub
   Else
      MsgBox Error$(Err), 0, "Error Closing GPIB Form"
   End If
   Exit Sub

End Sub

Sub DoDevClear(Device As Integer)

   If Device < 0 Then
      address% = NOADDR 'send DCL rather than SDC
      Dev$ = "All devices"
   Else
      address% = mauCfg(Device).address
      Dev$ = mauCfg(Device).Name
   End If
   DevClear 0, address%
   
   PrintStatus GPDevClear, Dev$, "DevClear (0, " & Format$(address%, "0") & ")", True

End Sub

Sub WaitTillDone()

   DoEvents

End Sub

Sub WriteGPIB(Device As Integer, GPIBCommand As String)

   address% = mauCfg(Device).address
   Send 0, address%, GPIBCommand, DABend
   Dev$ = mauCfg(Device).Name
   PrintStatus GPSend, Dev$, "Send (0, " & Format$(Device, "0") & _
      ", " & GPIBCommand & ", 0x" & Hex$(DABend), 0

End Sub

Sub DestroyGPIBBoards()

   mnInitGPIB = False

End Sub
