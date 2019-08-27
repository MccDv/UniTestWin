Attribute VB_Name = "LIBINFO"
'Description
'  obtains information from the Universal Library
'  using mainly cbGetConfig()

'Dependencies
'     cbw.bas
'     errorhand.frm

Global Const WIN3X95 = 0
Global Const WIN95NT = 1
Global Const DOTNET = 2

Global geErrFlow As Integer, gnLocalErrDisp As Integer
Global geWarnFlow As Integer, gnLocalWarnDisp As Integer
Global gnErrReporting As Integer, gnErrHandling As Integer
Global gnHexVals As Integer, gnLiveDisp As Integer

Global gnLibVer As Integer

Dim mfNoForm As Form

Function BadULFunction(FuncName As String) As Integer

   BadULFunction = NOTWINDOWSFUNC

End Function

Sub ErrorHandling(Optional NotAnError As Integer = 0)

   HistString$ = GetHistoryText()
   If Not NotAnError Then DateString$ = "----  " & Now & "  ----" & vbCrLf
   frmErrors.txtErrorMessage.Text = DateString$ & HistString$
   If NotAnError Then frmErrors.SetNoError
   ModalForm% = 1
   If gnLiveDisp Then ModalForm% = 0
   If Not frmErrors.Visible Then frmErrors.Show (ModalForm%)
   
End Sub

Function GetNameOfBoard(ByVal BoardNum As Integer, Optional LibType As Integer = UNILIB) As String

   'redirects based on optional parameter
   Select Case LibType
      Case UNILIB
         GetNameOfBoard = ""
         BoardName$ = Space$(BOARDNAMELEN)
         ULStat = cbGetBoardName(BoardNum, BoardName$)
         If Len(BoardName$) Then
            'Drop the space characters
            BoardName$ = RTrim$(BoardName$)
            StringSize% = Len(BoardName$)
            'lop off the null
            If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
            GetNameOfBoard = BoardName$
         Else
            GetNameOfBoard = ""
         End If
         If SaveFunc(mfNoForm, GetBoardName, ULStat, BoardNum, BoardName$, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
      Case NETLIB
         BoardName$ = gnBoardEnum(BoardNum)
      Case MSGLIB
         BoardName$ = GetNameOfMsgBoard(BoardNum)
         GetNameOfBoard = BoardName$
         Exit Function
   End Select

End Function

Function GetLibDLLRev(ByRef Filename As String) As String

   ReDim Props(1 To 7) As String
   
   CurrentDir$ = CurDir()
   InfoFileExists$ = Dir(CurrentDir$ & "\VB5STKIT.DLL")
   If InfoFileExists$ = "" Then
      NewPath$ = App.Path & "\VB5STKIT.DLL"
      InfoFileExists$ = Dir(NewPath$)
      If InfoFileExists$ = "" Then
         GetLibDLLRev = "NoInfoFile"
         Exit Function
      Else
         ChDir (App.Path)
      End If
   End If
   
   If LCase$(Right$(Filename$, 4)) = ".dll" Then _
      FullPath$ = GetSysInfo(DLLINF, Filename$)
   GetFileVersion Filename$, Props$()
   GetLibDLLRev = Props(6)
   
End Function

Function GetNumInstalled() As Integer
   
   'Get number of installed boards
   
   InfoType% = GLOBALINFO
   DevNum% = 0
   ConfigItem% = GINUMBOARDS
   'to do - this beeps on startup sometimes (USB-QUAD08 installed)
   ULStat = GetConfig520(InfoType%, 0, DevNum%, ConfigItem%, ValConfig&)
   If ULStat = CFGFILENOTFOUND Then
      GetNumInstalled = 0
      Exit Function
   End If
   ConfigVal& = ValConfig&
   'ULStat = cbGetConfig(InfoType%, 0, DevNum%, ConfigItem%, ConfigVal&)
   If SaveFunc(mfNoForm, GetConfig, ULStat, InfoType%, 0, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   
   NumBoards% = ConfigVal&          'this number is actually MAX_NUM_BOARDS (should be GI_NUMBOARDS)
                                    'the following For/Next compensates for this
   If NumBoards% > 50 And (Not WIN32APP) Then NumBoards% = 10
                                    'this previous line is required because of a bug in the 16 bit
                                    'library.. if cfg file was created by Inscal32 then the number
                                    'returned above will be 200.. 16 bit lib crashes if following
                                    'For loop is run with board% = 67 (bug #1021)
   For board% = 0 To NumBoards% - 1
      InfoType% = BOARDINFO
      DevNum% = 0
      ConfigItem% = BIBOARDTYPE
      ULStat = GetConfig520(InfoType%, board%, DevNum%, ConfigItem%, ValConfig&)
      ConfigVal& = ValConfig&
      If SaveFunc(mfNoForm, GetConfig, ULStat, InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
      If ConfigVal& <> 0 Then
         'determine boards relative position (second board installed could be board 5, etc.)
         ReDim Preserve gnBoardEnum(Installed%)
         gnBoardEnum(Installed%) = board%
         Installed% = Installed% + 1
         
         'check if this board has an associated memory board
         InfoType% = BOARDINFO
         DevNum% = 0
         ConfigItem% = BIDTBOARD
         ULStat = GetConfig520(InfoType%, board%, DevNum%, ConfigItem%, ValConfig&)
         ConfigVal& = ValConfig&
         'ULStat = cbGetConfig(InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&)
         If SaveFunc(mfNoForm, GetConfig, ULStat, InfoType%, board%, DevNum%, ConfigItem%, ConfigVal&, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
         If ConfigVal& > -1 Then
            'determine boards relative position (second board installed could be board 5, etc.)
            'and if it has already been detected as its own board number (UL16)
            ReDim Preserve gnBoardEnum(Installed%)
            gnBoardEnum(Installed%) = ConfigVal&
            Installed% = Installed% + 1
         End If
      End If
   Next board%

   GetNumInstalled = Installed%

End Function

Function GetValidRanges(BoardNum As Integer, Ranges() As Integer) As Integer
      
   ULStat = GetConfig520(BOARDINFO, BoardNum, 0, BIRANGE, ValConfig&)
   ConfigVal% = ValConfig&
   'ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, BIRANGE, ConfigVal%)
   If SaveFunc(mfNoForm, GetConfig, ULStat, BOARDINFO, BoardNum, 0, BIRANGE, ConfigVal%, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   If C = -1 Then
      GetValidRanges = 0
      Exit Function
   End If

   For i% = -1 To 12
      If cbAIn(BoardNum, 0, i%, 0) <> BADRANGE Then
         Ranges(ValidRanges%) = i%
         ValidRanges% = ValidRanges% + 1
      End If
   Next i%
   For i% = 100 To 110
      If cbAIn(BoardNum, 0, i%, 0) <> BADRANGE Then
         Ranges(ValidRanges%) = i%
         ValidRanges% = ValidRanges% + 1
      End If
   Next i%
   For i% = 200 To 203
      If cbAIn(BoardNum, 0, i%, 0) <> BADRANGE Then
         Ranges(ValidRanges%) = i%
         ValidRanges% = ValidRanges% + 1
      End If
   Next i%
   ReDim Preserve Ranges(ValidRanges%)

End Function

Function TestRange(ByVal FormType As Integer, ByVal BoardNum As Integer, ByVal Range As Integer) As Integer

   Dim NoForm As Form
   ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
   If SaveFunc(NoForm, ErrHandling, ULStat, DONTPRINT, DONTSTOP, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   If FormType = ANALOG_OUT Then
      If BoardNum < 0 Then
         'DeviceNum& = Abs(mnBoardNum + 1)
         TestRange = True
      Else
         TestRange = (cbAOut(BoardNum, 0, Range, 0) <> BADRANGE)
      End If
   Else
      TestRange = (cbAIn(BoardNum, 0, Range, 0) <> BADRANGE)
   End If
   ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
   If SaveFunc(NoForm, ErrHandling, ULStat, gnErrReporting, gnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Function

End Function

Public Function GetCallAsString(ULCall As Integer, Optional A0, Optional A1, _
Optional A2, Optional A3, Optional A4, Optional A5, Optional A6, Optional A7, _
Optional A8, Optional A9, Optional A10) As String

   CallString$ = Format(ULCall, "0") & ","
   NumberOfArgs% = GetNumArgs(ULCall)
   For Arg% = 1 To NumberOfArgs%
      ArgVal = Choose(Arg%, A0, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10)
      If Not (IsNull(ArgVal) Or IsMissing(ArgVal)) Then
         CallString$ = CallString$ & ArgVal & ","
      End If
   Next
   CallString$ = Left(CallString$, Len(CallString$) - 1)
   GetCallAsString = CallString$
   
End Function


