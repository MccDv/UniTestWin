Attribute VB_Name = "LOCINFO"
'Description
'  gets board specific information not available from
'  the Universal Library itself

'Dependencies
'  cbw.bas (for constants used in some functions)

'Global Const BIP30VOLTS = 23

Dim msBoardName As String

Function GetADResolution(BoardName As String, BoardNum As Integer) As Integer

   'determine boards input resolution
   'the resolution may be listed in the configuration file (required for res > 16 bit)
   'if the file exists, check if resolution listed for this board
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = BoardName$
   lpKeyName$ = "ADResolution"
   lpDefault$ = "0"
   nSize% = 16
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   If Len(CustomParam$) Then Resolution& = Val(CustomParam$)

   'if Resolution& is zero, derive resolution from cbToEngUnits() if using UL or statically otherwise
   'for some devices (>16 bit), EngUnits now returns error, so err handling needs squelching
   If Not (gnLibType = MSGLIB) Then
      ErrRep% = gnErrReporting
      ErrHand% = gnErrHandling
      LocRep% = gnLocalErrDisp
      gnLocalErrDisp = False
      ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
      InfoType% = BOARDINFO
      ConfigItem% = BIADRES
      DevNum% = 0
      ULStat = GetConfig520(InfoType%, BoardNum, DevNum%, ConfigItem%, ValConfig&)
      If ULStat = 0 Then Resolution& = ValConfig&
      If Resolution& = 0 Then
         Dim NoForm As Form
         Sample% = 32767
         'BoardNum% = Val(BoardName)
         For RangeList% = 1 To 4
            Range% = Choose(RangeList%, BIP5VOLTS, BIP10VOLTS, UNI10VOLTS, UNI5VOLTS)
            TestV% = 4
            ULStat = cbToEngUnits(BoardNum, Range%, Sample%, EngUnits!)
            If Not (ULStat = BADRANGE) Then
               x% = SaveFunc(NoForm, ToEngUnits, ULStat, BoardNum, Range%, Sample%, EngUnits!, A5, A6, A7, A8, A9, A10, A11, 0)
               If Abs(EngUnits!) > TestV% Then
                  'check if 11 bit
                  Sample% = 2048
                  ULStat = cbToEngUnits(BoardNum, Range%, Sample%, EngUnits!)
                  If Abs(EngUnits!) > TestV% Then
                     Resolution& = 11
                     If EngUnits! < 0 Then Resolution& = -12
                  Else
                     If EngUnits! = -3.75 Then
                        Resolution& = 14
                     ElseIf EngUnits! = -2.5 Then
                        Resolution& = 13
                     Else
                        Resolution& = 12
                        If EngUnits! < 0 Then Resolution& = -12
                     End If
                  End If
               Else
                  Resolution& = 16
               End If
               Exit For
            End If
         Next RangeList%
      Else
         Resolution& = Resolution&
      End If
      
      gnLocalErrDisp = LocRep%
      ULStat = cbErrHandling(ErrRep%, ErrHand%)
   End If
   GetADResolution = Resolution&
   
End Function

Function GetBoardInfoString(BIType As Integer) As String

   Select Case BIType
      Case BIBASEADR
         GetBoardInfoString = "BIBASEADR"
      Case BIBOARDTYPE
         GetBoardInfoString = "BIBOARDTYPE"
      Case BIINTLEVEL
         GetBoardInfoString = "BIINTLEVEL"
      Case BIDMACHAN
         GetBoardInfoString = "BIDMACHAN"
      Case BIINITIALIZED
         GetBoardInfoString = "BIINITIALIZED"
      Case BICLOCK
         GetBoardInfoString = "BICLOCK"
      Case BIRANGE
         GetBoardInfoString = "BIRANGE"
      Case BINUMADCHANS
         GetBoardInfoString = "BINUMADCHANS"
      Case BIUSESEXPS
         GetBoardInfoString = "BIUSESEXPS"
      Case BIDINUMDEVS
         GetBoardInfoString = "BIDINUMDEVS"
      Case BIDIDEVNUM
         GetBoardInfoString = "BIDIDEVNUM"
      Case BICINUMDEVS
         GetBoardInfoString = "BICINUMDEVS"
      Case BICIDEVNUM
         GetBoardInfoString = "BICIDEVNUM"
      Case BINUMDACHANS
         GetBoardInfoString = "BINUMDACHANS"
      Case BIWAITSTATE
         GetBoardInfoString = "BIWAITSTATE"
      Case BINUMIOPORTS
         GetBoardInfoString = "BINUMIOPORTS"
      Case BIPARENTBOARD
         GetBoardInfoString = "BIPARENTBOARD"
      Case BIDTBOARD
         GetBoardInfoString = "BIDTBOARD"
   End Select

End Function

Function GetCounterInfoString(CIType As Integer) As String

   Select Case CIType
      Case CIBASEADR
         GetCounterInfoString = "CIBASEADR"
      Case CIINITIALIZED
         GetCounterInfoString = "CIINITIALIZED"
      Case CICTRTYPE
         GetCounterInfoString = "CICTRTYPE"
      Case CICTRNUM
         GetCounterInfoString = "CICTRNUM"
      Case CICONFIGBYTE
         GetCounterInfoString = "CICONFIGBYTE"
   End Select

End Function

Function GetDAResolution(BoardName As String, BoardNum As Integer, Range As Integer) As Integer

   'determine boards input resolution
   'the resolution may be listed in the configuration file (required for res > 16 bit)
   'if the file exists, check if resolution listed for this board
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = BoardName$
   lpKeyName$ = "DAResolution"
   lpDefault$ = "0"
   nSize% = 16
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   If Len(CustomParam$) Then Resolution& = Val(CustomParam$)

   'if Resolution& is zero, derive resolution from cbFromEngUnits()
   'for some devices (>16 bit), EngUnits now returns error, so err handling needs squelching
   If Not (gnLibType = MSGLIB) Then
      ErrRep% = gnErrReporting
      ErrHand% = gnErrHandling
      LocRep% = gnLocalErrDisp
      gnLocalErrDisp = False
      ULStat = cbErrHandling(DONTPRINT, DONTSTOP)
      If Resolution& = 0 Then
         'derive resolution from cbFromEngUnits()
         Dim NoForm As Form
         EngUnits! = 10
         ULStat = cbFromEngUnits(BoardNum, Range, EngUnits!, Sample%)
         If Not ((ULStat = BADRANGE) Or ((ULStat = CBUSB_BAD_RANGE) And (CBUSB_BAD_RANGE <> 0))) Then
            x% = SaveFunc(NoForm, FromEngUnits, ULStat, BoardNum, Range%, EngUnits!, Sample%, A5, A6, A7, A8, A9, A10, A11, 0)
            Select Case Sample%
               Case Is < 0
                  GetDAResolution = 16
               Case Is < 300
                  GetDAResolution = 8
               Case Is < 1100
                  GetDAResolution = 10
               Case 0
                  GetDAResolution = -12
               Case Else
                  GetDAResolution = 12
            End Select
         End If
      Else
         GetDAResolution = Resolution&
      End If
      
      gnLocalErrDisp = LocRep%
      ULStat = cbErrHandling(ErrRep%, ErrHand%)
   Else
      GetDAResolution = Resolution&
   End If

End Function

Function GetDigitalInfoString(DIType As Integer) As String

   Select Case DIType
      Case DIBASEADR
         GetDigitalInfoString = "DIBASEADR"
      Case DIINITIALIZED
         GetDigitalInfoString = "DIINITIALIZED"
      Case DIDEVTYPE
         GetDigitalInfoString = "DIDEVTYPE"
      Case DIMASK
         GetDigitalInfoString = "DIMASK"
      Case DIREADWRITE
         GetDigitalInfoString = "DIREADWRITE"
      Case DICONFIG
         GetDigitalInfoString = "DICONFIG"
      Case DINUMBITS
         GetDigitalInfoString = "DINUMBITS"
      Case DICURVAL
         GetDigitalInfoString = "DICURVAL"
   End Select

End Function

Function GetErrorConst(ErrNum As Variant) As String

   Select Case ErrNum
      Case 0
         ConstString$ = "NOERRORS"
         If NOERRORS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1
         ConstString$ = "BADBOARD"
         If BADBOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 2
         ConstString$ = "DEADDIGITALDEV"
         If DEADDIGITALDEV <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 3
         ConstString$ = "DEADCOUNTERDEV"
         If DEADCOUNTERDEV <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 4
         ConstString$ = "DEADDADEV"
         If DEADDADEV <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 5
         ConstString$ = "DEADADDEV"
         If DEADADDEV <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 6
         ConstString$ = "NOTDIGITALCONF"
         If NOTDIGITALCONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 7
         ConstString$ = "NOTCOUNTERCONF"
         If NOTCOUNTERCONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 8
         ConstString$ = "NOTDACONF"
         If NOTDACONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 9
         ConstString$ = "NOTADCONF"
         If NOTADCONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 10
         ConstString$ = "NOTMUXCONF"
         If NOTMUXCONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 11
         ConstString$ = "BADPORTNUM"
         If BADPORTNUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 12
         ConstString$ = "BADCOUNTERDEVNUM"
         If BADCOUNTERDEVNUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 13
         ConstString$ = "BADDADEVNUM"
         If BADDADEVNUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 14
         ConstString$ = "BADSAMPLEMODE"
         If BADSAMPLEMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 15
         ConstString$ = "BADINT"
         If BADINT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 16
         ConstString$ = "BADADCHAN"
         If BADADCHAN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 17
         ConstString$ = "BADCOUNT"
         If BADCOUNT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 18
         ConstString$ = "BADCNTRCONFIG"
         If BADCNTRCONFIG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 19
         ConstString$ = "BADDAVAL"
         If BADDAVAL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 20
         ConstString$ = "BADDACHAN"
         If BADDACHAN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 22
         ConstString$ = "ALREADYACTIVE"
         If ALREADYACTIVE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 23
         ConstString$ = "PAGEOVERRUN"
         If PAGEOVERRUN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 24
         ConstString$ = "BADRATE"
         If BADRATE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 25
         ConstString$ = "COMPATMODE"
         If COMPATMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 26
         ConstString$ = "TRIGSTATE"
         If TRIGSTATE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 27
         ConstString$ = "ADSTATUSHUNG"
         If ADSTATUSHUNG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 28
         ConstString$ = "TOOFEW"
         If TOOFEW <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 29
         ConstString$ = "OVERRUN"
         If OVERRUN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 30
         ConstString$ = "BADRANGE"
         If BADRANGE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 31
         ConstString$ = "NOPROGGAIN"
         If NOPROGGAIN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 32
         ConstString$ = "BADFILENAME"
         If BADFILENAME <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 33
         ConstString$ = "DISKISFULL"
         If DISKISFULL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 34
         ConstString$ = "COMPATWARN"
         If COMPATWARN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 35
         ConstString$ = "BADPOINTER"
         If BADPOINTER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 36
         ConstString$ = "TOOMANYGAINS"
         If TOOMANYGAINS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 37
         ConstString$ = "RATEWARNING"
         If RATEWARNING <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 38
         ConstString$ = "CONVERTDMA"
         If CONVERTDMA <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 39
         ConstString$ = "DTCONNECTERR"
         If DTCONNECTERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 40
         ConstString$ = "FORECONTINUOUS"
         If FORECONTINUOUS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 41
         ConstString$ = "BADBOARDTYPE"
         If BADBOARDTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 42
         ConstString$ = "WRONGDIGCONFIG"
         If WRONGDIGCONFIG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 43
         ConstString$ = "NOTCONFIGURABLE"
         If NOTCONFIGURABLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 44
         ConstString$ = "BADPORTCONFIG"
         If BADPORTCONFIG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 45
         ConstString$ = "BADFIRSTPOINT"
         If BADFIRSTPOINT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 46
         ConstString$ = "ENDOFFILE"
         If ENDOFFILE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 47
         ConstString$ = "NOT8254CTR"
         If NOT8254CTR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 48
         ConstString$ = "NOT9513CTR"
         If NOT9513CTR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 49
         ConstString$ = "BADTRIGTYPE"
         If BADTRIGTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 50
         ConstString$ = "BADTRIGVALUE"
         If BADTRIGVALUE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 52
         ConstString$ = "BADOPTION"
         If BADOPTION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 53
         ConstString$ = "BADPRETRIGCOUNT"
         If BADPRETRIGCOUNT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 55
         ConstString$ = "BADDIVIDER"
         If BADDIVIDER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 56
         ConstString$ = "BADSOURCE"
         If BADSOURCE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 57
         ConstString$ = "BADCOMPARE"
         If BADCOMPARE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 58
         ConstString$ = "BADTIMEOFDAY"
         If BADTIMEOFDAY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 59
         ConstString$ = "BADGATEINTERVAL"
         If BADGATEINTERVAL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 60
         ConstString$ = "BADGATECNTRL"
         If BADGATECNTRL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 61
         ConstString$ = "BADCOUNTEREDGE"
         If BADCOUNTEREDGE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 62
         ConstString$ = "BADSPCLGATE"
         If BADSPCLGATE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 63
         ConstString$ = "BADRELOAD"
         If BADRELOAD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 64
         ConstString$ = "BADRECYCLEFLAG"
         If BADRECYCLEFLAG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 65
         ConstString$ = "BADBCDFLAG"
         If BADBCDFLAG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 66
         ConstString$ = "BADDIRECTION"
         If BADDIRECTION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 67
         ConstString$ = "BADOUTCONTROL"
         If BADOUTCONTROL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 68
         ConstString$ = "BADBITNUMBER"
         If BADBITNUMBER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 69
         ConstString$ = "NONEENABLED"
         If NONEENABLED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 70
         ConstString$ = "BADCTRCONTROL"
         If BADCTRCONTROL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 71
         ConstString$ = "BADMUXCHAN"
         If BADMUXCHAN <> ErrNum Then
            If BADEXPCHAN <> ErrNum Then
               ConstString$ = LCase$(ConstString$)
            Else
               'UL 32 introduced following constant
               ConstString$ = "BADEXPCHAN"
            End If
         End If
      Case 72
         ConstString$ = "WRONGADRANGE"
         If WRONGADRANGE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 73
         ConstString$ = "OUTOFRANGE"
         If OUTOFRANGE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 74
         ConstString$ = "BADTEMPSCALE"
         If BADTEMPSCALE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 75
         ConstString$ = "BADERRCODE"
         If BADERRCODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 76
         ConstString$ = "NOQUEUE"
         If NOQUEUE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 77
         ConstString$ = "CONTINUOUSCOUNT"
         If CONTINUOUSCOUNT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 78
         ConstString$ = "UNDERRUN"
         If UNDERRUN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 79
         ConstString$ = "BADMEMMODE"
         If BADMEMMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 80
         ConstString$ = "FREQOVERRUN"
         If FREQOVERRUN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 81
         ConstString$ = "NOCJCCHAN"
         If NOCJCCHAN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 82
         ConstString$ = "BADCHIPNUM"
         If BADCHIPNUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 83
         ConstString$ = "DIGNOTENABLED"
         If DIGNOTENABLED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 84
         ConstString$ = "CONVERT16BITS"
         If CONVERT16BITS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 85
         ConstString$ = "NOMEMBOARD"
         If NOMEMBOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 86
         ConstString$ = "DTACTIVE"
         If DTACTIVE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 87
         ConstString$ = "NOTMEMCONF"
         If NOTMEMCONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 88
         ConstString$ = "ODDCHAN"
         If ODDCHAN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 89
         ConstString$ = "CTRNOINIT"
         If CTRNOINIT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 90
         ConstString$ = "NOT8536CTR"
         If NOT8536CTR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 91
         ConstString$ = "FREERUNNING"
         If FREERUNNING <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 92
         ConstString$ = "INTERRUPTED"
         If INTERRUPTED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 93
         ConstString$ = "NOSELECTORS"
         If NOSELECTORS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 94
         ConstString$ = "NOBURSTMODE"
         If NOBURSTMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 95
         ConstString$ = "NOTWINDOWSFUNC"
         If NOTWINDOWSFUNC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 96
         ConstString$ = "NOTSIMULCONF"
         If NOTSIMULCONF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 97
         ConstString$ = "EVENODDMISMATCH"
         If EVENODDMISMATCH <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 98
         ConstString$ = "M1RATEWARNING"
         If M1RATEWARNING <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 99
         ConstString$ = "NOTRS485"
         If NOTRS485 <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 100
         ConstString$ = "NOTDOSFUNC"
         If NOTDOSFUNC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 101
         ConstString$ = "RANGEMISMATCH"
         If RANGEMISMATCH <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 102
         ConstString$ = "CLOCKTOOSLOW"
         If CLOCKTOOSLOW <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 103
         ConstString$ = "BADCALFACTORS"
         If BADCALFACTORS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 104
         ConstString$ = "BADCONFIGTYPE"
         If BADCONFIGTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 105
         ConstString$ = "BADCONFIGITEM"
         If BADCONFIGITEM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 106
         ConstString$ = "NOPCMCIABOARD"
         If NOPCMCIABOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 107
         ConstString$ = "NOBACKGROUND"
         If NOBACKGROUND <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 108
         ConstString$ = "STRINGTOOSHORT"
         If STRINGTOOSHORT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 109
         ConstString$ = "CONVERTEXTMEM"
         If CONVERTEXTMEM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 110
         ConstString$ = "BADEUADD"
         If BADEUADD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 111
         ConstString$ = "DAS16JRRATEWARNING"
         If DAS16JRRATEWARNING <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 112
         ConstString$ = "DAS08TOOLOWRATE"
         If DAS08TOOLOWRATE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 114
         ConstString$ = "AMBIGSENSORTYPE"
         If AMBIGSENSORTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 115
         ConstString$ = "NOSENSORTYPEONGP"
         If NOSENSORTYPEONGP <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 116
         ConstString$ = "NOCONVERSIONNEEDED"
         If NOCONVERSIONNEEDED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 117
         ConstString$ = "NOEXTCONTINUOUS"
         If NOEXTCONTINUOUS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 118
         ConstString$ = "INVALIDPRETRIGCONVERT"
         If INVALIDPRETRIGCONVERT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 119
         ConstString$ = "BADCTRREG"
         If BADCTRREG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 120
         ConstString$ = "BADTRIGTHRESHOLD"
         If BADTRIGTHRESHOLD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 121
         ConstString$ = "BADPCMSLOTREF"
         If BADPCMSLOTREF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 122
         ConstString$ = "AMBIGPCMSLOTREF"
         If AMBIGPCMSLOTREF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 123
         ConstString$ = "BADSENSORTYPE"
         If BADSENSORTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 124
         ConstString$ = "DELBOARDNOTEXIST"
         If DELBOARDNOTEXIST <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 125
         ConstString$ = "NOBOARDNAMEFILE"
         If NOBOARDNAMEFILE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 126
         ConstString$ = "CFGFILENOTFOUND"
         If CFGFILENOTFOUND <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 127
         ConstString$ = "NOVDDINSTALLED"
         If NOVDDINSTALLED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 128
         ConstString$ = "NOWINDOWSMEMORY"
         If NOWINDOWSMEMORY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 129
         ConstString$ = "OUTOFDOSMEMORY"
         If OUTOFDOSMEMORY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 130
         ConstString$ = "OBSOLETEOPTION"
         If OBSOLETEOPTION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 131
         ConstString$ = "NOPCMREGKEY"
         If NOPCMREGKEY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 132
         ConstString$ = "NOCBUL32SYS"
         If NOCBUL32SYS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 133
         ConstString$ = "NODMAMEMEMORY"
         If NODMAMEMEMORY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 134
         ConstString$ = "IRQNOTAVAILABLE"
         If IRQNOTAVAILABLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 135
         ConstString$ = "NOT7266CTR"
         If NOT7266CTR <> ErrNum Then ConstString$ = "NOT7266CTR"
      Case 136
         ConstString$ = "BADQUADRATURE"
         If BADQUADRATURE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 137
         ConstString$ = "BADCOUNTMODE"
         If BADCOUNTMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 138
         ConstString$ = "BADENCODING"
         If BADENCODING <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 139
         ConstString$ = "BADINDEXMODE"
         If BADINDEXMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 140
         ConstString$ = "BADINVERTINDEX"
         If BADINVERTINDEX <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 141
         ConstString$ = "BADFLAGPINS"
         If BADFLAGPINS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 142
         ConstString$ = "NOCTRSTATUS"
         If NOCTRSTATUS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 143
         ConstString$ = "NOGATEALLOWED"
         If NOGATEALLOWED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 144
         ConstString$ = "NOINDEXALLOWED"
         If NOINDEXALLOWED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 145
         ConstString$ = "OPENCONNECTION"
         If OPENCONNECTION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 146
         ConstString$ = "BMCONTINUOUSCOUNT"
         If BMCONTINUOUSCOUNT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 147
         ConstString$ = "BADCALLBACKFUNC"
         If BADCALLBACKFUNC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 148
         ConstString$ = "MBUSINUSE"
         If MBUSINUSE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 149
         ConstString$ = "MBUSNOCTLR"
         If MBUSNOCTLR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 150
         ConstString$ = "BADEVENTTYPE"
         If BADEVENTTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 151
         ConstString$ = "ALREADYENABLED"
         If ALREADYENABLED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 152
         ConstString$ = "BADEVENTSIZE"
         If BADEVENTSIZE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 153
         ConstString$ = "CANTINSTALLEVENT"
         If CANTINSTALLEVENT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 154
         ConstString$ = "BADBUFFERSIZE"
         If BADBUFFERSIZE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 155
         ConstString$ = "BADAIMODE"
         If BADAIMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 156
         ConstString$ = "BADSIGNAL"
         If BADSIGNAL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 157
         ConstString$ = "BADCONNECTION"
         If BADCONNECTION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 158
         ConstString$ = "BADINDEX"
         If BADINDEX <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 159
         ConstString$ = "NOCONNECTION"
         If NOCONNECTION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 160
         ConstString$ = "BADBURSTIOCOUNT"
         If BADBURSTIOCOUNT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 161
         ConstString$ = "DEADDEV"
         If DEADDEV <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 163
         ConstString$ = "INVALIDACCESS"
         If INVALIDACCESS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 164
         ConstString$ = "UNAVAILABLE"
         If UNAVAILABLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 165
         ConstString$ = "NOTREADY"
         If NOTREADY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 169
         ConstString$ = "BITUSEDFORALARM"
         If BITUSEDFORALARM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 170
         ConstString$ = "PORTUSEDFORALARM"
         If PORTUSEDFORALARM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 171
         ConstString$ = "PACEROVERRUN"
         If PACEROVERRUN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 172
         ConstString$ = "BADCHANTYPE"
         If BADCHANTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 173
         ConstString$ = "BADTRIGSENSE"
         If BADTRIGSENSE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 174
         ConstString$ = "BADTRIGCHAN"
         If BADTRIGCHAN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 175
         ConstString$ = "BADTRIGLEVEL"
         If BADTRIGLEVEL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 176
         ConstString$ = "NOPRETRIGMODE"
         If NOPRETRIGMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 177
         ConstString$ = "BADDEBOUNCETIME"
         If BADDEBOUNCETIME <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 178
         ConstString$ = "BADDEBOUNCETRIGMODE"
         If BADDEBOUNCETRIGMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 179
         ConstString$ = "BADMAPPEDCOUNTER"
         If BADMAPPEDCOUNTER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 180
         ConstString$ = "BADCOUNTERMODE"
         If BADCOUNTERMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 181
         ConstString$ = "BADTCCHANMODE"
         If BADTCCHANMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 182
         ConstString$ = "BADFREQUENCY"
         If BADFREQUENCY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 183
         ConstString$ = "BADEVENTPARAM"
         If BADEVENTPARAM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 184
         ConstString$ = "NONETIFC"
         If NONETIFC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 185
         ConstString$ = "DEADNETIFC"
         If DEADNETIFC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 186
         ConstString$ = "NOREMOTEACK"
         If NOREMOTEACK <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 187
         ConstString$ = "INPUTTIMEOUT"
         If INPUTTIMEOUT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 188
         ConstString$ = "MISMATCHSETPOINTCOUNT"
         If MISMATCHSETPOINTCOUNT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 189
         ConstString$ = "INVALIDSETPOINTLEVEL"
         If INVALIDSETPOINTLEVEL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 190
         ConstString$ = "INVALIDSETPOINTOUTPUTTYPE"
         If INVALIDSETPOINTOUTPUTTYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 191
         ConstString$ = "INVALIDSETPOINTOUTPUTVALUE"
         If INVALIDSETPOINTOUTPUTVALUE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 192
         ConstString$ = "INVALIDSETPOINTLIMITS"
         If INVALIDSETPOINTLIMITS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 193
         ConstString$ = "STRINGTOOLONG"
         If STRINGTOOLONG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 194
         ConstString$ = "INVALIDLOGIN"
         If INVALIDLOGIN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 195
         ConstString$ = "SESSIONINUSE"
         If SESSIONINUSE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 196
         ConstString$ = "NOEXTPOWER"
         If NOEXTPOWER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 197
         ConstString$ = "BADDUTYCYCLE"
         If BADDUTYCYCLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 199
         ConstString$ = "BADINITIALDELAY"
         If BADINITIALDELAY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      
      Case 200
         ConstString$ = "INTERNALERR"
         If INTERNALERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 201
         ConstString$ = "CANT_LOCK_DMA_BUF"
         If CANT_LOCK_DMA_BUF <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 202
         ConstString$ = "DMA_IN_USE"
         If DMA_IN_USE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 203
         ConstString$ = "BAD_MEM_HANDLE"
         If BAD_MEM_HANDLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      
      Case 300
         ConstString$ = "INTERNALERR32"
         If INTERNALERR32 <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 304
         ConstString$ = "CFG_FILE_READ_FAILURE"
         If CFG_FILE_READ_FAILURE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 305
         ConstString$ = "CFG_FILE_WRITE_FAILURE"
         If CFG_FILE_WRITE_FAILURE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 308
         ConstString$ = "CFGFILE_CANT_OPEN"
         If CFGFILE_CANT_OPEN <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 325
         ConstString$ = "BAD_RTD_CONVERSION"
         If BAD_RTD_CONVERSION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 326
         ConstString$ = "NO_PCI_BIOS"
         If NO_PCI_BIOS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 327
         ConstString$ = "BAD_PCI_INDEX"
         If BAD_PCI_INDEX <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 328
         ConstString$ = "NO_PCI_BOARD"
         If NO_PCI_BOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 334
         ConstString$ = "CANT_INSTALL_INT"
         If CANT_INSTALL_INT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 343
         ConstString$ = "NO_USB_BOARD"
         If NO_USB_BOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 344
         ConstString$ = "NOMOREFILES"
         If NOMOREFILES <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 345
         ConstString$ = "BADFILENUMBER"
         If BADFILENUMBER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 346
         ConstString$ = "INVALIDSTRUCTSIZE"
         If INVALIDSTRUCTSIZE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 347
         ConstString$ = "LOSSOFDATA"
         If LOSSOFDATA <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 348
         ConstString$ = "INVALIDBINARYFILE"
         If INVALIDBINARYFILE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 349
         ConstString$ = "INVALIDDELIMITER"
         If INVALIDDELIMITER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 350
         ConstString$ = "NO_BTH_BOARD"
         If NO_BTH_BOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 351
         ConstString$ = "NO_NET_BOARD"
         If NO_NET_BOARD <> ErrNum Then ConstString$ = LCase$(ConstString$)

      Case 400
         ConstString$ = "PCMCIAERRS"
         If PCMCIAERRS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 501
         ConstString$ = "DOSBADFUNC"
         If DOSBADFUNC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 502
         ConstString$ = "DOSFILENOTFOUND"
         If DOSFILENOTFOUND <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 503
         ConstString$ = "DOSPATHNOTFOUND"
         If DOSPATHNOTFOUND <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 504
         ConstString$ = "DOSNOHANDLES"
         If DOSNOHANDLES <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 505
         ConstString$ = "DOSACCESSDENIED"
         If DOSACCESSDENIED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 506
         ConstString$ = "DOSINVALIDHANDLE"
         If DOSINVALIDHANDLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 507
         ConstString$ = "DOSNOMEMORY"
         If DOSNOMEMORY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 515
         ConstString$ = "DOSBADDRIVE"
         If DOSBADDRIVE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 518
         ConstString$ = "DOSTOOMANYFILES"
         If DOSTOOMANYFILES <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 519
         ConstString$ = "DOSWRITEPROTECT"
         If DOSWRITEPROTECT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 521
         ConstString$ = "DOSDRIVENOTREADY"
         If DOSDRIVENOTREADY <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 525
         ConstString$ = "DOSSEEKERROR"
         If DOSSEEKERROR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 529
         ConstString$ = "DOSWRITEFAULT"
         If DOSWRITEFAULT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 530
         ConstString$ = "DOSREADFAULT"
         If DOSREADFAULT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 531
         ConstString$ = "DOSGENERALFAULT"
         If DOSGENERALFAULT <> ErrNum Then ConstString$ = LCase$(ConstString$)

      Case 603
         ConstString$ = "WIN_CANNOT_ENABLE_INT"
         If WIN_CANNOT_ENABLE_INT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 605
         ConstString$ = "WIN_CANNOT_DISABLE_INT"
         If WIN_CANNOT_DISABLE_INT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 606
         ConstString$ = "WIN_CANT_PAGE_LOCK_BUFFER"
         If WIN_CANT_PAGE_LOCK_BUFFER <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 630
         ConstString$ = "NO_PCM_CARD"
         If NO_PCM_CARD <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1000
         ConstString$ = "NOTEDSSENSOR"
         If NOTEDSSENSOR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1001
         ConstString$ = "INVALIDTEDSSENSOR"
         If INVALIDTEDSSENSOR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1002
         ConstString$ = "CALIBRATIONFAILED"
         If CALIBRATIONFAILED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1003
         ConstString$ = "BITUSEDFORTERMINALCOUNTSTATUS"
         If BITUSEDFORTERMINALCOUNTSTATUS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1004
         ConstString$ = "PORTUSEDFORTERMINALCOUNTSTATUS"
         If PORTUSEDFORTERMINALCOUNTSTATUS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1005
         ConstString$ = "BADEXCITATION"
         If BADEXCITATION <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1006
         ConstString$ = "BADBRIDGETYPE"
         If BADBRIDGETYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1007
         ConstString$ = "BADLOADVAL"
         If BADLOADVAL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1008
         ConstString$ = "BADTICKSIZE"
         If BADTICKSIZE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1013
         ConstString$ = "BTHCONNECTIONFAILED"
         If BTHCONNECTIONFAILED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1014
         ConstString$ = "INVALIDBTHFRAME"
         If INVALIDBTHFRAME <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1015
         ConstString$ = "BADTRIGEVENT"
         If BADTRIGEVENT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1016
         ConstString$ = "NETCONNECTIONFAILED"
         If NETCONNECTIONFAILED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1017
         ConstString$ = "DATASOCKETCONNECTIONFAILED"
         If DATASOCKETCONNECTIONFAILED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1018
         ConstString$ = "INVALIDNETFRAME"
         If INVALIDNETFRAME <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1019
         ConstString$ = "NETTIMEOUT"
         If NETTIMEOUT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1020
         ConstString$ = "NETDEVNOTFOUND"
         If NETDEVNOTFOUND <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1021
         ConstString$ = "INVALIDCONNECTIONCODE"
         If INVALIDCONNECTIONCODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1022
         ConstString$ = "CONNECTIONCODEIGNORED"
         If CONNECTIONCODEIGNORED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1023
         ConstString$ = "NETDEVINUSE"
         If NETDEVINUSE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1024
         ConstString$ = "NETDEVINUSEBYANOTHERPROC"
         If NETDEVINUSEBYANOTHERPROC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1025
         ConstString$ = "SOCKETDISCONNECTED"
         If SOCKETDISCONNECTED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1026
         ConstString$ = "BOARDNUMINUSE"
         If BOARDNUMINUSE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1027
         ConstString$ = "DEVALREADYCREATED"
         If DEVALREADYCREATED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1028
         ConstString$ = "BOARDNOTEXIST"
         If BOARDNOTEXIST <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1029
         ConstString$ = "INVALIDNETHOST"
         If INVALIDNETHOST <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1030
         ConstString$ = "INVALIDNETPORT"
         If INVALIDNETPORT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1031
         ConstString$ = "INVALIDIFC"
         If INVALIDIFC <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1032
         ConstString$ = "INVALIDAIINPUTMODE"
         If INVALIDAIINPUTMODE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1033
         ConstString$ = "AIINPUTMODENOTCONFIGURABLE"
         If AIINPUTMODENOTCONFIGURABLE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1034
         ConstString$ = "INVALIDEXTPACEREDGE"
         If INVALIDEXTPACEREDGE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case 1035
         ConstString$ = "CMREXCEEDED"
         If CMREXCEEDED <> ErrNum Then ConstString$ = LCase$(ConstString$)

      Case &HE0001001
         ConstString$ = "CBUSB_FILE_OPEN_FAILED"
         If CBUSB_FILE_OPEN_FAILED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001002
         ConstString$ = "CBUSB_REGISTRY_OPEN_ERROR"
         If CBUSB_REGISTRY_OPEN_ERROR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001003
         ConstString$ = "CBUSB_REGISTRY_QUERY_ERROR"
         If CBUSB_REGISTRY_QUERY_ERROR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001004
         ConstString$ = "CBUSB_DEVICE_OPEN_FAILED"
         If CBUSB_DEVICE_OPEN_FAILED <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001005
         ConstString$ = "CBUSB_NO_DETECT"
         If CBUSB_NO_DETECT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001006
         ConstString$ = "CBUSB_INVALID_DEVICE_NUM"
         If CBUSB_INVALID_DEVICE_NUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001007
         ConstString$ = "CBUSB_INVALID_PORT_NUM"
         If CBUSB_INVALID_PORT_NUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001008
         ConstString$ = "CBUSB_INVALID_BIT_NUM"
         If CBUSB_INVALID_BIT_NUM <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001009
         ConstString$ = "CBUSB_NO_DEVICES_IN_CFG"
         If CBUSB_NO_DEVICES_IN_CFG <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000100A
         ConstString$ = "CBUSB_AOUT_ERR"
         If CBUSB_AOUT_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000100B
         ConstString$ = "CBUSB_AIN_REQUEST_ERR"
         If CBUSB_AIN_REQUEST_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000100C
         ConstString$ = "CBUSB_AIN_READ_ERR"
         If CBUSB_AIN_READ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000100D
         ConstString$ = "CBUSB_CINIT_ERR"
         If CBUSB_CINIT_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000100E
         ConstString$ = "CBUSB_CIN32_REQUEST_ERR"
         If CBUSB_CIN32_REQUEST_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000100F
         ConstString$ = "CBUSB_CIN32_READ_ERR"
         If CBUSB_CIN32_READ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001010
         ConstString$ = "CBUSB_DCONFIG_ERR"
         If CBUSB_DCONFIG_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001011
         ConstString$ = "CBUSB_DIN_REQUEST_ERR"
         If CBUSB_DIN_REQUEST_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001012
         ConstString$ = "CBUSB_DIN_READ_ERR"
         If CBUSB_DIN_READ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001013
         ConstString$ = "CBUSB_DOUT_ERR"
         If CBUSB_DOUT_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001014
         ConstString$ = "CBUSB_AINSCAN_REQUEST_ERR"
         If CBUSB_AINSCAN_REQUEST_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001015
         ConstString$ = "CBUSB_AINSCAN_READ_ERR"
         If CBUSB_AINSCAN_READ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001016
         ConstString$ = "CBUSB_MAXRATE_ERR"
         If CBUSB_MAXRATE_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001017
         ConstString$ = "CBUSB_MINRATE_ERR"
         If CBUSB_MINRATE_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001018
         ConstString$ = "CBUSB_MAXCOUNT_ERR"
         If CBUSB_MAXCOUNT_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001019
         ConstString$ = "CBUSB_AINSTOP_ERR"
         If CBUSB_AINSTOP_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000101A
         ConstString$ = "CBUSB_AIGETDATA_ERR"
         If CBUSB_AIGETDATA_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000101B
         ConstString$ = "CBUSB_MAXCONT_RATE_ERR"
         If CBUSB_MAXCONT_RATE_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000101C
         ConstString$ = "CBUSB_DCONFIGBIT_ERR"
         If CBUSB_DCONFIGBIT_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000101D
         ConstString$ = "CBUSB_GETSNUM_ERR"
         If CBUSB_GETSNUM_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000101E
         ConstString$ = "CBUSB_SETSNUM_ERR"
         If CBUSB_SETSNUM_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000101F
         ConstString$ = "CBUSB_LOADQ_ERR"
         If CBUSB_LOADQ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001020
         ConstString$ = "CBUSB_BAD_CHANNEL"
         If CBUSB_BAD_CHANNEL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001021
         ConstString$ = "CBUSB_BAD_RANGE"
         If CBUSB_BAD_RANGE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001022
         ConstString$ = "CBUSB_INVALID_DATA"
         If CBUSB_INVALID_DATA <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001023
         ConstString$ = "CBUSB_MAXQ_CHANS"
         If CBUSB_MAXQ_CHANS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001024
         ConstString$ = "CBUSB_DATA_RANGE_ERR"
         If CBUSB_DATA_RANGE_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001025
         ConstString$ = "CBUSB_OVERRUN_ERR"
         If CBUSB_OVERRUN_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001026
         ConstString$ = "CBUSB_SCAN_IN_PROGRESS"
         If CBUSB_SCAN_IN_PROGRESS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001027
         ConstString$ = "CBUSB_THREAD_FAIL"
         If CBUSB_THREAD_FAIL <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001028
         ConstString$ = "CBUSB_BUFSIZ_ERR"
         If CBUSB_BUFSIZ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001029
         ConstString$ = "CBUSB_RESET_ERR"
         If CBUSB_RESET_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000102A
         ConstString$ = "CBUSB_PACKET_SYNC_ERR"
         If CBUSB_PACKET_SYNC_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000102B
         ConstString$ = "CBUSB_INVALID_ADDRESS"
         If CBUSB_INVALID_ADDRESS <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000102C
         ConstString$ = "CBUSB_MEMREAD_ERR"
         If CBUSB_MEMREAD_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000102D
         ConstString$ = "CBUSB_MEMWRITE_ERR"
         If CBUSB_MEMWRITE_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000102E
         ConstString$ = "CBUSB_CAL_READ_ERR"
         If CBUSB_CAL_READ_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE000102F
         ConstString$ = "CBUSB_NOSUPPORT"
         If CBUSB_NOSUPPORT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001030
         ConstString$ = "CBUSB_SETTRIG_ERR"
         If CBUSB_SETTRIG_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001031
         ConstString$ = "CBUSB_INVALID_TRIG_TYPE"
         If CBUSB_INVALID_TRIG_TYPE <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001032
         ConstString$ = "CBUSB_OPTIONS_ERR"
         If CBUSB_OPTIONS_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001033
         ConstString$ = "CBUSB_INVALID_TIMEOUT"
         If CBUSB_INVALID_TIMEOUT <> ErrNum Then ConstString$ = LCase$(ConstString$)
      Case &HE0001034
         ConstString$ = "CBUSB_PORT_DIR_ERR"
         If CBUSB_PORT_DIR_ERR <> ErrNum Then ConstString$ = LCase$(ConstString$)

      Case Else
         ConstString$ = "No constant defined"
   End Select
   GetErrorConst = ConstString$

End Function

Function GetExpInfoString(XIType As Integer) As String

   Select Case XIType
      Case XIBOARDTYPE
         GetExpInfoString = "XIBOARDTYPE"
      Case XIMUXADCHAN1
         GetExpInfoString = "XIMUXADCHAN1"
      Case XIMUXADCHAN2
         GetExpInfoString = "XIMUXADCHAN2"
      Case XIRANGE1
         GetExpInfoString = "XIRANGE1"
      Case XIRANGE2
         GetExpInfoString = "XIRANGE2"
      Case XICJCCHAN
         GetExpInfoString = "XICJCCHAN"
      Case XITHERMTYPE
         GetExpInfoString = "XITHERMTYPE"
      Case XINUMEXPCHANS
         GetExpInfoString = "XINUMEXPCHANS"
      Case XIPARENTBOARD
         GetExpInfoString = "XIPARENTBOARD"
      Case XISPARE0
         GetExpInfoString = "XISPARE0"
   End Select

End Function

Private Sub GetRangeInfo(ByVal BoardName As String, ByVal Range As Integer, RangeString As String, RangeVolts As Single)
   
   Select Case Range
      Case NOTUSED
         RangeString = "NOTUSED"
         RangeVolts = 0
      Case BIP5VOLTS
         RangeString = "BIP5VOLTS"
         RangeVolts = 10
      Case BIP10VOLTS
         Select Case BoardName
            Case "PCI-DAC6702", "PCI-DAC6703", "PCI-DAC6704"
               RangeString = "BIP10VOLTS"
               RangeVolts = 20.48
            'Case "USB-6003"
            '   RangeString = "BIP10.5VOLTS"
            '   RangeVolts = 21#
            Case Else
              RangeString = "BIP10VOLTS"
              RangeVolts = 20
         End Select
      Case BIP2PT5VOLTS
         RangeString = "BIP2PT5VOLTS"
         RangeVolts = 5
      Case BIP1PT25VOLTS
         RangeString = "BIP1PT25VOLTS"
         RangeVolts = 2.5
      Case BIP1VOLTS
         RangeString = "BIP1VOLTS"
         RangeVolts = 2
      Case BIPPT625VOLTS
         RangeString = "BIPPT625VOLTS"
         RangeVolts = 1.25
      Case BIPPT5VOLTS
         RangeString = "BIPPT5VOLTS"
         RangeVolts = 1
      Case BIPPT1VOLTS
         RangeString = "BIPPT1VOLTS"
         RangeVolts = 0.2
      Case BIPPT05VOLTS
         RangeString = "BIPPT05VOLTS"
         RangeVolts = 0.1
      Case BIPPT01VOLTS
         RangeString = "BIPPT01VOLTS"
         RangeVolts = 0.02
      Case BIPPT005VOLTS
         RangeString = "BIPPT005VOLTS"
         RangeVolts = 0.01
      Case BIP1PT67VOLTS
         RangeString = "BIP1PT67VOLTS"
         RangeVolts = 3.34
      Case BIPPT312VOLTS
         RangeString = "BIPPT312VOLTS"
         RangeVolts = 0.625
      Case BIPPT156VOLTS
         RangeString = "BIPPT156VOLTS"
         RangeVolts = 0.3125
      Case BIPPT078VOLTS
         RangeString = "BIPPT078VOLTS"
         RangeVolts = 0.15625
      Case BIP60VOLTS
         RangeString = "BIP60VOLTS"
         RangeVolts = 120  '125.274
         'to do - switch to toengunits
      Case BIP15VOLTS
         RangeString = "BIP15VOLTS"
         RangeVolts = 30
      Case BIPPT125VOLTS
         RangeString = "BIPPT125VOLTS"
         RangeVolts = 0.25
      Case BIPPT25VOLTS
         RangeString = "BIPPT25VOLTS"
         RangeVolts = 0.5
      Case BIPPT2VOLTS
         RangeString = "BIPPT2VOLTS"
         RangeVolts = 0.4
      Case BIP2VOLTS
         RangeString = "BIP2VOLTS"
         RangeVolts = 4
      Case BIP20VOLTS
         RangeString = "BIP20VOLTS"
         RangeVolts = 40
      Case BIP4VOLTS
         RangeString = "BIP4VOLTS"
         RangeVolts = 8
      Case BIP30VOLTS
         RangeString = "BIP30VOLTS"
         RangeVolts = 60
      Case BIPPT025VOLTSPERVOLT
         RangeString = "BIPPT025VOLTSPERVOLT"
         RangeVolts = 0.05
      Case BIPPT073125VOLTS
         RangeString = "BIPPT073125VOLTS"
         RangeVolts = 0.14625
      Case UNI10VOLTS
         RangeString = "UNI10VOLTS"
         RangeVolts = 10
      Case UNI5VOLTS
         RangeString = "UNI5VOLTS"
         RangeVolts = 5
      Case UNI2PT5VOLTS
         RangeString = "UNI2PT5VOLTS"
         RangeVolts = 2.5
      Case UNI2VOLTS
         RangeString = "UNI2VOLTS"
         RangeVolts = 2
      Case UNI1PT25VOLTS
         RangeString = "UNI1PT25VOLTS"
         RangeVolts = 1.25
      Case UNI1VOLTS
         RangeString = "UNI1VOLTS"
         RangeVolts = 1
      Case UNIPT1VOLTS
         RangeString = "UNIPT1VOLTS"
         RangeVolts = 0.1
      Case UNIPT01VOLTS
         RangeString = "UNIPT01VOLTS"
         RangeVolts = 0.01
      Case UNIPT02VOLTS
         RangeString = "UNIPT02VOLTS"
         RangeVolts = 0.02
      Case UNI1PT67VOLTS
         RangeString = "UNI1PT67VOLTS"
         RangeVolts = 1.67
      Case UNIPT5VOLTS
         RangeString = "UNIPT5VOLTS"
         RangeVolts = 0.5
      Case UNIPT25VOLTS
         RangeString = "UNIPT25VOLTS"
         RangeVolts = 0.25
      Case UNIPT2VOLTS
         RangeString = "UNIPT2VOLTS"
         RangeVolts = 0.2
      Case UNIPT05VOLTS
         RangeString = "UNIPT05VOLTS"
         RangeVolts = 0.05
      Case UNI4VOLTS
         RangeString = "UNI4VOLTS"
         RangeVolts = 4.096
      Case MA4TO20
         RangeString = "MA4TO20"
         RangeVolts = 16
      Case MA2to10
         RangeString = "MA2to10"
         RangeVolts = 8
      Case MA1TO5
         RangeString = "MA1TO5"
         RangeVolts = 4
      Case MAPT5TO2PT5
         RangeString = "MAPT5TO2PT5"
         RangeVolts = 2
      Case MA0TO20
         RangeString = "MA0TO20"
         RangeVolts = 20
      Case BIPPT025AMPS
         RangeString = "BIPPT025AMPS"
         RangeVolts = 0.05

'      Case GAIN2_DIFF
'         RangeString = "GAIN2_DIFF"
'         RangeVolts = 20
'      Case GAIN4_DIFF
'         RangeString = "GAIN4_DIFF"
'         RangeVolts = 10
'      Case GAIN5_DIFF
'         RangeString = "GAIN5_DIFF"
'         RangeVolts = 8
'      Case GAIN8_DIFF
'         RangeString = "GAIN8_DIFF"
'         RangeVolts = 5
'      Case GAIN10_DIFF
'         RangeString = "GAIN10_DIFF"
'         RangeVolts = 4
'      Case GAIN16_DIFF
'         RangeString = "GAIN16_DIFF"
'         RangeVolts = 2.5
      'Case GAIN20_DIFF
      '   RangeString = "GAIN20_DIFF"
      '   RangeVolts = 2
      
   End Select

End Sub

Function GetRangeVolts(Range As Integer) As Single

   GetRangeInfo msBoardName, Range, RangeString$, RangeVolts!
   GetRangeVolts = RangeVolts!

End Function

Private Sub GetRangeInfoD(ByVal BoardName As String, ByVal Range As Integer, RangeString As String, RangeVolts As Double)
   
   Select Case Range
      Case NOTUSED
         RangeString = "NOTUSED"
         RangeVolts = 0
      Case BIP5VOLTS
         RangeString = "BIP5VOLTS"
         RangeVolts = 10
      Case BIP10VOLTS
         Select Case BoardName
            Case "PCI-DAC6702", "PCI-DAC6703", "PCI-DAC6704"
               RangeString = "BIP10VOLTS"
               RangeVolts = 20.48
            Case Else
              RangeString = "BIP10VOLTS"
              RangeVolts = 20
         End Select
      Case BIP2PT5VOLTS
         RangeString = "BIP2PT5VOLTS"
         RangeVolts = 5
      Case BIP1PT25VOLTS
         RangeString = "BIP1PT25VOLTS"
         RangeVolts = 2.5
      Case BIP1VOLTS
         RangeString = "BIP1VOLTS"
         RangeVolts = 2
      Case BIPPT625VOLTS
         RangeString = "BIPPT625VOLTS"
         RangeVolts = 1.25
      Case BIPPT5VOLTS
         RangeString = "BIPPT5VOLTS"
         RangeVolts = 1
      Case BIPPT1VOLTS
         RangeString = "BIPPT1VOLTS"
         RangeVolts = 0.2
      Case BIPPT05VOLTS
         RangeString = "BIPPT05VOLTS"
         RangeVolts = 0.1
      Case BIPPT01VOLTS
         RangeString = "BIPPT01VOLTS"
         RangeVolts = 0.02
      Case BIPPT005VOLTS
         RangeString = "BIPPT005VOLTS"
         RangeVolts = 0.01
      Case BIP1PT67VOLTS
         RangeString = "BIP1PT67VOLTS"
         RangeVolts = 3.34
      Case BIPPT25VOLTS
         RangeString = "BIPPT25VOLTS"
         RangeVolts = 0.5
      Case BIPPT2VOLTS
         RangeString = "BIPPT2VOLTS"
         RangeVolts = 0.4
      Case BIP2VOLTS
         RangeString = "BIP2VOLTS"
         RangeVolts = 4
      Case BIP20VOLTS
         RangeString = "BIP20VOLTS"
         RangeVolts = 40
      Case BIP4VOLTS
         RangeString = "BIP4VOLTS"
         RangeVolts = 8
      Case BIPPT312VOLTS
         RangeString = "BIPPT312VOLTS"
         RangeVolts = 0.625
      Case BIPPT156VOLTS
         RangeString = "BIPPT156VOLTS"
         RangeVolts = 0.3125
      Case BIPPT078VOLTS
         RangeString = "BIPPT078VOLTS"
         RangeVolts = 0.15625
      Case BIP60VOLTS
         RangeString = "BIP60VOLTS"
         RangeVolts = 120  '125.274
         'to do - switch to toengunits
      Case BIP15VOLTS
         RangeString = "BIP15VOLTS"
         RangeVolts = 30
      Case BIPPT125VOLTS
         RangeString = "BIPPT125VOLTS"
         RangeVolts = 0.25
      Case BIPPT025VOLTSPERVOLT
         RangeString = "BIPPT025VOLTSPERVOLT"
         RangeVolts = 0.05
      Case BIP30VOLTS
         RangeString = "BIP30VOLTS"
         RangeVolts = 60
      
      Case UNI10VOLTS
         RangeString = "UNI10VOLTS"
         RangeVolts = 10
      Case UNI5VOLTS
         RangeString = "UNI5VOLTS"
         RangeVolts = 5
      Case UNI2PT5VOLTS
         RangeString = "UNI2PT5VOLTS"
         RangeVolts = 2.5
      Case UNI2VOLTS
         RangeString = "UNI2VOLTS"
         RangeVolts = 2
      Case UNI1PT25VOLTS
         RangeString = "UNI1PT25VOLTS"
         RangeVolts = 1.25
      Case UNI1VOLTS
         RangeString = "UNI1VOLTS"
         RangeVolts = 1
      Case UNIPT1VOLTS
         RangeString = "UNIPT1VOLTS"
         RangeVolts = 0.1
      Case UNIPT01VOLTS
         RangeString = "UNIPT01VOLTS"
         RangeVolts = 0.01
      Case UNIPT02VOLTS
         RangeString = "UNIPT02VOLTS"
         RangeVolts = 0.02
      Case UNI1PT67VOLTS
         RangeString = "UNI1PT67VOLTS"
         RangeVolts = 1.67
      Case UNIPT5VOLTS
         RangeString = "UNIPT5VOLTS"
         RangeVolts = 0.5
      Case UNIPT25VOLTS
         RangeString = "UNIPT25VOLTS"
         RangeVolts = 0.25
      Case UNIPT2VOLTS
         RangeString = "UNIPT2VOLTS"
         RangeVolts = 0.2
      Case UNIPT05VOLTS
         RangeString = "UNIPT05VOLTS"
         RangeVolts = 0.05
      Case UNI4VOLTS
         RangeString = "UNI4VOLTS"
         RangeVolts = 4.096
      Case MA4TO20
         RangeString = "MA4TO20"
         RangeVolts = 16
      Case MA2to10
         RangeString = "MA2to10"
         RangeVolts = 8
      Case MA1TO5
         RangeString = "MA1TO5"
         RangeVolts = 4
      Case MAPT5TO2PT5
         RangeString = "MAPT5TO2PT5"
         RangeVolts = 2
      Case MA0TO20
         RangeString = "MA0TO20"
         RangeVolts = 20
      Case BIPPT025AMPS
         RangeString = "BIPPT025AMPS"
         RangeVolts = 0.05
   End Select

End Sub

Function GetRangeVoltsD(Range As Integer) As Double

   GetRangeInfoD msBoardName, Range, RangeString$, RangeVolts#
   GetRangeVoltsD = RangeVolts#

End Function

Function GetTrigString(TrigVal As Integer) As String

   Select Case TrigVal
      Case TRIGABOVE
         GetTrigString = "TRIGABOVE"
      Case TRIGBELOW
         GetTrigString = "TRIGBELOW"
      Case GATENEGHYS
         GetTrigString = "GATENEGHYS"
      Case GATEPOSHYS
         GetTrigString = "GATEPOSHYS"
      Case GATEABOVE
         GetTrigString = "GATEABOVE"
      Case GATEBELOW
         GetTrigString = "GATEBELOW"
      Case GATEINWINDOW
         GetTrigString = "GATEINWINDOW"
      Case GATEOUTWINDOW
         GetTrigString = "GATEOUTWINDOW"
      Case GATEHIGH
         GetTrigString = "GATEHIGH"
      Case GATELOW
         GetTrigString = "GATELOW"
      Case TRIGHIGH
         GetTrigString = "TRIGHIGH"
      Case TRIGLOW
         GetTrigString = "TRIGLOW"
      Case TRIGPOSEDGE
         GetTrigString = "TRIGPOSEDGE"
      Case TRIGNEGEDGE
         GetTrigString = "TRIGNEGEDGE"
      Case Else
         GetTrigString = "Invalid trigger code"
   End Select
   

End Function

Sub InitLocModule(BoardName As String)

   msBoardName = BoardName

End Sub

Function GetCfgItemIndex(CfgItem As Integer) As Integer

   'obsolete - moved to config form
   Stop
   LastItem% = 39
   For ListItem% = 0 To LastItem%
      MatchItem = Choose(ListItem% + 1, BIBASEADR, BIBOARDTYPE, BIINTLEVEL, BIDMACHAN, _
      BICLOCK, BIRANGE, BINUMADCHANS, BIUSESEXPS, BIDINUMDEVS, _
      BICINUMDEVS, BINUMDACHANS, BIWAITSTATE, BINUMIOPORTS)
      If IsNull(MatchItem) Then
         MatchItem = Choose(ListItem% - 12, BIDTBOARD, BINUMEXPS, BINUMTEMPCHANS, _
         BIFWVERSION, BISERIALNUM, BIDACUPDATEMODE, BIDACUPDATECMD, BIDACSTARTUP, _
         BIADTRIGCOUNT, BIADFIFOSIZE, BIADSOURCE, BICALOUTPUT, BISRCADPACER, _
         BIDIALARMMASK, BIDIDEBOUNCESTATE)
      End If
      If IsNull(MatchItem) Then
         MatchItem = Choose(ListItem% - 27, BIDIDEBOUNCETIME, BIPANID, BIRFCHANNEL, _
         BIRSS, BINODEID, BIDEVNOTES, BINETIOTIMEOUT, BIFACTORYID, BIHIDELOGINDLG, BIDACTRIGCOUNT, BIADRES, BIDACRES)
      End If
      If Val(MatchItem) = CfgItem Then Exit For
   Next
   
   If ListItem% > LastItem% Then
      'there is no constant defined so use actual value
      GetCfgItemIndex = CfgItem
   Else
      GetCfgItemIndex = ListItem%
   End If

   'cmbConfigItem.AddItem "BIINITIALIZED"   'obsolete
   'cmbConfigItem.AddItem "BIPARENTBOARD"   'obsolete

End Function

Function GetFixedRange(BoardName As String) As Long

   'the fixed range may be listed in the configuration file
   'this is required because the UL doesn't consistently return
   'the value of a fixed range - it returns NOTUSED for some
   'if the file exists, check if range is listed for this board
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = BoardName
   lpKeyName$ = "FixedRange"
   lpDefault$ = "-1"
   nSize% = 16
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   If Len(CustomParam$) Then Range& = Val(CustomParam$)
   GetFixedRange = Range&
   
End Function

Function GetFixedAOutRange(BoardName As String) As Long

   'the fixed range may be listed in the configuration file
   'this is required because the UL doesn't return
   'the value of a analog output ranges
   'if the file exists, check if range is listed for this board
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = BoardName
   lpKeyName$ = "FixedAOutRange"
   lpDefault$ = "-1"
   nSize% = 16
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   If Len(CustomParam$) Then Range& = Val(CustomParam$)
   GetFixedAOutRange = Range&
   
End Function

Public Function GetCustomRange(BoardName As String) As Variant
   
   lpFileName$ = "UniBoards.ini"
   lpApplicationName$ = BoardName
   lpKeyName$ = "NominalRange"
   lpDefault$ = ""
   nSize% = 16
   CustomParam$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, CustomParam$, nSize%, lpFileName$)
   CustomParam$ = Left$(CustomParam$, StringSize%)
   If Len(CustomParam$) Then CustomRange = Val(CustomParam$)
   GetCustomRange = CustomRange

End Function

Public Function GetManualRange(ByVal BoardName As String, ByVal BoardNum As Integer, FixedRange As Long) As Long

   'returns current range setting and FixedRange if any
   'first check if range is listed in extended UL info
   Dim NoForm As Form
   Dim NoExtInfo As Boolean
   
   ConfigItem& = BIADAIMODE 'BIADNUMCHANMODES
   'ConfigItem& = BIADCHANMODE
   
   ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
   If (ULStat = 0) Then
      Select Case ConfigVal&
         Case SINGLE_ENDED
            ConfigItem& = BIADNUMSERANGES
            ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
            If (ULStat = 0) And (ConfigVal& = 1) Then
               ConfigItem& = BIADSERANGE
               ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
               FixedRange = ConfigVal&
               GetManualRange = ConfigVal&
            Else
               NoExtInfo = True
            End If
         Case DIFFERENTIAL
            ConfigItem& = BIADNUMDIFFRANGES
            ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
            If (ULStat = 0) And (ConfigVal& = 1) Then
               ConfigItem& = BIADDIFFRANGE
               ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
               FixedRange = ConfigVal&
               GetManualRange = ConfigVal&
            Else
               NoExtInfo = True
            End If
         Case Else
            NoExtInfo = True
      End Select
   Else
      NoExtInfo = True
   End If
   
   If NoExtInfo Then
      'if not, check if manual range is listed in UniBoards.ini
      ManualRange& = GetFixedRange(BoardName)
      Select Case ManualRange&
         Case -1
            InfoType% = BOARDINFO
            ConfigItem& = BIRANGE
            DevNum% = 0
            ULStat = GetConfig520(InfoType%, BoardNum, DevNum%, ConfigItem&, ValConfig&)
            x% = SaveFunc(NoForm, GetConfig, ULStat, InfoType%, BoardNum, DevNum%, ConfigItem&, ValConfig&, A6, A7, A8, A9, A10, A11, 0)
            GetManualRange = ValConfig&
            FixedRange = ValConfig&
         Case Else
            FixedRange = ManualRange&
            GetManualRange = ManualRange&
      End Select
   End If
      
End Function

Public Function GetManualAOutRange(ByVal BoardName As String, ByVal BoardNum As Integer, FixedRange As Long) As Long

   'returns current range setting and FixedRange if any
   'first check if manual range is listed in UniBoards.ini
   Dim NoForm As Form
   
   ConfigItem& = BIDACNUMRANGES
   ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
   If ConfigVal& = 1 Then
      ConfigItem& = BIDACDEVRANGE
      ULStat = cbGetConfig(BOARDINFO, BoardNum, 0, ConfigItem&, ConfigVal&)
      FixedRange = ConfigVal&
      GetManualAOutRange = ConfigVal&
   Else
      ManualRange& = GetFixedAOutRange(BoardName)
      Select Case ManualRange&
         Case -1
            'ULStat = GetConfig520(InfoType%, BoardNum, DevNum%, ConfigItem%, ValConfig&)
            FixedRange = ValConfig&
         Case Else
            FixedRange = ManualRange&
            GetManualAOutRange = ManualRange&
      End Select
   End If
End Function


Function GetLoggerRange(RangeCode As String) As Integer

   Select Case RangeCode
      Case "5V"
         RangeValue% = BIP5VOLTS
      Case "10V"
         RangeValue% = BIP10VOLTS
      Case "1V"
         RangeValue% = BIP1VOLTS
      Case "30V"
         RangeValue% = BIP30VOLTS
      Case Else
         RangeValue% = "Unsupported"
   End Select
   GetLoggerRange = RangeValue%
   
End Function

