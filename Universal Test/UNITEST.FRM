VERSION 5.00
Begin VB.MDIForm mfmUniTest 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Universal Test"
   ClientHeight    =   2685
   ClientLeft      =   660
   ClientTop       =   3285
   ClientWidth     =   11790
   HelpContextID   =   10000
   Icon            =   "UNITEST.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "mfmUniTest"
   ScrollBars      =   0   'False
   Tag             =   "UniTest"
   Begin VB.PictureBox picCommands 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   588
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11760
      TabIndex        =   1
      Top             =   2100
      Width           =   11790
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "Logger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   1035
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "GPIB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.Timer tmrDDERespond 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   9000
         Top             =   60
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "Config&'n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   6120
         TabIndex        =   9
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdUtils 
         Appearance      =   0  'Flat
         Caption         =   "Uti&lities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7140
         TabIndex        =   8
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "&Misc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   5100
         TabIndex        =   7
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "&Counters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "Digital Ou&t"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3060
         TabIndex        =   5
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "Digital I&n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "Analog &Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1020
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdFormType 
         Appearance      =   0  'Flat
         Caption         =   "Analog &In"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label lblDDEInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   4515
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11760
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11790
   End
   Begin VB.Menu mnuFileGroup 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "mfmUniTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mfNoForm As Form

#If NETOPS Then
   Dim WithEvents ULNetLibrary As ULNet.ULNetClass
Attribute ULNetLibrary.VB_VarHelpID = -1
#End If

Private Sub cmdFormType_Click(Index As Integer)

   mfmUniTest.Picture1.Visible = False
   Select Case Index
      Case 7
         Success% = OpenGPIB()
         If Not Success% Then CloseGPIB 0
      Case Else
         LoadChildForm Index + 1
   End Select

End Sub

Private Sub cmdUtils_Click()

   If gnUniScript Then
      picCommands.Visible = False
   Else
      frmMain.SetFocus
   End If

End Sub

Private Sub cmdUtils_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then ErrorHandling True

End Sub

Private Function InitLibrary() As Integer

   On Error GoTo ConfigProblem
   Dim numDevices As Long
   Dim LibFound As Boolean
   
   LibStatus% = 7   'bitfield: 1=UL, 2=DotNet, 4=MsgDaq

   geErrFlow = 1
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "MainForm"
   lpKeyName$ = "Revision"
   nSize% = 6
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = Format$(CURRENTREVNUM, "0.00")

   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   Rev$ = Left$(lpReturnedString$, x%)
   RevNum! = Val(Rev$)
   
   lpKeyName$ = "Library"
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "Unilib"

   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, _
   lpReturnedString$, nSize%, lpFileName$)
   LibString$ = Left(lpReturnedString$, x%)
   Select Case LibString$
      Case "None"
         gnLibType = INVALIDLIB
         MsgBox "No library is selected. Choose a library under the File menu. " & _
         "If none are enabled, install one of the MCC libraries" & vbCrLf & _
         "(or copy the " & "required DLLs to the application directory" & _
         " - " & CurDir() & ").", vbInformation, "Select a Library"
      Case "Unilib"
         gnLibType = UNILIB
      Case "DotNet"
         gnLibType = NETLIB
      Case "MsgDaq"
         gnLibType = MSGLIB
      Case Else
         gnLibType = UNILIB
   End Select
   If gnLibType = UNILIB Then
      gnICalDisable = True
      lpKeyName$ = "InstaCal"
      nSize% = 16
      lpReturnedString$ = Space$(nSize%)
      lpDefault$ = "Enabled"
   
      x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, _
      lpReturnedString$, nSize%, lpFileName$)
      LibString$ = Left(lpReturnedString$, x%)
      If LibString$ = "Enabled" Then gnICalDisable = False
   End If
   DoEvents
   
   lpKeyName$ = "ErrorSettings"
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "001101"

   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, _
      lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   gnErrReporting = Val(Mid$(lpReturnedString$, 1, 1))
   gnErrHandling = Val(Mid$(lpReturnedString$, 2, 1))
   geErrFlow = Val(Mid$(lpReturnedString$, 3, 1))
   gnLocalErrDisp = Val(Mid$(lpReturnedString$, 4, 1))
   geWarnFlow = Val(Mid$(lpReturnedString$, 5, 1))
   gnLocalWarnDisp = Val(Mid$(lpReturnedString$, 6, 1))
   
   LibRevString$ = GetLibDLLRev("cbw32.dll")
   If LibRevString$ = "NoInfoFile" Then
      gfDLLRev = -2
   Else
      gfDLLRev = Val(LibRevString$)
   End If
   gbULLoaded = Not (LibRevString$ = "")
   
   CheckForUL% = True
   If (gnLibType = UNILIB) And gbULLoaded Then
      If gnICalDisable Then
         If LibSupportsFunction(IgnoreInstaCal) Then
            ULStat = cbIgnoreInstaCal()
            ErrReturn& = SaveFunc(mfNoForm, IgnoreInstaCal, ULStat, A1, _
                A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
            DevsFound = DiscoverDevices(USB_IFC + BLUETOOTH_IFC, True)
         End If
      End If
   End If
   
AfterDiscover:
   
   'ULStat = cbDeclareRevision(RevNum!)
   'x% = SaveFunc(Me, DeclareRevision, ULStat, RevNum!, _
   '   A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   'LibStat& = ULStat
   'ULStat = cbGetConfig(GIVERSION, 0, 0, 0, ConfigVal&)
   'If (ULStat = CFGFILENOTFOUND) And Not (gnLibType = UNILIB) Then
   '   If Not (LibStat& = 0) Then
   '      MsgBox "Although Universal Test is not currently configured to use " & _
   '      "the Universal Library, a small number of UL functions may be used." & vbCrLf & _
   '      "However, the Universal Library configuration file was not detected. " & _
   '      "These library functions may not work. " & vbCrLf & "Run InstaCal or " & _
   '      "uncheck the /Utilies/UL Creates Buffer menu item.", _
   '      vbCritical, "Universal Library Configuration File Not Found"
   '   End If
   'End If
   CheckForUL% = False
   
   Select Case gnLibType
      Case UNILIB
         ULStat = cbErrHandling(gnErrReporting, gnErrHandling)
         If SaveFunc(Me, ErrHandling, ULStat, gnErrReporting, _
            gnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, _
            A11, 0) Then Exit Function
      Case NETLIB
   End Select
   
   InitLibrary = LibStatus%
   gnInitializing = False
   Exit Function
   
ConfigProblem:
   If gnICalDisable Then
      gnICalDisable = False
      MsgBox "The current configuration for Universal Test is not " & _
      "supported. Configuration has been changed to compatible mode.", _
      vbCritical, "Universal Library Not Compatible"
      'mfmUniTest.picCommands.Enabled = False
      'frmMain.mnuScripter.Enabled = False
      DoEvents
      InitLibrary = LibStatus%
      Resume AfterDiscover
   End If
   If CheckForUL% Then LibStatus% = (LibStatus% Xor 1)
   If gnLibType = UNILIB Then
      gnLibType = INVALIDLIB
      MsgBox "The current configuration for Universal Test " _
         & "is for Universal Library, but the Universal Library " _
         & "was not detected. Select a different library.", _
         vbCritical, "Universal Library Not Available"
   End If
   InitLibrary = LibStatus%
   Exit Function

End Function

Private Sub lblDDEInfo_Change()

   'Stop
   'lblDDEInfo.Caption = "Change event triggered"
   'lblDDEInfo.LinkPoke

End Sub

Private Sub MDIForm_Activate()

   On Error GoTo CheckExternalCtl   'Resume Next
   If gnRunDefault Then
      gfOwner.cmdGo = True
   Else
      Screen.ActiveForm.cmdGo.SetFocus
   End If
   If gnCancel Then gfOwner.cmdStop = True
   Set gfOwner = Nothing
   gnRunDefault = False
   gnCancel = False
   If gnUniScript Then
      Me.HelpContextID = 30000
   Else
      Me.HelpContextID = 10000
   End If
   Exit Sub

CheckExternalCtl:
   If Screen.ActiveForm.Tag = "Main" Then Exit Sub
   If Screen.ActiveForm.Tag = "UniTest" Then Exit Sub
   If gnXternalCtl And gnRunDefault Then
      Screen.ActiveForm.cmdGo = True
   Else
      'Screen.ActiveForm.cmdGo.SetFocus
   End If
   gnRunDefault = False
   Exit Sub

End Sub

Private Sub MDIForm_Load()

#If NETOPS Then
   Set NetLibrary = New ULNet.ULNetClass
   mnThreading = True
#End If
   
   Dim ConfigFound As Boolean
   
   gnInitializing = True
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "DataFormat"
   nDefault% = 0
   lpKeyName$ = "Notation"
   gnHexVals = GetPrivateProfileInt(lpApplicationName$, _
   lpKeyName$, nDefault%, lpFileName$)
   
   lpApplicationName$ = "HistBuffer"
   lpKeyName$ = "BufferSize"
   nSize% = 6
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "10"
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
   lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   sDefault$ = Left$(lpReturnedString$, StringSize%)
   NumFuncs% = Val(sDefault$)
   If IsNumeric(sDefault$) Then
      If Val(sDefault$) < 2 Then NumFuncs% = 10
   Else
      NumFuncs% = 10
   End If
   SetHistory NumFuncs%, 0
   
   lpKeyName$ = "BufferedFuncs"
   nSize% = 80
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "0000000000000000000000000000000000000000000000000000000000000000"
   StringSize% = GetPrivateProfileString(lpApplicationName$, _
   lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   sDefault$ = Left$(lpReturnedString$, StringSize%)
   ReDim FuncList%(StringSize%)
   For i% = 1 To StringSize%
      FuncList%(i%) = Val(Mid$(lpReturnedString$, i%, 1)) = 1
   Next i%
   SetFuncsSaved FuncList%()

   lpApplicationName$ = "MsgLibrary"
   nDefault% = -1
   lpKeyName$ = "NameFormat"
   NameDef% = GetPrivateProfileInt(lpApplicationName$, _
   lpKeyName$, nDefault%, lpFileName$)
   If NameDef% < 0 Then
      NameSpec% = NAME_SERNO
   Else
      NameSpec% = NameDef%
   End If
   SetNameFormat NameSpec%
   
   If Not WIN32APP Then
      gnLibVer = WIN3X95
      Caption = "Universal Test: Universal Library for Windows 3.X and 95"
      If WIN32APP Then Caption = Caption & "   (Thunk)"
   Else
      gnLibVer = WIN95NT
      Caption = "Universal Test: Universal Library for Windows"
   End If
   InitVB
   LibStatus% = InitLibrary()
   'gbULLoaded = ((LibStatus% And 1) = 1)
   PositionForm MAIN_FORM, 0
   Select Case gnLibType
      Case INVALIDLIB
         mfmUniTest.picCommands.ENABLED = False
         frmMain.mnuScripter.ENABLED = False
         frmMain.mnuMCCCtl.Checked = False
      Case UNILIB
         If WIN32APP Then
            dum% = GetRegGroup(HKEY_LOCAL_MACHINE, "SOFTWARE\Universal Library", result&)
         Else
            dum% = GetRegGroup(HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Universal Library", result&)
         End If
         If dum% Then
            dumdum% = GetKeyValue(result&, "ConfigDir", KeyVal$)
            If (Not dumdum%) Or KeyVal$ = "" Then
               dumdum% = GetKeyValue(result&, "RootDir", KeyVal$)
            End If
         End If
         If dumdum% Then
            CFile$ = Dir(KeyVal$ & "*.cfg")
            ConfigFound = Not (CFile$ = "")
            ConfigFound = ConfigFound And (Not (CFile$ = "dummy.cfg"))
         
         End If
         If ConfigFound Then gnNumBoards = GetNumInstalled()
         frmMain.mnuLibrary(UNILIB).ENABLED = (gnLibType > INVALIDLIB)
         frmMain.mnuNoIcal.Checked = gnICalDisable
      Case NETLIB
         gnNumBoards = GetNumInstalled()
         'instantiate each board in .Net library
         For board% = 0 To gnNumBoards - 1
            BoardNumber& = gnBoardEnum(board%)
            BoardName$ = NetLibrary.CreateBoard(BoardNumber&)
         Next
         NumDevs% = NetLibrary.NumberOfDevices
      Case MSGLIB
         gnNumBoards = GetNumMsgBoards()
         frmMain.mnuLibrary(MSGLIB).ENABLED = (gnLibType > INVALIDLIB)
   End Select
   ConfigureLibrary gnLibType
   Load frmMain

End Sub

Private Sub MDIForm_Resize()

   If Me.WindowState <> 1 Then
      If Height < 2130 Then Height = 2130
      Picture1.Height = mfmUniTest.Height - picCommands.ScaleHeight
      If gnUniScript Then
         For CurForm% = 0 To Forms.Count - 1
            If Forms(CurForm%).Tag = "Scripter" Then
               If picCommands.Visible Then Offset& = picCommands.ScaleHeight
               Forms(CurForm%).Width = Me.ScaleWidth
               Forms(CurForm%).Top = Me.ScaleHeight - (Forms(CurForm%).Height - Offset&)
               Forms(CurForm%).Left = 0   'Me.Left
            End If
         Next CurForm%
      End If
   End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

   NameFormat% = GetNameFormat()
   lplFileName$ = "UniTest.ini"
   lpApplicationName$ = "MsgLibrary"
   lpKeyName$ = "NameFormat"
   lpString$ = Format(NameFormat%, "0")
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)
   
   If gnCurFunc > 0 Then
      If mnHistoryFile = SAVEALL Then
         HistoryText$ = GetHistoryText()
         Print #1, HistoryText$
      End If
   End If
   
   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%
   UnLoadChildForm Me, 0, 0

End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash

End Sub

Private Sub mnuQuit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub picCommands_Click()

   frmMain.SetFocus

End Sub

Private Sub picCommands_Resize()

   PicDiv! = picCommands.ScaleWidth / 8
   For i% = 0 To 6
      cmdFormType(i%).Width = PicDiv!
      cmdFormType(i%).Left = PicDiv! * i%
   Next
   cmdUtils.Width = PicDiv!
   cmdUtils.Left = PicDiv! * i%

End Sub

Private Sub Picture1_Click()

   mfmUniTest.Picture1.Visible = False
   mfmUniTest.cmdUtils.ENABLED = True

End Sub

