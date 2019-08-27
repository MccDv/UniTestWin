VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menus"
   ClientHeight    =   750
   ClientLeft      =   1800
   ClientTop       =   1815
   ClientWidth     =   7365
   ControlBox      =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   750
   ScaleWidth      =   7365
   Tag             =   "Main"
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuGPIB 
         Caption         =   "&GPIB Interface"
      End
      Begin VB.Menu mnuMCCCtl 
         Caption         =   "Force MCC Control"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close A&ll"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogger 
         Caption         =   "Logger Functions"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "Universal Library"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   ".Net Library"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "DAQFlex"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoIcal 
         Caption         =   "Ignore InstaCal"
      End
      Begin VB.Menu mnuRefreshDevs 
         Caption         =   "Refresh Devices"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuManRefreshDevs 
         Caption         =   "Manual Refresh Devices"
      End
      Begin VB.Menu mnuManageDevs 
         Caption         =   "Manage Devices..."
      End
      Begin VB.Menu mnuSepScript 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScripter 
         Caption         =   "&Scripting"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFuncs 
      Caption         =   "F&unctions"
      Begin VB.Menu mnuGetRev 
         Caption         =   "cbGetRevision()"
      End
      Begin VB.Menu mnuDeclRev 
         Caption         =   "cbDeclareRevision()"
      End
      Begin VB.Menu mnuLoadConf 
         Caption         =   "cbLoadConfig()"
      End
      Begin VB.Menu mnuSaveConfig 
         Caption         =   "cbSaveConfig()"
      End
      Begin VB.Menu mnuErrHandling 
         Caption         =   "&Error Handling"
      End
      Begin VB.Menu mnuMsgSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevName 
         Caption         =   "Device Name Format"
         Begin VB.Menu mnuDevNameSpec 
            Caption         =   "Name Only"
            Index           =   0
         End
         Begin VB.Menu mnuDevNameSpec 
            Caption         =   "Name - Serial Number"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuDevNameSpec 
            Caption         =   "Name - ID"
            Index           =   2
         End
         Begin VB.Menu mnuDevNameSpec 
            Caption         =   "Name - Serial Number - ID"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&*Information"
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu mnuSelFileProps 
         Caption         =   "File Properties"
         Begin VB.Menu mnuFileProps 
            Caption         =   "cbw.dll"
            Index           =   0
         End
         Begin VB.Menu mnuFileProps 
            Caption         =   "cbw16.dll"
            Index           =   1
         End
         Begin VB.Menu mnuFileProps 
            Caption         =   "cbw32.dll"
            Index           =   2
         End
         Begin VB.Menu mnuFileProps 
            Caption         =   "cbul.386"
            Index           =   4
         End
         Begin VB.Menu mnuFileProps 
            Caption         =   "Other"
            Index           =   5
         End
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegSel 
         Caption         =   "Registry"
         Begin VB.Menu mnuReg 
            Caption         =   "CBDIREC"
            Index           =   1
         End
         Begin VB.Menu mnuReg 
            Caption         =   "DAS Components"
            Index           =   2
         End
         Begin VB.Menu mnuReg 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuReg 
            Caption         =   ""
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuArrange 
         Caption         =   "Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "Tile Horizontally"
         Index           =   1
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "Tile Vertically"
         Index           =   2
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPosition 
         Caption         =   "&Position"
         Begin VB.Menu mnuPosSel 
            Caption         =   "&Top of Screen"
            Index           =   0
         End
         Begin VB.Menu mnuPosSel 
            Caption         =   "&Bottom of Screen"
            Index           =   1
         End
         Begin VB.Menu mnuPosSel 
            Caption         =   "&Width of Screen"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnLibType As Integer
Dim mfNoForm As Form
Dim mnCreateDevs As Integer, mnRemoveDevs As Integer

Private Sub ConfigureControls()

   If mnLibType = MSGLIB Then
      mnuGetRev.ENABLED = False
      mnuDeclRev.ENABLED = False
      mnuLoadConf.ENABLED = False
   Else
      mnuGetRev.ENABLED = True
      mnuDeclRev.ENABLED = True
      mnuLoadConf.ENABLED = True
   End If
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   If Shift And vbAltMask Then
      Select Case KeyCode
         Case Asc("I")
            mfmUniTest.cmdFormType(0) = True
         Case Asc("O")
            mfmUniTest.cmdFormType(1) = True
         Case Asc("N")
            mfmUniTest.cmdFormType(2) = True
         Case Asc("T")
            mfmUniTest.cmdFormType(3) = True
         Case Asc("C")
            mfmUniTest.cmdFormType(4) = True
         Case Asc("M")
            mfmUniTest.cmdFormType(5) = True
         Case Asc("'")
            mfmUniTest.cmdFormType(6) = True
         Case Asc("L")
            mfmUniTest.cmdUtils = True
      End Select
   ElseIf Shift And vbCtrlMask Then
      Select Case KeyCode
         Case Asc("X")  'Ctl X
            'program run by SendKeys
            gnXternalCtl = True
      End Select
   Else
      Select Case KeyCode
         Case vbKeyF9 'F9 - set height to 1/3 screen
            mfmUniTest.Height = Screen.Height / 3
         Case vbKeyF11 'F11 - set to screen bottom
            mfmUniTest.Move 0, Screen.Height - mfmUniTest.Height, Screen.Width
         Case vbKeyF12 'F12 - set to screen top
            mfmUniTest.Move 0, 0, Screen.Width
      End Select
   End If

End Sub

Private Sub Form_Load()
    
#If NETOPS Then
   Me.mnuLibrary(NETLIB).ENABLED = True
#End If
#If Not MSGOPS Then
   Me.mnuLibrary(MSGLIB).ENABLED = False
#End If
   
   If Not gbULLoaded Then
      mnuLibrary(UNILIB).Checked = False
      mnuLibrary(UNILIB).ENABLED = False
   End If
   'frmMain.mnuLogger.Enabled = gbULLoaded
   mnuLibrary(MSGLIB).Checked = (gnLibType = MSGLIB)
   mnuLibrary(NETLIB).Checked = (gnLibType = NETLIB)
   mnuLibrary(UNILIB).Checked = (gnLibType = UNILIB)
   mnuNoIcal.Checked = gnICalDisable
   mnLibType = gnLibType

   Dim CurStyle As Long
   Dim NewStyle As Long
   
   mnuScripter.ENABLED = mfmUniTest.picCommands.ENABLED
   CurStyle = GetWindowLong(frmMain.hWnd, GWL_STYLE)
   NewStyle = SetWindowLong(frmMain.hWnd, GWL_STYLE, CurStyle And Not (WS_CAPTION))
   DoEvents
   PositionMain
   If gnLibVer = WIN95NT Then
      Load mnuFileProps(6)
      Load mnuFileProps(7)
      Load mnuFileProps(8)
      Load mnuFileProps(9)
      mnuFileProps(1).Caption = "Cbrtt32.dll"
      mnuFileProps(5).Caption = "cbi_cal.dll"
      mnuFileProps(6).Caption = "cbi_node.dll"
      mnuFileProps(7).Caption = "cbi_prop.dll"
      mnuFileProps(8).Caption = "cbi_test.dll"
      mnuFileProps(9).Caption = "Other"
   End If

   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "ControlConfig"
   lpKeyName$ = "ForceMCC"
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = "0"
   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   CtlConfig$ = Left$(lpReturnedString$, x%)
   If Len(CtlConfig$) Then mnuMCCCtl.Checked = Val(CtlConfig$)
   If gnUniScript Then
      Me.HelpContextID = 30000
   Else
      Me.HelpContextID = 10000
   End If
   
   NameSpec% = GetNameFormat()
   mnuDevNameSpec_Click (NameSpec%)
   ConfigureControls

   mnRemoveDevs = 1
   mnCreateDevs = 1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "ControlConfig"
   lpKeyName$ = "ForceMCC"
   lpString$ = "0"
   If mnuMCCCtl.Checked Then lpString$ = "-1"
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "MainForm"
   lpKeyName$ = "Library"
   Select Case gnLibType
      Case INVALIDLIB
         lpString$ = "None"
      Case UNILIB
         lpString$ = "Unilib"
      Case NETLIB
         lpString$ = "DotNet"
      Case MSGLIB
         lpString$ = "MsgDaq"
   End Select
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)

   lpKeyName$ = "InstaCal"
   nSize% = 16
   lpReturnedString$ = Space$(nSize%)
   lpString$ = "Enabled"
   If gnICalDisable Then lpString$ = "Disabled"

   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)


End Sub

Private Sub Form_LostFocus()

   mfmUniTest.Picture1.Visible = False
   mfmUniTest.cmdUtils.ENABLED = True

End Sub

Private Sub Form_Resize()

   PositionMain

End Sub

Private Sub mnuAbout_Click()

   frmSplash.Show 1
   Unload frmSplash
   
End Sub

Private Sub mnuArrange_Click(Index As Integer)

   mfmUniTest.Picture1.Visible = False
   mfmUniTest.Arrange Index
   PositionMain

End Sub

Private Sub mnuClose_Click()

   mfmUniTest.Picture1.Visible = False
   Unload Me

End Sub

Private Sub mnuCloseAll_Click()

   If Forms.Count > 2 Then
      For i% = Forms.Count - 1 To 0 Step -1
         If Not ((Forms(i%).Tag = "Main") Or (Forms(i%).Tag _
            = "UniTest")) Then Unload Forms(i%)
      Next i%
   End If

End Sub

Private Sub mnuConfig_Click()

   mfmUniTest.Picture1.Cls
   mfmUniTest.Picture1.Visible = True
   DoEvents
   
   If gnLibVer = WIN95NT Then
      DLLName$ = "cbw32.dll"
   Else
      DLLName$ = "cbw.dll"
   End If
   ColWidth = mfmUniTest.Picture1.ScaleWidth / 4
   OperatingSystem$ = GetSysInfo(OPSYS, 0)
   mfmUniTest.Picture1.Print "Library revision: " & Val(CURRENTREVNUM)
   mfmUniTest.Picture1.Print "Operating system: " & OperatingSystem$
   DLLPath$ = GetSysInfo(DLLINF, DLLName$)
   mfmUniTest.Picture1.Print "Universal Library DLL path: " & DLLPath$;
   If Not (DLLPath$ = "Not Found") Then
      mfmUniTest.Picture1.CurrentX = ColWidth * 2
      mfmUniTest.Picture1.Print "DLL date: " & GetSysInfo(DLLINF + FILEDATE, DLLName$);
      mfmUniTest.Picture1.CurrentX = ColWidth * 3
      ReDim Props$(1 To 7)
      GetFileVersion DLLPath$ & DLLName$, Props$
      mfmUniTest.Picture1.Print "DLL revision: " & Props$(6) 'GetSysInfo(DLLINF + FILEREV, DLLName$)
      ULStat = cbGetRevision(DLLRev!, VXDRev!)
      If SaveFunc(Me, GetRevision, ULStat, DLLRev!, VXDRev!, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      mfmUniTest.Picture1.CurrentX = ColWidth * 2
      mfmUniTest.Picture1.Print "cbGetRevision()"; 'Chr$(9) &
      mfmUniTest.Picture1.CurrentX = ColWidth * 3
      mfmUniTest.Picture1.Print "DLL: " & Format$(DLLRev!, "0.00")
   Else
      mfmUniTest.Picture1.Print
   End If
   If OperatingSystem$ = "Win NT " Then
      mnuRegSel.ENABLED = True
      KeyName$ = "SYSTEM\CurrentControlSet\Services\CBUL32\ImagePath"
      RegInfo = GetSysInfo(REGINF + HKLM, KeyName$)
      mfmUniTest.Picture1.Print "UL32 NT Driver path: " & RegInfo;
      DriverName$ = "cbul32.sys"
   Else
      'DriverName$ = "cbul.386"
      'VxDPath$ = GetSysInfo(VXDINF, DriverName$)
   End If
   'mfmUniTest.Picture1.Print "Universal Library VxD path: " & VxDPath$
   'If Not (VxDPath$ = "Not Found") Then
   '   mfmUniTest.Picture1.CurrentX = ColWidth * 2
   '   mfmUniTest.Picture1.Print "VxD date: " & GetSysInfo(VXDINF + FILEDATE, DriverName$);
   '   mfmUniTest.Picture1.CurrentX = ColWidth * 3
   '   mfmUniTest.Picture1.Print "VxD revision: " & GetSysInfo(VXDINF + FILEREV, DriverName$)
   '   mfmUniTest.Picture1.CurrentX = ColWidth * 2
   '   mfmUniTest.Picture1.Print "cbGetRevision()"; 'Chr$(9) &
   '   mfmUniTest.Picture1.CurrentX = ColWidth * 3
   '   mfmUniTest.Picture1.Print "VxD: " & Format$(VXDRev!, "0.00")
   'End If
   mfmUniTest.Picture1.Print
   If Not (DLLPath$ = "Not Found") Then
      ReDim BoardEnum(0) As Integer
      NumBoards% = GetNumInstalled()
      mfmUniTest.Picture1.Print "Number of boards installed: " & NumBoards%
      For i% = 0 To NumBoards% - 1
         BoardNum% = gnBoardEnum(i%)
         mfmUniTest.Picture1.Print Chr$(9) & "Board " & BoardNum% & ")  " & GetNameOfBoard(BoardNum%)
      Next i%
   End If

End Sub

Private Sub mnuDeclRev_Click()

   mfmUniTest.Picture1.Visible = False
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "MainForm"
   lpKeyName$ = "Revision"
   nSize% = 6
   lpReturnedString$ = Space$(nSize%)
   lpDefault$ = Format$(CURRENTREVNUM, "0.00")

   x% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lpFileName$)
   Rev$ = Left$(lpReturnedString$, x%)
   sDefault$ = Rev$
   Revision = InputBox("Enter revision", "Declare Revision", sDefault$)
   If Len(Revision) Then
      RevNum! = Val(Revision)
      ULStat = cbDeclareRevision(RevNum!)
      If SaveFunc(Me, DeclareRevision, ULStat, RevNum!, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
      lpString$ = Format$(RevNum!, "0.00")
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
   End If

End Sub

Private Sub mnuDevNameSpec_Click(Index As Integer)

   If mnuDevNameSpec(Index).Checked Then Exit Sub
   For MenuIndex% = 0 To mnuDevNameSpec.Count - 1
      mnuDevNameSpec(MenuIndex%).Checked = False
   Next
   mnuDevNameSpec(Index).Checked = True
   NameFormat% = Choose(Index + 1, NAME_ONLY, NAME_SERNO, NAME_ID, NAME_SERNO_ID)
   SetNameFormat Index
   
End Sub

Private Sub mnuErrHandling_Click()

   mfmUniTest.Picture1.Visible = False
   ErrorHandling True

End Sub

Private Sub mnuExit_Click()

   For i% = Forms.Count - 1 To 0 Step -1
      Unload Forms(i%)
   Next i%

End Sub

Private Sub mnuFileProps_Click(Index As Integer)
   
   ReDim Props$(1 To 7)
   'mfmUniTest.Picture1.Visible = True
   'mfmUniTest.Image1.Visible = False
   'mfmUniTest.Picture1.Cls
   LastMenu% = 9
   If CURRENTREVNUM < 5 Then LastMenu% = 5
   DoEvents
   If Index < LastMenu% Then
      Filename$ = mnuFileProps(Index).Caption
      If LCase$(Right$(Filename$, 4)) = ".dll" Then FullPath$ = GetSysInfo(DLLINF, Filename$)
      If LCase$(Right$(Filename$, 4)) = ".386" Then FullPath$ = GetSysInfo(VXDINF, Filename$)
      Filename$ = FullPath$ & Filename$
   Else
      Default$ = App.EXEName
      Filename$ = InputBox$("Enter file name", "File Properties", Default$)
      frmMain.SetFocus
   End If
   GetFileVersion Filename$, Props$()
   
   DoEvents
   'mfmUniTest.Picture1.Visible = True
   If mnuFileProps(Index).Caption = "cbw32.dll" Then gfDLLRev = Val(Props$(6))
   PrintMain "Properties for " & UCase$(Filename$) & ":  " & Props$(6)
   'mfmUniTest.Picture1.Print
   If 0 Then
   For i% = 1 To 7
      If Len(Props$(i%)) > 100 Then
         Position% = 80
         Do
            ThisChar$ = Mid$(Props$(i%), Position%, 1)
            If ThisChar$ = " " Then Exit Do
            Position% = Position% + 1
         Loop While Position% < Len(Props$(i%))
         mfmUniTest.Picture1.Print Left$(Props$(i%), Position%)
         mfmUniTest.Picture1.Print Mid$(Props$(i%), Position% + 1)
      Else
         mfmUniTest.Picture1.Print Props$(i%)
      End If
      If (i% = 3) Or (i% = 6) Then mfmUniTest.Picture1.Print
   Next i%
   End If

End Sub

Private Sub mnuGetRev_Click()

   mfmUniTest.Picture1.Visible = False
   Me.Cls
   ULStat = cbGetRevision(DLLRevNum!, VXDRevNum!)
   PrintMain "DLL Revision: " & DLLRevNum! & "   VxD Revision: " & VXDRevNum!

End Sub

Private Sub mnuGPIB_Click()

   mfmUniTest.Picture1.Visible = False
   Success% = OpenGPIB()
   If Not Success% Then CloseGPIB 0
   DoEvents

End Sub

Private Sub mnuLibrary_Click(Index As Integer)

   For LibMenu% = 0 To mnuLibrary.Count - 1
      mnuLibrary(LibMenu%).Checked = False
   Next
   
   mnuLibrary(Index).Checked = True
   mnLibType = Index
   ConfigureLibrary mnLibType, True
   If gnLibType = INVALIDLIB Then
      If Not (mnLibType = gnLibType) Then
         mnuLibrary(mnLibType).Checked = False
         mnuLibrary(mnLibType).ENABLED = False
         mnLibType = gnLibType
         ConfigureLibrary gnLibType, True
      End If
   Else
      gnLibType = mnLibType
   End If
   ConfigureControls

End Sub

Private Sub mnuManageDevs_Click()

   If LibSupportsFunction(GetDaqDeviceInventory) Then
      frmDiscovery.chkCreate.value = mnCreateDevs
      frmDiscovery.chkRemoveUndisc.value = mnRemoveDevs
      frmDiscovery.Show 1
      mnCreateDevs = frmDiscovery.chkCreate.value
      mnRemoveDevs = frmDiscovery.chkRemoveUndisc.value
      Unload frmDiscovery
   End If
   
End Sub

Private Sub mnuMCCCtl_Click()

   mnuMCCCtl.Checked = Not mnuMCCCtl.Checked
   
End Sub

Private Sub mnuLoadConf_Click()

   Dim Resp As VbMsgBoxResult
   If Forms.Count > 2 Then
      Resp = MsgBox("All open child forms will be closed " & _
         "before executing cbLoadConfig(). Unload child forms?", _
         4, "Unload Forms?")
      If Resp = vbNo Then Exit Sub
      For i% = Forms.Count - 1 To 0 Step -1
         If Not ((Forms(i%).Tag = "Main") Or (Forms(i%).Tag = "UniTest")) Then Unload Forms(i%)
      Next i%
   End If
   mfmUniTest.Picture1.Visible = False
   Filename$ = "cb.cfg"
   UserFileName$ = InputBox("Configuration file name", "File Name", Filename$)
   ULStat = GetCfgFile(UserFileName$)
   gnNumBoards = GetNumInstalled()
   'the value returned by cbLoadConfig() is
   'not an error code

   NumFuncs% = GetHistory() - 1
   ReDim MyArray(NumFuncs%, 14)
   GetHistoryArray MyArray()
   PrintMain "cbLoadConfig() = " & ULStat

End Sub

Private Sub mnuLogger_Click()

   'frmLogger.Show
   LoadChildForm LOGFUNC

End Sub

Private Sub mnuNoIcal_Click()

   If LibSupportsFunction(IgnoreInstaCal) Then
      mnuNoIcal.Checked = Not mnuNoIcal.Checked
      gnICalDisable = mnuNoIcal.Checked
      If gnICalDisable Then
         MsgBox "Restart Universal Test to detect devices " & _
            "within this application and ignore Instacal settings.", _
            vbInformation, "Restart Universal Test"
      Else
         MsgBox "Restart Universal Test to use devices " & _
            "configured through Instacal.", _
            vbInformation, "Restart Universal Test"
      End If
   End If
   
End Sub

Private Sub mnuPosSel_Click(Index As Integer)
   
   Select Case Index
      Case 0
         mfmUniTest.Top = 0
      Case 1
         mfmUniTest.Top = Screen.Height - mfmUniTest.Height
      Case 2
         mfmUniTest.Width = Screen.Width
         mfmUniTest.Left = 0
   End Select

End Sub

Private Sub mnuRefreshDevs_Click()

   If LibSupportsFunction(GetDaqDeviceInventory) Then
      mfmUniTest.MousePointer = vbHourglass
      DoEvents
      x& = DiscoverDevices(ANY_IFC, True)
      mfmUniTest.MousePointer = vbDefault
      Unload frmDiscovery
   End If

End Sub

Private Sub mnuManRefreshDevs_Click()

   Dim HostString As String
   Dim HostPort As Long, Timeout As Long
   
   If LibSupportsFunction(GetDaqDeviceInventory) Then
      Timeout = 5000
      HostString = InputBox("Enter host name or IP address:", _
         "Add Remote Device", "173.76.198.250")
      HP = InputBox("Enter host port:", "Add Remote Device", "54211")
      HostPort = Val(HP)
      
      If (HostString = "") Or (HP = "") Then Exit Sub
      x& = DiscoverDevices(ETHERNET_IFC, True, HostString, HostPort, Timeout)
   End If
   
End Sub

Private Sub mnuReg_Click(Index As Integer)

   Select Case Index
      Case 1
         KeyName$ = "Environment\CBDIREC"
         RegHKey& = HKCU
      Case 2
         KeyName$ = "SYSTEM\CurrentControlSet\Services\Class\DAS Component\"
         RegHKey& = HKLM
      Case 3
      Case 4
   End Select
   RegInfo = GetSysInfo(REGINF + RegHKey&, KeyName$)

End Sub

Private Sub mnuSaveConfig_Click()

   If False Then
      'obsolete start
      If Forms.Count > 2 Then
         Resp = MsgBox("All open child forms will be closed " & _
            "before executing cbSaveConfig(). Unload child forms?", _
            4, "Unload Forms?")
         If Resp = 7 Then Exit Sub
         For i% = Forms.Count - 1 To 0 Step -1
            If Not ((Forms(i%).Tag = "Main") Or (Forms(i%).Tag _
               = "UniTest")) Then Unload Forms(i%)
         Next i%
      End If
      'obsolete end
   End If
   mfmUniTest.Picture1.Visible = False
   Filename$ = "cb.cfg"
   UserFileName$ = InputBox("Configuration file name", "File Name", Filename$)
   ULStat& = cbSaveConfig(UserFileName$)
   If SaveFunc(Me, SaveConfig, ULStat&, UserFileName$, A2, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   'the value returned by cbLoadConfig() is
   'not an error code

   NumFuncs% = GetHistory() - 1
   ReDim MyArray(NumFuncs%, 14)
   GetHistoryArray MyArray()
   PrintMain "cbSaveConfig() = " & ULStat

End Sub

Private Sub mnuScripter_Click()

   mnuScripter.Checked = Not mnuScripter.Checked
   If mnuScripter.Checked Then
      CurGroupKey$ = "SOFTWARE\Measurement Computing\Universal Test Suite"
      KeyName$ = "ScriptPath"
      ProgExists% = GetRegGroup(HKEY_LOCAL_MACHINE, CurGroupKey$, hProgResult&)
      ScriptRegistered% = GetKeyValue(hProgResult&, KeyName$, KeyVal$)
      If ScriptRegistered% Then
         ChDrive (Left(KeyVal$, 2))
         ChDir (KeyVal$)
      End If
      frmScript.Show
   Else
      mfmUniTest.Show
      mfmUniTest.cmdUtils.Caption = "Uti&lities"
      mfmUniTest.picCommands.Visible = True
      Unload frmScript
      gnUniScript = False
   End If

End Sub

Public Sub SetDefaultLibType(ByVal LibType As Integer)

   If Not (gnLibType = LibType) Then
      Select Case LibType
         Case UNILIB
            If gbULLoaded Then
               gnNumBoards = GetNumInstalled()
            Else
               Exit Sub
            End If
         Case MSGLIB
            gnNumBoards = GetNumMsgBoards()
      End Select
   End If
   mnuLibrary_Click (LibType)

End Sub

Public Sub SetMCCControl(ByVal TrueFalse As Boolean)

   mnuMCCCtl.Checked = TrueFalse
   If mnuNoIcal.Checked = True Then
      mnuNoIcal_Click
      'ConfigureLibrary mnLibType, True
      'ConfigureControls
      gnErrFlag = True
   End If

End Sub
