VERSION 5.00
Begin VB.Form frmGPIBCtl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "GPIB Control"
   ClientHeight    =   1665
   ClientLeft      =   1905
   ClientTop       =   1470
   ClientWidth     =   5520
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1665
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdToggleREN 
      Appearance      =   0  'Flat
      Caption         =   "REN"
      Height          =   315
      Left            =   4800
      TabIndex        =   13
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdRead 
      Appearance      =   0  'Flat
      Caption         =   "&Read"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   1035
   End
   Begin VB.CommandButton cmdConfigure 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdSelDevClear 
      Appearance      =   0  'Flat
      Caption         =   "SDC"
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdDevClear 
      Appearance      =   0  'Flat
      Caption         =   "DCL"
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdTrig 
      Appearance      =   0  'Flat
      Caption         =   "Trigger"
      Height          =   315
      Left            =   2340
      TabIndex        =   9
      Top             =   60
      Width           =   915
   End
   Begin VB.ComboBox cmbArg2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4620
      TabIndex        =   7
      Top             =   420
      Width           =   735
   End
   Begin VB.ComboBox cmbArg1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3780
      TabIndex        =   6
      Top             =   420
      Width           =   735
   End
   Begin VB.CommandButton cmdWrite 
      Appearance      =   0  'Flat
      Caption         =   "&Write"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Tag             =   "GPIBCtl"
      Top             =   420
      Width           =   1035
   End
   Begin VB.TextBox txtCommand 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   780
      Width           =   4155
   End
   Begin VB.ComboBox cmbDevice 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "No Boards Installed"
      Top             =   60
      Width           =   2175
   End
   Begin VB.ComboBox cmbFuncList 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   420
      Width           =   2475
   End
   Begin VB.Label lblCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   1200
      Width           =   4035
   End
End
Attribute VB_Name = "frmGPIBCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnRenState As Integer, mnGPIBCtlrs As Integer
Dim mnUseMCCBoard As Integer, mnMCCCtlrs As Integer

Private Sub cmbArg1_Change()

   DoEvents
   txtCommand = lblCommand & cmbArg1.Text & cmbArg2.Text

End Sub

Private Sub cmbArg1_Click()

   DoEvents
   txtCommand = lblCommand & cmbArg1.Text & cmbArg2.Text

End Sub

Private Sub cmbArg2_Change()

   DoEvents
   txtCommand = lblCommand & cmbArg1.Text & cmbArg2.Text

End Sub

Private Sub cmbArg2_Click()

   DoEvents
   txtCommand = lblCommand & cmbArg1.Text & cmbArg2.Text

End Sub

Private Sub cmbDevice_Change()

   Instance% = Val(Me.Tag)
   DeviceName$ = cmbDevice.Text
   Dummy$ = GetCommands(Instance%, DeviceName$, -1)
   DoEvents
   'cmbDevice.ListIndex = 0
   cmbFuncList.ListIndex = 0
   cmdTrig.ENABLED = Not (DeviceName$ = "ENCODER")

End Sub

Private Sub cmbDevice_Click()

   Device% = cmbDevice.ListIndex
   Instance% = Val(Right$(Me.Tag, 2))
   DeviceName$ = cmbDevice.Text
   Dummy$ = GetCommands(Instance%, DeviceName$, -1)
   'cmbDevice.ListIndex = 0
   If cmbFuncList.ListCount > 0 Then cmbFuncList.ListIndex = 0
   cmdTrig.ENABLED = Not (DeviceName$ = "ENCODER")

End Sub

Private Sub cmbFuncList_Click()

   UpdateCommand

End Sub

Private Sub cmdConfigure_Click()

   'this exists to give menu access to the scripting
   'form when running scripts
   CmdStr$ = Left$(cmdConfigure.Caption, 1)
   value& = Val(Mid$(cmdConfigure.Caption, 2))
   Select Case CmdStr$
      Case "C" 'device clear
         cmdDevClear = True
      Case "D" 'set device
         Device$ = Mid$(cmdConfigure.Caption, 2)
         NotFound% = True
         For SearchIndex% = 0 To cmbDevice.ListCount - 1
            If Not (InStr(cmbDevice.List(SearchIndex%), Device$) = 0) Then
               cmbDevice.ListIndex = SearchIndex%
               NotFound% = False
               Exit For
            End If
         Next SearchIndex%
         If NotFound% Then
            If mnUseMCCBoard Then
               result% = InitControlBoards(DevNum%, Device$)
               MsgBox Device$ & " equivalent not found in list of currently installed devices. " & _
               vbCrLf & "Make sure MCC Control board is installed at designated board number (Board " & _
               Format(DevNum%) & " in Instacal)." & vbCrLf & "Aborting script.", , _
               "Requested GPIB Equivalent Device Not Found"
            Else
               MsgBox Device$ & " not found in list of currently installed devices. " & _
               "If using MCC boards for signal sources, make sure MCC Control board is installed at designated board number. Aborting script.", , "Requested GPIB Device Not Found"
            End If
            gnScriptRun = False
         End If
      Case "L" 'set REN (using ibsre)
         mnRenState = Not value&
         cmdToggleREN = True
      Case "R" 'read device
         cmdRead = True
      Case "S" 'selected device clear
         cmdSelDevClear = True
      Case "T" 'trigger (using GET)
         cmdTrig = True
      Case "W" 'write string in txtCommand
         cmdWrite = True
   End Select

End Sub

Private Sub cmdDevClear_Click()

   If mnUseMCCBoard Then
      Cmd$ = "STOPBG"
      WriteMCC -1, Cmd$
   Else
      DoDevClear -1
   End If

End Sub

Private Sub cmdRead_Click()

   lblResult.Caption = ""
   Device% = cmbDevice.ListIndex
   Cmd$ = Space$(100)
   DisableCommands True
   If mnUseMCCBoard Then
   Else
      ReadGPIB Device%, Cmd$
   End If
   DisableCommands False
   lblResult = Cmd$

End Sub

Private Sub cmdSelDevClear_Click()

   Device% = cmbDevice.ListIndex
   If mnUseMCCBoard Then
      Cmd$ = "STOPBG"
      WriteMCC Device%, Cmd$
   Else
      DoDevClear Device%
   End If

End Sub

Private Sub cmdToggleREN_Click()

   mnRenState = Not mnRenState
   Temp$ = txtCommand.Text
   txtCommand.Text = mnRenState
   DoEvents
   If Not mnUseMCCBoard Then DoIBSre 0, mnRenState
   txtCommand.Text = Temp$

End Sub

Private Sub cmdTrig_Click()

   Device% = cmbDevice.ListIndex
   If mnUseMCCBoard Then
      Cmd$ = "TRIG"
      WriteMCC Device%, Cmd$
   Else
      TriggerGPIB Device%
   End If

End Sub

Private Sub cmdWrite_Click()

   Device% = cmbDevice.ListIndex
   Cmd$ = txtCommand.Text
   If Not mnUseMCCBoard Then
      'check if GPIB device exists at index
      If Not Device% < mnGPIBCtlrs Then
         SigSwitch% = True
         Device% = mnGPIBCtlrs - Device%
      End If
   End If
   DisableCommands True
   If mnUseMCCBoard Or SigSwitch% Then
      WriteMCC Device%, Cmd$
      OrgCmd$ = txtCommand.Text
      If Not OrgCmd$ = Cmd$ Then
         CmdArray = Split(OrgCmd$, " ")
         CmdParts& = UBound(CmdArray)
         For Part& = 0 To CmdParts&
            PartString$ = CmdArray(Part&)
            If IsNumeric(PartString$) Then PartString$ = Cmd$
            NewCmd$ = NewCmd$ & PartString$ & " "
         Next
         NewCmd$ = Trim(NewCmd$)
      End If
   Else
      FuncType$ = Left(Cmd$, 2)
      If FuncType$ = "FU" Then
         FuncVal% = Val(Right(Cmd$, 1))
         If FuncVal% > 5 Then
            FuncVal% = FuncVal% - 5
            Cmd$ = FuncType$ & Format(FuncVal%, "0")
         End If
      End If
      WriteGPIB Device%, Cmd$
   End If
   If Not NewCmd$ = "" Then txtCommand.Text = NewCmd$
   DisableCommands False

End Sub

Private Sub DisableCommands(CommandState As Integer)

   cmdWrite.ENABLED = Not CommandState
   cmdRead.ENABLED = Not CommandState
   
End Sub

Private Sub Form_Load()

   Me.Height = 2150
   Me.Width = 5700
   mnRenState = True
   
   SetInitialValues
   If frmMain.mnuMCCCtl.Checked Then
      'force use of MCC boards as signal sources
      TryMCCBoards% = True
   Else
      If Not InitGPIB() Then
         If gbULLoaded Then TryMCCBoards% = True
      End If
   End If

   If TryMCCBoards% Then
      'Mcc boards forced, no library exists, or no GPIB board installed
      ' - try checking ini file for controller boards
      Device% = 0
      If Not gbULLoaded Then
         MsgBox "MCC Devices cannot be used for control since the Universal Library isn't loaded.", vbCritical, "Universal Library Required"
         gnErrFlag = True
         Exit Sub
      End If
      For CheckAddress% = 0 To 63
         UsingIni% = True
         GetGPIBBoardName CheckAddress%, BdDevName$, UsingIni%
         If Len(BdDevName$) Then
            If InitControlBoards(Device%, BdDevName$) > 0 Then mnUseMCCBoard = True
            If mnUseMCCBoard Then ConfigCtrlBoard cmbDevice.ListIndex '- 1
         End If
      Next
      If Not mnUseMCCBoard Then
         If Not InitGPIB() Then
            MsgBox "There are no GPIB or MCC Control devices configured.", vbCritical, "No Control Interface Configured"
            gnErrFlag = True
            Exit Sub
         End If
      End If
   Else
      For CheckAddress% = 0 To 63
         UsingIni% = True
         GetGPIBBoardName CheckAddress%, BdDevName$, UsingIni%
         If Len(BdDevName$) Then
            x% = InitControlBoards(Device%, BdDevName$, True)
            If x% Then
               ConfigCtrlBoard cmbDevice.ListIndex '- 1
               mnMCCCtlrs = mnMCCCtlrs + 1
            End If
         End If
      Next
   End If

   If gnScriptSave Then
      FormID$ = Me.Tag
      Print #2, FormID$ & ", 5001, " & Format$(GPIB_CTL, "0") & ", 1,,,,,,,,,,,"
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   If gnScriptSave Then
      FormID$ = Me.Tag
      Print #2, FormID$ & ", 5002, " & Format$(GPIB_CTL, "0") & ", 1,,,,,,,,,,,"
   End If
   gnGPIBCtlForms = gnGPIBCtlForms - 1
   cmbDevice.Clear
   DestroyCtrlBoards
   DestroyGPIBBoards

End Sub

Private Sub ParseArg(ArgList As String, Destination As Integer)
   
   If Destination = 1 Then cmbArg1.Clear
   If Destination = 2 Then cmbArg2.Clear
   CmdLen% = Len(ArgList)
   Start% = 1
   Do
      Position% = Position% + 1
      If (Mid$(ArgList, Position%, 1) = "/") Then
         Pos% = Position% - 1
         NewArg$ = Mid$(ArgList, Start%, Position% - Start%)
         If Destination = 1 Then cmbArg1.AddItem NewArg$
         If Destination = 2 Then cmbArg2.AddItem NewArg$
         Position% = Position% + 1
         Start% = Position%
      End If
   Loop While (Position% < CmdLen%)
   NewArg$ = Mid$(ArgList, Start%, CmdLen%)
   If Destination = 1 Then cmbArg1.AddItem NewArg$
   If Destination = 2 Then cmbArg2.AddItem NewArg$
   cmbArg1.ListIndex = 0
   cmbArg2.ListIndex = 0

End Sub

Private Sub ParseCommand(Cmd As String)

   cmbArg1.Clear
   cmbArg2.Clear
   cmbArg1.AddItem ""
   cmbArg2.AddItem ""
   CmdLen% = Len(Cmd)
   If CmdLen% = 0 Then Exit Sub
   Do
      Position% = Position% + 1
      If (Asc(Mid$(Cmd, Position%, 1)) = 0) And Not Parse1% Then
         lblCommand = Left$(Cmd, Position%)
         Position% = Position% + 1
         Start2% = Position%
         Parse1% = True
      End If
      If (Asc(Mid$(Cmd, Position%, 1)) = 0) And Not Parse2% Then
         ArgList1$ = Mid$(Cmd, Start2%, Position% - Start2%)
         Position% = Position% + 1
         Start3% = Position%
         ParseArg ArgList1$, 1
         Parse2% = True
      End If
      If (Asc(Mid$(Cmd, Position%, 1)) = 0) And Not Parse3% Then
         ArgList2$ = Mid$(Cmd, Start3%, Position% - Start3%)
         ParseArg ArgList2$, 2
         Parse3% = True
      End If
   Loop While (Position% < CmdLen%)
   If Not Parse1% Then lblCommand = Cmd
   If Parse1% And Not Parse2% Then
      ArgList1$ = Mid$(Cmd, Start2%, CmdLen%)
      ParseArg ArgList1$, 1
   End If
   If Parse2% And Not Parse3% Then
      ArgList2$ = Mid$(Cmd, Start3%, CmdLen%)
      ParseArg ArgList2$, 2
   End If
      
End Sub

Private Sub UpdateCommand()
   
   ComNum% = cmbFuncList.ListIndex
   DevNum% = cmbDevice.ListIndex
   Instance% = Val(Me.Tag)
   DeviceName$ = cmbDevice.Text
   Comnd$ = GetCommands(Instance%, DeviceName$, ComNum%)
   'Comnd$ = GetCommands(Instance%, DevNum%, ComNum%)
   ParseCommand Comnd$
   txtCommand = lblCommand & cmbArg1.Text & cmbArg2.Text
   DoEvents

End Sub

Public Function Get488ValueRead() As String

   Get488ValueRead = lblResult.Caption
   
End Function

Public Function GetReturnVal() As String

   GetReturnVal = Me.txtCommand.Text
   
End Function

Public Sub SetNumGPIBCtlrs(ByVal NumFound As Integer)

   mnGPIBCtlrs = NumFound
   
End Sub
