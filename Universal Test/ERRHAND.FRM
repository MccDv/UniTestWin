VERSION 5.00
Begin VB.Form frmErrors 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Error Handling Setup"
   ClientHeight    =   2565
   ClientLeft      =   360
   ClientTop       =   1740
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "ERRHAND.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmErrors"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2565
   ScaleWidth      =   9195
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   9015
   End
   Begin VB.Frame fraBufferSetup 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Setup Function History Buffer"
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CommandButton cmdDisableAll 
         Appearance      =   0  'Flat
         Caption         =   "Di&sable All"
         Height          =   435
         Left            =   5640
         TabIndex        =   10
         Top             =   1440
         Width           =   1275
      End
      Begin VB.CommandButton cmdEnableAll 
         Appearance      =   0  'Flat
         Caption         =   "&Enable All"
         Height          =   435
         Left            =   4140
         TabIndex        =   9
         Top             =   1440
         Width           =   1275
      End
      Begin VB.PictureBox picSelFunctions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   180
         ScaleHeight     =   -4
         ScaleMode       =   0  'User
         ScaleTop        =   4
         ScaleWidth      =   23
         TabIndex        =   8
         Top             =   240
         Width           =   8775
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   21
            X1              =   22.001
            X2              =   22.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   20
            X1              =   21.001
            X2              =   21.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   19
            X1              =   19.999
            X2              =   19.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   18
            X1              =   19
            X2              =   19
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   17
            X1              =   18
            X2              =   18
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   16
            X1              =   17.001
            X2              =   17.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   15
            X1              =   15.999
            X2              =   15.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   14
            X1              =   14.999
            X2              =   14.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   13
            X1              =   14
            X2              =   14
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   12
            X1              =   13
            X2              =   13
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   11
            X1              =   12.001
            X2              =   12.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   10
            X1              =   10.999
            X2              =   10.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   9
            X1              =   10
            X2              =   10
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   8
            X1              =   9
            X2              =   9
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   7
            X1              =   8.001
            X2              =   8.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   6
            X1              =   7.001
            X2              =   7.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   5
            X1              =   5.999
            X2              =   5.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   4
            X1              =   5
            X2              =   5
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   3
            X1              =   4
            X2              =   4
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   2
            X1              =   3.001
            X2              =   3.001
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   1.999
            X2              =   1.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   0.999
            X2              =   0.999
            Y1              =   4
            Y2              =   0
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   0
            X2              =   22.976
            Y1              =   3.001
            Y2              =   3.001
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   0
            X2              =   22.976
            Y1              =   2.006
            Y2              =   2.006
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   0
            X2              =   22.976
            Y1              =   0.999
            Y2              =   0.999
         End
      End
      Begin VB.CommandButton cmdDoneBuffer 
         Appearance      =   0  'Flat
         Caption         =   "&Done"
         Height          =   435
         Left            =   7980
         TabIndex        =   6
         Top             =   1440
         Width           =   915
      End
   End
   Begin VB.Frame fraFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Format"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   7800
      TabIndex        =   2
      Top             =   60
      Width           =   1335
      Begin VB.OptionButton optFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hex"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Decimal"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txtErrorMessage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1995
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   7635
   End
   Begin VB.CommandButton cmdDone 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   8040
      TabIndex        =   0
      Top             =   1620
      Width           =   915
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "C&lear Buffer"
      Height          =   435
      Left            =   7800
      TabIndex        =   7
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Menu mnuSelLibHandling 
      Caption         =   "Li&brary Error Handling"
      Begin VB.Menu mnuSelPrint 
         Caption         =   "Error &Printing"
         Begin VB.Menu mnuPrint 
            Caption         =   "DONTPRINT"
            Index           =   0
            Shortcut        =   ^X
         End
         Begin VB.Menu mnuPrint 
            Caption         =   "PRINTWARNINGS"
            Index           =   1
         End
         Begin VB.Menu mnuPrint 
            Caption         =   "PRINTFATAL"
            Index           =   2
         End
         Begin VB.Menu mnuPrint 
            Caption         =   "PRINTALL"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSelLibFlow 
         Caption         =   "Program &Flow"
         Begin VB.Menu mnuFlow 
            Caption         =   "DONTSTOP"
            Index           =   0
         End
         Begin VB.Menu mnuFlow 
            Caption         =   "STOPFATAL"
            Index           =   1
         End
         Begin VB.Menu mnuFlow 
            Caption         =   "STOPALL"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuSelLocHandling 
      Caption         =   "Lo&cal Error Handling"
      Begin VB.Menu mnuLocHandling 
         Caption         =   "None"
         Index           =   0
      End
      Begin VB.Menu mnuLocHandling 
         Caption         =   "Abort Routine"
         Index           =   1
      End
      Begin VB.Menu mnuLocHandling 
         Caption         =   "Stop On Error"
         Index           =   2
      End
      Begin VB.Menu mnuLocHandling 
         Caption         =   "Continue After Error"
         Index           =   3
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocPrint 
         Caption         =   "Print Local Errors"
      End
      Begin VB.Menu mnuLiveDisplay 
         Caption         =   "Live Display"
      End
   End
   Begin VB.Menu mnuSelWarnHandling 
      Caption         =   "Local Warning Handling"
      Begin VB.Menu mnuLocWarnHandling 
         Caption         =   "None"
         Index           =   0
      End
      Begin VB.Menu mnuLocWarnHandling 
         Caption         =   "Abort Routine"
         Index           =   1
      End
      Begin VB.Menu mnuLocWarnHandling 
         Caption         =   "Stop On Error"
         Index           =   2
      End
      Begin VB.Menu mnuLocWarnHandling 
         Caption         =   "Continue After Error"
         Index           =   3
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocWarnPrint 
         Caption         =   "Print Local Warnings"
      End
   End
   Begin VB.Menu mnuHist 
      Caption         =   "&History Buffer"
   End
   Begin VB.Menu mnuSelHistOpts 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuBufSize 
         Caption         =   "History Buffer Si&ze"
      End
      Begin VB.Menu mnuSaveBufToFile 
         Caption         =   "&Save To File"
         Begin VB.Menu mnuFileOpt 
            Caption         =   "&Don't Save"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuFileOpt 
            Caption         =   "Save &Errors"
            Index           =   1
         End
         Begin VB.Menu mnuFileOpt 
            Caption         =   "Save &All"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim manFuncList() As Integer
Dim mnNumFuncs As Integer, mnFileOpt As Integer
Dim mnMaxFuncs As Integer
Dim mnErrReporting As Integer, mnErrHandling As Integer
Dim meErrFlow As Integer, mnErrDisp As Integer
Dim meWarnFlow As Integer, mnWarnDisp As Integer
Dim mnNoError As Integer

Private Sub cmdClear_Click()

   ClearBuffer
   ErrorHandling True

End Sub

Private Sub cmdDisableAll_Click()

   For i% = 1 To UBound(manFuncList)
      manFuncList(i%) = True
   Next
   UPDATEGRID

End Sub

Private Sub cmdDone_Click()

   ErrChange% = (gnErrReporting <> mnErrReporting) Or (gnErrHandling <> mnErrHandling) Or _
   (geErrFlow <> meErrFlow) Or (mnErrDisp <> gnLocalErrDisp) Or _
   (geWarnFlow <> meWarnFlow) Or (mnWarnDisp <> gnLocalWarnDisp)
   If ErrChange% Then
      gnErrReporting = mnErrReporting
      gnErrHandling = mnErrHandling
      geErrFlow = meErrFlow
      gnLocalErrDisp = mnErrDisp
      gnLocalWarnDisp = mnWarnDisp
      geWarnFlow = meWarnFlow
      
      ErrorDisp$ = "0": If gnLocalErrDisp Then ErrorDisp$ = "1"
      lpFileName$ = "UniTest.ini"
      lpApplicationName$ = "MainForm"
      lpKeyName$ = "ErrorSettings"
      lpString$ = gnErrReporting & gnErrHandling & geErrFlow & ErrorDisp$ & geWarnFlow & gnLocalWarnDisp
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      If gnScriptSave Then
         A1 = geErrFlow
         Print #2, Me.Tag & ", " & Format$(SErrorFlow, "0") & ", 0, "; A1; ","; A2; ","; A3; ","; A4; ","; A5; ","; A6; ","; A7; ","; A8; ","; A9; ","; A10; ","; A11; ","; AuxHandle
      End If
   End If
   Unload frmErrors

End Sub

Private Sub cmdDoneBuffer_Click()

   fraBufferSetup.Visible = False
   txtErrorMessage.Visible = True
   fraFormat.Visible = True
   cmdDone.Visible = True
   cmdClear.Visible = True
   If mnNumFuncs < gnCurFunc Then gnCurFunc = 0
   SetFuncsSaved manFuncList()
   NumberOfFunctions% = mnNumFuncs
   If mnNoError Then NumberOfFunctions% = -1
   SetHistory NumberOfFunctions%, mnFileOpt
   mnuHist.Visible = True
   mnuSelLocHandling.Visible = True
   mnuSelLibHandling.Visible = True
   mnuSelHistOpts.Visible = False

End Sub

Private Sub cmdEnableAll_Click()

   For i% = 1 To UBound(manFuncList)
      manFuncList(i%) = False
   Next
   UPDATEGRID

End Sub

Private Sub Form_Activate()

   mnuFileOpt(0).Checked = False
   mnuFileOpt(mnFileOpt).Checked = True
   
End Sub

Private Sub Form_Load()

   lpFileName$ = "UniTest.ini"
   lpHKeyName$ = "Height"
   nHDefault% = 3165
   lpWKeyName$ = "Width"
   nWDefault% = 9315
   lpTKeyName$ = "Top"
   nTDefault% = Screen.Height / 2 - frmErrors.Height / 2
   lpLKeyName$ = "Left"
   nLDefault% = Screen.Width / 2 - frmErrors.Width / 2
   
   FormName$ = "ErrorForm"
   lpApplicationName$ = FormName$

   Me.Height = 1
   Me.Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
   Me.Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
   Me.Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
   Me.Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
   
   mnErrReporting = gnErrReporting
   mnErrHandling = gnErrHandling
   meErrFlow = geErrFlow
   mnErrDisp = gnLocalErrDisp
   meWarnFlow = geWarnFlow
   mnWarnDisp = gnLocalWarnDisp
   
   mnuLocHandling(geErrFlow).Checked = True
   mnuPrint(gnErrReporting).Checked = True
   mnuFlow(gnErrHandling).Checked = True
   mnuLocPrint.Checked = gnLocalErrDisp
   mnuLiveDisplay.Checked = gnLiveDisp
   mnuLocWarnHandling(geWarnFlow).Checked = True
   mnuLocWarnPrint.Checked = mnWarnDisp
   mnMaxFuncs = GetMaxFuncs()
   ReDim manFuncList(1 To mnMaxFuncs)
   UpdateStatus
   optFormat(1).value = gnHexVals

End Sub

Private Sub Form_Resize()

   If Not (frmErrors.WindowState = 1) Then
      'lblStatus.Top = frmErrors.ScaleHeight - 465
      txtStatus.Top = frmErrors.ScaleHeight - 465
      If frmErrors.Height < 3000 Then frmErrors.Height = 3000
      txtErrorMessage.Height = frmErrors.ScaleHeight - 570
      frmErrors.Width = 9500
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

   If Me.WindowState = 0 Then
      lpFileName$ = "UniTest.ini"
      FormName$ = "ErrorForm"
      lpApplicationName$ = FormName$
      lpKeyName$ = "Height"
      lpString$ = Str$(Me.Height)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      lpKeyName$ = "Width"
      lpString$ = Str$(Me.Width)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      lpKeyName$ = "Top"
      lpString$ = Str$(Me.Top)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
      lpKeyName$ = "Left"
      lpString$ = Str$(Me.Left)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lpFileName$)
   End If
   mnNoError = False

End Sub

Private Sub mnuBufSize_Click()

   Default = mnNumFuncs
   NumStored = InputBox("Enter number of functions to store in circular buffer", "Set Buffer Size", Default)
   If Len(NumStored) > 0 Then
      NumFuncs% = Val(NumStored)
      If Not NumFuncs% = mnNumFuncs Then
         mnNoError = False
         mnNumFuncs = NumFuncs%
      End If
   End If
   
End Sub

Private Sub mnuFileOpt_Click(Index As Integer)

   mnuFileOpt(mnFileOpt).Checked = False
   mnFileOpt = Index
   mnuFileOpt(mnFileOpt).Checked = True

End Sub

Private Sub mnuFlow_Click(Index As Integer)

   mnuFlow(mnErrHandling).Checked = False
   mnuFlow(Index).Checked = True
   mnErrHandling = Choose(Index + 1, DONTSTOP, STOPFATAL, STOPALL)
   ULStat = cbErrHandling(mnErrReporting, mnErrHandling)
   If SaveFunc(Me, ErrHandling, ULStat, mnErrReporting, mnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   UpdateStatus

End Sub

Private Sub mnuHist_Click()

   mnuHist.Visible = False
   mnuSelLocHandling.Visible = False
   mnuSelLibHandling.Visible = False
   mnuSelHistOpts.Visible = True
   GetFuncsSaved manFuncList()
   mnNumFuncs = GetHistory()
   'txtNumFuncs.Text = GetHistory()
   UPDATEGRID
      
   txtErrorMessage.Visible = False
   fraBufferSetup.Visible = True
   fraFormat.Visible = False
   cmdDone.Visible = False
   cmdClear.Visible = False

End Sub

Private Sub mnuLiveDisplay_Click()
   
   mnuLiveDisplay.Checked = Not mnuLiveDisplay.Checked
   gnLiveDisp = 0
   If mnuLiveDisplay.Checked Then gnLiveDisp = True
   UpdateStatus

End Sub

Private Sub mnuLocHandling_Click(Index As Integer)

   mnuLocHandling(meErrFlow).Checked = False
   meErrFlow = Index
   mnuLocHandling(meErrFlow).Checked = True
   UpdateStatus

End Sub

Private Sub mnuLocPrint_Click()
   
   mnuLocPrint.Checked = Not mnuLocPrint.Checked
   mnErrDisp = mnuLocPrint.Checked
   UpdateStatus

End Sub

Private Sub mnuLocWarnHandling_Click(Index As Integer)

   mnuLocWarnHandling(meWarnFlow).Checked = False
   meWarnFlow = Index
   mnuLocWarnHandling(meWarnFlow).Checked = True
   UpdateStatus

End Sub

Private Sub mnuLocWarnPrint_Click()
   
   mnuLocWarnPrint.Checked = Not mnuLocWarnPrint.Checked
   mnWarnDisp = 0
   If mnuLocWarnPrint.Checked Then mnWarnDisp = 1
   UpdateStatus

End Sub

Private Sub mnuPrint_Click(Index As Integer)

   mnuPrint(mnErrReporting).Checked = False
   mnuPrint(Index).Checked = True
   mnErrReporting = Choose(Index + 1, DONTPRINT, PRINTWARNINGS, PRINTFATAL, PRINTALL)
   ULStat = cbErrHandling(mnErrReporting, mnErrHandling)
   If SaveFunc(Me, ErrHandling, ULStat, mnErrReporting, mnErrHandling, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub
   UpdateStatus

End Sub

Private Sub optFormat_Click(Index As Integer)

   gnHexVals = optFormat(1).value
   If Me.Visible Then ErrorHandling
   lplFileName$ = "UniTest.ini"
   lpApplicationName$ = "DataFormat"
   lpKeyName$ = "Notation"
   lpString$ = Str$(gnHexVals)
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)

End Sub

Private Sub picSelFunctions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   Row% = Int(y) + 1
   If Row% > 4 Then Row% = 4
   Col% = 4 * Int(x)
   'lblStatus.Caption = GetFunctionString(Row% + Col%)
   txtStatus.Text = GetFunctionString(Row% + Col%)

End Sub

Private Sub picSelFunctions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   Row% = Int(y) + 1
   If Row% > 4 Then Row% = 4
   Col% = 4 * Int(x)
   FuncID% = Row% + Col%
   manFuncList(FuncID%) = Not (manFuncList(FuncID%))
   SetColor Int(x), Int(y), manFuncList(FuncID%)
   picSelFunctions.CurrentX = Int(x) + 0.1
   picSelFunctions.CurrentY = Int(y) + 1
   picSelFunctions.Print Mid$(GetFunctionString(FuncID%), 3, 4)

End Sub

Private Sub picSelFunctions_Paint()
   
   UPDATEGRID

End Sub

Private Sub SetColor(x As Integer, y As Integer, NotEnabled As Integer)

   TheColor& = &HFF00&
   If NotEnabled Then TheColor& = &HFF&
   
   picSelFunctions.Line (x + 0.1, y + 0.1)-(x + 0.9, y + 0.9), TheColor&, BF

End Sub

Private Sub UPDATEGRID()

   picSelFunctions.Cls
   For x% = 0 To 23
      For y% = 0 To 3
         z% = 4 * x% + (y% + 1)
         If z% > mnMaxFuncs Then
            SetColor x%, y%, -1
         Else
            SetColor x%, y%, manFuncList(z%)
            picSelFunctions.CurrentX = x% + 0.1
            picSelFunctions.CurrentY = y% + 1
            picSelFunctions.Print Mid$(GetFunctionString(z%), 3, 4)
         End If
      Next y%
   Next x%

End Sub

Private Sub UpdateStatus()
   
   For i% = 0 To 3
      If mnuPrint(i%).Checked Then PrintErr$ = mnuPrint(i%).Caption
      If mnuLocHandling(i%).Checked Then LocHandling$ = mnuLocHandling(i%).Caption
      If mnuLocWarnHandling(i%).Checked Then LocWarnHandling$ = mnuLocWarnHandling(i%).Caption
   Next i%
   For i% = 0 To 2
      If mnuFlow(i%).Checked Then
         FlowErr$ = mnuFlow(i%).Caption
         Exit For
      End If
   Next i%
   LocalPrint$ = ", Print Errors"
   LocalWarnPrint$ = ", Print Warnings"
   If Not (mnuLocPrint.Checked) Then LocalPrint$ = ", Don't Print Errors"
   If Not (mnuLocWarnPrint.Checked) Then LocalWarnPrint$ = ", Don't Print Warnings"
   'lblStatus.Caption = "cbErrHandling(" & PrintErr$ & ", " & FlowErr$ & ")"
   txtStatus.Text = "cbErrHandling(" & PrintErr$ & ", " & FlowErr$ & ")"
   'lblStatus.Caption = lblStatus.Caption & " ... Local Error Handling: " & LocHandling$ & LocalPrint$ & _
   lblStatus.Caption = "cbErrHandling(" & PrintErr$ & ", " & FlowErr$ & ")"
   txtStatus.Text = txtStatus.Text & " ... Local Error Handling: " & LocHandling$ & LocalPrint$ & vbCrLf '& _
   'txtStatus.Text = "cbErrHandling(" & PrintErr$ & ", " & FlowErr$ & ")"
   'lblStatus.Caption = lblStatus.Caption & " ... Local Error Handling: " & LocHandling$ & LocalPrint$ & _
   LocWarnHandling$ & LocalWarnPrint$
   txtStatus.Text = txtStatus.Text & "Local Warning Handling: " & LocWarnHandling$ & LocalWarnPrint$
   
End Sub

Public Sub SetNoError()

   mnNoError = True
   
End Sub
