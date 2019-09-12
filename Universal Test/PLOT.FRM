VERSION 5.00
Begin VB.Form frmPlot 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Plot Form"
   ClientHeight    =   3285
   ClientLeft      =   3765
   ClientTop       =   6060
   ClientWidth     =   5505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3285
   ScaleWidth      =   5505
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   915
      Left            =   60
      ScaleHeight     =   3.688
      ScaleMode       =   4  'Character
      ScaleWidth      =   29.375
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.TextBox txtShow 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox picPlot 
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
      Height          =   1455
      Left            =   1200
      MousePointer    =   2  'Cross
      ScaleHeight     =   1425
      ScaleMode       =   0  'User
      ScaleWidth      =   5000
      TabIndex        =   2
      Top             =   60
      Width           =   2415
      Begin VB.Label lblAnalyzing 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Analyzing..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   700
         Left            =   300
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Height          =   615
         Left            =   1080
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line linCursor 
         BorderColor     =   &H00FF0000&
         Visible         =   0   'False
         X1              =   4654.088
         X2              =   4654.088
         Y1              =   0
         Y2              =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   2  'Dash
         X1              =   0
         X2              =   13081.76
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   13081.76
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   13081.76
         Y1              =   1620
         Y2              =   1620
      End
   End
   Begin VB.Label lblSampNumLabel 
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   1005
   End
   Begin VB.Label lblAmplLabel 
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   1005
   End
   Begin VB.Label lblSampNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label lblAmpl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   1095
   End
   Begin VB.Menu mnuChan 
      Caption         =   "&Channel"
      Begin VB.Menu mnuChPlot 
         Caption         =   "Plot"
         Begin VB.Menu mnuChanUnselect 
            Caption         =   "Un&select Chan"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAllChans 
            Caption         =   "&All (Ctl A)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOneChan 
            Caption         =   "Ch0 (Ctl 0)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuCursorChan 
         Caption         =   "Cursor"
         Begin VB.Menu mnuCursChan 
            Caption         =   "Ch0 (Ctl-Alt 0)"
            Checked         =   -1  'True
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuSetBlocks 
      Caption         =   "Set &Blocks"
      Begin VB.Menu mnuBlockSize 
         Caption         =   "Block &Size"
      End
      Begin VB.Menu mnuFirstPoint 
         Caption         =   "First &Point"
      End
      Begin VB.Menu mnuTextRange 
         Caption         =   "Text Range"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuSelRes 
         Caption         =   "Resolution"
         Begin VB.Menu mnuResolution 
            Caption         =   "Normal"
            Index           =   0
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "Autodetect"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "4 Bit"
            Index           =   2
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "8 Bit"
            Index           =   3
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "12 Bit"
            Index           =   4
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "13 Bit"
            Index           =   5
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "14 Bit"
            Index           =   6
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "16 Bit"
            Index           =   7
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "24 Bit"
            Index           =   8
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "32 Bit"
            Index           =   9
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "48 Bit"
            Index           =   10
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "64 Bit"
            Index           =   11
         End
         Begin VB.Menu mnuSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSetManScale 
            Caption         =   "Set Scale..."
         End
      End
      Begin VB.Menu mnuTitle 
         Caption         =   "Title"
         Begin VB.Menu mnuTitleType 
            Caption         =   "Scale Mode"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuTitleType 
            Caption         =   "Data Origin"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRetain 
         Caption         =   "Retain Plot"
      End
      Begin VB.Menu mnuHex 
         Caption         =   "Hex Text"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "View Text Window"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveData 
         Caption         =   "Save to File"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuDetails 
      Caption         =   "Details"
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnLoaded As Integer, mvRefVal As Variant, mvRefSample As Variant
Dim mbHandleMouseEvent As Integer, mnLastChan As Integer
Dim mRefPoint As Integer, mnShowDiff As Integer, mnDetails As Integer
Dim mnCursorChan As Integer, mnTitleType As Integer
Dim mnBlocksEnabled As Integer, mnRetainPlot As Integer
Dim mnSavedRes, mnUnitCode As Integer
Dim msUnit As String

Private Sub Form_Activate()

   mnuHex.Checked = gnHexVals

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   If Shift = 0 Then
      Select Case KeyCode
         Case 27
            'escape to end background scans
            ReturnToOwner
         Case 33
            'page up - prev block plot
         Case 34
            'page down - next block plot
            UpdateBlock
         Case vbKeyF6
            ReturnToOwner
            If Not gfOwner Is Nothing Then
               gfOwner.ShowBufferInfo 1
            End If
      End Select
   End If

   If (Shift And 6) = 6 Then  'if Ctl-Alt keys
      Select Case KeyCode
         Case 48 To 57
            'if a number from 0 to 9
            PlusTen = 0
            If (Shift And 1) = 1 Then PlusTen = 10
            Ch% = Val(Chr$(KeyCode)) + PlusTen
            mnuCursChan_Click (Ch%)
         Case 73  '"i" - buffer info
            ReturnToOwner
         Case 78  '"n" - next block
            NextBlock mnCursorChan
      End Select
   ElseIf (Shift And 2) = 2 Then  'if Ctl key
      Select Case KeyCode
         Case 13  'Enter
            'run owners default function
            gnRunDefault = True
            ReturnToOwner
         Case Asc("'")
            mfmUniTest.cmdFormType(6) = True
         Case 48 To 57
            PlusTen = 0
            If (Shift And 1) = 1 Then PlusTen = 10
            Ch% = Val(Chr$(KeyCode)) + PlusTen
            mnuOneChan_Click (Ch%)
         Case 96 To 105
            PlusTen = 20
            'If (Shift And 1) = 1 Then PlusTen = 30
            Ch% = (KeyCode - 96) + PlusTen
            mnuOneChan_Click (Ch%)
         Case Asc("A")
            mnuAllChans_Click
         Case Asc("C")
            mfmUniTest.cmdFormType(4) = True
         Case Asc("G")
            'run owners default function
            gnRunDefault = True
            ReturnToOwner
         Case Asc("I")
            mfmUniTest.cmdFormType(0) = True
         Case Asc("L")
            mfmUniTest.cmdUtils = True
         Case Asc("M")
            mfmUniTest.cmdFormType(5) = True
         Case Asc("N")
            mfmUniTest.cmdFormType(2) = True
         Case Asc("O")
            mfmUniTest.cmdFormType(1) = True
         Case Asc("T")
            mfmUniTest.cmdFormType(3) = True
         Case 191    'Alt /
            'set focus to owner
            ReturnToOwner
      End Select
   ElseIf (Shift And 4) = 4 Then  'if Alt key
      If gnScriptRun Then
         Select Case KeyCode
            Case 82  '"R"
               frmScript.cmdScript(2).value = True
            Case 84  '"T"
               frmScript.cmdScript(3).value = True
         End Select
      End If
   End If

End Sub

Private Sub Form_Load()

   PositionMe
   SetChannel -1
   DoEvents
   FormLoading
   mnLoaded = True
   mnuHex.Checked = gnHexVals
   
   lplFileName$ = "PlotForm.ini"
   lpApplicationName$ = "Options"
   lpKeyName$ = "Details"
   lpDefault$ = "-1"
   nSize% = 4
   lpReturnedString$ = Space$(nSize%)
   StringSize% = GetPrivateProfileString(lpApplicationName$, lpKeyName$, lpDefault$, lpReturnedString$, nSize%, lplFileName$)
   mnDetails = Val(lpReturnedString$)
   picDetails.Visible = mnDetails
   ResizeObjects
   CurSize& = GetBlockSize()

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   UpdateBlock
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   linCursor.Visible = False

End Sub

Private Sub Form_Resize()

   mbHandleMouseEvent = False
   linCursor.Visible = False
   
   If Me.WindowState <> 1 Then
      picPlot.Cls
      ResizeObjects
      frmPlot.Visible = True
   End If
   mbHandleMouseEvent = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

   If WindowState = 0 Then
      lplFileName$ = "PlotForm.ini"
      lpApplicationName$ = "Size"
      lpKeyName$ = "Width"
      lpString$ = Str$(Width)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)
      lpKeyName$ = "Height"
      lpString$ = Str$(Height)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)
   
      lpApplicationName$ = "Position"
      lpKeyName$ = "Left"
      lpString$ = Str$(Left)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)
      lpKeyName$ = "Top"
      lpString$ = Str$(Top)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)
   
      lpApplicationName$ = "Options"
      lpKeyName$ = "Details"
      lpString$ = Str$(mnDetails)
      x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)

      SaveBlockSize

   End If

   FormUnLoading

End Sub

Private Sub mnuAllChans_Click()

   SetChannel -1
   RePlot False

End Sub

Private Sub mnuBlockSize_Click()

   mbHandleMouseEvent = False
   linCursor.Visible = False
   CurSize& = GetBlockSize()
   
   sDefault$ = Format$(CurSize&, "0")
   Block = InputBox("Enter block size", "Set Block Size", sDefault$)
   If Len(Block) Then
      BlockValue& = Val(Block)
      If BlockValue& <> CurSize& Then
         If (Not WIN32APP) And (BlockValue& > 32767) Then
            MsgBox "16 bit VB allows only 32k array sizes. Block size will be set to 32766", , "Block Size Error"
            BlockValue& = 32766
         End If
         RestorePlot True
         SetBlockSize BlockValue&, True
         RePlot False
      End If
   End If
   mbHandleMouseEvent = True
   
End Sub

Private Sub mnuChanUnselect_Click()

   mnuChanUnselect.Checked = Not mnuChanUnselect.Checked
   
End Sub

Private Sub mnuCursChan_Click(Index As Integer)

   mnLastChan = GetLastChan()
   If Index > mnLastChan Then Exit Sub
   mnuCursChan(mnCursorChan).Checked = False
   mnCursorChan = Index
   mnuCursChan(mnCursorChan).Checked = True
   NewColor = glaPlotColor(Index Mod 8)
   lblAmpl.ForeColor = NewColor
   lblSampNum.ForeColor = NewColor
   linCursor.BorderColor = NewColor
   If Not mnShowDiff Then
      lblAmplLabel.ForeColor = NewColor
      lblSampNumLabel.ForeColor = NewColor
   End If

End Sub

Private Sub mnuDetails_Click()

   mnDetails = Not mnDetails
   picDetails.Visible = mnDetails
   ResizeObjects
   
   lpApplicationName$ = "Options"
   lpKeyName$ = "Details"
   lpString$ = Str$(mnDetails)
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)

End Sub

Private Sub mnuFirstPoint_Click()

   StartBlock& = GetFirstPoint()
   DefVal$ = Format(StartBlock&, "0")
   FP = InputBox("Enter first point", "Set First Point", DefVal$)
   If Not FP = "" Then
      FirstPointVal& = Val(FP)
      SetFirstPoint FirstPointVal&
      frmPlot.mnuRestore.ENABLED = Not (FirstPointVal& = 0)
      'BufferHandle& = GetBufHandle()
      'If Not BufferHandle& = 0 Then TransferBlock
      RePlot True
   End If

End Sub

Private Sub mnuHex_Click()

   mnuHex.Checked = Not mnuHex.Checked
   gnHexVals = mnuHex.Checked
   lplFileName$ = "UniTest.ini"
   lpApplicationName$ = "DataFormat"
   lpKeyName$ = "Notation"
   lpString$ = Str$(gnHexVals)
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, lpString$, lplFileName$)
   UpdateBlock
   
End Sub

Private Sub mnuOneChan_Click(Index As Integer)

   SetChannel Index
   RePlot False

End Sub

Private Sub mnuResolution_Click(Index As Integer)

   Dim ManualSet As Boolean
   
   ResVal% = Choose(Index + 1, -1, 0, 4, 8, 12, 13, 14, 16, 24, 32, 48, 64)
   If ResVal% < 0 Then
      ManualSet = False
      ResVal% = mnSavedRes
   Else
      ManualSet = True
   End If
   SetPlotResolution ResVal%, ManualSet

End Sub

Private Sub mnuSaveData_Click()

   mnuSaveData.Checked = True
   If mnuSaveData.Checked Then
      frmScriptFiles.Caption = "Save Data As Text"
      frmScriptFiles.cmbPattern.ENABLED = False
      frmScriptFiles.cmbPattern.AddItem "*.txt;*.csv"
      frmScriptFiles.cmbPattern.ListIndex = 0
      frmScriptFiles.Show 1
      filespec$ = frmScriptFiles.txtHeader.Text
      FileSaved% = SaveDataAsText(filespec$)
   End If
   mnuSaveData.Checked = False
   
End Sub

Private Sub mnuSetManScale_Click()

   YMaxStr$ = InputBox("Enter max Y value: ", "Manual Scale", "4095")
   If YMaxStr$ = "" Then Exit Sub
   YMinStr$ = InputBox("Enter min Y value: ", "Manual Scale", "0")
   If YMinStr$ = "" Then Exit Sub
   YMax! = Val(YMaxStr$)
   picPlot.ScaleTop = YMax!
   YMin! = Val(YMinStr$)
   NumChoices& = mnuResolution.Count - 1
   For mnuIndex& = 0 To NumChoices&
      mnuResolution(mnuIndex&).Checked = False
   Next
   SetManualScale YMin!, YMax!

End Sub

Private Sub mnuRestore_Click()
   
   RestorePlot True

End Sub

Private Sub mnuRetain_Click()

   mnuRetain.Checked = Not mnuRetain.Checked
   mnRetainPlot = mnuRetain.Checked
   SetPlotOptions mnRetainPlot, mnTitleType

End Sub

Private Sub mnuTextRange_Click()

   GetTextRange StartText&, TextCount&, MaxEnd&
   sDefault$ = Format(StartText&, "0") & ", " & Format(TextCount&, "0")
   
   Block = InputBox("Enter beginning sample and number of samples" & _
   vbCrLf & "(up to " & Format(MaxEnd&, "0") & ") separated by a comma", _
   "Set Text Range", sDefault$)
   
   If Block = "" Then Exit Sub
   TextRange = Split(Block, ",")
   If Not UBound(TextRange) = 1 Then
      MsgBox "Improper format for Text Range.", vbCritical, "Range Not Set"
      Exit Sub
   End If
   StartText& = Val(TextRange(0))
   TextCount& = Val(TextRange(1))
   If StartText& + TextCount& > MaxEnd& Then
      If StartText& < (MaxEnd& - 1) Then
         TextCount& = MaxEnd& - StartText&
         MsgBox "Number of samples reset to " & Format(TextCount&, "0") & _
         " to avoid reading past end of buffer.", vbCritical, "Range Reset"
      Else
         MsgBox "Starting sample (" & Format(StartText&, "0") & _
         ") is invalid.", vbCritical, "Range Not Set"
         Exit Sub
      End If
   End If
   SetTextRange StartText&, TextCount&
   RePlot 0
   
End Sub

Private Sub mnuTitleType_Click(Index As Integer)

   mnuTitleType(mnTitleType).Checked = False
   mnTitleType = Index
   mnuTitleType(mnTitleType).Checked = True
   SetPlotOptions mnRetainPlot, mnTitleType

End Sub

Private Sub mnuWindow_Click()

   mnuWindow.Checked = Not mnuWindow.Checked
   picPlot.Visible = Not mnuWindow.Checked
   txtShow.Visible = mnuWindow.Checked
   
End Sub

Private Sub picPlot_DblClick()
   
   lblSampNumLabel.Caption = ""
   lblAmplLabel.Caption = ""
   SetAutoScale
   
End Sub

Private Sub picPlot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   lblSampNumLabel.Caption = ""
   lblAmplLabel.Caption = ""
   If Button = 2 Then
      If mnShowDiff Then
         mnShowDiff = False
         lblSampNumLabel.FontName = "Arial"
         lblAmplLabel.FontName = "Arial"
      Else
         mnShowDiff = True
         lblSampNumLabel.FontName = "Symbol"
         lblAmplLabel.FontName = "Symbol"
         SampleNum = GetSampleNum(Int(x + 0.5), SampValue, mnCursorChan)
         mRefPoint = CInt(x)
         mvRefSample = SampleNum
         mvRefVal = SampValue
         NewColor = glaPlotColor(mnCursorChan Mod 8)
         lblSampNumLabel.ForeColor = NewColor
         lblAmplLabel.ForeColor = NewColor
      End If
   Else
      Shape1.Left = x
      Shape1.Top = y
   End If
   
End Sub

Private Sub picPlot_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 0 Then
      If mnuCursChan(0).Checked Then
         'If # of chans has been reduced, plott.bas will
         'set mnuCursChan(0).Checked to true but the
         'event must be triggered here
         mnCursorChan = 0
         mnuCursChan_Click (0)
      End If
      linCursor.Visible = True
      linCursor.X1 = x
      linCursor.X2 = x
      If mbHandleMouseEvent Then
         CursorChan% = mnCursorChan
         If mnShowDiff Then
            lblSampNumLabel.Caption = "D " & mRefPoint & ":"
            SampleNum& = GetSampleNum(Int(x + 0.5), SampValue, CursorChan%) - mvRefSample
            If SampleNum& < 0 Then
               lblSampNum.Caption = "???"
            Else
               lblSampNum.Caption = SampleNum&
               SampValue = SampValue - mvRefVal
               ValString$ = "D " & mvRefVal & ":"
            End If
         Else
            GetSettleStep SettleStep&, SettleRate&
            If SettleStep& > 0 Then
               lblSampNumLabel.Caption = "Time (µs): "
               SampleNum& = GetSampleNum(Int(x + 0.5), SampValue, CursorChan%)
               If SampleNum& < 0 Then
                  lblSampNum.Caption = "???"
               Else
                  lblSampNum.Caption = Format$((1 / (SettleRate& - SettleStep& * SampleNum&)) * 500000, "0.000")
               End If
            Else
               lblSampNumLabel.Caption = "Sample #: "
               SampleNum& = GetSampleNum(Int(x + 0.5), SampValue, CursorChan%)
               If SampleNum& < 0 Then
                  lblSampNum.Caption = "???"
               Else
                  lblSampNum.Caption = SampleNum&
               End If
            End If
            ValString$ = "Ch " & CursorChan% & " value: "
         End If
         lblAmplLabel.Caption = ValString$
         SampString$ = SampValue
         If gnHexVals Then
            'x = SampValue#
            'If x > 2147483647 Then x = CLng(SampValue# - 2 ^ 32)
            lblAmpl.Caption = SampValue
         Else
            If InStr(1, SampString$, ".") > 0 Then
               If mnUnitCode = 4 Then
                  If SampValue < 0.01 Then
                     SampString$ = Format(SampValue * 1000, "0.0###mV")
                  Else
                     SampString$ = Format(SampValue, "0.0###V")
                  End If
               Else
                  SampString$ = Format(SampValue, "0.0###" & msUnit)
               End If
            End If
            lblAmpl.Caption = SampString$
         End If
      End If
   ElseIf Button = 1 Then
      If (picPlot.ScaleWidth > 10) Then
         linCursor.Visible = False
         Shape1.Visible = True
         MousePosX! = x - Shape1.Left
         If MousePosX! > 0 Then Shape1.Width = MousePosX!
         Shape1.Height = (Me.picPlot.ScaleHeight * -0.95)
         Shape1.Top = Me.picPlot.ScaleHeight * -0.975  '/ 2
      End If
   End If

End Sub

Private Sub picPlot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If (Shape1.Visible And (picPlot.ScaleWidth > 10)) Then
      PortionSelected! = ((Shape1.Width + picPlot.ScaleWidth) / (picPlot.ScaleWidth) - 1) * 100
      If PortionSelected! > 2 Then
         CursorChan% = 0
         Start% = CInt(Shape1.Left)
         Ending% = CInt(Shape1.Left + Shape1.Width)
         SampNumStart& = GetSampleNum(Start%, SampValue, CursorChan%)
         CursorChan% = 0
         SampNumEnd& = GetSampleNum(Ending%, SampValue, CursorChan%)
         BlockValue& = (SampNumEnd& - SampNumStart&) / (mnLastChan + 1) '* -1
         SetBlockSize BlockValue&, False
         SetFirstPoint SampNumStart&
         'BufferHandle& = GetBufHandle()
         InitBlock False
         
         RePlot True
         'If Not BufferHandle& = 0 Then TransferBlock
      End If
   End If
   Shape1.Visible = False

End Sub

Private Sub PositionMe()

   lpFileName$ = "PlotForm.ini"
   lpApplicationName$ = "Size"
   nDefault% = 4000
   lpKeyName$ = "Height"
   Height = GetPrivateProfileInt(lpApplicationName$, lpKeyName$, nDefault%, lpFileName$)
   nDefault% = 6040
   lpKeyName$ = "Width"
   Width = GetPrivateProfileInt(lpApplicationName$, lpKeyName$, nDefault%, lpFileName$)
   
   lpApplicationName$ = "Position"
   lpKeyName$ = "Left"
   lpString$ = Str$(Left)
   nDefault% = Screen.Width - (frmPlot.Width + 200)
   Left = GetPrivateProfileInt(lpApplicationName$, lpKeyName$, nDefault%, lpFileName$)
   lpKeyName$ = "Top"
   nDefault% = Screen.Height * 0.45
   lpString$ = Str$(Top)
   Top = GetPrivateProfileInt(lpApplicationName$, lpKeyName$, nDefault%, lpFileName$)

End Sub

Private Sub ResizeObjects()
   
   If Me.Width - 300 < 2000 Then Me.Width = 2300
   SaveWidth% = picPlot.ScaleWidth
   picPlot.Width = Me.Width - 1500
   picPlot.ScaleWidth = SaveWidth%
   If mnDetails Then
      If Me.Height - 1350 < 1200 Then Me.Height = 2550
      picPlot.Height = Me.ScaleHeight - 1100
      picDetails.Width = Me.Width - 350
      picDetails.Top = Me.ScaleHeight - 1000
   Else
      If Me.Height - 1350 < 400 Then Me.Height = 1650
      picPlot.Height = Me.ScaleHeight - 120
   End If
   txtShow.Width = picPlot.Width
   txtShow.Height = picPlot.Height

   Line1.X2 = picPlot.ScaleWidth
   Line1.Y1 = picPlot.ScaleTop + picPlot.ScaleHeight / 2
   Line1.Y2 = picPlot.ScaleTop + picPlot.ScaleHeight / 2
   Line2.X2 = picPlot.ScaleWidth
   Line2.Y1 = picPlot.ScaleTop + picPlot.ScaleHeight / 4
   Line2.Y2 = picPlot.ScaleTop + picPlot.ScaleHeight / 4
   Line3.X2 = picPlot.ScaleWidth
   Line3.Y1 = picPlot.ScaleTop + (picPlot.ScaleHeight / 4) * 3
   Line3.Y2 = picPlot.ScaleTop + (picPlot.ScaleHeight / 4) * 3
   linCursor.Y1 = picPlot.ScaleTop
   linCursor.Y2 = picPlot.ScaleHeight
   picPlot.ForeColor = &HFF0000
   DoEvents
   If mnLoaded Then RePlot False

End Sub

Private Sub txtShow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   RePlot False

End Sub

Sub UpdateBlock()
   
   If mnuRestore.ENABLED Then Exit Sub
   Chan% = mnCursorChan
   NextBlock (Chan%)
   Exit Sub
   BufferHandle& = GetBufHandle()
   If BufferHandle& = 0 Then Exit Sub
   'SetBlock
   TransferBlock
   Chan% = mnCursorChan
   If Not txtShow.Visible Then
      If mnShowDiff Then
         lblSampNumLabel.Caption = "D " & mRefPoint & ":"
         lblSampNum.Caption = GetSampleNum(0, SampValue, Chan%) - mvRefSample
         SampValue = SampValue - mvRefVal
         ValString$ = "D " & mvRefVal & ":"
      Else
         lblSampNumLabel.Caption = "Sample #: "
         lblSampNum.Caption = GetSampleNum(0, SampValue, Chan%)
         ValString$ = "Ch " & Chan% & " value: "
      End If
      lblAmplLabel.Caption = ValString$
      lblAmpl.Caption = SampValue
   End If

End Sub

Public Sub ShowText(TrueOrFalse As Boolean)

   If Not (mnuWindow.Checked = TrueOrFalse) _
      Then mnuWindow_Click
      
End Sub

Public Sub SetUnit(UnitCode As Integer)

   mnUnitCode = UnitCode
   Select Case UnitCode
      Case 0
         msUnit = ""
      Case 1
         msUnit = "°C"
      Case 2
         msUnit = "°F"
      Case 3
         msUnit = "K"
      Case 4
         msUnit = "V"
      Case Else
         msUnit = ""
   End Select
   
End Sub

