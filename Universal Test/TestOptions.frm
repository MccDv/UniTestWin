VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTestOptions 
   Caption         =   "Select Test Options"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSetup 
      Caption         =   "Select Test and Device"
      Height          =   2535
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   7035
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5880
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Default         =   -1  'True
         Height          =   375
         Left            =   5880
         TabIndex        =   22
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame fraTestGroups 
         Caption         =   "Test Groups"
         Height          =   1095
         Left            =   4560
         TabIndex        =   19
         Top             =   240
         Width           =   2355
         Begin VB.ComboBox cmbTest 
            Height          =   315
            Left            =   180
            TabIndex        =   21
            Top             =   660
            Width           =   2000
         End
         Begin VB.ComboBox cmbTestCat 
            Height          =   315
            Left            =   180
            TabIndex        =   20
            Text            =   "Category"
            Top             =   300
            Width           =   2000
         End
      End
      Begin VB.Frame fraProducts 
         Caption         =   "Product Groups"
         Height          =   1095
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   4275
         Begin VB.ComboBox cmbGroup 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Text            =   "Group"
            Top             =   300
            Width           =   1455
         End
         Begin VB.ComboBox cmbSubProduct 
            Height          =   315
            Left            =   1680
            TabIndex        =   18
            Text            =   "SubProduct"
            Top             =   660
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox cmbProduct 
            Height          =   315
            Left            =   1680
            TabIndex        =   17
            Text            =   "Product"
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label lblFileName 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Visible         =   0   'False
            Width           =   1395
         End
      End
      Begin VB.Label lblTestFile 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   180
         TabIndex        =   25
         Top             =   1500
         Width           =   5535
      End
      Begin VB.Label lblDParmPath 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Top             =   1980
         Width           =   5535
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3942
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Common"
      TabPicture(0)   =   "TestOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCommon"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Test Blocks"
      TabPicture(1)   =   "TestOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTestBlocks"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Parameters"
      TabPicture(2)   =   "TestOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraParameters"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraParameters 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   6375
         Begin VB.ComboBox cmbParamSelect 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   28
            Text            =   "ParamSelect"
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtParam 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblParam 
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   14
            Top             =   180
            Width           =   2235
         End
      End
      Begin VB.Frame fraTestBlocks 
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   -74940
         TabIndex        =   10
         Top             =   360
         Width           =   6375
         Begin VB.CheckBox chkTest 
            Caption         =   "Test Name"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.Frame fraCommon 
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   6315
         Begin VB.CheckBox chkUseDF 
            Caption         =   "Use DAQFlex"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   180
            Width           =   1935
         End
         Begin VB.CheckBox chkSE 
            Caption         =   "Use Single-ended input"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   420
            Width           =   2415
         End
         Begin VB.TextBox txtHighChan 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "0"
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox txtLowChan 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Text            =   "0"
            Top             =   780
            Width           =   495
         End
         Begin VB.Label lblHighChannel 
            Caption         =   "High Channel"
            Height          =   195
            Left            =   720
            TabIndex        =   7
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblLowChan 
            Caption         =   "Low Channel"
            Height          =   195
            Left            =   720
            TabIndex        =   5
            Top             =   840
            Width           =   1575
         End
      End
   End
   Begin VB.Label lblSerialNumber 
      Height          =   195
      Left            =   3060
      TabIndex        =   2
      Top             =   60
      Width           =   2775
   End
   Begin VB.Label lblDeviceName 
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "frmTestOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msRoot As String, msTestDir As String
Dim mnGroupIndex As Integer
Dim mbUseDropCommand As Boolean, mbDevSelected As Boolean

Public Sub SetParamList(ByVal ParamList As String)

   ListArray = Split(ParamList$, ";")
   ListSize& = UBound(ListArray)
   If Not ListSize& < 0 Then Me.SSTab1.TabVisible(1) = True
   DefaultWidth& = Me.Width
   FirstBoxTop& = frmTestOptions.chkTest(0).Top
   HorizPos& = frmTestOptions.chkTest(0).Left
   TopOffset& = FirstBoxTop&
   VertPos& = 1
   For ItemInList& = 0 To ListSize&
      CurListItem$ = ListArray(ItemInList&)
      If ItemInList& > 0 Then
         Load Me.chkTest(ItemInList&)
         chkTest(ItemInList&).Top = (260 * VertPos&) + TopOffset&
         VertPos& = VertPos& + 1
         chkTest(ItemInList&).Left = HorizPos&
         chkTest(ItemInList&).Visible = True
         chkTest(ItemInList&).value = 0
         If VertPos& = 6 Then
            VertPos& = 0
            HorizPos& = HorizPos& + 3000
            If HorizPos& > (DefaultWidth& - 1500) Then Wider& = Wider& + 2400
         End If
      End If
      Me.chkTest(ItemInList&).Caption = CurListItem$
      OptVal% = frmScript.GetTestOptValue(ItemInList&)
      If Not (OptVal% = 0) Then Me.chkTest(ItemInList&).value = 1
   Next
   Me.Width = DefaultWidth& + Wider&

End Sub

Public Sub SetValueList(ByVal ParamList As String)

   Dim Args(0)
   ListArray = Split(ParamList$, ";")
   ListSize& = UBound(ListArray)
   VertPos& = 0
   HorizPos& = 0
   DefaultWidth& = Me.Width
   If Not ListSize& < 0 Then Me.SSTab1.TabVisible(2) = True
   FirstBoxTop& = frmTestOptions.txtParam(0).Top
   FirstBoxLeft& = frmTestOptions.txtParam(0).Left
   FirstLabelLeft& = lblParam(0).Left
   Dim CurControl As Control
   For ItemInList& = 0 To ListSize&
      CurListItem$ = ListArray(ItemInList&)
      TestOptName$ = "param" & Format(ItemInList&, "0")
      Args(0) = TestOptName$
      VarFound% = frmScript.CheckForVariables(Args)
      Parameter$ = Args(0)
      MultiSelArray = Split(Parameter$, "$")
      NumSel& = UBound(MultiSelArray)
      If ItemInList& > 0 Then
         'load the control and its label
         Load Me.lblParam(ItemInList&)
         If NumSel& > 0 Then
            'if control contains a list, make it a combo box
            Load Me.cmbParamSelect(TextCount&)
            Set CurControl = cmbParamSelect(ComboCount&)
            ComboCount& = ComboCount& + 1
         Else
            'if control contains one item, make it a text box
            Load Me.txtParam(TextCount&)
            Set CurControl = txtParam(TextCount&)
            TextCount& = TextCount& + 1
         End If
         If VertPos& = 4 Then
            VertPos& = 0
            HorizPos& = HorizPos& + 1
            If HorizPos& > (DefaultWidth& - 1500) Then Wider& = Wider& + 2400
         Else
            VertPos& = VertPos& + 1
         End If
      Else
         'for the first control, no need to load it or position it
         If NumSel& > 0 Then
            'if control contains a list, show the combo box and hide the text box
            txtParam(TextCount&).Visible = False
            ComboCount& = ComboCount& + 1
         Else
            'if control contains one item, leave the combo box hidden
            Set CurControl = txtParam(TextCount&)
            TextCount& = TextCount& + 1
         End If
      End If
      Me.lblParam(ItemInList&).Caption = CurListItem$
      Me.lblParam(ItemInList&).Left = FirstLabelLeft& + (3000 * HorizPos&)
      lblParam(ItemInList&).Top = (320 * VertPos&) + FirstBoxTop&
      lblParam(ItemInList&).Visible = True
      CurControl.Tag = ItemInList&
      CurControl.Visible = True
      CurControl.Top = lblParam(ItemInList&).Top
      CurControl.Left = FirstBoxLeft& + (3000 * HorizPos&)
      If TypeOf CurControl Is TextBox Then
         CurControl.Text = Parameter$
      Else
         For ParamItem& = 0 To NumSel&
            CurControl.AddItem MultiSelArray(ParamItem&)
         Next
         CurControl.ListIndex = 0
         lblParam(ItemInList&).Left = lblParam(ItemInList&).Left + 500
      End If
   Next
   Me.Width = DefaultWidth& + Wider&

End Sub

Public Sub SetPaths(ByVal Root As String, ByVal ProdGroups As String)

   msRoot = Root
   GroupArray = Split(ProdGroups, ";")
   GroupSize& = UBound(GroupArray)
   For GNum& = 0 To GroupSize&
      cmbGroup.AddItem GroupArray(GNum&)
   Next
   cmbGroup.ListIndex = 0

End Sub

Public Sub SetTestDir(ByVal TestDir As String)

   msTestDir = TestDir
   SetTestCats
   
End Sub

Private Sub SetTestCats()

   Dim FileType As VbFileAttribute
   
   Me.cmbTestCat.Clear
   Root$ = msRoot
   If Not Right(msRoot, 1) = "\" Then Root$ = msRoot & "\"
   TestDir$ = Root$ & msTestDir & "\"
   Cat$ = Dir(TestDir$, vbDirectory)
   Do While Cat$ <> ""
      Cat$ = Dir()
      FileType = GetAttr(TestDir$ & Cat$)
      TypeOfFile = FileType And &HFF
      If (TypeOfFile = vbDirectory) And _
      (Cat$ <> "..") And (Cat$ <> "") Then
         Me.cmbTestCat.AddItem Cat$
      End If
   Loop
   If cmbTestCat.ListCount > 0 Then
      cmbTestCat.ListIndex = 0
      SetTestList
   End If

End Sub

Private Sub SetTestList()

   Dim FileType As VbFileAttribute
   
   Me.cmbTest.Clear
   Root$ = msRoot
   If Not Right(msRoot, 1) = "\" Then Root$ = msRoot & "\"
   TestDir$ = Root$ & msTestDir & "\" & Me.cmbTestCat.Text & "\"
   Cat$ = Dir(TestDir$, vbDirectory)
   Do While Cat$ <> ""
      Cat$ = Dir()
      FileType = GetAttr(TestDir$ & Cat$)
      If (FileType = vbDirectory) And _
      (Cat$ <> "..") And (Cat$ <> "") Then
         Me.cmbTest.AddItem Cat$
      End If
   Loop
   If cmbTest.ListCount > 0 Then cmbTest.ListIndex = 0
   UpdateTest

End Sub

Private Sub chkSE_Click(Index As Integer)

   SEVal = chkSE(Index).value
   ItemName$ = chkSE(Index).Caption
   If Index = 0 Then
      If ItemName$ = "Use Single-ended input" Then _
      VarSet% = frmScript.SetVariable("usesemode", SEVal)
   End If
   If InStr(1, ItemName$, "PORT") > 0 Then
      ItemNum$ = Format(Index, "0")
      VarSet% = frmScript.SetVariable("testport" & ItemNum$, SEVal)
   End If

End Sub

Private Sub chkSE_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
      NumItems& = chkSE.Count - 1
      StateSet% = chkSE(Index).value
      SetState% = 0
      If StateSet% = 0 Then SetState% = 1
      For CheckBoxNum& = 0 To NumItems&
         chkSE(CheckBoxNum&).value = SetState%
      Next
   End If

End Sub

Private Sub chkTest_Click(Index As Integer)

   If frmTestOptions.Visible Then
      ValToSet% = (chkTest(Index).value = 1)
      frmScript.SetTestOptValue Index, ValToSet%
   End If
   
End Sub

Private Sub chkTest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
      NumItems& = chkTest.Count - 1
      StateSet% = chkTest(Index).value
      SetState% = 0
      If StateSet% = 0 Then SetState% = 1
      For CheckBoxNum& = 0 To NumItems&
         chkTest(CheckBoxNum&).value = SetState%
      Next
   End If
   
End Sub

Private Sub chkUseDF_Click()

   DFVal = chkUseDF.value
   VarSet% = frmScript.SetVariable("usedaqflex", DFVal)

End Sub

Private Sub cmbGroup_Click()

   mnGroupIndex = Me.cmbGroup.ListIndex
   UpdateProdList
   
End Sub

Private Sub cmbParamSelect_Change(Index As Integer)

   ParamVal = cmbParamSelect(Index).Text
   ParamName$ = "param" & cmbParamSelect(Index).Tag
   VarSet% = frmScript.SetVariable(ParamName$, ParamVal)
   
End Sub

Private Sub cmbParamSelect_Click(Index As Integer)

   ParamVal = cmbParamSelect(Index).Text
   ParamName$ = "param" & cmbParamSelect(Index).Tag
   VarSet% = frmScript.SetVariable(ParamName$, ParamVal)
   
End Sub

Private Sub cmbProduct_Change()

   If InStr(1, cmbProduct.Text, " ") > 0 Then
      Me.cmbSubProduct.Visible = True
      UpdateSubProdList
   Else
      Me.cmbSubProduct.Visible = False
      UpdateButton
   End If
   
End Sub

Private Sub cmbProduct_Click()

   If InStr(1, cmbProduct.Text, " ") > 0 Then
      Me.cmbSubProduct.Visible = True
      UpdateSubProdList
   Else
      Me.cmbSubProduct.Visible = False
      UpdateButton
   End If

End Sub

Private Sub cmbProduct_Validate(Cancel As Boolean)

   mbDevSelected = True

End Sub

Private Sub cmbSubProduct_Click()

   UpdateButton
   
End Sub

Private Sub cmbSubProduct_Validate(Cancel As Boolean)
   
   mbDevSelected = True

End Sub

Private Sub cmbTest_Click()

   UpdateTest
   
End Sub

Private Sub cmbTestCat_Click()

   SetTestList
   UpdateTest
   
End Sub

Private Sub cmdCancel_Click()

   lblDParmPath.Caption = ""
   DoEvents
   Me.Hide
   
End Sub

Private Sub cmdOpen_Click()

   Dim response As VbMsgBoxResult
   
   PGroup$ = Me.cmbGroup.Text
   Prod$ = Me.cmbProduct.Text
   If mbUseDropCommand And Not mbDevSelected Then
      response = MsgBox("Continue using " & Prod$ & _
      " as the test device?", vbYesNo, "Confirm Test Device")
      If response = vbNo Then Exit Sub
   End If
   If cmbSubProduct.Visible Then SubProd$ = Me.cmbSubProduct.Text
   TGroup$ = Me.cmbTestCat.Text
   Test$ = Me.cmbTest.Text
   
   lpFileName$ = "UniTest.ini"
   lpApplicationName$ = "TestOptions"
   lpKeyName$ = "ProductGroup"
   Param$ = PGroup$
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, Param$, lpFileName$)
   lpKeyName$ = "Product"
   Param$ = Prod$
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, Param$, lpFileName$)
   lpKeyName$ = "SubProduct"
   Param$ = SubProd$
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, Param$, lpFileName$)
   lpKeyName$ = "TestGroup"
   Param$ = TGroup$
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, Param$, lpFileName$)
   Param$ = Test$
   lpKeyName$ = "TestName"
   x% = WritePrivateProfileString(lpApplicationName$, lpKeyName$, Param$, lpFileName$)
   Me.Hide

End Sub

Private Sub Form_Activate()

   Dim ComboControl As Control
   
   If Me.fraSetup.Visible Then
      lpFileName$ = "UniTest.ini"
      lpApplicationName$ = "TestOptions"
      lpKeyName$ = "ProductGroup"
      lpDefault$ = ""
      nSize% = 36
      StdParam$ = Space$(nSize%)
      For cBox% = 1 To 5
         If Not ((cBox% > 3) And mbUseDropCommand) Then
            StdParam$ = Space$(nSize%)
            Set ComboControl = Choose(cBox%, cmbGroup, cmbProduct, cmbSubProduct, cmbTestCat, cmbTest)
            lpKeyName$ = Choose(cBox%, "ProductGroup", "Product", "SubProduct", "TestGroup", "TestName")
            StringSize% = GetPrivateProfileString(lpApplicationName$, _
            lpKeyName$, lpDefault$, StdParam$, nSize%, lpFileName$)
            StdParam$ = Left$(StdParam$, StringSize%)
            SetSelection ComboControl, StdParam$
         End If
      Next
   End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Then
      If Me.fraSetup.Visible Then
         Me.cmdOpen = True
      Else
         Me.Hide
      End If
   End If
   
End Sub

Private Sub Form_Load()
   
   Me.Top = mfmUniTest.Top + 800
   Me.Left = mfmUniTest.Left + 500

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   LoadedControls& = chkSE.Count - 1
   For SECtl& = 0 To LoadedControls&
      SEVal& = chkSE(SECtl&).value
      If Not (SEVal& = 0) Then
         ItemNum$ = Format(SECtl&, "0")
         PortArray$ = PortArray$ & ItemNum$ & ";"
      End If
   Next
   If Not (PortArray$ = "") Then PortArray$ = Left(PortArray$, Len(PortArray$) - 1)
   VarSet% = frmScript.SetVariable("portarray", PortArray$)

End Sub

Private Sub Form_Resize()

   If Not Me.fraSetup.Visible Then
      Me.SSTab1.Width = Me.Width - 390
      Me.fraTestBlocks.Width = SSTab1.Width - 330
   End If
   
End Sub

Private Sub txtHighChan_Change()

   ChanVal = txtHighChan.Text
   VarSet% = frmScript.SetVariable("highchan", ChanVal)

End Sub

Private Sub txtLowChan_Change()

   ChanVal = txtLowChan.Text
   VarSet% = frmScript.SetVariable("lowchan", ChanVal)
   
End Sub

Private Sub txtParam_Change(Index As Integer)

   ParamVal = txtParam(Index).Text
   ParamName$ = "param" & txtParam(Index).Tag
   VarSet% = frmScript.SetVariable(ParamName$, ParamVal)

End Sub

Sub UpdateProdList()

   Dim FileType As VbFileAttribute
   Me.cmbProduct.Clear
   Root$ = msRoot
   If Not Right(msRoot, 1) = "\" Then Root$ = msRoot & "\"
   ProdDir$ = Root$ & cmbGroup.Text & "\"
   Prod$ = Dir(ProdDir$, vbDirectory)
   Do While Prod$ <> ""
      Prod$ = Dir()
      FileType = GetAttr(ProdDir$ & Prod$)
      If (FileType = vbDirectory) And _
      (Prod$ <> "..") And (Prod$ <> "") Then
         Me.cmbProduct.AddItem Prod$
      End If
   Loop
   If cmbProduct.ListCount > 0 Then cmbProduct.ListIndex = 0
   UpdateButton
   
End Sub

Sub UpdateSubProdList()

   Dim FileType As VbFileAttribute
   Me.cmbSubProduct.Clear
   Root$ = msRoot
   If Not Right(msRoot, 1) = "\" Then Root$ = msRoot & "\"
   ProdDir$ = Root$ & cmbGroup.Text & _
   "\" & cmbProduct.Text & "\"
   Prod$ = Dir(ProdDir$, vbDirectory)
   Do While Prod$ <> ""
      Prod$ = Dir()
      FileType = GetAttr(ProdDir$ & Prod$)
      If (FileType = vbDirectory) And _
      (Prod$ <> "..") And (Prod$ <> "") Then
         Me.cmbSubProduct.AddItem Prod$
      End If
   Loop
   If cmbSubProduct.ListCount > 0 Then cmbSubProduct.ListIndex = 0
   UpdateButton
   
End Sub

Sub UpdateButton()

   lblDParmPath.Caption = ""
   ParamPath$ = msRoot
   CatPath$ = cmbGroup.Text & "\"
   ProdPath$ = Me.cmbProduct.Text & "\"
   If cmbSubProduct.Visible Then
      If Not (cmbSubProduct.Text = "") Then
         SubPath$ = Me.cmbSubProduct.Text & "\"
      End If
   Else
      SubPath$ = ""
   End If
   PDir$ = ParamPath$ & CatPath$ & ProdPath$ & SubPath$ & "DeviceParams.uts"
   FileFound$ = Dir(PDir$, vbNormal)
   ValidFile% = Not (FileFound$ = "")
   cmdOpen.Enabled = ValidFile%
   If ValidFile% Then lblDParmPath.Caption = PDir$
   
End Sub

Sub UpdateTest()

   RootDir$ = msRoot & msTestDir & "\"
   CatDir$ = cmbTestCat.Text & "\"
   If Not (cmbTest.Text = "") Then TestDir$ = cmbTest.Text & "\"
   TPath$ = RootDir$ & CatDir$ & TestDir$
   'If mbUseDropCommand Then
   TFile$ = Dir(TPath$ & "*.utm", vbNormal)
   If TFile$ = "" Then TFile$ = Dir(TPath$ & "*.uss", vbNormal)
   'Else
   '   TFile$ = Dir(TPath$ & "*.utm", vbNormal)
   'End If
   ValidFile% = Not (TFile$ = "")
   If ValidFile% Then lblTestFile.Caption = TPath$ & TFile$
   Me.lblFileName.Caption = TFile$
   
End Sub

Public Sub SetSelection(ByVal ComboControl As ComboBox, ByVal StdParam As String)

   For lstItem% = 0 To ComboControl.ListCount - 1
      If ComboControl.List(lstItem%) = StdParam Then ComboControl.ListIndex = lstItem%
   Next
   
End Sub

Public Sub UseDropCommand(ByVal TrueFalse As Boolean)

   mbUseDropCommand = TrueFalse
   mbDevSelected = False
   
End Sub
