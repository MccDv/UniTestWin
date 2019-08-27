VERSION 5.00
Begin VB.Form frmDiscovery 
   Caption         =   "Device Discovery and Removal"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRemoveUndisc 
      Caption         =   "Remove undiscovered devices"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkCreate 
      Caption         =   "Create device during discovery"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.Frame fraDevManager 
      Caption         =   "Add or Remove Devices"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   6375
      Begin VB.CheckBox chkInterface 
         Caption         =   "Demo Board"
         Height          =   195
         Index           =   4
         Left            =   2940
         TabIndex        =   15
         Top             =   600
         Width           =   1275
      End
      Begin VB.CheckBox chkInterface 
         Caption         =   "All"
         Height          =   195
         Index           =   3
         Left            =   4200
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkInterface 
         Caption         =   "Ethernet"
         Height          =   195
         Index           =   2
         Left            =   1740
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chkInterface 
         Caption         =   "Bluetooth"
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkInterface 
         Caption         =   "USB"
         Height          =   195
         Index           =   0
         Left            =   2940
         TabIndex        =   11
         Top             =   360
         Width           =   795
      End
      Begin VB.CommandButton cmdCreateDev 
         Caption         =   "Create Device"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         TabIndex        =   10
         Top             =   600
         Width           =   1300
      End
      Begin VB.CommandButton cmdDiscover 
         Caption         =   "Discover Devices"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Right click for manual network discovery."
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdReleaseDev 
         Caption         =   "Release Device"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   1300
      End
      Begin VB.Label lblStatus 
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   5235
      End
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   4020
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveConfig 
      Caption         =   "Save Config"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4020
      Width           =   1455
   End
   Begin VB.TextBox txtBoardInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3555
   End
   Begin VB.ListBox lstInstalledDevs 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmDiscovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnLibType As Integer
Dim mlBoardNumber As Long
Dim mfNoForm As Form

Private Sub chkInterface_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   If (Index = 3) Then
      If chkInterface(3).value = 1 Then
         For CheckItem% = 0 To 2
            chkInterface(CheckItem%).value = 0
         Next
      End If
   Else
      AllValue% = 1
      For CheckItem% = 0 To 2
         If chkInterface(CheckItem%).value = 1 Then AllValue% = 0
      Next
      chkInterface(3).value = AllValue%
   End If
   
End Sub

Private Sub cmdCreateDev_Click()

   Dim Inventory As DaqDeviceDescriptor

   'CurBoardNum& = GetFirstAvailableBoardNum(os)
   BoardNumber& = lstInstalledDevs.ItemData(lstInstalledDevs.ListIndex)
   If BoardNumber& > 99 Then
      ItemIndex& = BoardNumber& - 100
      Inventory = GetInventoryItem(ItemIndex&, "")
      CurBoardNum& = GetFirstAvailableBoardNum(0)
      AddBoardToLibrary CurBoardNum&, Inventory
      RawName$ = Inventory.ProductName
      ConvName$ = StrConv(RawName$, vbUnicode)
      NewName = Left(ConvName$, InStr(1, ConvName$, Chr(0)) - 1)
      lstInstalledDevs.ItemData(lstInstalledDevs.ListIndex) = CurBoardNum&
      'RefreshList False
      PrintMain NewName & " added as board " & Format(CurBoardNum&, "0")
      lstInstalledDevs.ListIndex = -1
   End If
   
End Sub

Private Sub cmdDiscover_LostFocus()

   Me.lblStatus.Caption = ""
   
End Sub

Private Sub cmdDiscover_MouseUp(Button As Integer, _
   Shift As Integer, x As Single, y As Single)

   Dim HostString As String
   Dim HostPort As Long, Timeout As Long
   Dim DevsFound As Long
   Dim CreateDev As Boolean, AddDemo As Boolean
   
   CreateDev = (Me.chkCreate.value = 1)
   If Button = vbRightButton Then
      chkInterface(2).value = 1
      Timeout = 5000
      HostString = InputBox("Enter host name or IP address:", _
         "Add Remote Device", "173.76.198.250")
      HP = InputBox("Enter host port:", "Add Remote Device", "54211")
      HostPort = Val(HP)
      
      If (HostString = "") Or (HP = "") Then Exit Sub
      Me.MousePointer = vbHourglass
      DoEvents
      DevsFound = DiscoverDevices(ETHERNET_IFC, _
         CreateDev, HostString, HostPort, Timeout)
      Me.MousePointer = vbDefault
   Else
      Dim Ifc As DaqDeviceInterface
      For IfcSelect = 0 To 2
         If chkInterface(IfcSelect).value Then
            Ifc = Ifc Or Choose(IfcSelect + 1, USB_IFC, _
               BLUETOOTH_IFC, ETHERNET_IFC)
         End If
      Next
      If chkInterface(3).value = 1 Then _
         Ifc = ANY_IFC
      If chkInterface(4).value = 1 Then AddDemo = True
      Me.MousePointer = vbHourglass
      DoEvents
      DevsFound = DiscoverDevices(Ifc, CreateDev, , , , AddDemo)
      Me.MousePointer = vbDefault
   End If
      
   RefreshList Not CreateDev
   DevString$ = " devices"
   If DevsFound = 1 Then DevString$ = " device"
   Me.lblStatus.Caption = DevsFound & _
      DevString$ & " discovered."

End Sub

Private Sub cmdDone_Click()

   Me.Hide
   
End Sub

Private Sub cmdReleaseDev_Click()
   
   RemoveBoardFromLibrary mlBoardNumber
   RefreshList False
   
End Sub

Private Sub cmdSaveConfig_Click()

   Filename$ = "cb.cfg"
   UserFileName$ = InputBox("Configuration file name", _
      "File Name", Filename$)
   ULStat& = cbSaveConfig(UserFileName$)
   If SaveFunc(Me, SaveConfig, ULStat&, UserFileName$, A2, _
      A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then Exit Sub

   NumFuncs% = GetHistory() - 1
   ReDim MyArray(NumFuncs%, 14)
   GetHistoryArray MyArray()
   PrintMain "cbSaveConfig() = " & ULStat

End Sub

Private Sub Form_Load()

   mnLibType = UNILIB
   RefreshList False
   
End Sub

Private Sub lstInstalledDevs_Click()

   Dim DevInterface As DaqDeviceInterface
   Dim Inventory As DaqDeviceDescriptor
   Dim MatchFound As Boolean
   
   If lstInstalledDevs.ListIndex < 0 Then
      txtBoardInfo.Text = ""
      Me.cmdCreateDev.ENABLED = False
      Exit Sub
   End If
   BoardNumber& = lstInstalledDevs.ItemData(lstInstalledDevs.ListIndex)
   If BoardNumber& > 99 Then
      ItemIndex& = BoardNumber& - 100
      Inventory = GetInventoryItem(ItemIndex&, "")
      BoardInInventory& = cbGetBoardNumber(Inventory)
      Me.txtBoardInfo.Text = "Board number: " & _
         vbTab & Format(BoardInInventory&, "0")
      cmdCreateDev.ENABLED = (chkCreate.value = 0)
      cmdReleaseDev.ENABLED = False
      MatchFound = True
   Else
      cmdCreateDev.ENABLED = False
      mlBoardNumber = BoardNumber&
      Me.txtBoardInfo.Text = "Board number: " _
         & vbTab & Format(mlBoardNumber, "0")
      Me.cmdReleaseDev.ENABLED = True
      NumItems& = GetInventorySize
      MatchFound = False
      For i& = 0 To NumItems& - 1
         Inventory = GetInventoryItem(i&, "")
         BoardInInventory& = cbGetBoardNumber(Inventory)
         If BoardInInventory& = mlBoardNumber Then
            MatchFound = True
            Exit For
         End If
      Next
   End If
   If MatchFound Then
      ProductName$ = Inventory.ProductName
      ProductID$ = Hex(Inventory.ProductID)
      UniqueID$ = Inventory.UniqueID
      DevString$ = Inventory.DevString
      DevInterface = Inventory.InterfaceType
      If DevInterface = 0 Then
         InterfaceString$ = "PCI_IFC"
      Else
      InterfaceString$ = Switch( _
         DevInterface = BLUETOOTH_IFC, "BLUETOOTH_IFC", _
         DevInterface = ETHERNET_IFC, "ETHERNET_IFC", _
         DevInterface = USB_IFC, "USB_IFC")
      End If
      For NuidItem% = 7 To 1 Step -1
         NUID$ = NUID$ & Hex(Inventory.NUID(NuidItem%)) & ", "
      Next
      NUID$ = NUID$ & Hex(Inventory.NUID(0))
      Prod$ = NullTermByteToString(ProductName$)
      If (Prod$ = "") And (Inventory.ProductID = 45) Then Prod$ = "DEMO-BOARD"
      Filler$ = ""
      If Len(Prod$) < 8 Then Filler$ = vbTab
      txtBoardInfo.Text = txtBoardInfo.Text & vbCrLf _
         & Prod$ & vbTab & Filler$ _
         & "(type 0x" & ProductID$ & ")" & vbCrLf _
         & "Device string: " & vbTab _
         & NullTermByteToString(DevString$) & vbCrLf _
         & "Unique ID: " & vbTab _
         & NullTermByteToString(UniqueID$) & vbCrLf _
         & "Interface string: " & vbTab & _
         InterfaceString$ & vbCrLf & vbCrLf _
         & "NUID:  " & NUID$
   End If
   If BoardNumber& < 100 Then
      Dim VerFound As String
      For FWType& = 0 To 4
         ConfigLen& = 64
         VerFound = Space(ConfigLen&)
         CurFWType& = Choose(FWType& + 1, VER_FW_MAIN, VER_FW_MEASUREMENT, _
            VER_FW_RADIO, VER_FPGA, VER_FW_MEASUREMENT_EXP)
         ULStat& = cbGetConfigString(BOARDINFO, BoardNumber&, _
            CurFWType&, BIDEVVERSION, VerFound, ConfigLen&)
         If ConfigLen& > 0 Then
            VerFound = Left(VerFound, ConfigLen&)
            FWTypeString$ = Choose(FWType& + 1, _
               "Mfw ", "Ifw ", "Rfw ", "FPGA ", "Xfw ")
            TypeList$ = TypeList$ & FWTypeString$ & VerFound & ", "
         End If
      Next
      If Not (TypeList$ = "") Then
         TypeList$ = Left(TypeList$, Len(TypeList$) - 2)
         txtBoardInfo.Text = txtBoardInfo.Text & _
            vbCrLf & vbCrLf & TypeList$
      End If
   End If
   
End Sub

Private Sub RefreshList(ByVal GetInventory As Boolean)

   Me.lstInstalledDevs.Clear
   If Not GetInventory Then
      For i% = 0 To gnNumBoards - 1
         If gnLibType = MSGLIB Then
            BoardNum% = i%
            BoardName$ = GetNameOfMsgBoard(BoardNum%)
            SplitName = Split(BoardName$, "::")
            DisplayName$ = SplitName(0)
         Else
            BoardNum% = gnBoardEnum(i%)
            BoardName$ = GetNameOfBoard(BoardNum%)
            DisplayName$ = BoardName$
         End If
         lstInstalledDevs.AddItem DisplayName$
         lstInstalledDevs.ItemData _
            (lstInstalledDevs.NewIndex) = BoardNum%
      Next i%
   End If
   If GetInventory Then
      Dim CurInvItem As DaqDeviceDescriptor
      NumItems& = GetInventorySize
      For ItemIndex% = 0 To NumItems& - 1
         CurInvItem = GetInventoryItem(ItemIndex%, "")
         BoardInInventory& = cbGetBoardNumber(CurInvItem)
         ProductName$ = CurInvItem.ProductName
         Prod$ = NullTermByteToString(ProductName$)
         UIDString$ = CurInvItem.UniqueID
         UID$ = NullTermByteToString(UIDString$)
         Me.lstInstalledDevs.AddItem Prod$
         Me.lstInstalledDevs.ItemData(lstInstalledDevs.NewIndex) _
            = 100 + ItemIndex%
      Next
   End If
   
End Sub
