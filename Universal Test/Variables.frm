VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVariables 
   Caption         =   "Current Variables"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdVGrid 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvVarData As Variant

Public Sub LoadGrid(ByVal GridData As Variant)

   mvVarData = GridData

End Sub

Private Sub cmdCopy_Click()

    Clipboard.Clear
    Clipboard.SetText grdVGrid.Clip
    
End Sub

Private Sub cmdSort_Click()

   If cmdSort.Caption = "Sort" Then
      grdVGrid.Col = 1
      grdVGrid.Sort = 1
      cmdSort.Caption = "UnSort"
   Else
      cmdSort.Caption = "Sort"
      ReLoadGrid
   End If
End Sub

Private Sub Form_Activate()

   ReLoadGrid

End Sub

Private Sub Form_Load()

   grdVGrid.Col = 0
   grdVGrid.Row = 0
   grdVGrid.Text = "VarNum"
   grdVGrid.ColWidth(0) = 700
   grdVGrid.ColWidth(1) = 2000
   grdVGrid.ColWidth(2) = 4000
   grdVGrid.Col = 1
   grdVGrid.Text = "Name"
   grdVGrid.Col = 2
   grdVGrid.Text = "Value"
   grdVGrid.ColAlignment(1) = flexAlignLeftCenter
   grdVGrid.ColAlignment(2) = flexAlignLeftCenter
   
End Sub

Private Sub Form_Resize()

   grdVGrid.Width = Me.Width - 400
   SBarComp& = 0
   If grdVGrid.Rows > 10 Then SBarComp& = 505
   WidthCol0! = grdVGrid.ColWidth(0) + SBarComp&
   grdVGrid.ColWidth(1) = (grdVGrid.Width - WidthCol0!) * 0.3
   grdVGrid.ColWidth(2) = (grdVGrid.Width - WidthCol0!) * 0.68
   
End Sub

Sub ReLoadGrid()

   On Error GoTo NoData
   
   grdVGrid.Clear
   NumRows& = grdVGrid.Rows
   NumVars& = UBound(mvVarData, 2)
   AddRows% = (NumRows& < NumVars&)
   For GridRow& = 1 To NumVars& + 1
      If AddRows% Then
         grdVGrid.AddItem Format(GridRow&, "0"), GridRow&
      Else
         grdVGrid.Col = 0
         grdVGrid.Row = GridRow&
         grdVGrid.Text = Format(GridRow&, "0")
      End If
      grdVGrid.Col = 1
      grdVGrid.Row = GridRow&
      grdVGrid.Text = mvVarData(0, GridRow& - 1)
      grdVGrid.Col = 2
      grdVGrid.Text = mvVarData(1, GridRow& - 1)
   Next
   Exit Sub
   
NoData:
   Exit Sub
   
End Sub

Private Sub grdVGrid_Click()

    If grdVGrid.RowSel > 0 Then cmdCopy.Enabled = True
    
End Sub

Private Sub grdVGrid_SelChange()

    cmdCopy.Enabled = (grdVGrid.RowSel > 0)
    
End Sub
