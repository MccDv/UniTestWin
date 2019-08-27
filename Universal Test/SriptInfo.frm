VERSION 5.00
Begin VB.Form frmScriptInfo 
   Caption         =   "Script Information"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   2340
      Width           =   675
   End
   Begin VB.PictureBox picMasterStat 
      AutoRedraw      =   -1  'True
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   2340
      Width           =   4095
   End
   Begin VB.PictureBox picScriptStat 
      AutoRedraw      =   -1  'True
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox txtScriptInfo 
      Height          =   2145
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "frmScriptInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()

   txtScriptInfo.Text = ""

End Sub

Private Sub Form_Load()
   
   lpFileName$ = "UniTest.ini"
   lpHKeyName$ = "Height"
   nHDefault% = 3210
   lpWKeyName$ = "Width"
   nWDefault% = 5205
   lpTKeyName$ = "Top"
   nTDefault% = 0
   lpLKeyName$ = "Left"
   nLDefault% = 0
   
   lpApplicationName$ = Me.Name
   Me.Height = GetPrivateProfileInt(lpApplicationName$, lpHKeyName$, nHDefault%, lpFileName$)
   Me.Width = GetPrivateProfileInt(lpApplicationName$, lpWKeyName$, nWDefault%, lpFileName$)
   Me.Top = GetPrivateProfileInt(lpApplicationName$, lpTKeyName$, nTDefault%, lpFileName$)
   Me.Left = GetPrivateProfileInt(lpApplicationName$, lpLKeyName$, nLDefault%, lpFileName$)
   
End Sub

Private Sub Form_Resize()

   Me.txtScriptInfo.Width = Me.Width - 350
   Me.txtScriptInfo.Height = Me.Height - 1100
   Me.picScriptStat.Width = Me.Width - 1400
   Me.picScriptStat.Top = Me.Height - 700
   picMasterStat.Top = Me.picScriptStat.Top - 180
   picMasterStat.Width = Me.picScriptStat.Width
   Me.cmdClear.Left = Me.Width - 1100
   Me.cmdClear.Top = Me.picMasterStat.Top
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   lpFileName$ = "UniTest.ini"
   lpHKeyName$ = "Height"
   nHDefault% = 1815
   lpWKeyName$ = "Width"
   nWDefault% = 8685
   lpTKeyName$ = "Top"
   nTDefault% = 0
   lpLKeyName$ = "Left"
   nLDefault% = 0
   
   If Me.WindowState = 0 Then
      lpApplicationName$ = Me.Name
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
   frmScript.mnuScriptInf.Checked = False

End Sub

Private Sub txtScriptInfo_Change()

   Me.txtScriptInfo.SelStart = Len(Me.txtScriptInfo.Text)
   
End Sub
