VERSION 5.00
Begin VB.Form frmConvertFile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Convert File Settings"
   ClientHeight    =   3015
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   6795
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3015
   ScaleWidth      =   6795
   Tag             =   "frmConvertFile"
   Begin VB.TextBox txtStart 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   13
      Text            =   "0"
      Top             =   1860
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4740
      TabIndex        =   12
      Top             =   2460
      Width           =   915
   End
   Begin VB.TextBox txtCount 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   2220
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   435
      Left            =   5760
      TabIndex        =   9
      Top             =   2460
      Width           =   915
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   2220
      TabIndex        =   8
      Top             =   180
      Width           =   1875
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   1875
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   1860
      Width           =   1155
   End
   Begin VB.TextBox txtDestFile 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   2580
      Width           =   4335
   End
   Begin VB.Frame fraConvert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delimiters"
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   4740
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      Begin VB.OptionButton optDelimiter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Comma"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDelimiter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Semicolon"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.OptionButton optDelimiter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Space"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   1275
      End
      Begin VB.OptionButton optDelimiter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tab"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   1275
      End
   End
   Begin VB.Label lblStart 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Starting Sample"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1740
      TabIndex        =   14
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sample Count"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2280
      Width           =   1515
   End
End
Attribute VB_Name = "frmConvertFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

   txtDestFile.Text = "Cancel"
   Unload Me

End Sub

Private Sub cmdOK_Click()

   Me.Hide

End Sub

Private Sub Dir1_Change()
   
   FilePath$ = Dir1.Path
   File1.Path = Dir1.Path   ' Set file path.
   txtDestFile.Text = FilePath$

End Sub

Private Sub Drive1_Change()
On Error GoTo NoDrive

   Dir1.Path = Drive1.Drive ' Set directory path.

   Exit Sub

NoDrive:
   MsgBox Error$(Err), , "Error Selecting Drive"
   Resume Next

End Sub

Private Sub File1_Click()
   
   FilePath$ = Dir1.Path
   'txtDestFile.Text = FilePath$
   If Not (Len(FilePath$) = 0) Then
      If Not (Right$(FilePath$, 1) = "\") Then FilePath$ = Dir1.Path & "\"
      Filename$ = FilePath$ & File1.Filename
      txtDestFile.Text = Filename$
   Else
      MsgBox "No file path", , "No Path"
   End If

End Sub

Private Sub Form_Load()

   Me.Left = Screen.Width / 2 - Me.Width / 2
   Me.Top = Screen.Height / 3
   txtDestFile.Text = Dir1.Path

End Sub

