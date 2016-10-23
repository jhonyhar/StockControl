VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   myStock"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   600
      Top             =   1080
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Special license to Asia Baru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "myStock"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   2880
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3030
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
On Error Resume Next
Call OKOK
End Sub

Private Sub Form_DblClick()
On Error Resume Next
Call OKOK
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then Call OKOK
End Sub

Private Sub Form_Load()
On Error GoTo aa
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Timer1_Timer()
Call OKOK
Timer1.Enabled = False
End Sub

Private Sub OKOK()
On Error Resume Next
Unload frmSplash
fLog.Show
Unload frmSplash
End Sub
