VERSION 5.00
Begin VB.Form fLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login User"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox t2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   720
      Width           =   2385
   End
   Begin VB.TextBox t1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2385
   End
   Begin VB.CommandButton c2 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton c1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   780
      Width           =   945
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   945
   End
End
Attribute VB_Name = "fLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xkali As Integer

Private Sub c1_Click()
On Error GoTo aa
If xkali = 3 Then
  End
Else
  xkali = xkali + 1
End If

Dim oRS As New ADODB.Recordset
Set oRS = aData.AmbilCommand("Select * from Pengguna where nama='" & _
AmanOi(t1.Text) & "'")
If oRS.RecordCount = 1 Then
  If oRS!Pass = RegC.HashString(t2.Text) Then
    nAMA = t1.Text
    Unload fLog
    Utama.Caption = "Asia Baru " & "- " & LokasiFile
    Utama.Show
    Unload fLog
  Else
    GoTo trl
  End If
Else
trl:
  MsgBox "User name atau password salah..", vbCritical, "Login"
  t1.Text = "": t2.Text = "": t1.SetFocus
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub c2_Click()
On Error Resume Next
End
End Sub

Private Sub Form_Load()
On Error Resume Next
xkali = 0
End Sub
