VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   2310
   ClientLeft      =   5640
   ClientTop       =   4560
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1364.824
   ScaleMode       =   0  'User
   ScaleWidth      =   5830.854
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   225
      Width           =   2370
   End
   Begin VB.TextBox Text1 
      Height          =   870
      Left            =   2610
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2025
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3825
      TabIndex        =   2
      Top             =   180
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4905
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   765
      Width           =   5760
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   1080
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginOK As Boolean
Public Jenis As String
Private Sub cmdCancel_Click()
    LoginOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo aa

Dim rsPass As New ADODB.Recordset, aPass As String
Set rsPass = aData.AmbilCommand("select * from xPass")
If rsPass.EOF Then
Unload Me
Exit Sub
End If
If Jenis = "Faktur" Then
aPass = rsPass!PassFaktur
ElseIf Jenis = "Cust" Then
aPass = rsPass!PassCust
Else
aPass = rsPass!PassRugi
End If
If RegC.HashString(TxtPassword.Text) = aPass Then
Unload Me
LoginOK = True
Else
MsgBox "Password yang anda masukkan salah", vbInformation, "Password"
Unload Me
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_Load()
On Error GoTo aa
LoginOK = False
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

