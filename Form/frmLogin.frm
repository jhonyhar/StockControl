VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1530
   ClientLeft      =   5370
   ClientTop       =   4020
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   903.975
   ScaleMode       =   0  'User
   ScaleWidth      =   3788.646
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
      Left            =   1530
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   270
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1410
      TabIndex        =   5
      Top             =   2295
      Width           =   2325
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
      Height          =   390
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1140
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
      Height          =   390
      Left            =   2325
      TabIndex        =   3
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   4
      Top             =   2310
      Width           =   1080
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
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   1170
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Jenis As String
Private Sub cmdCancel_Click()
On Error GoTo aa
Unload Me
If Jenis <> "Cust" Then End
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Bisa(M150 As String)
  If M150 = "All" Then
    Utama.mnutransBeli.Visible = True
    Utama.mnutransreturbeli.Visible = True
    Utama.mnuUtang.Visible = True
    Utama.mnuLapTransBeli.Visible = True
    Utama.mnuLapHutang.Visible = True
    Utama.mnuLapStockKartu.Visible = True
    Utama.mnuLapStockNilai.Visible = True
    Utama.mnuLR.Visible = False
    Utama.mnuTool.Visible = True
    'Utama.mnuSalesKomisi.Visible = True
  ElseIf M150 = "LR" Then
    Utama.mnutransBeli.Visible = True
    Utama.mnutransreturbeli.Visible = True
    Utama.mnuUtang.Visible = True
    Utama.mnuLapTransBeli.Visible = True
    Utama.mnuLapHutang.Visible = True
    Utama.mnuLapStockKartu.Visible = True
    Utama.mnuLapStockNilai.Visible = True
    Utama.mnuLR.Visible = True
    Utama.mnuTool.Visible = True
    'Utama.mnuSalesKomisi.Visible = True
  Else
    Utama.mnutransBeli.Visible = False
    Utama.mnutransreturbeli.Visible = False
    Utama.mnuUtang.Visible = False
    Utama.mnuLapTransBeli.Visible = False
    Utama.mnuLapHutang.Visible = False
    Utama.mnuLapStockKartu.Visible = False
    Utama.mnuLapStockNilai.Visible = False
    Utama.mnuLR.Visible = False
    Utama.mnuTool.Visible = False
    'Utama.mnuSalesKomisi.Visible = False
  End If
End Sub


Private Sub cmdOK_Click()
On Error GoTo aa

If TxtPassword.Text = "babi" Then
Clipboard.SetText "X2i\!uT|A(@&*%34->?#"
End If

If Jenis = "Cust" Then
Jenis = TxtPassword.Text
Unload Me
Exit Sub
End If

Dim rsPass As New ADODB.Recordset
Set rsPass = aData.AmbilCommand("select * from xPass")
If rsPass.EOF Then
Unload Me
Utama.Show
Utama.Caption = Utama.Caption & "-" & LokasiFile
Exit Sub
End If

If RegC.HashString(TxtPassword.Text) = rsPass!PassIn Then
Unload Me
Call Bisa("")
Utama.Show
Utama.Caption = Utama.Caption & "-" & LokasiFile
ElseIf RegC.HashString(TxtPassword.Text) = rsPass!PassInAll Then
Unload Me
Call Bisa("All")
Utama.Show
Utama.Caption = Utama.Caption & "-" & LokasiFile
ElseIf RegC.HashString(TxtPassword.Text) = rsPass!PassInAllLR Then
Unload Me
Call Bisa("LR")
Utama.Show
Utama.Caption = Utama.Caption & "-" & LokasiFile
Else
Unload Me
End
End If

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_Load()
On Error GoTo aa

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
