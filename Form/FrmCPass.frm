VERSION 5.00
Begin VB.Form FrmCPass 
   Caption         =   "Ubah Password"
   ClientHeight    =   10005
   ClientLeft      =   1140
   ClientTop       =   1920
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   5430
      TabIndex        =   6
      Top             =   1020
      Width           =   1230
   End
   Begin VB.TextBox t1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2790
      PasswordChar    =   "•"
      TabIndex        =   5
      Top             =   1755
      Width           =   2385
   End
   Begin VB.TextBox t1 
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
      Index           =   1
      Left            =   2790
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   1380
      Width           =   2385
   End
   Begin VB.TextBox t1 
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
      Index           =   0
      Left            =   2790
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   1020
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password &Baru : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   630
      TabIndex        =   2
      Top             =   1440
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Ulangi Password Baru : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   630
      TabIndex        =   4
      Top             =   1785
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password &Lama : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   630
      TabIndex        =   0
      Top             =   1020
      Width           =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ubah Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   210
      TabIndex        =   7
      Top             =   210
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2220
      Left            =   315
      Top             =   360
      Width           =   7395
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   2490
      Left            =   225
      Top             =   210
      Width           =   7665
   End
End
Attribute VB_Name = "FrmCPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSimpan_Click(Index As Integer)
Dim xLama As String
On Error GoTo aa

Dim rsPass As New ADODB.Recordset
zas:
Set rsPass = aData.AmbilCommand("select * from pengguna where nama='" & AmanOi(nAMA) & "'")
If rsPass.EOF Then
Err.Raise vbObjectError + 123, "User Management", "Data user tidak ada.. Kemungkinan telah dihapus"
End If

Select Case Index
Case 0
If t1(1).Text <> t1(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t1(1).SetFocus
  Exit Sub
End If
xLama = rsPass!Pass
If xLama <> RegC.HashString(t1(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t1(0).SetFocus
  Exit Sub
End If
Dim oSRT As String
oSRT = aData.ExecCommand("update pengguna set Pass='" & RegC.HashString(t1(1)) & "' where  nama='" & AmanOi(nAMA) & "'")
If oSRT = "" Then
  MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
  t1(0).Text = "": t1(1).Text = "": t1(2).Text = ""
Else
  MsgBox oSRT, vbInformation, "Ubah Password"
End If
End Select

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Label1_Click(Index As Integer)

End Sub
