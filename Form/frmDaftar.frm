VERSION 5.00
Begin VB.Form frmDaftar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "License"
   ClientHeight    =   1530
   ClientLeft      =   5655
   ClientTop       =   5460
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   903.974
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.983
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1605
      TabIndex        =   3
      Top             =   510
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nomor Pass : "
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nomor Registrasi : "
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   570
      Width           =   1530
   End
End
Attribute VB_Name = "frmDaftar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RegA As New Registri.RegistrySetting
Dim RegB As New Registri.GetDiscID
Dim RegC As New Registri.Enkripsi
Dim k As String, l As String, Coba As Boolean

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
RegA.SaveSettingString HKEY_LOCAL_MACHINE, "Software\DataX Active Object", "DataNya", txtPassword.Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
txtUserName.Text = RegB.BacaDrive(App.Path)
txtPassword.SetFocus
End Sub
