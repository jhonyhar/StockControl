VERSION 5.00
Begin VB.Form frmCekLokasi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1260
   ClientLeft      =   3255
   ClientTop       =   5895
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   744.45
   ScaleMode       =   0  'User
   ScaleWidth      =   4154.835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Batal"
      Height          =   420
      Left            =   1890
      TabIndex        =   4
      Top             =   675
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   630
      TabIndex        =   3
      Top             =   675
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3735
      TabIndex        =   2
      Top             =   135
      Width           =   375
   End
   Begin VB.TextBox t 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Data File :"
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   945
   End
End
Attribute VB_Name = "frmCekLokasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo aa
Dim a As New myComDialog.cComDialog
On Error Resume Next
a.CancelError = True
a.DialogTitle = "Pilih file Database :"
a.Filter = "Database File|*.mdb"
a.FilterIndex = 1
a.InitDir = IIf(t.Text = "", App.Path, t.Text)
a.Flags = cdlOFNFileMustExist Or OFN_PATHMUSTEXIST
a.FileName = IIf(t.Text = "", App.Path, t.Text)
a.ShowOpen

If Err.Number = 0 Then t.Text = a.FileName
  
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command1_Click pada Form frmCekLokasi"
End Sub

Private Sub Command2_Click()
On Error GoTo aa
LokasiFile = t.Text
Unload Me
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command2_Click pada Form frmCekLokasi"
End Sub

Private Sub Command3_Click()
On Error GoTo aa
End
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command3_Click pada Form frmCekLokasi"
End Sub
