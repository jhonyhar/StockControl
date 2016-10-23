VERSION 5.00
Begin VB.Form fBackUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Back Up Data"
   ClientHeight    =   2625
   ClientLeft      =   4425
   ClientTop       =   3885
   ClientWidth     =   5490
   ClipControls    =   0   'False
   Icon            =   "fBackUp.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1811.822
   ScaleMode       =   0  'User
   ScaleWidth      =   5155.393
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox t 
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   1230
      Width           =   4620
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
      Left            =   4890
      TabIndex        =   2
      Top             =   1215
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Back Up"
      Default         =   -1  'True
      Height          =   465
      Left            =   3840
      TabIndex        =   0
      Top             =   1950
      Width           =   1500
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   5040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama File : "
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   960
      Width           =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   140.858
      X2              =   4972.278
      Y1              =   1200.979
      Y2              =   1200.979
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   140.858
      X2              =   4972.278
      Y1              =   1221.686
      Y2              =   1221.686
   End
End
Attribute VB_Name = "fBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CD1 As New myComDialog.cComDialog

Private Sub cmdOK_Click()
      Dim NamaFile As String, DiMana As String, i As Integer
5     On Error GoTo aa
10    DiMana = RegC.Dekrip(RegA.GetSettingString(HKEY_LOCAL_MACHINE, "Software\DataX Active Object\Service", "DataLoc", "2"))
15      If Right(DiMana, 1) = "\" Then
20      DiMana = Left(DiMana, Len(DiMana) - 1)
25      End If
30    CopyFile LokasiFile, t.Text
35    MsgBox "Data telah selesai dibackup ke " & t.Text, vbInformation, "Backup Data"
40      Unload Me
45    Exit Sub
aa:
      Dim Err_Setering As String
50    Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur cmdOK_Click pada " & "Form fBackUp di baris " & Erl
55    Select Case MsgBox(Err_Setering, vbRetryCancel, App.Title & "-fBackUp Error")
        Case vbCancel: Resume Exit_cmdOK_Click:
60      Case vbRetry: Resume
65      Case Else: End
70    End Select
Exit_cmdOK_Click:
75    App.LogEvent "myAS=>" & Format(Date, "dd-mmmm-yyyy") & Format(Time, "(hh:mm:ss)") & _
      vbCrLf & Err_Setering & vbCrLf, vbLogEventTypeError

End Sub

Private Sub Command1_Click()
On Error Resume Next
CD1.CancelError = True
CD1.DialogTitle = "Pilih nama file backup :"
CD1.Filter = "Database File|*.mdb"
CD1.FilterIndex = 1
CD1.InitDir = IIf(t.Text = "", App.Path, t.Text)
CD1.Flags = 2048 Or 4096
CD1.FileName = Format(Date, "ddmmmyyyy") & Format(Time, "hhmm")
CD1.ShowSave
If Err.Number = 0 Then t.Text = CD1.FileName & ".mdb"
End Sub

Private Sub Form_Load()
    lblDescription.Caption = "Tool ini berfungsi untuk membackup data ke media lain yang dapat dipergunakan kembali sewaktu-waktu bila terjadi sesuatu.." & _
    vbCrLf & "Pilih nama file sebagai tujuan backup, kemudian tekan tombol backup.."
End Sub

