VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form fBTandaTerima 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilih data tanda terima penagihan"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCekFaktur 
      Caption         =   "&Proses"
      Height          =   540
      Left            =   270
      TabIndex        =   6
      Top             =   2070
      Width           =   1395
   End
   Begin VB.TextBox t3 
      DataField       =   "No Faktur"
      Height          =   405
      Left            =   2295
      TabIndex        =   0
      Top             =   1095
      Width           =   3435
   End
   Begin TDBDate6Ctl.TDBDate t1 
      Height          =   405
      Left            =   2295
      TabIndex        =   1
      Top             =   240
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calendar        =   "fBTandaTerima.frx":0000
      Caption         =   "fBTandaTerima.frx":0136
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "fBTandaTerima.frx":01A4
      Keys            =   "fBTandaTerima.frx":01C2
      Spin            =   "fBTandaTerima.frx":0220
      AlignHorizontal =   0
      AlignVertical   =   2
      Appearance      =   2
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd mmm yyyy;;Data tidak valid"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   3
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   39191
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate t2 
      Height          =   405
      Left            =   2295
      TabIndex        =   4
      Top             =   675
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calendar        =   "fBTandaTerima.frx":0248
      Caption         =   "fBTandaTerima.frx":037E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "fBTandaTerima.frx":03EC
      Keys            =   "fBTandaTerima.frx":040A
      Spin            =   "fBTandaTerima.frx":0468
      AlignHorizontal =   0
      AlignVertical   =   2
      Appearance      =   2
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd mmm yyyy;;Data tidak valid"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   3
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   39191
      CenturyMode     =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   2310
      TabIndex        =   7
      Top             =   1605
      Width           =   3420
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sampai Tgl:"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   675
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dari Tgl:"
      Height          =   375
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nama Toko :"
      Height          =   375
      Index           =   3
      Left            =   255
      TabIndex        =   2
      Top             =   1125
      Width           =   1815
   End
End
Attribute VB_Name = "fBTandaTerima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Jenis As String

Private Sub CmdCekFaktur_Click()
On Error GoTo aa
  HasilTT = t1.Value & ";;;" & t2.Value & ";;;" & t3.Text
  Unload Me
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_Load()
On Error Resume Next
t1.Value = Date - 30
t2.Value = Date
t3.Text = ""
HasilTT = ""
End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     If Jenis = "H" Then
     Bantu.GridData = "SELECT Supplier.Kode, Supplier.Nama, Supplier.Alamat,Supplier.Wilayah FROM Supplier order by Nama;"
     Else
     Bantu.GridData = "SELECT Kode, Nama, Alamat,Wilayah FROM Konsumen order by Nama"
     End If
     Bantu.Left = t3.Left + Me.Left + 50
     Bantu.Top = t3.Top + Me.Top + t3.Height + 500
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     t3.Text = BarisGrid(0)
     t3_Validate False
     End If
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t3_Validate(Cancel As Boolean)
On Error GoTo aa
Dim rsKodok As New ADODB.Recordset
If Jenis = "H" Then
  Set rsKodok = aData.AmbilCommand("Select * from Supplier where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
   Label6.Caption = IIf(IsNull(rsKodok!Nama), "", rsKodok!Nama) & "(" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & ")"
  End If
Else
  Set rsKodok = aData.AmbilCommand("Select * from Konsumen where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
  Label6.Caption = IIf(IsNull(rsKodok!Nama), "", rsKodok!Nama) & "(" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & ")"
  End If
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur t3_Validate pada Form frmPembelian"
End Sub

