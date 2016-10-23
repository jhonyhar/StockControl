VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUtangReal 
   Caption         =   "Realisasi Pembayaran Utang"
   ClientHeight    =   9195
   ClientLeft      =   570
   ClientTop       =   1065
   ClientWidth     =   14085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14085
   WindowState     =   2  'Maximized
   Begin VB.TextBox t8 
      DataField       =   "No Faktur"
      Height          =   570
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   7440
      Width           =   6660
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Pembayaran &Ditolak/dibatalkan"
      Height          =   420
      Left            =   10170
      TabIndex        =   4
      Top             =   7470
      Width           =   2850
   End
   Begin VB.CommandButton CmdBaru 
      Caption         =   "&Pembayaran diterima dengan baik"
      Height          =   420
      Left            =   7245
      TabIndex        =   3
      Top             =   7470
      Width           =   2850
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   6300
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   12660
      _cx             =   22331
      _cy             =   11112
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   10013642
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16711134
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14938588
      GridColor       =   -2147483633
      GridColorFixed  =   12579766
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmUtangReal.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   1
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+F4=Close  F4=Bantuan  F5=Refresh"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   8040
      Width           =   6285
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Keterangan : "
      Height          =   255
      Index           =   11
      Left            =   510
      TabIndex        =   1
      Top             =   7185
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Realisasi Pembayaran Utang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   165
      TabIndex        =   6
      Top             =   105
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   8145
      Left            =   285
      Top             =   210
      Width           =   13200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Konsumen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   -3960
      TabIndex        =   5
      Top             =   195
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   8385
      Left            =   165
      Top             =   105
      Width           =   13410
   End
End
Attribute VB_Name = "frmUtangReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intNoFaktur As Single
Private Sub CmdBaru_Click()
On Error GoTo aa
Dim aHasil As String
aHasil = aData.UtangPiutangOK(SerbaGuna.AmanOi(Master.TextMatrix(Master.Row, 8)), SerbaGuna.AmanOi(t8.Text), "Utang")
If aHasil <> "" Then
MsgBox aHasil, vbInformation, Me.Caption & "#Error"
End If
Call aData.UtangPiutangTransOK("Utang")
Call Form_Load
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_Click()
On Error GoTo aa
Dim aHasil As String
aHasil = aData.UtangPiutangBatal(SerbaGuna.AmanOi(Master.TextMatrix(Master.Row, 8)), SerbaGuna.AmanOi(t8.Text), "Utang")
If aHasil <> "" Then
MsgBox aHasil, vbInformation, Me.Caption & "#Error"
End If
Call Form_Load
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
Call Form_Load
End If
End Sub

Private Sub Form_Load()
On Error GoTo aa
Dim rsMaster As New ADODB.Recordset
Set rsMaster = aData.AmbilCommand("SELECT utanganak.intno,Utang.KodeBayar, Utang.tgl, Utang.Kepada+'-'+Supplier.Nama AS Kepada, UtangAnak.Total, IIf(UtangAnak.JenisBayar='T','Cash',IIf(UtangAnak.JenisBayar='G','Giro','Transfer')) AS JenisBayar, UtangAnak.JatuhTempo, UtangAnak.NoGiro, UtangAnak.NamaBank " & _
" FROM (Supplier INNER JOIN Utang ON Supplier.Kode=Utang.Kepada) INNER JOIN UtangAnak ON Utang.intno=UtangAnak.KodeBayar " & _
"WHERE ((([status])='')) " & _
"ORDER BY JatuhTempo, JenisBayar;")
    Master.Rows = 1
    Do While Not rsMaster.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = rsMaster![KodeBayar]
    Master.TextMatrix(Master.Rows - 1, 1) = rsMaster![tgl]
    Master.TextMatrix(Master.Rows - 1, 2) = rsMaster![Kepada]
    Master.TextMatrix(Master.Rows - 1, 3) = rsMaster!Total
    Master.TextMatrix(Master.Rows - 1, 4) = rsMaster!jenisbayar
    Master.TextMatrix(Master.Rows - 1, 5) = rsMaster!namabank
    Master.TextMatrix(Master.Rows - 1, 6) = rsMaster![nogiro]
    Master.TextMatrix(Master.Rows - 1, 7) = rsMaster![jatuhtempo]
    Master.TextMatrix(Master.Rows - 1, 8) = rsMaster!intNo
    rsMaster.MoveNext
    Loop
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub


