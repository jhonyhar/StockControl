VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPiutangReal 
   Caption         =   "Realisasi Pembayaran Piutang"
   ClientHeight    =   9195
   ClientLeft      =   600
   ClientTop       =   1440
   ClientWidth     =   14085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14085
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Pembayaran &Ditolak/dibatalkan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10350
      TabIndex        =   2
      Top             =   7770
      Width           =   2850
   End
   Begin VB.CommandButton CmdBaru 
      Caption         =   "&Pembayaran diterima dengan baik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7425
      TabIndex        =   1
      Top             =   7770
      Width           =   2850
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   6915
      Left            =   450
      TabIndex        =   0
      Top             =   765
      Width           =   13845
      _cx             =   24421
      _cy             =   12197
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPiutangReal.frx":0000
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   7935
      Width           =   6285
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Realisasi Pembayaran Piutang"
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
      TabIndex        =   5
      Top             =   105
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   9855
      Left            =   285
      Top             =   210
      Width           =   14145
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F4=Input, F5=Sort, F6=Filter, F7=Form View, F8=Print, F9=Refresh, F10=Search, Alt+X=Close"
      Height          =   495
      Left            =   420
      TabIndex        =   6
      Top             =   7455
      Width           =   7875
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
      TabIndex        =   4
      Top             =   195
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   10095
      Left            =   165
      Top             =   105
      Width           =   14400
   End
   Begin VB.Label Label1 
      Height          =   795
      Left            =   13170
      TabIndex        =   3
      Top             =   8730
      Width           =   1665
   End
End
Attribute VB_Name = "frmPiutangReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JenisUtPi As String
Dim intNoFaktur As Single
Private Sub CmdBaru_Click()
On Error GoTo aa
Dim aHasil As String, i As Single

For i = 1 To Master.Rows - 1
  If Master.TextMatrix(i, 9) Then
    aHasil = aData.UtangPiutangOK(SerbaGuna.AmanOi(Master.TextMatrix(i, 8)), SerbaGuna.AmanOi(Master.TextMatrix(i, 10)), JenisUtPi)
    If aHasil <> "" Then
    MsgBox aHasil, vbInformation, Me.Caption & "#Error"
    End If
  End If
Next i
Call aData.UtangPiutangTransOK(JenisUtPi)
Call Form_Load
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_Click()
On Error GoTo aa
Dim aHasil As String, i As Single

For i = 1 To Master.Rows - 1
  If Master.TextMatrix(i, 9) Then
    aHasil = aData.UtangPiutangBatal(SerbaGuna.AmanOi(Master.TextMatrix(i, 8)), SerbaGuna.AmanOi(Master.TextMatrix(i, 10)), JenisUtPi)
    If aHasil <> "" Then
    MsgBox aHasil, vbInformation, Me.Caption & "#Error"
    End If
  End If
Next i
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
If JenisUtPi = "Utang" Then
Set rsMaster = aData.AmbilCommand("SELECT UtangAnak.intNo,Utang.KodeBayar, Utang.tgl, Utang.Kepada+'-'+Supplier.Nama AS Kepada, UtangAnak.Total, IIf(UtangAnak.JenisBayar='T','Cash',IIf(UtangAnak.JenisBayar='G','Giro','Transfer')) AS JenisBayar, UtangAnak.JatuhTempo, UtangAnak.NoGiro, UtangAnak.NamaBank " & _
" FROM (Supplier INNER JOIN Utang ON Supplier.Kode=Utang.Kepada) INNER JOIN UtangAnak ON Utang.intno=UtangAnak.KodeBayar " & _
"WHERE ((([status])='')) " & _
"ORDER BY JatuhTempo, JenisBayar;")
Label2.Caption = "Realisasi Pembayaran Utang"
Else
Set rsMaster = aData.AmbilCommand("SELECT PiutangAnak.intNo,Piutang.KodeBayar, Piutang.tgl, Piutang.Kepada+'-'+Konsumen.Nama AS Kepada, PiutangAnak.Total, IIf(PiutangAnak.JenisBayar='T','Cash',IIf(PiutangAnak.JenisBayar='G','Giro','Transfer')) AS JenisBayar, PiutangAnak.JatuhTempo, PiutangAnak.NoGiro, PiutangAnak.NamaBank " & _
" FROM (Konsumen INNER JOIN Piutang ON Konsumen.Kode=Piutang.Kepada) INNER JOIN PiutangAnak ON Piutang.intno=PiutangAnak.KodeBayar " & _
"WHERE ((([status])='')) " & _
"ORDER BY JatuhTempo, JenisBayar;")
Label2.Caption = "Realisasi Pembayaran Piutang"
End If

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
    Master.TextMatrix(Master.Rows - 1, 9) = False
    Master.TextMatrix(Master.Rows - 1, 10) = ""
    rsMaster.MoveNext
    Loop
Master.RowHeight(0) = 350
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub


Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Col
Case 0 To 8
Cancel = True
Case Else
Cancel = False
End Select
End Sub
