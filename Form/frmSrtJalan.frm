VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSrtJalan 
   Caption         =   "Surat Pengantar Barang"
   ClientHeight    =   8610
   ClientLeft      =   1200
   ClientTop       =   2370
   ClientWidth     =   13830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   13830
   WindowState     =   2  'Maximized
   Begin VB.TextBox t7 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2580
      TabIndex        =   25
      Top             =   2304
      Width           =   2160
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cek SPB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6120
      TabIndex        =   2
      Top             =   675
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2595
      TabIndex        =   19
      Top             =   2700
      Width           =   2160
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   11190
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6075
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6075
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox t8 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   6585
      Width           =   6660
   End
   Begin VB.CommandButton CmdCekFaktur 
      Caption         =   "&Cek Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5100
      TabIndex        =   20
      Top             =   2460
      Width           =   1515
   End
   Begin VB.CommandButton CmdBaru 
      Caption         =   "S&PB Baru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7380
      TabIndex        =   12
      Top             =   6495
      Width           =   1230
   End
   Begin VB.TextBox t4 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2580
      TabIndex        =   8
      Top             =   1908
      Width           =   3435
   End
   Begin VB.TextBox t3 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2580
      TabIndex        =   6
      Top             =   1512
      Width           =   3435
   End
   Begin VB.TextBox t2 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2580
      TabIndex        =   4
      Top             =   1116
      Width           =   3435
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8715
      TabIndex        =   13
      Top             =   6495
      Width           =   1230
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10050
      TabIndex        =   14
      Top             =   6495
      Width           =   1230
   End
   Begin VB.TextBox t1 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2580
      TabIndex        =   1
      Top             =   720
      Width           =   3435
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   3000
      Left            =   435
      TabIndex        =   9
      Top             =   3300
      Width           =   10890
      _cx             =   19209
      _cy             =   5292
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSrtJalan.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      Editable        =   1
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
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lama Kredit : "
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
      Index           =   9
      Left            =   525
      TabIndex        =   26
      Top             =   2379
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Berdasar Faktur :"
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
      Index           =   0
      Left            =   540
      TabIndex        =   18
      Top             =   2775
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+F4=Close  F4=Bantuan "
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
      Left            =   465
      TabIndex        =   22
      Top             =   7875
      Width           =   6285
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Keterangan : "
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
      Index           =   11
      Left            =   450
      TabIndex        =   10
      Top             =   6330
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Salesman :"
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
      Index           =   4
      Left            =   525
      TabIndex        =   7
      Top             =   1983
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Kepada :"
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
      Index           =   3
      Left            =   525
      TabIndex        =   5
      Top             =   1587
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tgl :"
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
      Index           =   2
      Left            =   525
      TabIndex        =   3
      Top             =   1191
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&No SPB :"
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
      Index           =   1
      Left            =   525
      TabIndex        =   0
      Top             =   795
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surat Pengantar Barang"
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
      TabIndex        =   17
      Top             =   105
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   8055
      Left            =   270
      Top             =   210
      Width           =   13200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F4=Input, F5=Sort, F6=Filter, F7=Form View, F8=Print, F9=Refresh, F10=Search, Alt+X=Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   420
      TabIndex        =   21
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
      TabIndex        =   16
      Top             =   195
      Width           =   3255
   End
   Begin VB.Label Label1 
      Height          =   795
      Left            =   13170
      TabIndex        =   15
      Top             =   8730
      Width           =   1665
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   8295
      Left            =   165
      Top             =   105
      Width           =   13410
   End
End
Attribute VB_Name = "frmSrtJalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JenisJenis As String
Dim edRow As Integer, edCol  As Integer
Dim intNoFaktur As Single
Dim PtoScreen As Boolean
Private Sub CmdBaru_Click()
On Error GoTo aa
t1 = SerbaGuna.NoFaktur("SJ", "SuratJalan", "Baa")
t2 = Date
Master.Rows = 1
Master.Rows = 2
Master.Row = 1
Master.Col = 0
t3 = ""
t4 = ""
intNoFaktur = 0
t7 = ""
t8 = ""
PtoScreen = False
Text1.Text = ""
On Error Resume Next
t3.SetFocus
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdCekFaktur_Click()
  On Error GoTo aa
'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = Diskon    (5) = CaraPembayaran c/k
'(6) = LamaKredit (7) = jenistrans b/rb  (8) = Keterangan

'0=Kodebrg 1=Namabrg 2=Harga   3=QtyKArton   4=qtypcs  5=Diskon  6=subtotal 7=qty/karton
  
Dim rsMaster As New ADODB.Recordset
Set rsMaster = aData.AmbilCommand("select * from transjual where [No Faktur]='" & _
Text1.Text & "' and hapus=0")
  If Not rsMaster.EOF Then
  Text1.Text = rsMaster![No Faktur]
  t2 = rsMaster!tgl
  t3.Text = rsMaster!Kepada
  t4.Text = rsMaster!Salesman
  t7.Text = rsMaster!jatuhtempo
  t8.Text = IIf(IsNull(rsMaster!Keterangan), "", (rsMaster!Keterangan))
  intNoFaktur = rsMaster!intNo
    Set rsMaster = aData.AmbilCommand("select detailjual.*,barang.[Nama Barang] " & _
    "from detailjual,barang where detailjual.[Kode Barang]=barang.[Kode Barang] and " & _
    "inttrans=" & intNoFaktur)
    Master.Rows = 1
    Do While Not rsMaster.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = rsMaster![Kode Barang]
    Master.TextMatrix(Master.Rows - 1, 1) = rsMaster![Nama Barang]
    Master.TextMatrix(Master.Rows - 1, 2) = rsMaster!qty
    Master.TextMatrix(Master.Rows - 1, 3) = 0
    rsMaster.MoveNext
    Loop
  Else
  MsgBox "No Faktur " & Text1 & vbCrLf & "tidak ditemukan pada database atau telah dihapus", vbInformation, "Cek Faktur"
  Call CmdBaru_Click
  End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdHapus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = vbShiftMask Then PtoScreen = True
End Sub

Private Sub CmdSimpan_Click()
On Error GoTo aa
'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = LamaKredit    (5) = jenistrans
'(6) =Keterangan
Dim aHead(7) As String, aBodi() As String, i As Byte, aHasil As String
ReDim aBodi(Master.Rows - 1, 4)
aHead(0) = SerbaGuna.AmanOi(t1.Text) 'NoFaktur
aHead(1) = SerbaGuna.AmanTgl(t2.Text) 'Tgl
aHead(2) = SerbaGuna.AmanOi(t3.Text) 'Supplier
aHead(3) = SerbaGuna.AmanOi(t4.Text) 'Salesman
aHead(4) = Val(SerbaGuna.AmanOi(t7.Text)) 'LamaKredit
aHead(5) = ""
aHead(6) = SerbaGuna.AmanOi(t8.Text) 'Keterangan
For i = 1 To Master.Rows - 1
If Master.TextMatrix(i, 0) <> "" Then
aBodi(i, 0) = SerbaGuna.AmanOi(Master.TextMatrix(i, 0)) 'Kode
aBodi(i, 1) = Val(Master.TextMatrix(i, 2))  'Qty
End If
Next i

aHasil = aData.SimpanSPB(aHead, aBodi)
If aHasil <> "" Then
MsgBox aHasil, vbInformation, Me.Caption & "#Error"
Else
If MsgBox("Cetak Faktur..?", vbInformation + vbYesNo, "Faktur") = vbYes Then
Call CmdHapus_Click
End If
Call CmdBaru_Click
End If

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub CmdHapus_Click()
On Error GoTo aa
Dim Rrpt As New CRAXDdRT.Application
Screen.MousePointer = vbHourglass
Set fRpt.Report = Rrpt.OpenReport(App.Path & "\SPB.rpt")
 ' SuratJalan.Kepada+'-'+
fRpt.Report.Database.SetDataSource aData.AmbilCommand("SELECT SuratJalan.intno, SuratJalan.[No Faktur], " & _
"Konsumen.Nama AS Kepada, Konsumen.Alamat, Konsumen.Kota , SuratJalan.Salesman AS Salesman, SuratJalan.Tgl, SuratJalan.JatuhTempo, SuratJalan.Keterangan, SuratJalan.JenisTrans, SuratJalanDetail.[Kode Barang], Barang.[Nama Barang], SuratJalanDetail.Qty " & _
"FROM Konsumen INNER JOIN (Barang INNER JOIN (SuratJalan INNER JOIN SuratJalanDetail ON SuratJalan.intno=SuratJalanDetail.intTrans) ON Barang.[Kode Barang]=SuratJalanDetail.[Kode Barang]) ON Konsumen.Kode=SuratJalan.Kepada " & _
"WHERE SuratJalan.[No Faktur]='" & t1.Text & "'"), 3
If PtoScreen Then
PtoScreen = False
fRpt.aView.ReportSource = fRpt.Report
fRpt.aView.ViewReport
fRpt.Show 'vbModal
Else
fRpt.Report.PrintOut
End If
Screen.MousePointer = vbDefault
Exit Sub
aa:
Screen.MousePointer = vbDefault
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = vbShiftMask Then PtoScreen = True
End Sub

Private Sub Command1_Click()
On Error GoTo aa
    CD1.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "ExportJurnal", "\ExportJurnal")
    CD1.FileName = ""
    On Error Resume Next
    CD1.CancelError = True
    CD1.DialogTitle = "Masukkan nama file Transaksi yang akan diimport.."
    CD1.Filter = "Transaksi File|*.trns"
    CD1.ShowOpen
    If Err.Number = 0 Then
    On Error GoTo aa
    Dim k As New ADODB.Recordset
    k.Open CD1.FileName
    k.MoveFirst
    JenisJenis = k(6)
    Call Form_Load
    t2.Text = k(0)
    t3.Text = k(1)
    t4.Text = k(2)
    t7.Text = k(5)
    t8.Text = k(7)
    k.MoveNext
    Master.Rows = 1
     Do While Not k.EOF
     Master.Rows = Master.Rows + 1
     Master.TextMatrix(Master.Rows - 1, 0) = k(0)
     Master.TextMatrix(Master.Rows - 1, 1) = k(1)
     Master.TextMatrix(Master.Rows - 1, 2) = k(2)
     Master.TextMatrix(Master.Rows - 1, 3) = k(4)
     Master.TextMatrix(Master.Rows - 1, 4) = k(3)
     Master.TextMatrix(Master.Rows - 1, 5) = k(5)
     Master.TextMatrix(Master.Rows - 1, 6) = k(6)
     Master.TextMatrix(Master.Rows - 1, 7) = k(7)
     k.MoveNext
     Loop
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Command2_Click()
On Error GoTo aa
    CD1.InitDir = App.Path & IIf(Right(App.Path, 1) = "\", "ExportJurnal", "\ExportJurnal")
    CD1.FileName = ""
    On Error Resume Next
    CD1.CancelError = True
    CD1.DialogTitle = "Masukkan nama file Transaksi yang diexport.."
    CD1.Filter = "Transaksi File|*.Trns"
    CD1.ShowSave
    If Err.Number = 0 Then
    On Error GoTo aa
    Dim k As New ADODB.Recordset, i As Integer
    '(0)=Tgl (1)= Supplier  (2)=Salesman (3)=Diskon (4)=CaraPembayaran c/k
    '(5)=LamaKredit (6)=jenistrans b/rb  (7)=Keterangan
    '0=Kodebrg 1=NAmabrg 2=Harga   3=QtyKArton   4=qtypcs  5=Diskon  6=subtotal 7=qty/karton
    k.Fields.Append "Kode", adVarChar, 50
    k.Fields.Append "Nama", adVarChar, 200
    k.Fields.Append "Harga", adVarChar, 50
    k.Fields.Append "Qty", adVarChar, 50
    k.Fields.Append "QtyKarton", adVarChar, 50
    k.Fields.Append "Diskon", adVarChar, 50
    k.Fields.Append "Total", adVarChar, 50
    k.Fields.Append "Karton", adVarChar, 200
    k.Open
    k.AddNew:
    k!Kode = t2.Text
    k!Nama = t3.Text
    k!Harga = t4.Text
    k!Diskon = t7.Text
    k!Total = JenisJenis
    k!Karton = t8.Text
    k.Update
     For i = 1 To Master.Rows - 1
     If Master.TextMatrix(i, 1) <> "" Then
     k.AddNew
     k!Kode = Master.TextMatrix(i, 0)
     k!Nama = Master.TextMatrix(i, 1)
     k!Harga = Master.TextMatrix(i, 2)
     k!qty = Master.TextMatrix(i, 4)
     k!QtyKarton = Master.TextMatrix(i, 3)
     k!Diskon = Master.TextMatrix(i, 5)
     k!Total = Master.TextMatrix(i, 6)
     k!Karton = Master.TextMatrix(i, 7)
     k.Update
     End If
     Next i
    k.Save CD1.FileName, adPersistADTG
    MsgBox "Data telah disimpan..", vbInformation, "Simpan Transaksi"
    End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Command3_Click()
  On Error GoTo aa
'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = Diskon    (5) = CaraPembayaran c/k
'(6) = LamaKredit (7) = jenistrans b/rb  (8) = Keterangan

'0=Kodebrg 1=Namabrg 2=Harga   3=QtyKArton   4=qtypcs  5=Diskon  6=subtotal 7=qty/karton
  
Dim rsMaster As New ADODB.Recordset
Set rsMaster = aData.AmbilCommand("select * from suratjalan where [No Faktur]='" & _
t1.Text & "'")
  If Not rsMaster.EOF Then
  t1.Text = rsMaster![No Faktur]
  t2 = rsMaster!tgl
  t3.Text = rsMaster!Kepada
  t4.Text = rsMaster!Salesman
  t7.Text = rsMaster!jatuhtempo
  t8.Text = rsMaster!Keterangan
  intNoFaktur = rsMaster!intNo
    Set rsMaster = aData.AmbilCommand("select suratjalandetail.*,barang.[Nama Barang] " & _
    "from suratjalandetail,barang where suratjalandetail.[Kode Barang]=barang.[Kode Barang] and " & _
    "inttrans=" & intNoFaktur)
    Master.Rows = 1
    Do While Not rsMaster.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = rsMaster![Kode Barang]
    Master.TextMatrix(Master.Rows - 1, 1) = rsMaster![Nama Barang]
    Master.TextMatrix(Master.Rows - 1, 2) = rsMaster!qty
    Master.TextMatrix(Master.Rows - 1, 3) = 0
    rsMaster.MoveNext
    Loop
  Else
  MsgBox "No Faktur " & t1 & vbCrLf & "tidak ditemukan pada database atau telah dihapus", vbInformation, "Cek Faktur"
  Call CmdBaru_Click
  End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyX And Shift = 4 Then   '############ TUTUP FORM ############
Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo aa
Call CmdBaru_Click
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Master_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyInsert Then
Master.Rows = Master.Rows + 1
Master.Row = Master.Rows - 1
Master.Col = 0
End If
If KeyCode = vbKeyDelete And Master.Rows <> 1 Then
Master.RemoveItem Master.Row
End If

End Sub

Private Sub Master_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
    Select Case Col
    Case 2
     edRow = Master.Row
     edCol = Master.Col
     Bantu.Apaan = "Angka"
     Bantu.Left = Master.CellLeft + Master.Left + 50
     Bantu.Top = Master.CellTop + Master.Top + Master.CellHeight + 400
     Bantu.Show vbModal
     Master.Text = Bantu.NilaiBantu
     Master.TextMatrix(Row, 6) = IIf(InStr(1, Master.TextMatrix(Row, 5), "%"), (1 - Val(Master.TextMatrix(Row, 5)) / 100) * Val(Master.TextMatrix(Row, 2)), Val(Master.TextMatrix(Row, 2)) - Val(Master.TextMatrix(Row, 5))) * (Val(Master.TextMatrix(Row, 3)) * Val(Master.TextMatrix(Row, 7)) + Val(Master.TextMatrix(Row, 4)))
    End Select
End If

If Master.Col = 0 And KeyCode = vbKeyReturn Then
'     Cancel = True
     edRow = Master.Row
     edCol = Master.Col
     KeyCode = 0
     BantuBarang.Show vbModal
     If Not BantuBarang.Batal Then
     Master.TextMatrix(edRow, edCol) = BarisGrid(0)
     Master.TextMatrix(edRow, edCol + 1) = BarisGrid(1)
     Master.TextMatrix(edRow, edCol + 2) = 0
     Master.TextMatrix(edRow, edCol + 3) = 0
     Master.Col = 2
     End If
End If

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
'0=Kodebrg 1=NAmabrg 2=Harga   3=QtyKArton   4=qtypcs  5=Diskon  6=subtotal 7=qty/karton

Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
    Select Case Col
    Case 1
    Cancel = True
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
'0=Kodebrg 1=NAmabrg 2=Harga   3=QtyKArton   4=qtypcs  5=Diskon  6=subtotal 7=qty/karton

Private Sub Master_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
    Select Case Col
    Case 0
      If Master.TextMatrix(Row, Col) <> "" And Master.TextMatrix(Row, 1) = "" Then
      Dim rsSbrg As New ADODB.Recordset
      Set rsSbrg = aData.AmbilCommand("SELECT Barang.[Kode Barang], Barang.[Nama Barang], Barang.Qty, Barang.[Harga Beli], Barang.[Harga Jual]," & _
      "[Qty per Karton] FROM Barang where Barang.[Kode Barang]='" & AmanOi(Master.TextMatrix(Row, Col)) & "';")
        If rsSbrg.EOF Then
         MsgBox "Kode barang tersebut tidak ada, harap cek kembali kode barang anda", vbInformation, "Kode Barang"
         Master.TextMatrix(Row, Col) = ""
         Exit Sub
        Else
         Master.TextMatrix(Row, 0) = rsSbrg![Kode Barang]
         Master.TextMatrix(Row, 1) = rsSbrg![Nama Barang]
         Master.TextMatrix(Row, 2) = 0
         Master.TextMatrix(Row, 3) = 0
         Master.Col = 2
        End If
      End If
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Tanggal"
     Bantu.Left = t2.Left + Me.Left + 50
     Bantu.Top = t2.Top + Me.Top + t2.Height + 500
     Bantu.Show vbModal
     t2.Text = Bantu.NilaiBantu
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     Bantu.GridData = "SELECT Kode, Nama, Alamat, Kota, Telepon FROM Konsumen"
     Bantu.Left = t3.Left + Me.Left + 50
     Bantu.Top = t3.Top + Me.Top + t3.Height + 500
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     t3.Text = BarisGrid(0)
     End If
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     Bantu.GridData = "Salesman"
     Bantu.Left = t4.Left + Me.Left + 50
     Bantu.Top = t4.Top + Me.Top + t4.Height + 500
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     t4.Text = BarisGrid(0)
     End If
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t7_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Angka"
     Bantu.Left = t7.Left + Me.Left + 50
     Bantu.Top = t7.Top + Me.Top + t7.Height + 500
     Bantu.Show vbModal
     t7.Text = Bantu.NilaiBantu
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     Bantu.GridData = "SELECT intNo,[No Faktur],Kepada,Tgl,Total,Salesman FROM TransJual"
     Bantu.Left = t3.Left + Me.Left + 50
     Bantu.Top = t3.Top + Me.Top + t3.Height + 500
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     Text1.Text = BarisGrid(1)
     End If
     Call CmdCekFaktur_Click
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub





