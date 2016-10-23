VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRetur 
   Caption         =   "Transaksi retur"
   ClientHeight    =   8820
   ClientLeft      =   840
   ClientTop       =   765
   ClientWidth     =   13830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   13830
   WindowState     =   2  'Maximized
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
      Height          =   345
      Left            =   7770
      TabIndex        =   19
      Top             =   7005
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
      Height          =   345
      Left            =   7770
      TabIndex        =   18
      Top             =   7515
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
      TabIndex        =   8
      Top             =   6660
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
      Height          =   465
      Left            =   6075
      TabIndex        =   11
      Top             =   765
      Width           =   1755
   End
   Begin VB.TextBox t3 
      DataField       =   "No Faktur"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   5
      Top             =   1530
      Width           =   3435
   End
   Begin VB.TextBox t2 
      DataField       =   "No Faktur"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   1155
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
      Height          =   465
      Left            =   10395
      TabIndex        =   9
      Top             =   6375
      Width           =   1230
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11745
      TabIndex        =   10
      Top             =   6375
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
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   780
      Width           =   3435
   End
   Begin VB.TextBox txtFields 
      DataField       =   "JenisTrans"
      Height          =   315
      Left            =   390
      TabIndex        =   12
      Top             =   8445
      Visible         =   0   'False
      Width           =   315
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   4320
      Left            =   435
      TabIndex        =   6
      Top             =   1980
      Width           =   12555
      _cx             =   22146
      _cy             =   7620
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRetur.frx":0000
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
      TabIndex        =   17
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
      TabIndex        =   7
      Top             =   6345
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dari:"
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
      TabIndex        =   4
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tgl:"
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
      TabIndex        =   2
      Top             =   1188
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&No Faktur:"
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
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pembelian"
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
      TabIndex        =   15
      Top             =   105
      Width           =   3255
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
      Height          =   495
      Left            =   420
      TabIndex        =   16
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
      TabIndex        =   14
      Top             =   195
      Width           =   3255
   End
   Begin VB.Label Label1 
      Height          =   795
      Left            =   13170
      TabIndex        =   13
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
Attribute VB_Name = "frmRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JenisJenis As String
Dim edRow As Integer, edCol  As Integer
Dim intNoFaktur As Single
Private Sub CmdBaru_Click()
On Error GoTo aa
If JenisJenis = "RB" Then
lblLabels(3).Caption = "Kepada :"
ElseIf JenisJenis = "RJ" Then
lblLabels(3).Caption = "Dari :"
End If
T1 = ""
t2 = ""
Master.Rows = 1
Master.Rows = 2
Master.Row = 1
Master.Col = 0
t3 = ""
intNoFaktur = 0
t8 = ""
On Error Resume Next
T1.SetFocus
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdCekFaktur_Click()
  On Error GoTo aa
'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = Diskon    (5) = CaraPembayaran c/k
'(6) = LamaKredit (7) = jenistrans b/rb  (8) = Keterangan

  
Dim rsMaster As New ADODB.Recordset, rsUdaBayar As New ADODB.Recordset, strExec1 As String
If JenisJenis = "RB" Then
Set rsMaster = aData.AmbilCommand("select * from transbeli where jenistrans='" & _
JenisJenis & "' and [No Faktur]='" & T1.Text & "' and hapus=0 and (jenis='K')")
ElseIf JenisJenis = "RJ" Then
Set rsMaster = aData.AmbilCommand("select * from transjual where jenistrans='" & _
JenisJenis & "' and [No Faktur]='" & T1.Text & "' and hapus=0 and (jenis='K')")
End If
  If Not rsMaster.EOF Then
    strExec1 = "Select intno from " & IIf(JenisJenis = "RB", "InputUtang", "InputPiutang") & " Where intno=" & rsMaster!intNo & " and bayar<>0"
    Set rsUdaBayar = aData.AmbilCommand(strExec1)
    If Not rsUdaBayar.EOF Then
    MsgBox "Transaksi tersebut sudah dibayar.. tidak dapat diproses lagi", vbInformation, "Pengembalian Retur"
    Exit Sub
    End If
  T1.Text = rsMaster![no faktur]
  t2 = rsMaster!tgl
  t3.Text = rsMaster!Kepada
  t8.Text = rsMaster!keterangan
  intNoFaktur = rsMaster!intNo
    If JenisJenis = "RB" Then
    Set rsMaster = aData.AmbilCommand("select detailbeli.*,barang.[Nama Barang], 0 as HRata " & _
    "from detailbeli,barang where detailbeli.[Kode Barang]=barang.[Kode Barang] and " & _
    "inttrans=" & intNoFaktur & " order by detailbeli.[Kode Barang], detailbeli.qty desc")
    ElseIf JenisJenis = "RJ" Then
    Set rsMaster = aData.AmbilCommand("select detailjual.*,barang.[Nama Barang] " & _
    "from detailjual,barang where detailjual.[Kode Barang]=barang.[Kode Barang] and " & _
    "inttrans=" & intNoFaktur & " order by detailjual.[Kode Barang], detailjual.qty desc")
    End If
    Master.Rows = 1
    '0=Kodebrg 1=Namabrg 2=Qty   3=QtyMinus   4=0  5=3-4-5  6=intTrans 7=harga 8=diskon
    
    Do While Not rsMaster.EOF
    
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = rsMaster![Kode Barang]
    Master.TextMatrix(Master.Rows - 1, 1) = rsMaster![Nama Barang]
    Master.TextMatrix(Master.Rows - 1, 2) = rsMaster!Qty
    Master.TextMatrix(Master.Rows - 1, 3) = 0
    Master.TextMatrix(Master.Rows - 1, 4) = 0
    Master.TextMatrix(Master.Rows - 1, 6) = rsMaster!intTrans
    Master.TextMatrix(Master.Rows - 1, 7) = rsMaster!Harga
    Master.TextMatrix(Master.Rows - 1, 8) = rsMaster!Diskon
    Master.TextMatrix(Master.Rows - 1, 9) = IIf(JenisJenis = "RB", 0, rsMaster!HRata)
    rsMaster.MoveNext
    If Not (rsMaster.EOF) Then
    If Master.TextMatrix(Master.Rows - 1, 0) = rsMaster![Kode Barang] Then
    Master.TextMatrix(Master.Rows - 1, 3) = -rsMaster!Qty
    Master.TextMatrix(Master.Rows - 1, 5) = Master.TextMatrix(Master.Rows - 1, 2) - Master.TextMatrix(Master.Rows - 1, 3) - Master.TextMatrix(Master.Rows - 1, 4)
    rsMaster.MoveNext
    End If
    End If
    Loop
  
  Master.Row = 1
  Else
  MsgBox "No Faktur " & T1 & vbCrLf & "tidak ditemukan pada database atau telah dihapus", vbInformation, "Cek Faktur"
  Call CmdBaru_Click
  End If
  
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_Click()
On Error GoTo aa
Dim i As Byte, aHasil As String
'Call TotalOi
aHasil = aData.SimpanRetur(Me.JenisJenis, Master)
If aHasil = "" Then
Call CmdBaru_Click
Else: MsgBox aHasil
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub CmdHapus_Click()
On Error GoTo aa
Call CmdBaru_Click
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
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
    k!Total = JenisJenis
    k!Karton = t8.Text
    k.Update
     For i = 1 To Master.Rows - 1
     If Master.TextMatrix(i, 1) <> "" Then
     k.AddNew
     k!Kode = Master.TextMatrix(i, 0)
     k!Nama = Master.TextMatrix(i, 1)
     k!Harga = Master.TextMatrix(i, 2)
     k!Qty = Master.TextMatrix(i, 4)
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyX And Shift = 4 Then   '############ TUTUP FORM ############
Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo aa
If JenisJenis = "RB" Then
Me.Caption = "Tukar Retur Pembelian"
Label2.Caption = "Tukar Retur Pembelian"
ElseIf JenisJenis = "RJ" Then
Me.Caption = "Tukar Retur Penjualan"
Label2.Caption = "Tukar Retur Penjualan"
End If
Call CmdBaru_Click
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub


Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
    Select Case Col
    Case 0, 1, 2, 3, 5
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
    Case 4
aac:
      Master.TextMatrix(Row, 4) = Val(Master.TextMatrix(Row, 4))
      If Val(Master.TextMatrix(Row, 4)) < 0 Then Master.TextMatrix(Row, 4) = -Val(Master.TextMatrix(Row, 4))
      Master.TextMatrix(Row, 5) = Val(Master.TextMatrix(Row, 2)) - Val(Master.TextMatrix(Row, 3)) - Val(Master.TextMatrix(Row, 4))
      If Val(Master.TextMatrix(Row, 5)) < 0 Then
      Master.TextMatrix(Row, 4) = 0
      GoTo aac
      End If
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub



Private Sub t1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     If JenisJenis = "RB" Then
     Bantu.GridData = "SELECT intno as [Trans No],[No Faktur],Kepada,Tgl as Tanggal from TransBeli where JenisTrans='RB' and jenis='K' and hapus=0;"
     ElseIf JenisJenis = "RJ" Then
     Bantu.GridData = "SELECT intno as [Trans No],[No Faktur],Kepada,Tgl as Tanggal from TransJual where JenisTrans='RJ' and jenis='K' and hapus=0;"
     End If
     Bantu.Left = t3.Left + Me.Left + 50
     Bantu.Top = t3.Top + Me.Top + t3.Height + 500
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     T1.Text = BarisGrid(1)
     End If
     CmdCekFaktur_Click
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


