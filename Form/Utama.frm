VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Utama 
   BackColor       =   &H8000000C&
   Caption         =   "Utama"
   ClientHeight    =   7845
   ClientLeft      =   3585
   ClientTop       =   2115
   ClientWidth     =   7260
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7575
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4577
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "18/09/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "16:05"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu1File 
      Caption         =   "&File"
      Begin VB.Menu mnu2Produk 
         Caption         =   "&Produk"
      End
      Begin VB.Menu mnufSupplier 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu mnufKonsumen 
         Caption         =   "&Konsumen"
      End
      Begin VB.Menu mnufSalesman 
         Caption         =   "Sa&lesman"
      End
      Begin VB.Menu mnuFileMataUang 
         Caption         =   "&Mata Uang"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu2grs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipe 
         Caption         =   "&Tipe"
      End
      Begin VB.Menu mnuWilayah 
         Caption         =   "&Wilayah"
      End
      Begin VB.Menu mnugrs2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu2exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnutransBeli 
         Caption         =   "Pem&Belian"
      End
      Begin VB.Menu mnutransreturbeli 
         Caption         =   "Retur Pem&belian"
         Begin VB.Menu mnuTtrBeliInput 
            Caption         =   "&Input Retur Pembelian"
         End
         Begin VB.Menu mnuRtrBeliTukar 
            Caption         =   "&Tukar Retur Pembelian"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnutranspenjualan 
         Caption         =   "Pen&jualan"
         Begin VB.Menu mnuInJual 
            Caption         =   "&Input Penjualan"
         End
         Begin VB.Menu mnuSPB 
            Caption         =   "&Cetak SPB"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnutransreturpenjualan 
         Caption         =   "&Retur Penjualan"
         Begin VB.Menu mnuRtrJualInput 
            Caption         =   "&Input Retur Penjualan"
         End
         Begin VB.Menu mnuRtrJualTukar 
            Caption         =   "&Tukar Retur Penjualan"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuTransBatch 
         Caption         =   "&Proses Offline Transaksi"
      End
      Begin VB.Menu mnuHisTrans 
         Caption         =   "&Histori Transaksi"
      End
   End
   Begin VB.Menu mnuUtang 
      Caption         =   "&Utang"
      Begin VB.Menu mnuUtangBayar 
         Caption         =   "&Pembayaran Utang"
      End
      Begin VB.Menu mnuUtangReal 
         Caption         =   "&Realisasi Utang (Giro/Bank/Batal)"
      End
      Begin VB.Menu mnuUtangTT 
         Caption         =   "&Tanda Terima Hutang"
      End
   End
   Begin VB.Menu mnuPiutang 
      Caption         =   "&Piutang"
      Begin VB.Menu mnuPiutangbayarpiutang 
         Caption         =   "&Pembayaran Piutang"
      End
      Begin VB.Menu mnuPiutangRealisasi 
         Caption         =   "&Realisasi Piutang (Giro/Bank/Batal)"
      End
      Begin VB.Menu mnuPiutangTT 
         Caption         =   "&Tanda Terima Piutang"
      End
   End
   Begin VB.Menu mnuBiaya 
      Caption         =   "&Biaya"
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnuLapFile 
         Caption         =   "&File/Master"
         Begin VB.Menu mnuLapFileSupp 
            Caption         =   "&Supplier"
         End
         Begin VB.Menu mnuLApFileKonsumen 
            Caption         =   "&Konsumen"
         End
         Begin VB.Menu mnuLapFileSalesman 
            Caption         =   "&Salesman"
         End
      End
      Begin VB.Menu mnuLapTransBeli 
         Caption         =   "Pem&belian/Retur Pembelian"
         Begin VB.Menu mnuLapBeliFaktur 
            Caption         =   "Lihat/Cek &Faktur"
         End
         Begin VB.Menu mnuLapBeliTokoBulanFaktur 
            Caption         =   "Per Toko=>Bulan=>Faktur"
         End
         Begin VB.Menu mnuLapBeliSupp 
            Caption         =   "Pembelian/Retur per &Supplier"
         End
         Begin VB.Menu mnuLapBeliBarang 
            Caption         =   "Pembelian/Retur per Item Barang"
         End
         Begin VB.Menu mnuLapBeliTotalBeli 
            Caption         =   "&Total Pembelian"
         End
         Begin VB.Menu mnuLapBeliRetur 
            Caption         =   "Total &Retur Pembelian"
         End
         Begin VB.Menu mnuLapBeliItemTgl 
            Caption         =   "&Item Barang Masuk per Tanggal"
         End
         Begin VB.Menu mnuTransBFak 
            Caption         =   "&Transaksi per faktur"
         End
      End
      Begin VB.Menu mnuLapTransJual 
         Caption         =   "Pen&jualan/ReturPenjualan"
         Begin VB.Menu mnuLapJualFaktur 
            Caption         =   "Lihat/Cek &Faktur"
         End
         Begin VB.Menu mnuLapJualTokoBulanFaktur 
            Caption         =   "Per Toko=>Bulan=>Faktur"
         End
         Begin VB.Menu mnuLapJualKonsumen 
            Caption         =   "Penjualan/Retur per &Konsumen"
         End
         Begin VB.Menu mnuLapJualSalesman 
            Caption         =   "Penjualan/Retur per &Salesman"
         End
         Begin VB.Menu mnuLapJualBarang 
            Caption         =   "Penjualan/Retur per Item Barang"
         End
         Begin VB.Menu mnuLapJualOmzetPerSales 
            Caption         =   "&Omzet per Salesman"
         End
         Begin VB.Menu mnuLapJualTotal 
            Caption         =   "&Total Penjualan"
         End
         Begin VB.Menu mnuLapPenjualanTotalRetur 
            Caption         =   "Total &Retur Penjualan"
         End
         Begin VB.Menu mnuTransJFak 
            Caption         =   "&Transaksi per Faktur"
         End
      End
      Begin VB.Menu mnuLapStock 
         Caption         =   "&Stock"
         Begin VB.Menu mnuLapStockMin 
            Caption         =   "&Daftar Stock Hampir Habis"
         End
         Begin VB.Menu mnuLapStockQty 
            Caption         =   "&Jumlah Stock"
         End
         Begin VB.Menu mnuLapStockKartu 
            Caption         =   "&Kartu Stock"
         End
         Begin VB.Menu mnuLapStockKartuQty 
            Caption         =   "Kartu Stock &Qty"
         End
         Begin VB.Menu mnuLapStockNilai 
            Caption         =   "Persediaan Barang Dagangan"
         End
      End
      Begin VB.Menu mnuLapHutang 
         Caption         =   "&Hutang"
         Begin VB.Menu mnuLapHutangDaftar 
            Caption         =   "&Daftar Seluruh Hutang"
         End
         Begin VB.Menu mnuDafUtExGiro 
            Caption         =   "&Daftar Seluruh Hutang (Ex.Giro)"
         End
         Begin VB.Menu mnuLapHutangDaftarArea 
            Caption         =   "Daftar Hutang per &Area"
         End
         Begin VB.Menu mnuLapHutangJatuh 
            Caption         =   "&Hutang Jatuh Tempo"
         End
         Begin VB.Menu mnuLapHutangBayar 
            Caption         =   "&Pembayaran Hutang per Supplier"
         End
         Begin VB.Menu mnuLapHutangBayartanggal 
            Caption         =   "&Pembayaran Hutang per Tanggal"
         End
         Begin VB.Menu mnuLapHutangGiro 
            Caption         =   "&Kontrol Giro/Bank"
         End
      End
      Begin VB.Menu mnuLapPiutang 
         Caption         =   "&Piutang"
         Begin VB.Menu mnuLapPiutangDaftar 
            Caption         =   "&Daftar Seluruh Piutang"
         End
         Begin VB.Menu mnuDafPiutExGiro 
            Caption         =   "&Daftar Seluruh Piutang (ex. Giro)"
         End
         Begin VB.Menu mnuLapPiutangDaftarArea 
            Caption         =   "Daftar Piutang per &Area"
         End
         Begin VB.Menu mnuLapPiutangJatuh 
            Caption         =   "&Piutang Jatuh Tempo"
         End
         Begin VB.Menu mnuLapPiutangBayar 
            Caption         =   "&Pembayaran Piutang per Konsumen"
         End
         Begin VB.Menu mnuLapPiutangBayartanggal 
            Caption         =   "&Pembayaran Piutang per Tanggal"
         End
         Begin VB.Menu mnuLapPiutangGiro 
            Caption         =   "&Kontrol Giro/Bank"
         End
      End
      Begin VB.Menu mnuSalesKomisi 
         Caption         =   "Penjualan Sales Per Barang"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLR 
         Caption         =   "&Laba/Rugi"
         Begin VB.Menu mnuLRItem 
            Caption         =   "Laba Rugi Item"
         End
         Begin VB.Menu mnuLRSales 
            Caption         =   "Laba Rugi per &Sales"
         End
         Begin VB.Menu mnuLRFaktur 
            Caption         =   "Laba Rugi per Faktur"
         End
      End
      Begin VB.Menu mnuLapBiaya 
         Caption         =   "&Biaya"
      End
      Begin VB.Menu mnuRekapData 
         Caption         =   "&Rekap Transaksi"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "T&ool"
      Begin VB.Menu mnuSETPASS 
         Caption         =   "&Ubah Password"
      End
      Begin VB.Menu mnuToolTutupBuku 
         Caption         =   "&Tutup Buku"
      End
      Begin VB.Menu mnuHitUlang 
         Caption         =   "&Hitung Ulang L/R"
      End
      Begin VB.Menu mnuCekInput 
         Caption         =   "&Cek Siapa Yang Input"
      End
      Begin VB.Menu mnuGrsT1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&BackUp Data"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "&Register"
      Visible         =   0   'False
      Begin VB.Menu mnuRegisterLicense 
         Caption         =   "&License Registration"
      End
   End
End
Attribute VB_Name = "Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OOQ As Boolean

Private Sub MDIForm_Load()
'If Pendaftaran Then mnuRegister.Visible = False
OOQ = False
Me.Caption = "Asia Baru " & "- " & LokasiFile
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Unload RptPalsu
'Set RptPalsu = Nothing
Set Lap = Nothing
Set fRpt = Nothing
Set aData = Nothing
OOQ = True
Unload BantuBarang
Unload frmSplash
End Sub

Private Sub mnu2exit_Click()
Unload Me
End
End Sub

Private Sub mnu2Produk_Click()
fBarang.Show
fBarang.ZOrder 0
End Sub

Private Sub mnuBackup_Click()
On Error Resume Next
fBackUp.Show vbModal
End Sub

Private Sub mnuBiaya_Click()
fBiaya.Show
fBiaya.ZOrder 0
End Sub

Private Sub mnuCekInput_Click()
 frmCekSiapa.Show
 frmCekSiapa.ZOrder 0
End Sub

Private Sub mnuDafPiutExGiro_Click()
On Error GoTo aa
Call AturLaporan("PDaftarNoGiro")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuDafUtExGiro_Click()
On Error GoTo aa
Call AturLaporan("HDaftarNoGiro")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuFileMataUang_Click()
'fMataUang.Show
End Sub

Private Sub mnufKonsumen_Click()
fKonsumen.Show
fKonsumen.ZOrder 0
End Sub

Private Sub mnufSalesman_Click()
fSalesman.Show
fSalesman.ZOrder 0
End Sub

Private Sub mnufSupplier_Click()
fSupplier.Show
fSupplier.ZOrder 0
End Sub

Private Sub mnuHisTrans_Click()
frmHisTrans.Show
frmHisTrans.ZOrder 0
End Sub

Private Sub mnuHitUlang_Click()
FrmHitungUlang.Show vbModal
End Sub

Private Sub mnuInJual_Click()
Dim aJual As New frmPembelian
aJual.JenisJenis = "J"
aJual.Show
End Sub

Private Sub mnuLapBeliBarang_Click()
On Error GoTo aa
Call AturLaporan("ItemBarangB")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBeliFaktur_Click()
On Error GoTo aa
Call AturLaporan("FakturB")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBeliItemTgl_Click()
On Error GoTo aa
Call AturLaporan("BItemBrgTgl")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBeliRetur_Click()
On Error GoTo aa
Call AturLaporan("BPlus", "RB")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBeliSupp_Click()
On Error GoTo aa
Call AturLaporan("BSupp")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBeliTokoBulanFaktur_Click()
On Error GoTo aa
Call AturLaporan("TokoBulanFakturB")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBeliTotalBeli_Click()
On Error GoTo aa
Call AturLaporan("BPlus", "B")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapBiaya_Click()
On Error GoTo aa
Call AturLaporan("Biaya")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLApFileKonsumen_Click()
On Error GoTo aa
Call AturLaporan("MasterKonsumen")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapFileSalesman_Click()
On Error GoTo aa
Call AturLaporan("MasterSalesman")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapFileSupp_Click()
On Error GoTo aa
Call AturLaporan("MasterSupplier")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapHutangBayar_Click()
On Error GoTo aa
Call AturLaporan("HBayarSupp")
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuLapHutangBayar_Click pada Form Utama"
End Sub

Private Sub mnuLapHutangBayartanggal_Click()
On Error GoTo aa
Call AturLaporan("HBayarTgl")
Exit Sub
aa:
Dim Err_Setering As String
Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur mnuLapHutangBayartanggal_Click pada " & "Form Utama di baris " & Erl
Select Case MsgBox(Err_Setering, vbRetryCancel, App.Title & "-Utama Error")
  Case vbCancel: Resume Exit_mnuLapHutangBayartanggal_Click:
  Case vbRetry: Resume
  Case Else: End
End Select
Exit_mnuLapHutangBayartanggal_Click:
End Sub

Private Sub mnuLapHutangDaftar_Click()
On Error GoTo aa
Call AturLaporan("HDaftar")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapHutangDaftarArea_Click()
On Error GoTo aa
Call AturLaporan("HArea")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapHutangGiro_Click()
On Error GoTo aa
Call AturLaporan("HGiro")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapHutangJatuh_Click()
On Error GoTo aa
Call AturLaporan("HJatuh")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualBarang_Click()
On Error GoTo aa
Call AturLaporan("ItemBarangJ")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualFaktur_Click()
On Error GoTo aa
Call AturLaporan("FakturJ")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualKonsumen_Click()
On Error GoTo aa
Call AturLaporan("JCust")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualOmzetPerSales_Click()
On Error GoTo aa
Call AturLaporan("OmzetSales")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualSalesman_Click()
On Error GoTo aa
Call AturLaporan("JualSales")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualTokoBulanFaktur_Click()
On Error GoTo aa
Call AturLaporan("TokoBulanFakturJ")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapJualTotal_Click()
On Error GoTo aa
Call AturLaporan("JPlus", "J")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapPenjualanTotalRetur_Click()
On Error GoTo aa
Call AturLaporan("JPlus", "RJ")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapPiutangBayar_Click()
On Error GoTo aa
Call AturLaporan("PBayarCust")
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuLapPiutangBayar_Click pada Form Utama"
End Sub

Private Sub mnuLapPiutangBayartanggal_Click()
On Error GoTo aa
Call AturLaporan("PBayarTgl")
Exit Sub
aa:
Dim Err_Setering As String
Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur mnuLapPiutangBayartanggal_Click pada " & "Form Utama di baris " & Erl
Select Case MsgBox(Err_Setering, vbRetryCancel, App.Title & "-Utama Error")
  Case vbCancel: Resume Exit_mnuLapPiutangBayartanggal_Click:
  Case vbRetry: Resume
  Case Else: End
End Select
Exit_mnuLapPiutangBayartanggal_Click:
End Sub

Private Sub mnuLapPiutangDaftar_Click()
On Error GoTo aa
Call AturLaporan("PDaftar")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapPiutangDaftarArea_Click()
On Error GoTo aa
Call AturLaporan("PArea")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapPiutangGiro_Click()
On Error GoTo aa
Call AturLaporan("PGiro")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapPiutangJatuh_Click()
On Error GoTo aa
Call AturLaporan("PJatuh")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapStockKartu_Click()
On Error GoTo aa
Call AturLaporan("StockKartu")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapStockKartuQty_Click()
On Error GoTo aa
Call AturLaporan("StockKartuQty")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapStockMin_Click()
On Error GoTo aa
Call AturLaporan("StockMin")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapStockNilai_Click()
On Error GoTo aa
Call AturLaporan("StockPersBD")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLapStockQty_Click()
Dim QtyRpt As New CRAXDdRT.Report
On Error GoTo aa
Call AturLaporan("StockQty")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuLRFaktur_Click()
On Error GoTo aa
Call AturLaporan("LRFaktur")
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuLRFaktur_Click pada Form Utama"
End Sub

Private Sub mnuLRItem_Click()
On Error GoTo aa
Call AturLaporan("LRItem")
Exit Sub
aa:
Screen.MousePointer = vbNormal
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuLRItem_Click pada Form Utama"
End Sub

Private Sub mnuLRSales_Click()
On Error GoTo aa
Call AturLaporan("LRSales")
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuLRSales_Click pada Form Utama"
End Sub

Private Sub mnuPiutangbayarpiutang_Click()
Dim aUtPi As New frmPiutang
aUtPi.JenisUtPi = "Piutang"
aUtPi.Show
End Sub

Private Sub mnuPiutangRealisasi_Click()
Dim aPiuRe As New frmPiutangReal
aPiuRe.JenisUtPi = "Piutang"
aPiuRe.Show
End Sub

Private Sub mnuPiutangTT_Click()
Dim aUtPi As New frmTT
aUtPi.JenisJenis = "TTP"
aUtPi.Show
End Sub

Private Sub mnuRegisterLicense_Click()
'frmDaftar.Show vbModal
End Sub

Private Sub mnuRekapData_Click()
On Error GoTo aa
Call AturLaporan("RekapData")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuRtrBeliTukar_Click()
Dim aRtrBeli As New frmRetur
aRtrBeli.JenisJenis = "RB"
aRtrBeli.Show
End Sub

Private Sub mnuRtrJualInput_Click()
Dim aReturJual As New frmPembelian
aReturJual.JenisJenis = "RJ"
aReturJual.Show
End Sub

Private Sub mnuRtrJualTukar_Click()
Dim aRtrJual As New frmRetur
aRtrJual.JenisJenis = "RJ"
aRtrJual.Show
End Sub

Private Sub mnuSalesKomisi_Click()
On Error GoTo aa
Dim tgAw As Date, tgAk As Date
Tgkaco:
tgAw = AmanTgl(InputBox("Masukkan tanggal awal periode perhitungan :", "Jual Sales", Date - 30))
tgAk = AmanTgl(InputBox("Masukkan tanggal akhir periode perhitungan :", "Jual Sales", Date))
If Not (IsDate(tgAw) And IsDate(tgAk)) Then
  If MsgBox("Format tanggal yang anda masukkan salah!" & _
  vbCrLf & "masukkan tanggal dalam format dd/mm/yyyy" & _
  vbCrLf & "Ulangi sekali lagi..?", vbYesNo + vbQuestion, "Laba Rugi") = vbYes Then
  GoTo Tgkaco:
  Else
  Exit Sub
  End If
End If
If (tgAw) > (tgAk) Then
  If MsgBox("Tanggal awal yang anda masukkan lebih besar dari tanggal akhir periode!" & _
  vbCrLf & "Ulangi sekali lagi..?", vbYesNo + vbQuestion, "Laba Rugi") = vbYes Then
  GoTo Tgkaco:
  Else
  Exit Sub
  End If
End If

Screen.MousePointer = vbHourglass
Dim Rrpt As New CRAXDdRT.Application, LLap As CRAXDdRT.Report
Set LLap = Rrpt.OpenReport(App.Path & "\PenjItemSales.rpt")
Set fRpt = New Form2
Set fRpt.Report = LLap
fRpt.Report.FormulaFields(1).Text = Chr(34) & _
"Dari " & Format(tgAw, "dd mmm yyyy") & " s/d " & Format(tgAk, "dd mmm yyyy") & Chr(34)
'Dim ForX As CRAXdDRT.FormulaFieldDefinition
'For Each ForX In fRpt.Report.FormulaFields
'If ForX.Name = "{@DKTanggal}" Then ForX.Text = Chr(34) & _
'"Dari " & Format(tgAw, "dd mmm yyyy") & " s/d " & Format(tgAw, "dd mmm yyyy") & Chr(34)
'Next
fRpt.Report.Database.Tables(1).SetDataSource aData.DataLapSalesKomisi(tgAw, tgAk, "1"), 3
'fRpt.Report.Database.Tables(2).SetDataSource aData.DataLapSalesKomisi(tgAw, tgAk, "2"), 3
fRpt.aView.ReportSource = fRpt.Report
fRpt.aView.ViewReport
fRpt.Show
Screen.MousePointer = vbNormal
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub mnuSETPASS_Click()
FrmCPass.Show
End Sub

Private Sub mnuSPB_Click()
frmSrtJalan.Show
End Sub

Private Sub mnuTipe_Click()
fTipe.Show
fTipe.ZOrder 0
End Sub

Private Sub mnuToolTutupBuku_Click()
FrmTutupBuku.Show vbModal
End Sub

Private Sub mnuTransBatch_Click()
frmTransBatch.Show
frmTransBatch.ZOrder 0
End Sub

Private Sub mnutransBeli_Click()
Dim aBeli As New frmPembelian
aBeli.JenisJenis = "B"
aBeli.Show
End Sub


Private Sub mnuTransBFak_Click()
On Error GoTo aa
Call AturLaporan("BTransPerFak")
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuTransBFak_Click pada Form Utama"
End Sub

Private Sub mnuTransJFak_Click()
On Error GoTo aa
Call AturLaporan("JTransPerFak")
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur mnuTransJFak_Click pada Form Utama"
End Sub

Private Sub mnuTtrBeliInput_Click()
Dim aReturBeli As New frmPembelian
aReturBeli.JenisJenis = "RB"
aReturBeli.Show
End Sub

Private Sub mnuUtangBayar_Click()
Dim aUtPi As New frmPiutang
aUtPi.JenisUtPi = "Utang"
aUtPi.Show
End Sub

Private Sub mnuUtangReal_Click()
Dim aURe As New frmPiutangReal
aURe.JenisUtPi = "Utang"
aURe.Show
End Sub

Private Sub mnuUtangTT_Click()
Dim aUtPi As New frmTT
aUtPi.JenisJenis = "TTH"
aUtPi.Show
End Sub

Private Sub mnuWilayah_Click()
fWilayah.Show
fWilayah.ZOrder 0
End Sub

