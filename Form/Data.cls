VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const aDBPass As String = ""
Dim a As New ADODB.Connection


Private Sub Class_Initialize()
On Error GoTo aa
a.CursorLocation = adUseClient

a.Open Lokasi

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub
Private Sub Class_Terminate()
On Error GoTo aa
a.Close
Set a = Nothing
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Public Function AmbilData(Tabel As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset
k.CursorLocation = adUseClient
k.Open Tabel, a, adOpenStatic, adLockOptimistic
Set AmbilData = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function ExecCommand(CmdSetering As String) As String
On Error GoTo aa
a.Execute CmdSetering
ExecCommand = ""
Exit Function
aa:
ExecCommand = Err.Number & vbCrLf & Err.Description
End Function
Public Function AmbilCommand(CmdSetering As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
Set k = a.Execute(CmdSetering)
Set AmbilCommand = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function


Public Function SimpanBeli(Head() As String, Bodi() As String) As String
On Error GoTo aa
Dim aString As String, rsCek As New ADODB.Recordset
a.BeginTrans

aString = "Select [No Faktur] from TransBeli where [No Faktur]='" & Head(0) & "' and hapus<>0"
Set rsCek = a.Execute(aString)
If Not rsCek.EOF Then
  'Tanya Password
    'frmPass.Jenis = "Faktur"
    'frmPass.lblLabels(0).Caption = "No. Faktur yang anda masukkan telah pernah digunakan/dihapus.." & vbCrLf & _
    '"Untuk memakainya kembali harap isi password berikut dengan benar."
    'frmPass.Show vbModal
    'If Not frmPass.LoginOK Then
    'SimpanBeli = "Transaksi tidak dapat disimpan, no faktur telah ada.."
    'a.RollbackTrans
    'Exit Function
    'Else
  a.Execute "Delete from TransBeli where [No Faktur]='" & Head(0) & "'"
    'End If
End If

'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = Diskon    (5) = CaraPembayaran c/k
'(6) = LamaKredit (7) = jenistrans b/rb  (8) = Keterangan (9) = Total
aString = "INSERT INTO TransBeli( [No Faktur],Tgl,Kepada,Salesman,Diskon," & _
"Jenis,JatuhTempo,JenisTrans,Keterangan,nama,Total) " & _
"values('" & Head(0) & "','" & Head(1) & _
"','" & Head(2) & "','" & Head(3) & _
"','" & Head(4) & "','" & Head(5) & _
"','" & Head(6) & "','" & Head(7) & _
"','" & Head(8) & "','" & AmanOi(nAMA) & "','" & Head(9) & "')"
a.Execute aString

Dim i As Byte, b As New ADODB.Recordset
Set b = a.Execute("select intno from transbeli where [No Faktur]='" & Head(0) & "'")

For i = 1 To UBound(Bodi)
If Bodi(i, 0) <> "" Then

aString = "INSERT INTO DetailBeli (intTrans, [Kode Barang],harga,qty,diskon,QtyB,QtyS,Pengali) " & _
"values(" & b(0) & ",'" & _
Bodi(i, 0) & "','" & _
Bodi(i, 1) & "','" & _
Bodi(i, 2) & "','" & _
Bodi(i, 3) & "','" & _
Bodi(i, 5) & "','" & _
Bodi(i, 6) & "','" & _
Bodi(i, 4) & "')"
a.Execute aString
If Head(7) = "B" Or Head(7) = "RJ" Then
aString = "update barang set qty=qty+'" & Bodi(i, 2) & _
"' where [kode barang]='" & Bodi(i, 0) & "'"
ElseIf Head(7) = "RB" Or Head(7) = "J" Then
aString = "update barang set qty=qty-'" & Bodi(i, 2) & _
"' where [kode barang]='" & Bodi(i, 0) & "'"
End If
a.Execute aString
End If
Next i
a.CommitTrans
SimpanBeli = ""
Exit Function
aa:
a.RollbackTrans
SimpanBeli = Err.Number & Err.Description
End Function
Public Function HapusBeli(intNo As String, Jenis As String) As String
On Error GoTo aa
Dim aString As String
a.BeginTrans

Dim b As New ADODB.Recordset
Set b = a.Execute("Select bayar from inpututang where intno=" & intNo)
If Not b.EOF Then
If b!Bayar <> 0 Then Err.Raise vbObjectError + 123, "Hapus data Pembelian", "Data pembelian tersebut telah dibayar, harap cek kembali"
End If
aString = "update TransBeli set hapus=1 where [intno]=" & Val(intNo)
a.Execute aString
Set b = a.Execute("select intno,[Kode Barang],qty from detailbeli where [inttrans]=" & Val(intNo))
Do While Not b.EOF
aString = "update DetailBeli set qty=0 where intno=" & b!intNo
a.Execute aString
If Jenis = "B" Or Jenis = "RJ" Then
aString = "update barang set qty=qty-'" & b!qty & _
"' where [kode barang]='" & b![Kode Barang] & "'"
ElseIf Jenis = "RB" Or Jenis = "J" Then
aString = "update barang set qty=qty+'" & b!qty & _
"' where [kode barang]='" & b![Kode Barang] & "'"
End If
a.Execute aString
b.MoveNext
Loop
a.Execute "delete from transbeli where intno=" & Val(intNo)
a.CommitTrans
HapusBeli = ""
Exit Function
aa:
a.RollbackTrans
HapusBeli = Err.Number & Err.Description
End Function



Public Function SimpanJual(Head() As String, Bodi() As String) As String
On Error GoTo aa
Dim aString As String, rsCek As New ADODB.Recordset
a.BeginTrans
LagiSimpan = True
SimpanSalah = ""
aString = "Select [No Faktur] from TransJual where [No Faktur]='" & Head(0) & "' and hapus<>0"
Set rsCek = a.Execute(aString)
If Not rsCek.EOF Then
  'Tanya Password
    'frmPass.Jenis = "Faktur"
    'frmPass.lblLabels(0).Caption = "No. Faktur yang anda masukkan telah pernah digunakan/ dihapus.." & vbCrLf & _
    '"Untuk memakainya kembali harap isi password berikut dengan benar."
    'frmPass.Show vbModal
    'If Not frmPass.LoginOK Then
    'SimpanJual = "Transaksi tidak dapat disimpan, no faktur telah ada.."
    'a.RollbackTrans
    'Exit Function
    'Else
  a.Execute "Delete from TransJual where [No Faktur]='" & Head(0) & "'"
    'End If
End If

'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = Diskon    (5) = CaraPembayaran c/k
'(6) = LamaKredit (7) = jenistrans b/rb  (8) = Keterangan  (9) = Total
aString = "INSERT INTO TransJual( [No Faktur],Tgl,Kepada,Salesman,Diskon," & _
"Jenis,JatuhTempo,JenisTrans,Keterangan,nama,Total,Posting) " & _
"values('" & Head(0) & "','" & Head(1) & _
"','" & Head(2) & "','" & Head(3) & _
"','" & Head(4) & "','" & Head(5) & _
"','" & Head(6) & "','" & Head(7) & _
"','" & Head(8) & "','" & AmanOi(nAMA) & "','" & Head(9) & "'," & Head(10) & ")"
a.Execute aString

Dim i As Byte, b As New ADODB.Recordset
Set b = a.Execute("select intno from transjual where [No Faktur]='" & Head(0) & "'")

Dim hRt2 As Currency, LogPass As Boolean
LogPass = False
For i = 1 To UBound(Bodi)
If Bodi(i, 0) <> "" Then
'(0) = Kode  '(1) = Harga  '(2) = Qty  '(3) = Diskon

If Head(7) = "J" Then
  hRt2 = QtyRata(Bodi(i, 0), Bodi(i, 2))
Else
  hRt2 = Bodi(i, 1) - Bodi(i, 3)
End If

aString = "INSERT INTO DetailJual (intTrans, [Kode Barang],harga,qty,diskon,QtyB,QtyS,Pengali,HRata) " & _
"values(" & b(0) & ",'" & _
Bodi(i, 0) & "','" & _
Bodi(i, 1) & "','" & _
Bodi(i, 2) & "','" & _
Bodi(i, 3) & "','" & _
Bodi(i, 5) & "','" & _
Bodi(i, 6) & "','" & _
Bodi(i, 4) & "','" & _
IIf(hRt2 = -1, Val(Bodi(i, 1)) - Val(Bodi(i, 3)), hRt2) & "')"
  
  'If hRt2 = -1 Then
  '   SimpanJual = "Barang " & Bodi(i, 0) & " harga modalnya tidak dapat ditentukan.." & _
  '   vbCrLf & "Kemungkinan qty barang tersebut 0 atau tidak mencukupi.."
  '   a.RollbackTrans
  '   Exit Function
  'End If
 'Tanya Password penj dibawah modal
  'If ((Bodi(i, 1)) - Val(Bodi(i, 3))) < hRt2 And Head(7) = "J" Then
  '  If Not LogPass Then
  '    frmPass.Jenis = "Rugi"
  '    frmPass.lblLabels(0).Caption = "Barang " & Bodi(i, 0) & _
  '    " harga penjualannya dibawah modal.." & vbCrLf & _
  '    "HargaPenjualan=" & Val((Bodi(i, 1)) - Val(Bodi(i, 3))) & ", HargaModal=" & hRt2 & vbCrLf & "Untuk melanjutkan Transaksi masukkan password."
  '    frmPass.Show vbModal
  '      If Not frmPass.LoginOK Then
  '      SimpanJual = "Transaksi tidak dapat disimpan, ada penjualan yang dibawah modal.."
  '      a.RollbackTrans
  '      Exit Function
  '      Else
  '      LogPass = True
  '      End If
  '  Else
  '    If MsgBox("Barang " & Bodi(i, 0) & _
  '    "harga penjualannya dibawah modal.." & vbCrLf & _
  '    "HargaPenjualan=" & Val((Bodi(i, 1)) - Val(Bodi(i, 3))) & ", HargaModal=" & hRt2 & vbCrLf & _
  '    "Apakah anda ingin melanjutkan..?", vbInformation + vbYesNo, "Simpan Penjualan") = vbNo Then
  '      SimpanJual = "Transaksi tidak dapat disimpan, ada penjualan yang dibawah modal.."
  '      a.RollbackTrans
  '      Exit Function
  '    End If
  '  End If
  'End If

a.Execute aString

If Head(7) = "B" Or Head(7) = "RJ" Then
aString = "update barang set qty=qty+'" & Bodi(i, 2) & _
"' where [kode barang]='" & Bodi(i, 0) & "'"
ElseIf Head(7) = "RB" Or Head(7) = "J" Then
aString = "update barang set qty=qty-'" & Bodi(i, 2) & _
"' where [kode barang]='" & Bodi(i, 0) & "'"
End If
a.Execute aString
End If
Next i
'Set rsCek = a.Execute("CreditCheck '" & Head(2) & "'")
'If (Head(5) <> "C") Then
'If (rsCek![Credit Limit] >= rsCek!Sisa) Then
'a.CommitTrans
'SimpanJual = ""
'Else
'  'Tanya Password proses yg password bener
'  frmPass.Jenis = "Cust"
'  frmPass.lblLabels(0).Caption = "Kredit Limit konsumen yang anda masukkan telah melewati batas.." & vbCrLf & _
'  "Untuk menyimpan transksi ini, harap isi password berikut dengan benar."
'  frmPass.Show vbModal
'  If Not frmPass.LoginOK Then
'  SimpanJual = "Total transaksi melebihi credit limit.." & vbCrLf & "Harap hubungi administrator"
'  a.RollbackTrans
'  Exit Function
'  Else
'  a.CommitTrans
'  SimpanJual = ""
'  End If
'End If
'Else
a.CommitTrans
SimpanJual = ""
'End If
LagiSimpan = False
'If SimpanSalah <> "" Then _
'MsgBox "Qty tidak cukup pada barang : " & SimpanSalah
Exit Function
aa:
a.RollbackTrans
LagiSimpan = False
'If SimpanSalah <> "" Then _
'MsgBox "Qty tidak cukup pada barang : " & SimpanSalah
SimpanJual = Err.Number & Err.Description
End Function
Public Function HapusJual(intNo As String, Jenis As String) As String
On Error GoTo aa
Dim aString As String
a.BeginTrans
Dim b As New ADODB.Recordset
Set b = a.Execute("Select bayar from inputpiutang where intno=" & intNo)
If Not b.EOF Then
If b!Bayar <> 0 Then Err.Raise vbObjectError + 123, "Hapus data Penjualan", "Data penjualan tersebut telah dibayar, harap cek kembali"
End If
aString = "update TransJual set hapus=1 where [intno]=" & Val(intNo)
a.Execute aString
Set b = a.Execute("select intno,[Kode Barang],qty from detailJual where [inttrans]=" & Val(intNo))
Do While Not b.EOF
aString = "update DetailJual set qty=0 where intno=" & b!intNo
a.Execute aString
If Jenis = "B" Or Jenis = "RJ" Then
aString = "update barang set qty=qty-'" & b!qty & _
"' where [kode barang]='" & b![Kode Barang] & "'"
ElseIf Jenis = "RB" Or Jenis = "J" Then
aString = "update barang set qty=qty+'" & b!qty & _
"' where [kode barang]='" & b![Kode Barang] & "'"
End If
a.Execute aString
b.MoveNext
Loop
a.Execute "delete from transjual where intno=" & Val(intNo)
a.CommitTrans
HapusJual = ""
Exit Function
aa:
a.RollbackTrans
HapusJual = Err.Number & Err.Description
End Function


Public Function SimpanUtangPiutang(Head() As String, BodiA As VSFlexGrid, BodiB As VSFlexGrid, Jenis As String, PotHar As VSFlexGrid) As String
On Error GoTo aa
Dim aString As String, xTabel As String
a.BeginTrans
If Jenis = "Utang" Then
xTabel = "Utang"
Else
xTabel = "Piutang"
End If

'Head=>  0=NoBayar  1=Tgl  2=Supplier  3=Total  4=Keterangan
'BodiA=> 0=Pembayaran([T]unai [G]iro [B]ank/lain)   1=NoGiro
'        2=Bank   3=JatuhTempo    4=SubTotal    5=Keterangan
'BodiB=> 0=NoFaktur(nga pake)   1=TotalFaktur(nga pake)
'        2=TotalBayar    3=Potongan(dlm %/Rp nga pake cuma tampilan)
'        4=Tampilan%      5=Sisa UtangUtang(Nga pake)
'        6=SubTotalBayar   7=intNoFaktur   8=diskon(yg dipake)
'PotHar=>0=intNoFaktur  1=intNoDetailTrans  2=Potongan
aString = "INSERT INTO " & xTabel & "(KodeBayar , tgl, Kepada, Total, nama, Keterangan) " & _
"values('" & Head(0) & "','" & Head(1) & _
"','" & Head(2) & "','" & Head(3) & _
 "','" & AmanOi(nAMA) & "','" & Head(4) & "')"
a.Execute aString
Dim i As Byte, b As New ADODB.Recordset, intNoUtPi As Currency
Set b = a.Execute("select intno from " & xTabel & " where [KodeBayar]='" & Head(0) & "'")
intNoUtPi = b(0)
For i = 1 To BodiA.Rows - 1
If BodiA.TextMatrix(i, 0) <> "" Then
aString = "INSERT INTO " & xTabel & "Anak (KodeBayar,JenisBayar,Total,JatuhTempo,NoGiro,NamaBank,Keterangan) " & _
"values(" & intNoUtPi & ",'" & BodiA.TextMatrix(i, 0) & "','" & _
BodiA.TextMatrix(i, 4) & "','" & IIf(BodiA.TextMatrix(i, 3) = "", Date, BodiA.TextMatrix(i, 3)) & "','" & _
BodiA.TextMatrix(i, 1) & "','" & BodiA.TextMatrix(i, 2) & "','" & _
BodiA.TextMatrix(i, 5) & "')"
a.Execute aString
  If BodiA.TextMatrix(i, 0) = "LB" Then
  a.Execute "Update " & IIf(Jenis = "Utang", "Supplier", "Konsumen") & _
  " set [Lebih Bayar]=[Lebih Bayar]-'" & BodiA.ValueMatrix(i, 4) & _
  "' where Kode='" & Head(2) & "'"
  a.Execute "update " & xTabel & "Anak set [Lebih Bayar]='" & -(BodiA.ValueMatrix(i, 4)) & _
  "' where JenisBayar='LB' and KodeBayar=" & intNoUtPi
  End If
End If
Next i

Set b = a.Execute("select intno,KodeBayar,Total from " & xTabel & "Anak where [KodeBayar]=" & intNoUtPi & "")
Dim TSemen As Currency, UdaMsk As Currency, UdaDisc As Currency
For i = 1 To BodiB.Rows - 1
If BodiB.TextMatrix(i, 0) = "" Then BodiB.RemoveItem i
Next i
'        6=SubTotalBayar   7=intNoFaktur   8=diskon(yg dipake)
    TSemen = b!Total
yyy:
  If BodiB.Rows = 1 And TSemen <> 0 Then
    a.Execute "Update " & IIf(Jenis = "Utang", "Supplier", "Konsumen") & _
    " set [Lebih Bayar]=[Lebih Bayar]+'" & TSemen & "' where Kode='" & Head(2) & "'"
    a.Execute "Update " & IIf(Jenis = "Utang", "Utang", "Piutang") & _
    "Anak set [Lebih Bayar]='" & TSemen & "' where intno=" & b!intNo
    TSemen = 0
    GoTo xxx
  End If
  If TSemen > BodiB.TextMatrix(1, 6) Then
    TSemen = TSemen - BodiB.TextMatrix(1, 6)
    UdaMsk = BodiB.TextMatrix(1, 6)
    UdaDisc = BodiB.TextMatrix(1, 8)
    BodiB.TextMatrix(1, 6) = 0
  Else
    BodiB.TextMatrix(1, 6) = BodiB.TextMatrix(1, 6) - TSemen
    UdaDisc = BodiB.TextMatrix(1, 8)
    BodiB.TextMatrix(1, 8) = 0
    UdaMsk = TSemen
    TSemen = 0
  End If
aString = "INSERT INTO Detail" & xTabel & " (KodeBayar,KodeFaktur,Total,Potongan) " & _
"values(" & b!intNo & ",'" & BodiB.TextMatrix(1, 7) & "','" & _
UdaMsk & "','" & UdaDisc & "')"
a.Execute aString
    If BodiB.TextMatrix(1, 6) = 0 Then BodiB.RemoveItem 1
xxx:
    If TSemen = 0 Then
    b.MoveNext
      If b.EOF Then GoTo zzz
    TSemen = b!Total
    End If
GoTo yyy
zzz:

'//Potongan Harga//
If Jenis = "Piutang" Then
Set b = a.Execute("SELECT Piutang.intno, PiutangAnak.intno, DetailPiutang.KodeFaktur " & _
"FROM Piutang INNER JOIN (PiutangAnak INNER JOIN DetailPiutang ON PiutangAnak.intno = " & _
"DetailPiutang.KodeBayar) ON Piutang.intno = PiutangAnak.KodeBayar " & _
"where piutang.intno=" & intNoUtPi)
Else
Set b = a.Execute("SELECT Utang.intno, UtangAnak.intno, DetailUtang.KodeFaktur " & _
"FROM Utang INNER JOIN (UtangAnak INNER JOIN DetailUtang ON UtangAnak.intno = " & _
"DetailUtang.KodeBayar) ON Utang.intno = UtangAnak.KodeBayar " & _
"where Utang.intno=" & intNoUtPi)
End If
Dim kPotX As Single
Do While Not b.EOF
  For kPotX = 1 To PotHar.Rows - 1
    If PotHar.ValueMatrix(kPotX, 0) = b!KodeFaktur Then
    a.Execute "update detail" & IIf(Jenis = "Piutang", "jual", "beli") & " set ph=ph+'" & PotHar.ValueMatrix(kPotX, 2) & _
    "' where intno=" & PotHar.ValueMatrix(kPotX, 1)
    End If
  Next kPotX
  b.MoveNext
Loop

Dim kSS As Currency
Set b = a.Execute("SELECT sum(Detail" & xTabel & ".total) FROM " & xTabel & " INNER JOIN (" & xTabel & "Anak INNER JOIN Detail" & xTabel & " ON " & xTabel & "Anak.intno=Detail" & xTabel & ".KodeBayar) ON " & xTabel & ".intno=" & xTabel & "Anak.KodeBayar " & _
"where " & xTabel & "Anak.kodebayar=" & intNoUtPi & " group by " & xTabel & ".kodebayar")
kSS = b(0)
Set b = a.Execute("select sum(total) from " & xTabel & "Anak where kodebayar=" & intNoUtPi & _
" group by kodebayar")
kSS = kSS - b(0)
If kSS > 0 Then Err.Raise vbObjectError + 8, "Simpan Utang", "Pembayaran dan pelunasan " & xTabel & " tidak balance"

'*************Cash langsung OK statusnya*****************
a.Execute "update " & xTabel & "Anak set status='OK', keterangan=keterangan & ' OK'," & _
"tglstatus='" & SerbaGuna.AmanTgl(Date) & "' where jenisbayar='T' and kodebayar=" & intNoUtPi
'********************************************************
a.Execute "update " & xTabel & "Anak set status='OK', keterangan=keterangan & ' OK'," & _
"tglstatus='" & SerbaGuna.AmanTgl(Date) & "' where jenisbayar='LB' and kodebayar=" & intNoUtPi
'********************************************************
a.CommitTrans
SimpanUtangPiutang = ""
Call UtangPiutangTransOK(Jenis)
Exit Function
aa:
a.RollbackTrans
SimpanUtangPiutang = Err.Number & Err.Description
End Function

Public Function UtangPiutangOK(Kode As String, Ket As String, Jenis As String) As String
On Error GoTo aa
Dim aString As String, xTabel As String
a.BeginTrans
If Jenis = "Utang" Then
xTabel = "Utang"
Else
xTabel = "Piutang"
End If
aString = "update " & xTabel & "Anak set status='OK', keterangan=keterangan & ' OK',tglstatus='" & SerbaGuna.AmanTgl(Date) & _
"', ketstatus='" & Ket & "' where intNo=" & Kode & " "
a.Execute aString
a.CommitTrans
UtangPiutangOK = ""
Exit Function
aa:
a.RollbackTrans
UtangPiutangOK = Err.Number & Err.Description
End Function

Public Function UtangPiutangHapus(Jenis As String, Kode As String) As String
On Error GoTo aa
Dim kRec As New ADODB.Recordset, intNoAnak As Single, Siapa As String
a.BeginTrans

If Jenis = "Utang" Then
  Set kRec = a.Execute("SELECT Utang.KodeBayar as A, UtangAnak.KodeBayar as B, Utang.Kepada, " & _
  "DetailUtang.KodeFaktur FROM (Utang INNER JOIN UtangAnak ON Utang.intno  " & _
  "= UtangAnak.KodeBayar) INNER JOIN DetailUtang ON UtangAnak.intno =  " & _
  "DetailUtang.KodeBayar where utang.kodeBayar='" & Kode & "'")
  Do While Not kRec.EOF
    intNoAnak = kRec!b
    Siapa = kRec!Kepada
    a.Execute ("update transbeli set jenis='K' where intno=" & kRec!KodeFaktur)
    kRec.MoveNext
  Loop
Else
  Set kRec = a.Execute("SELECT Piutang.KodeBayar as A, PiutangAnak.KodeBayar as B, Piutang.Kepada, " & _
  "DetailPiutang.KodeFaktur FROM (Piutang INNER JOIN PiutangAnak ON Piutang.intno  " & _
  "= PiutangAnak.KodeBayar) INNER JOIN DetailPiutang ON PiutangAnak.intno =  " & _
  "DetailPiutang.KodeBayar where Piutang.kodeBayar='" & Kode & "'")
  Do While Not kRec.EOF
    intNoAnak = kRec!b
    Siapa = kRec!Kepada
    a.Execute ("update transjual set jenis='K' where intno=" & kRec!KodeFaktur)
    kRec.MoveNext
  Loop

End If

'--Kurangi Lebih Bayar--
Dim rsLbh As New ADODB.Recordset
Set rsLbh = a.Execute("Select intNo,[Lebih Bayar] from " & _
IIf(Jenis = "Utang", "Utang", "Piutang") & "Anak where KodeBayar=" & intNoAnak)
Do While Not rsLbh.EOF
a.Execute "Update " & IIf(Jenis = "Utang", "Supplier", "Konsumen") & _
" set [Lebih Bayar]=[Lebih Bayar]-'" & rsLbh![Lebih Bayar] & "' where Kode='" & AmanOi(Siapa) & "'"
rsLbh.MoveNext
Loop
'--
  a.Execute "delete from " & IIf(Jenis = "Utang", "Utang", "Piutang") & _
  " where kodeBayar='" & Kode & "'"
a.CommitTrans
UtangPiutangHapus = ""
Exit Function
aa:
a.RollbackTrans
UtangPiutangHapus = "Error:" & Err.Number & " (" & Err.Description & ") di prosedur UtangPiutangHapus pada Class Module Data"
End Function

Public Function UtangPiutangBatal(Kode As String, Ket As String, Jenis As String) As String
On Error GoTo aa
Dim aString As String, xTabel As String
a.BeginTrans
If Jenis = "Utang" Then
xTabel = "Utang"
Else
xTabel = "Piutang"
End If
aString = "update " & xTabel & "Anak set status='Batal', keterangan=keterangan & ' Batal', tglstatus='" & SerbaGuna.AmanTgl(Date) & _
"', ketstatus='" & Ket & "' where intno=" & Kode & " "
a.Execute aString

'--Cek batal lebih bayar--
Dim rsLbh As New ADODB.Recordset, BnykLebih As Currency
Set rsLbh = a.Execute("select iif(isnull([Lebih Bayar]),0,[Lebih Bayar]) as [Lebih Bayar],KodeBayar from " & _
xTabel & "Anak where intno=" & Kode & " ")
If rsLbh![Lebih Bayar] <> 0 Then
BnykLebih = rsLbh![Lebih Bayar]
a.Execute "Update " & IIf(Jenis = "Utang", "Utang", "Piutang") & _
"Anak set [Lebih Bayar]=0 where intno=" & Kode
Set rsLbh = a.Execute("select Kepada from " & _
xTabel & " where intno=" & rsLbh!KodeBayar & " ")
a.Execute "Update " & IIf(Jenis = "Utang", "Supplier", "Konsumen") & _
" set [Lebih Bayar]=[Lebih Bayar]-'" & BnykLebih & "' where Kode='" & AmanOi(rsLbh!Kepada) & "'"
End If

a.CommitTrans
UtangPiutangBatal = ""
Exit Function
aa:
a.RollbackTrans
UtangPiutangBatal = Err.Number & Err.Description
End Function

Public Function UtangPiutangTransOK(Jenis As String) As String
On Error GoTo aa
Dim aString As String, xTabel As String
Dim b As New ADODB.Recordset, intBayar As String
a.BeginTrans

If Jenis = "Utang" Then
aString = "SELECT TransBeli.intno, TransBeli.JenisTrans, TransBeli.Total-transBeli.diskon AS Total, " & _
"Sum([DetailUtang.Total]+[Potongan]) AS TotBayar, " & _
"Sum(DetailUtang.Potongan) AS TotPot, " & _
"DateDiff('d',TransBeli.Tgl,Max(UtangAnak.TglStatus)) AS TglOK, TransBeli.Jenis " & _
"FROM TransBeli INNER JOIN (UtangAnak INNER JOIN DetailUtang ON " & _
"UtangAnak.intno = DetailUtang.KodeBayar) ON TransBeli.intno = " & _
"DetailUtang.KodeFaktur " & _
"GROUP BY TransBeli.intno, TransBeli.JenisTrans,TransBeli.Jenis, TransBeli.Total, TransBeli.Diskon,  " & _
"UtangAnak.Status, TransBeli.Tgl, TransBeli.Hapus, TransBeli.Jenis " & _
"HAVING (((UtangAnak.Status)='OK') AND ((TransBeli.Hapus)=0)) AND ((TransBeli.Jenis)<>'OK');"
Else
aString = "SELECT TransJual.intno, TransJual.JenisTrans, TransJual.Total-transjual.diskon AS Total, " & _
"Sum([DetailPiutang.Total]+[Potongan]) AS TotBayar, " & _
"Sum(DetailPiutang.Potongan) AS TotPot,  " & _
"DateDiff('d',TransJual.Tgl,Max(PiutangAnak.TglStatus)) AS TglOK, TransJual.Jenis " & _
"FROM TransJual INNER JOIN (PiutangAnak INNER JOIN DetailPiutang ON " & _
"PiutangAnak.intno = DetailPiutang.KodeBayar) ON TransJual.intno = " & _
"DetailPiutang.KodeFaktur " & _
"GROUP BY TransJual.intno, TransJual.Total, TransJual.Diskon, " & _
"PiutangAnak.Status, TransJual.Jenis, TransJual.Tgl, TransJual.Hapus, TransJual.JenisTrans " & _
"HAVING (((PiutangAnak.Status)='OK')  " & _
"AND ((TransJual.Jenis)<>'OK') AND ((TransJual.Hapus)=0));"
End If
Set b = a.Execute(aString)

Do While Not b.EOF
  If Jenis = "Utang" Then
  If IIf(b!JenisTrans = "B", b!Total, -b!Total) = b!TotBayar Then
    aString = "update transBeli set jenis='OK' where intno=" & b!intNo
    a.Execute aString
  End If
  Else
  If IIf(b!JenisTrans = "J", b!Total, -b!Total) = b!TotBayar Then
    aString = "update transJual set jenis='OK' where intno=" & b!intNo
    a.Execute aString
    a.Execute "Insert into JualSales values(" & b!intNo & _
    ",'" & b!TotPot & "','" & b!tglOK & "')"
  End If
  End If
b.MoveNext
Loop

a.CommitTrans
UtangPiutangTransOK = ""
Exit Function
aa:
a.RollbackTrans
UtangPiutangTransOK = Err.Number & Err.Description
End Function






'Laporan*******************************************************
'**************************************************************
Public Function DataLapFaktur(Jenis As String, Optional No As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "B" Then
CmadString = "SELECT * from xBeliDet"
ElseIf Jenis = "A" Then
CmadString = "SELECT * from xBeliDet union SELECT * from xJualDet order by tgl"
Else
CmadString = "SELECT * from xJualDet"
End If
If No <> "" Then
CmadString = CmadString & " where [no faktur]='" & No & "'"
End If
Set k = a.Execute(CmadString)
Set DataLapFaktur = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapBJ_Supp(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "B" Then
CmadString = "SELECT * FROM xBeliGen"
Else
CmadString = "SELECT * FROM xJualGen"
End If
Set k = a.Execute(CmadString)
Set DataLapBJ_Supp = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapBJ_ItemBarang(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "B" Then
CmadString = "Select * from xBeliDet"
Else
CmadString = "SELECT * from xJualDet"
End If
Set k = a.Execute(CmadString)
Set DataLapBJ_ItemBarang = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapBJ_Total(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "B" Then
CmadString = "SELECT * from xBeliGen WHERE jenistrans='B'"
ElseIf Jenis = "RB" Then
CmadString = "SELECT * from xBeliGen WHERE jenistrans='RB'"
ElseIf Jenis = "J" Then
CmadString = "SELECT * from xJualGen WHERE jenistrans='J'"
ElseIf Jenis = "RJ" Then
CmadString = "SELECT * from xJualGen WHERE jenistrans='RJ'"
End If
Set k = a.Execute(CmadString)
Set DataLapBJ_Total = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapBJ_Tgl(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
CmadString = "SELECT TransBeli.Tgl, DetailBeli.[Kode Barang], Barang.[Nama Barang], Sum(DetailBeli.Qty) AS Qty, sum(DetailBEli.harga-DetailBeli.Diskon) as Harga " & _
"FROM TransBeli INNER JOIN (Barang INNER JOIN DetailBeli ON Barang.[Kode Barang] = DetailBeli.[Kode Barang]) ON TransBeli.intno = DetailBeli.intTrans " & _
"Where hapus=0 and transbeli.jenistrans='B' " & _
"GROUP BY TransBeli.Tgl, DetailBeli.[Kode Barang], Barang.[Nama Barang];"
Set k = a.Execute(CmadString)
Set DataLapBJ_Tgl = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function

Public Function DataLapJ_Sales(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
CmadString = "SELECT * from xJualGen"
Set k = a.Execute(CmadString)
Set DataLapJ_Sales = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function



Public Function DataLapHP_Bayar(Jenis As String, Tgl As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "H" Then
CmadString = "SELECT 'H' as Judul, Utang.intno, Utang.KodeBayar, Utang.tgl, 'U' as JenisTrans,Utang.Kepada, Supplier.Nama, Supplier.Alamat, " & _
"Utang.Total, Utang.Keterangan,UtangAnak.JenisBayar,UtangAnak.NoGiro, UtangAnak.NamaBank, UtangAnak.JatuhTempo,UtangAnak.Total as AnakTotal, " & _
"UtangAnak.Status, UtangAnak.TglStatus,UtangAnak.Keterangan as AnakKeterangan,UtangAnak.KetStatus, TransBeli.[No Faktur],DetailUtang.Total as TotalDetail, " & _
"DetailUtang.Potongan, UtangAnak.intno  as AnakIntNo FROM (((Supplier INNER JOIN TransBeli ON Supplier.Kode = TransBeli.Kepada) INNER JOIN Utang ON Supplier.Kode = Utang.Kepada) INNER JOIN UtangAnak ON Utang.intno = UtangAnak.KodeBayar) INNER JOIN DetailUtang ON (TransBeli.intno = DetailUtang.KodeFaktur) AND (UtangAnak.intno = DetailUtang.KodeBayar)"
Else
CmadString = "SELECT 'P' as Judul, Piutang.intno, Piutang.KodeBayar, Piutang.tgl, 'P' as JenisTrans,Piutang.Kepada, Konsumen.Nama, Konsumen.Alamat, " & _
"Piutang.Total, Piutang.Keterangan,PiutangAnak.JenisBayar,PiutangAnak.NoGiro, PiutangAnak.NamaBank, PiutangAnak.JatuhTempo, " & _
"PiutangAnak.Total as AnakTotal,PiutangAnak.Status, PiutangAnak.TglStatus,PiutangAnak.Keterangan as AnakKeterangan, " & _
"PiutangAnak.KetStatus, TransJual.[No Faktur],DetailPiutang.Total as TotalDetail,DetailPiutang.Potongan, PiutangAnak.intno  as AnakIntNo " & _
"FROM (((Konsumen INNER JOIN TransJual ON Konsumen.Kode = TransJual.Kepada) INNER JOIN Piutang ON Konsumen.Kode = Piutang.Kepada) INNER JOIN PiutangAnak ON Piutang.intno = PiutangAnak.KodeBayar) INNER JOIN DetailPiutang ON (TransJual.intno = DetailPiutang.KodeFaktur) AND (PiutangAnak.intno = DetailPiutang.KodeBayar)"
End If
If Tgl <> "" Then
CmadString = CmadString & " where " & IIf(Jenis = "H", "Utang", "Piutang") & ".tgl between " & Tgl
End If
Set k = a.Execute(CmadString)
Set DataLapHP_Bayar = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapHP_Giro(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "H" Then
CmadString = "SELECT Utang.KodeBayar, Utang.tgl, Utang.Kepada+'-'+Supplier.Nama AS Kepada, UtangAnak.JenisBayar, UtangAnak.JatuhTempo, UtangAnak.NoGiro, UtangAnak.NamaBank, UtangAnak.Total, detailutang.Total AS Detail_Total, detailutang.Potongan, 'H' AS Jenis " & _
"FROM (Supplier INNER JOIN (Utang INNER JOIN UtangAnak ON Utang.intno = UtangAnak.KodeBayar) ON Supplier.Kode = Utang.Kepada) INNER JOIN detailutang ON UtangAnak.intno = detailutang.KodeBayar " & _
"WHERE ((([status])='') AND (([jenisbayar])='G' Or ([jenisbayar])='B'));"
Else
CmadString = "SELECT Piutang.KodeBayar, Piutang.tgl, Piutang.Kepada+'-'+Konsumen.Nama AS Kepada, PiutangAnak.JenisBayar, PiutangAnak.JatuhTempo, PiutangAnak.NoGiro, PiutangAnak.NamaBank, PiutangAnak.Total, detailPiutang.Total AS Detail_Total, detailPiutang.Potongan, 'P' AS Jenis " & _
"FROM (Konsumen INNER JOIN (Piutang INNER JOIN PiutangAnak ON Piutang.intno = PiutangAnak.KodeBayar) ON Konsumen.Kode = Piutang.Kepada) INNER JOIN detailPiutang ON PiutangAnak.intno = detailPiutang.KodeBayar " & _
"WHERE ((([status])='') AND (([jenisbayar])='G' Or ([jenisbayar])='B'));"
End If
Set k = a.Execute(CmadString)
Set DataLapHP_Giro = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapHP_Daftar(Jenis As String, TampilUdaBayar As Boolean) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "HDaftar" Then
CmadString = "SELECT 'H' as Judul, TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada, Supplier.Nama AS Nama, IIf(TransBeli.JenisTrans='B','PemBelian','Retur PemBelian') AS Transaksi, TransBeli.Tgl, IIf(transBeli.jenistrans='B',TransBeli.Total-transBeli.diskon,-(TransBeli.Total-transBeli.diskon)) AS Total, IIf(IsNull(Sum(UtangStatusNoBatal.Total+UtangStatusNoBatal.Potongan)),0,Sum(UtangStatusNoBatal.Total+UtangStatusNoBatal.Potongan)) AS Bayar, total-bayar AS Sisa " & _
"FROM (Supplier INNER JOIN TransBeli ON Supplier.Kode=TransBeli.Kepada) LEFT JOIN UtangStatusNoBatal ON TransBeli.intno=UtangStatusNoBatal.KodeFaktur " & _
"WHERE TransBeli.Jenis<>'C' AND TransBeli.Hapus=0 " & _
"GROUP BY TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada, Supplier.Nama, TransBeli.Tgl, TransBeli.JenisTrans, TransBeli.Diskon, TransBeli.JatuhTempo, TransBeli.Jenis, TransBeli.Total " & _
"ORDER BY TransBeli.Kepada, TransBeli.[No Faktur];"
ElseIf Jenis = "PDaftar" Then
CmadString = "SELECT 'P' as Judul, TransJual.intno, TransJual.[No Faktur], TransJual.Kepada, Konsumen.Nama AS Nama, IIf(TransJual.JenisTrans='J','Penjualan','Retur Penjualan') AS Transaksi, TransJual.Tgl, IIf(transJual.jenistrans='J',TransJual.Total-transjual.diskon,-(TransJual.Total-transjual.diskon)) AS Total, IIf(IsNull(Sum(PiutangStatusNoBatal.Total+PiutangStatusNoBatal.Potongan)),0,Sum(PiutangStatusNoBatal.Total+PiutangStatusNoBatal.Potongan)) AS Bayar, total-bayar AS Sisa " & _
"FROM Konsumen INNER JOIN (TransJual LEFT JOIN PiutangStatusNoBatal ON TransJual.intno = PiutangStatusNoBatal.KodeFaktur) ON Konsumen.Kode = TransJual.Kepada " & _
"WHERE TransJual.Jenis<>'C' AND (TransJual.Hapus)=0 " & _
"GROUP BY TransJual.intno, TransJual.[No Faktur], TransJual.Kepada, Konsumen.Nama, TransJual.Tgl, TransJual.JenisTrans, TransJual.Diskon, TransJual.JatuhTempo, TransJual.Jenis, TransJual.Total " & _
"ORDER BY TransJual.Kepada, TransJual.[No Faktur];"
ElseIf Jenis = "HDaftarNoGiro" Then
CmadString = "SELECT DISTINCTROW 'HG' as Judul, TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada, Supplier.Nama AS Nama, IIf(TransBeli.JenisTrans='B','PemBelian','Retur PemBelian') AS Transaksi, TransBeli.Tgl, IIf(transBeli.jenistrans='B',TransBeli.Total-transBeli.diskon,-(TransBeli.Total-transBeli.diskon)) AS Total, IIf(IsNull(Sum(UtangStatusOK.Total+UtangStatusOK.Potongan)),0,Sum(UtangStatusOK.Total+UtangStatusOK.Potongan)) AS Bayar, total-bayar AS Sisa " & _
"FROM (Supplier INNER JOIN TransBeli ON Supplier.Kode=TransBeli.Kepada) LEFT JOIN UtangStatusOK ON TransBeli.intno=UtangStatusOK.KodeFaktur " & _
"WHERE TransBeli.Jenis<>'C'  AND TransBeli.Hapus=0 " & _
"GROUP BY TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada, Supplier.Nama, TransBeli.Tgl, TransBeli.JenisTrans, TransBeli.Diskon, TransBeli.JatuhTempo, TransBeli.Jenis, TransBeli.Total " & _
"ORDER BY TransBeli.Kepada, TransBeli.[No Faktur];"
ElseIf Jenis = "PDaftarNoGiro" Then
CmadString = "SELECT DISTINCTROW 'PG' as Judul, TransJual.intno, TransJual.[No Faktur], TransJual.Kepada, Konsumen.Nama AS Nama, IIf(TransJual.JenisTrans='J','Penjualan','Retur Penjualan') AS Transaksi, TransJual.Tgl, IIf(transJual.jenistrans='J',TransJual.Total-transjual.diskon,-(TransJual.Total-transjual.diskon)) AS Total, IIf(IsNull(Sum(PiutangStatusOK.Total+PiutangStatusOK.Potongan)),0,Sum(PiutangStatusOK.Total+PiutangStatusOK.Potongan)) AS Bayar, total-bayar AS Sisa " & _
"FROM Konsumen INNER JOIN (TransJual LEFT JOIN PiutangStatusOK ON TransJual.intno = PiutangStatusOK.KodeFaktur) ON Konsumen.Kode = TransJual.Kepada " & _
"WHERE TransJual.Jenis<>'C' AND TransJual.Hapus=0 " & _
"GROUP BY TransJual.intno, TransJual.[No Faktur], TransJual.Kepada, Konsumen.Nama, TransJual.Tgl, TransJual.JenisTrans, TransJual.Diskon, TransJual.JatuhTempo, TransJual.Jenis, TransJual.Total " & _
"ORDER BY TransJual.Kepada, TransJual.[No Faktur];"
End If
If TampilUdaBayar = False Then
CmadString = Replace(CmadString, "WHERE ", "WHERE Trans" & IIf(Left(Jenis, 1) = "H", "Beli", "Jual") & ".Jenis<>'OK' AND ")
End If
Set k = a.Execute(CmadString)
Set DataLapHP_Daftar = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapHP_Daftar_Rinci(Jenis As String, NoFak As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "H" Then
CmadString = "SELECT Utang.intno, Utang.KodeBayar, Utang.tgl, 'P' as JenisTrans, " & _
"Utang.Kepada, Supplier.Nama, Supplier.Alamat, " & _
"iif(transbeli.jenistrans='B',TransBeli.total-TransBeli.diskon,-(TransBeli.total-TransBeli.diskon)) as TotalFaktur, " & _
"TransBeli.tgl As TglFaktur,Utang.Total, Utang.Keterangan,UtangAnak.JenisBayar, " & _
"UtangAnak.NoGiro, UtangAnak.NamaBank, UtangAnak.JatuhTempo,UtangAnak.Total as AnakTotal, " & _
"UtangAnak.Status, UtangAnak.TglStatus,UtangAnak.Keterangan as AnakKeterangan, " & _
"UtangAnak.KetStatus, TransBeli.[No Faktur],DetailUtang.Total as TotalDetail, " & _
"DetailUtang.Potongan, UtangAnak.intno  as AnakIntNo  " & _
"FROM (((Supplier INNER JOIN TransBeli ON Supplier.Kode = TransBeli.Kepada) INNER JOIN  " & _
"Utang ON Supplier.Kode = Utang.Kepada) INNER JOIN UtangAnak ON  " & _
"Utang.intno = UtangAnak.KodeBayar) INNER JOIN DetailUtang ON  " & _
"(TransBeli.intno = DetailUtang.KodeFaktur) AND (UtangAnak.intno = DetailUtang.KodeBayar) " & _
"where TransBeli.[No Faktur]='" & AmanOi(NoFak) & "'"
ElseIf Jenis = "P" Then
CmadString = "SELECT Piutang.intno, Piutang.KodeBayar, Piutang.tgl, 'P' as JenisTrans, " & _
"Piutang.Kepada, Konsumen.Nama, Konsumen.Alamat, " & _
"iif(transjual.jenistrans='J',TransJual.total-TransJual.diskon,-(TransJual.total-TransJual.diskon)) as TotalFaktur, " & _
"TransJual.tgl As TglFaktur,Piutang.Total, Piutang.Keterangan,PiutangAnak.JenisBayar, " & _
"PiutangAnak.NoGiro, PiutangAnak.NamaBank, PiutangAnak.JatuhTempo,PiutangAnak.Total as AnakTotal, " & _
"PiutangAnak.Status, PiutangAnak.TglStatus,PiutangAnak.Keterangan as AnakKeterangan, " & _
"PiutangAnak.KetStatus, TransJual.[No Faktur],DetailPiutang.Total as TotalDetail, " & _
"DetailPiutang.Potongan, PiutangAnak.intno  as AnakIntNo  " & _
"FROM (((Konsumen INNER JOIN TransJual ON Konsumen.Kode = TransJual.Kepada) INNER JOIN  " & _
"Piutang ON Konsumen.Kode = Piutang.Kepada) INNER JOIN PiutangAnak ON  " & _
"Piutang.intno = PiutangAnak.KodeBayar) INNER JOIN DetailPiutang ON  " & _
"(TransJual.intno = DetailPiutang.KodeFaktur) AND (PiutangAnak.intno = DetailPiutang.KodeBayar) " & _
"where TransJual.[No Faktur]='" & AmanOi(NoFak) & "'"
End If
Set k = a.Execute(CmadString)
Set DataLapHP_Daftar_Rinci = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function
Public Function DataLapHP_Jatuh(Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
If Jenis = "H" Then
CmadString = "SELECT DISTINCTROW 'H' as Judul, TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada, Supplier.Nama AS Nama, IIf(TransBeli.JenisTrans='J','Pembelian','Retur Pembelian') AS Transaksi, TransBeli.Tgl, transBeli.jatuhtempo, IIf(transBeli.jenistrans='B',TransBeli.Total-transBeli.diskon,-(TransBeli.Total-transBeli.diskon)) AS Total, IIf(IsNull(Sum(UtangStatusNoBatal.Total+UtangStatusNoBatal.Potongan)),0,Sum(UtangStatusNoBatal.Total+UtangStatusNoBatal.Potongan)) AS Bayar, total-bayar AS Sisa " & _
"FROM Supplier INNER JOIN (TransBeli LEFT JOIN UtangStatusNoBatal ON TransBeli.intno = UtangStatusNoBatal.KodeFaktur) ON Supplier.Kode = TransBeli.Kepada " & _
"WHERE (((TransBeli.Jenis)<>'C' And (TransBeli.Jenis)<>'OK') AND ((TransBeli.Hapus)=0)) and (date()-(transBeli.tgl+transBeli.jatuhtempo))>0 " & _
"GROUP BY TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada, Supplier.Nama, TransBeli.Tgl, TransBeli.JenisTrans, TransBeli.Diskon, TransBeli.JatuhTempo, TransBeli.Jenis, TransBeli.Total " & _
"ORDER BY TransBeli.Kepada, TransBeli.[No Faktur];"
Else
CmadString = "SELECT DISTINCTROW 'P' as Judul, TransJual.intno, TransJual.[No Faktur], TransJual.Kepada, Konsumen.Nama AS Nama, IIf(TransJual.JenisTrans='J','Penjualan','Retur Penjualan') AS Transaksi, TransJual.Tgl, transjual.jatuhtempo, IIf(transJual.jenistrans='J',TransJual.Total-transjual.diskon,-(TransJual.Total-transjual.diskon)) AS Total, IIf(IsNull(Sum(PiutangStatusNoBatal.Total+PiutangStatusNoBatal.Potongan)),0,Sum(PiutangStatusNoBatal.Total+PiutangStatusNoBatal.Potongan)) AS Bayar, total-bayar AS Sisa " & _
"FROM Konsumen INNER JOIN (TransJual LEFT JOIN PiutangStatusNoBatal ON TransJual.intno = PiutangStatusNoBatal.KodeFaktur) ON Konsumen.Kode = TransJual.Kepada " & _
"WHERE (((TransJual.Jenis)<>'C' And (TransJual.Jenis)<>'OK') AND ((TransJual.Hapus)=0)) and (date()-(transjual.tgl+transjual.jatuhtempo))>0 " & _
"GROUP BY TransJual.intno, TransJual.[No Faktur], TransJual.Kepada, Konsumen.Nama, TransJual.Tgl, TransJual.JenisTrans, TransJual.Diskon, TransJual.JatuhTempo, TransJual.Jenis, TransJual.Total " & _
"ORDER BY TransJual.Kepada, TransJual.[No Faktur];"
End If
Set k = a.Execute(CmadString)
Set DataLapHP_Jatuh = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function

Public Function DataLapStock_PersBD() As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String
k.CursorLocation = adUseClient
k.CursorType = adOpenStatic
k.LockType = adLockReadOnly
CmadString = "PersBDFinal"
Set k = a.Execute(CmadString)
Set DataLapStock_PersBD = k
Exit Function
aa:
'a.Execute "update transbeli set hapus=1 where [No Faktur]='Saldo'"
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function



Public Function TutupBuku(Optional SdTgl As Date) As String
Dim kkkO
With FrmTutupBuku
On Error Resume Next
a.Execute "insert into Supplier(Kode) values('-')"
On Error GoTo aa
Dim aString As String, intNoSaldo As Single, NamaFile As String, DiMana As String
Dim i As Integer, Direktori As String
On Error GoTo aa
DiMana = RegC.Dekrip(RegA.GetSettingString(HKEY_LOCAL_MACHINE, "Software\DataX Active Object\StockJual", "DataLoc", "2"))
  If Right(DiMana, 1) = "\" Then
  DiMana = Left(DiMana, Len(DiMana) - 1)
  End If

Dim koKj() As String
koKj = Split(DiMana, "\")
Direktori = ""
For i = LBound(koKj) To UBound(koKj) - 1
  Direktori = Direktori & "\" & koKj(i)
Next i
  If Right(Direktori, 1) = "\" Then
  Direktori = Left(Direktori, Len(Direktori) - 1)
  End If
  If Left(Direktori, 1) = "\" Then
  Direktori = Right(Direktori, Len(Direktori) - 1)
  End If

NamaFile = "TB" & Format(Date, "ddmmyyyy") '& Format(Time, "hhmmss") & ".mdb"
CopyFile DiMana, Direktori & "\" & NamaFile
SetAttr Direktori & "\" & NamaFile, vbHidden + vbReadOnly + vbSystem

Dim b As New ADODB.Recordset, rsSaldo As New ADODB.Recordset
a.BeginTrans
.Label3.Caption = "Persiapan Analisa Data.."
.P1.Value = 5
DoEvents
'#############################################################################
'Ambil intNo SaldoAwal di Trans Beli
CS:
 Set b = a.Execute("Select intno from transBeli where [No Faktur]='SAwal'")
 If b.EOF Then
  aString = "INSERT INTO TransBeli( [No Faktur],Kepada,Salesman," & _
  "Jenis,JatuhTempo,JenisTrans,Total) " & _
  "values('SAwal','-','-','C',0,'B',0)"
  a.Execute aString
  GoTo CS
 Else
  intNoSaldo = b!intNo
 End If


'#############################################################################
'hapus transaksi pembelian, penjualan yang dibatalkan
 aString = "delete from transbeli where hapus <> 0 and TransBeli.tgl <= #" & AmanTgl(.D1.Value) & "#" 'and jurnal=false
 a.Execute aString
 aString = "delete from transjual where hapus <> 0 and TransJual.tgl <= #" & AmanTgl(.D1.Value) & "#" 'and jurnal=false
 a.Execute aString, kkkO



.Label3.Caption = "Proses Transaksi Tunai.."
.P1.Value = 20
DoEvents
'#############################################################################
'Proses Semua Transaksi yg Cash kemudian hapus

'Pembelian & Retur Pembelian--------------------------------------------------
 If b.State = adStateOpen Then b.Close
 b.Open "SELECT TransBeli.intno, TransBeli.[No Faktur], TransBeli.JenisTrans, " & _
 "TransBeli.Tgl, TransBeli.Jenis, DetailBeli.[Kode Barang], iif(JenisTrans='B',Qty,-Qty) as Qty, " & _
 "(DetailBeli.Harga-DetailBeli.Diskon) as Harga FROM TransBeli " & _
 "INNER JOIN DetailBeli ON TransBeli.intno = DetailBeli.intTrans WHERE TransBeli.Jenis='C' " & _
 "and [No Faktur]<>'SAwal' and TransBeli.tgl<= #" & AmanTgl(.D1.Value) & "# ", a, adOpenStatic, adLockPessimistic

 Do While Not b.EOF
 Set rsSaldo = a.Execute("Select [Kode Barang] from DetailBeli where intTrans=" & intNoSaldo & " and [kode barang]='" & b![Kode Barang] & "'")

    If rsSaldo.EOF Then
      a.Execute "Insert into DetailBeli(IntTrans, [Kode Barang], Qty, Harga) values(" & _
      intNoSaldo & ",'" & b![Kode Barang] & "','" & b![qty] & "','" & b![Harga] & "')"
    Else
      a.Execute "Update Detailbeli set Qty = Qty + '" & b!qty & _
      "',Harga = ((Harga*Qty) + ('" & b!Harga * b!qty & "')) / (Qty+'" & b!qty & "') where intTrans=" & intNoSaldo & _
      " and [kode barang]='" & b![Kode Barang] & "'"
    End If
  b.MoveNext
 Loop
 a.Execute "Delete from TransBeli where TransBeli.Jenis='C' " & _
 "and [No Faktur]<>'SAwal' and TransBeli.tgl<= #" & AmanTgl(.D1.Value) & "# ", kkkO

'Penjualan & Retur Penjualan--------------------------------------------------
 If b.State = adStateOpen Then b.Close
 b.Open "SELECT TransJual.intno, TransJual.[No Faktur], TransJual.JenisTrans, " & _
 "TransJual.Tgl, TransJual.Jenis, DetailJual.[Kode Barang], iif(JenisTrans='RJ',Qty,-Qty) as Qty, " & _
 "(DetailJual.Harga-DetailJual.Diskon) as Harga FROM TransJual " & _
 "INNER JOIN DetailJual ON TransJual.intno = DetailJual.intTrans WHERE TransJual.Jenis='C' " & _
 "and [No Faktur]<>'SAwal' and TransJual.tgl<= #" & AmanTgl(.D1.Value) & "# ", a, adOpenStatic, adLockPessimistic

 Do While Not b.EOF
  Set rsSaldo = a.Execute("Select [Kode Barang] from DetailBeli where intTrans=" & intNoSaldo & " and [kode barang]='" & b![Kode Barang] & "'")
    If rsSaldo.EOF Then
      a.Execute "Insert into DetailBeli(IntTrans, [Kode Barang], Qty, Harga) values(" & _
      intNoSaldo & ",'" & b![Kode Barang] & "','" & b![qty] & "','" & b![Harga] & "')"
    Else
      a.Execute "Update Detailbeli set Qty = Qty + '" & b!qty & _
      "' where intTrans=" & intNoSaldo & _
      " and [kode barang]='" & b![Kode Barang] & "'"
    End If
  b.MoveNext
 Loop
 a.Execute "Delete from TransJual where TransJual.Jenis='C' " & _
 "and [No Faktur]<>'SAwal' and TransJual.tgl<= #" & AmanTgl(.D1.Value) & "# ", kkkO


.Label3.Caption = "Proses Transaksi Kredit.."
.P1.Value = 40
DoEvents
'#############################################################################
'Proses Semua Transaksi yg Kredit & Sudah OK kemudian hapus
'Pembelian & Retur Pembelian--------------------------------------------------
 If b.State = adStateOpen Then b.Close
 b.Open "SELECT TransBeli.intno, TransBeli.[No Faktur], TransBeli.JenisTrans, " & _
 "TransBeli.Tgl, TransBeli.Jenis, DetailBeli.[Kode Barang], iif(JenisTrans='B',Qty,-Qty) as Qty, " & _
 "(DetailBeli.Harga-DetailBeli.Diskon) as Harga FROM TransBeli " & _
 "INNER JOIN DetailBeli ON TransBeli.intno = DetailBeli.intTrans WHERE TransBeli.Jenis='OK' " & _
 "and [No Faktur]<>'SAwal' and TransBeli.tgl<= #" & AmanTgl(.D1.Value) & "# ", a, adOpenStatic, adLockPessimistic
 Debug.Print "*****Beli Credit********"
 Do While Not b.EOF
  Set rsSaldo = a.Execute("Select [Kode Barang] from DetailBeli where intTrans=" & intNoSaldo & " and [kode barang]='" & b![Kode Barang] & "'")
  Debug.Print b![Kode Barang]
    If rsSaldo.EOF Then
        a.Execute "Insert into DetailBeli(IntTrans, [Kode Barang], Qty, Harga) values(" & _
      intNoSaldo & ",'" & b![Kode Barang] & "','" & b![qty] & "','" & b![Harga] & "')"
    Else
           
      Dim rsOO1 As New ADODB.Recordset
      Set rsOO1 = a.Execute("Select * from DetailBeli where intTrans=" & intNoSaldo & " and [kode barang]='" & b![Kode Barang] & "'")
      If rsOO1!qty + b!qty = 0 Then
        a.Execute "Update Detailbeli set Qty=0,Harga=0 where intTrans=" & intNoSaldo & _
        " and [kode barang]='" & b![Kode Barang] & "'"
      Else
        a.Execute "Update Detailbeli set Qty = Qty + '" & b!qty & _
        "',Harga = ((Harga*Qty) + ('" & b!Harga * b!qty & "')) / (Qty+'" & b!qty & "') where intTrans=" & intNoSaldo & _
        " and [kode barang]='" & b![Kode Barang] & "'"
      End If
      
    End If
      Dim rsUrusUt As New ADODB.Recordset
      If rsUrusUt.State = adStateOpen Then rsUrusUt.Close
      rsUrusUt.Open "Select * from DetailUtang where KodeFaktur=" & b!intNo, a, adOpenStatic, adLockReadOnly
      Do While Not rsUrusUt.EOF
      a.Execute "update UtangAnak set Keterangan=Keterangan & chr(13) & '[" & b![No Faktur] & _
      " =" & rsUrusUt!Total & "+" & rsUrusUt!Potongan & "(Pot)]XXX' " & _
      "where intno=" & rsUrusUt!KodeBayar
      rsUrusUt.MoveNext
      Loop
  b.MoveNext
 Loop
 a.Execute "Delete from TransBeli where TransBeli.Jenis='OK' " & _
 "and [No Faktur]<>'SAwal' and TransBeli.tgl<= #" & AmanTgl(.D1.Value) & "# ", kkkO

'Penjualan & Retur Penjualan--------------------------------------------------
 If b.State = adStateOpen Then b.Close
 b.Open "SELECT TransJual.intno, TransJual.[No Faktur], TransJual.JenisTrans, " & _
 "TransJual.Tgl, TransJual.Jenis, DetailJual.[Kode Barang], iif(JenisTrans='RJ',Qty,-Qty) as Qty, " & _
 "(DetailJual.Harga-DetailJual.Diskon) as Harga FROM TransJual " & _
 "INNER JOIN DetailJual ON TransJual.intno = DetailJual.intTrans WHERE TransJual.Jenis='OK' " & _
 "and [No Faktur]<>'SAwal' and TransJual.tgl<= #" & AmanTgl(.D1.Value) & "# ", a, adOpenStatic, adLockPessimistic
 Debug.Print "*****Jual Credit********"
 Do While Not b.EOF
  Set rsSaldo = a.Execute("Select [Kode Barang] from DetailBeli where intTrans=" & intNoSaldo & " and [kode barang]='" & AmanOi(b![Kode Barang]) & "'")
  Debug.Print b![Kode Barang]
    If rsSaldo.EOF Then
    a.Execute "Insert into DetailBeli(IntTrans, [Kode Barang], Qty, Harga) values(" & _
      intNoSaldo & ",'" & AmanOi(b![Kode Barang]) & "','" & b![qty] & "','" & b![Harga] & "')"
    Else
      a.Execute "Update Detailbeli set Qty = Qty + '" & b!qty & _
      "' where intTrans=" & intNoSaldo & _
      " and [kode barang]='" & AmanOi(b![Kode Barang]) & "'"
    End If
      Dim rsUrusPi As New ADODB.Recordset
      If rsUrusPi.State = adStateOpen Then rsUrusPi.Close
      rsUrusPi.Open "Select * from DetailPiutang where KodeFaktur=" & b!intNo, a, adOpenStatic, adLockReadOnly
      Do While Not rsUrusPi.EOF
      a.Execute "update PiutangAnak set Keterangan=Keterangan & chr(13) & '[" & b![No Faktur] & _
      " =" & rsUrusPi!Total & "+" & rsUrusPi!Potongan & "(Pot)]XXX' " & _
      "where intno=" & rsUrusPi!KodeBayar
      rsUrusPi.MoveNext
      Loop
  b.MoveNext
 Loop
 a.Execute "Delete from TransJual where TransJual.Jenis='OK' " & _
 "and [No Faktur]<>'SAwal' and TransJual.tgl<= #" & AmanTgl(.D1.Value) & "# ", kkkO

'hapus barang di saldo awal yg qty=0
a.Execute "Delete from DetailBeli where qty=0 and inttrans=" & intNoSaldo, kkkO


.Label3.Caption = "Proses Pembayaran Hutang Piutang.."
.P1.Value = 70
DoEvents
'#############################################################################
'hapus transaksi Utang & Piutang
Dim rsA As New ADODB.Recordset, rsB As ADODB.Recordset
'Utang
Set rsA = a.Execute("Select intNo from UtangAnak")
Do While Not rsA.EOF
  Set rsB = a.Execute("select KodeBayar from DetailUtang where KodeBayar=" & rsA!intNo)
  If rsB.EOF Then a.Execute "Delete from UtangAnak where intno=" & rsA!intNo
rsA.MoveNext
Loop
Set rsA = a.Execute("Select intNo from Utang")
Do While Not rsA.EOF
  Set rsB = a.Execute("select KodeBayar from UtangAnak where KodeBayar=" & rsA!intNo)
  If rsB.EOF Then a.Execute "Delete from Utang where intno=" & rsA!intNo
rsA.MoveNext
Loop
'Piutang
Set rsA = a.Execute("Select intNo from PiutangAnak")
Do While Not rsA.EOF
  Set rsB = a.Execute("select KodeBayar from DetailPiutang where KodeBayar=" & rsA!intNo)
  If rsB.EOF Then a.Execute "Delete from PiutangAnak where intno=" & rsA!intNo
rsA.MoveNext
Loop
Set rsA = a.Execute("Select intNo from Piutang")
Do While Not rsA.EOF
  Set rsB = a.Execute("select KodeBayar from PiutangAnak where KodeBayar=" & rsA!intNo)
  If rsB.EOF Then a.Execute "Delete from Piutang where intno=" & rsA!intNo
rsA.MoveNext
Loop

.Label3.Caption = "Proses Surat Jalan.."
.P1.Value = 90
DoEvents

'#############################################################################
'Ubah tgl Saldo awal jadi tgl akhir tutup buku
Set b = a.Execute("select min(Tgl) as A from TransBeli group by intNo")
 a.Execute "Update transBeli set tgl='" & AmanTgl(b!a) & "' where [No Faktur]='SAwal'"

'#############################################################################
'SuratJalan
a.Execute "Delete from SuratJalan where Tgl <= #" & AmanTgl(.D1.Value) & "# "
.P1.Value = 100
a.CommitTrans
TutupBuku = ""
End With
Exit Function
aa:
a.RollbackTrans
TutupBuku = "Telah terjadi kesalahan.." & vbCrLf & "No : " & Err.Number & _
vbCrLf & "Keterangan : " & Err.Description
End Function


Public Function SimpanRetur(tTabel As String, vMaster As VSFlexGrid) As String
On Error GoTo aa
a.BeginTrans
If tTabel = "RB" Then
tTabel = "DetailBeli"
Else: tTabel = "DetailJual"
End If
Dim i As Single
i = 1
'0=Kodebrg 1=Namabrg 2=Qty   3=QtyMinus   4=0  5=3-4-5  6=intTrans 7=harga 8=diskon 9=HRata
With vMaster
Do While i <= (.Rows - 1)
  If Val(.TextMatrix(i, 4)) <> 0 Then
    a.Execute "delete from " & tTabel & " where [kode Barang]='" & _
    AmanOi(.TextMatrix(i, 0)) & "' and intTrans=" & .TextMatrix(i, 6) & _
    " and qty=-'" & Val(.TextMatrix(i, 3)) & "'"
  If tTabel = "DetailJual" Then
    a.Execute "insert into " & tTabel & " (intTrans,[Kode Barang],Qty,Harga,Diskon,HRata) values(" & _
    .TextMatrix(i, 6) & ",'" & AmanOi(.TextMatrix(i, 0)) & "','" & _
    -(Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4))) & _
    "','" & .TextMatrix(i, 7) & "','" & .TextMatrix(i, 8) & "','" & .TextMatrix(i, 9) & "')"
  Else
    a.Execute "insert into " & tTabel & " (intTrans,[Kode Barang],Qty,Harga,Diskon) values(" & _
    .TextMatrix(i, 6) & ",'" & AmanOi(.TextMatrix(i, 0)) & "','" & _
    -(Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4))) & _
    "','" & .TextMatrix(i, 7) & "','" & .TextMatrix(i, 8) & "')"
  End If
    a.Execute "update barang set qty=qty" & _
    IIf(tTabel = "DetailBeli", "+'", "-'") & .TextMatrix(i, 4) & _
    "' where [kode Barang]='" & AmanOi(.TextMatrix(i, 0)) & "'"
  End If
i = i + 1
Loop
Dim rsSmpn As New ADODB.Recordset
If .TextMatrix(.Rows - 1, 6) <> "" Then
Set rsSmpn = a.Execute("select sum(qty*(harga-diskon)) from " & tTabel & " where inttrans=" & .TextMatrix(.Rows - 1, 6))
  
If tTabel = "DetailBeli" Then
tTabel = "TransBeli"
Else: tTabel = "TransJual"
End If
    a.Execute "update " & tTabel & " set total='" & rsSmpn(0) & "' where intno=" & .TextMatrix(.Rows - 1, 6)
End If
End With
a.CommitTrans
SimpanRetur = ""
Exit Function
aa:
a.RollbackTrans
SimpanRetur = "Telah terjadi kesalahan.." & vbCrLf & "No : " & Err.Number & _
vbCrLf & "Keterangan : " & Err.Description
End Function



Public Function DataLapLRItem() As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String, rsOk As New ADODB.Recordset
CmadString = "SELECT TransJual.intno, TransJual.Tgl, TransJual.JenisTrans, TransJual.Jenis, " & _
"DetailJual.[Kode Barang], Barang.[Nama Barang], " & _
"iif(JenisTrans='J',DetailJual.Qty,-DetailJual.Qty) as Qty, " & _
"(1-(transjual.diskon/iif(transjual.total=0,1,transjual.total)))*(DetailJual.Harga-DetailJual.Diskon) AS Harga,HRata " & _
"FROM Barang INNER JOIN (TransJual INNER JOIN DetailJual ON TransJual.intno=DetailJual.intTrans) " & _
"ON Barang.[Kode Barang]=DetailJual.[Kode Barang] " & _
"WHERE transjual.hapus=0;"
Set k = a.Execute(CmadString)
Set DataLapLRItem = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function


'no faktur 'Saldo' hutangnya harus ada di utang &detail utang kalo nga tidak balance di pembelian


Public Function DataLapLRStock(tAwal As Date, tAkhir As Date) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String, rsOk As New ADODB.Recordset
Dim NamaFile As String
tAkhir = tAkhir
rsOk.CursorLocation = adUseClient
rsOk.CursorType = adOpenStatic
rsOk.LockType = adLockReadOnly
k.Fields.Append "A", adVarChar, 50
k.Fields.Append "B", adVarChar, 50
k.Fields.Append "C", adVarChar, 50
k.Fields.Append "D", adCurrency
k.Fields.Append "E", adBoolean
k.Fields.Append "Dari", adDate
k.Fields.Append "Sampai", adDate
k.Open

'Awal
CmadString = "LRFinal '" & tAwal & "'"
Set rsOk = a.Execute(CmadString)
If Not rsOk.EOF Then
k.AddNew
k!a = "Harga Pokok Penjualan": k!b = "Persediaan Awal"
k!C = "aw": k!D = IIf(IsNull(rsOk!Total), 0, rsOk!Total): k!E = 0
k!Dari = tAwal: k!sampai = tAkhir
k.Update
End If

'Beli
CmadString = "LRTrans '" & tAwal & "','" & tAkhir & "'"
Set rsOk = a.Execute(CmadString)
If Not rsOk.EOF Then
  rsOk.MoveFirst
  rsOk.Find "JenisTrans='B'"
  If Not rsOk.EOF Then
    k.AddNew
    k!a = "Harga Pokok Penjualan": k!b = "Pembelian"
    k!C = "Pembelian": k!D = rsOk!Total: k!E = 1
    k!Dari = tAwal: k!sampai = tAkhir
    k.Update
  End If
  rsOk.MoveFirst
  rsOk.Find "JenisTrans='RB'"
  If Not rsOk.EOF Then
    k.AddNew
    k!a = "Harga Pokok Penjualan": k!b = "Pembelian"
    k!C = "Retur Pembelian": k!D = -rsOk!Total: k!E = 1
    k!Dari = tAwal: k!sampai = tAkhir
    k.Update
  End If
End If

'Akhir
CmadString = "LRFinal '" & tAkhir + 1 & "'"
Set rsOk = a.Execute(CmadString)
If Not rsOk.EOF Then
k.AddNew
k!a = "Harga Pokok Penjualan": k!b = "Persediaan Akhir"
k!C = "ak": k!D = IIf(IsNull(rsOk!Total), 0, -rsOk!Total): k!E = 0
k!Dari = tAwal: k!sampai = tAkhir
k.Update
End If

'Jual
CmadString = "LRTrans '" & tAwal & "','" & tAkhir & "'"
Set rsOk = a.Execute(CmadString)
If Not rsOk.EOF Then
  rsOk.MoveFirst
  rsOk.Find "JenisTrans='J'"
  If Not rsOk.EOF Then
    k.AddNew
    k!a = "Penjualan": k!b = "Penjualan"
    k!C = "Penjualan": k!D = -rsOk!Total: k!E = 1
    k!Dari = tAwal: k!sampai = tAkhir
    k.Update
  End If
  rsOk.MoveFirst
  rsOk.Find "JenisTrans='RJ'"
  If Not rsOk.EOF Then
    k.AddNew
    k!a = "Penjualan": k!b = "Penjualan"
    k!C = "Retur Penjualan": k!D = rsOk!Total: k!E = 1
    k!Dari = tAwal: k!sampai = tAkhir
    k.Update
  End If
End If

Set DataLapLRStock = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function



Public Function SimpanSPB(Head() As String, Bodi() As String) As String
On Error GoTo aa
Dim aString As String
a.BeginTrans

'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = LamaKredit    (5) = jenistrans
'(6) =Keterangan
aString = "INSERT INTO SuratJalan( [No Faktur],Tgl,Kepada,Salesman," & _
"JatuhTempo,JenisTrans,Keterangan) " & _
"values('" & Head(0) & "','" & Head(1) & _
"','" & Head(2) & "','" & Head(3) & _
"'," & Head(4) & ",'" & Head(5) & _
"','" & Head(6) & "')"
a.Execute aString

Dim i As Byte, b As New ADODB.Recordset
Set b = a.Execute("select intno from SuratJalan where [No Faktur]='" & Head(0) & "'")

For i = 1 To UBound(Bodi)
If Bodi(i, 0) <> "" Then
'(0) = Kode  (1) = Qty
aString = "INSERT INTO SuratJalanDetail (intTrans, [Kode Barang],qty) " & _
"values(" & b(0) & ",'" & _
Bodi(i, 0) & "'," & _
Bodi(i, 1) & ")"
a.Execute aString
End If
Next i
a.CommitTrans
SimpanSPB = ""
Exit Function
aa:
a.RollbackTrans
SimpanSPB = Err.Number & Err.Description
End Function








Public Function DataLapSalesKomisi(tAwal As Date, tAkhir As Date, Jenis As String) As ADODB.Recordset
On Error GoTo aa
Dim k As New ADODB.Recordset, CmadString As String

If Jenis = "2" Then
CmadString = "SELECT Salesman.Kode as aKode, " & _
"Sum(iif(transjual.jenistrans='J',JualSales.PotBayar,-JualSales.PotBayar)) AS PotBayar, " & _
"Sum(iif(transjual.jenistrans='J',TransJual.Diskon,-TransJual.Diskon)) AS Diskon " & _
"FROM (Salesman INNER JOIN TransJual ON Salesman.Kode = TransJual.Salesman) INNER JOIN JualSales2 ON TransJual.intno = JualSales2.intNoFaktur " & _
"WHERE (((TransJual.Jenis)<>'K') AND ((TransJual.Hapus)=0) AND " & _
"((TransJual.Tgl) Between #" & AmanTgl(tAwal) & "# And #" & AmanTgl(tAkhir) & "#)) " & _
"GROUP BY Salesman.Kode;"
Else
CmadString = "SELECT Salesman.Kode as Kode, Salesman.Nama, DetailJual.[Kode Barang], Barang.[Nama Barang], Sum(IIf(TransJual.JenisTrans='J',DetailJual.Qty,-DetailJual.Qty)) AS Qty, Avg([DetailJual.Harga]-[DetailJual.Diskon]-[DetailJual.PH]) AS Harga " & _
"FROM ((Salesman INNER JOIN TransJual ON Salesman.Kode = TransJual.Salesman) INNER JOIN (Barang INNER JOIN DetailJual ON Barang.[Kode Barang] = DetailJual.[Kode Barang]) ON TransJual.intno = DetailJual.intTrans) INNER JOIN JualSales2 ON TransJual.intno = JualSales2.intNoFaktur " & _
"WHERE (((TransJual.Jenis)<>'K') AND ((TransJual.Hapus)=0) AND " & _
"((TransJual.Tgl) Between #" & AmanTgl(tAwal) & "# And #" & AmanTgl(tAkhir) & "#) ) " & _
"GROUP BY Salesman.Kode, Salesman.Nama, DetailJual.[Kode Barang], Barang.[Nama Barang];"

End If
Set k = a.Execute(CmadString)
Set DataLapSalesKomisi = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function




Public Function SimpanKurs(Head As VSFlexGrid, BTgl As String) As String
On Error GoTo aa
Dim aString As String, i As Single
a.BeginTrans

a.Execute "Delete from matakurs where tgl=#" & Format(BTgl, "mm/dd/yyyy") & "#"

'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
For i = 1 To Head.Rows - 1
aString = "INSERT INTO MataKurs(Tgl,KodeMata,[Nilai Kurs]) " & _
"values('" & BTgl & "','" & Head.TextMatrix(i, 2) & _
"','" & Head.TextMatrix(i, 1) & "')"
a.Execute aString
Next i

a.CommitTrans
SimpanKurs = ""
Exit Function
aa:
a.RollbackTrans
SimpanKurs = Err.Number & Err.Description
End Function


Public Function QtyModal(KdBrg As String) As ADODB.Recordset
On Error GoTo aa
Dim rsQty As New ADODB.Recordset, Jlh As Currency
Dim k As New ADODB.Recordset
Set rsQty = a.Execute("Select qty from barang where [Kode Barang]='" & KdBrg & "'")
Jlh = rsQty!qty
  If Jlh > 0 Then
    k.Fields.Append "Kode", adVarChar, 50
    k.Fields.Append "Tanggal", adDate, 50
    k.Fields.Append "No Faktur", adVarChar, 50
    k.Fields.Append "Qty", adCurrency
    k.Fields.Append "Harga", adCurrency
    k.Fields.Append "Jenis", adVarChar, 50
    k.Open
      Set rsQty = a.Execute("SELECT DetailBeli.intNo, DetailBeli.[Kode Barang], DetailBeli.Qty, " & _
      "(1-(transbeli.diskon/iif(transbeli.total=0,1,transbeli.total)))*([Harga]-[DetailBeli.Diskon]) AS Total, 'Beli' AS Jenis, TransBeli.Tgl, " & _
      "TransBeli.intno,[No Faktur] FROM TransBeli INNER JOIN DetailBeli ON " & _
      "TransBeli.intno = DetailBeli.intTrans WHERE (TransBeli.JenisTrans)='B' " & _
      "and [Kode Barang]='" & KdBrg & "' and hapus=0 ORDER BY Jenis, Tgl desc, " & _
      "[No Faktur] desc Union " & _
      "SELECT DetailJual.intNo, DetailJual.[Kode Barang], DetailJual.Qty, [HRata] AS Total, " & _
      "'Retur Penjualan' AS Jenis, TransJual.Tgl, TransJual.intno,[No Faktur] " & _
      "FROM TransJual INNER JOIN DetailJual ON TransJual.intno = DetailJual.intTrans " & _
      "WHERE (TransJual.JenisTrans)='RJ' and [Kode Barang]='" & KdBrg & "' and hapus=0 " & _
      "ORDER BY Jenis, Tgl desc, [No Faktur] desc")
'      Set rsQty = a.Execute("FIFO2 '" & KdBrg & "'")
'      SELECT DetailBeli.[Kode Barang], DetailBeli.Qty, " & _
'      "[Harga]-[DetailBeli.Diskon] AS Total, 'Beli' AS Jenis, TransBeli.Tgl, TransBeli.intno,[No Faktur] " & _
'      "FROM TransBeli INNER JOIN DetailBeli ON TransBeli.intno = DetailBeli.intTrans " & _
'      "WHERE (TransBeli.JenisTrans)='B' and [Kode Barang]='" & KdBrg & _
'      "' ORDER BY TransBeli.Tgl desc, TransBeli.intno desc;")
      Do While Jlh > 0
        k.AddNew: k!Kode = KdBrg: k!qty = IIf(Jlh > rsQty!qty, rsQty!qty, Jlh)
        k!Jenis = rsQty!Jenis
        k!Tanggal = Format(rsQty!Tgl, "dd mmm yyyy"): k![No Faktur] = rsQty![No Faktur]
        k!Harga = rsQty!Total: k.Update
        Jlh = Jlh - rsQty!qty
        rsQty.MoveNext
      Loop
      k.Sort = "Jenis Asc, [Tanggal] ASC, [No Faktur] ASC"
  Else
    k.Fields.Append "Kode", adVarChar, 500
    k.Open
    k.AddNew
    k!Kode = "Barang dengan Kode " & KdBrg & "memiliki Qty 0 atau tidak mencukupi"
    k.Update
  End If
    Set QtyModal = k
Exit Function
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur QtyModal pada Class Module Data"
End Function

Public Function QtyRata(KdBrg As String, ByVal qty As Currency) As Currency
On Error GoTo aa
KdBrg = AmanOi(KdBrg)
Dim rsOOK As New ADODB.Recordset
Dim xJlh As Currency, xTotal As Currency
  Set rsOOK = QtyModal(KdBrg)
  If rsOOK.Fields(0).DefinedSize = 500 Then GoTo aa
  rsOOK.Sort = "[Tanggal] ASC, [No Faktur] ASC"
  xJlh = 0: xTotal = 0
  Do While qty > 0
    xJlh = xJlh + IIf(qty > rsOOK!qty, rsOOK!qty, qty)
    xTotal = xTotal + (IIf(qty > rsOOK!qty, rsOOK!qty, qty) * rsOOK!Harga)
    qty = qty - rsOOK!qty
    rsOOK.MoveNext
  Loop
QtyRata = xTotal / xJlh
Exit Function
aa:
Set rsOOK = a.Execute("select * from barang where [Kode Barang]='" & KdBrg & "'")
xTotal = rsOOK![Harga Beli] - (rsOOK![Disc Beli] / 100 * rsOOK![Harga Beli])
If Not LagiSimpan Then
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur QtyRata pada Class Module Data" & _
" Kode Barang:" & KdBrg & ", Qty:" & qty & ", Modal:" & xTotal
Else
SimpanSalah = SimpanSalah & vbCrLf & " Kode Barang:" & KdBrg & ", Qty:" & qty & ", Modal:" & xTotal
End If
QtyRata = xTotal
End Function

Public Sub HitungUlang(aTgl As String)
5     On Error GoTo aa
      Dim rsJual As New ADODB.Recordset, rsJualDet As New ADODB.Recordset
      Dim NamaFile As String, DiMana As String, i As Integer, Direktori As String
10    On Error GoTo aa
15    DiMana = RegC.Dekrip(RegA.GetSettingString(HKEY_LOCAL_MACHINE, "Software\DataX Active Object\StockJual", "DataLoc", "2"))
20      If Right(DiMana, 1) = "\" Then
25      DiMana = Left(DiMana, Len(DiMana) - 1)
30      End If

      Dim koKj() As String
35    koKj = Split(DiMana, "\")
40    Direktori = ""
45    For i = LBound(koKj) To UBound(koKj) - 1
50      Direktori = Direktori & "\" & koKj(i)
55    Next i
60      If Right(Direktori, 1) = "\" Then
65      Direktori = Left(Direktori, Len(Direktori) - 1)
70      End If
75      If Left(Direktori, 1) = "\" Then
80      Direktori = Right(Direktori, Len(Direktori) - 1)
85      End If

90    NamaFile = "HU" & Format(Date, "ddmmyyyy") '& Format(Time, "hhmmss") & ".mdb"
95    CopyFile DiMana, Direktori & "\" & NamaFile
100   SetAttr Direktori & "\" & NamaFile, vbHidden + vbReadOnly + vbSystem

105   a.Execute "update TransJual set UdaHit=1 where jenistrans='RJ'"
110   a.Execute "update TransJual set UdaHit=1 where Tgl<cdate('" & aTgl & "') and jenistrans='J'"
115   a.Execute "update TransJual set UdaHit=0 where Tgl>=cdate('" & aTgl & "') and jenistrans='J'"
120   a.BeginTrans
125   Set rsJual = a.Execute("select TransJual.intNo, TransJual.[No Faktur], TransJual.JenisTrans, TransJual.Tgl from TransJual where Tgl>=cdate('" & aTgl & "') and hapus=0 and jenistrans='J' order by Tgl,[No Faktur]")
130   FrmHitungUlang.P1.Max = IIf(rsJual.RecordCount = 0, 100, rsJual.RecordCount)
135   FrmHitungUlang.P1.Value = 1
140   Do While Not rsJual.EOF
145     Set rsJualDet = a.Execute("SELECT DetailJual.[Kode Barang], Sum(DetailJual.Qty) AS SumOfQty From DetailJual where intTrans=" & rsJual!intNo & " GROUP BY DetailJual.[Kode Barang]")
150     Do While Not rsJualDet.EOF
        'Debug.Assert (rsJual![No faktur] <> "B000644")
155     a.Execute "update DetailJual set Hrata='" & HitQtyRata(rsJualDet![Kode Barang], rsJualDet!SumOfQty, rsJual!Tgl) & "' where [Kode Barang]='" & rsJualDet![Kode Barang] & _
        "' and intTrans=" & rsJual!intNo
160     Debug.Assert (rsJualDet![Kode Barang] <> "AR-TDI-S1/2")
165     rsJualDet.MoveNext
170     FrmHitungUlang.P1.Value = rsJual.AbsolutePosition
175     DoEvents
180     Loop
185   a.Execute "update transjual set udahit=1 where intNo=" & rsJual!intNo
190   rsJual.MoveNext
195   Loop
200   a.CommitTrans
205   MsgBox "Proses telah selesai.." & _
      vbCrLf & "Data laba-rugi transaksi anda telah dihitung ulang..", vbInformation, "Tutup Buku"
210   Exit Sub
aa:
      Dim Err_Setering As String
215   a.RollbackTrans
220   Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur HitungUlang pada " & "Class Module Data di baris " & Erl
225   MsgBox Err_Setering, vbRetryCancel, App.Title & "-Data Error"
Exit_HitungUlang:
230   App.LogEvent "myAS=>" & Format(Date, "dd-mmmm-yyyy") & Format(Time, "(hh:mm:ss)") & _
      vbCrLf & Err_Setering & vbCrLf, vbLogEventTypeError
235   Exit Sub
End Sub


Public Function HitQtyModal(KdBrg As String, aTgl As String) As ADODB.Recordset
Dim rsQty As New ADODB.Recordset, Jlh As Currency
Dim k As New ADODB.Recordset, qOO As New ADODB.Recordset
On Error GoTo aa
Set rsQty = a.Execute("SELECT Sum(iif(TransBeli.JenisTrans='B',DetailBeli.Qty,-DetailBeli.Qty)) AS Qty, " & _
"DetailBeli.[Kode Barang] FROM TransBeli INNER JOIN DetailBeli ON TransBeli.intno = DetailBeli.intTrans " & _
"where (((DetailBeli.[Kode Barang])='" & KdBrg & "') AND ((TransBeli.Hapus)=0) AND  " & _
"((TransBeli.Tgl)<=cdate('" & aTgl & "'))) " & _
"GROUP BY DetailBeli.[Kode Barang]")
If rsQty.EOF Then
Jlh = 0
Else
Debug.Assert (KdBrg <> "AR-TDI-S1/2")
Jlh = rsQty!qty
End If
Set rsQty = a.Execute("SELECT Sum(iif(TransJual.JenisTrans='RJ',DetailJual.Qty,-DetailJual.Qty)) AS Qty, " & _
"DetailJual.[Kode Barang] FROM TransJual INNER JOIN DetailJual ON TransJual.intno = DetailJual.intTrans " & _
"where (((DetailJual.[Kode Barang])='" & KdBrg & "') AND ((TransJual.Hapus)=0) AND  " & _
"((TransJual.Tgl)<=cdate('" & aTgl & "')) and udahit<>0) " & _
"GROUP BY DetailJual.[Kode Barang]")
If Not rsQty.EOF Then
Jlh = Jlh + rsQty!qty
End If
    k.Fields.Append "Kode", adVarChar, 50
    k.Fields.Append "Tanggal", adDate, 50
    k.Fields.Append "No Faktur", adVarChar, 50
    k.Fields.Append "Qty", adCurrency
    k.Fields.Append "Harga", adCurrency
    k.Fields.Append "Jenis", adVarChar, 50
    k.Open
      Set rsQty = a.Execute("SELECT DetailBeli.intNo,DetailBeli.[Kode Barang], DetailBeli.Qty, " & _
      "(1-(transbeli.diskon/iif(transbeli.total=0,1,transbeli.total)))*([Harga]-[DetailBeli.Diskon]) AS Total, 'Beli' AS Jenis, TransBeli.Tgl, " & _
      "TransBeli.intno,[No Faktur] FROM TransBeli INNER JOIN DetailBeli ON " & _
      "TransBeli.intno = DetailBeli.intTrans WHERE (TransBeli.JenisTrans)='B' " & _
      "and [Kode Barang]='" & KdBrg & "' and hapus=0 and TransBeli.Tgl<=cdate('" & aTgl & "') ORDER BY Jenis, Tgl desc, " & _
      "[No Faktur] desc Union " & _
      "SELECT DetailJual.intNo, DetailJual.[Kode Barang], DetailJual.Qty, [HRata] AS Total, " & _
      "'Retur Penjualan' AS Jenis, TransJual.Tgl, TransJual.intno,[No Faktur] " & _
      "FROM TransJual INNER JOIN DetailJual ON TransJual.intno = DetailJual.intTrans " & _
      "WHERE (TransJual.JenisTrans)='RJ' and [Kode Barang]='" & KdBrg & "' and hapus=0 and TransJual.Tgl<=cdate('" & aTgl & "') " & _
      "ORDER BY Jenis, Tgl desc, [No Faktur] desc")
      Do While Jlh > 0
        k.AddNew: k!Kode = KdBrg: k!qty = IIf(Jlh > rsQty!qty, rsQty!qty, Jlh)
        k!Jenis = rsQty!Jenis
        k!Tanggal = Format(rsQty!Tgl, "dd mmm yyyy"): k![No Faktur] = rsQty![No Faktur]
        k!Harga = rsQty!Total: k.Update
        Jlh = Jlh - rsQty!qty
        rsQty.MoveNext
      Loop
Set qOO = a.Execute("select * from barang where [Kode Barang]='" & KdBrg & "'")
Dim HrgBl As Currency
HrgBl = IIf(qOO![Harga Beli] = 0, 1, qOO![Harga Beli])
HrgBl = HrgBl - (qOO![Disc Beli] / 100 * HrgBl)
k.AddNew: k!Kode = KdBrg: k!qty = 1000000
k!Jenis = "ErrCor"
k!Tanggal = Format(Date, "dd mmm yyyy"): k![No Faktur] = "AA"
k!Harga = HrgBl: k.Update
k.Sort = "Jenis Asc, [Tanggal] ASC, [No Faktur] ASC"
    Set HitQtyModal = k
Exit Function
aa:
Dim Err_Setering As String
Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur HitQtyModal pada " & "Class Module Data di baris " & Erl
MsgBox Err_Setering, vbRetryCancel, App.Title & "-Data Error"
Exit_HitQtyModal:
App.LogEvent "myAS=>" & Format(Date, "dd-mmmm-yyyy") & Format(Time, "(hh:mm:ss)") & _
vbCrLf & Err_Setering & vbCrLf, vbLogEventTypeError
End Function

Public Function HitQtyRata(KdBrg As String, ByVal qty As Currency, aTgl As String) As Currency
On Error GoTo aa
KdBrg = AmanOi(KdBrg)
Dim rsOOK As New ADODB.Recordset
Dim xJlh As Currency, xTotal As Currency
If qty = 0 Then
HitQtyRata = 0
GoTo ab
End If
  Set rsOOK = HitQtyModal(KdBrg, aTgl)
  rsOOK.Sort = "[Tanggal] ASC, [No Faktur] ASC"
  xJlh = 0: xTotal = 0
  Do While qty > 0
    xJlh = xJlh + IIf(qty > rsOOK!qty, rsOOK!qty, qty)
    xTotal = xTotal + (IIf(qty > rsOOK!qty, rsOOK!qty, qty) * rsOOK!Harga)
    qty = qty - rsOOK!qty
    rsOOK.MoveNext
  Loop
HitQtyRata = xTotal / xJlh
Debug.Assert (KdBrg <> "AR-TDI-S1/2")
ab:
Exit Function
aa:
Dim Err_Setering As String
Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur HitQtyRata pada " & "Class Module Data di baris " & Erl
MsgBox Err_Setering, vbRetryCancel, App.Title & "-Data Error"
Exit_HitQtyRata:
App.LogEvent "myAS=>" & Format(Date, "dd-mmmm-yyyy") & Format(Time, "(hh:mm:ss)") & _
vbCrLf & Err_Setering & vbCrLf, vbLogEventTypeError
End Function


