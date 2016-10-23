Attribute VB_Name = "SerbaGuna"
Option Explicit
Public Pendaftaran As Boolean
Public BarisGrid()
Public fRpt As New Form2
Public aData As New Data
'Public Lap As New LaporanStockAtiong.File
Public Lap As New xLaporan
Public Lokasi As String
Public LokasiFile As String
'Public RptPalsu As New Form2
Public RegA As New Registri.RegistrySetting
Public RegB As New Registri.GetDiscID
Public RegC As New Registri.Enkripsi
Public RegD As New Registri.Rupiah
Public xPotHar()
Public HasilTT As String
Public nAMA As String

Declare Function apiCopyFile Lib "kernel32" Alias "CopyFileA" _
(ByVal lpExistingFileName As String, _
ByVal lpNewFileName As String, _
ByVal bFailIfExists As Long) As Long

Public LagiSimpan As Boolean, SimpanSalah As String
'xp style
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200


'#########Win Special Folder##########
' Declare Public variables.
Public Type ShortItemId
  cb As Long
  abID As Byte
End Type
Public Type ITEMIDLIST
  mkid As ShortItemId
End Type
' Declare constants.
Const CSIDL_TEMPLATES = &H15
' Declare API functions.
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
(ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib _
"shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder _
As Long, pidl As ITEMIDLIST) As Long

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function





Public Function NoFaktur(Awal As String, Tabel As String, JenisTrans As String) As String
On Error GoTo aa
Dim rsData As New ADODB.Recordset, KKO As String

Select Case JenisTrans
Case "B"
  Set rsData = aData.AmbilCommand("select max(val(right([no faktur],5))) from transbeli where len([no faktur])=7 and jenistrans='B'")
  If IsNull(rsData(0)) Then
 ' KKO = 1
  Else
'  KKO = rsData(0) + 1
  End If
  'NoFaktur = Chr(Year(Date) - 2007 + 65) & Chr(Month(Date) + 64) & Format(KKO, "00000")
  NoFaktur = ""
Case "J"
  Set rsData = aData.AmbilCommand("select max(val(right([no faktur],6))) from transJual where len([no faktur])=13 and jenistrans='J' and mid([no faktur],5,2)='" & Right(Year(Date), 2) & "' and posting=false and left([no faktur],2)='AB'")
  If IsNull(rsData(0)) Then
  KKO = 1
  Else
  KKO = Val(Right(rsData(0), 6)) + 1
  End If
  NoFaktur = "AB" & Format(Date, "mmyy") & "/" & Format(KKO, "000000")
Case "RB"
  Set rsData = aData.AmbilCommand("select max(val(mid([no faktur],3,3))) from transbeli where left([no faktur],2)='RB' and jenistrans='RB'")
  KKO = IIf(IsNull(rsData(0)), 0, rsData(0)) + 1
  NoFaktur = "RB" & Format(KKO, "000") & "/" & Romawi(Month(Date)) & "/" & Year(Date)
Case "RJ"
  Set rsData = aData.AmbilCommand("select max(val(mid([no faktur],3,3))) from transJual where left([no faktur],2)='RJ' and jenistrans='RJ'")
  KKO = IIf(IsNull(rsData(0)), 0, rsData(0)) + 1
  NoFaktur = "RJ" & Format(KKO, "000") & "/" & Romawi(Month(Date)) & "/" & Year(Date)
Case "Baa"
  Set rsData = aData.AmbilData("select max(val(mid([no faktur],3,4))) from [" & Tabel & "] " & _
  "where JenisTrans='' and left([no faktur],2)='" & Awal & "' and right([no faktur],6)='" & _
  Format(Date, "/mm/yy") & "'")
  KKO = IIf(IsNull(rsData(0)), 0, rsData(0)) + 1
  NoFaktur = Awal & Format(KKO, "0000") & Format(Date, "/mm/yy")
End Select
 
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function

Public Function Romawi(a As Integer) As String
  Select Case a
  Case 1: Romawi = "I"
  Case 2: Romawi = "II"
  Case 3: Romawi = "III"
  Case 4: Romawi = "IV"
  Case 5: Romawi = "V"
  Case 6: Romawi = "VI"
  Case 7: Romawi = "VII"
  Case 8: Romawi = "VIII"
  Case 9: Romawi = "IX"
  Case 10: Romawi = "X"
  Case 11: Romawi = "XI"
  Case 12: Romawi = "XII"
  End Select
End Function

Public Function AmanOi(Pengacau As String) As String
On Error GoTo aa
AmanOi = Replace(Pengacau, "'", "''")
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function

Public Function AmanTgl(TglKaco As Date) As String
On Error GoTo aa
'AmanTgl = Format(TglKaco, "mm/dd/yyyy")
AmanTgl = Format(TglKaco, "dd/mm/yyyy")
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function


Sub Main()
Dim k As String, l As String, Coba As Boolean
   InitCommonControlsVB
Dim aWinPath As String

aWinPath = Left(GetSpecialFolder(CSIDL_TEMPLATES), Len(GetSpecialFolder(CSIDL_TEMPLATES)) - 9)


'-------------------------
Dim isiFile As String
Dim iFNo As Integer
'-----------------------------------------------------------------------
Lokasi = RegC.Dekrip(RegA.GetSettingString(HKEY_LOCAL_MACHINE, "Software\DataX Active Object\StockJual", "DataLoc", "2"))
If Lokasi = "" Then
Lokasi = App.Path
End If
frmCekLokasi.t.Text = Lokasi
frmCekLokasi.Show vbModal
RegA.SaveSettingString HKEY_LOCAL_MACHINE, "Software\DataX Active Object\StockJual", "DataLoc", RegC.Enkrip(LokasiFile)
Lokasi = RegC.Dekrip(RegA.GetSettingString(HKEY_LOCAL_MACHINE, "Software\DataX Active Object\StockJual", "DataLoc", "2"))
LokasiFile = Lokasi
Dim PDB As String
PDB = RegC.Dekrip("CE844026980376A603FAF1E87EA8D57DFA46282891C30CEA2096E5857F3E57375D002FF2B421F7A8ACED0F484C52B6BE")
  If Right(Lokasi, 1) = "\" Then
  Lokasi = Left(Lokasi, Len(Lokasi) - 1)
  End If
Lokasi = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Lokasi & ";Persist Security Info=False;Jet OLEDB:Database Password=" & PDB

DoEvents
SaveBack
frmSplash.Show vbModal
End Sub

Public Function NoUtang(Awal As String, Tabel As String) As String
On Error GoTo aa
  'HXXXX/mm/yy
 Dim rsData As New ADODB.Recordset, KKO As String
 Set rsData = aData.AmbilData("select max(val(mid([KodeBayar],2,4))) from [" & Tabel & "] " & _
 "where left([KodeBayar],1)='" & Awal & "' and right([KodeBayar],6)='" & _
 Format(Date, "/mm/yy") & "'")
 KKO = IIf(IsNull(rsData(0)), 0, rsData(0)) + 1
 NoUtang = Awal & Format(KKO, "0000") & Format(Date, "/mm/yy")
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Function


Public Function Fakktur() As CRAXDdRT.Report
Dim Rrpt As New CRAXDdRT.Application
Set Fakktur = Rrpt.OpenReport(App.Path & "\faktur.rpt")
End Function



Public Sub CopyFile(SourceFile As String, DestFile As String)
'---------------------------------------------------------------
' PURPOSE: Copy a file on disk from one location to another.
' ACCEPTS: The name of the source file and destination file.
' RETURNS: Nothing
'---------------------------------------------------------------
  Dim Result As Long
   If Dir(SourceFile) = "" Then
      MsgBox Chr(34) & SourceFile & Chr(34) & _
         " is not valid file name."
   Else
      Result = apiCopyFile(SourceFile, DestFile, False)
   End If
End Sub

Public Function GetSpecialFolder(CSIDL As Long) As String
  Dim idlstr As Long
  Dim sPath As String
  Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260

On Error GoTo Err_GetFolder
' Fill the idl structure with the specified folder item.
idlstr = SHGetSpecialFolderLocation(Utama.hWnd, CSIDL, IDL)
          If idlstr = NOERROR Then
' Get the path from the idl list, and return the folder with a slash at the end.
sPath = Space$(MAX_LENGTH)
idlstr = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
If idlstr Then
GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1) & "\"
End If
End If
Exit_GetFolder:
Exit Function
Err_GetFolder:
GetSpecialFolder = ""
End Function

Public Function SaveBack() As String
Dim NamaFile As String, DiMana As String, i As Integer, Direktori As String
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

NamaFile = "BU" & Format(Date, "ddmmyyyy") '& Format(Time, "hhmmss") & ".mdb"
CopyFile DiMana, Direktori & "\" & NamaFile
SetAttr Direktori & "\" & NamaFile, vbHidden + vbReadOnly + vbSystem

For i = 7 To 14
NamaFile = "BU" & Format(Date - i, "ddmmyyyy")
On Error Resume Next
SetAttr Direktori & "\" & NamaFile, vbNormal
Kill Direktori & "\" & NamaFile
If Err.Number <> 53 And Err.Number <> 0 Then GoTo aa
On Error GoTo aa
Next i

Exit Function
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur SaveBack pada Module SerbaGuna"
End Function



Public Sub AturLaporan(Jenis As String, Optional JenisTrans As String, Optional NoFak As String, Optional PrintLangsung As Boolean)
On Error GoTo aa
Dim kTabel As String, strJudul As String, Ikut As Boolean
Dim ForX As CRAXDdRT.FormulaFieldDefinition
Screen.MousePointer = vbHourglass
Set fRpt = New Form2
Select Case Jenis
Case "FakturPrint"
  Set fRpt.Report = Lap.LapBJFaktur
  strJudul = ""
  If JenisTrans = "" Then
    kTabel = "select * from xBeliDet where [No Faktur]='" & AmanOi(NoFak) & "' union " & _
    "select * from xJualDet where [No Faktur]='" & AmanOi(NoFak) & "'"
  Else
    kTabel = "select * from " & IIf(JenisTrans = "B", "xBeliDet ", "xJualDet ") & "where [No Faktur]='" & AmanOi(NoFak) & "'"
  End If
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3
  
Case "MasterSupplier", "MasterKonsumen", "MasterSalesman"
  Set fRpt.Report = Lap.LapFileMaster
  kTabel = Right(Jenis, Len(Jenis) - 6)
  fRpt.Report.Database.SetDataSource aData.AmbilData(kTabel), 3

Case "FakturB", "FakturJ", "TokoBulanFakturB", "TokoBulanFakturJ", _
     "ItemBarangB", "ItemBarangJ"
  If Jenis = "FakturB" Or Jenis = "FakturJ" Then
    Set fRpt.Report = Lap.LapBJFaktur
  ElseIf Jenis = "TokoBulanFakturB" Or Jenis = "TokoBulanFakturJ" Then
    Set fRpt.Report = Lap.LapTokoBulanFaktur
  ElseIf Jenis = "ItemBarangB" Or Jenis = "ItemBarangJ" Then
    Set fRpt.Report = Lap.LapBJItemBarang
  End If
  Screen.MousePointer = vbNormal
  BantuDataLap.Show vbModal
  kTabel = "select * from " & IIf(Right(Jenis, 1) = "B", "xBeliDet ", "xJualDet ")
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & "where tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    strJudul = "Semua Data"
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3

Case "BSupp", "JCust", "JualSales", "BPlus", "JPlus", "BTransPerFak", "JTransPerFak"
  If Jenis = "BSupp" Or Jenis = "JCust" Then
    Set fRpt.Report = Lap.LapBJPihak
  ElseIf Jenis = "JualSales" Then
    Set fRpt.Report = Lap.LapJualSales
  ElseIf Jenis = "BPlus" Or Jenis = "JPlus" Then
    Set fRpt.Report = Lap.LapBJRetur
  ElseIf Jenis = "BTransPerFak" Or Jenis = "JTransPerFak" Then
    Set fRpt.Report = Lap.LapTransPerFak
  End If
  Screen.MousePointer = vbNormal
  BantuDataLap.Show vbModal
  kTabel = "select * from " & IIf(Left(Jenis, 1) = "B", "xBeliGen ", "xJualGen ")
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & "where tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
    If Right(Jenis, 4) = "Plus" Then kTabel = kTabel & " and Jenistrans='" & JenisTrans & "'"
  Else
    strJudul = "Semua Data"
    If Right(Jenis, 4) = "Plus" Then kTabel = kTabel & " where Jenistrans='" & JenisTrans & "'"
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3

Case "BItemBrgTgl"
  Set fRpt.Report = Lap.LapBeliBrgPerTgl
  Screen.MousePointer = vbNormal
  BantuDataLap.Show vbModal
  kTabel = "SELECT TransBeli.Tgl, DetailBeli.[Kode Barang], Barang.[Nama Barang], Sum(DetailBeli.Qty) AS Qty, sum(DetailBEli.harga-DetailBeli.Diskon) as Harga " & _
  "FROM TransBeli INNER JOIN (Barang INNER JOIN DetailBeli ON Barang.[Kode Barang] = DetailBeli.[Kode Barang]) ON TransBeli.intno = DetailBeli.intTrans " & _
  "Where hapus=0 and transbeli.jenistrans='B' "
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & "and tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    strJudul = "Semua Data"
  End If
  Unload BantuDataLap
  kTabel = kTabel & "GROUP BY TransBeli.Tgl, DetailBeli.[Kode Barang], Barang.[Nama Barang];"
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3

Case "StockMin"
  Set fRpt.Report = Lap.LapStockMin
  kTabel = "Select [Kode Barang], [Nama Barang], Qty, [Qty Min] from Barang where Qty<[Qty Min] "
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3

Case "StockQty"
  Set fRpt.Report = Lap.LapStockQty
  'kTabel = "Select [Kode Barang], [Nama Barang], Qty from Barang "
  kTabel = "Select * from Barang "
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3

Case "StockKartu", "StockKartuQty"
  If Jenis = "StockKartu" Then
    Set fRpt.Report = Lap.LapStockKartu
  ElseIf Jenis = "StockKartuQty" Then
    Set fRpt.Report = Lap.LapStockKartuQty
  End If
  
  Screen.MousePointer = vbNormal
  BantuDataLap.t2.Enabled = False
  BantuDataLap.Show vbModal
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = "SELECT xBeliDet.*, Barang.Qty AS AkhirQty FROM Barang INNER JOIN xBeliDet ON Barang.[Kode Barang] = xBeliDet.[Kode Barang] " & _
             "where tgl >= #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
             "# " & _
             "Union " & _
             "SELECT xJualDet.*, Barang.Qty AS AkhirQty FROM Barang INNER JOIN xJualDet ON Barang.[Kode Barang] = xJualDet.[Kode Barang]" & _
             "where tgl >= #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
             "# "
    strJudul = "Mulai dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy")
  Else
    kTabel = "SELECT xBeliDet.*, Barang.Qty AS AkhirQty FROM Barang INNER JOIN xBeliDet ON Barang.[Kode Barang] = xBeliDet.[Kode Barang] " & _
             "Union " & _
             "SELECT xJualDet.*, Barang.Qty AS AkhirQty FROM Barang INNER JOIN xJualDet ON Barang.[Kode Barang] = xJualDet.[Kode Barang]"
    strJudul = "Semua Data"
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3

Case "StockPersBD"
  Set fRpt.Report = Lap.LapStockPersBD
  fRpt.Report.Database.SetDataSource aData.DataLapStock_PersBD, 3
  
Case "HDaftar", "PDaftar", "HDaftarNoGiro", "PDaftarNoGiro"
  Set fRpt.Report = Lap.LapUtPiDaftar
  Screen.MousePointer = vbNormal
  If MsgBox("Tampilkan data faktur pembayaran yang telah selesai..?", vbInformation + vbYesNo + vbDefaultButton2, "Pilih Data") = vbYes Then
  Ikut = True
  Else
  Ikut = False
  End If
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.DataLapHP_Daftar(Jenis, Ikut), 3

Case "HPDaftarAnak"
  Set fRpt.Report = Lap.LapUtPiDaftarAnak
  fRpt.Report.Database.SetDataSource aData.DataLapHP_Daftar_Rinci(JenisTrans, NoFak), 3

Case "HJatuh", "PJatuh"
 Set fRpt.Report = Lap.LapUtPiJthTempo
 fRpt.Report.Database.SetDataSource aData.DataLapHP_Jatuh(Left(Jenis, 1)), 3

Case "HBayarSupp", "PBayarCust", "HBayarTgl", "PBayarTgl"
  If Right(Jenis, 3) = "Tgl" Then
    Set fRpt.Report = Lap.LapUtPiBayarTgl
  Else
    Set fRpt.Report = Lap.LapUtPiBayarSupplier
  End If
  Screen.MousePointer = vbNormal
  BantuDataLap.Show vbModal
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & "#" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    kTabel = ""
    strJudul = "Semua Data"
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.DataLapHP_Bayar(Left(Jenis, 1), kTabel), 3

Case "HGiro", "PGiro"
  Set fRpt.Report = Lap.LapUtPiGiro
  kTabel = ""
  strJudul = ""
  fRpt.Report.Database.SetDataSource aData.DataLapHP_Giro(Left(Jenis, 1)), 3

Case "LRItem", "LRSales", "LRFaktur"
  If Jenis = "LRItem" Then
    Set fRpt.Report = Lap.LapLRItem
  ElseIf Jenis = "LRSales" Then
    Set fRpt.Report = Lap.LapLRSales
  ElseIf Jenis = "LRFaktur" Then
    Set fRpt.Report = Lap.LapLRFaktur
  End If
  Screen.MousePointer = vbNormal
  kTabel = "SELECT TransJual.intno, " & _
           "TransJual.[No Faktur], TransJual.Tgl, TransJual.JenisTrans, TransJual.Jenis, " & _
           "DetailJual.[Kode Barang], Barang.[Nama Barang], " & _
           "IIf(JenisTrans='J',DetailJual.Qty,-DetailJual.Qty) AS Qty, " & _
           "iif(IIf(JenisTrans='J',DetailJual.Qty,-DetailJual.Qty)* " & _
           "((DetailJual.Harga-DetailJual.Diskon-DetailJual.PH)-DetailJual.HRata )<0,'Rugi','Laba') as LaR, " & _
           "(1-(transjual.diskon/iif(transjual.total=0,1,transjual.total)))*(DetailJual.Harga-DetailJual.Diskon)-DetailJual.PH AS Harga, DetailJual.HRata as HModal, " & _
           "TransJual.Salesman, Salesman.Nama, TransJual.Diskon AS BigDisc, TransJual.Kepada, " & _
           "Konsumen.Nama FROM Konsumen INNER JOIN ((Salesman INNER JOIN TransJual " & _
           "ON Salesman.Kode = TransJual.Salesman) INNER JOIN (Barang INNER JOIN DetailJual " & _
           "ON Barang.[Kode Barang] = DetailJual.[Kode Barang]) ON TransJual.intno = " & _
           "DetailJual.intTrans) ON Konsumen.Kode = TransJual.Kepada " & _
           "WHERE TransJual.Hapus=0"
  
  BantuDataLap.Show vbModal
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & " and tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    strJudul = "Semua Data"
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3


Case "Biaya"
  Set fRpt.Report = Lap.LapBiaya
  Screen.MousePointer = vbNormal
  kTabel = "SELECT * FROM Biaya " & _
           "WHERE 0=0 "
  BantuDataLap.Show vbModal
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & " and tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    strJudul = "Semua Data"
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3
  
  
Case "RekapData"
  Set fRpt.Report = Lap.LapRekap
  Screen.MousePointer = vbNormal
  kTabel = "SELECT * FROM Biaya " & _
           "WHERE 0=0 "
  BantuDataLap.Show vbModal
  Dim aCrit As String
  If BantuDataLap.Check1.Value = vbUnchecked Then
    aCrit = " and tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "#"
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    strJudul = "Semua Data"
    aCrit = ""
  End If
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  
Dim aRSet As New ADODB.Recordset, kSet As New ADODB.Recordset
  kSet.Fields.Append "A", adVarChar, 50
  kSet.Fields.Append "b", adVarChar, 50
  kSet.Fields.Append "c", adCurrency
  kSet.Open
  
  'Beli Kredit
  kTabel = "SELECT sum(Total- Diskon) as a From TransBeli " & _
  " WHERE (TransBeli.Jenis<>'C') AND (TransBeli.Hapus=0) AND (TransBeli.JenisTrans='B')" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Pembelian": kSet!b = "Pembelian Kredit": kSet!C = IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  'Beli Cash
  kTabel = "SELECT sum(Total- Diskon) as a From TransBeli " & _
  " WHERE (TransBeli.Jenis='C') AND (TransBeli.Hapus=0) AND (TransBeli.JenisTrans='B')" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Pembelian": kSet!b = "Pembelian Cash": kSet!C = IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  'Retur Beli Kredit
  kTabel = "SELECT sum(Total- Diskon) as a From TransBeli " & _
  " WHERE (TransBeli.Hapus=0) AND (TransBeli.JenisTrans='RB')" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Pembelian": kSet!b = "Retur Pembelian": kSet!C = -1 * IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  'Jual Kredit
  kTabel = "SELECT sum(Total- Diskon) as a From TransJual " & _
  " WHERE (TransJual.Jenis<>'C') AND (TransJual.Hapus=0) AND (TransJual.JenisTrans='J')" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Penjualan": kSet!b = "Penjualan Kredit": kSet!C = IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  'Jual Cash
  kTabel = "SELECT sum(Total- Diskon) as a From TransJual " & _
  " WHERE (TransJual.Jenis='C') AND (TransJual.Hapus=0) AND (TransJual.JenisTrans='J')" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Penjualan": kSet!b = "Penjualan Cash": kSet!C = IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  'Retur Jual Kredit
  kTabel = "SELECT sum(Total- Diskon) as a From TransJual " & _
  " WHERE (TransJual.Hapus=0) AND (TransJual.JenisTrans='RJ')" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Penjualan": kSet!b = "Retur Penjualan": kSet!C = -1 * IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  'Biaya
  kTabel = "SELECT sum(Jumlah) as a From Biaya " & _
  " WHERE (0=0)" & _
  aCrit
  Set aRSet = aData.AmbilCommand(kTabel)
  kSet.AddNew: kSet!a = "Biaya": kSet!b = "Biaya": kSet!C = IIf(IsNull(aRSet!a), 0, aRSet!a): kSet.Update
  
  
  fRpt.Report.Database.SetDataSource kSet, 3
   
Case "OmzetSales"
  Set fRpt.Report = Lap.LapOmzetSales
  Screen.MousePointer = vbNormal
  
  kTabel = "SELECT TransJual.Salesman, Salesman.Nama as SalesNama, InputPiutang.Kepada, " & _
    "InputPiutang.Nama, InputPiutang.Transaksi, Sum(InputPiutang.Total) AS " & _
    "SumOfTotal, Sum(InputPiutang.Bayar) AS SumOfBayar, Sum(InputPiutang.Sisa) " & _
    "AS SumOfSisa FROM Salesman INNER JOIN (TransJual INNER JOIN InputPiutang " & _
    "ON TransJual.[No Faktur] = InputPiutang.[No Faktur]) ON Salesman.Kode = " & _
    "TransJual.Salesman " & _
    "WHERE 0=0 "
  BantuDataLap.Show vbModal
  If BantuDataLap.Check1.Value = vbUnchecked Then
    kTabel = kTabel & " and InputPiutang.tgl between #" & Format(BantuDataLap.t1.Value, "mm/dd/yyyy") & _
    "# and #" & Format(BantuDataLap.t2.Value, "mm/dd/yyyy") & "# "
    strJudul = "Dari tanggal " & Format(BantuDataLap.t1.Value, "dd mmmm yyyy") & _
               " sampai " & Format(BantuDataLap.t2.Value, "dd mmmm yyyy")
  Else
    strJudul = "Semua Data"
  End If
  kTabel = kTabel & " GROUP BY TransJual.Salesman, Salesman.Nama, " & _
    "InputPiutang.Kepada, InputPiutang.Nama, InputPiutang.Transaksi "
  Unload BantuDataLap
  Screen.MousePointer = vbHourglass
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3
  
  Case "PArea", "HArea"
  Set fRpt.Report = Lap.LapUtpiArea
  If Jenis = "HArea" Then
    kTabel = "SELECT 'H' AS Judul, " & _
    "Supplier.Kode, Supplier.Nama as Nama, Supplier.Wilayah, Wilayah.Nama, Sum(InputUtang.Total) AS SumOfTotal, Sum(InputUtang.Bayar) AS SumOfBayar, Sum(InputUtang.Sisa) AS SumOfSisa " & _
    "FROM InputUtang INNER JOIN (Wilayah INNER JOIN Supplier ON Wilayah.Kode = Supplier.Wilayah) ON InputUtang.Kepada = Supplier.Kode " & _
    "GROUP BY Supplier.Kode, Supplier.Nama, Supplier.Wilayah, Wilayah.Nama"
  Else
    kTabel = "SELECT 'P' AS Judul, " & _
    "Konsumen.Kode, Konsumen.Nama as Nama, Konsumen.Wilayah, Wilayah.Nama, Sum(InputPiutang.Total) AS SumOfTotal, Sum(InputPiutang.Bayar) AS SumOfBayar, Sum(InputPiutang.Sisa) AS SumOfSisa " & _
    "FROM InputPiutang INNER JOIN (Wilayah INNER JOIN Konsumen ON Wilayah.Kode = Konsumen.Wilayah) ON InputPiutang.Kepada = Konsumen.Kode " & _
    "GROUP BY Konsumen.Kode, Konsumen.Nama, Konsumen.Wilayah, Wilayah.Nama"
    'kTabel = "SELECT 'P' AS Judul, TransJual.Kepada, Konsumen.Nama AS Nama, Sum(IIf(TransJual.jenistrans='J',TransJual.Total-TransJual.diskon,-(TransJual.Total-TransJual.diskon))) AS Total, IIf(IsNull(Sum(PiutangStatusNoBatal.Total+PiutangStatusNoBatal.Potongan)),0,Sum(PiutangStatusNoBatal.Total+PiutangStatusNoBatal.Potongan)) AS Bayar, total-bayar AS Sisa, Konsumen.Wilayah as Kota " & _
    '"FROM Konsumen INNER JOIN (TransJual LEFT JOIN PiutangStatusNoBatal ON TransJual.intno = PiutangStatusNoBatal.KodeFaktur) ON Konsumen.Kode = TransJual.Kepada " & _
    '"WHERE (((TransJual.Jenis)<>'C') AND ((TransJual.Hapus)=0)) and TransJual.Jenis<>'OK' " & _
    '"GROUP BY TransJual.Kepada, Konsumen.Nama, Konsumen.Wilayah, TransJual.JenisTrans " & _
    '"ORDER BY TransJual.Kepada"
  End If
  Screen.MousePointer = vbNormal
  fRpt.Report.Database.SetDataSource aData.AmbilCommand(kTabel), 3
  
End Select
  
  For Each ForX In fRpt.Report.FormulaFields
  If ForX.Name = "{@DKTgl}" Then ForX.Text = Chr(34) & strJudul & Chr(34)
  Next
If PrintLangsung Then
fRpt.Report.PrintOut
Else
fRpt.aView.ReportSource = fRpt.Report
fRpt.aView.ViewReport
fRpt.Show
End If
Screen.MousePointer = vbNormal

Exit Sub
aa:
Dim Err_Setering As String
Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur AturLaporan pada " & "Module SerbaGuna di baris " & Erl
Select Case MsgBox(Err_Setering, vbRetryCancel, App.Title & "-SerbaGuna Error")
  Case vbCancel: Resume Exit_AturLaporan:
  Case vbRetry: Resume
  Case Else: End
End Select
Exit_AturLaporan:
App.LogEvent "myAS=>" & Format(Date, "dd-mmmm-yyyy") & Format(Time, "(hh:mm:ss)") & _
vbCrLf & Err_Setering & vbCrLf, vbLogEventTypeError
Screen.MousePointer = vbNormal
End Sub
