VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form cBantu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4860
   ClientLeft      =   1755
   ClientTop       =   3450
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
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
      Left            =   9810
      TabIndex        =   3
      Top             =   90
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   555
      Left            =   -2000
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin TDBNumber6Ctl.TDBNumber t1 
      Height          =   405
      Left            =   2025
      TabIndex        =   1
      Top             =   90
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calculator      =   "cBantu.frx":0000
      Caption         =   "cBantu.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "cBantu.frx":008C
      Keys            =   "cBantu.frx":00AA
      Spin            =   "cBantu.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   2
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "##,###,###,##0.##"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   1
      ForeColor       =   -2147483640
      Format          =   "##,###,###,##0.##"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   5
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   -99999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   1
      ValueVT         =   96731137
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VSFlex8Ctl.VSFlexGrid VGrid 
      Height          =   2910
      Left            =   135
      TabIndex        =   2
      Top             =   630
      Width           =   11370
      _cx             =   20055
      _cy             =   5133
      Appearance      =   1
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
      BackColorSel    =   16711390
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"cBantu.frx":011C
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
      ExplorerBar     =   3
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5670
      TabIndex        =   6
      Top             =   180
      Width           =   3990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Potongan Faktur :"
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
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   1890
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   135
      TabIndex        =   4
      Top             =   3600
      Width           =   11370
   End
End
Attribute VB_Name = "cBantu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IntNoFak As String, Keterangan As String, Jenis As String
Public Batal As Boolean
Public NilaiBantu
Dim TotalForm As Currency

Private Sub Command2_Click()
On Error GoTo aa
Batal = True
Unload Me
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command2_Click pada Form cBantu"
End Sub

Private Sub Command3_Click()
Dim BolehRugi As Boolean
On Error GoTo aa
BolehRugi = False
NilaiBantu = TotalForm
Dim i As Single
ReDim xPotHar(VGrid.Rows - 1, 3)

For i = 1 To VGrid.Rows - 1
  If (VGrid.ValueMatrix(i, 3) - VGrid.ValueMatrix(i, 5) - VGrid.ValueMatrix(i, 6)) < VGrid.ValueMatrix(i, 4) And Jenis = "Piutang" Then
      If Not BolehRugi Then
      frmPass.Jenis = "Rugi"
      frmPass.lblLabels(0).Caption = "Potongan harga untuk " & Me.Caption & _
      "melewati harga modal dan ditandai oleh baris yg bewarna merah.." & vbCrLf & _
      "Untuk melanjutkan potongan transaksi masukkan password."
      frmPass.Show vbModal
        If Not frmPass.LoginOK Then
        ReDim xPotHar(0, 3)
        Exit Sub
        Else
        BolehRugi = True
        End If
      End If
    End If
  xPotHar(i - 1, 0) = IntNoFak
  xPotHar(i - 1, 1) = VGrid.ValueMatrix(i, 7)
  xPotHar(i - 1, 2) = VGrid.ValueMatrix(i, 6)
Next i
Batal = False
Unload Me
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command3_Click pada Form cBantu"
End Sub

Private Sub Form_Load()
On Error GoTo aa
Dim RSSData As New ADODB.Recordset, rsTotData As New ADODB.Recordset
Dim PotGlob As Currency

Me.Caption = Keterangan
Batal = False
t1.Value = 0
If Jenis = "Utang" Then
Set RSSData = aData.AmbilCommand("SELECT TransBeli.intno, DetailBeli.intno as xNo,TransBeli.[No Faktur], " & _
"DetailBeli.[Kode Barang], Barang.[Nama Barang], DetailBeli.Qty, " & _
"([DetailBeli].[Harga]-[DetailBeli].[Diskon]) AS Harga, " & _
"0 AS Modal, DetailBeli.PH " & _
"FROM TransBeli INNER JOIN (Barang INNER JOIN DetailBeli ON " & _
"Barang.[Kode Barang] = DetailBeli.[Kode Barang]) ON " & _
"TransBeli.intno = DetailBeli.intTrans where transBeli.intno=" & _
IntNoFak)
Set rsTotData = aData.AmbilCommand("SELECT TransBeli.intno, " & _
"Sum(DetailUtang.Potongan) AS SumOfPotongan " & _
"FROM TransBeli INNER JOIN ((Utang INNER JOIN UtangAnak " & _
"ON Utang.intno = UtangAnak.KodeBayar) INNER JOIN DetailUtang " & _
"ON UtangAnak.intno = DetailUtang.KodeBayar) ON TransBeli.intno " & _
"= DetailUtang.KodeFaktur where transBeli.intno=" & _
IntNoFak & " GROUP BY TransBeli.intno, TransBeli.[No Faktur], " & _
"UtangAnak.Status HAVING (UtangAnak.Status)<>'Batal'")
Else
Set RSSData = aData.AmbilCommand("SELECT TransJual.intno, DetailJual.intno as xNo,TransJual.[No Faktur], " & _
"DetailJual.[Kode Barang], Barang.[Nama Barang], DetailJual.Qty, " & _
"([DetailJual].[Harga]-[DetailJual].[Diskon]) AS Harga, " & _
"DetailJual.HRata AS Modal, DetailJual.PH " & _
"FROM TransJual INNER JOIN (Barang INNER JOIN DetailJual ON " & _
"Barang.[Kode Barang] = DetailJual.[Kode Barang]) ON " & _
"TransJual.intno = DetailJual.intTrans where transjual.intno=" & _
IntNoFak)
Set rsTotData = aData.AmbilCommand("SELECT TransJual.intno, " & _
"Sum(DetailPiutang.Potongan) AS SumOfPotongan " & _
"FROM TransJual INNER JOIN ((Piutang INNER JOIN PiutangAnak " & _
"ON Piutang.intno = PiutangAnak.KodeBayar) INNER JOIN DetailPiutang " & _
"ON PiutangAnak.intno = DetailPiutang.KodeBayar) ON TransJual.intno " & _
"= DetailPiutang.KodeFaktur where transjual.intno=" & _
IntNoFak & " GROUP BY TransJual.intno, TransJual.[No Faktur], " & _
"PiutangAnak.Status HAVING (PiutangAnak.Status)<>'Batal'")
End If

With VGrid
.Rows = 1
PotGlob = 0
  Do While Not RSSData.EOF
  .Rows = .Rows + 1
  .Row = .Rows - 1
  .TextMatrix(.Row, 0) = RSSData![Kode Barang]
  .TextMatrix(.Row, 1) = RSSData![Nama Barang]
  .TextMatrix(.Row, 2) = RSSData![qty]
  .TextMatrix(.Row, 3) = RSSData![Harga]
  .TextMatrix(.Row, 4) = RSSData!Modal
  .TextMatrix(.Row, 5) = RSSData!PH
  .TextMatrix(.Row, 6) = 0
  .TextMatrix(.Row, 7) = RSSData!xNo
  PotGlob = PotGlob + (RSSData!PH * RSSData![qty])
  RSSData.MoveNext
  Loop
  On Error Resume Next: .Row = 1: .Col = 6
  On Error GoTo aa
  If Jenis <> "Piutang" Then .ColHidden(4) = True
End With

If Not (rsTotData.EOF) Then
    Label1.Caption = "Faktur Lama : " & vbCrLf & _
    "  Total Potongan : " & FormatNumber(rsTotData!SumOfPotongan, 2) & vbCrLf & _
    "  Potongan Harga : " & FormatNumber(PotGlob, 2) & vbCrLf & _
    "  Potongan Faktur : " & FormatNumber(rsTotData!SumOfPotongan - PotGlob, 2)
Else
    Label1.Caption = "Faktur Lama : " & vbCrLf & _
    "  Total Potongan : " & 0 & vbCrLf & _
    "  Potongan Harga : " & 0 & vbCrLf & _
    "  Potongan Faktur : " & 0
End If

Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Load pada Form bBantu"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo aa
If Batal Then
NilaiBantu = 0
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Unload pada Form bBantu"
End Sub

Private Sub t1_Validate(Cancel As Boolean)
On Error GoTo aa
Call Hitung
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur t5_Validate pada Form cBantu"
End Sub

Private Sub VGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
Call Hitung
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur VGrid_AfterEdit pada Form bBantu"
End Sub

Private Sub VGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
If Col <> 6 Then Cancel = True
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur VGrid_StartEdit pada Form cBantu"
End Sub

Private Sub Hitung()
On Error GoTo aa
  Dim i As Single, PotHar As Currency
  With VGrid
  PotHar = 0
  For i = 1 To .Rows - 1
  PotHar = PotHar + .ValueMatrix(i, 2) * .ValueMatrix(i, 6)
    If (.ValueMatrix(i, 3) - .ValueMatrix(i, 5) - .ValueMatrix(i, 6)) < .ValueMatrix(i, 4) And Jenis = "Piutang" Then
    .Cell(flexcpForeColor, i, 0, i, 7) = vbRed
    Else
    .Cell(flexcpForeColor, i, 0, i, 7) = vbBlack
    End If
  Next i
  End With
  TotalForm = PotHar + t1.Value
  Label4.Caption = "Total : " & FormatNumber(TotalForm, 2)
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Hitung pada Form cBantu"
End Sub
