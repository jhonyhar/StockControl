VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHisTrans 
   Caption         =   "Histori Transaksi"
   ClientHeight    =   9600
   ClientLeft      =   315
   ClientTop       =   1050
   ClientWidth     =   14835
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pembelian"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Penjualan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   1560
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox t3 
      DataField       =   "No Faktur"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   12
      Top             =   1350
      Width           =   3435
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6120
      Top             =   9900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load &Data"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtFields 
      DataField       =   "JenisTrans"
      Height          =   315
      Left            =   540
      TabIndex        =   5
      Top             =   9990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   5520
      Left            =   435
      TabIndex        =   2
      Top             =   3225
      Width           =   13860
      _cx             =   24447
      _cy             =   9737
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
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
      AllowUserResizing=   3
      SelectionMode   =   0
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
      FormatString    =   $"frmHisTrans.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin TDBDate6Ctl.TDBDate t1 
      Height          =   405
      Left            =   2040
      TabIndex        =   4
      Top             =   825
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calendar        =   "frmHisTrans.frx":0128
      Caption         =   "frmHisTrans.frx":025E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHisTrans.frx":02CC
      Keys            =   "frmHisTrans.frx":02EA
      Spin            =   "frmHisTrans.frx":0348
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
      Left            =   7125
      TabIndex        =   10
      Top             =   825
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calendar        =   "frmHisTrans.frx":0370
      Caption         =   "frmHisTrans.frx":04A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHisTrans.frx":0514
      Keys            =   "frmHisTrans.frx":0532
      Spin            =   "frmHisTrans.frx":0590
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   1800
      Width           =   4740
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nama Toko :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sampai Tgl :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5790
      TabIndex        =   11
      Top             =   855
      Width           =   1260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+F4=Close F4=Bantuan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   465
      TabIndex        =   3
      Top             =   9000
      Width           =   6285
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dari Tgl :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1020
      TabIndex        =   0
      Top             =   855
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Histori Transaksi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
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
      TabIndex        =   8
      Top             =   105
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   9135
      Left            =   270
      Top             =   210
      Width           =   14280
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "F4=Input, F5=Sort, F6=Filter, F7=Form View, F8=Print, F9=Refresh, F10=Search, Alt+X=Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   420
      TabIndex        =   9
      Top             =   7860
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
      TabIndex        =   7
      Top             =   195
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   9375
      Left            =   165
      Top             =   105
      Width           =   14535
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9000
      TabIndex        =   6
      Top             =   9495
      Width           =   1665
   End
End
Attribute VB_Name = "frmHisTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BaruBuka As Boolean
Public JenisJenis As String
Dim edRow As Integer, edCol  As Integer
Dim intNoFaktur As Single
Dim PtoScreen As Boolean

Private Sub Command1_Click()
On Error Resume Next
Master.ColWidth(0) = 2000
Master.ColWidth(1) = 1335
Master.ColWidth(2) = 3870
Master.ColWidth(3) = 1245
Master.ColWidth(4) = 1455
Master.ColWidth(5) = 1125
Master.ColWidth(6) = 1725
Master.ColWidth(7) = 800

Dim aStr As String
If Option1.Value Then
aStr = "SELECT TransJual.intno, TransJual.[No Faktur], TransJual.Kepada & ' - ' & Konsumen.Nama AS Kepada, TransJual.Salesman AS Salesman, TransJual.Tgl, " & _
"DetailJual.[Kode Barang], Barang.[Nama Barang], DetailJual.Qty, DetailJual.Harga, DetailJual.Diskon AS Diskon, DetailJual.intno AS DetintNo, DetailJual.QtyB & ' ' & Barang.Satuan & ' ' & DetailJual.QtyS & ' ' & Barang.[Satuan Kecil] AS TampilSat " & _
"FROM (Salesman INNER JOIN (Konsumen INNER JOIN TransJual ON Konsumen.Kode=TransJual.Kepada) ON Salesman.Kode=TransJual.Salesman) INNER JOIN (Barang INNER JOIN DetailJual ON Barang.[Kode Barang]=DetailJual.[Kode Barang]) ON TransJual.intno=DetailJual.intTrans " & _
"WHERE (TransJual.Hapus=0) and (TransJual.JenisTrans='J') and " & _
"(TransJual.Tgl between #" & Format(t1.Value, "mm/dd/yyyy") & _
"# and #" & Format(t2.Value, "mm/dd/yyyy") & "#) and " & _
"TransJual.Kepada='" & AmanOi(t3.Text) & "'"
Else
aStr = "SELECT TransBeli.intno, TransBeli.[No Faktur], TransBeli.Kepada & ' - ' & Supplier.Nama AS Kepada, '-' AS Salesman, TransBeli.Tgl, " & _
"DetailBeli.[Kode Barang], Barang.[Nama Barang], DetailBeli.Qty, DetailBeli.Harga, DetailBeli.Diskon AS Diskon, DetailBeli.intno AS DetintNo, DetailBeli.QtyB & ' ' & Barang.Satuan & ' ' & DetailBeli.QtyS & ' ' & Barang.[Satuan Kecil] AS TampilSat " & _
"FROM (Supplier INNER JOIN TransBeli ON Supplier.Kode=TransBeli.Kepada) INNER JOIN (Barang INNER JOIN DetailBeli ON Barang.[Kode Barang]=DetailBeli.[Kode Barang]) ON TransBeli.intno=DetailBeli.intTrans " & _
"WHERE (TransBeli.Hapus=0) and (TransBeli.JenisTrans='B') and " & _
"(TransBeli.Tgl between #" & Format(t1.Value, "mm/dd/yyyy") & _
"# and #" & Format(t2.Value, "mm/dd/yyyy") & "#) and " & _
"TransBeli.Kepada='" & AmanOi(t3.Text) & "'"
End If
Dim kSet As New ADODB.Recordset
Set kSet = aData.AmbilCommand(aStr)
Master.Rows = 1
  'Label8.Caption = ""
  Do While Not kSet.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = kSet![No Faktur]
    Master.TextMatrix(Master.Rows - 1, 1) = kSet![Tgl]
    Master.TextMatrix(Master.Rows - 1, 2) = kSet![Nama Barang]
    Master.TextMatrix(Master.Rows - 1, 3) = kSet![qty]
    Master.TextMatrix(Master.Rows - 1, 4) = kSet![Harga]
    Master.TextMatrix(Master.Rows - 1, 5) = kSet![Diskon]
    Master.TextMatrix(Master.Rows - 1, 6) = (kSet![Harga] - kSet![Diskon]) * kSet![qty]
    Master.TextMatrix(Master.Rows - 1, 7) = kSet![Salesman]
    Master.TextMatrix(Master.Rows - 1, 8) = kSet![TampilSat]
    'Master.TextMatrix(Master.Rows - 1, 10) = kset![jatuhTempo]
    'Master.TextMatrix(Master.Rows - 1, 11) = kset![Sales]
    kSet.MoveNext
  Loop
  Master.MergeCells = flexMergeFree
  Master.MergeCol(2) = True
  Master.Col = 2
  Master.Sort = flexSortGenericAscending

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyX And Shift = 4 Then   '############ TUTUP FORM ############
Unload Me
End If
End Sub

Private Sub Form_Load()
On Error GoTo aa
BaruBuka = True
t1.Value = Date - 30
t2.Value = Date
DoEvents
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
Cancel = True
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     If Option2.Value Then
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
If Option2.Value Then
  Set rsKodok = aData.AmbilCommand("Select * from Supplier where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
   Label6.Caption = IIf(IsNull(rsKodok!nAMA), "", rsKodok!nAMA) & "(" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & ")"
  End If
Else
  Set rsKodok = aData.AmbilCommand("Select * from Konsumen where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
  Label6.Caption = IIf(IsNull(rsKodok!nAMA), "", rsKodok!nAMA) & "(" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & ")"
  End If
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur t3_Validate pada Form frmPembelian"
End Sub
