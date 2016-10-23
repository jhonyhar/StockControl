VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTT 
   Caption         =   "Tanda Terima"
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton pTagih 
      Caption         =   "Cetak &Daftar Tagihan"
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
      Left            =   9000
      TabIndex        =   18
      Top             =   2865
      Width           =   2325
   End
   Begin VB.CommandButton pTanda 
      Caption         =   "&Cetak Tanda Terima"
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
      Left            =   3090
      TabIndex        =   17
      Top             =   2880
      Width           =   2325
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6120
      Top             =   9900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load &Tanda Terima"
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
      Left            =   660
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton CmdBaru 
      Caption         =   "Load Daftar Tagihan"
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
      Left            =   6390
      TabIndex        =   4
      Top             =   2880
      Width           =   2385
   End
   Begin VB.TextBox txtFields 
      DataField       =   "JenisTrans"
      Height          =   315
      Left            =   540
      TabIndex        =   12
      Top             =   9990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   4680
      Left            =   435
      TabIndex        =   5
      Top             =   3585
      Width           =   13860
      _cx             =   24447
      _cy             =   8255
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTT.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
      ExplorerBar     =   5
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
   Begin TDBNumber6Ctl.TDBNumber tot1 
      Height          =   405
      Left            =   8010
      TabIndex        =   9
      Top             =   8400
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmTT.frx":0163
      Caption         =   "frmTT.frx":0183
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTT.frx":01EF
      Keys            =   "frmTT.frx":020D
      Spin            =   "frmTT.frx":0257
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
      ReadOnly        =   -1
      Separator       =   "."
      ShowContextMenu =   1
      ValueVT         =   2089877505
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBDate6Ctl.TDBDate t2 
      Height          =   405
      Left            =   4005
      TabIndex        =   11
      Top             =   690
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calendar        =   "frmTT.frx":027F
      Caption         =   "frmTT.frx":03B5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTT.frx":0423
      Keys            =   "frmTT.frx":0441
      Spin            =   "frmTT.frx":049F
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
   Begin TDBNumber6Ctl.TDBNumber Tot2 
      Height          =   405
      Left            =   2700
      TabIndex        =   7
      Top             =   8400
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmTT.frx":04C7
      Caption         =   "frmTT.frx":04E7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTT.frx":0553
      Keys            =   "frmTT.frx":0571
      Spin            =   "frmTT.frx":05BB
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
      ReadOnly        =   -1
      Separator       =   "."
      ShowContextMenu =   1
      ValueVT         =   2089877505
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   645
      TabIndex        =   1
      Top             =   1200
      Width           =   4770
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Banyak :"
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
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   8475
      Width           =   1815
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6390
      TabIndex        =   3
      Top             =   1200
      Width           =   4935
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
      TabIndex        =   10
      Top             =   9000
      Width           =   6285
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total : "
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
      Index           =   8
      Left            =   6150
      TabIndex        =   8
      Top             =   8475
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tgl:"
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
      Left            =   3390
      TabIndex        =   0
      Top             =   735
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanda Terima Piutang"
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
      TabIndex        =   15
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
      TabIndex        =   16
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   9495
      Width           =   1665
   End
End
Attribute VB_Name = "frmTT"
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

Private Sub pTagih_Click()
On Error GoTo aa
Dim ForX As CRAXDdRT.FormulaFieldDefinition
Screen.MousePointer = vbHourglass
Set fRpt = New Form2
  Set fRpt.Report = Lap.LapTagihan
  
  Dim rsO As New ADODB.Recordset, i As Integer
  rsO.Fields.Append "Tgl", adDate
  rsO.Fields.Append "No Faktur", adVarChar, 50
  rsO.Fields.Append "Kepada", adVarChar, 50
  rsO.Fields.Append "Nama", adVarChar, 50
  rsO.Fields.Append "Sales", adVarChar, 50
  rsO.Fields.Append "Total", adCurrency
  rsO.Fields.Append "Bayar", adCurrency
  rsO.Fields.Append "Sisa", adCurrency
  rsO.Fields.Append "Alamat", adVarChar, 50
  rsO.Fields.Append "JatuhTempo", adInteger
    rsO.Open
  For i = 1 To Master.Rows - 1
    If Master.ValueMatrix(i, 0) Then
      rsO.AddNew
      rsO!Tgl = Master.TextMatrix(i, 2)
      rsO![No Faktur] = Master.TextMatrix(i, 1)
      rsO!Kepada = Master.TextMatrix(i, 7)
      rsO!nAMA = Master.TextMatrix(i, 8)
      rsO!Sales = Master.TextMatrix(i, 11)
      rsO!Total = Master.ValueMatrix(i, 4)
      rsO!Bayar = Master.ValueMatrix(i, 5)
      rsO!Sisa = Master.ValueMatrix(i, 6)
      rsO!Alamat = Master.TextMatrix(i, 9)
      rsO!jatuhtempo = Master.ValueMatrix(i, 10)
      rsO.Update
    End If
  Next i
  fRpt.Report.Database.SetDataSource rsO
  
  
Dim bs() As String
'HasilTT = TglAw;;;TglAk;;;KodW;;;NamW;;;KodS;;;NamS
bs = Split(Label6.Tag, ";;;")
'    strJudul = "Dari tanggal " & Format(bs(2), "dd mmmm yyyy") & _
'               " sampai " & Format(bs(3), "dd mmmm yyyy")
Dim Pa As String, Pb As String
If Right(JenisJenis, 1) = "H" Then
  Pb = "Asia Baru": Pa = bs(1)
Else
  Pa = "Asia Baru": Pb = bs(1)
End If
  For Each ForX In fRpt.Report.FormulaFields
  'If ForX.Name = "{@DS}" Then ForX.Text = Chr(34) & strJudul & Chr(34)
  If ForX.Name = "{@Tgl}" Then ForX.Text = Chr(34) & Format(t2.Value, "dd mmmm yyyy") & Chr(34)
  'If ForX.Name = "{@TBilang}" Then ForX.Text = Chr(34) & RegD.Ubah(tot1.Value, False) & Chr(34)
  If ForX.Name = "{@a}" Then ForX.Text = Chr(34) & Pa & Chr(34)
  'If ForX.Name = "{@b}" Then ForX.Text = Chr(34) & Pb & Chr(34)
  Next
fRpt.aView.ReportSource = fRpt.Report
fRpt.aView.ViewReport
fRpt.Show
Screen.MousePointer = vbNormal

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
Screen.MousePointer = vbNormal
End Sub


Private Sub CmdBaru_Click()
On Error GoTo aa
fBDaftarTagihan.Jenis = Right(JenisJenis, 1)
fBDaftarTagihan.Show vbModal
'HasilTT = TglAw;;;TglAk;;;KodW;;;NamW;;;KodS;;;NamS
If HasilTT <> "" Then
  Dim aStr As String, bs() As String
  bs = Split(HasilTT, ";;;")
  If JenisJenis = "TTP" Then
    aStr = "  SELECT InputPiutang.*, Konsumen.Alamat, TransJual.JatuhTempo,TransJual.Salesman as Sales FROM InputPiutang INNER JOIN (Konsumen INNER JOIN TransJual ON Konsumen.Kode = TransJual.Kepada) ON InputPiutang.intno = TransJual.intno " & _
    "where InputPiutang.tgl between #" & _
    Format(bs(0), "mm/dd/yyyy") & "# and #" & _
    Format(bs(1), "mm/dd/yyyy") & "# and Konsumen.Wilayah='" & AmanOi(bs(2)) & "'"
  Else
    aStr = "  SELECT InputUtang.*, Supplier.Alamat, TransBeli.JatuhTempo,TransBeli.Salesman as Sales FROM InputUtang INNER JOIN (Supplier INNER JOIN TransBeli ON Supplier.Kode = TransBeli.Kepada) ON InputUtang.intno = TransBeli.intno " & _
    "where InputUtang.tgl between #" & _
    Format(bs(0), "mm/dd/yyyy") & "# and #" & _
    Format(bs(1), "mm/dd/yyyy") & "# and Supplier.Wilayah='" & AmanOi(bs(2)) & "'"
  End If
  Dim aRSet As New ADODB.Recordset
  Set aRSet = aData.AmbilCommand(aStr)
  Master.Rows = 1
  Label8.Caption = ""
  Master.ColAlignment(1) = flexAlignLeftCenter
  Do While Not aRSet.EOF
    If aRSet![Sisa] <> 0 Then
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = 1
    Master.TextMatrix(Master.Rows - 1, 1) = aRSet![No Faktur]
    Master.TextMatrix(Master.Rows - 1, 2) = aRSet![Tgl]
    Master.TextMatrix(Master.Rows - 1, 3) = aRSet![Transaksi]
    Master.TextMatrix(Master.Rows - 1, 4) = aRSet![Total]
    Master.TextMatrix(Master.Rows - 1, 5) = aRSet![Bayar]
    Master.TextMatrix(Master.Rows - 1, 6) = aRSet![Sisa]
    Master.TextMatrix(Master.Rows - 1, 7) = aRSet![Kepada]
    Master.TextMatrix(Master.Rows - 1, 8) = aRSet![nAMA]
    Master.TextMatrix(Master.Rows - 1, 9) = IIf(IsNull(aRSet![Alamat]), "", aRSet![Alamat])
    Master.TextMatrix(Master.Rows - 1, 10) = aRSet![jatuhtempo]
    Master.TextMatrix(Master.Rows - 1, 11) = aRSet![Sales]
    Else
     'MsgBox aRSet![No Faktur]
    End If
    aRSet.MoveNext
  Loop
  Master.ColHidden(7) = False
  Label6.Caption = "Daftar Tagihan " & _
  vbCrLf & "Untuk Wilayah : " & bs(2) & "-" & bs(3) & _
  vbCrLf & "Dengan penagih : " & bs(4) & "-" & bs(5) & _
  vbCrLf & "Dari tanggal " & Format(bs(0), "dd mmm yyyy") & _
  " sampai tanggal " & Format(bs(1), "dd mmm yyyy")
  Label6.Tag = HasilTT ' Label8.Tag & bS(0) & ";;;" & bS(1)
  TotalOi
End If


Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = vbShiftMask Then PtoScreen = True
End Sub


Private Sub pTanda_Click()
On Error GoTo aa
Dim kTabel As String, strJudul As String, Ikut As Boolean
Dim ForX As CRAXDdRT.FormulaFieldDefinition
Screen.MousePointer = vbHourglass
Set fRpt = New Form2
  Set fRpt.Report = Lap.LapTandaTerima
  strJudul = ""
  
  Dim rsO As New ADODB.Recordset, i As Integer
  rsO.Fields.Append "Tgl", adDate
  rsO.Fields.Append "No Faktur", adVarChar, 50
  rsO.Fields.Append "Total", adCurrency
  rsO.Open
  For i = 1 To Master.Rows - 1
    If Master.ValueMatrix(i, 0) Then
      rsO.AddNew
      rsO!Tgl = Master.TextMatrix(i, 2)
      rsO![No Faktur] = Master.TextMatrix(i, 1)
      rsO!Total = Master.ValueMatrix(i, 4)
      rsO.Update
    End If
  Next i
  fRpt.Report.Database.SetDataSource rsO
  
Dim bs() As String
'HasilTT = aRSet![Kepada] & ";;;" & aRSet![Nama]tglAw;;;tglAk
bs = Split(Label8.Tag, ";;;")
    strJudul = "Dari tanggal " & Format(bs(2), "dd mmmm yyyy") & _
               " sampai " & Format(bs(3), "dd mmmm yyyy")
Dim Pa As String, Pb As String
If Right(JenisJenis, 1) = "H" Then
  Pb = "Asia Baru": Pa = bs(1)
Else
  Pa = "Asia Baru": Pb = bs(1)
End If
  For Each ForX In fRpt.Report.FormulaFields
  If ForX.Name = "{@DS}" Then ForX.Text = Chr(34) & strJudul & Chr(34)
  If ForX.Name = "{@Tgl}" Then ForX.Text = Chr(34) & Format(t2.Value, "dd mmmm yyyy") & Chr(34)
  If ForX.Name = "{@TBilang}" Then ForX.Text = Chr(34) & terbilang.terbilang(tot1.Value) & Chr(34)
  'If ForX.Name = "{@TBilang}" Then ForX.Text = Chr(34) & RegD.Ubah(tot1.Value, False) & Chr(34)
  If ForX.Name = "{@a}" Then ForX.Text = Chr(34) & Pa & Chr(34)
  If ForX.Name = "{@b}" Then ForX.Text = Chr(34) & Pb & Chr(34)
  Next
fRpt.aView.ReportSource = fRpt.Report
fRpt.aView.ViewReport
fRpt.Show
Screen.MousePointer = vbNormal

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
Screen.MousePointer = vbNormal
End Sub


Private Sub Command1_Click()
On Error GoTo aa
fBTandaTerima.Jenis = Right(JenisJenis, 1)
fBTandaTerima.Show vbModal
'HasilTT = tglAw;;;tglAk;;;Pihak
If HasilTT <> "" Then
  Dim aStr As String, bs() As String
  bs = Split(HasilTT, ";;;")
  If JenisJenis = "TTP" Then
    aStr = "select * from inputpiutang where tgl between #" & _
    Format(bs(0), "mm/dd/yyyy") & "# and #" & _
    Format(bs(1), "mm/dd/yyyy") & "# and Kepada='" & AmanOi(bs(2)) & "'"
  Else
    aStr = "select * from inpututang where tgl between #" & _
    Format(bs(0), "mm/dd/yyyy") & "# and #" & _
    Format(bs(1), "mm/dd/yyyy") & "# and Kepada='" & AmanOi(bs(2)) & "'"
  End If
  Dim aRSet As New ADODB.Recordset
  Set aRSet = aData.AmbilCommand(aStr)
  Master.Rows = 1
  Label6.Caption = "": Label8.Caption = ""
  Master.ColHidden(0) = True:
  Master.ColAlignment(1) = flexAlignLeftCenter
  Do While Not aRSet.EOF
    If aRSet![Sisa] <> 0 Then
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = 1
    Master.TextMatrix(Master.Rows - 1, 1) = aRSet![No Faktur]
    Master.TextMatrix(Master.Rows - 1, 2) = aRSet![Tgl]
    Master.TextMatrix(Master.Rows - 1, 3) = aRSet![Transaksi]
    Master.TextMatrix(Master.Rows - 1, 4) = aRSet![Total]
    'Master.TextMatrix(Master.Rows - 1, 0) = aRSet![Bayar]
    'Master.TextMatrix(Master.Rows - 1, 0) = aRSet![Sisa]
    Master.TextMatrix(Master.Rows - 1, 7) = aRSet![Kepada] & "-" & aRSet![nAMA]
    Label8.Caption = "Daftar Tanda Terima " & _
    vbCrLf & aRSet![Kepada] & "-" & aRSet![nAMA]
    Label8.Tag = aRSet![Kepada] & ";;;" & aRSet![nAMA] & ";;;"
    Else
    'MsgBox aRSet![No Faktur]
    End If
    aRSet.MoveNext
  Loop
  Master.ColHidden(7) = True
  Label8.Caption = Label8.Caption & vbCrLf & _
  "Dari tanggal " & Format(bs(0), "dd mmm yyyy") & " sampai tanggal " & _
  Format(bs(1), "dd mmm yyyy")
  Label8.Tag = Label8.Tag & bs(0) & ";;;" & bs(1)
  TotalOi
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
BaruBuka = True
If JenisJenis = "TTP" Then
Me.Caption = "Tanda Terima Piutang"
Label2.Caption = "Tanda Terima Piutang"
ElseIf JenisJenis = "TTH" Then
Me.Caption = "Tanda Terima Hutang"
Label2.Caption = "Tanda Terima Hutang"
End If
t2.Value = Date
DoEvents
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
   If Col <> 0 Then Cancel = True
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Master_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
    Select Case Col
    Case 0
      Call TotalOi
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub TotalOi()
On Error GoTo aa
Dim k As Currency, l As Integer
k = 0: l = 0
  Dim i As Integer
  For i = 1 To Master.Rows - 1
  k = k + IIf(Master.TextMatrix(i, 0) = 0, 0, Master.ValueMatrix(i, 4))
  l = l + IIf(Master.TextMatrix(i, 0) = 0, 0, 1)
  Next i
  tot1.Value = k:  Tot2.Value = l
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


Private Sub HapusKosong()
On Error Resume Next
Dim IoTo As Integer: IoTo = 1
 Do While IoTo < Master.Rows
   If Master.TextMatrix(IoTo, 0) = "" Then
     Master.RemoveItem IoTo
   Else
     IoTo = IoTo + 1
   End If
 Loop
End Sub

