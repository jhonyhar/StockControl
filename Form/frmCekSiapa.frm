VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCekSiapa 
   Caption         =   "Cek Input Transaksi"
   ClientHeight    =   9600
   ClientLeft      =   315
   ClientTop       =   1050
   ClientWidth     =   15240
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
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox t1 
      DataField       =   "No Faktur"
      Height          =   345
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   3435
   End
   Begin VB.OptionButton Opt4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pembayaran Piutang"
      Height          =   270
      Left            =   3855
      TabIndex        =   3
      Top             =   1500
      Width           =   1890
   End
   Begin VB.OptionButton Opt3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pembayaran Hutang"
      Height          =   270
      Left            =   3855
      TabIndex        =   2
      Top             =   1215
      Width           =   1935
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Penjualan && Retur Penjualan"
      Height          =   270
      Left            =   825
      TabIndex        =   1
      Top             =   1485
      Width           =   2700
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pembelian && Retur Pembelian"
      Height          =   270
      Left            =   825
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   2700
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6120
      Top             =   9900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load Data"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5265
      TabIndex        =   8
      Top             =   2040
      Width           =   1725
   End
   Begin VB.TextBox txtFields 
      DataField       =   "JenisTrans"
      Height          =   315
      Left            =   540
      TabIndex        =   11
      Top             =   9990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   4980
      Left            =   435
      TabIndex        =   9
      Top             =   3180
      Width           =   13995
      _cx             =   24686
      _cy             =   8784
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCekSiapa.frx":0000
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
      ExplorerBar     =   1
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
   Begin TDBDate6Ctl.TDBDate tgl1 
      Height          =   405
      Left            =   840
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   714
      Calendar        =   "frmCekSiapa.frx":0107
      Caption         =   "frmCekSiapa.frx":023D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCekSiapa.frx":02AB
      Keys            =   "frmCekSiapa.frx":02C9
      Spin            =   "frmCekSiapa.frx":0327
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
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   1080
      Left            =   3000
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1875
      _cx             =   3307
      _cy             =   1905
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
      FormatString    =   $"frmCekSiapa.frx":034F
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
      ExplorerBar     =   0
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
   Begin TDBNumber6Ctl.TDBNumber num1 
      Height          =   405
      Left            =   840
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   714
      Calculator      =   "frmCekSiapa.frx":04B2
      Caption         =   "frmCekSiapa.frx":04D2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCekSiapa.frx":053E
      Keys            =   "frmCekSiapa.frx":055C
      Spin            =   "frmCekSiapa.frx":05A6
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
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   2490
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   635
      Calendar        =   "frmCekSiapa.frx":05CE
      Caption         =   "frmCekSiapa.frx":0704
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCekSiapa.frx":0772
      Keys            =   "frmCekSiapa.frx":0790
      Spin            =   "frmCekSiapa.frx":07EE
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
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&No Bukti :"
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   4
      Top             =   2100
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tgl :"
      Height          =   330
      Index           =   2
      Left            =   525
      TabIndex        =   6
      Top             =   2505
      Width           =   1845
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Jenis Transaksi : "
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
      Height          =   315
      Left            =   660
      TabIndex        =   18
      Top             =   900
      Width           =   1515
   End
   Begin VB.Shape Shape3 
      Height          =   870
      Left            =   525
      Top             =   1035
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Cek Input Transaksi"
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
      TabIndex        =   14
      Top             =   105
      Width           =   5055
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
      TabIndex        =   15
      Top             =   8880
      Width           =   6285
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
      TabIndex        =   13
      Top             =   195
      Width           =   3255
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
      TabIndex        =   12
      Top             =   9495
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   9135
      Left            =   240
      Top             =   255
      Width           =   14400
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   9375
      Left            =   165
      Top             =   105
      Width           =   14655
   End
End
Attribute VB_Name = "frmCekSiapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
On Error GoTo aa
If t1.Text = "" And t2.Text = "" Then
  MsgBox "Harap masukkan kriteria pencarian..", vbInformation, "Cari Data"
  Exit Sub
End If
Dim aStr As String
  If Opt1.Value Then
    aStr = "SELECT TransBeli.[No Faktur] as [No],tgl, TransBeli.Kepada & ' - ' & Supplier.Nama as Nama, " & _
    "iif(TransBeli.JenisTrans='B','Pembelian','Retur Pembelian') as Jenis, TransBeli.nama as [User], TransBeli.waktu as Waktu " & _
    "FROM Supplier INNER JOIN TransBeli ON Supplier.Kode = TransBeli.Kepada"
  ElseIf Opt2.Value Then
    aStr = "SELECT TransJual.[No Faktur] as [No], tgl,TransJual.Kepada & ' - ' & Konsumen.Nama as Nama, " & _
    "iif(TransJual.JenisTrans='J','Penjualan','Retur Penjualan') as Jenis, TransJual.nama as [User], TransJual.waktu as Waktu " & _
    "FROM Konsumen INNER JOIN TransJual ON Konsumen.Kode = TransJual.Kepada"
  ElseIf Opt3.Value Then
    aStr = "SELECT Utang.[KodeBayar] as [No],  tgl,Utang.Kepada & ' - ' & Supplier.Nama as Nama, " & _
    "'Pembayaran Utang' as Jenis, Utang.nama as [User], Utang.waktu as Waktu " & _
    "FROM Supplier INNER JOIN Utang ON Supplier.Kode = Utang.Kepada"
  Else
    aStr = "SELECT Piutang.[KodeBayar] as [No],  tgl,Piutang.Kepada & ' - ' & Konsumen.Nama as Nama, " & _
    "'Pembayaran Piutang' as Jenis, Piutang.nama as [User], Piutang.waktu as Waktu " & _
    "FROM Konsumen INNER JOIN Piutang ON Konsumen.Kode = Piutang.Kepada"
  End If
  
  If t1.Text <> "" Then
    If Opt1.Value Or Opt2.Value Then
      aStr = aStr & " where [No Faktur]='" & AmanOi(t1.Text) & "' "
    Else
      aStr = aStr & " where [KodeBayar]='" & AmanOi(t1.Text) & "' "
    End If
    t1.Text = ""
  ElseIf Not IsNull(t2.Value) Then
    aStr = aStr & " where tgl between #" & Format(t2.Value, "mm/dd/yyyy") & "# and #" & Format(t2.Value, "mm/dd/yyyy") & "# "
    t2.Value = Null
  End If
  
  Dim oRS As New ADODB.Recordset
  Set oRS = aData.AmbilCommand(aStr)
  Master.Rows = 1
  Do While Not oRS.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = oRS!No
    Master.TextMatrix(Master.Rows - 1, 1) = oRS!Tgl
    Master.TextMatrix(Master.Rows - 1, 2) = oRS![nAMA]
    Master.TextMatrix(Master.Rows - 1, 3) = oRS!Jenis
    Master.TextMatrix(Master.Rows - 1, 4) = IIf(IsNull(oRS!User), "", oRS!User)
    Master.TextMatrix(Master.Rows - 1, 5) = IIf(IsNull(oRS!Waktu), "", oRS!Waktu)
    oRS.MoveNext
  Loop
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub



