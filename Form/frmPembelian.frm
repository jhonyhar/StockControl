VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPembelian 
   Caption         =   "Transaksi Pembelian"
   ClientHeight    =   9600
   ClientLeft      =   315
   ClientTop       =   1050
   ClientWidth     =   14430
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
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Histori Transaksi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      TabIndex        =   43
      Top             =   3360
      Width           =   2175
   End
   Begin VSFlex8Ctl.VSFlexGrid VSG1 
      Height          =   3000
      Left            =   8760
      TabIndex        =   40
      Top             =   360
      Width           =   5595
      _cx             =   9869
      _cy             =   5292
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPembelian.frx":0000
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
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Bukan No. Faktur Kita"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      TabIndex        =   39
      Top             =   690
      Width           =   2115
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   10
      Top             =   2385
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6120
      Top             =   9900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export"
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
      Left            =   7410
      TabIndex        =   32
      Top             =   3210
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Import"
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
      Left            =   6090
      TabIndex        =   33
      Top             =   3210
      Width           =   1230
   End
   Begin VB.TextBox t8 
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
      Height          =   1170
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   7065
      Width           =   6660
   End
   Begin VB.CommandButton CmdCekFaktur 
      Caption         =   "&Cek Faktur"
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
      Left            =   6165
      TabIndex        =   22
      Top             =   1170
      Width           =   1230
   End
   Begin VB.CommandButton CmdBaru 
      Caption         =   "Faktur &Baru"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9795
      TabIndex        =   19
      Top             =   8235
      Width           =   1425
   End
   Begin VB.ComboBox t6 
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
      ItemData        =   "frmPembelian.frx":00A8
      Left            =   2565
      List            =   "frmPembelian.frx":00B2
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2745
      Width           =   3435
   End
   Begin VB.TextBox t4 
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
      Left            =   2565
      TabIndex        =   7
      Top             =   1935
      Width           =   3435
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
      Left            =   2565
      TabIndex        =   5
      Top             =   1530
      Width           =   3435
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11325
      TabIndex        =   20
      Top             =   8235
      Width           =   1290
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12660
      TabIndex        =   21
      Top             =   8235
      Width           =   1290
   End
   Begin VB.TextBox t1 
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
      Left            =   2565
      TabIndex        =   1
      Top             =   720
      Width           =   3435
   End
   Begin VB.TextBox txtFields 
      DataField       =   "JenisTrans"
      Height          =   315
      Left            =   540
      TabIndex        =   23
      Top             =   9990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   3000
      Left            =   435
      TabIndex        =   16
      Top             =   3705
      Width           =   13995
      _cx             =   24686
      _cy             =   5292
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
      FormatString    =   $"frmPembelian.frx":00C4
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
   Begin TDBNumber6Ctl.TDBNumber t5 
      Height          =   405
      Left            =   2565
      TabIndex        =   9
      Top             =   2340
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calculator      =   "frmPembelian.frx":0221
      Caption         =   "frmPembelian.frx":0241
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":02AD
      Keys            =   "frmPembelian.frx":02CB
      Spin            =   "frmPembelian.frx":0315
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
      ValueVT         =   99942401
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tot1 
      Height          =   405
      Left            =   11610
      TabIndex        =   34
      Top             =   6720
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmPembelian.frx":033D
      Caption         =   "frmPembelian.frx":035D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":03C9
      Keys            =   "frmPembelian.frx":03E7
      Spin            =   "frmPembelian.frx":0431
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
      ValueVT         =   144572417
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tot2 
      Height          =   405
      Left            =   11610
      TabIndex        =   35
      Top             =   7110
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmPembelian.frx":0459
      Caption         =   "frmPembelian.frx":0479
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":04E5
      Keys            =   "frmPembelian.frx":0503
      Spin            =   "frmPembelian.frx":054D
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
      ValueVT         =   144572417
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tot3 
      Height          =   405
      Left            =   11610
      TabIndex        =   36
      Top             =   7515
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmPembelian.frx":0575
      Caption         =   "frmPembelian.frx":0595
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":0601
      Keys            =   "frmPembelian.frx":061F
      Spin            =   "frmPembelian.frx":0669
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
      ValueVT         =   144572417
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBDate6Ctl.TDBDate t2 
      Height          =   405
      Left            =   2565
      TabIndex        =   3
      Top             =   1125
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   714
      Calendar        =   "frmPembelian.frx":0691
      Caption         =   "frmPembelian.frx":07C7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":0835
      Keys            =   "frmPembelian.frx":0853
      Spin            =   "frmPembelian.frx":08B1
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
   Begin TDBNumber6Ctl.TDBNumber t7 
      Height          =   405
      Left            =   2565
      TabIndex        =   14
      Top             =   3150
      Width           =   2160
      _Version        =   65536
      _ExtentX        =   3810
      _ExtentY        =   714
      Calculator      =   "frmPembelian.frx":08D9
      Caption         =   "frmPembelian.frx":08F9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":0965
      Keys            =   "frmPembelian.frx":0983
      Spin            =   "frmPembelian.frx":09CD
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
      ValueVT         =   22675457
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber TDB1 
      Height          =   315
      Left            =   480
      TabIndex        =   41
      Top             =   8880
      Width           =   1080
      _Version        =   65536
      _ExtentX        =   1905
      _ExtentY        =   556
      Calculator      =   "frmPembelian.frx":09F5
      Caption         =   "frmPembelian.frx":0A15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPembelian.frx":0A81
      Keys            =   "frmPembelian.frx":0A9F
      Spin            =   "frmPembelian.frx":0AE9
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
      ValueVT         =   2088828933
      Value           =   5
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "baris"
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
      Index           =   12
      Left            =   1680
      TabIndex        =   42
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6120
      TabIndex        =   38
      Top             =   1972
      Width           =   2580
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6120
      TabIndex        =   37
      Top             =   1560
      Width           =   2580
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+F4=Close F4=Bantuan F5=FIFO F6=HrgModal"
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
      TabIndex        =   31
      Top             =   8280
      Width           =   6285
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Keterangan : "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   450
      TabIndex        =   17
      Top             =   6735
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "hari"
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
      Index           =   10
      Left            =   4800
      TabIndex        =   15
      Top             =   3225
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Lama Kredit : "
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
      Index           =   9
      Left            =   525
      TabIndex        =   13
      Top             =   3225
      Width           =   1815
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
      Left            =   9750
      TabIndex        =   30
      Top             =   6795
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Diskon : "
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
      Index           =   7
      Left            =   9750
      TabIndex        =   29
      Top             =   7185
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total : "
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
      Index           =   6
      Left            =   9750
      TabIndex        =   28
      Top             =   7590
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Diskon:"
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
      Index           =   5
      Left            =   525
      TabIndex        =   8
      Top             =   2415
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Salesman:"
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
      Index           =   4
      Left            =   525
      TabIndex        =   6
      Top             =   2010
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dari:"
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
      Index           =   3
      Left            =   525
      TabIndex        =   4
      Top             =   1605
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
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&No Faktur:"
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
      Left            =   525
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Pembayaran:"
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
      Index           =   0
      Left            =   525
      TabIndex        =   11
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pembelian"
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
      TabIndex        =   26
      Top             =   105
      Width           =   3255
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
      TabIndex        =   27
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   9495
      Width           =   1665
   End
End
Attribute VB_Name = "frmPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**0=Kode 1=nama  2=hrg 3=qtyKoktak 4=qty 5=Satuan
'**6=Disc  7=% 8=sub 9=qtyperkarton 10=RealDisc
'0=kode 1=nama 2=qtyB 3=satB 4=qtyS 5=satS 6=harga 7=disc 8=%
'9=sub 10=qtyperkarton 11=RealDisc
Option Explicit
Dim BaruBuka As Boolean
Public JenisJenis As String
Dim edRow As Integer, edCol  As Integer
Dim intNoFaktur As Single
Dim PtoScreen As Boolean

Private Sub Check1_Click()
Call TotalOi
End Sub

Private Sub CmdBaru_Click()
On Error GoTo aa
If BaruBuka Then
  BaruBuka = False
Else
  If MsgBox("Anda yakin ingin melanjutkan dengan transaksi baru ..?", vbQuestion + vbYesNo + vbDefaultButton2, "Transaksi Baru") = vbNo Then
    Exit Sub
  End If
End If
Master.MergeRow(0) = True
Check2.Visible = False
If JenisJenis = "B" Then
t1 = SerbaGuna.NoFaktur("B", "TransBeli", "B")
lblLabels(3).Caption = "Dari :"
lblLabels(4).Caption = "No Surat Jalan :"
ElseIf JenisJenis = "RB" Then
t1 = SerbaGuna.NoFaktur("RB", "TransBeli", "RB")
lblLabels(3).Caption = "Kepada :"
ElseIf JenisJenis = "J" Then
t1 = SerbaGuna.NoFaktur("J", "TransJual", "J")
lblLabels(3).Caption = "Kepada :"
Check2.Visible = True
Check2.Value = vbUnchecked
ElseIf JenisJenis = "RJ" Then
t1 = SerbaGuna.NoFaktur("RJ", "TransJual", "RJ")
lblLabels(3).Caption = "Dari :"
End If
t2.Value = Date
Master.Rows = 1
Master.Rows = 2
Master.Row = 1
Master.Col = 0
t3 = ""
t4 = ""
t5.Value = 0
t6.ListIndex = 1
intNoFaktur = 0
Check1.Value = Unchecked
t7.Value = 0
t8 = ""
tot1.Value = 0
tot2.Value = 0
tot3.Value = 0
Label6.Caption = ""
Label7.Caption = ""
PtoScreen = False
LagiSimpan = False
On Error Resume Next
t1.SetFocus
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdCekFaktur_Click()
  On Error GoTo aa
'(0) = NoFaktur  (1) = Tgl  (2) = Supplier  (3) = Salesman
'(4) = Diskon    (5) = CaraPembayaran c/k
'(6) = LamaKredit (7) = jenistrans b/rb  (8) = Keterangan

Check1.Value = Unchecked
Dim rsMaster As New ADODB.Recordset
If JenisJenis = "B" Or JenisJenis = "RB" Then
Set rsMaster = aData.AmbilCommand("select * from transbeli where jenistrans='" & _
JenisJenis & "' and [No Faktur]='" & t1.Text & "' and hapus=0 and (jenis='K' or jenis='C')")
ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
Set rsMaster = aData.AmbilCommand("select * from transjual where jenistrans='" & _
JenisJenis & "' and [No Faktur]='" & t1.Text & "' and hapus=0 and (jenis='K' or jenis='C')")
End If
  If Not rsMaster.EOF Then
  t1.Text = rsMaster![No Faktur]
  t2.Value = rsMaster!Tgl
  t3.Text = rsMaster!Kepada
  Call t3_Validate(False)
  t4.Text = rsMaster!Salesman
  Call t4_Validate(False)
  t5.Value = rsMaster!Diskon
  t6.Text = IIf(rsMaster!Jenis = "K", "Kredit", "Cash")
  t7.Value = rsMaster!jatuhtempo
  t8.Text = rsMaster!Keterangan
  intNoFaktur = rsMaster!intNo
    If JenisJenis = "B" Or JenisJenis = "RB" Then
    Set rsMaster = aData.AmbilCommand("select detailbeli.*,barang.[Nama Barang],barang.[Qty Satuan Kecil],barang.[Satuan Kecil],Barang.Satuan " & _
    "from detailbeli,barang where detailbeli.[Kode Barang]=barang.[Kode Barang] and " & _
    "inttrans=" & intNoFaktur)
    ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
    Set rsMaster = aData.AmbilCommand("select detailjual.*,barang.[Nama Barang],barang.[Qty Satuan Kecil],barang.[Satuan Kecil], Barang.Satuan " & _
    "from detailjual,barang where detailjual.[Kode Barang]=barang.[Kode Barang] and " & _
    "inttrans=" & intNoFaktur)
    End If
    Master.Rows = 1
    Do While Not rsMaster.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = rsMaster![Kode Barang]
    Master.TextMatrix(Master.Rows - 1, 1) = rsMaster![Nama Barang]
    Master.TextMatrix(Master.Rows - 1, 2) = rsMaster![QtyB]
    Master.TextMatrix(Master.Rows - 1, 3) = rsMaster![Satuan]
    Master.TextMatrix(Master.Rows - 1, 4) = rsMaster![QtyS]
    Master.TextMatrix(Master.Rows - 1, 5) = rsMaster![Satuan Kecil]
    Master.TextMatrix(Master.Rows - 1, 6) = rsMaster!Harga
    Master.TextMatrix(Master.Rows - 1, 7) = rsMaster!Diskon
    Master.TextMatrix(Master.Rows - 1, 8) = False
    Master.TextMatrix(Master.Rows - 1, 10) = rsMaster!Pengali
    Master.TextMatrix(Master.Rows - 1, 11) = rsMaster!Diskon
    Call Master_AfterEdit(Master.Rows - 1, 2)
    rsMaster.MoveNext
    Loop
    HapusKosong
    Call TotalOi
  Else
  MsgBox "No Faktur " & t1 & vbCrLf & "tidak ditemukan pada database atau telah dihapus", vbInformation, "Cek Faktur"
  Call CmdBaru_Click
  End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Function CekSimpanJual(aO() As String) As Boolean
On Error Resume Next
Dim i As Integer, aRSet As ADODB.Recordset, aHsl As String
If t6.ListIndex <> 0 Then
  Set aRSet = aData.AmbilCommand("CreditCheck '" & SerbaGuna.AmanOi(t3.Text) & "'")
  If aRSet.EOF Then Set aRSet = aData.AmbilCommand("select [Credit Limit],'0' as Sisa from Konsumen where kode='" & SerbaGuna.AmanOi(t3.Text) & "'")
  If Not aRSet.EOF Then
    If (aRSet![Credit Limit] < aRSet!Sisa + tot1.Value - tot2.Value) Then
    aHsl = "Konsumen yang bertransaksi (" & t3.Text & ") melewati batas kredit limit.." & vbCrLf & _
    "(" & aRSet![Credit Limit] & ";" & aRSet!Sisa & ";" & tot1.Value - tot2.Value & ")" & vbCrLf & vbCrLf
    End If
  End If
End If
For i = LBound(aO, 1) To UBound(aO, 1)
If aO(i, 0) <> "" Then
  Set aRSet = aData.AmbilCommand("select Qty from Barang where [Kode Barang]='" & AmanOi(aO(i, 0)) & "'")
  If aRSet.RecordCount = 1 And aRSet(0) < Val(aO(i, 2)) Then
    aHsl = aHsl & "Barang dengan kode " & aO(i, 0) & " memiliki qty yang kurang " & vbCrLf
  End If
End If
Next i
If aHsl <> "" Then
  If MsgBox(aHsl & vbCrLf & "Apakah anda akan melanjutkan transaksi ini..?", vbQuestion + vbYesNo + vbDefaultButton2, "Simpan Transaksi") = vbNo Then
    CekSimpanJual = False
  Else
    CekSimpanJual = True
  End If
Else
  CekSimpanJual = True
End If
End Function


Private Sub CmdSimpan_Click()
On Error GoTo aa
Dim aHead(11) As String, aBodi() As String, i As Byte, aHasil As String
Call TotalOi
ReDim aBodi(Master.Rows - 1, 7)
aHead(0) = SerbaGuna.AmanOi(t1.Text) 'NoFaktur
aHead(1) = SerbaGuna.AmanTgl(t2.Value)  'Tgl
aHead(2) = SerbaGuna.AmanOi(t3.Text) 'Supplier
aHead(3) = SerbaGuna.AmanOi(t4.Text) 'Salesman
aHead(4) = tot2.Value  'Diskon
aHead(5) = IIf(t6.ListIndex = 0, "C", "K") 'CaraPembayaran
aHead(6) = Val(SerbaGuna.AmanOi(t7.Value))  'LamaKredit
aHead(7) = SerbaGuna.AmanOi(JenisJenis) 'JenisTrans
aHead(8) = SerbaGuna.AmanOi(t8.Text) 'Keterangan
aHead(9) = tot1.Value  'Total
aHead(10) = IIf(Check2.Value = vbChecked, True, False)
For i = 1 To Master.Rows - 1
If Master.TextMatrix(i, 0) <> "" Then
aBodi(i, 0) = SerbaGuna.AmanOi(Master.TextMatrix(i, 0)) 'Kode
aBodi(i, 1) = Master.ValueMatrix(i, 6)  'Harga
aBodi(i, 2) = Master.ValueMatrix(i, 2) + (Master.ValueMatrix(i, 4) / Master.ValueMatrix(i, 10)) 'Qty
aBodi(i, 3) = Master.ValueMatrix(i, 11) 'Diskon
aBodi(i, 4) = Master.ValueMatrix(i, 10) 'Pengali
aBodi(i, 5) = Master.ValueMatrix(i, 2) 'QtyB
aBodi(i, 6) = Master.ValueMatrix(i, 4) 'QtyS
End If
Next i

If JenisJenis = "J" Then
  If Not CekSimpanJual(aBodi) Then Exit Sub
End If
If JenisJenis = "B" Or JenisJenis = "RB" Then
aHasil = aData.SimpanBeli(aHead, aBodi)
ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
aHasil = aData.SimpanJual(aHead, aBodi)
End If
If aHasil <> "" Then
MsgBox aHasil, vbInformation, Me.Caption & "#Error"
Else
'  If JenisJenis = "B" Then
'  Dim DiiS As Currency
    
'    On Error Resume Next
'    For i = 1 To Master.Rows - 1
'    If Master.TextMatrix(i, 0) <> "" Then
'      If Master.ValueMatrix(i, 7) Then
'      DiiS = Master.ValueMatrix(i, 6)
'      Else
'      DiiS = Val(FormatNumber(Master.ValueMatrix(i, 6) / Master.ValueMatrix(i, 2) * 100, 2))
'      End If
'    aData.ExecCommand "update Barang set [Harga Beli]='" & Master.ValueMatrix(i, 2) & _
'    "', [Disc Beli]='" & DiiS & _
'    "' where [Kode Barang]='" & SerbaGuna.AmanOi(Master.TextMatrix(i, 0)) & "'"
'    End If
'    Next i
'    On Error GoTo aa
'  End If
If MsgBox("Cetak Faktur..?", vbInformation + vbYesNo, "Faktur") = vbYes Then
Set fRpt.Report = Fakktur 'Lap.GetFaktur
If JenisJenis = "B" Or JenisJenis = "RB" Then
  Call AturLaporan("FakturPrint", "B", t1.Text, Not PtoScreen)
ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
  Call AturLaporan("FakturPrint", "J", t1.Text, Not PtoScreen)
End If
PtoScreen = False
End If
BaruBuka = True
Call CmdBaru_Click
End If

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub CmdHapus_Click()
On Error GoTo aa
If MsgBox("Anda yakin ingin melanjutkan proses ini..?", vbQuestion + vbYesNo + vbDefaultButton2, "Hapus Transaksi") = vbNo Then
 Exit Sub
End If
Dim aaa As String
If JenisJenis = "B" Or JenisJenis = "RB" Then
aaa = aData.HapusBeli(CStr(intNoFaktur), JenisJenis)
ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
aaa = aData.HapusJual(CStr(intNoFaktur), JenisJenis)
End If
If aaa <> "" Then
MsgBox aaa, vbInformation, Me.Caption
Else
  CmdBaru_Click
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = vbShiftMask Then PtoScreen = True
End Sub

Private Sub Command1_Click()
Dim CD2 As New myComDialog.cComDialog
On Error GoTo aa
    CD2.InitDir = App.Path '& IIf(Right(App.Path, 1) = "\", "Export", "\Export")
    CD2.FileName = ""
    On Error Resume Next
    CD2.CancelError = True
    CD2.DialogTitle = "Masukkan nama file Transaksi yang akan diimport.."
    CD2.Filter = "Transaksi File|*.trns"
    If JenisJenis = "B" Then
    CD2.Flags = cdlOFNFileMustExist Or OFN_PATHMUSTEXIST Or cdlOFNAllowMultiselect
    Else
    CD2.Flags = cdlOFNFileMustExist Or OFN_PATHMUSTEXIST
    End If
    CD2.ShowOpen
    Dim zk() As String, i As Long, xPath As String
    
    If Err.Number = 0 Then
        Master.Rows = 1
        t4.Text = ""
    zk = Split(CD2.FileName, vbNullChar)
    If UBound(zk) = 0 Then
      xPath = Left(zk(0), Len(zk(0)) - Len(CD2.FileTitle))
      ReDim zk(0 To 1)
      zk(0) = xPath: zk(1) = CD2.FileTitle
    Else
      xPath = IIf(Right(zk(0), 1) <> "\", zk(0) & "\", zk(0))
    End If
    For i = 1 To UBound(zk)
    On Error GoTo aa
    Dim k As New ADODB.Recordset
    If k.State = adStateOpen Then k.Close
    k.Open IIf(zk(1) = "", zk(0), xPath & zk(i))
    k.MoveFirst
    
    If JenisJenis <> k!Diskon Then
      MsgBox "Jenis transaksi yang berlainan tidak dapat diproses..", vbInformation, "Import Data"
      Exit Sub
    End If
    On Error Resume Next
    t1.Text = k!Kode
    On Error GoTo aa
    t2.Value = k!nAMA
    t3.Text = k!QtyB
    t4.Text = k!SatB 't4.Text & IIf(t4.Text = "", k!OO, "," & k!OO)
    t5.Value = Val(k!QtyS)
    t6.Text = k!SatS
    t7.Value = Val(k!Harga)
    t8.Text = k!Karton
    k.MoveNext
     Do While Not k.EOF
     Master.Rows = Master.Rows + 1
      Master.TextMatrix(Master.Rows - 1, 0) = k!Kode
     Master.TextMatrix(Master.Rows - 1, 1) = k!nAMA
     Master.TextMatrix(Master.Rows - 1, 2) = Val(k!QtyB)
     Master.TextMatrix(Master.Rows - 1, 3) = k!SatB
     Master.TextMatrix(Master.Rows - 1, 4) = Val(k!QtyS)
     Master.TextMatrix(Master.Rows - 1, 5) = k!SatS
     Master.TextMatrix(Master.Rows - 1, 6) = Val(k!Harga)
     Master.TextMatrix(Master.Rows - 1, 7) = Val(k!Diskon)
     Master.TextMatrix(Master.Rows - 1, 8) = False
     Master.TextMatrix(Master.Rows - 1, 10) = Val(k!Karton)
     Master.TextMatrix(Master.Rows - 1, 11) = Val(k!Diskon)
     Call Master_AfterEdit(Master.Rows - 1, 2)
     k.MoveNext
     Loop
     HapusKosong
     On Error Resume Next
    If zk(i + 1) = "" Then i = UBound(zk)
     On Error GoTo aa
    Next i
     Call TotalOi
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
    CD1.FileName = t1.Text
    CD1.ShowSave
    If Err.Number = 20477 Then ' error file name
      CD1.FileName = "":     Err.Clear:     CD1.ShowSave
    End If
    If Err.Number = 0 Then
    On Error GoTo aa
    Dim k As New ADODB.Recordset, i As Integer
    '(0)=Tgl (1)= Supplier  (2)=Salesman (3)=Diskon (4)=CaraPembayaran c/k
    '(5)=LamaKredit (6)=jenistrans b/rb  (7)=Keterangan
    k.Fields.Append "Kode", adVarChar, 50
    k.Fields.Append "Nama", adVarChar, 200
    k.Fields.Append "QtyB", adVarChar, 50
    k.Fields.Append "SatB", adVarChar, 50
    k.Fields.Append "QtyS", adVarChar, 50
    k.Fields.Append "SatS", adVarChar, 50
    k.Fields.Append "Harga", adVarChar, 50
    k.Fields.Append "Diskon", adVarChar, 50
    k.Fields.Append "Karton", adVarChar, 200
    
    k.Open
    k.AddNew:
    k!Kode = t1.Text
    k!nAMA = Str(t2.Value)
    k!QtyB = t3.Text
    k!SatB = t4.Text
    k!QtyS = tot2.Value
    k!SatS = t6.Text
    k!Harga = t7.Value
    k!Diskon = JenisJenis
    k!Karton = t8.Text
    k.Update
     For i = 1 To Master.Rows - 1
     If Master.TextMatrix(i, 1) <> "" Then
       k.AddNew
       k!Kode = Master.TextMatrix(i, 0)
       k!nAMA = Master.TextMatrix(i, 1)
       k!QtyB = Master.ValueMatrix(i, 2)
       k!SatB = Master.TextMatrix(i, 3)
       k!QtyS = Master.ValueMatrix(i, 4)
       k!SatS = Master.TextMatrix(i, 5)
       k!Harga = Master.ValueMatrix(i, 6)
       k!Diskon = Master.ValueMatrix(i, 11)
       k!Karton = Master.ValueMatrix(i, 10)
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
BaruBuka = True
If JenisJenis = "B" Then
Me.Caption = "Transaksi Pembelian"
Label2.Caption = "Pembelian"
ElseIf JenisJenis = "RB" Then
Me.Caption = "Transaksi Retur Pembelian"
Label2.Caption = "Retur Pembelian"
ElseIf JenisJenis = "J" Then
Me.Caption = "Transaksi Penjualan"
Label2.Caption = "Penjualan"
ElseIf JenisJenis = "RJ" Then
Me.Caption = "Transaksi Retur Penjualan"
Label2.Caption = "Retur Penjualan"
End If
DoEvents
Call CmdBaru_Click
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub


Private Sub Master_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
  Dim rsAA As New ADODB.Recordset
  If Not (Check3.Value = vbChecked) Then Exit Sub
  If Master.TextMatrix(NewRow, 0) = "" Then Exit Sub
  If OldRow = NewRow Then Exit Sub
Dim kStr As String
Dim Top As Integer
Top = 5: If TDB1.Value > 5 Then Top = TDB1.Value
If JenisJenis = "B" Or JenisJenis = "RB" Then
kStr = "SELECT top " & Top & " TransBeli.Tgl, TransBeli.[No Faktur] AS [No], DetailBeli.QtyB & ' ' " & _
     "& Barang.Satuan & ' ' & DetailBeli.QtyS & ' ' & Barang.[Satuan Kecil] " & _
     "AS Qty, DetailBeli.Harga, DetailBeli.Diskon FROM TransBeli INNER JOIN  " & _
     "(Barang INNER JOIN DetailBeli ON Barang.[Kode Barang] = DetailBeli.[Kode Barang]) " & _
     "ON TransBeli.intno = DetailBeli.intTrans where TransBeli.JenisTrans='B' and TransBeli.Kepada='" & _
     AmanOi(t3.Text) & "' and DetailBeli.[Kode Barang]='" & _
     AmanOi(Master.TextMatrix(NewRow, 0)) & "' order by TransBeli.Tgl Desc, DetailBeli.intNo desc"
Else
kStr = "SELECT top " & Top & " TransJual.Tgl, TransJual.[No Faktur] AS [No], DetailJual.QtyB & ' ' " & _
     "& Barang.Satuan & ' ' & DetailJual.QtyS & ' ' & Barang.[Satuan Kecil] " & _
     "AS Qty, DetailJual.Harga, DetailJual.Diskon FROM TransJual INNER JOIN  " & _
     "(Barang INNER JOIN DetailJual ON Barang.[Kode Barang] = DetailJual.[Kode Barang]) " & _
     "ON TransJual.intno = DetailJual.intTrans where TransJual.JenisTrans='J' and TransJual.Kepada='" & _
     AmanOi(t3.Text) & "' and DetailJual.[Kode Barang]='" & _
     AmanOi(Master.TextMatrix(NewRow, 0)) & "' order by TransJual.Tgl Desc, DetailJual.intNo desc"
End If

  Set rsAA = aData.AmbilCommand(kStr)
  Set VSG1.DataSource = rsAA
End Sub

Private Sub Master_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
On Error Resume Next
If KeyCode = vbKeyInsert Then
Master.Rows = Master.Rows + 1
Master.Row = Master.Rows - 1
Master.Col = 0
End If
If KeyCode = vbKeyDelete And Master.Rows <> 1 Then
Master.RemoveItem Master.Row
Call TotalOi
End If
If KeyCode = vbKeyF5 And Master.Rows <> 0 Then
      Set aBantu.VGrid.DataSource = aData.QtyModal(Master.TextMatrix(Master.Row, 0))
      aBantu.Show vbModal
End If
If KeyCode = vbKeyF6 And Master.Rows <> 0 Then
   With Master
    MsgBox "Harga Modal " & .TextMatrix(.Row, 1) & _
    " = " & aData.QtyRata(.TextMatrix(Master.Row, 0), .ValueMatrix(.Row, 2) + (.ValueMatrix(.Row, 4) / .ValueMatrix(.Row, 10))), vbInformation, "Modal FIFO"
   End With
End If
If KeyCode = vbKeyReturn And Master.Col = 0 Then
     KeyCode = 0
     BantuBarang.Show vbModal
          KeyCode = 0
     Dim iKi As Integer
     If Not BantuBarang.Batal Then
     KeyCode = 0
     Master.TextMatrix(Master.Row, 0) = BarisGrid(0)
     Call Brg2(Master.Row, 0)
     End If
End If
If KeyCode = vbKeyF4 And Master.Col = 8 Then  'multi Diskon
     KeyCode = 0
     fDiskon.Show vbModal
     KeyCode = 0
     If Not fDiskon.Batal Then
     KeyCode = 0
     Master.TextMatrix(Master.Row, 7) = fDiskon.NilaiDiskon
     Master.TextMatrix(Master.Row, 8) = 1
     Call Master_AfterEdit(Master.Row, 2)
     End If
End If
Master.SetFocus
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Master_KeyDown pada Form frmPembelian"
End Sub

Private Sub Master_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
    Select Case Col
    Case 2, 3, 4 ', 5
     edRow = Master.Row
     edCol = Master.Col
     Bantu.Apaan = "Angka"
     Bantu.Left = Master.CellLeft + Master.Left + 50
     Bantu.Top = Master.CellTop + Master.Top + Master.CellHeight + 400
     Bantu.Show vbModal
     Master.Text = Bantu.NilaiBantu
     Call Master_AfterEdit(Row, Col)
     Call TotalOi
    End Select
End If

'If Master.Col = 0 And KeyCode = vbKeyReturn Then
'     Cancel = True
'     edRow = Master.Row
'     edCol = Master.Col
'     KeyCode = 0
'     BantuBarang.Show vbModal
'     If Not BantuBarang.Batal Then
'     Master.TextMatrix(edRow, 0) = BarisGrid(0)
'     End If
'End If




Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
    Select Case Col
    Case 0
      If Screen.ActiveControl.Parent.Caption = "Barang" Then Cancel = True
    Case 1, 3, 5, 9
    Cancel = True
    
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


Private Sub Master_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
    Select Case Col
    Case 0
      If Master.TextMatrix(Row, Col) <> "" Then Call Brg2(Row, Col)
    Case 41
      If JenisJenis = "RJ" Then
        Dim KRtr As New ADODB.Recordset
        Set KRtr = aData.AmbilCommand("SELECT TransJual.[No Faktur], [Kepada] & '-' & [Nama] AS Konsumen, DetailJual.[Kode Barang], " & _
        "Barang.[Nama Barang], DetailJual.Qty, ([DetailJual.Harga]-[DetailJual.Diskon]) AS Harga, " & _
        "DetailJual.HRata as Modal FROM Konsumen INNER JOIN (TransJual INNER JOIN " & _
        "(Barang INNER JOIN DetailJual ON Barang.[Kode Barang] = DetailJual.[Kode Barang]) " & _
        "ON TransJual.intno = DetailJual.intTrans) ON Konsumen.Kode = TransJual.Kepada " & _
        "Where Kepada='" & AmanOi(t3.Text) & "' and barang.[kode barang]='" & _
        AmanOi(Master.TextMatrix(Row, 0)) & "' and transjual.hapus=0 and transjual.jenistrans='J'")
        Set bBantu.RSSData = KRtr
        bBantu.Show vbModal
        If bBantu.Batal Then
          Master.TextMatrix(Row, 4) = 0
        Else
          If BarisGrid(1) <> Master.TextMatrix(Row, 4) Then
          MsgBox "Error:" & vbCrLf & "Qty barang " & Master.TextMatrix(Row, 1) & " yg diretur adalah " & _
          Master.TextMatrix(Row, 4) & " sedangkan qty perhitungan modal adalah " & BarisGrid(1), vbInformation, "Retur Penjualan"
          Master.TextMatrix(Row, 4) = 0
          Else
          MsgBox "Harga Modal utk barang " & Master.TextMatrix(Row, 1) & " adalah " & BarisGrid(0), vbInformation, "Retur Penjualan"
          Master.TextMatrix(Row, 3) = BarisGrid(0)
          If Master.ValueMatrix(Row, 2) = 0 Then Master.TextMatrix(Row, 2) = BarisGrid(2)
          End If
        End If
      End If
      Call Master_AfterEdit(Row, 3)
    Case 2, 4, 6, 7, 8
      If (Master.TextMatrix(Row, 8)) Then
        Master.TextMatrix(Row, 11) = (Master.ValueMatrix(Row, 7) / 100) * Master.ValueMatrix(Row, 6)
      Else
        Master.TextMatrix(Row, 11) = Master.ValueMatrix(Row, 7)
      End If
      Master.TextMatrix(Row, 9) = FormatNumber((Master.ValueMatrix(Row, 6) - _
      Master.ValueMatrix(Row, 11)) * _
      (Master.ValueMatrix(Row, 2) + (Master.ValueMatrix(Row, 4) / Master.ValueMatrix(Row, 10))) _
      , 2)
      Call TotalOi
      If Row = Master.Rows - 1 Then Master.Rows = Master.Rows + 1
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub TotalOi()
On Error GoTo aa
Dim k As Currency
k = 0
  Dim i As Byte
  For i = 1 To Master.Rows - 1
  k = k + IIf(Master.TextMatrix(i, 9) = "", 0, Master.TextMatrix(i, 9))
  Next i
  tot1.Value = k
        Dim nK As Currency
      If Check1.Value = Checked Then
        nK = (t5.Value / 100) * k
      Else
        nK = t5.Value
      End If
  tot2.Value = nK
  tot3.Value = tot1.Value - tot2.Value
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

'Private Sub t1_GotFocus()
'   On Error GoTo t1_GotFocus_Error
'If BaruBuka Then
't3.SetFocus
'BaruBuka = False
'End If
'   On Error GoTo 0
'   Exit Sub
't1_GotFocus_Error:
'  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure t1_GotFocus of Form frmPembelian"
'End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     If JenisJenis = "B" Or JenisJenis = "RB" Then
     'Bantu.GridData = "SELECT Supplier.Kode, Supplier.Nama, Supplier.Alamat, Supplier.Kota, Supplier.Telepon, Supplier.Diskon FROM Supplier order by Nama;"
     Bantu.GridData = "SELECT Supplier.Kode, Supplier.Nama, Supplier.Alamat,Supplier.Wilayah FROM Supplier order by Nama;"
     ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
     'Bantu.GridData = "SELECT Kode, Nama, Alamat, Kota, Telepon, Diskon FROM Konsumen order by Nama"
     Bantu.GridData = "SELECT Kode, Nama, Alamat,Wilayah FROM Konsumen order by Nama"
     End If
     Bantu.Left = t3.Left + Me.Left + 50
     Bantu.Top = t3.Top + Me.Top + t3.Height + 500
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     t3.Text = BarisGrid(0)
     End If
     t5.Text = 0 'BarisGrid(5) & "%"
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t3_Validate(Cancel As Boolean)
'Tampilkan Salesman dari Konsumen
On Error GoTo aa
Dim rsKodok As New ADODB.Recordset
If JenisJenis = "J" Or JenisJenis = "RJ" Then
Set rsKodok = aData.AmbilCommand("Select * from Konsumen where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
  Label6.Caption = IIf(IsNull(rsKodok!nAMA), "", rsKodok!nAMA) & "(" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & ")"
  t4.Text = IIf(IsNull(rsKodok!Salesman), "", rsKodok!Salesman)
  t7.Value = IIf(IsNull(rsKodok![Lama Credit]), 0, rsKodok![Lama Credit])
  Call t4_Validate(False)
  End If
Else
  Set rsKodok = aData.AmbilCommand("Select * from Supplier where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
  t7.Value = IIf(IsNull(rsKodok![Lama Credit]), 0, rsKodok![Lama Credit])
  Label6.Caption = IIf(IsNull(rsKodok!nAMA), "", rsKodok!nAMA) & "(" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & ")"
  End If
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur t3_Validate pada Form frmPembelian"
End Sub

Private Sub t4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If JenisJenis = "B" Then Exit Sub
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     Bantu.GridData = "Salesman"
     Bantu.Left = t4.Left + Me.Left + 50
     Bantu.Top = t4.Top + Me.Top + t4.Height + 500
     Bantu.Show vbModal
     If Not (Bantu.Batal) Then
     If JenisJenis = "B" Or JenisJenis = "RB" Then
     t4.Text = BarisGrid(0) & "-" & BarisGrid(1)
     ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
     t4.Text = BarisGrid(0)
     End If
     End If
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub t4_Validate(Cancel As Boolean)
On Error GoTo aa
Dim rsKodok As New ADODB.Recordset
Label7.Caption = ""
If JenisJenis = "J" Or JenisJenis = "RJ" Then
Set rsKodok = aData.AmbilCommand("Select * from Salesman where kode='" & AmanOi(t4.Text) & "'")
  If Not rsKodok.EOF Then
  Label7.Caption = rsKodok!nAMA
  End If
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur t4_Validate pada Form frmPembelian"
End Sub

Private Sub t5_Validate(Cancel As Boolean)
Call TotalOi
End Sub


Private Sub Brg2(Row As Long, Col As Long)
Dim rsSbrg As New ADODB.Recordset
On Error GoTo aa
Set rsSbrg = aData.AmbilCommand("SELECT * FROM Barang where Barang.[Kode Barang]='" & AmanOi(Master.TextMatrix(Row, Col)) & "';")
If rsSbrg.EOF Then
  MsgBox "Kode barang tersebut tidak ada, harap cek kembali kode barang anda", vbInformation, "Kode Barang"
  Master.TextMatrix(Row, Col) = ""
  Exit Sub
Else
    '***********0=Kode 1=nama  2=hrg 3=qtyKoktak 4=qty 5=Satuan 6=Disc  7=% 8=sub 9=qtyperkarton 10=RealDisc
    '0=kode 1=nama 2=qtyB 3=satB 4=qtyS 5=satS 6=harga 7=disc 8=%
    '9=sub 10=qtyperkarton 11=RealDisc
  Master.TextMatrix(Row, 0) = rsSbrg![Kode Barang]
  Master.TextMatrix(Row, 1) = rsSbrg![Nama Barang]
  Master.TextMatrix(Row, 2) = 0
  Master.TextMatrix(Row, 3) = rsSbrg![Satuan]
  Master.TextMatrix(Row, 4) = 0
  Master.TextMatrix(Row, 5) = rsSbrg![Satuan Kecil]
    
    If JenisJenis = "B" Or JenisJenis = "RB" Then
    Master.TextMatrix(Row, 6) = rsSbrg![Harga Beli]
    ElseIf JenisJenis = "J" Or JenisJenis = "RJ" Then
    Master.TextMatrix(Row, 6) = rsSbrg![Harga Jual]
    End If
  
  Master.TextMatrix(Row, 7) = 0
  Master.TextMatrix(Row, 8) = False
  
  Master.TextMatrix(Row, 9) = 0
  Master.TextMatrix(Row, 10) = rsSbrg![Qty Satuan Kecil]
  Master.TextMatrix(Row, 11) = 0
  Call TotalOi
  Master.Col = 2
  Call Master_AfterRowColChange(0, ByVal 2, ByVal Row, 2)
  End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Brg2 pada Form frmPembelian"
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
