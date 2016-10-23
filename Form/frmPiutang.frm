VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmPiutang 
   Caption         =   "Pembayaran Piutang"
   ClientHeight    =   9660
   ClientLeft      =   750
   ClientTop       =   645
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   14055
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&Hapus"
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
      Left            =   6660
      TabIndex        =   23
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CmdCekFaktur 
      Caption         =   "&Cek Piutang"
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
      Left            =   4905
      TabIndex        =   15
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CmdBaru 
      Caption         =   "&Pembayaran Baru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10575
      TabIndex        =   14
      Top             =   8235
      Width           =   2385
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
      Left            =   2475
      TabIndex        =   6
      Top             =   1530
      Width           =   2175
   End
   Begin VB.TextBox t2o 
      DataField       =   "No Faktur"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   10395
      Visible         =   0   'False
      Width           =   2175
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
      Height          =   420
      Left            =   10575
      TabIndex        =   13
      Top             =   7755
      Width           =   2385
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
      Left            =   2475
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   2265
      Left            =   345
      TabIndex        =   10
      Top             =   4965
      Width           =   12585
      _cx             =   22199
      _cy             =   3995
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
      AllowUserResizing=   1
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
      FormatString    =   $"frmPiutang.frx":0000
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
      Height          =   810
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   7590
      Width           =   6660
   End
   Begin VSFlex8Ctl.VSFlexGrid Utang 
      Height          =   2265
      Left            =   345
      TabIndex        =   8
      Top             =   2025
      Width           =   12585
      _cx             =   22199
      _cy             =   3995
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPiutang.frx":0137
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
   Begin TDBDate6Ctl.TDBDate T2 
      Height          =   405
      Left            =   2475
      TabIndex        =   3
      Top             =   1125
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   714
      Calendar        =   "frmPiutang.frx":0231
      Caption         =   "frmPiutang.frx":0367
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPiutang.frx":03D5
      Keys            =   "frmPiutang.frx":03F3
      Spin            =   "frmPiutang.frx":0451
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
      Text            =   "19/04/2007"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39191
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber TotBayar 
      Height          =   405
      Left            =   10470
      TabIndex        =   20
      Top             =   4305
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmPiutang.frx":0479
      Caption         =   "frmPiutang.frx":0499
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPiutang.frx":0505
      Keys            =   "frmPiutang.frx":0523
      Spin            =   "frmPiutang.frx":056D
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   2
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "Rp ##,###,###,##0.##"
      EditMode        =   0
      Enabled         =   0
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
      ValueVT         =   148439045
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber TotFaktur 
      Height          =   405
      Left            =   10515
      TabIndex        =   21
      Top             =   7245
      Width           =   2430
      _Version        =   65536
      _ExtentX        =   4286
      _ExtentY        =   714
      Calculator      =   "frmPiutang.frx":0595
      Caption         =   "frmPiutang.frx":05B5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPiutang.frx":0621
      Keys            =   "frmPiutang.frx":063F
      Spin            =   "frmPiutang.frx":0689
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   2
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "Rp ##,###,###,##0.##"
      EditMode        =   0
      Enabled         =   0
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
      ValueVT         =   148439045
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VSFlex8Ctl.VSFlexGrid VGrid 
      Height          =   1290
      Left            =   10485
      TabIndex        =   24
      Top             =   315
      Visible         =   0   'False
      Width           =   2460
      _cx             =   4339
      _cy             =   2275
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPiutang.frx":06B1
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
      Height          =   255
      Left            =   405
      TabIndex        =   25
      Top             =   4365
      Width           =   6285
   End
   Begin VB.Label Label6 
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
      Height          =   345
      Left            =   4725
      TabIndex        =   22
      Top             =   1620
      Width           =   8160
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&F"
      Height          =   255
      Index           =   5
      Left            =   1740
      TabIndex        =   9
      Top             =   5820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&B"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   7
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total : "
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
      Left            =   8535
      TabIndex        =   19
      Top             =   4350
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl+F4=Close  F4=Bantuan   Ins=Tambah Grid   Del=Hapus Grid"
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
      Left            =   405
      TabIndex        =   18
      Top             =   8520
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
      Left            =   390
      TabIndex        =   11
      Top             =   7290
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total : "
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
      Index           =   8
      Left            =   8640
      TabIndex        =   17
      Top             =   7305
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Kepada : "
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
      Left            =   465
      TabIndex        =   5
      Top             =   1605
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
      Left            =   465
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "K&ode Pembayaran Piutang : "
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
      Left            =   465
      TabIndex        =   0
      Top             =   780
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Piutang"
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
      TabIndex        =   16
      Top             =   105
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   8865
      Left            =   240
      Top             =   210
      Width           =   12885
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   9090
      Left            =   165
      Top             =   120
      Width           =   13095
   End
End
Attribute VB_Name = "frmPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'0=nofaktur  1=total  2=bayar 3=potongan  4=% 5=sisa  6=subtotal  7=intno 8=realDiskon
Option Explicit
Public JenisUtPi As String
Dim intNoFaktur As Single
Dim LbhByr As Currency
Dim BaruBuka As Boolean


Private Sub Command1_Click()
On Error GoTo aa
Dim kExec As String
  kExec = aData.UtangPiutangHapus(JenisUtPi, AmanOi(t1.Text))
  If kExec <> "" Then
  MsgBox kExec
  Else
  MsgBox "Data pembayaran telah dihapus..", vbInformation, "Hapus Data"
  Call CmdBaru_Click
  End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command1_Click pada Form frmPiutang"
End Sub

Private Sub Form_Load()
On Error GoTo aa
BaruBuka = True
If JenisUtPi = "Utang" Then
Me.Caption = "Pembayaran Utang"
CmdCekFaktur.Caption = "&Cek Utang"
lblLabels(1).Caption = "K&ode Pembayaran Utang :"
lblLabels(3).Caption = "&Kepada :"
Label2.Caption = "Utang"
Else
Me.Caption = "Pembayaran Piutang"
CmdCekFaktur.Caption = "&Cek Piutang"
lblLabels(1).Caption = "K&ode Pembayaran Piutang :"
Label2.Caption = "Piutang"
lblLabels(3).Caption = "&Dari :"
End If
Utang.ColComboList(0) = "#T;Cash|#G;Giro|#B;Transfer/Lain"
Call CmdBaru_Click
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdBaru_Click()
On Error GoTo aa
If BaruBuka Then
  BaruBuka = False
Else
  If MsgBox("Anda yakin ingin melanjutkan dengan pembayaran baru ..?", vbQuestion + vbYesNo + vbDefaultButton2, "Transaksi Baru") = vbNo Then
    Exit Sub
  End If
End If
CmdSimpan.Visible = True
If JenisUtPi = "Utang" Then
t1 = SerbaGuna.NoUtang("H", "Utang")
Else
t1 = SerbaGuna.NoUtang("P", "Piutang")
End If
VGrid.Rows = 1
Label1.Caption = ""
t2.Value = Date
Master.Rows = 1
Master.Rows = 2
Master.Row = 1
Master.Col = 0
Utang.Rows = 1
Utang.Rows = 2
Utang.Row = 1
Utang.Col = 0
Label6.Caption = ""
t3 = ""
intNoFaktur = 0
t8 = ""
TotFaktur.Value = 0
TotBayar.Value = 0
On Error Resume Next
t1.SetFocus
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub




Private Sub CmdCekFaktur_Click()
  On Error GoTo aa
  
Dim rsMaster As New ADODB.Recordset
If JenisUtPi = "Utang" Then
Set rsMaster = aData.AmbilCommand("select * from Utang where [KodeBayar]='" & t1.Text & "'")
Else
Set rsMaster = aData.AmbilCommand("select * from Piutang where [KodeBayar]='" & t1.Text & "'")
End If

  If Not rsMaster.EOF Then
  t1.Text = rsMaster![KodeBayar]
  t2.Value = rsMaster!Tgl
  t3.Text = rsMaster!Kepada
  Call t3_Validate(False)
  t8.Text = rsMaster!Keterangan
  TotBayar.Value = rsMaster!Total
  TotFaktur.Value = rsMaster!Total
  intNoFaktur = rsMaster!intNo
  
    
    Master.Rows = 1
    
    If JenisUtPi = "Utang" Then
   Set rsMaster = aData.AmbilCommand("SELECT * from UtangAnak where KodeBayar=" & intNoFaktur & "")
Else
    Set rsMaster = aData.AmbilCommand("SELECT * from PiutangAnak where KodeBayar=" & intNoFaktur & "")
End If
    Utang.Rows = 1
    Do While Not rsMaster.EOF
    Utang.Rows = Utang.Rows + 1
    Utang.TextMatrix(Utang.Rows - 1, 0) = rsMaster!jenisbayar
    Utang.TextMatrix(Utang.Rows - 1, 1) = rsMaster!nogiro
    Utang.TextMatrix(Utang.Rows - 1, 2) = rsMaster!namabank
    Utang.TextMatrix(Utang.Rows - 1, 3) = rsMaster!jatuhtempo
    Utang.TextMatrix(Utang.Rows - 1, 4) = rsMaster!Total
    Utang.TextMatrix(Utang.Rows - 1, 5) = rsMaster!Keterangan
Dim rsMasterA As New ADODB.Recordset
    If JenisUtPi = "Utang" Then
   Set rsMasterA = aData.AmbilCommand("SELECT DetailUtang.Total, DetailUtang.Potongan, TransBeli.intno as NooI,TransBeli.[No Faktur], TransBeli.Total AS TotFaktur " & _
"FROM (Supplier INNER JOIN TransBeli ON Supplier.Kode=TransBeli.Kepada) INNER JOIN DetailUtang ON TransBeli.intno=DetailUtang.KodeFaktur " & _
"where KodeBayar=" & rsMaster!intNo & "")
Else
    Set rsMasterA = aData.AmbilCommand("SELECT DetailPiutang.Total, DetailPiutang.Potongan, TransJual.intno as NooI,TransJual.[No Faktur], TransJual.Total as TotFaktur " & _
"FROM (Konsumen INNER JOIN TransJual ON Konsumen.Kode=TransJual.Kepada) INNER JOIN DetailPiutang ON TransJual.intno=DetailPiutang.KodeFaktur " & _
"where KodeBayar=" & rsMaster!intNo & "")
End If
    Do While Not rsMasterA.EOF
    Master.Rows = Master.Rows + 1
    Master.TextMatrix(Master.Rows - 1, 0) = rsMasterA![No Faktur]
    Master.TextMatrix(Master.Rows - 1, 1) = rsMasterA!TotFaktur
    Master.TextMatrix(Master.Rows - 1, 2) = rsMasterA!Total + rsMasterA!Potongan
    Master.TextMatrix(Master.Rows - 1, 3) = rsMasterA!Potongan
    Master.TextMatrix(Master.Rows - 1, 4) = False
    Master.TextMatrix(Master.Rows - 1, 5) = 0
    Master.TextMatrix(Master.Rows - 1, 6) = Master.ValueMatrix(Master.Rows - 1, 2) - Master.ValueMatrix(Master.Rows - 1, 3)
    Master.TextMatrix(Master.Rows - 1, 7) = rsMasterA!nooi
    Master.TextMatrix(Master.Rows - 1, 8) = Master.ValueMatrix(Master.Rows - 1, 3)
    rsMasterA.MoveNext
    Loop
    rsMaster.MoveNext
    Loop
    Call TotalOi

Else
  MsgBox "No bukti " & t1 & vbCrLf & "tidak ditemukan pada database atau telah dihapus", vbInformation, "Cek Faktur"
  Call CmdBaru_Click
  End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub CmdSimpan_Click()
On Error GoTo aa
Call TotalOi
Dim aHead(5) As String, aBodi() As String, i As Byte, aHasil As String
Dim kLbh As Single, Dpt As Boolean
  For kLbh = 1 To Utang.Rows - 1
    If Utang.TextMatrix(kLbh, 0) = "LB" Then Utang.RemoveItem kLbh
  Next kLbh

If TotBayar.Value < TotFaktur.Value Then
  If LbhByr - (TotFaktur.Value - TotBayar.Value) >= 0 Then
    If MsgBox("Total Faktur lebih besar dari Total Pembayaran." & vbCrLf & _
    "Selisih " & FormatNumber(TotBayar.Value - TotFaktur.Value, 2) & _
    " akan diambil dari lebih bayar..?", vbInformation + vbYesNo, "Simpan Data") = vbYes Then
      Utang.Rows = Utang.Rows + 1
      Utang.TextMatrix(Utang.Rows - 1, 0) = "LB"
      Utang.TextMatrix(Utang.Rows - 1, 4) = TotFaktur.Value - TotBayar.Value
      Call TotalOi
    Else
      GoTo abc
    End If
  Else
abc:
  MsgBox "Total Pembayaran kurang dari Total Faktur. Pembayaran tidak dapat disimpan.." & vbCrLf & _
  "Harap data yang dimasukkan dicek sekali lagi..", vbInformation, "Simpan Data"
  Exit Sub
  End If
ElseIf TotBayar.Value > TotFaktur.Value Then
  MsgBox "Total Pembayaran lebih besar dari Total Faktur." & vbCrLf & _
  "Selisih " & FormatNumber(TotBayar.Value - TotFaktur.Value, 2) & _
  " akan dialokasikan ke lebih bayar..", vbInformation, "Simpan Data"
End If
'0=NoBayar  1=Tgl  2=Supplier  3=Total  4=Keterangan
aHead(0) = SerbaGuna.AmanOi(t1.Text) 'NoBayar
aHead(1) = SerbaGuna.AmanTgl(t2.Value)  'Tgl
aHead(2) = SerbaGuna.AmanOi(t3.Text) 'Konsumen
aHead(3) = SerbaGuna.AmanOi(TotBayar.Value)   'Total
aHead(4) = SerbaGuna.AmanOi(t8.Text) 'Keterangan

    If JenisUtPi = "Utang" Then
    aHasil = aData.SimpanUtangPiutang(aHead, Utang, Master, "Utang", VGrid)
    Else
    aHasil = aData.SimpanUtangPiutang(aHead, Utang, Master, "Piutang", VGrid)
    End If
If aHasil <> "" Then
MsgBox aHasil, vbInformation, Me.Caption & "#Error"
Else
BaruBuka = True
Call CmdBaru_Click
End If

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
'****************************************************************

'****************************************************************
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

If Master.Col = 0 And KeyCode = vbKeyReturn Then
     Bantu.Apaan = "Grid"
     If JenisUtPi = "Utang" Then
     Bantu.GridData = "select * from InputUtang where kepada='" & t3.Text & "'"
     Else
     Bantu.GridData = "select * from InputPiutang where kepada='" & t3.Text & "'"
     End If
     Bantu.Left = Master.CellLeft + Master.Left + 50
     Bantu.Top = Master.CellTop + Master.Top + Master.CellHeight + 400
     Bantu.Show vbModal
     If Not Bantu.Batal Then
     Master.Text = BarisGrid(1)
     Master.TextMatrix(Row, 1) = ""
     Call Master_AfterEdit(Row, Col)
End If
End If

If KeyCode = vbKeyReturn And Col = 3 Then
     cBantu.IntNoFak = Master.TextMatrix(Row, 7)
     cBantu.Keterangan = t3.Text & "-" & Label6.Caption & " =>" & Master.TextMatrix(Row, 0)
     cBantu.Jenis = JenisUtPi
     cBantu.Show vbModal
     If Not cBantu.Batal Then
      Dim i As Single
      i = 1
      Do While i < VGrid.Rows
        If VGrid.ValueMatrix(i, 0) = Master.ValueMatrix(Row, 7) Then
        VGrid.RemoveItem (i)
        Else
        i = i + 1
        End If
      Loop
      For i = LBound(xPotHar, 1) To UBound(xPotHar, 1)
        If xPotHar(i, 2) <> 0 Then
        VGrid.Rows = VGrid.Rows + 1
        VGrid.TextMatrix(VGrid.Rows - 1, 0) = xPotHar(i, 0)
        VGrid.TextMatrix(VGrid.Rows - 1, 1) = xPotHar(i, 1)
        VGrid.TextMatrix(VGrid.Rows - 1, 2) = xPotHar(i, 2)
        End If
      Next i
     Master.Text = IIf(Master.ValueMatrix(Row, 1) < 0, -cBantu.NilaiBantu, cBantu.NilaiBantu)
     Master.TextMatrix(Row, 4) = 0
     Call Master_AfterEdit(Row, 2)
     
     End If
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
'0=nofaktur  1=total  2=bayar 3=potongan  4=% 5=sisa  6=subtotal  7=intno 8=realDiskon
Private Sub Master_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo aa
    Select Case Col
    Case 0
    Case 1, 5, 6
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
     If Master.TextMatrix(Row, Col) <> "" And Master.TextMatrix(Row, 1) = "" Then
     Dim rsSbrg As New ADODB.Recordset
     If JenisUtPi = "Utang" Then
     Set rsSbrg = aData.AmbilCommand("select * from InputUtang where [No Faktur]='" & Master.TextMatrix(Master.Row, 0) & "' and Kepada='" & t3.Text & "'")
     Else
     Set rsSbrg = aData.AmbilCommand("select * from InputPiutang where [No FAktur]='" & Master.TextMatrix(Master.Row, 0) & "' and Kepada='" & t3.Text & "'")
     End If
     If rsSbrg.EOF Then
     MsgBox "Nomor Faktur tersebut tidak ada, harap cek kembali nomor Faktur anda", vbInformation, "No. Faktur"
     Master.TextMatrix(Row, Col) = ""
     Exit Sub
     Else
     Master.Text = rsSbrg![No Faktur]
     Master.TextMatrix(Row, Col + 1) = rsSbrg![Sisa]
     Master.TextMatrix(Row, Col + 2) = rsSbrg![Sisa]
     Master.TextMatrix(Row, Col + 3) = 0
     Master.TextMatrix(Row, Col + 4) = False
     Master.TextMatrix(Row, Col + 5) = 0
     Master.TextMatrix(Row, Col + 6) = rsSbrg![Sisa]
     Master.TextMatrix(Row, Col + 7) = rsSbrg![intNo]
     Master.TextMatrix(Row, 8) = 0
     Call TotalOi
     Master.Col = 3
     End If
     End If
    Case 2, 3, 4
      
     '0=nofaktur  1=total  2=bayar 3=potongan  4=% 5=sisa  6=subtotal  7=intno 8=realDiskon
     Master.TextMatrix(Row, 2) = FormatNumber(Master.TextMatrix(Row, 2), 2)
     Master.TextMatrix(Row, 3) = FormatNumber(Master.TextMatrix(Row, 3), 2)
     Master.TextMatrix(Row, 8) = FormatNumber(IIf(Master.TextMatrix(Row, 4), (Master.TextMatrix(Row, 3) / 100) * Master.TextMatrix(Row, 1), Master.TextMatrix(Row, 3)), 2)
     Master.TextMatrix(Row, 5) = Master.TextMatrix(Row, 1) - Master.TextMatrix(Row, 2)
     Master.TextMatrix(Row, 6) = Master.TextMatrix(Row, 2) - Master.TextMatrix(Row, 8)
     Call TotalOi
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub TotalOi()
On Error GoTo aa
TotBayar.Value = 0
  Dim i As Byte
  For i = 1 To Utang.Rows - 1
  TotBayar.Value = TotBayar.Value + IIf(Utang.TextMatrix(i, 4) = "", 0, Utang.TextMatrix(i, 4))
  Next i
  
TotFaktur.Value = 0
  For i = 1 To Master.Rows - 1
  TotFaktur.Value = TotFaktur.Value + IIf(Master.TextMatrix(i, 6) = "", 0, Master.TextMatrix(i, 6))
  Next i
If (TotBayar.Value - TotFaktur.Value) = 0 Then
Label1.Caption = ""
ElseIf (TotBayar.Value - TotFaktur.Value) > 0 Then
Label1.Caption = "Lebih Bayar : " & _
FormatCurrency((TotBayar.Value - TotFaktur.Value), 2)
Else
Label1.Caption = "Kurang Bayar : " & _
FormatCurrency(-(TotBayar.Value - TotFaktur.Value), 2)
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub T2_GotFocus()
t2.SelStart = 0
End Sub

Private Sub t3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyF4 Then
     Bantu.Apaan = "Grid"
     If JenisUtPi = "Utang" Then
     Bantu.GridData = "select Kode,Nama, Alamat, Kota,Telepon from Supplier order by Nama"
     Else
     Bantu.GridData = "select Kode,Nama, Alamat, Kota,Telepon from Konsumen order by Nama"
     End If
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

Private Sub t3_Validate(Cancel As Boolean)
'Tampilkan Salesman dari Konsumen
On Error GoTo aa
Dim rsKodok As New ADODB.Recordset
If JenisUtPi = "Piutang" Then
Set rsKodok = aData.AmbilCommand("Select * from Konsumen where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
  LbhByr = IIf(IsNull(rsKodok![Lebih Bayar]), 0, rsKodok![Lebih Bayar])
  Label6.Caption = IIf(IsNull(rsKodok!Nama), "", rsKodok!Nama) & "-" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & " (" & FormatNumber(LbhByr, 2) & ")"
  End If
Else
  Set rsKodok = aData.AmbilCommand("Select * from Supplier where kode='" & AmanOi(t3.Text) & "'")
  Label6.Caption = ""
  If Not rsKodok.EOF Then
  LbhByr = IIf(IsNull(rsKodok![Lebih Bayar]), 0, rsKodok![Lebih Bayar])
  Label6.Caption = IIf(IsNull(rsKodok!Nama), "", rsKodok!Nama) & "-" & IIf(IsNull(rsKodok!Alamat), "", rsKodok!Alamat) & " (" & FormatNumber(LbhByr, 2) & ")"
  End If
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur t3_Validate pada Form frmPiutang"
End Sub

Private Sub Utang_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
Dim kLbh As Single, BnykDpt As Byte
    Select Case Col
    Case 4
     Call TotalOi
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Utang_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyInsert Then
Utang.Rows = Utang.Rows + 1
Utang.Row = Utang.Rows - 1
Utang.Col = 0
End If
If KeyCode = vbKeyDelete And Utang.Rows <> 1 Then
Utang.RemoveItem Utang.Row
End If
End Sub

Private Sub Utang_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyF4 Then
    Select Case Col
    Case 4
     Bantu.Apaan = "Angka"
     Bantu.Left = Utang.CellLeft + Utang.Left + 50
     Bantu.Top = Utang.CellTop + Utang.Top + Utang.CellHeight + 400
     Bantu.Show vbModal
     Utang.Text = Bantu.NilaiBantu
     Call TotalOi
    Case 3
     Bantu.Apaan = "Tanggal"
     Bantu.Left = Utang.CellLeft + Utang.Left + 50
     Bantu.Top = Utang.CellTop + Utang.Top + Utang.CellHeight + 400
     Bantu.Show vbModal
     Utang.Text = Bantu.NilaiBantu
     Call TotalOi
    End Select
  End If
End Sub

Private Sub Utang_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Select Case Col
   Case 1, 2, 3
    If Utang.TextMatrix(Row, 0) = "T" Or Utang.TextMatrix(Row, 0) = "LB" Then Cancel = True
  End Select
End Sub
