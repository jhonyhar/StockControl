VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form Bantu 
   BorderStyle     =   0  'None
   ClientHeight    =   3720
   ClientLeft      =   3405
   ClientTop       =   4665
   ClientWidth     =   7980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid VGrid 
      Height          =   3585
      Left            =   45
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   7770
      _cx             =   13705
      _cy             =   6324
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   42
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Bantu.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
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
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   555
      Left            =   -2000
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   525
      Left            =   -2000
      TabIndex        =   3
      Top             =   4440
      Width           =   675
   End
   Begin TDBNumber6Ctl.TDBNumber BAngka 
      Height          =   360
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   6006
      _ExtentY        =   635
      Calculator      =   "Bantu.frx":0335
      Caption         =   "Bantu.frx":0355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Bantu.frx":03C1
      Keys            =   "Bantu.frx":03DF
      Spin            =   "Bantu.frx":0429
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "#,###,###,###.##;-#,###,###,###.##;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,###.##"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999999999
      MinValue        =   -9999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   144572417
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBText6Ctl.TDBText BTeks 
      Height          =   405
      Left            =   135
      TabIndex        =   0
      Top             =   1140
      Visible         =   0   'False
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   6006
      _ExtentY        =   714
      Caption         =   "Bantu.frx":0451
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Bantu.frx":04BD
      Key             =   "Bantu.frx":04DB
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate BTanggal 
      Height          =   360
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   6006
      _ExtentY        =   635
      Calendar        =   "Bantu.frx":051F
      Caption         =   "Bantu.frx":064B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Bantu.frx":06B7
      Keys            =   "Bantu.frx":06D5
      Spin            =   "Bantu.frx":0733
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd mmmm yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
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
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "04/03/2006"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   38810
      CenturyMode     =   0
   End
End
Attribute VB_Name = "Bantu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Penanda As Boolean
Public Batal As Boolean

Public NilaiBantu
Public Apaan As String
Public GridData As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Batal = True
Unload Me
End Sub

Private Sub Form_Load()
BAngka = 0
BTanggal = Date
BTeks = ""

BAngka.Top = 0
BAngka.Left = 0
BTanggal.Top = 0
BTanggal.Left = 0
BTeks.Top = 0
BTeks.Left = 0
VGrid.Top = 0
VGrid.Left = 0
Me.Height = BAngka.Height
Me.Width = BAngka.Width

Penanda = True
Batal = False


Select Case Apaan
Case "Tanggal"
BTanggal.Visible = True
Case "Angka"
BAngka.Visible = True
Case "Grid"
VGrid.Visible = True
Me.Height = VGrid.Height
Me.Width = VGrid.Width
Set VGrid.DataSource = aData.AmbilData(GridData)
Case Else
BTeks.Visible = True
End Select
SendKeys "{F4}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Batal Then
NilaiBantu = 0
ReDim SerbaGuna.BarisGrid(1)
SerbaGuna.BarisGrid(0) = ""
Exit Sub
End If
Select Case Apaan
Case "Tanggal"
NilaiBantu = BTanggal.Value
Case "Angka"
NilaiBantu = BAngka.Value
Case "Grid"
ReDim SerbaGuna.BarisGrid(VGrid.Cols - 1)
Dim i As Byte
For i = LBound(SerbaGuna.BarisGrid) To UBound(SerbaGuna.BarisGrid)
SerbaGuna.BarisGrid(i) = VGrid.TextMatrix(VGrid.Row, i)
Next i
NilaiBantu = ""
Case Else
NilaiBantu = BTeks.Text
End Select
End Sub

Private Sub VGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If Penanda Then
Penanda = False
Exit Sub
End If
If KeyCode = vbKeyReturn Then
Unload Me
End If
End Sub
