VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTransBatch 
   Caption         =   "Proses Offline Transaksi"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tandai transaksi yang diproses dengan sukses"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   5235
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Proses"
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
      Left            =   480
      TabIndex        =   2
      Top             =   8280
      Width           =   1290
   End
   Begin VB.TextBox txtFields 
      DataField       =   "JenisTrans"
      Height          =   315
      Left            =   540
      TabIndex        =   3
      Top             =   9990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   6720
      Left            =   435
      TabIndex        =   1
      Top             =   1440
      Width           =   13995
      _cx             =   24686
      _cy             =   11853
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
      FormatString    =   $"frmTransBatch.frx":0000
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
   Begin TDBDate6Ctl.TDBDate tgl1 
      Height          =   405
      Left            =   840
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   714
      Calendar        =   "frmTransBatch.frx":0108
      Caption         =   "frmTransBatch.frx":023E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTransBatch.frx":02AC
      Keys            =   "frmTransBatch.frx":02CA
      Spin            =   "frmTransBatch.frx":0328
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
      TabIndex        =   10
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
      FormatString    =   $"frmTransBatch.frx":0350
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
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   714
      Calculator      =   "frmTransBatch.frx":04B3
      Caption         =   "frmTransBatch.frx":04D3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTransBatch.frx":053F
      Keys            =   "frmTransBatch.frx":055D
      Spin            =   "frmTransBatch.frx":05A7
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
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Batch Transaction Process"
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
      TabIndex        =   6
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   9495
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   9135
      Left            =   240
      Top             =   240
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
Attribute VB_Name = "frmTransBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim CD2 As New myComDialog.cComDialog
On Error GoTo aa
    CD2.InitDir = App.Path '& IIf(Right(App.Path, 1) = "\", "Export", "\Export")
    CD2.FileName = ""
    On Error Resume Next
    CD2.CancelError = True
    CD2.DialogTitle = "Masukkan nama file Transaksi yang akan diimport.."
    CD2.Filter = "Transaksi File|*.trns"
    CD2.Flags = cdlOFNFileMustExist Or OFN_PATHMUSTEXIST Or OFN_ALLOWMULTISELECT
    CD2.ShowOpen
    Dim zk() As String, i As Long, xPath As String
    
If Err.Number = 0 Then
  Master.Rows = 1
  zk = Split(CD2.FileName, vbNullChar)
           
  If Right(zk(0), 5) = ".Trns" Then
    ProsesFile zk(0)
  Else
    xPath = IIf(Right(zk(0), 1) <> "\", zk(0) & "\", zk(0))
    i = 1
    Do While zk(i) <> ""
      ProsesFile xPath & zk(i)
      i = i + 1
    Loop
  End If
End If

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


Private Sub ProsesFile(a As String)
On Error GoTo aa
Dim myDo As New ADODB.Recordset, aRo As Integer
myDo.Open a
With Master
  .Rows = .Rows + 1
  aRo = .Rows - 1
  '0=no 1=tgl 2=pihak 3=jentrans 4=file 5=status 6=datA 7=datB 8=""
  myDo.MoveFirst
  .TextMatrix(aRo, 0) = myDo!Kode
  .TextMatrix(aRo, 1) = myDo!Nama
  .TextMatrix(aRo, 2) = myDo!QtyB
  .TextMatrix(aRo, 3) = NamaTrans(myDo!Diskon)
  .TextMatrix(aRo, 4) = a
  .TextMatrix(aRo, 5) = ""
  .TextMatrix(aRo, 6) = "" 'k!
  .TextMatrix(aRo, 7) = "" 'k!
  .TextMatrix(aRo, 8) = ""
End With
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


Private Sub CmdSimpan_Click()
On Error GoTo aa
Dim i As Long
For i = 1 To Master.Rows - 1
If Master.TextMatrix(i, 0) <> "" Then
  ProsesTrans Master.TextMatrix(i, 4), i
End If
Next i
i = 1
Do While i < Master.Rows
If Master.TextMatrix(i, 5) = "OK" Then
  On Error Resume Next
  FileCopy Master.TextMatrix(i, 4), Master.TextMatrix(i, 4) & "OK"
  'CopyFile Master.TextMatrix(i, 4), Master.TextMatrix(i, 4) & "OK"
  If Err.Number = 0 Then
    Kill Master.TextMatrix(i, 4)
  End If
  Master.RemoveItem i
Else
  i = i + 1
End If
Loop
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub ProsesTrans(a As String, iRo As Long)
On Error GoTo aa
Dim myDo As New ADODB.Recordset, aRo As Integer
myDo.Open a

Dim aHead(11) As String
aHead(0) = SerbaGuna.AmanOi(myDo!Kode) 'NoFaktur
tgl1.Value = myDo!Nama: aHead(1) = SerbaGuna.AmanTgl(tgl1.Value)  'Tgl
aHead(2) = SerbaGuna.AmanOi(myDo!QtyB) 'Supplier
aHead(3) = SerbaGuna.AmanOi(myDo!SatB) 'Salesman
num1.Value = Val(myDo!QtyS): aHead(4) = num1.Value  'Diskon
aHead(5) = IIf(myDo!SatS = "Cash", "C", "K") 'CaraPembayaran
num1.Value = Val(myDo!Harga): aHead(6) = Val(SerbaGuna.AmanOi(num1.Value))   'LamaKredit
aHead(7) = SerbaGuna.AmanOi(myDo!Diskon) 'JenisTrans
aHead(8) = SerbaGuna.AmanOi(myDo!Karton) 'Keterangan

Grid1.Rows = 1
Dim aBodi() As String, i As Byte, aHasil As String
     myDo.MoveNext
     Do While Not myDo.EOF
       Grid1.Rows = Grid1.Rows + 1
       Grid1.TextMatrix(Grid1.Rows - 1, 0) = myDo!Kode
       Grid1.TextMatrix(Grid1.Rows - 1, 1) = myDo!Nama
       Grid1.TextMatrix(Grid1.Rows - 1, 2) = Val(myDo!QtyB)
       Grid1.TextMatrix(Grid1.Rows - 1, 3) = myDo!SatB
       Grid1.TextMatrix(Grid1.Rows - 1, 4) = Val(myDo!QtyS)
       Grid1.TextMatrix(Grid1.Rows - 1, 5) = myDo!SatS
       Grid1.TextMatrix(Grid1.Rows - 1, 6) = Val(myDo!Harga)
       Grid1.TextMatrix(Grid1.Rows - 1, 7) = Val(myDo!Diskon)
       Grid1.TextMatrix(Grid1.Rows - 1, 8) = False
       Grid1.TextMatrix(Grid1.Rows - 1, 10) = Val(myDo!Karton)
       Grid1.TextMatrix(Grid1.Rows - 1, 11) = Val(myDo!Diskon)
       Call grid1_AfterEdit(Grid1.Rows - 1, 2)
       myDo.MoveNext
     Loop
     ReDim aBodi(Grid1.Rows - 1, 7)
  
For i = 1 To Grid1.Rows - 1
If Grid1.TextMatrix(i, 0) <> "" Then
aBodi(i, 0) = SerbaGuna.AmanOi(Grid1.TextMatrix(i, 0)) 'Kode
aBodi(i, 1) = Grid1.ValueMatrix(i, 6)  'Harga
aBodi(i, 2) = Grid1.ValueMatrix(i, 2) + (Grid1.ValueMatrix(i, 4) / Grid1.ValueMatrix(i, 10)) 'Qty
aBodi(i, 3) = Grid1.ValueMatrix(i, 11) 'Diskon
aBodi(i, 4) = Grid1.ValueMatrix(i, 10) 'Pengali
aBodi(i, 5) = Grid1.ValueMatrix(i, 2) 'QtyB
aBodi(i, 6) = Grid1.ValueMatrix(i, 4) 'QtyS
End If
Next i

num1.Value = Val(TotalOi): aHead(9) = num1.Value  'Total
aHead(10) = False
  
  
If Master.TextMatrix(iRo, 3) = "Penjualan" Then
  'If Not CekSimpanJual(aBodi) Then Exit Sub
End If
If Master.TextMatrix(iRo, 3) = "Pembelian" Or Master.TextMatrix(iRo, 3) = "Retur Pembelian" Then
aHasil = aData.SimpanBeli(aHead, aBodi)
ElseIf Master.TextMatrix(iRo, 3) = "Penjualan" Or Master.TextMatrix(iRo, 3) = "Retur Penjualan" Then
aHasil = aData.SimpanJual(aHead, aBodi)
End If
If aHasil <> "" Then
  Master.TextMatrix(iRo, 5) = aHasil
Else
  Master.TextMatrix(iRo, 5) = "OK"
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Function NamaTrans(a As String) As String
On Error Resume Next
If a = "J" Then
  NamaTrans = "Penjualan"
ElseIf a = "B" Then
  NamaTrans = "Pembelian"
ElseIf a = "RJ" Then
  NamaTrans = "Retur Penjualan"
Else
  NamaTrans = "Retur Pembelian"
End If
End Function

Private Function TotalOi() As Currency
On Error GoTo aa
Dim k As Currency
k = 0
  Dim i As Byte
  For i = 1 To Grid1.Rows - 1
  k = k + IIf(Grid1.TextMatrix(i, 9) = "", 0, Grid1.TextMatrix(i, 9))
  Next i
  TotalOi = k
Exit Function
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Function

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
    Select Case Col
    Case 2, 4, 6, 7, 8
      If (Grid1.TextMatrix(Row, 8)) Then
        Grid1.TextMatrix(Row, 11) = (Grid1.ValueMatrix(Row, 7) / 100) * Grid1.ValueMatrix(Row, 6)
      Else
        Grid1.TextMatrix(Row, 11) = Grid1.ValueMatrix(Row, 7)
      End If
      Grid1.TextMatrix(Row, 9) = FormatNumber((Grid1.ValueMatrix(Row, 6) - _
      Grid1.ValueMatrix(Row, 11)) * _
      (Grid1.ValueMatrix(Row, 2) + (Grid1.ValueMatrix(Row, 4) / Grid1.ValueMatrix(Row, 10))) _
      , 2)
    End Select
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub



Private Sub Master_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete And Master.Rows <> 1 Then
Master.RemoveItem Master.Row
End If
End Sub


