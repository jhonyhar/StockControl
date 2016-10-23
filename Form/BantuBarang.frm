VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form BantuBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barang"
   ClientHeight    =   4500
   ClientLeft      =   3105
   ClientTop       =   4365
   ClientWidth     =   10245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1530
      TabIndex        =   1
      Top             =   120
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6090
      TabIndex        =   3
      Top             =   120
      Width           =   3075
   End
   Begin VSFlex8Ctl.VSFlexGrid VGrid 
      Height          =   3585
      Left            =   225
      TabIndex        =   4
      Top             =   660
      Width           =   9810
      _cx             =   17304
      _cy             =   6324
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"BantuBarang.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Kode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   165
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   165
      Width           =   1260
   End
End
Attribute VB_Name = "BantuBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aSet As New ADODB.Recordset, Penanda As Boolean, Tambah As Boolean
Public GridData As String
Public Batal As Boolean

Private Sub Form_Activate()
On Error Resume Next
Penanda = True
Tambah = False
Batal = False
ReDim SerbaGuna.BarisGrid(VGrid.Cols - 1)
Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyEscape Then
Batal = True
Unload Me
ElseIf KeyCode = vbKeyInsert Then
Exit Sub
Dim a As String, b As String
a = InputBox("Masukkan kode barang yang baru", "Tambah Data Barang")
b = InputBox("Masukkan nama barang yang baru", "Tambah Data Barang")
    If a = "" Or b = "" Then
    MsgBox "Data yang anda masukkan tidak lengkap", vbInformation, "Tambah Data Barang"
    Exit Sub
    End If
Dim HslTambah As String
HslTambah = aData.ExecCommand("insert into barang([Kode Barang], [Nama Barang]) values('" & AmanOi(a) & "','" & AmanOi(b) & "')")
If HslTambah = "" Then
SerbaGuna.BarisGrid(0) = a
SerbaGuna.BarisGrid(1) = b
SerbaGuna.BarisGrid(2) = 0
SerbaGuna.BarisGrid(3) = 0
SerbaGuna.BarisGrid(4) = 0
SerbaGuna.BarisGrid(5) = 1
Tambah = True
Unload Me
Else
MsgBox "Kode barang tersebut telah ada", vbInformation, "Tambah data barang"
End If
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Form_Load()
On Error GoTo aa
Penanda = True
Tambah = False
Batal = False
Set aSet = aData.AmbilData("SELECT [Kode Barang], [Nama Barang], Qty, Satuan, [Harga Beli], [Harga Jual], " & _
"'1 ' & [Satuan] & ' = ' & [Qty Satuan Kecil] & ' ' & [Satuan Kecil] as [Satuan Kecil] FROM Barang order by [Nama Barang];")
Set VGrid.DataSource = aSet
ReDim SerbaGuna.BarisGrid(VGrid.Cols - 1)
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo aa
If Tambah Then Exit Sub
If Not Batal Then
Dim i As Byte
For i = LBound(SerbaGuna.BarisGrid) To UBound(SerbaGuna.BarisGrid)
SerbaGuna.BarisGrid(i) = VGrid.TextMatrix(VGrid.Row, i)
Next i
End If
If Not Utama.OOQ Then Cancel = True
Me.Hide
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyReturn Then
Call FilterAA("Nama")
ElseIf KeyCode = vbKeyDown Then
VGrid.SetFocus
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub


Private Sub FilterAA(a As String)
On Error GoTo aa
aSet.Filter = ""

If Text1.Text = "" And Text2.Text = "" Then
  Set aSet = aData.AmbilData("SELECT [Kode Barang], [Nama Barang], Qty, Satuan, [Harga Beli], [Harga Jual], " & _
  "'1 ' & [Satuan] & ' = ' & [Qty Satuan Kecil] & ' ' & [Satuan Kecil] as [Satuan Kecil] FROM Barang order by [Nama Barang];")
ElseIf a = "Nama" And Text1.Text <> "" Then
  Set aSet = aData.AmbilData("SELECT [Kode Barang], [Nama Barang], Qty, Satuan, [Harga Beli], [Harga Jual], " & _
  "'1 ' & [Satuan] & ' = ' & [Qty Satuan Kecil] & ' ' & [Satuan Kecil] as [Satuan Kecil] FROM Barang " & _
  "Where [Nama Barang] like '%" & AmanOi(Text1.Text) & "%' " & _
  "order by [Nama Barang];")
  'aSet.Filter = "[Nama Barang] like '%" & AmanOi(Text1.Text) & "%'"
ElseIf a = "Kode" And Text2.Text <> "" Then
  Set aSet = aData.AmbilData("SELECT [Kode Barang], [Nama Barang], Qty, Satuan, [Harga Beli], [Harga Jual], " & _
  "'1 ' & [Satuan] & ' = ' & [Qty Satuan Kecil] & ' ' & [Satuan Kecil] as [Satuan Kecil] FROM Barang " & _
  "Where [Kode Barang] like '%" & AmanOi(Text2.Text) & "%' " & _
  "order by [Nama Barang];")
  'aSet.Filter = "[Kode Barang] like '%" & AmanOi(Text2.Text) & "%'"
End If
Set VGrid.DataSource = aSet
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

Private Sub VGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If Penanda Then
Penanda = False
Exit Sub
End If
If KeyCode = vbKeyReturn Then
KeyCode = 0
Unload Me
ElseIf KeyCode = vbKeyUp And VGrid.Row = 1 Then
Text1.SetFocus
End If
End Sub


Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
If KeyCode = vbKeyReturn Then
Call FilterAA("Kode")
ElseIf KeyCode = vbKeyDown Then
VGrid.SetFocus
End If
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "#Error"
End Sub

