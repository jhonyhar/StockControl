VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form bBantu 
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   1740
   ClientTop       =   4200
   ClientWidth     =   11565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid VGrid 
      Height          =   3585
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   11415
      _cx             =   20135
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
      FormatString    =   $"bBantu.frx":0000
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
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   555
      Left            =   -2000
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   525
      Left            =   -2000
      TabIndex        =   0
      Top             =   4440
      Width           =   675
   End
End
Attribute VB_Name = "bBantu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Penanda As Boolean
Public Batal As Boolean
Public RSSData As New ADODB.Recordset
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
On Error GoTo aa
VGrid.Top = 10
VGrid.Left = 10
Penanda = True
Batal = False
VGrid.Visible = True
Me.Height = VGrid.Height + 20
Me.Width = VGrid.Width + 20
With VGrid
.Rows = 1
Do While Not RSSData.EOF
.Rows = VGrid.Rows + 1
.Row = VGrid.Rows - 1
.TextMatrix(.Row, 0) = RSSData![No Faktur]
.TextMatrix(.Row, 1) = RSSData!Konsumen
.TextMatrix(.Row, 2) = RSSData![Kode Barang]
.TextMatrix(.Row, 3) = RSSData![Nama Barang]
.TextMatrix(.Row, 4) = RSSData!qty
.TextMatrix(.Row, 5) = 0
.TextMatrix(.Row, 6) = RSSData!Harga
.TextMatrix(.Row, 7) = RSSData!Modal
RSSData.MoveNext
Loop
.Rows = VGrid.Rows + 1
.Row = VGrid.Rows - 1
.TextMatrix(.Row, 0) = "Saldo Awal"
.TextMatrix(.Row, 1) = ""
.TextMatrix(.Row, 2) = ""
.TextMatrix(.Row, 3) = ""
.TextMatrix(.Row, 4) = 0
.TextMatrix(.Row, 5) = 0
.TextMatrix(.Row, 6) = 0
.TextMatrix(.Row, 7) = 0
End With

Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Load pada Form bBantu"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo aa
If Not Batal Then
Set RSSData = Nothing
Dim i As Single, aNilai As Currency, aQty As Currency, aHrg
With VGrid
  For i = 1 To .Rows - 1
    If .ValueMatrix(i, 5) <> 0 Then
      aNilai = aNilai + (.ValueMatrix(i, 5) * .ValueMatrix(i, 7))
      aHrg = aHrg + (.ValueMatrix(i, 5) * .ValueMatrix(i, 6))
      aQty = aQty + .ValueMatrix(i, 5)
    End If
  Next i
  ReDim BarisGrid(3)
  BarisGrid(0) = aNilai / aQty
  BarisGrid(1) = aQty
  BarisGrid(2) = aHrg / aQty
End With
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Unload pada Form bBantu"
End Sub

Private Sub VGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo aa
  VGrid.TextMatrix(Row, 8) = VGrid.TextMatrix(Row, 5) * VGrid.TextMatrix(Row, 7)
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur VGrid_AfterEdit pada Form bBantu"
End Sub

Private Sub VGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> 5 Then Cancel = True
If Col = 7 And VGrid.TextMatrix(Row, 0) = "Saldo Awal" Then Cancel = False
If Col = 6 And VGrid.TextMatrix(Row, 0) = "Saldo Awal" Then Cancel = False
End Sub
