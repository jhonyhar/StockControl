VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form fDiskon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diskon (dalam %)"
   ClientHeight    =   5055
   ClientLeft      =   4290
   ClientTop       =   3945
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6255
   Begin VB.CheckBox Check1 
      Caption         =   "Pembulatan"
      Height          =   240
      Left            =   225
      TabIndex        =   5
      Top             =   3870
      Width           =   2265
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Batal"
      Height          =   510
      Left            =   5010
      TabIndex        =   2
      Top             =   3870
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   3960
      TabIndex        =   1
      Top             =   3870
      Width           =   1005
   End
   Begin VSFlex8Ctl.VSFlexGrid VGrid 
      Height          =   3585
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   5790
      _cx             =   10213
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
      Rows            =   11
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"fDiskon.frx":0000
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
      Height          =   435
      Left            =   225
      TabIndex        =   4
      Top             =   4500
      Width           =   5745
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ins=Tambah Grid   Del=Hapus Grid"
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
      Left            =   225
      TabIndex        =   3
      Top             =   4230
      Width           =   5790
   End
End
Attribute VB_Name = "fDiskon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Batal As Boolean, NilaiDiskon As Currency

Private Sub Check1_Click()
On Error GoTo aa
Call OO
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Check1_Click pada Form fDiskon"
End Sub

Private Sub Command1_Click()
On Error GoTo aa
Unload Me
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command1_Click pada Form fDiskon"
End Sub

Private Sub Command2_Click()
On Error GoTo aa
Unload Me
Batal = True
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Command2_Click pada Form fDiskon"
End Sub

Private Sub Form_Load()
On Error GoTo aa
Batal = False
NilaiDiskon = 0
Check1.Value = vbChecked
Call AturGrid
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Load pada Form fDiskon"
End Sub

Private Sub AturGrid()
Dim i As Byte
On Error GoTo aa
VGrid.RowHeight(0) = 450
For i = 1 To VGrid.Rows - 1
VGrid.TextMatrix(i, 0) = "Diskon ke-" & i
Next i
Call BackWarna
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur AturGrid pada Form fDiskon"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo aa
 If Not Batal Then Call OO
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Unload pada Form fDiskon"
End Sub

Private Sub VGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Call OO
End Sub

Private Sub VGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
On Error Resume Next
If KeyCode = vbKeyInsert Then
VGrid.Rows = VGrid.Rows + 1
Call AturGrid
VGrid.Row = VGrid.Rows - 1
VGrid.Col = 1
VGrid.ShowCell VGrid.Rows - 1, 1
End If
If KeyCode = vbKeyDelete And VGrid.Rows <> 1 Then
VGrid.RemoveItem VGrid.Row
Call AturGrid
End If
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur VGrid_KeyDown pada Form fDiskon"
End Sub


Private Sub BackWarna()
On Error GoTo aa
VGrid.Select 0, 0, VGrid.Rows - 1, 0
VGrid.FillStyle = flexFillRepeat
VGrid.CellBackColor = VGrid.BackColorFixed
VGrid.FillStyle = flexFillSingle
VGrid.Select 1, 1
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur BackWarna pada Form fDiskon"
End Sub


Private Sub OO()
Dim i As Single, a As Currency, b As Double
a = 1
b = 1
For i = 1 To VGrid.Rows - 1
  If VGrid.ValueMatrix(i, 1) <> 0 Then
  a = a - (VGrid.ValueMatrix(i, 1) / 100 * a)
  b = b - (VGrid.ValueMatrix(i, 1) / 100 * b)
  End If
Next i
If Check1.Value = vbChecked Then
NilaiDiskon = (1 - a) * 100
Else
NilaiDiskon = (1 - b) * 100
End If
Label1.Caption = "Nilai diskon : " & NilaiDiskon & "%"
End Sub


Private Sub VGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then Cancel = True
End Sub
