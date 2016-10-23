VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form fMataKurs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kurs Mata Uang"
   ClientHeight    =   5235
   ClientLeft      =   2700
   ClientTop       =   2685
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   435
      Left            =   6900
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin TDBDate6Ctl.TDBDate T1 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   660
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   556
      Calendar        =   "fKurs.frx":0000
      Caption         =   "fKurs.frx":0136
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "fKurs.frx":01A4
      Keys            =   "fKurs.frx":01C2
      Spin            =   "fKurs.frx":0220
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
   Begin VSFlex8Ctl.VSFlexGrid Master 
      Height          =   3360
      Left            =   360
      TabIndex        =   5
      Top             =   1020
      Width           =   7890
      _cx             =   13917
      _cy             =   5927
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"fKurs.frx":0248
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal : "
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kurs Mata Uang"
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
      Left            =   135
      TabIndex        =   1
      Top             =   105
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F4=Input, F5=Sort, F6=Filter, F7=Form View, F9=Refresh, F10=Search, Alt+X=Close"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   4620
      Width           =   7875
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   4740
      Left            =   240
      Top             =   210
      Width           =   8160
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   4935
      Left            =   135
      Top             =   105
      Width           =   8370
   End
End
Attribute VB_Name = "fMataKurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aSet As New ADODB.Recordset
Dim rsPdukung1 As New ADODB.Recordset, rsPdukung2 As New ADODB.Recordset

Private Sub IkatData()
On Error GoTo aa
Set aSet = aData.AmbilData("SELECT * FROM MataKurs where tgl=#" & Format(AmanTgl(T1.Value), "mm/dd/yyyy") & "#")

Dim rsSem As New ADODB.Recordset
Set rsSem = aData.AmbilCommand("Select * from MataUang")
'bBData = bBData & "#" & rsSem![Kode Mata Uang] & ";" & rsSem![Kode Mata Uang] & "(" & rsSem![Nama Mata Uang] & ")|"

Master.Rows = 1
Do While Not rsSem.EOF
Master.Rows = Master.Rows + 1
Master.TextMatrix(Master.Rows - 1, 0) = rsSem![Kode Mata Uang] & " (" & rsSem![Nama Mata Uang] & ")"
Master.TextMatrix(Master.Rows - 1, 2) = rsSem![Kode Mata Uang]
If Not aSet.EOF Then aSet.MoveFirst
aSet.Find "KodeMata='" & rsSem![Kode Mata Uang] & "'"
Dim kNilai As Currency
If aSet.EOF Then
kNilai = 1
Else
kNilai = aSet![Nilai Kurs]
End If
Master.TextMatrix(Master.Rows - 1, 1) = kNilai
rsSem.MoveNext
Loop

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub IkatGrid()
On Error GoTo aa
'Grid.Columns(0).Width = 3200
'Grid.Columns(1).Width = 6500
'Grid.Columns(5).Alignment = dbgRight
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Command1_Click()
Dim kHasil As String
kHasil = aData.SimpanKurs(Master, AmanTgl(T1.Value))
If kHasil = "" Then
MsgBox "Data telah disimpan..", vbInformation, "Simpan Kurs (" & T1.Text & ")"
Else
MsgBox kHasil, vbInformation, "Simpan Kurs (" & T1.Text & ")"
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
Dim tmp As String
Select Case KeyCode

Case vbKeyF5  '############ SORTING ############

Case vbKeyF6  '############ FILTER ############

Case vbKeyF7  '############ FORM VIEW ############

Case vbKeyF10  '############ CARI DATA ############

Case vbKeyF9 '############ REFRESH DATA ############
Call IkatData
Call IkatGrid

Case vbKeyX And Shift = 4 '############ TUTUP FORM ############
Unload Me

Screen.MousePointer = vbNormal
End Select

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
On Error GoTo aa
Dim XaX As String
T1.Value = Date
Call IkatData
Call IkatGrid
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set aData = Nothing
Set aSet = Nothing
End Sub

Private Sub T1_Validate(Cancel As Boolean)
Call IkatData
Call IkatGrid
End Sub
