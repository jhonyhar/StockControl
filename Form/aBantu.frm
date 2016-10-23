VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form aBantu 
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
      TabIndex        =   1
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
      FormatString    =   $"aBantu.frx":0000
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
Attribute VB_Name = "aBantu"
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
VGrid.Top = 10
VGrid.Left = 10

Penanda = True
Batal = False

VGrid.Visible = True
Me.Height = VGrid.Height + 20
Me.Width = VGrid.Width + 20

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
