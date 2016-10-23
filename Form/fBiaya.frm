VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fBiaya 
   Caption         =   "Input Biaya"
   ClientHeight    =   9975
   ClientLeft      =   600
   ClientTop       =   1065
   ClientWidth     =   13800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   13800
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid80.TDBGrid Grid 
      Height          =   8820
      Left            =   450
      TabIndex        =   1
      Top             =   630
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   15558
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   1
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      LayoutName      =   "aa"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DataView        =   2
      GroupByCaption  =   ""
      DeadAreaBackColor=   16777215
      RowDividerColor =   12632256
      RowSubDividerColor=   16777215
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HBC1616&"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H98CBCA&,.fgcolor=&H0&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1,.appearance=1,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(43)  =   ":id=34,.strikethrough=0,.charset=0"
      _StyleDefs(44)  =   ":id=34,.fontname=MS Sans Serif"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HF7F2C1&,.fgcolor=&H8000&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HE3F1DC&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
      _StyleDefs(61)  =   "Named:id=0:"
      _StyleDefs(62)  =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(63)  =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(64)  =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(65)  =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(66)  =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(67)  =   ":id=0,.fontname=MS Sans Serif"
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biaya"
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
      TabIndex        =   0
      Top             =   105
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F4=Input, F5=Sort, F6=Filter, F7=Form View, F8=Export, F9=Refresh, F10=Search, Alt+X=Close"
      Height          =   315
      Left            =   270
      TabIndex        =   2
      Top             =   9675
      Width           =   7875
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   9390
      Left            =   270
      Top             =   255
      Width           =   14490
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   9870
      Left            =   135
      Top             =   90
      Width           =   14775
   End
End
Attribute VB_Name = "fBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aSet As New ADODB.Recordset
Dim rsPdukung1 As New ADODB.Recordset, rsPdukung2 As New ADODB.Recordset

Private Sub IkatData()
On Error GoTo aa
Set aSet = aData.AmbilData("SELECT format(Tgl, 'yyyy MMM') AS Periode, " & _
"Tgl, Rincian, Jumlah, intNo FROM Biaya;")
Grid.DataSource = aSet
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub IkatGrid()
On Error GoTo aa
Grid.Columns(0).Width = 2000
Grid.Columns(0).Locked = True
Grid.Columns(1).Width = 2000
Grid.Columns(1).Alignment = dbgLeft
Grid.Columns(1).DefaultValue = Date

Grid.Columns(2).Width = 5000
Grid.Columns(3).Width = 3000
Grid.Columns(3).NumberFormat = "Standard"
Grid.Columns(4).Visible = False
Grid.HeadLines = 2.5
Grid.Font.Bold = True
Grid.Font.Size = 10

Grid.Columns(0).FilterText = Format(Date, "yyyy MMM")
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo aa
Dim tmp As String
Select Case KeyCode

Case vbKeyF5  '############ SORTING ############
    If Shift = vbShiftMask Then
    tmp = " desc"
    Else
    tmp = " asc"
    End If
aSet.Sort = "[" & aSet.Fields(Grid.Col).Name & "]" & tmp
Grid.ReBind
Call IkatGrid

Case vbKeyF6  '############ FILTER ############
Call FilterOi(KeyCode, Shift)
Call IkatGrid

Case vbKeyF7  '############ FORM VIEW ############
If Grid.Splits.Count > 1 Then Exit Sub
If Shift = vbShiftMask Then
Grid.DataView = dbgFormView
Else
Grid.DataView = dbgGroupView
End If

Case vbKeyF10  '############ CARI DATA ############
Call CariOi(KeyCode, Shift)
Call IkatGrid

Case vbKeyF9 '############ REFRESH DATA ############
Call IkatData
Call IkatGrid

Case vbKeyX And Shift = 4 '############ TUTUP FORM ############
Unload Me

Case vbKeyF8 '############ CETAK ############
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    CD1.DialogTitle = "Masukkan nama file untuk diexport ke Html(Excel Compatible)"
    CD1.Filter = "Html File|*.htm;*.html"
    CD1.ShowSave
    If Err.Number = 0 Then
    Grid.ExportToFile CD1.FileName, False
    End If
    On Error GoTo aa
    Screen.MousePointer = vbNormal
End Select

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
On Error GoTo aa
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

Private Sub FilterOi(a As Integer, b As Integer)
Dim Col, tmp As String, n As Single
On Error GoTo aa
If b = vbShiftMask Then
For Each Col In Grid.Columns
        Col.FilterText = ""
Next Col
tmp = ""
Else
    For Each Col In Grid.Columns
        If Trim(Col.FilterText) <> "" Then
        n = n + 1
         If n > 1 Then
         tmp = tmp & " AND "
         End If
        tmp = tmp & "[" & Col.DataField & "]" & IIf(IsNumeric(Col.FilterText), "=", " LIKE '") & Col.FilterText & IIf(IsNumeric(Col.FilterText), "", "*'")
        End If
    Next Col
End If
aSet.Filter = tmp
Grid.ReBind
Exit Sub
aa:
MsgBox "Kolom ini tidak bisa difilter.."
Grid.SetFocus
End Sub

Private Sub CariOi(a As Integer, b As Integer)
On Error GoTo aa
Dim Col, tmp As String, n As Single
If "'" & Grid.Columns(Grid.Col).FilterText & "*'" = "'*'" Then Exit Sub
   tmp = "[" & Grid.Columns(Grid.Col).DataField & "]" & _
   IIf(IsNumeric(Grid.Columns(Grid.Col).FilterText), "=", " LIKE '") & _
   Grid.Columns(Grid.Col).FilterText & _
   IIf(IsNumeric(Grid.Columns(Grid.Col).FilterText), "", "*'")
If b = vbShiftMask Then
    If aSet.EOF Then
    aSet.MoveFirst
    End If
aSet.MoveNext
End If
aSet.Find tmp

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub
