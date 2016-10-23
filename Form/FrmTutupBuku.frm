VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTutupBuku 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutup Buku"
   ClientHeight    =   4305
   ClientLeft      =   3990
   ClientTop       =   3405
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Proses Tutup Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   495
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3465
      Width           =   4815
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   315
      Left            =   630
      TabIndex        =   2
      Top             =   2880
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin TDBDate6Ctl.TDBDate D1 
      Height          =   405
      Left            =   2070
      TabIndex        =   6
      Top             =   1935
      Width           =   2595
      _Version        =   65536
      _ExtentX        =   4577
      _ExtentY        =   714
      Calendar        =   "FrmTutupBuku.frx":0000
      Caption         =   "FrmTutupBuku.frx":0136
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTutupBuku.frx":01A4
      Keys            =   "FrmTutupBuku.frx":01C2
      Spin            =   "FrmTutupBuku.frx":0220
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sampai Tanggal :"
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
      Left            =   570
      TabIndex        =   5
      Top             =   2010
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   630
      TabIndex        =   4
      Top             =   2655
      Width           =   4455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   495
      TabIndex        =   3
      Top             =   2475
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proses ini akan menghapus semua transaksi pembelian dan penjualan beserta riwayat hutang piutang yang ada."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   5445
   End
End
Attribute VB_Name = "FrmTutupBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo aa
Dim DataBaru As New Data, Hasil As String
Command1.Enabled = False
DoEvents
Hasil = DataBaru.TutupBuku
If Hasil = "" Then
MsgBox "Proses telah selesai.." & _
vbCrLf & "Sekarang data transaksi lama anda yang telah selesai tidak ada lagi..", vbInformation, "Tutup Buku"
Else
MsgBox Hasil, vbInformation, "Tutup Buku"
End If
Me.Enabled = True
Unload Me
Exit Sub
aa:
'MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_Load()
On Error GoTo aa
Label1.Caption = "Proses ini akan menghapus semua " & _
"transaksi pembelian dan penjualan beserta riwayat " & _
"hutang piutang yang ada." & _
vbCrLf & "Harap anda melakukan proses backup terlebih dahulu " & _
"dan pastikan semua user tidak sedang melakukan transaksi apapun"
D1.Value = Date
Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

