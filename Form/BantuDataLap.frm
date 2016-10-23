VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form BantuDataLap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilih Data"
   ClientHeight    =   2925
   ClientLeft      =   5565
   ClientTop       =   2580
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   465
      Left            =   3750
      TabIndex        =   6
      Top             =   2280
      Width           =   1470
   End
   Begin VB.Frame FrmTgl 
      Caption         =   "Tanggal : "
      Height          =   2055
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   5025
      Begin VB.CheckBox Check1 
         Caption         =   "&Semua Data"
         Height          =   390
         Left            =   180
         TabIndex        =   5
         Top             =   1575
         Width           =   2730
      End
      Begin TDBDate6Ctl.TDBDate t2 
         Height          =   405
         Left            =   1110
         TabIndex        =   4
         Top             =   795
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   714
         Calendar        =   "BantuDataLap.frx":0000
         Caption         =   "BantuDataLap.frx":0136
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BantuDataLap.frx":01A4
         Keys            =   "BantuDataLap.frx":01C2
         Spin            =   "BantuDataLap.frx":0220
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
      Begin TDBDate6Ctl.TDBDate t1 
         Height          =   405
         Left            =   1110
         TabIndex        =   2
         Top             =   345
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   714
         Calendar        =   "BantuDataLap.frx":0248
         Caption         =   "BantuDataLap.frx":037E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BantuDataLap.frx":03EC
         Keys            =   "BantuDataLap.frx":040A
         Spin            =   "BantuDataLap.frx":0468
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
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   4830
         X2              =   180
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Dari :"
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sampai : "
         Height          =   360
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   825
         Width           =   1050
      End
   End
End
Attribute VB_Name = "BantuDataLap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
 If Check1.Value = vbChecked Then
  t1.Enabled = False
  t2.Enabled = False
 Else
  t1.Enabled = True
  t2.Enabled = True
 End If
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
t1.Value = Date - 30
t2.Value = Date
Check1.Value = vbUnchecked
End Sub
