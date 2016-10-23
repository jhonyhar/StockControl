VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSetPass 
   Caption         =   "Setting Password"
   ClientHeight    =   10005
   ClientLeft      =   1140
   ClientTop       =   1920
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   675
      TabIndex        =   2
      Top             =   1440
      Width           =   6585
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   5040
         TabIndex        =   20
         Top             =   2325
         Width           =   1230
      End
      Begin VB.TextBox t2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   19
         Top             =   2925
         Width           =   2385
      End
      Begin VB.TextBox t2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   17
         Top             =   2625
         Width           =   2385
      End
      Begin VB.TextBox t2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   15
         Top             =   2325
         Width           =   2385
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   6
         Top             =   735
         Width           =   2385
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   8
         Top             =   1035
         Width           =   2385
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   10
         Top             =   1335
         Width           =   2385
      End
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   5040
         TabIndex        =   11
         Top             =   735
         Width           =   1230
      End
      Begin VB.TextBox t3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   24
         Top             =   3975
         Width           =   2385
      End
      Begin VB.TextBox t3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   26
         Top             =   4275
         Width           =   2385
      End
      Begin VB.TextBox t3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   28
         Top             =   4575
         Width           =   2385
      End
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   5040
         TabIndex        =   29
         Top             =   3975
         Width           =   1230
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Password Login semua fungsi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   225
         TabIndex        =   13
         Top             =   2025
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   360
         TabIndex        =   16
         Top             =   2700
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ulangi Password Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   360
         TabIndex        =   18
         Top             =   2970
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Lama : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   360
         TabIndex        =   14
         Top             =   2340
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   35
         Left            =   210
         TabIndex        =   12
         Top             =   1905
         Width           =   6195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Password Login semua fungsi + L/R "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   180
         TabIndex        =   4
         Top             =   465
         Width           =   4635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Lama : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   360
         TabIndex        =   5
         Top             =   735
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ulangi Password Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   360
         TabIndex        =   9
         Top             =   1410
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   360
         TabIndex        =   7
         Top             =   1095
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   30
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   6195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Password Masuk Program Client : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   180
         TabIndex        =   22
         Top             =   3705
         Width           =   4635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Lama : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   360
         TabIndex        =   23
         Top             =   3975
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ulangi Password Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   360
         TabIndex        =   27
         Top             =   4650
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   360
         TabIndex        =   25
         Top             =   4335
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   20
         Left            =   210
         TabIndex        =   21
         Top             =   3555
         Width           =   6195
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   675
      TabIndex        =   30
      Top             =   1440
      Width           =   6585
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   5040
         TabIndex        =   57
         Top             =   3975
         Width           =   1230
      End
      Begin VB.TextBox t6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   56
         Top             =   4575
         Width           =   2385
      End
      Begin VB.TextBox t6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   54
         Top             =   4275
         Width           =   2385
      End
      Begin VB.TextBox t6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   52
         Top             =   3975
         Width           =   2385
      End
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   5040
         TabIndex        =   39
         Top             =   735
         Width           =   1230
      End
      Begin VB.TextBox t4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   38
         Top             =   1335
         Width           =   2385
      End
      Begin VB.TextBox t4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   36
         Top             =   1035
         Width           =   2385
      End
      Begin VB.TextBox t4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   34
         Top             =   735
         Width           =   2385
      End
      Begin VB.TextBox t5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   43
         Top             =   2325
         Width           =   2385
      End
      Begin VB.TextBox t5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   45
         Top             =   2625
         Width           =   2385
      End
      Begin VB.TextBox t5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   47
         Top             =   2925
         Width           =   2385
      End
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   5040
         TabIndex        =   48
         Top             =   2325
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   360
         TabIndex        =   53
         Top             =   4335
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ulangi Password Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   360
         TabIndex        =   55
         Top             =   4650
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Lama : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   360
         TabIndex        =   51
         Top             =   3975
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Password Pakai Kembali No Faktur yang dihapus : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   180
         TabIndex        =   50
         Top             =   3705
         Width           =   4635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   1095
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ulangi Password Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   37
         Top             =   1410
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Lama : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   33
         Top             =   735
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Password Edit Customer, Credit Limit Lewat : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   225
         TabIndex        =   32
         Top             =   465
         Width           =   4635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Lama : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   42
         Top             =   2340
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ulangi Password Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   46
         Top             =   2970
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Baru : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   44
         Top             =   2700
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Password Penjualan Rugi : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   225
         TabIndex        =   41
         Top             =   2025
         Width           =   3375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   14
         Left            =   210
         TabIndex        =   31
         Top             =   315
         Width           =   6195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   4
         Left            =   210
         TabIndex        =   40
         Top             =   1905
         Width           =   6195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   29
         Left            =   210
         TabIndex        =   49
         Top             =   3555
         Width           =   6195
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6225
      Left            =   450
      TabIndex        =   1
      Top             =   945
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   10980
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Login   "
            Key             =   "Log"
            Object.ToolTipText     =   "Setting password untuk login ke aplikasi"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Function   "
            Key             =   "func"
            Object.ToolTipText     =   "Setting password yang berhubungan dengan fungsi aplikasi"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setting Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   7650
      Left            =   225
      Top             =   210
      Width           =   7665
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   7350
      Left            =   315
      Top             =   360
      Width           =   7395
   End
End
Attribute VB_Name = "FrmSetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSimpan_Click(Index As Integer)
Dim xLama As String
On Error GoTo aa

Dim rsPass As New ADODB.Recordset
zas:
Set rsPass = aData.AmbilCommand("select * from xPass")
If rsPass.EOF Then
aData.ExecCommand ("insert into xPass values('','','','','','')")
GoTo zas
End If

Select Case Index
Case 0    'Masuk All +LR
If t1(1).Text <> t1(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t1(1).SetFocus
  Exit Sub
End If
If Not (IsNull(rsPass!PassInAllLR) Or rsPass!PassInAllLR = "") Then
xLama = rsPass!PassInAllLR
  If xLama <> RegC.HashString(t1(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t1(0).SetFocus
  Exit Sub
  End If
End If
aData.ExecCommand "update xpass set PassInAllLR='" & RegC.HashString(t1(1)) & "'"
MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
t1(0).Text = "": t1(1).Text = "": t1(2).Text = ""

Case 1    'Masuk All
If t2(1).Text <> t2(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t2(1).SetFocus
  Exit Sub
End If
If Not (IsNull(rsPass!PassInAll) Or rsPass!PassInAll = "") Then
xLama = rsPass!PassInAll
  If xLama <> RegC.HashString(t2(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t2(0).SetFocus
  Exit Sub
  End If
End If
aData.ExecCommand "update xpass set PassInAll='" & RegC.HashString(t2(1)) & "'"
MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
t2(0).Text = "": t2(1).Text = "": t2(2).Text = ""

Case 2    'Masuk Client
If t3(1).Text <> t3(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t3(1).SetFocus
  Exit Sub
End If
If Not (IsNull(rsPass!PassIn) Or rsPass!PassIn = "") Then
xLama = rsPass!PassIn
  If xLama <> RegC.HashString(t3(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t3(0).SetFocus
  Exit Sub
  End If
End If
aData.ExecCommand "update xpass set PassIn='" & RegC.HashString(t3(1)) & "'"
MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
t3(0).Text = "": t3(1).Text = "": t3(2).Text = ""

Case 3    'Customer & Credit Limit
If t4(1).Text <> t4(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t4(1).SetFocus
  Exit Sub
End If
If Not (IsNull(rsPass!PassCust) Or rsPass!PassCust = "") Then
xLama = rsPass!PassCust
  If xLama <> RegC.HashString(t4(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t4(0).SetFocus
  Exit Sub
  End If
End If
aData.ExecCommand "update xpass set PassCust='" & RegC.HashString(t4(1)) & "'"
MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
t4(0).Text = "": t4(1).Text = "": t4(2).Text = ""

Case 4    'Penjualan Rugi
If t5(1).Text <> t5(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t5(1).SetFocus
  Exit Sub
End If
If Not (IsNull(rsPass!PassRugi) Or rsPass!PassRugi = "") Then
xLama = rsPass!PassRugi
  If xLama <> RegC.HashString(t5(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t5(0).SetFocus
  Exit Sub
  End If
End If
aData.ExecCommand "update xpass set PassRugi='" & RegC.HashString(t5(1)) & "'"
MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
t5(0).Text = "": t5(1).Text = "": t5(2).Text = ""

Case 5    'No Faktur Hapus
If t6(1).Text <> t6(2).Text Then
  MsgBox "Password baru yang anda masukkan tidak sesuai..", vbInformation, "Ubah Password"
  t6(1).SetFocus
  Exit Sub
End If
If Not (IsNull(rsPass!PassFaktur) Or rsPass!PassFaktur = "") Then
xLama = rsPass!PassFaktur
  If xLama <> RegC.HashString(t6(0).Text) Then
  MsgBox "Password lama yang anda masukkan salah..", vbInformation, "Ubah Password"
  t6(0).SetFocus
  Exit Sub
  End If
End If
aData.ExecCommand "update xpass set PassFaktur='" & RegC.HashString(t6(1)) & "'"
MsgBox "Password baru anda telah disimpan..", vbInformation, "Ubah Password"
t6(0).Text = "": t6(1).Text = "": t6(2).Text = ""
End Select

Exit Sub
aa:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

Private Sub Form_Load()
  Call TabStrip1_Click
End Sub

Private Sub TabStrip1_Click()
  If TabStrip1.SelectedItem.Index = 1 Then
    Frame2.Visible = True
    Frame1.Visible = False
  Else
    Frame1.Visible = True
    Frame2.Visible = False
  End If
End Sub
