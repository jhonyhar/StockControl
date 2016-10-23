VERSION 5.00
Begin VB.Form frmParent 
   Caption         =   "frmParent"
   ClientHeight    =   3240
   ClientLeft      =   2820
   ClientTop       =   2370
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4635
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2205
      Left            =   90
      ScaleHeight     =   2205
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   -30
      Width           =   4515
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "This is the container of the MDI"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   2220
      End
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   180
      Top             =   2280
      _ExtentX        =   1270
      _ExtentY        =   1270
      ShowSysCommands =   0
      ShowInactiveState=   0   'False
      Skin            =   98
      ChangeSkinButton=   0   'False
      SkinPicture     =   "frmParent.frx":0000
      LcK1            =   ".42(-,+(423(+,-+,"
      AmbientB        =   $"frmParent.frx":7F12
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pLoaded As Boolean
Private WithEvents pFormMDI As MDIForm
Attribute pFormMDI.VB_VarHelpID = -1
Private pFormRegion As Long
Private pFormHwnd As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nindex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nindex As Long, ByVal dwnewlong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Option Explicit

Private Sub Form_Activate()
    pFormMDI.ZOrder
    pFormMDI.SetFocus
End Sub

Private Sub Form_Load()
    Dim lngResult As Long
    
    Picture1.Top = 0
    Picture1.Left = 90
    pFormHwnd = pFormMDI.hwnd
    If pFormMDI.WindowState = 0 Then
        Move pFormMDI.Left, pFormMDI.Top
    End If
    lngResult = GetWindowLong(pFormHwnd, -16)
    SetWindowLong pFormHwnd, -16, lngResult And Not &HC00000
    SetParent pFormHwnd, Me.Picture1.hwnd
    pFormMDI.Move -40, -40, pFormMDI.Width, pFormMDI.Height
    Picture1.Width = pFormMDI.Width - 100
    Picture1.Height = pFormMDI.Height - 80
    Me.Width = Picture1.Width + 100
    Me.Height = Picture1.Height + 200
    Picture1.BackColor = pFormMDI.BackColor
    
    pFormMDI.Show
    pFormRegion = CreateRectRgn(0, 0, 0, 0)
    If pFormMDI.WindowState = 2 Then
        Call SetWindowRgn(hwnd, pFormRegion, True)
    End If
    WindowState = pFormMDI.WindowState
    If pFormMDI.WindowState = 1 Then
        pFormMDI.WindowState = 0
    End If
    Me.Caption = pFormMDI.Caption
    pLoaded = True
End Sub

Private Sub Form_Resize()
    If Not pLoaded Then Exit Sub
    If WindowState = vbMinimized Then Exit Sub
    If pFormMDI.WindowState = 2 Then
        pFormMDI.WindowState = vbNormal
        WindowState = vbMaximized
    End If
    pFormMDI.Move -40, -40, Me.Width - 100, Me.Height - 500
    Picture1.Width = pFormMDI.Width - 100
    Picture1.Height = pFormMDI.Height - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Cancel Then
        If IsWindowVisible(pFormHwnd) <> 0 Then
            Call ShowWindow(pFormHwnd, SW_HIDE)
        End If
        SetParent pFormHwnd, 0&
        Unload pFormMDI
        Set pFormMDI = Nothing
        DeleteObject pFormRegion
    End If
End Sub

Private Sub pFormMDI_Unload(Cancel As Integer)
    If Not Cancel Then
        Unload Me
    End If
End Sub

Public Property Get Form() As MDIForm
    Set Form = pFormMDI
End Property

Public Property Set Form(ByVal nForm As MDIForm)
    Set pFormMDI = nForm
End Property

