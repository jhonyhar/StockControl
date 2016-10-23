VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   Caption         =   "Laporan"
   ClientHeight    =   7845
   ClientLeft      =   2235
   ClientTop       =   1935
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   8790
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":09A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1258
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1574
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Printer Setup"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export to Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export to Acrobat PDF"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export to HTML"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export"
         EndProperty
      EndProperty
   End
   Begin CRVIEWER9LibCtl.CRViewer9 aView 
      Height          =   7005
      Left            =   165
      TabIndex        =   0
      Top             =   495
      Width           =   5805
      lastProp        =   600
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.Menu mnuAA 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAAPmbyran 
         Caption         =   "&Pembayaran"
      End
      Begin VB.Menu mnuAAFaktur 
         Caption         =   "&Faktur"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Report As CRAXDdRT.Report
Dim Lontong As String, JenisLontong As String
Private Sub aView_OnLaunchHyperlink(Hyperlink As String, UseDefault As Boolean)
On Error GoTo aa
If Left(Hyperlink, 1) = "#" Then
  Lontong = Right(Hyperlink, Len(Hyperlink) - 3)
  JenisLontong = Mid(Hyperlink, 2, 1)
  PopupMenu mnuAA
Else
  Call AturLaporan("FakturPrint", , Hyperlink)
End If
Exit Sub
aa:
Dim Err_Setering As String
Err_Setering = "Error:" & Err.Number & " => " & Err.Description & vbCrLf & "Di prosedur aView_OnLaunchHyperlink pada " & "Form Form2 di baris " & Erl
Select Case MsgBox(Err_Setering, vbRetryCancel, App.Title & "-Form2 Error")
  Case vbCancel: Resume Exit_aView_OnLaunchHyperlink:
  Case vbRetry: Resume
  Case Else: End
End Select
Exit_aView_OnLaunchHyperlink:
App.LogEvent "myAS=>" & Format(Date, "dd-mmmm-yyyy") & Format(Time, "(hh:mm:ss)") & _
vbCrLf & Err_Setering & vbCrLf, vbLogEventTypeError

End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'  Case vbKeyLeft
'  aView.ShowNextPage
'  Case vbKeyRight
'  aView.ShowPreviousPage
'  End Select
'End Sub

Private Sub Form_Load()
On Error GoTo aa
  aView.EnableSearchExpertButton = True
  aView.EnableSelectExpertButton = True
  aView.EnableSearchControl = True
Exit Sub
aa:
MsgBox "Error:" & Err.Number & " (" & Err.Description & ") di prosedur Form_Load pada Form Form2"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Dim iTop As Integer
    Dim iAdjustment As Integer
        iTop = Toolbar.Height
        iAdjustment = Toolbar.Height
    aView.Top = iTop
    aView.Left = 0
    aView.Height = Me.Height - iAdjustment - 500
    aView.Width = Me.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Report = Nothing
End Sub

Private Sub mnuAAFaktur_Click()
  Call AturLaporan("FakturPrint", , Lontong)
End Sub

Private Sub mnuAAPmbyran_Click()
  Call AturLaporan("HPDaftarAnak", JenisLontong, Lontong)
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo aa
Dim NamaFile As String
    Select Case Button.Index
    Case 2 'print Setup
    Me.Report.PrinterSetup Me.hWnd
    aView.Refresh
    Case 3
    Me.Report.ExportOptions.Reset
    Me.Report.ExportOptions.FormatType = crEFTExcel80Tabular
    Me.Report.ExportOptions.DestinationType = crEDTDiskFile
    NamaFile = InputBox("Masukkan nama file untuk laporan yg diexport ke Excel", "Excel Export", Format(Date, "ddmmmmyyyy"))
    Me.Report.ExportOptions.DiskFileName = App.Path & "\export\" & NamaFile & ".xls"
    Me.Report.Export False
    Case 4
    Me.Report.ExportOptions.Reset
    Me.Report.ExportOptions.FormatType = crEFTPortableDocFormat
    Me.Report.ExportOptions.DestinationType = crEDTDiskFile
    Me.Report.ExportOptions.PDFExportAllPages = True
    NamaFile = InputBox("Masukkan nama file untuk laporan yg diexport ke PDF(Acrobat)", "PDF Export", Format(Date, "ddmmmmyyyy"))
    Me.Report.ExportOptions.DiskFileName = App.Path & "\export\" & NamaFile & ".pdf"
    Me.Report.Export False
    Case 5
    Me.Report.ExportOptions.Reset
    Me.Report.ExportOptions.FormatType = crEFTHTML40
    Me.Report.ExportOptions.DestinationType = crEDTDiskFile
    NamaFile = InputBox("Masukkan nama file untuk laporan yg diexport ke Format Html", "HTML Export", Format(Date, "ddmmmmyyyy"))
    Me.Report.ExportOptions.DiskFileName = App.Path & "\export\" & NamaFile & ".htm"
    Me.Report.ExportOptions.HTMLEnableSeparatedPages = True
    Me.Report.ExportOptions.HTMLHasPageNavigator = True
    Me.Report.ExportOptions.HTMLFileName = App.Path & "\export\" & NamaFile & ".htm"
    Me.Report.Export False
    Case 6
    Me.Report.ExportOptions.Reset
    Me.Report.ExportOptions.PromptForExportOptions
    If Me.Report.ExportOptions.FormatType = 0 Then Me.Report.ExportOptions.FormatType = crEFTExcel80Tabular
    Me.Report.ExportOptions.PDFExportAllPages = True
    If Me.Report.ExportOptions.FormatType <> crEFTCrystalReport Then Me.Report.Export False
    End Select
    
    Exit Sub
aa:
MsgBox Err.Source & Err.Number & vbCrLf & Err.Description, vbInformation, Me.Caption & "#Error"
End Sub

