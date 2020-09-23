VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2205
   ClientLeft      =   2895
   ClientTop       =   3015
   ClientWidth     =   7125
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAbout_Home.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout_Home.frx":000C
   ScaleHeight     =   2205
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCompanyName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   390
      UseMnemonic     =   0   'False
      Width           =   5130
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout_Home.frx":11D6
      ForeColor       =   &H00C0C0FF&
      Height          =   1110
      Left            =   2040
      TabIndex        =   3
      Top             =   1380
      UseMnemonic     =   0   'False
      Width           =   4965
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   2040
      TabIndex        =   2
      Top             =   645
      UseMnemonic     =   0   'False
      Width           =   5100
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   5085
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail ME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "morganh@hartcom.net"
      Top             =   1965
      Width           =   855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA& = 48
Private Type Rect
     left As Long
     top As Long
     right As Long
     bottom As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public PreventClose As Boolean
Public AlwaysOnTop As Boolean
Public SleepTime As Integer
Private Sub CenterForm()
  Dim Rc As Rect
  Dim rVal As Long
  Dim T As Long, B As Long, L As Long, r As Long
  Dim mT As Long, mL As Long
     
    rVal = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
    
    T = Rc.top * Screen.TwipsPerPixelY
    B = Rc.bottom * Screen.TwipsPerPixelY
    L = Rc.left * Screen.TwipsPerPixelX
    r = Rc.right * Screen.TwipsPerPixelX
    
    mT = Abs((B / 2.8) - (Me.Height / 2))
    mL = Abs((r / 2) - (Me.Width / 2))
    
    If mT < T Then mT = T
    If mT > B - Me.Height Then mT = B - Me.Height
    If mL < L Then mL = L
    
    Me.Move mL, mT

End Sub

Private Sub Form_Click()
    If Not PreventClose Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Call CenterForm
    
    On Error Resume Next
    lblTitle.Caption = App.ProductName
    lblCompanyName.Caption = "MorganWare™" 'App.CompanyName
    
    lblVersion.Caption = "By: Morgan Haueisen" & vbCrLf & _
        "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
        App.LegalCopyright
    
    Show
    DoEvents
    
    If AlwaysOnTop Then Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    If SleepTime > 0 Then Sleep SleepTime

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.Font.Underline = False
    'lblWebSite.Font.Underline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub

Private Sub lblDisclaimer_Click()
    If Not PreventClose Then Unload Me
End Sub

Private Sub lblEMail_Click()
    ShellExecute Me.hWnd, "open", "mailto:" & lblEMail.ToolTipText & _
                            "?subject=" & App.ProductName, vbNullString, "C:\", 5
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.Font.Underline = True
End Sub

Private Sub lblWebSite_Click()
    'ShellExecute Me.hWnd, "open", "http://ebrain.8m.net/", vbNullString, "C:\", 5
End Sub

Private Sub lblWebSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lblWebSite.Font.Underline = True
End Sub

Private Sub lblTitle_Click()
    If Not PreventClose Then Unload Me
End Sub

Private Sub lblVersion_Click()
    If Not PreventClose Then Unload Me
End Sub


