VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E1FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   2730
   ClientTop       =   2580
   ClientWidth     =   5835
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmTimedMsgBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserText 
      Height          =   300
      Left            =   690
      TabIndex        =   6
      Top             =   930
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Help"
      Height          =   360
      Index           =   3
      Left            =   3405
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
      Height          =   360
      Index           =   2
      Left            =   2355
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
      Height          =   360
      Index           =   1
      Left            =   1305
      TabIndex        =   3
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
      Height          =   360
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   5430
      ToolTipText     =   "Close"
      Top             =   45
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   2
      Left            =   5295
      Picture         =   "frmTimedMsgBox.frx":000C
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   1
      Left            =   4935
      Picture         =   "frmTimedMsgBox.frx":0410
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   0
      Left            =   4575
      Picture         =   "frmTimedMsgBox.frx":0837
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label txtMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1020
      TabIndex        =   1
      Top             =   495
      UseMnemonic     =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009EF5F3&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
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
      Height          =   285
      Left            =   375
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 2003-2004
'/*************************************/
'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

'/* Used for Manifest files (Win XP style controls)
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()

'/* Used to keep form always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'/* Used to get screen size
Private Type Rect
     left   As Long
     top    As Long
     right  As Long
     bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA As Long = 48

'/* Used to get positions of cursor
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private CursorXY As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'/* Button and Icon types
Public Enum ShowIconTypes
    None_i = 0
    vbCritical = 16         '/* Display Critical Message icon.
    vbQuestion = 32         '/* Display Warning Query icon.
    vbExclamation = 48      '/* Display Warning Message icon.
    vbInformation = 64      '/* Display Information Message icon.
    WinLogo_i = 128         '/* Display WinLogo icon.
    Folder_i = 144          '/* Display Folder icon.
    Printer_i = 160         '/* Display Printer icon.
    Find_i = 176            '/* Display Find icon.
    Save_i = 240            '/* Display Save icon.
    Hourglass_i = 80        '/* Display Hourglass icon.
    
    vbDefaultButton1 = 0    '/* First button is default.
    vbDefaultButton2 = 256  '/* Second button is default.
    vbDefaultButton3 = 512  '/* Third button is default.
    vbDefaultButton4 = 768  '/* Fourth button is default.
    
    vbOKCancel = 1          '/* Display OK and Cancel buttons.
    vbAbortRetryIgnore = 2  '/* Display Abort, Retry, and Ignore buttons.
    vbYesNoCancel = 3       '/* Display Yes, No, and Cancel buttons.
    vbYesNo = 4             '/* Display Yes and No buttons.
    vbRetryCancel = 5       '/* Display Retry and Cancel buttons.
    vbOkButton = 6          '/* Display OK button only.
    vbMsgBoxHelpButton = 16384 '/* Display the Help button
    
    vbHelp = 8              '/* Help button pressed
End Enum

'/* Used for moving the form around by draging the caption bar
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'/* Used to draw the form's border
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal left As Long, ByVal top As Long, ByVal right As Long, ByVal bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'/* Used to round the corners of the form and make trasnparent
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, ByVal RectY2 As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'/* Used to play system sounds
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Const MB_IconAsterisk    As Long = &H10&
Private Const MB_IconQuestion    As Long = &H20&
Private Const MB_IconExclamation As Long = &H30&
Private Const MB_IconInformation As Long = &H40&

'/* Used to draw system icons
Private Enum SystemIconConstants
    IDI_Application = 32512
    IDI_Error = 32513       'vbCritical (Critical)
    IDI_Question = 32514    'vbQuestion
    IDI_Warning = 32515     'vbExlamation (Exclamation)
    IDI_Information = 32516 'vbInformation (Asterisk)
    IDI_WinLogo = 32517
End Enum
Private Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As SystemIconConstants) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

'/* Used to draw system icons from Shell32.dll
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'/* GradientFill API - Requires Windows 2000 or later; Requires Windows 98 or later
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Type TRIVERTEX
    X     As Long
    Y     As Long
    Red   As Integer '/* Ushort value (-256 to 0)
    Green As Integer '/* Ushort value (-256 to 0)
    Blue  As Integer '/* Ushort value (-256 to 0)
    Alpha As Integer '/* Ushort value (-256 to 0)
End Type
Private Const GRADIENT_FILL_RECT_H   As Long = &H0&
Private Const GRADIENT_FILL_RECT_V   As Long = &H1&
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2&
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" _
    (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, _
     pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

'/* Operating system version information
Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long

'/* Form Variables
Private oStandardIcon     As Long
Private oCaption          As String
Private oAutoCloseSeconds As Long
Private oButtonResponse   As Integer
Private oButtonFocus      As Byte
Private oNonModal         As Boolean
Private oInputBox         As Boolean
Private oCountDown        As Long

Private Sub cmdButton_Click(Index As Integer)
    
    oButtonResponse = cmdButton(Index).Tag
    Me.Hide
    If oNonModal Then Unload Me
    
End Sub

Private Sub Form_Activate()
    
    If cmdButton(oButtonFocus).Visible Then
        cmdButton(oButtonFocus).SetFocus
    End If
    If oInputBox Then txtUserText.SetFocus
    
    Me.ZOrder
    
End Sub


Private Sub Form_Initialize()
    
    '/* Used for Manifest files (Win XP style controls)
    Call InitCommonControls
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgX(0).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMsgBox = Nothing
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then imgClose.Picture = imgX(2).Picture
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgClose.Picture <> imgX(2).Picture Then
        imgClose.Picture = imgX(1).Picture
    End If
End Sub

Private Sub DisplayInputBox(ByVal sPrompt As String, _
                        ByVal sTitle As String, _
                        Optional ByVal sDefault As String = vbNullString, _
                        Optional ByVal bShowClose As Boolean = True, _
                        Optional ByVal bCenter As Boolean = False, _
                        Optional sFont = "Tahoma")
    
  Dim lPosX  As Long
  Dim lPosY  As Long
  Dim lWidth As Long
    
    '/* Set defaults
    On Error Resume Next
        Me.ScaleMode = vbPixels
        Me.DrawWidth = 1
        Me.FillStyle = 1
        Me.Font = sFont
        txtMessage.Font = sFont
        txtMessage.FontSize = 10
        lblTitle.Font = sFont
        imgClose.Picture = imgX(0).Picture
    On Error GoTo 0
    
    '/* Get display position from mouse position
    Call GetCursorPos(CursorXY)
    lPosX = CursorXY.X * Screen.TwipsPerPixelX
    lPosY = CursorXY.Y * Screen.TwipsPerPixelY
    
    oCaption = sTitle
    txtUserText.Text = sDefault
            
    '/* Resize the Form's width to fit the title bar/messagebox width
    Me.FontSize = 10
    lWidth = 5000
    Me.FontSize = 8
    If lWidth < (Me.TextWidth(sPrompt) + 90) * Screen.TwipsPerPixelX Then
        lWidth = (Me.TextWidth(sPrompt) + 90) * Screen.TwipsPerPixelX
    End If
    lblTitle.Caption = sTitle
    
    Me.Width = lWidth
    Me.Height = 800
    
    '/* Resize the Form's height based on the amount of text to display
    txtMessage.Move 8, 40, Me.ScaleWidth - 20, Me.ScaleHeight - 50
    txtMessage.Caption = sPrompt
    If txtMessage.top + txtMessage.Height >= Me.ScaleHeight - 10 Then
        Me.Height = (txtMessage.top + txtMessage.Height + 10) * Screen.TwipsPerPixelY
    End If
    
    txtUserText.Move 25, txtMessage.top + txtMessage.Height + 10, txtMessage.Width - 25
    
    '/* Locate Buttons and resize Form if required
    If Val(cmdButton(0).Tag) > 0 Or Val(cmdButton(3).Tag) > 0 Then
        Dim i As Byte
        '/* How many buttons are visible?
        If Val(cmdButton(1).Tag) > 0 Then i = 1
        If Val(cmdButton(2).Tag) > 0 Then i = 2
        If Val(cmdButton(3).Tag) > 0 Then i = 3
        
        cmdButton(0).top = txtUserText.top + txtUserText.Height + 10
        cmdButton(1).top = txtUserText.top + txtUserText.Height + 10
        cmdButton(2).top = txtUserText.top + txtUserText.Height + 10
        cmdButton(3).top = txtUserText.top + txtUserText.Height + 10
        If Me.Width < (cmdButton(i).left + cmdButton(i).Width + 15) * Screen.TwipsPerPixelX Then
            Me.Width = (cmdButton(i).left + cmdButton(i).Width + 15) * Screen.TwipsPerPixelX
        End If
        
        Me.Height = (cmdButton(0).top + cmdButton(0).Height + 10) * Screen.TwipsPerPixelY
    End If
    
    '/* Show or don't show the close button
    If bShowClose Then
        imgClose.Visible = True
    Else
        imgClose.Visible = False
    End If
    
    '/* Locate title bar and close button
    imgClose.Move (Me.ScaleWidth - imgClose.Width) - 8, 4
    lblTitle.Move 2, 5, Me.ScaleWidth, 25
    
    Call GradientFill
        
    '/* Draw box around Title Bar
    Me.Line (0, 0)-(Me.ScaleWidth, 25), &HB1FFFF, BF
    
    '/* Draw border around the Form
    Me.ForeColor = &H80000015
    RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
    Me.ForeColor = &H8000000F
    RoundRect Me.hdc, 1, 1, (Me.Width / Screen.TwipsPerPixelX) - 2, (Me.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
  
    '/* Make corners transparent
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, _
            (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), _
            25, 25), True

    '/* Position form on screen
    If Not bCenter Then
        Me.Move lPosX, lPosY
    End If
    Call PositionForm(bCenter)
    
    Dim hIcon As Long
    hIcon = LoadStandardIcon(0&, IDI_Question)
    Call DrawIcon(Me.hdc, 4&, 4&, hIcon)
    DestroyIcon hIcon
    
    txtUserText.SelStart = 0
    txtUserText.SelLength = Len(txtUserText.Text)
    
End Sub

Private Sub DisplayMessage(ByVal sText As String, _
                        Optional sIcon As ShowIconTypes = None_i, _
                        Optional ByVal sTitle As String = vbNullString, _
                        Optional lAutoCloseSeconds As Long = 0, _
                        Optional ByVal bShowClose As Boolean = True, _
                        Optional ByVal bCenter As Boolean = False, _
                        Optional lWidth As Long = -1, _
                        Optional sFont = "Tahoma", _
                        Optional OwnerForm As Form)
    

  Dim lPosX As Long
  Dim lPosY As Long
  Dim Shell32Icon As Boolean
    
    '/* Set defaults
    On Error Resume Next
        Me.ScaleMode = vbPixels
        Me.DrawWidth = 1
        Me.FillStyle = 1
        Me.Font = sFont
        txtMessage.Font = sFont
        txtMessage.FontSize = 10
        lblTitle.Font = sFont
        imgClose.Picture = imgX(0).Picture
    On Error GoTo 0
    
    '/* Get display position from mouse position
    Call GetCursorPos(CursorXY)
    lPosX = CursorXY.X * Screen.TwipsPerPixelX
    lPosY = CursorXY.Y * Screen.TwipsPerPixelY
    
    '/* Set Title bar
    Select Case sIcon
        Case vbInformation '/* The "i" icon - Information
            If sTitle = vbNullString Then sTitle = "Information"
            MessageBeep MB_IconInformation
            oStandardIcon = IDI_Information
            
        Case vbCritical '/* The "x" icon - Critical
            If sTitle = vbNullString Then sTitle = "ERROR!"
            MessageBeep MB_IconAsterisk
            oStandardIcon = IDI_Error
            
        Case vbExclamation '/* The "!" icon - Exclamation
            If sTitle = vbNullString Then sTitle = "Warning!"
            MessageBeep MB_IconExclamation
            oStandardIcon = IDI_Warning
            
        Case vbQuestion '/* The "?" icon - Question
            If sTitle = vbNullString Then sTitle = "Question?"
            MessageBeep MB_IconQuestion
            oStandardIcon = IDI_Question
            
        Case WinLogo_i '/* Winlogo icon
            oStandardIcon = IDI_WinLogo
            
        Case Printer_i '/* Printer icon
            If sTitle = vbNullString Then sTitle = "Printing.. Please Wait"
            MessageBeep MB_IconInformation
            oStandardIcon = 16
            Shell32Icon = True
            
        Case Folder_i '/* Open folder icon
            MessageBeep MB_IconInformation
            oStandardIcon = 4
            Shell32Icon = True
        
        Case Find_i '/* Find icon
            MessageBeep MB_IconInformation
            oStandardIcon = 22
            Shell32Icon = True

        Case Save_i '/* Save icon
            MessageBeep MB_IconInformation
            oStandardIcon = 6
            Shell32Icon = True
        
        Case Hourglass_i '/* Hourglass icon
            If sTitle = vbNullString Then sTitle = "Working.. Please Wait"
            MessageBeep MB_IconInformation
            oStandardIcon = 76
            Shell32Icon = True

        Case Else 'Use no icon

    End Select
    oCaption = sTitle
            
    '/* Resize the Form's width to fit the title bar/messagebox width
    Me.FontSize = 10
    If lWidth = -1 Then
        lWidth = (Me.TextWidth(sText) + 20) * Screen.TwipsPerPixelX
        If lWidth > 5000 Then lWidth = 5000
    End If
    If lWidth < 1500 Then lWidth = 1500
    If lAutoCloseSeconds > 0 Then
        If sTitle > vbNullString Then
            sTitle = sTitle & " -" & CStr(lAutoCloseSeconds)
        Else
            sTitle = CStr(lAutoCloseSeconds)
        End If
    End If
    Me.FontSize = 8
    If lWidth < (Me.TextWidth(sTitle) + 90) * Screen.TwipsPerPixelX Then
        lWidth = (Me.TextWidth(sTitle) + 90) * Screen.TwipsPerPixelX
    End If
    lblTitle.Caption = sTitle
    
    Me.Width = lWidth
    Me.Height = 800
    
    '/* Resize the Form's height based on the amount of text to display
    txtMessage.Move 8, 40, Me.ScaleWidth - 20, Me.ScaleHeight - 50
    txtMessage.Caption = sText
    If txtMessage.top + txtMessage.Height >= Me.ScaleHeight - 10 Then
        Me.Height = (txtMessage.top + txtMessage.Height + 10) * Screen.TwipsPerPixelY
    End If
    
    '/* Locate Buttons and resize Form if required
    If Val(cmdButton(0).Tag) > 0 Or Val(cmdButton(3).Tag) > 0 Then
        Dim i As Byte
        '/* How many buttons are visible?
        If Val(cmdButton(1).Tag) > 0 Then i = 1
        If Val(cmdButton(2).Tag) > 0 Then i = 2
        If Val(cmdButton(3).Tag) > 0 Then i = 3
        
        'Me.Height = Me.Height + 500
        cmdButton(0).top = txtMessage.top + txtMessage.Height + 10
        cmdButton(1).top = txtMessage.top + txtMessage.Height + 10
        cmdButton(2).top = txtMessage.top + txtMessage.Height + 10
        cmdButton(3).top = txtMessage.top + txtMessage.Height + 10
        If Me.Width < (cmdButton(i).left + cmdButton(i).Width + 15) * Screen.TwipsPerPixelX Then
            Me.Width = (cmdButton(i).left + cmdButton(i).Width + 15) * Screen.TwipsPerPixelX
        End If
        Me.Height = (cmdButton(0).top + cmdButton(0).Height + 10) * Screen.TwipsPerPixelY 'Me.Height + 500
    End If
    
    '/* Show or don't show the close button
    If bShowClose Then
        imgClose.Visible = True
    Else
        imgClose.Visible = False
    End If
    
    '/* Enable or disable auto close timer
    If lAutoCloseSeconds = 0 Then
        Timer1.Enabled = False
    Else
        If oCaption > vbNullString Then oCaption = oCaption & " -"
        oAutoCloseSeconds = lAutoCloseSeconds
        Timer1.Enabled = True
    End If
    
    Call GradientFill

    '/* Locate title bar and close button
    imgClose.Move (Me.ScaleWidth - imgClose.Width) - 8, 4
    lblTitle.Move 2, 5, Me.ScaleWidth, 25
        
    '/* Draw box around Title Bar
    Me.Line (0, 0)-(Me.ScaleWidth, 25), &HB1FFFF, BF
    
    '/* Draw border around the Form
    Me.ForeColor = &H80000015
    RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
    Me.ForeColor = &H8000000F
    RoundRect Me.hdc, 1, 1, (Me.Width / Screen.TwipsPerPixelX) - 2, (Me.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
  
    '/* Draw Icon
    If Shell32Icon Then
        Call LoadShell32Icon(oStandardIcon)
    Else
        Dim hIcon As Long
        hIcon = LoadStandardIcon(0&, oStandardIcon)
        Call DrawIcon(Me.hdc, 4&, 4&, hIcon)
        DestroyIcon hIcon
    End If
    
    '/* Make corners transparent
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, _
            (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), _
            25, 25), True

    '/* Position form on screen
    If Not bCenter Then
        Me.Move lPosX, lPosY
    End If
    Call PositionForm(bCenter)
    
End Sub

Public Function SMessageModal(ByVal sText As String, _
                        Optional ByVal sIcon As ShowIconTypes = None_i, _
                        Optional ByVal sTitle As String = vbNullString, _
                        Optional ByVal lAutoCloseSeconds As Long = 0, _
                        Optional ByVal bShowClose As Boolean = True, _
                        Optional ByVal bCenter As Boolean = True, _
                        Optional ByVal lWidth As Long = -1, _
                        Optional ByVal sFont As String = "Tahoma", _
                        Optional OwnerForm As Form) As Integer
                        
 Dim MsgType As ShowIconTypes
 Dim TesthDC As Boolean
     
    Call CheckIfLoaded
    
    '/* Separate Message Icon from input
    MsgType = sIcon And 240
    
    '/* Separate button default from input
    Select Case sIcon And 1792
    Case 256
        oButtonFocus = 1 '/* Second button is default.
    Case 512
        oButtonFocus = 2 '/* Third button is default.
    Case 768
        oButtonFocus = 3 '/* Fourth button is default.
    Case Else
        oButtonFocus = 0 '/* First button is default.
    End Select
        
    '/* Separate Button type from input
    If lAutoCloseSeconds = 0 Then bShowClose = True
    Select Case sIcon And 15
    Case vbRetryCancel
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Retry"
        cmdButton(0).Tag = vbRetry
        cmdButton(1).Visible = True
        cmdButton(1).Caption = "Cancel"
        cmdButton(1).Tag = vbCancel
        cmdButton(1).Cancel = True
        bShowClose = False
        lAutoCloseSeconds = 0
    Case vbYesNo
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Yes"
        cmdButton(0).Tag = vbYes
        cmdButton(1).Visible = True
        cmdButton(1).Caption = "No"
        cmdButton(1).Tag = vbNo
        bShowClose = False
        lAutoCloseSeconds = 0
    Case vbYesNoCancel
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Yes"
        cmdButton(0).Tag = vbYes
        cmdButton(1).Visible = True
        cmdButton(1).Caption = "No"
        cmdButton(1).Tag = vbNo
        cmdButton(2).Visible = True
        cmdButton(2).Caption = "Cancel"
        cmdButton(2).Tag = vbCancel
        cmdButton(2).Cancel = True
        bShowClose = False
        lAutoCloseSeconds = 0
    Case vbAbortRetryIgnore
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Abort"
        cmdButton(0).Tag = vbAbort
        cmdButton(1).Visible = True
        cmdButton(1).Caption = "Retry"
        cmdButton(1).Tag = vbRetry
        cmdButton(2).Visible = True
        cmdButton(2).Caption = "Ignore"
        cmdButton(2).Tag = vbIgnore
        bShowClose = False
    Case vbOKCancel
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Ok"
        cmdButton(0).Tag = vbOK
        cmdButton(1).Visible = True
        cmdButton(1).Caption = "Cancel"
        cmdButton(1).Tag = vbCancel
        cmdButton(1).Cancel = True
        bShowClose = False
        lAutoCloseSeconds = 0
    Case Else
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Ok"
        cmdButton(0).Tag = vbOK
        cmdButton(0).Cancel = True
    
    End Select
    '/* Show Help button?
    If sIcon And 16384 Then
        cmdButton(3).Visible = True
        cmdButton(3).Tag = vbHelp
    End If
    
    DisplayMessage sText, MsgType, sTitle, lAutoCloseSeconds, bShowClose, bCenter, lWidth, sFont
    DoEvents
    
    On Local Error Resume Next
    TesthDC = OwnerForm.HasDC
    If Not TesthDC Then Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)

    Me.Show vbModal
    SMessageModal = oButtonResponse
    DoEvents
    
    Unload Me
                        
End Function
Public Function SInputBox(ByVal sPrompt As String, _
                        Optional sTitle As String = vbNullString, _
                        Optional sDefault As String = vbNullString, _
                        Optional ByVal bShowClose As Boolean = False, _
                        Optional ByVal bCenter As Boolean = True, _
                        Optional ByVal sFont As String = "Tahoma") As String
                        
 
    Call CheckIfLoaded

    With cmdButton(0)
        .Visible = True
        .Caption = "Ok"
        .Tag = vbOK
        .Default = True
    End With
    With cmdButton(1)
        .Visible = True
        .Caption = "Cancel"
        .Tag = vbCancel
        .Cancel = True
    End With
    
    txtUserText.Visible = True
    oInputBox = True
    
    If sTitle = vbNullString Then sTitle = App.Title
    
    DisplayInputBox sPrompt, sTitle, sDefault, bShowClose, bCenter, sFont
    
    DoEvents
    Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    Me.Show vbModal
    
    If oButtonResponse = vbCancel Then
        SInputBox = vbNullString
    Else
        SInputBox = txtUserText.Text
    End If
    
    DoEvents
    Unload Me
                        
End Function

Public Function SMessage(ByVal sText As String, _
                        Optional ByVal sIcon As ShowIconTypes = None_i, _
                        Optional ByVal sTitle As String = vbNullString, _
                        Optional ByVal lAutoCloseSeconds As Long = 0, _
                        Optional ByVal bShowClose As Boolean = True, _
                        Optional ByVal bCenter As Boolean = True, _
                        Optional ByVal lWidth As Long = -1, _
                        Optional ByVal sFont As String = "Tahoma", _
                        Optional OwnerForm As Form) As Integer
 
 Dim MsgType As ShowIconTypes
 Dim TesthDC As Boolean
    
    Call CheckIfLoaded
    
    '/* Separate Message Icon from input
    MsgType = sIcon And 240
    
    '/* Only the OK button allowed for a non-modal message box
    If (sIcon And 15) = vbOkButton Then
        cmdButton(0).Visible = True
        cmdButton(0).Caption = "Ok"
        cmdButton(0).Tag = vbOK
        cmdButton(0).Cancel = True
        oNonModal = True
    End If
    
    DisplayMessage sText, MsgType, sTitle, lAutoCloseSeconds, bShowClose, bCenter, lWidth, sFont
    
    DoEvents
    On Local Error Resume Next
    TesthDC = OwnerForm.HasDC
    If Not TesthDC Then Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)

    
    Show , OwnerForm
    DoEvents
    Me.ZOrder
                        
End Function

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 0&
    End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgX(0).Picture
End Sub

Private Sub Timer1_Timer()
    
    oCountDown = oCountDown + 1
    If oCountDown >= oAutoCloseSeconds Then
        Unload Me
    Else
        lblTitle.Caption = oCaption & CStr(oAutoCloseSeconds - oCountDown)
    End If
    
End Sub

Private Sub LoadShell32Icon(ByVal Index As Long)
  Dim SysDir    As String
  Dim CurFile   As String
  Dim hIcon     As Long
  Dim IconCount As Long
  Dim Rv        As Long
  
    SysDir = Space(260)
    Rv = GetSystemDirectory(SysDir, 260)
    SysDir = left(SysDir, Rv) & "\"

    CurFile = SysDir & "Shell32.dll"
    IconCount = ExtractIconEx(CurFile, -1, 0, 0, 0)
    
    If IconCount >= Index Then
        Call ExtractIconEx(CurFile, Index, hIcon, 0&, 1&)
        Call DrawIcon(Me.hdc, 4&, 4&, hIcon)
        DestroyIcon hIcon
    End If
    
End Sub

Private Sub PositionForm(ByVal Center As Boolean)
  Dim Rc As Rect
  Dim T  As Long
  Dim B  As Long
  Dim L  As Long
  Dim r  As Long
  Dim mT As Long
  Dim mL As Long
  Const Offset As Long = 150
  
    '/* Get screen size
    SystemParametersInfo SPI_GETWORKAREA, 0&, Rc, 0&
    T = Rc.top * Screen.TwipsPerPixelY
    B = Rc.bottom * Screen.TwipsPerPixelY
    L = Rc.left * Screen.TwipsPerPixelX
    r = Rc.right * Screen.TwipsPerPixelX
    
    If Center Then
        '/* Center Form on screen
        mT = Abs((B / 2) - (Me.Height / 2))
        mL = Abs((r / 2) - (Me.Width / 2))
        
        If mT < T Then mT = T
        If mT > B - Me.Height Then mT = B - Me.Height
        If mL < L Then mL = L
    Else
        '/* Make sure all the Form is on the screen
        mT = Me.top
        mL = Me.left
            
        If Me.top - Offset < T Then mT = T + Offset
        If Me.left - Offset < L Then mL = L + Offset
        If Me.top + Me.Height + Offset > B Then mT = B - Me.Height - Offset
        If Me.left + Me.Width + Offset > r Then mL = r - Me.Width - Offset
    End If
    
    Me.Move mL, mT

End Sub

Private Sub CheckIfLoaded()
  Dim Frm As Form
  
    On Local Error Resume Next
    For Each Frm In Forms
        If LCase$(Frm.Name) = "frmmsgbox" Then
            Unload Frm
            Exit For
        End If
    Next Frm

End Sub
Private Sub GradientFill()
  
  Dim vert(4) As TRIVERTEX
  Dim gTRi(1) As GRADIENT_TRIANGLE
  Dim iOSver As Long
  Dim OSV        As OSVersionInfo
    
    '/* Get OS compatability flag
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then iOSver = 1 '/* Win 98/ME
        If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then iOSver = 2  '/* Win 2000/XP
    End If
 
    '/* Requires Windows 2000 or later; Requires Windows 98/ME
    If iOSver = 0 Then Exit Sub
    
    Me.AutoRedraw = True
    
    '/* Top Left Trangle
    vert(0).X = 0
    vert(0).Y = 0
    vert(0).Red = -256&
    vert(0).Green = -256&
    vert(0).Blue = -256&
    vert(0).Alpha = 0&
    
    '/* Top Right Trangle
    vert(1).X = Me.ScaleWidth * 2
    vert(1).Y = 0
    vert(1).Red = -100
    vert(1).Green = -256&
    vert(1).Blue = -256&
    vert(1).Alpha = 0&
    
    '/* Bottom Right Trangle
    vert(2).X = Me.ScaleWidth * 3
    vert(2).Y = Me.ScaleHeight * 3
    vert(2).Red = -100
    vert(2).Green = -256&
    vert(2).Blue = 0&
    vert(2).Alpha = 0&
    
    '/* Bottom Left Trangle
    vert(3).X = 0
    vert(3).Y = Me.ScaleHeight * 2
    vert(3).Red = -256&
    vert(3).Green = -256&
    vert(3).Blue = -256&
    vert(3).Alpha = 0&
    
    gTRi(0).Vertex1 = 0
    gTRi(0).Vertex2 = 1
    gTRi(0).Vertex3 = 2
    
    gTRi(1).Vertex1 = 0
    gTRi(1).Vertex2 = 2
    gTRi(1).Vertex3 = 3
    
    Call GradientFillTriangle(Me.hdc, vert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE)

End Sub

