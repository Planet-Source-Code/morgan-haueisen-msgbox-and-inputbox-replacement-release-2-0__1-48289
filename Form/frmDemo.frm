VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MsgBox Demo"
   ClientHeight    =   3900
   ClientLeft      =   4005
   ClientTop       =   2655
   ClientWidth     =   2505
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   2505
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FFFF&
      Caption         =   "New Input Box Example"
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3375
      Width           =   2220
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Display Changing Data"
      Height          =   375
      Left            =   105
      TabIndex        =   4
      Top             =   2308
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MessageBox w/ buttons"
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   2220
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Question"
      Height          =   375
      Left            =   105
      TabIndex        =   5
      Top             =   2840
      Width           =   2220
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Narrow and Long"
      Height          =   375
      Left            =   105
      TabIndex        =   2
      Top             =   1244
      Width           =   2220
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Display an Error"
      Height          =   375
      Left            =   105
      TabIndex        =   3
      Top             =   1776
      Width           =   2220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Wide and Thin"
      Height          =   375
      Left            =   105
      TabIndex        =   1
      Top             =   712
      Width           =   2220
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Command2_Click()

    frmMsgBox.SMessage "This is a messagebox with only a close button!  " & _
        "It is set to display at the current mouse position and has a fixed width of 10000.  " & _
        "It has an owner Form which means that when you minimize the main Form this messgaebox will also minimize.  " & _
        "It is shown as a non-modal form." & vbCrLf & vbCrLf & _
        "Move the main from closer to the right side of the screen and show it again; you will see that it will not fall of the edge.", _
        None_i, Command2.Caption, , , False, 10000, , Me
    
End Sub
Private Sub Command3_Click()

frmMsgBox.SMessage "This is a message box that does not display any buttons and will auto-close after ten " & _
    "seconds. It has a fixed width of 2000 and has no owner form which means it will always show even if " & _
    "the main form is minimized.", Save_i, , 10, False, False, 2000
    
End Sub

Private Sub Command4_Click()
  Dim msg As String

    msg = "Well what do you think?" & vbCrLf & _
        "This does everything the standard MsgBox does (except how it handles the additional Help button) plus more.  " & _
        "I have made it so that you can easily add this to any existing project by simply replacing " & _
        " MsgBox with frmMsgBox.SMessageModal"
    
    
    'MsgBox msg, vbQuestion + vbOKCancel + vbMsgBoxHelpButton
        
    If frmMsgBox.SMessageModal(msg, vbQuestion + vbOKCancel + vbMsgBoxHelpButton) = vbHelp Then
        frmAbout.Show , Me
    End If

End Sub

Private Sub Command5_Click()
  Dim msg As String, X As Integer
  
    '/* Create an error to display
    On Error GoTo DivisionError
    X = 1 / 0

Exit Sub

DivisionError:
    
    msg = "Error# " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf
    msg = msg & "An error occured trying to " & _
        "divide the numbers. Make sure they are both valid numbers, don't contain " & _
        "letters, etc., and try again." & vbCrLf & vbCrLf & "This " & _
        "msgbox uses the font Arial and is a modal form."
    
    frmMsgBox.SMessageModal msg, vbCritical + vbOkButton
    
End Sub
Private Sub Command1_Click()
Dim r As Integer
    r = frmMsgBox.SMessageModal("This is a sample message box to " & _
        "demonstrate the use of buttons! It is set to display Exclamation icon, " & _
        "Yes/No/Cancel buttons, make the second button the default, and is centered on the screen." & _
        vbCrLf & vbCrLf & "This purpose of this demo is to give you a small sample of the many different combinations possible.", _
        vbExclamation + vbYesNoCancel + vbDefaultButton2, Command1.Caption)
    
    Select Case r
    Case vbAbort
        frmMsgBox.SMessageModal "Abort button pressed", vbInformation + vbOkButton
    Case vbRetry
        frmMsgBox.SMessageModal "Retry button pressed", vbInformation + vbOkButton
    Case vbIgnore
        frmMsgBox.SMessageModal "Ignore button pressed", vbInformation + vbOkButton
    Case vbYes
        frmMsgBox.SMessageModal "Yes button pressed", vbInformation + vbOkButton
    Case vbNo
        frmMsgBox.SMessageModal "No button pressed", vbInformation + vbOkButton
    Case vbCancel
        frmMsgBox.SMessageModal "Cancel button pressed", vbInformation + vbOkButton
    Case vbHelp
        frmMsgBox.SMessageModal "Help button pressed", vbInformation + vbOkButton
    End Select
    
End Sub

Private Sub Command6_Click()
  Dim i As Byte
    
    frmMsgBox.SMessage "Building GRAPH..Please Wait" & vbCrLf & vbCrLf, Hourglass_i, , , False
    Sleep 1000
    
    For i = 1 To 52
            frmMsgBox.txtMessage = "Building GRAPH..Please Wait" & vbCrLf & "Week: " & CStr(i)
            frmMsgBox.txtMessage.Refresh
            Sleep 50 '/* allow some time so that you can see the week numbers change
    Next i
    Unload frmMsgBox
    Sleep 500 '/* allow some time so that you can see the last week displayed
    
    frmMsgBox.SMessageModal "This concludes the '" & Command6.Caption & "' demo", vbInformation + vbOkButton, "Demo By Morgan Haueisen"
    
End Sub


Private Sub Command7_Click()
Dim r As String
    r = frmMsgBox.SInputBox("Enter Something:", , "Default String", , False)
    frmMsgBox.SMessageModal "You entered """ & r & """", vbInformation + vbOkButton
    
End Sub


