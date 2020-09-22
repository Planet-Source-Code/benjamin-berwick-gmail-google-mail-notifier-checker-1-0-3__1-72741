VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GNotifier: "
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   1560
      Top             =   1080
   End
   Begin VB.Label lblSummary 
      Alignment       =   2  'Center
      Caption         =   "from names"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '//remove to erase popup notices
 '//remove to erase popup notices --- DELETE AND REMOVE FORM2
Option Explicit

Private Const SPI_GETWORKAREA& = 48
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'used to grab desktop area, for form2 popups
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'used to place popups ontop of current window
Dim myTransLevel As Byte
'transparency level of this form

Private Sub setLocation()
Dim rc As RECT, msg As String
  
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
    Me.Left = (rc.Right * Screen.TwipsPerPixelX) - Me.Width
    Me.Top = (rc.Bottom * Screen.TwipsPerPixelY) - (Me.Height + formsHeight)
    'sets form to bottom right hand of screen, above taskbar, and above any other new email displays
    
    Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2)
    'sets form on top, to ensure the person sees it
    
End Sub
'this aligns the form to the lower right hand corner just above the taskbar. using systemparainfo gathers desktop area rather
'than screen dimensions eliminating the taskbar within the calculations


Public Sub setDisplay(ByRef accountIndex, ByRef newEmailCount As Integer)
Dim emIndex As Integer

    Me.Caption = Me.Caption & gAccount(accountIndex).Alias
    'set email this is for
    lblSummary.Caption = "New Email[s] From: "
    
    For emIndex = gAccount(accountIndex).NewEmails To (newEmailCount - 1)
        lblSummary.Caption = lblSummary.Caption & gAccount(accountIndex).Data_NmFrom(emIndex) & ", "
    Next emIndex
    'apply the names of the emails from within the summary label seperated by a comma space
    lblSummary.Caption = Left$(lblSummary, Len(lblSummary.Caption) - 2)
    'strip the extra comma space
    If Len(lblSummary.Caption) > 40 Then
        lblSummary.Height = 615
        Me.Height = 1230
    Else
        lblSummary.Height = 375
        Me.Height = 975
    End If
    'this adjust the form and labels height based off from its content length. not very accurate but should help prevent the form
    'from  being excessively large in most cases
    setLocation 'placement
    formsHeight = formsHeight + Me.Height  'by adjusting this we raise the .top property by the height of previous notification popups
    
End Sub
'sets form and label caption, containing who emails are from and account of emails

Private Sub Form_Load()
    myTransLevel = 255
End Sub

Private Sub setTransparency()
Dim winStatus As Long
'0-255, 255 being fully visible

    winStatus = GetWindowLong(Me.hwnd, -20)
    SetWindowLong Me.hwnd, -20, &H80000
    SetLayeredWindowAttributes Me.hwnd, 0, myTransLevel, &H2&

End Sub
'alters the transparency of this form

Private Sub Form_Unload(Cancel As Integer)
    formsHeight = formsHeight - Me.Height
    'to adjust where form popsup
End Sub

Private Sub Timer1_Timer()
    
    Timer1.Interval = 120
    'adjust timer to swiftly adjust transparency
    If myTransLevel > 30 Then
        myTransLevel = myTransLevel - 10
        setTransparency
    'fade out till at 40 opacity then unload
    Else
        Unload Me
    End If

End Sub
'timer is only used to fade out then unload the form
'adjust default interval to change display time of notification popup
