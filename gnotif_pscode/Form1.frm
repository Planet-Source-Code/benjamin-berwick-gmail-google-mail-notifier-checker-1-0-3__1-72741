VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GMail Notifier"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   9120
      Top             =   3240
   End
   Begin VB.CommandButton cmdResetEmails 
      Caption         =   "Reset Alerts"
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdTray 
      Caption         =   "Hide Me"
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame FrameNotices 
      Caption         =   "OUTERTOUCH Productions"
      Height          =   2775
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   6975
      Begin VB.ListBox lstEmailsFrom 
         Appearance      =   0  'Flat
         Height          =   810
         Index           =   1
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtEmailContent 
         Appearance      =   0  'Flat
         Height          =   1455
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1200
         Width           =   4575
      End
      Begin VB.ListBox lstEmailsFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   0
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox lstAccountsEmails 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdNotices 
      Caption         =   "Mail Notices"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Frame FrameSettings 
      Caption         =   "  OUTERTOUCH Productions  "
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmdReadMe 
         Caption         =   "Read Me"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddAccount 
         Caption         =   "Add Account"
         Height          =   615
         Left            =   3840
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveModify 
         Caption         =   "Modify"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtAlias 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtAccount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox lstAccounts 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   3840
         Picture         =   "Form1.frx":1982
         Top             =   1480
         Width           =   1575
      End
      Begin VB.Label labelGeneric 
         Caption         =   "Interval"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label labelGeneric 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label labelGeneric 
         Caption         =   "Alias"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label labelGeneric 
         Caption         =   "Account Name"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer TimerCheckMail 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5880
      Top             =   720
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imageDetectingNet 
      Height          =   480
      Left            =   8520
      Picture         =   "Form1.frx":2689
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image ImageCheckingMail 
      Height          =   480
      Left            =   9120
      Picture         =   "Form1.frx":400B
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imageHasMail_MailError 
      Height          =   480
      Left            =   8520
      Picture         =   "Form1.frx":598D
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imageMailError 
      Height          =   480
      Left            =   9720
      Picture         =   "Form1.frx":730F
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imageHasMail 
      Height          =   480
      Left            =   7920
      Picture         =   "Form1.frx":8C91
      Top             =   1800
      Width           =   480
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "mnuRefresh"
      Visible         =   0   'False
      Begin VB.Menu mnuDispRefresh 
         Caption         =   "Check for new mail"
      End
      Begin VB.Menu mnud1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDispOpenBrowser 
         Caption         =   "Open Browser"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim formTrayed As Boolean
Dim inetBusy As Boolean
Dim mailExist As Boolean

Dim nid As NOTIFYICONDATA ' trayicon variable
Dim wordFilters() As String 'see filters routine for more info


Private Sub GatherAccounts()
Dim accString As String
Dim accountName As String, accountIndex As Byte
    
    lstAccounts.Clear
    With Form1
        .txtAccount.Text = vbNullString
        .txtAlias.Text = vbNullString
        .txtInterval.Text = vbNullString
        .txtPassword.Text = vbNullString
    End With
    ReDim gAccount(0)
    'clears
    
    If LenB(Dir$(App.Path & "\gNotifier.ini")) <> 0 Then
    'if the config file exist then read contents
        Open App.Path & "\gNotifier.ini" For Input As #1
        Do Until EOF(1)
        Line Input #1, accString
            If UBound(gAccount) < 200 Then
            'who really has 200 gmail accounts? but whatever just incase
                If Left$(accString, 1) = "[" And Right$(accString, 1) = "]" And LCase$(accString) <> "[notifier settings]" Then
                    accountName = Mid$(accString, 2, Len(accString) - 2)
                    accountIndex = UBound(gAccount)
                    'to make things easier
                    gAccount(accountIndex).Account = accountName
                    gAccount(accountIndex).Alias = getINI(accountName, "Alias")
                    gAccount(accountIndex).Interval = getINI(accountName, "Interval")
                    gAccount(accountIndex).Key = getINI(accountName, "Key")
                    gAccount(accountIndex).Password = getINI(accountName, "Password")
                    ReDim Preserve gAccount(UBound(gAccount) + 1)
                End If
            End If
        Loop
        Close #1
        If UBound(gAccount) > 0 Then ReDim Preserve gAccount(UBound(gAccount) - 1)
        'remove the extra null index
        For accountIndex = 0 To UBound(gAccount)
            lstAccounts.AddItem gAccount(accountIndex).Account
        Next accountIndex
        'add accounts to display list
    End If
    
    
End Sub
'loads accounts and their information into memory. password is loaded as encrypted value for safety purposes

Private Sub cmdAddAccount_Click()
Dim accName As String, AccPassword As String, AccAlias As String, AccInterval As String

    accName = InputBox("Account Login:", "gNotifier")
    If LenB(accName) <> 0 Then
        AccPassword = InputBox("Account Password:", "gNotifier")
        If LenB(AccPassword) <> 0 Then
            AccAlias = InputBox("Account Alias:", "gNotifier")
            If LenB(AccAlias) = 0 Then AccAlias = accName
            'if alias is not choosen, set account name as alias
            AccInterval = InputBox("Check Interval (In Minutes, Max 200)", "gNotifier")
            If Val#(AccInterval) < 1 Or Val#(AccInterval) > 200 Then AccInterval = 11
            AccInterval = Val#(AccInterval)
            'if interval is less than 1 minute or more than 200 default to 11
        End If
    End If
    'prompt for variables to save to config
    If LenB(accName) = 0 Or LenB(AccPassword) = 0 Or LenB(AccInterval) = 0 Then Exit Sub
    'exit if a vital field is empty

Dim arrIndex As Byte
    For arrIndex = 0 To UBound(gAccount)
        If LCase$(accName) = LCase$(gAccount(arrIndex).Account) Then
            MsgBox "The account name you have entered is already in the database." & vbCrLf & "Select the account in the list and choose 'Modify'", vbOKOnly
            Exit Sub
        End If
        'account is already in db, exit sub
    Next arrIndex
    
Dim keyHash As String
    keyHash = setKey(AccPassword)
    'hashes the password and sets the key
    ReDim Preserve gAccount(UBound(gAccount) + 1)
    gAccount(UBound(gAccount)).Account = accName
    gAccount(UBound(gAccount)).Alias = AccAlias
    gAccount(UBound(gAccount)).Interval = AccInterval
    gAccount(UBound(gAccount)).Key = keyHash
    gAccount(UBound(gAccount)).Password = AccPassword
    gAccount(UBound(gAccount)).Ticks = 201 'set to max so this gets checked next run
    SubmitToINI
    'write results to array then save to ini and reload
    lstAccounts.AddItem gAccount(UBound(gAccount)).Account
    'add account to display
    MsgBox "Account '" & accName & "' has been added!"
    
End Sub
'prompts when a user clicks add account

Private Sub cmdDelete_Click()
Dim accIndex As Byte
'Dim passLength As Byte

    accIndex = getAccountIndex(lstAccounts)
    If accIndex <> -1 Then
        If MsgBox("Are you sure you wish to remove account '" & gAccount(accIndex).Account & "', alias '" & gAccount(accIndex).Alias & "'?", vbOKCancel) <> 1 Then Exit Sub
        'if they canceled the delete exit sub
        gAccount(accIndex).Account = vbNullString
        lstAccounts.RemoveItem lstAccounts.ListIndex
        LockTextFields
        SubmitToINI
    End If
    
End Sub
'deletes an account, rewrites INI reloads INI

Private Function getAccountIndex(theList As ListBox) As Integer
Dim accName As String

    getAccountIndex = -1
    If theList.ListIndex <> -1 Then
    'if a user is selected in the listbox
        accName = theList.List(theList.ListIndex)
        If LenB(accName) <> 0 Then
        'if user isnt selecting a null account
            getAccountIndex = 0
            Do Until accName = gAccount(getAccountIndex).Account Or getAccountIndex > UBound(gAccount)
                getAccountIndex = getAccountIndex + 1
            Loop
            'search
            If getAccountIndex > UBound(gAccount) Then getAccountIndex = -1
            'couldnt find account, return -1
        End If
    End If
    
End Function
'gathers the index of an account within the array via its account name (not alias)

Private Sub SubmitToINI()
Dim accIndex As Byte

    Open App.Path & "\gNotifier.ini" For Output As #1
    Print #1, vbNullString
    Close #1
    'erase file w/o killing it, safe to use on vista+
    For accIndex = 0 To UBound(gAccount)
        If LenB(gAccount(accIndex).Account) <> 0 Then
            writeINI gAccount(accIndex).Account, "Alias", gAccount(accIndex).Alias
            writeINI gAccount(accIndex).Account, "Interval", gAccount(accIndex).Interval
            writeINI gAccount(accIndex).Account, "Key", gAccount(accIndex).Key
            writeINI gAccount(accIndex).Account, "Password", gAccount(accIndex).Password
            'write to the ini
        End If
    Next accIndex
 
End Sub
'is used to save alterations to an accounts details, used by multiple routines

Private Sub cmdNotices_Click()
    If lstAccounts.ListCount = 0 Then Exit Sub
    'they have no emails configured...
    FrameSettings.Top = 49999
    FrameSettings.Enabled = False
    FrameNotices.Enabled = True
    FrameNotices.Top = 120
    'hide noticies show settings
    cmdResetEmails.Enabled = True
    cmdResetEmails.Top = 3000
    'enable and set to proper place
    cmdSettings.Enabled = True
    cmdNotices.Enabled = False
    'toggle buttons
    Me.Width = 7335
    checkAccountsWithMail
End Sub
'check mail / notices is clicked

Private Sub checkAccountsWithMail()
Dim accIndex As Byte

    lstAccountsEmails.Clear
    lstEmailsFrom(0).Clear
    lstEmailsFrom(1).Clear
    txtEmailContent.Text = vbNullString
    For accIndex = 0 To UBound(gAccount)
        If LenB(gAccount(accIndex).Account) <> 0 Then
            If Val#(gAccount(accIndex).NewEmails) > 0 Then lstAccountsEmails.AddItem gAccount(accIndex).Account
        End If
    Next accIndex
    If lstAccountsEmails.ListCount <> 0 Then lstAccountsEmails.ListIndex = 0
    
End Sub
'checks which accounts have mail and adds them to the listbox displaying which do

Private Sub cmdReadMe_Click()
    ShellExecute Me.hwnd, "Open", App.Path & "\gnotif_ReadME.txt", vbNullString, "C:\", &H1
End Sub
'opens read me

Private Sub cmdResetEmails_Click()
Dim arrIndex As Integer

    For arrIndex = 0 To UBound(gAccount)
        gAccount(arrIndex).ResetEmails = gAccount(arrIndex).NewEmails
    Next arrIndex
    MsgBox "All new email notices have been reset!"

End Sub
'this routine just sets resetemails to the same value as new emails.

Private Sub cmdSaveModify_Click()
    
    If LenB(txtAccount.Text) = 0 Then Exit Sub
    'do not continue if an account is not selected or valid
    If cmdSaveModify.Caption = "Modify" Then
    'enables the text fields to allow editing and changes button caption to allow saving
        cmdSaveModify.Caption = "Save"
        With Form1
            .cmdDelete.Enabled = True
            .txtAccount.Enabled = True
            .txtAlias.Enabled = True
            .txtInterval.Enabled = True
            .txtPassword.Enabled = True
        End With
    Else
    'disables field and saves information submitted
        LockTextFields
        SaveTextFields
    End If
    
End Sub
'user clicks save/modify button

Private Sub LockTextFields()
    cmdSaveModify.Caption = "Modify"
    With Form1
        .cmdDelete.Enabled = False
        .txtAccount.Enabled = False
        .txtAlias.Enabled = False
        .txtInterval.Enabled = False
        .txtPassword.Enabled = False
    End With
End Sub
'this sub will disable text fields and rename modify/save button

Private Sub SaveTextFields()
Dim accIndex As Byte, keyHash As String, accPass As String
'Dim passLength As Byte

    accIndex = getAccountIndex(lstAccounts)
    If accIndex <> -1 Then
        With Form1
            If .txtPassword.Text <> gAccount(accIndex).Password Then
                accPass = .txtPassword.Text
                keyHash = setKey(accPass)
                gAccount(accIndex).Password = accPass
                gAccount(accIndex).Key = keyHash
            End If
            'save pw only if its been modified
            .txtInterval.Text = Val#(.txtInterval.Text)
            If Val#(.txtInterval.Text) < 1 Or Val#(.txtInterval) > 200 Then .txtInterval.Text = 11
            gAccount(accIndex).Interval = .txtInterval.Text
            'set interval
            gAccount(accIndex).Alias = .txtAlias.Text
        End With
        SubmitToINI
        'save and reload
    End If

End Sub
'called to save changes made when save button is clicked

Private Sub cmdSettings_Click()
    FrameNotices.Top = 49999
    FrameNotices.Enabled = False
    FrameSettings.Enabled = True
    FrameSettings.Top = 120
    'hide noticies show settings
    cmdResetEmails.Enabled = False
    cmdResetEmails.Top = 20000
    'disable and move off a little so its not shown on UI
    cmdSettings.Enabled = False
    cmdNotices.Enabled = True
    'toggle buttons
    Me.Width = 5895
End Sub
'user clicks settings button

Private Sub cmdTray_Click()
    If lstAccounts.ListCount = 0 Then Exit Sub
    'no point in being in the tray if they dont have accounts to check
    TrayForm
End Sub
'sends to tray via button click

Private Sub Form_Load()

    If App.PrevInstance Then End
    
    ReDim gAccount(0)
    
    Me.Width = 5895
    Me.Height = 3915
    'i just use this to resize to the proper size so I can stretch out the project to view all my objects
    
    Load_Filters
    cmdSettings_Click
    GatherAccounts
    
    DoEvents
    TimerCheckMail.Interval = 5000
    TimerCheckMail.Enabled = True
    'by doing this it triggers other routines that simply just wait 5 seconds before attempting to check for mail.
    'basic idea of this is when the application runs on bootup
    cmdTray_Click
    Load_MailCheck
    
    
    'Using setKey and getKey
    'setKey:
    '   KeyCode_IndentKeys = setKey(Text_To_Be_Hashed)
    'getKey:
    '   DeHashed_Text = getKey(KeyCode_IndentKeys, Text_To_Be_DeHashed)


End Sub
'application started

Private Sub Load_Filters()
Dim sData As String

    ReDim wordFilters(0) As String
    'so we can use
    If LenB(Dir$(App.Path & "\filters.txt")) <> 0 Then
    'only try to open file if it exist
        Open App.Path & "\filters.txt" For Input As #1
        Do Until EOF(1)
        Line Input #1, sData
            wordFilters(UBound(wordFilters)) = sData
            ReDim Preserve wordFilters(UBound(wordFilters) + 1)
        Loop
        Close #1
    End If
    
    If UBound(wordFilters) > 0 Then ReDim Preserve wordFilters(UBound(wordFilters) - 1)
    'if entry added, remove extra array no point in having it

End Sub
'loads filters  from filters.txt

Private Sub Load_MailCheck()
Dim accIndex As Byte

    For accIndex = 0 To UBound(gAccount)
        gAccount(accIndex).Ticks = 201
    Next accIndex

End Sub
'sets all accounts ready to be checked, used on reload emails or startup

'Private Sub Inet1_StateChanged(ByVal State As Integer)
'Dim inChunk As String, inData As String
'    Select Case State
'        Case icNone
'            '' "None"
'        Case icResolvingHost
'            '' "Resolving Host"
'        Case icHostResolved
'            '' "Resolved"
'        Case icConnecting
'            '' "Connecting"
'        Case icConnected
'            '' "Connected"
'        Case icResponseReceived
'            '' "AHHHHH
'        Case icDisconnecting
'            '' "Disconneting"
'        Case icDisconnected
'            '' "Disconnected"
'        Case icError
'            '' "ERR:" & vbCrLf & Inet1.ResponseCode & ": " & Inet1.ResponseInfo
'        Case icResponseCompleted
'            '' "Response Completed"
'    End Select
'
'End Sub
'state changes in inet, not really needed but keeping for reference

Private Sub CheckforEmails(ByRef parseData As String, ByRef accbeingChecked As Byte)

    If InStrB(LCase$(parseData), "<fullcount>") Then
    'if contains correct string
        Call setEmailData(parseData, accbeingChecked)
        'set data and new email count
        parseData = Mid$(parseData, InStr(parseData, "<fullcount>") + 11)
        parseData = Left$(parseData, InStr(parseData, "</fullcount>") - 1)
        'parses number of emails only
        If Val#(parseData) > gAccount(accbeingChecked).NewEmails Then
            Call sndPlaySound(App.Path & "\gnot_newmail.wav", &H1)
            'play sound if emails increased on this account '###SOUND###
            If Not isFullscreen Then '//remove to erase popup notices
                Dim newPopup As New Form2 '//remove to erase popup notices
                newPopup.Show '//remove to erase popup notices
                newPopup.setDisplay accbeingChecked, Val#(parseData) '//remove to erase popup notices
            End If '//remove to erase popup notices
            'if not using a fullscreen app this will show form2, and set the information to be displayed on it
        End If
        gAccount(accbeingChecked).NewEmails = Val#(parseData)
        'set amount of new emails
        gAccount(accbeingChecked).ErrorChecking = False
    Else
        If LenB(parseData) = 0 Then
            gAccount(accbeingChecked).Ticks = 201
            'set to be checked again, null data means no connection; try again in 5 seconds
            Form1.TimerCheckMail.Interval = 5000
        'no reply, suggesting no internet connection
        Else
            gAccount(accbeingChecked).ErrorChecking = True
        'invalid parse report error
        End If
    End If
    If formTrayed Then nidIcon 'updates tray icon

End Sub
'runs through data recieved from inet for new emails

Private Sub lstAccounts_Click()
Dim accName As String, accIndex As Byte
'Dim passLength As Byte

    accIndex = getAccountIndex(lstAccounts)
    If accIndex <> -1 Then
        'passLength = Len(getKey(gAccount(accIndex).Key, gAccount(accIndex).Password))
        'gets length of dehashed pass for display purposes only(currently not used)
        With Form1
            .txtAccount.Text = gAccount(accIndex).Account
            .txtAlias.Text = gAccount(accIndex).Alias
            .txtInterval.Text = Val#(gAccount(accIndex).Interval)
            .txtPassword.Text = gAccount(accIndex).Password
        End With
    End If

End Sub
'single clicked accounts on settings page

Private Sub lstAccounts_DblClick()
    LaunchBrowser lstAccounts
End Sub
'double clicked accounts on settings page, launch with browser

Private Sub lstAccountsEmails_Click()
Dim accIndex As Byte, entryIndex As Integer

    accIndex = getAccountIndex(lstAccountsEmails)
    lstEmailsFrom(0).Clear
    lstEmailsFrom(1).Clear
    txtEmailContent.Text = vbNullString
    'reset vars
    If accIndex <> -1 Then
        For entryIndex = 0 To (UBound(gAccount(accIndex).Data_EmFrom) - 1)
            lstEmailsFrom(0).AddItem gAccount(accIndex).Data_NmFrom(entryIndex) & ": "
            lstEmailsFrom(1).AddItem gAccount(accIndex).Data_EmTitle(entryIndex)
        Next entryIndex
        'adds our Name from and Email title to the listboxes when an email account is selected within the view emails section
        If lstEmailsFrom(0).ListCount <> 0 Then lstEmailsFrom(0).ListIndex = 0
        'if theres new emails select first (which there should be if they even got this far)
    End If
        
End Sub
'single clicked listbox containing accounts with new emails

Private Sub lstAccountsEmails_DblClick()
    cmdTray_Click
    LaunchBrowser lstAccountsEmails
End Sub
'double clicked listbox containing accounts with new emails, launch browser with specified email in url

Private Sub LaunchBrowser(theList As ListBox)
Dim accIndex As Integer
    accIndex = getAccountIndex(theList)
    'finds account in array
    If accIndex <> -1 Then ShellExecute Me.hwnd, "Open", "https://www.google.com/accounts/ServiceLoginAuth?continue=https://mail.google.com/mail&service=mail&Email=" & gAccount(accIndex).Account, vbNullString, "C:\", &H1
End Sub
'this will launch the selected account in default browser


Private Sub setEmailData(ByVal strData As String, ByRef accountIndex As Byte)
On Error GoTo Err
        
    ReDim gAccount(accountIndex).Data_EmFrom(0) As String
    ReDim gAccount(accountIndex).Data_EmSummary(0) As String
    ReDim gAccount(accountIndex).Data_EmTitle(0) As String
    ReDim gAccount(accountIndex).Data_NmFrom(0) As String
    
    Do Until InStrB(LCase$(strData), "<entry>") = 0
        With gAccount(accountIndex)
            strData = Mid$(strData, InStr(LCase$(strData), "<entry>") + 7)
            'parses past the <entry>
            strData = Mid$(strData, InStr(LCase$(strData), "<title>") + 7)
            .Data_EmTitle(UBound(.Data_EmTitle)) = Left$(strData, InStr(LCase$(strData), "</title>") - 1)
            ReDim Preserve .Data_EmTitle(UBound(.Data_EmTitle) + 1)
            'parses the title, and then past it
            strData = Mid$(strData, InStr(LCase$(strData), "<summary>") + 9)
            .Data_EmSummary(UBound(.Data_EmSummary)) = Left$(strData, InStr(LCase$(strData), "</summary>") - 1)
            ReDim Preserve .Data_EmSummary(UBound(.Data_EmSummary) + 1)
            'parses the summary, and then past it
            strData = Mid$(strData, InStr(LCase$(strData), "<author>") + 8)
            'parses past the <author>
            strData = Mid$(strData, InStr(LCase$(strData), "<name>") + 6)
            .Data_NmFrom(UBound(.Data_NmFrom)) = Left$(strData, InStr(LCase$(strData), "</name>") - 1)
            ReDim Preserve .Data_NmFrom(UBound(.Data_NmFrom) + 1)
            'parses from email, and then past it
            strData = Mid$(strData, InStr(LCase$(strData), "<email>") + 7)
           .Data_EmFrom(UBound(.Data_EmFrom)) = Left$(strData, InStr(LCase$(strData), "</email>") - 1)
            ReDim Preserve .Data_EmFrom(UBound(.Data_EmFrom) + 1)
            'parses from email, and then past it
        End With
    Loop

Err:

End Sub
'is called when new emails are found within routine checkforemails, stores subject, from name ect


Private Sub lstEmailsFrom_Click(Index As Integer)
Dim accIndex As Byte

    If lstEmailsFrom(Index).ListIndex <> -1 Then
        If Index = 0 Then
            If lstEmailsFrom(1).ListIndex <> lstEmailsFrom(0).ListIndex Then lstEmailsFrom(1).ListIndex = lstEmailsFrom(0).ListIndex
        ElseIf Index = 1 Then
            If lstEmailsFrom(0).ListIndex <> lstEmailsFrom(1).ListIndex Then lstEmailsFrom(0).ListIndex = lstEmailsFrom(1).ListIndex
        End If
        'highlights the other lstemailsfrom listbox just for show
        accIndex = getAccountIndex(lstAccountsEmails)
        txtEmailContent.Text = filterMailText(gAccount(accIndex).Data_EmSummary(lstEmailsFrom(Index).ListIndex))
        'pulls the index tied to the selected account, and then displays the summary tied to the selected email
    End If
    
End Sub
'called when a user selects an email subject or from within the new emails tab

Private Function filterMailText(ByRef emailText As String) As String
Dim findString As String, replacedString As String
Dim arrIndex As Integer

    filterMailText = emailText
    '
    For arrIndex = 0 To UBound(wordFilters)
        findString = LCase$(wordFilters(arrIndex))
        If Left$(findString, 5) = "find(" And InStrB(findString, "replace(") <> 0 Then
            replacedString = Mid$(findString, InStr(findString, "replace(") + 8)
            replacedString = Left$(replacedString, Len(replacedString) - 1)
            'sets the word to replace the find with
            findString = Left$(findString, InStr(findString, ")replace(") - 1)
            findString = Mid$(findString, 6)
            'sets string to search for
            filterMailText = Replace$(filterMailText, findString, replacedString)
            'cheezy replace...probably could advance on this system later
        End If
        
    Next arrIndex

End Function
'a simple function to replace text, our filters, within emails or any text for that matter

Private Sub mnuDispOpenBrowser_Click()
    ShellExecute Me.hwnd, "Open", "https://mail.google.com/", vbNullString, "C:\", &H1
End Sub
'opens browser via right click tray icon

Private Sub mnuDispRefresh_Click()
    Load_MailCheck
    TimerCheckMail_Timer
End Sub
'user clicked refresh on tray icon, check all accounts if internet connection is detected

Private Sub TimerCheckMail_Timer()
Dim accIndex As Byte, tmpStr As String
On Error GoTo Err

    If Not formTrayed Then Exit Sub
    'dont run while form is visible
    If TimerCheckMail.Interval = 5000 Then
        If InternetCheckConnection("http://www.google.com/", &H1, 0&) = 0 Then
            TimerCheckMail.Interval = 5000
            Exit Sub
        End If
    End If
    'if cannot connect to google then assume no network avail and set timer to check every 5seconds
    TimerCheckMail.Enabled = False
    'tmpdisable
    TimerCheckMail.Interval = 60000
    For accIndex = 0 To UBound(gAccount)
        If LenB(gAccount(accIndex).Account) <> 0 Then
            gAccount(accIndex).Ticks = Val#(gAccount(accIndex).Ticks) + 1
            'increase ticks
            If Val#(gAccount(accIndex).Ticks) >= Val#(gAccount(accIndex).Interval) Then
            'if ticks >= interval check mail!
                If Not inetBusy Then
                    inetBusy = True
                    nidIcon
                    gAccount(accIndex).Ticks = 0
                    'prep vars
                    DoEvents
                    Inet1.Protocol = icHTTPS
                    Inet1.URL = "https://mail.google.com/mail/feed/atom/"
                    Inet1.AccessType = icDirect
                    Inet1.UserName = gAccount(accIndex).Account
                    Inet1.Password = getKey(gAccount(accIndex).Key, gAccount(accIndex).Password)
                    tmpStr = Inet1.OpenURL
                    Do While Inet1.StillExecuting
                        DoEvents
                    Loop
                    inetBusy = False
                    'tells inet to download feed data
                    CheckforEmails tmpStr, accIndex
                    'processes feed data recieved
                End If
            End If
        End If
    Next accIndex

Err:

    inetBusy = False
    TimerCheckMail.Enabled = True
    
End Sub
'sixty second timer which determines and adjust when emails are to be checked, and performs the task of checking for new emails

Private Sub TrayForm()
    Me.Hide
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nidIcon 'the following sub will shell the icon updates
End Sub
'sets tray information

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As Long

    msg = x / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
            formTrayed = False
            'untray
            If mailExist Then
                cmdNotices.Enabled = True
                cmdNotices_Click
            Else
                cmdSettings.Enabled = True
                cmdSettings_Click
            End If
            'click appropriate button to show
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP
            PopupMenu mnuRefresh
        Case WM_RBUTTONDBLCLK
    End Select

End Sub
'tray icon is clicked, perform actions

Private Sub nidIcon()
Dim accIndex, errFound As Boolean, newMail As Boolean, tooltipData As String

    For accIndex = 0 To UBound(gAccount)
        If LenB(gAccount(accIndex).Account) <> 0 Then
            If gAccount(accIndex).NewEmails > 0 Then
                If gAccount(accIndex).ResetEmails <> gAccount(accIndex).NewEmails Then
                    'when we reset our notifications it simply just records the amount of emails already contained. and if reset
                    'value is the same as current value dont display that theres new emails, because this is the point at which
                    'they select reset. there are a few unlikely scenarios this could report inaccurately...should improve
                    'ie: they have 3 emails, they reset...then check their 3 emails. should they get 3 additional emails resulting in
                    'again 3 new emails before the next check cycle application will think its the same 3 that already existed.
                    gAccount(accIndex).ResetEmails = 0 'put this back at 0 so they either have to reset again or check their email
                    newMail = True
                    If gAccount(accIndex).ErrorChecking Then
                        tooltipData = tooltipData & "{E} "
                        errFound = True
                    End If
                End If
                tooltipData = tooltipData & gAccount(accIndex).Alias & ": " & gAccount(accIndex).NewEmails & vbCrLf
                'display tooltip data regardless of new email icon. this way they can mouse over and not forget they have new emails
                'that they reset.
            Else
                If gAccount(accIndex).ErrorChecking Then
                    tooltipData = tooltipData & "{E} " & gAccount(accIndex).Alias & ": ?" & vbCrLf
                    errFound = True
                End If
            End If
        End If
    Next accIndex
    'set tooltips
    With Form1
        If newMail Then mailExist = True
        If TimerCheckMail.Interval = 5000 Then
            nid.hIcon = .imageDetectingNet
            tooltipData = "Detecting internet connection..."
        ElseIf errFound And newMail Then
            nid.hIcon = .imageHasMail_MailError.Picture
        ElseIf errFound And Not newMail Then
            nid.hIcon = .imageMailError.Picture
        ElseIf newMail Then
            nid.hIcon = .imageHasMail.Picture
        ElseIf Not newMail Then
            nid.hIcon = Me.Icon
            If LenB(tooltipData) = 0 Then tooltipData = "No new mail."
            'only change if tooltil is null, if it isnt null then user has reset email count
            mailExist = False
        End If
        If inetBusy Then
            nid.hIcon = .ImageCheckingMail.Picture
            tooltipData = "Checking for mail..."
        End If
    End With

    If Right$(tooltipData, 2) = vbCrLf Then tooltipData = Left$(tooltipData, Len(tooltipData) - 2)
    'remove extra return on end of tooltip display
    nid.szTip = tooltipData & vbNullChar
    If Not formTrayed Then
        Shell_NotifyIcon NIM_ADD, nid
        formTrayed = True
    Else
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
    
End Sub
'alters tray icon to various graphics

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
    End
End Sub
'form is unloaded, delete tray icon and terminate
