Attribute VB_Name = "Module1"
Option Explicit

Public formsHeight As Long '//remove to erase popup notices
'amount of new email forms open. this is only used to determine location to display new popups

Public Type RECT '//remove to erase popup notices
    Left As Long '//remove to erase popup notices
    Top As Long '//remove to erase popup notices
    Right As Long '//remove to erase popup notices
    Bottom As Long '//remove to erase popup notices
End Type '//remove to erase popup notices
'used for various dimension checks

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'used to launch files, used for showing readme and opening default browser

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'plays a specified sound from location

Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
'used to establish if a connection can be made to an ip/dns

Public gAccount() As xyzgAccount
Private Type xyzgAccount
    Account As String
    Password As String
    Key As String
    Alias As String
    Interval As String
    Ticks As Byte
    Data_EmFrom() As String
    Data_NmFrom() As String
    Data_EmSummary() As String
    Data_EmTitle() As String
    NewEmails As Integer
    ResetEmails As Integer
    ErrorChecking As Boolean
End Type
'stores our unique account information

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'for reading and writing to an ini file

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'used to make various icon updates
    
Private Const WS_SYSMENU = &H80000 '//remove to erase popup notices
Private Const WS_MINIMIZEBOX = &H20000 '//remove to erase popup notices
Private Const WS_MAXIMIZEBOX = &H10000 '//remove to erase popup notices
Private Declare Function GetForegroundWindow Lib "user32" () As Long '//remove to erase popup notices
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long '//remove to erase popup notices
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long '//remove to erase popup notices
'used in our case to check if the foreground window is fullscreen, and portion of setting transparency

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '//remove to erase popup notices
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long '//remove to erase popup notices
'used to set transparency


Public Function isFullscreen() As Boolean '//remove to erase popup notices
Dim fWindow As Long
Dim winStatus As Long

    fWindow = GetForegroundWindow()
    winStatus = GetWindowLong(fWindow, -20)
    'grab foreground handle and windowstate
    If (winStatus And (WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_SYSMENU)) = 0 Then
    
    ' If winStatus = 262152 Or winStatus = 8 Then
     '   isFullscreen = True
   ' Else
    '    isFullscreen = False
    'End If '<--- seems if status is 8 or 262152 is also fullscreen, may be other codes though so am hesitant to use this method
        Dim rc As RECT
        Call GetWindowRect(fWindow, rc)
    
    
        If ((rc.Bottom - rc.Top) * Screen.TwipsPerPixelY) = Screen.Height And ((rc.Right - rc.Left) * Screen.TwipsPerPixelX) = Screen.Width Then
            isFullscreen = True
        Else
            isFullscreen = False
        End If
        'unsure if its a win7/vista issue only but while apps are full screen they are exact proportions but while
        'they are maximized the left and top alignment is off -8 returning incorrect screen allocation results
    End If

End Function
'checks if the currently viewed window is fullscreen
    
    
Public Function getINI(getSection As String, getKey As String)
Dim stringBuffer As String, charsCopied As Long

    stringBuffer = String(400, 0)
    charsCopied = GetPrivateProfileString(getSection, getKey, vbNullString, stringBuffer, 400, App.Path & "\gNotifier.ini")
    'NC is the number of characters copied to the buffer
    If charsCopied <> 0 Then
        getINI = Left$(stringBuffer, charsCopied)
    Else
    End If
    
End Function
'grabs a result from an INI file

Public Sub writeINI(writeSection As String, writeKey As String, writeData As String)
    WritePrivateProfileString writeSection, writeKey, writeData, App.Path & "\gNotifier.ini"
End Sub
'writes to an INI file

Public Function setKey(ByRef setData As String) As String
Dim tmpAsc As String, codeInjected As Byte
Dim tmpCode As Byte, tmpData As String


'this function will break down a string to its asc code 0-255
'then placing a random digit in front to confuse
    
    Randomize

    tmpData = setData
    setData = vbNullString
    Do Until LenB(tmpData) = 0
        codeInjected = 1
        'reset
        tmpAsc = Asc(Left$(tmpData, 1))
        tmpData = Mid$(tmpData, 2)
        'break down actual char to its asc
        Do Until Len(tmpAsc) = 4
            codeInjected = codeInjected + 1
            'how many digits were placed in front for this char *indent key*
            tmpCode = Rnd * 9
            tmpAsc = tmpCode & tmpAsc
        Loop
        'generates a random number
        setKey = setKey & codeInjected
        setData = setData & tmpAsc
        'add onto setkey and setdata
    Loop

End Function 'setKey returns the *indent key*,  setData returns the hashed text
'a very very weak encryption method but better than storing your passwords in plain text

Public Function getKey(ByVal setKey As String, ByVal setData As String) As String
Dim tmpKey As Byte, tmpData As String

    Do Until LenB(setKey) = 0
        tmpKey = Val#(Left$(setKey, 1))
        setKey = Mid$(setKey, 2)
        'stores first *indent key* and nips it from rest
        tmpData = Left$(setData, 4)
        setData = Mid$(setData, 5)
        'sets the first 4 numbers in the data and nips it from rest
        tmpData = Mid$(tmpData, tmpKey)
        getKey = getKey & Chr$(tmpData)
        'cuts off extra *indent keys* and converts to a letter
    Loop

End Function
'this reverses above, much easier

