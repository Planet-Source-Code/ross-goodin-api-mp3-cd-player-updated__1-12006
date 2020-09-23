VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   4020
   ClientTop       =   3840
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form2.frx":0442
   ScaleHeight     =   2250
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   4320
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   960
      Left            =   4440
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   4440
      Top             =   480
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   705
      ItemData        =   "Form2.frx":16414
      Left            =   120
      List            =   "Form2.frx":16416
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   2280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   495
      Width           =   4005
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Track 0 of 0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image10 
      Height          =   405
      Left            =   230
      Top             =   950
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   2880
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   3480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4080
      Top             =   960
      Width           =   255
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   4200
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   840
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2160
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1560
      Top             =   960
      Width           =   495
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Begin VB.Menu jhng 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnurest 
         Caption         =   "Restore"
         Visible         =   0   'False
      End
      Begin VB.Menu asdffd 
         Caption         =   "-"
      End
      Begin VB.Menu mnustuff 
         Caption         =   "PlayBack"
         Begin VB.Menu mnuopensong 
            Caption         =   "Open File"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuadddir 
            Caption         =   "Add Directory"
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuRem 
            Caption         =   "Remove Song"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuClear 
            Caption         =   "Clear List"
            Shortcut        =   ^C
         End
         Begin VB.Menu space6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlay 
            Caption         =   "Play"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuStop 
            Caption         =   "Stop"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuNext 
            Caption         =   "Next Song"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuPrev 
            Caption         =   "Prev. Song"
            Shortcut        =   {F7}
         End
         Begin VB.Menu space7 
            Caption         =   "-"
         End
         Begin VB.Menu mnurand 
            Caption         =   "Random"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnureteat 
            Caption         =   "Repeat"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Begin VB.Menu mnuformatwav 
            Caption         =   "Format Mp3 to WAV"
            Enabled         =   0   'False
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuformatmp3 
            Caption         =   "Format Wav to MP3"
            Enabled         =   0   'False
            Shortcut        =   ^E
         End
         Begin VB.Menu space8 
            Caption         =   "-"
         End
         Begin VB.Menu mnusetdir 
            Caption         =   "Set File Directory"
            Shortcut        =   ^D
         End
         Begin VB.Menu sadf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStay 
            Caption         =   "Stay On Top"
         End
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuskins 
         Caption         =   "Select Skin"
      End
      Begin VB.Menu sfghfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuopenlist 
         Caption         =   "Save List"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnulist 
         Caption         =   "Open List"
         Shortcut        =   ^O
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCDopen 
         Caption         =   "Open CD Rom"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuCDclose 
         Caption         =   "Close CD Rom"
         Shortcut        =   {F9}
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   {F12}
      End
      Begin VB.Menu afsg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTotal 
         Caption         =   "Total Number Of Songs Played"
      End
      Begin VB.Menu kluiuki 
         Caption         =   "-"
      End
      Begin VB.Menu mnuext 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const MAX_STRING_LEN As Long = 500

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Msg As Long, nid As NOTIFYICONDATA, j As Long, OpenError As Boolean

    Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONDBLCLK = &H206
    Private Const conHwndTopmost = -1
    Private Const conSwpNoActivate = &H10
    Private Const conSwpShowWindow = &H40
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Dim tmp As String * 255
Dim ShortPath As Long
Dim ShortPathAndFie As String
Dim ex As String
Dim ext As String
Dim dr As String
Dim dr1 As String
Dim s As String * 30
Dim Playing As Boolean
Dim paused As Boolean
Dim Min1 As Integer
Dim Min2 As Integer
Dim Sec1 As Integer
Dim Sec2 As Integer
Dim playlist As Integer
Dim Randy As Integer
Dim Reteat As Integer
Dim OnTop As Integer
Dim Min As Integer
Dim readINI As String
Dim TotalNumSongs As Integer
Dim MainPic As String
Dim PlayListPic As String
Dim AboutPic As String
Dim SplashPic As String
Dim DirPic As String

Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hwnd, &HA1, 2, 0&)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.ScaleMode = vbPixels Then
        Msg = X
    Else
        Msg = X / Screen.TwipsPerPixelX
    End If

    Select Case Msg
        Case WM_RBUTTONUP
        Me.PopupMenu Me.mnu
        Case WM_LBUTTONDBLCLK
        mnurest_Click
    End Select
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
'open file
Image5_Click
ElseIf KeyCode = vbKeyF2 Then
'remove file
List1.RemoveItem List1.ListIndex
List2.RemoveItem List1.ListIndex
frmPlayList.List1.RemoveItem frmPlayList.List1.ListIndex

ElseIf KeyCode = vbKeyF3 Then
'play song
Image3_Click
ElseIf KeyCode = vbKeyF4 Then
'pause song
Image1_Click
ElseIf KeyCode = vbKeyF5 Then
'stop song
Image2_Click
ElseIf KeyCode = vbKeyF6 Then
'next song in playlist
Image6_Click
ElseIf KeyCode = vbKeyF7 Then
'prev song in playlist
Image10_Click
ElseIf KeyCode = vbKeyF8 Then
'open cd rom door
mnuCDopen_Click
ElseIf KeyCode = vbKeyF12 Then
'about me
frmAbout.Show
ElseIf KeyCode = vbKeyF9 Then
'close cd rom door
mnuCDClose_Click
ElseIf KeyCode = vbKeyF11 Then
'close cd rom door
mnuadddir_Click
End If
If (Shift And vbCtrlMask) > 0 Then
Select Case KeyCode
Case vbKeyC
mnuClear_Click
Case vbKeyS
mnuOpenlist_Click
Case vbKeyO
mnuList_Click
Case vbKeyA
mnurand_Click
Case vbKeyR
mnureteat_Click
Case vbKeyD
mnusetdir_Click
End Select
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim buf As String * 256, a_line As String, length As Long, numfiles As String, file, fnum As Integer, lines As Integer, boo, i As Variant
Dim sRet As String
sRet = String(255, Chr(0))
mnu.Visible = False
frmPlayList.Visible = False
TheX = Picture1.ScaleWidth
SetWindowPos hwnd, conHwndTopmost, 250, 250, 302, 102, conSwpNoActivate Or conSwpShowWindow

Me.Hide
mciSendString "open cdaudio Alias cd wait shareable", 0, 0, 0
readINI = Left(sRet, GetPrivateProfileString("Settings", "File Dir", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
rep = Left(sRet, GetPrivateProfileString("Settings", "Random", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
ret = Left(sRet, GetPrivateProfileString("Settings", "Repeat", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
ont = Left(sRet, GetPrivateProfileString("Settings", "On Top", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
mmin = Left(sRet, GetPrivateProfileString("Settings", "Minimized", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
pplay = Left(sRet, GetPrivateProfileString("Settings", "PlayList", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
lste = Left(sRet, GetPrivateProfileString("Settings", "List Index", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
TotalNumSongs = Left(sRet, GetPrivateProfileString("Settings", "Total Number Of Songs", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
MainPic = Left(sRet, GetPrivateProfileString("Settings", "Main Pic", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
PlayListPic = Left(sRet, GetPrivateProfileString("Settings", "Playlist Pic", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
AboutPic = Left(sRet, GetPrivateProfileString("Settings", "About Pic", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
SplashPic = Left(sRet, GetPrivateProfileString("Settings", "Splash Pic", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
DirPic = Left(sRet, GetPrivateProfileString("Settings", "Dir Pic", "", sRet, Len(sRet), App.Path & "\" & "MP3-info.INI"))
Me.Picture = LoadPicture(MainPic)
frmPlayList.Picture = LoadPicture(PlayListPic)
frmAbout.Picture = LoadPicture(AboutPic)
frmSplash.Picture = LoadPicture(SplashPic)
frmDir.Picture = LoadPicture(DirPic)
If ret = 1 Then
mnureteat.Checked = True
Reteat = 1
Else
mnureteat.Checked = False
Reteat = 0
End If
If rep = 1 Then
mnurand.Checked = True
Randy = 1
Else
mnurand.Checked = False
Randy = 0
End If
If ont = 1 Then
mnuStay.Checked = True
OnTop = 1
Else
mnuStay.Checked = False
OnTop = 0
End If
If mmin = 1 Then
Min = 1
Timer1.Enabled = True
Else
Min = 0
End If
If pplay = 1 Then
frmPlayList.Show
playlist = 1
Else
playlist = 0
End If
C.InitDir = readINI
frmDir.Dir1.Path = readINI
frmDir.Hide
frmSplash.Visible = True
    List1.Clear
    List2.Clear
    frmPlayList.List1.Clear
    fnum = FreeFile
    Open App.Path & "\" & "MP3-info.INI" For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, a_line
        lines = lines + 1
    Loop
    Close fnum
    numfiles = GetPrivateProfileString( _
        "playlist", "NumberOfEntries", "", _
        buf, Len(buf), App.Path & "\" & "MP3-info.INI")
    If numfiles = "0" Then
    Exit Sub
    End If
Do Until List1.ListCount = lines - 17
    file = "File" & List1.ListCount + 1
        length = GetPrivateProfileString( _
        "playlist", file, "", _
        buf, Len(buf), App.Path & "\" & "MP3-info.INI")
    List1.AddItem Left$(buf, length)
    frmPlayList.List1.AddItem Left$(buf, length)
    List2.AddItem readINI & "\" & Left$(buf, length)
Loop
List1.ListIndex = lste
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu mnu, 1
Else
formdrag Me
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid

mciSendString "stop cd", 0, 0, 0
mciSendString "stop mpeg", 0, 0, 0
mciSendString "close all", 0, 0, 0
WritePrivateProfileString "Settings", "Random", Randy, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Repeat", Reteat, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "On Top", OnTop, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Minimized", Min, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "PlayList", playlist, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "List Index", List1.ListIndex, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Total Number Of Songs", TotalNumSongs, App.Path & "\" & "MP3-info.INI"

For i = 0 To List1.ListCount
    WritePrivateProfileString _
        "playlist", ("File" & i + 1), _
        List1.List(i), App.Path & "\" & "MP3-info.INI"
        Next i
    WritePrivateProfileString _
        "playlist", "NumberOfEntries", _
        List1.ListCount, App.Path & "\" & "MP3-info.INI"
End
End Sub

Private Sub Image1_Click()
On Error Resume Next
Dim ex As String
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
mciSendString "pause cd", 0, 0, 0
Else
mciSendString "pause mpeg", 0, 0, 0
End If
Label1 = List1.Text & " Paused"
 
Playing = False
paused = True
Timer3.Enabled = False
Timer2.Enabled = False

End Sub

Private Sub Image10_Click()
On Error Resume Next
If mnurand.Checked = True Then
Randomize Timer
song = Int((List1.ListCount * Rnd))
List1.ListIndex = song
Image2_Click
Image3_Click
Else
End If
If List1.ListIndex = 0 Then
    List1.ListIndex = List1.ListCount - 1
    Image2_Click
    Image3_Click
    Else
Image2_Click
List1.ListIndex = List1.ListIndex - 1
Image3_Click
End If
TotalNumSongs = TotalNumSongs + 1
End Sub

Private Sub Image2_Click()
On Error Resume Next
Dim ex As String
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
mciSendString "stop cd", 0, 0, 0
mciSendString "close cd", 0, 0, 0
Else
mciSendString "stop mpeg", 0, 0, 0
mciSendString "close mpeg", 0, 0, 0
End If
Label1 = List1.Text & " Stoped"
 
Playing = False
paused = False
Timer2.Enabled = False
Timer3.Enabled = False

End Sub


Private Sub Image3_Click()
On Error Resume Next
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
ext = LCase(Right(List1.Text, 6))
dr = LCase(Left(ext, 2))
    If paused = True Then
        mciSendString "play cd", 0, 0, 0
        Timer3.Enabled = True
        GoTo Q
    Else
    End If
mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
mciSendString "set cd time format tmsf wait", 0, 0, 0
mciSendString "seek cd to " & dr, 0, 0, 0
mciSendString "play cd", 0, 0, 0
TotalNumSongs = TotalNumSongs + 1

Timer3.Enabled = True
Else
    If paused = True Then
        mciSendString "play mpeg", 0, 0, 0
        Timer3.Enabled = True
        GoTo Q
    Else
    End If
If mnurand.Checked = True Then
Randomize Timer
song = Int((List1.ListCount * Rnd))
List1.ListIndex = song
List2.ListIndex = song
frmPlayList.List1.ListIndex = song
Else
End If
ShortPath = GetShortPathName(List2.Text, tmp, 255)
ShortPathAndFie = Left$(tmp, ShortPath)
mciSendString "close mpeg", 0, 0, 0
mciSendString "open " & ShortPathAndFie & " type MPEGVideo Alias mpeg", 0&, 0&, 0&
mciSendString "play mpeg", 0, 0, 0
TotalNumSongs = TotalNumSongs + 1

Min1 = "0"
Min2 = "0"
Sec1 = "0"
Sec2 = "0"
Timer3.Enabled = True
End If
Label1 = "Playing " & List1.Text
Label4 = "Track " & List1.ListIndex + 1 & " of " & List1.ListCount
 
Q:

Playing = True
Timer2.Enabled = True
End Sub


Private Sub Image4_Click()
On Error Resume Next
If frmPlayList.Visible = True Then
frmPlayList.Visible = False
playlist = 0
Else
frmPlayList.Visible = True
playlist = 1
End If
End Sub

Private Sub Image5_Click()
On Error Resume Next
mnuopensong_Click
End Sub

Private Sub Image6_Click()
On Error Resume Next
If mnurand.Checked = True Then
Randomize Timer
song = Int((List1.ListCount * Rnd))
List1.ListIndex = song
Image2_Click
Image3_Click
Else
End If
If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Image2_Click
    Image3_Click
Else
Image2_Click
List1.ListIndex = List1.ListIndex + 1
Image3_Click
End If
TotalNumSongs = TotalNumSongs + 1
End Sub

Private Sub Image8_Click()

mciSendString "stop cd", 0, 0, 0
mciSendString "stop mpeg", 0, 0, 0
mciSendString "close all", 0, 0, 0
WritePrivateProfileString "Settings", "Random", Randy, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Repeat", Reteat, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "On Top", OnTop, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Minimized", Min, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "PlayList", playlist, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "List Index", List1.ListIndex, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Total Number Of Songs", TotalNumSongs, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Main pic", MainPic, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "About pic", AboutPic, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "PlayList pic", PlayListPic, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Splash pic", SplashPic, App.Path & "\" & "MP3-info.INI"
WritePrivateProfileString "Settings", "Dir Pic", DirPic, App.Path & "\" & "MP3-info.INI"

Shell_NotifyIcon NIM_DELETE, nid

For i = 0 To List1.ListCount
    WritePrivateProfileString _
        "playlist", ("File" & i + 1), _
        List1.List(i), App.Path & "\" & "MP3-info.INI"
        Next i
    WritePrivateProfileString _
        "playlist", "NumberOfEntries", _
        List1.ListCount, App.Path & "\" & "MP3-info.INI"
End
End Sub


Private Sub List1_Click()
On Error Resume Next
List2.ListIndex = List1.ListIndex
frmPlayList.List1.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
On Error Resume Next
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
ext = LCase(Right(List1.Text, 6))
dr = LCase(Left(ext, 2))
mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
mciSendString "set cd time format tmsf wait", 0, 0, 0
mciSendString "seek cd to " & dr, 0, 0, 0
mciSendString "play cd", 0, 0, 0

Else
ShortPath = GetShortPathName(List2.Text, tmp, 255)
ShortPathAndFie = Left$(tmp, ShortPath)
mciSendString "close mpeg", 0, 0, 0
mciSendString "open " & ShortPathAndFie & " type MPEGVideo Alias mpeg", 0&, 0&, 0&
mciSendString "play mpeg", 0, 0, 0
End If
Playing = True
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu mnu, 1

End If
End Sub


Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuadddir_Click()
frmDir.Show
frmDir.Label2.Visible = True
frmDir.Label1.Visible = False
End Sub

Private Sub mnuCDClose_Click()
On Error Resume Next
mciSendString "set cd door closed", 0, 0, 0
End Sub

Private Sub mnuCDopen_Click()
On Error Resume Next
mciSendString "set cd door open", 0, 0, 0
End Sub

Private Sub mnumax_Click()
On Error Resume Next
frmPlayList.Hide
frmMain.Show
End Sub

Private Sub mnuClear_Click()
List1.Clear
List2.Clear
frmPlayList.List1.Clear
End Sub

Private Sub mnudir_Click()

End Sub

Private Sub mnuext_Click()
Image8_Click
End Sub

Private Sub mnuList_Click()
C.CancelError = True
On Error Resume Next
Dim buf As String * 256, a_line As String, length As Long, numfiles As String, file, fnum As Integer, lines As Integer, boo, i As Variant
C.Filter = "Sonique Playlist (*.PLS)|*.PLS|Winamp Playlist (*.M3U)|*.M3U" 'winamp or sonique file
C.DialogTitle = "Select a List to Load"
C.ShowOpen
boo = Right(C.FileName, 3)
On Error GoTo err
If boo = "m3u" Then
List1.Clear
    Dim strFileName As String, strText As String, strFilter As String, strBuffer As String, FileHandle%
        strFileName = C.FileName
        FileHandle% = FreeFile
        Open strFileName For Input As #FileHandle%
        Do While Not EOF(FileHandle%)
            
            Line Input #FileHandle%, strBuffer
            List1.AddItem (strBuffer)
    frmPlayList.List1.AddItem (strBuffer)
    List2.AddItem (strBuffer)
            strText = strText & strBuffer & vbCrLf
        Loop
        For i = 0 To List1.ListCount
        boo = Right(List1.List(i), 3)
        If boo = "mp3" Then
        ElseIf i < List1.ListCount Then
        List1.RemoveItem (i)
        List1.Refresh
        End If
        Next i
        List1.RemoveItem (0)
        Close #FileHandle%
        
        ElseIf boo = "PLS" Then
    List1.Clear
    List2.Clear
    frmPlayList.List1.Clear
    fnum = FreeFile
    On Error GoTo err
    Open C.FileName For Input As fnum
    Do While Not EOF(fnum)
    On Error GoTo err
        Line Input #fnum, a_line
        lines = lines + 1
    Loop
    Close fnum
    On Error GoTo err
    numfiles = GetPrivateProfileString( _
        "playlist", "NumberOfEntries", "", _
        buf, Len(buf), C.FileName)
        On Error GoTo err
Do Until List1.ListCount = lines - 3
    file = "File" & List1.ListCount + 1
        length = GetPrivateProfileString( _
        "playlist", file, "", _
        buf, Len(buf), C.FileName)
    List1.AddItem Left$(buf, length)
    frmPlayList.List1.AddItem Left$(buf, length)
    List2.AddItem Left$(buf, length)
Loop
End If
On Error GoTo err
err:
Exit Sub
End Sub

Private Sub mnuMin_Click()
Me.Visible = False
    nid = SetNotifyIconData(Me.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, Me.Icon, "VBAMP" & vbNullChar)
    j = Shell_NotifyIcon(NIM_ADD, nid)
    mnuMin.Visible = False
    mnurest.Visible = True
If frmPlayList.Visible = True Then
playlist = 1
frmPlayList.Visible = False
Else
playlist = 0
End If
Min = 1
End Sub
Private Function SetNotifyIconData(hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, tip As String) As NOTIFYICONDATA
          
    Dim nidTemp As NOTIFYICONDATA
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uId = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallBackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = tip & Chr$(0)
    SetNotifyIconData = nidTemp
          
End Function

Private Sub mnuNext_Click()
On Error Resume Next
If mnurand.Checked = True Then
Randomize Timer
song = Int((List1.ListCount * Rnd))
List1.ListIndex = song
Image2_Click
Image3_Click
Else
End If
If List1.ListIndex = List1.ListCount - 1 Then
    List1.ListIndex = 0
    Image2_Click
    Image3_Click
Else
Image2_Click
List1.ListIndex = List1.ListIndex + 1
Image3_Click
End If
TotalNumSongs = TotalNumSongs + 1
End Sub

Private Sub mnuOpenlist_Click()
On Error Resume Next
Dim i As Variant
C.Filter = "Save Playlist|*.PLS"
C.DialogTitle = "Save As a Play List"
C.ShowSave
Kill (C.FileName)
For i = 0 To List1.ListCount
    WritePrivateProfileString _
        "playlist", ("File" & i + 1), _
        List1.List(i), C.FileName
        Next i
    WritePrivateProfileString _
        "playlist", "NumberOfEntries", _
        List1.ListCount, C.FileName

End Sub

Private Sub mnuopensong_Click()
Dim buf As String * 256, a_line As String, length As Long, numfiles As String, file, fnum As Integer, lines As Integer, ext, i As Variant

On Error Resume Next
C.CancelError = True
On Error GoTo err
C.Filter = "Mp3's (*.Mp3)|*.mp3|Playlist's (*.PLS)|*.PLS|Winamp Playlist (*.M3U)|*.M3U|CD's (*.CDA)|*.cda|Wav's (*,Wav)|*.wav|"
C.ShowOpen
ext = Right(C.FileTitle, 3)
If ext = "M3U" Then
List1.Clear
    Dim strFileName As String, strText As String, strFilter As String, strBuffer As String, FileHandle%
        strFileName = C.FileName
        FileHandle% = FreeFile
        Open strFileName For Input As #FileHandle%
        Do While Not EOF(FileHandle%)
            
            Line Input #FileHandle%, strBuffer
            List1.AddItem (strBuffer)
    frmPlayList.List1.AddItem (strBuffer)
    List2.AddItem (strBuffer)
            strText = strText & strBuffer & vbCrLf
        Loop
        For i = 0 To List1.ListCount
        ext = Right(List1.List(i), 3)
        If ext = "mp3" Then
        ElseIf i < List1.ListCount Then
        List1.RemoveItem (i)
        List1.Refresh
        End If
        Next i
        List1.RemoveItem (0)
        Close #FileHandle%
        
        ElseIf ext = "PLS" Then
    List1.Clear
    List2.Clear
    frmPlayList.List1.Clear
    fnum = FreeFile
    On Error GoTo err
    Open C.FileName For Input As fnum
    Do While Not EOF(fnum)
    On Error GoTo err
        Line Input #fnum, a_line
        lines = lines + 1
    Loop
    Close fnum
    On Error GoTo err
    numfiles = GetPrivateProfileString( _
        "playlist", "NumberOfEntries", "", _
        buf, Len(buf), C.FileName)
        On Error GoTo err
Do Until List1.ListCount = lines - 3
    file = "File" & List1.ListCount + 1
        length = GetPrivateProfileString( _
        "playlist", file, "", _
        buf, Len(buf), C.FileName)
    List1.AddItem Left$(buf, length)
    frmPlayList.List1.AddItem Left$(buf, length)
    List2.AddItem Left$(buf, length)
Loop
Else
List1.AddItem C.FileTitle
frmPlayList.List1.AddItem C.FileTitle
List2.AddItem C.FileName
List1.ListIndex = List1.ListIndex + 1
frmPlayList.List1.ListIndex = frmPlayList.List1.ListIndex + 1
End If
err:
End Sub

Private Sub mnuPause_Click()
On Error Resume Next
Dim ex As String
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
mciSendString "pause cd", 0, 0, 0
Else
mciSendString "pause mpeg", 0, 0, 0
End If
Label1 = List1.Text & " Paused"
 
Playing = False
paused = True
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub mnuPlay_Click()
On Error Resume Next
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
ext = LCase(Right(List1.Text, 6))
dr = LCase(Left(ext, 2))
    If paused = True Then
        mciSendString "play cd", 0, 0, 0
        Timer3.Enabled = True
        GoTo Q
    Else
    End If
mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
mciSendString "set cd time format tmsf wait", 0, 0, 0
mciSendString "seek cd to " & dr, 0, 0, 0
mciSendString "play cd", 0, 0, 0
TotalNumSongs = TotalNumSongs + 1

Timer3.Enabled = True
Else
    If paused = True Then
        mciSendString "play mpeg", 0, 0, 0
        Timer3.Enabled = True
        GoTo Q
    Else
    End If
If mnurand.Checked = True Then
Randomize Timer
song = Int((List1.ListCount * Rnd))
List1.ListIndex = song
Else
End If
ShortPath = GetShortPathName(List2.Text, tmp, 255)
ShortPathAndFie = Left$(tmp, ShortPath)
mciSendString "close mpeg", 0, 0, 0
mciSendString "open " & ShortPathAndFie & " type MPEGVideo Alias mpeg", 0&, 0&, 0&
mciSendString "play mpeg", 0, 0, 0
TotalNumSongs = TotalNumSongs + 1

Min1 = "0"
Min2 = "0"
Sec1 = "0"
Sec2 = "0"
Timer3.Enabled = True
End If
Label1 = "Playing " & List1.Text
Label4 = "Track " & List1.ListIndex + 1 & " of " & List1.ListCount
 
Q:

Playing = True
Timer2.Enabled = True
End Sub

Private Sub mnuPrev_Click()
On Error Resume Next
If mnurand.Checked = True Then
Randomize Timer
song = Int((List1.ListCount * Rnd))
List1.ListIndex = song
Image2_Click
Image3_Click
Else
End If
If List1.ListIndex = 0 Then
    List1.ListIndex = List1.ListCount - 1
    Image2_Click
    Image3_Click
Else
Image2_Click
List1.ListIndex = List1.ListIndex - 1
Image3_Click
End If
TotalNumSongs = TotalNumSongs + 1
End Sub

Private Sub mnurand_Click()
If mnurand.Checked = True Then
mnurand.Checked = False
Randy = 0
Else
mnurand.Checked = True
Randy = 1
End If

End Sub

Private Sub mnuRem_Click()

List1.RemoveItem List1.ListIndex
List2.RemoveItem List1.ListIndex
frmPlayList.List1.RemoveItem frmPlayList.List1.ListIndex
err:

End Sub

Private Sub mnurest_Click()
Me.Show
Shell_NotifyIcon NIM_DELETE, nid
mnurest.Visible = False
mnuMin.Visible = True
If playlist = 1 Then
frmPlayList.Visible = True
End If
Min = 0
End Sub

Private Sub mnureteat_Click()
If mnureteat.Checked = True Then
mnureteat.Checked = False
Reteat = 0
Else
mnureteat.Checked = True
Reteat = 1
End If
End Sub

Private Sub mnusavelist_Click()

End Sub

Private Sub mnusetdir_Click()
frmDir.Show
End Sub

Private Sub mnuskins_Click()
C.CancelError = True
On Error Resume Next
C.Filter = "Bmp's (*.bmp)|*.bmp|jpeg's (*.jpeg)|*.jpeg|jpg's (*.jpg)|*.jpg|All File's (*.*)|*.*|"
C.DialogTitle = "Main Pic"
C.ShowOpen
MainPic = C.FileName
C.DialogTitle = "PlayList Pic"
C.ShowOpen
PlayListPic = C.FileName
C.DialogTitle = "About Pic"
C.ShowOpen
AboutPic = C.FileName
C.DialogTitle = "Dir Pic"
C.ShowOpen
DirPic = C.FileName
C.DialogTitle = "Splash Pic"
C.ShowOpen
SplashPic = C.FileName
Me.Picture = LoadPicture(MainPic)
frmPlayList.Picture = LoadPicture(PlayListPic)
frmAbout.Picture = LoadPicture(AboutPic)
frmSplash.Picture = LoadPicture(SplashPic)
frmDir.Picture = LoadPicture(DirPic)
End Sub

Private Sub mnuStay_Click()
If mnuStay.Checked = True Then
SetWindowPos hwnd, vbNormalFocus, 250, 250, 302, 102, conSwpNoActivate Or conSwpShowWindow
SetWindowPos frmPlayList.hwnd, vbNormalFocus, 250, 350, 300, 150, conSwpNoActivate Or conSwpShowWindow
mnuStay.Checked = False
OnTop = 0
Else
SetWindowPos hwnd, conHwndTopmost, 250, 250, 302, 102, conSwpNoActivate Or conSwpShowWindow
SetWindowPos frmPlayList.hwnd, conHwndTopmost, 250, 350, 300, 150, conSwpNoActivate Or conSwpShowWindow
mnuStay.Checked = True
OnTop = 1
End If
End Sub

Private Sub mnuStop_Click()
On Error Resume Next
Dim ex As String

ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
mciSendString "stop cd", 0, 0, 0
mciSendString "close cd", 0, 0, 0
Else
mciSendString "stop mpeg", 0, 0, 0
mciSendString "close mpeg", 0, 0, 0
End If
Label1 = List1.Text & " Stoped"
 
Playing = False
paused = False
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub Picture1_Click()

End Sub



Private Sub mnuTotal_Click()
MsgBox "You've Played " & TotalNumSongs & " Songs."
End Sub

Private Sub Timer1_Timer()
If frmSplash.Visible = False Then
mnuMin_Click
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()

On Error Resume Next
Dim e As String * 30
If mnureteat.Checked = True Then
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
ext = LCase(Right(List1.Text, 6))
dr = LCase(Left(ext, 2))
mciSendString "status cd mode", e, Len(e), 0
Playing = (Mid$(e, 1, 7) = "playing")
Else
mciSendString "status mpeg mode", e, Len(e), 0
Playing = (Mid$(e, 1, 7) = "playing")
Label2 = Playing
Label4 = "Track " & List1.ListIndex + 1 & " of " & List1.ListCount

End If
If Playing = True Then
Else
    If List1.ListIndex = List1.ListCount - 1 Then
        List1.ListIndex = -1
    Else
    End If
Image2_Click
List1.ListIndex = List1.ListIndex + 1
Image3_Click
Label4 = "Track " & List1.ListIndex + 1 & " of " & List1.ListCount

End If
ElseIf Playing = False Then
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Dim e As String * 30
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
ext = LCase(Right(List1.Text, 6))
dr = LCase(Left(ext, 2))
mciSendString "status cd position", s, Len(s), 0
Min = CInt(Mid$(s, 4, 2))
sec = CInt(Mid$(s, 7, 2))
Label3 = Format(Min, "00") & ":" & Format(sec, "00")
Else
If Sec1 = "9" Then
Sec1 = "0"
Sec2 = Sec2 + 1
Else
Sec1 = Sec1 + 1
End If
If Sec2 = "6" Then
Sec2 = "0"
Min1 = Min1 + 1
End If
If Min1 = "9" Then
Min2 = Min2 + 1
Min1 = 0
End If
Label3 = Min2 & Min1 & ":" & Sec2 & Sec1
End If
End Sub

Private Sub Timer4_Timer()
Label5 = TotalNumSongs
End Sub
