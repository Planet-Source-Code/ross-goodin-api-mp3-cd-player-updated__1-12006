VERSION 5.00
Begin VB.Form frmPlayList 
   BorderStyle     =   0  'None
   ClientHeight    =   2295
   ClientLeft      =   2910
   ClientTop       =   5430
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "playlist2.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   1590
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4065
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   4140
      Picture         =   "playlist2.frx":20F9A
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
Dim Playing As Boolean
Dim paused As Boolean
Dim s As String * 30
Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hwnd, &HA1, 2, 0&)
End Sub

Private Sub Form_Load()
 On Error Resume Next
SetWindowPos hwnd, conHwndTopmost, 250, 350, 300, 150, conSwpNoActivate Or conSwpShowWindow
mciSendString "open cdaudio Alias cd wait shareable", 0, 0, 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu mnu, 1
Else
formdrag Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
mciSendString "close all", 0, 0, 0
End Sub

Private Sub Image1_Click()
On Error Resume Next
Me.Visible = False
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
End Sub


Private Sub Image3_Click()
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

End Sub

Private Sub Image5_Click()
On Error Resume Next
mciSendString "set cd door open", 0, 0, 0
End Sub

Private Sub Image6_Click()
On Error Resume Next
mciSendString "set cd door closed", 0, 0, 0
End Sub

Private Sub Image8_Click()
On Error Resume Next
mciSendString "stop cd", 0, 0, 0
mciSendString "stop mpeg", 0, 0, 0
mciSendString "close all", 0, 0, 0
End
End Sub

Private Sub Label3_Click()
On Error Resume Next
Me.Hide
End Sub

Private Sub List1_Click()
On Error Resume Next
frmMain.List2.ListIndex = frmPlayList.List1.ListIndex
frmMain.List1.ListIndex = frmPlayList.List1.ListIndex
End Sub

Private Sub List1_DblClick()
On Error Resume Next
ex = LCase(Right(List1.Text, 4))
If ex = ".cda" Then
ext = LCase(Right(List1.Text, 6))
dr = LCase(Left(ext, 2))
    If paused = True Then
        mciSendString "play cd", 0, 0, 0
        frmMain.Timer3.Enabled = True
    Else
    End If
mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
mciSendString "set cd time format tmsf wait", 0, 0, 0
mciSendString "seek cd to " & dr, 0, 0, 0
mciSendString "play cd", 0, 0, 0
frmMain.Timer3.Enabled = True
Else
    If paused = True Then
        mciSendString "play mpeg", 0, 0, 0
        frmMain.Timer3.Enabled = True
        GoTo Q
    Else
    End If
ShortPath = GetShortPathName(frmMain.List2.Text, tmp, 255)
ShortPathAndFie = Left$(tmp, ShortPath)
mciSendString "close mpeg", 0, 0, 0
mciSendString "open " & ShortPathAndFie & " type MPEGVideo Alias mpeg", 0&, 0&, 0&
mciSendString "play mpeg", 0, 0, 0
frmMain.Timer3.Enabled = True
End If
Q:
frmMain.Label1 = "Playing " & List1.Text
frmMain.Label4 = "Track " & List1.ListIndex + 1 & " of " & List1.ListCount

Playing = True
frmMain.Timer2.Enabled = True
frmMain.Label5.Caption = frmMain.Label5.Caption + 1
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
'remove file
mnuRem_Click
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu frmMain.mnu, 1
End If
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

Private Sub mnuNext_Click()
On Error Resume Next
Call mnuStop_Click
List1.ListIndex = List1.ListIndex + 1
Call List1_DblClick
Timer1.Enabled = False
err:
List1.ListIndex = 0
End Sub

Private Sub mnuopensong_Click()
On Error Resume Next

C.CancelError = True
On Error GoTo err
C.Filter = "Mp3's (*.Mp3)|*.mp3|CD's (*.CDA)|*.cda|Wav's (*,Wav)|*.wav|"
C.ShowOpen
List1.AddItem C.FileTitle
frmPlayList.List1.AddItem C.FileTitle
List2.AddItem C.FileName
List1.ListIndex = List1.ListIndex + 1
frmPlayList.List1.ListIndex = frmPlayList.List1.ListIndex + 1
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
End Sub

Private Sub mnuPlay_Click()
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
End Sub

Private Sub mnuRem_Click()
On Error GoTo err
frmMain.List1.RemoveItem List1.ListIndex
frmMain.List2.RemoveItem List1.ListIndex
List1.RemoveItem List1.ListIndex
err:
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
End Sub

