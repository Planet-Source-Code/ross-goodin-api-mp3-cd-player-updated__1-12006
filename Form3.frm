VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   0  'None
   ClientHeight    =   2280
   ClientLeft      =   4845
   ClientTop       =   5670
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   2280
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   765
      Left            =   160
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   160
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   4140
      Picture         =   "Form3.frx":20F9A
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Const conHwndTopmost = -1
    Private Const conSwpNoActivate = &H10
    Private Const conSwpShowWindow = &H40

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
 On Error Resume Next
SetWindowPos hwnd, conHwndTopmost, 250, 350, 300, 150, conSwpNoActivate Or conSwpShowWindow

End Sub

Private Sub Image1_Click()
Me.Hide
End Sub

Private Sub Label1_Click()
WritePrivateProfileString "Settings", "File Dir", Dir1.Path, App.Path & "\" & "Mp3-info.INI"
frmMain.C.InitDir = Dir1.Path
Me.Hide
End Sub

Private Sub Label2_Click()
tmpStrg = Dir$(Dir1.Path & "\*.mp3")
If tmpStrg <> "" Then
mp3FileName = tmpStrg
frmMain.List1.AddItem mp3FileName
frmMain.List2.AddItem Dir1.Path & "\" & mp3FileName
frmPlayList.List1.AddItem mp3FileName
tmpStrg = Dir$
While Len(tmpStrg) > 0
mp3FileName = tmpStrg

frmMain.List1.AddItem mp3FileName
frmMain.List2.AddItem Dir1.Path & "\" & mp3FileName
frmPlayList.List1.AddItem mp3FileName
tmpStrg = Dir$
Wend
Else
End If
Label2.Visible = False
Label1.Visible = True
Me.Hide
End Sub
