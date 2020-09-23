VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   4410
   ClientTop       =   2070
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsplach.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmsplach.frx":000C
   ScaleHeight     =   4500
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim PlY As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
frmSplash.Visible = False
frmMain.Show
If PlY = 1 Then
frmPlayList.Visible = True
Else
frmPlayList.Visible = False
End If
End Sub

Private Sub Form_Load()
If frmPlayList.Visible = True Then
frmPlayList.Visible = False
PlY = 1
Else
PlY = 0
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.Show
frmSplash.Visible = False
If PlY = 1 Then
frmPlayList.Visible = True
Else
frmPlayList.Visible = False
End If
End Sub
