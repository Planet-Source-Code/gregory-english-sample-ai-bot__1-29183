VERSION 5.00
Begin VB.Form frmBot 
   BackColor       =   &H00000000&
   Caption         =   "Bot Movement"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BotPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1200
      Picture         =   "frmBot.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Menu mnuStart 
      Caption         =   "Start"
   End
   Begin VB.Menu mnuStop 
      Caption         =   "Stop"
   End
   Begin VB.Menu mnuStates 
      Caption         =   "States"
      Begin VB.Menu mnuSeek 
         Caption         =   "Seek"
      End
      Begin VB.Menu mnuFlee 
         Caption         =   "Flee"
      End
      Begin VB.Menu mnuRest 
         Caption         =   "Rest"
      End
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//Player Location
Dim Px As Integer
Dim Py As Integer

'//Bot Variables
Dim BotX As Integer
Dim BotY As Integer
Dim BotState As Integer

'//Bot Constants
Const S_SEEK = 1
Const S_FLEE = 2
Const S_REST = 3

'//Run Variables
Const LOOPRUNS = 60
Dim counter As Integer
Dim run As Boolean

Sub MoveBot()
'//These hold distances between the playerx and botx
'//and playery and boty
Dim Dx As Integer, Dy As Integer

Dx = Px - BotX 'The distances
Dy = Py - BotY '''''''''''''''

'//What state are we in?
Select Case BotState
    
    Case S_SEEK
        '//Find our player
        '//The X Axis
        If Px > BotX Then
            BotX = BotX + 1
        ElseIf Px < BotX Then
            BotX = BotX - 1
        End If
        
        
        '//The Y Axis
        If Py > BotY Then
            BotY = BotY + 1
        ElseIf Py < BotY Then
            BotY = BotY - 1
        End If
        
    Case S_FLEE
        '//Find our player..then RUN the other way
        '//The X Axis
            If Px > BotX Then
                BotX = BotX - 1
            ElseIf Px < BotX Then
                BotX = BotX + 1
            End If
        
        '//The Y Axis
            If Py > BotY Then
                BotY = BotY - 1
            ElseIf Py < BotY Then
                BotY = BotY + 1
            End If
            
    Case S_REST
    'Do NOTHING!
End Select

'Do Not Allow To Go Off Of Form Screen
If BotX > frmBot.ScaleWidth Then
    BotX = frmBot.ScaleWidth - BotPic.ScaleWidth
ElseIf BotX < 0 Then
    BotX = 0
End If

If BotY > frmBot.ScaleHeight Then
    BotY = frmBot.ScaleHeight - BotPic.ScaleHeight
ElseIf BotY < 0 Then
    BotY = 0
End If

BotPic.Left = BotX
BotPic.Top = BotY

End Sub

Private Sub Form_Load()
counter = 0
run = True
BotState = S_REST
BotX = BotPic.Left
BotY = BotPic.Top

Me.Show
Do
    If counter = LOOPRUNS Then
        If run = True Then
            Me.Cls
            MoveBot
            Line (BotX, BotY)-(Px, Py), vbRed
        End If
    counter = 0
    End If
    counter = counter + 1
    DoEvents
Loop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Px = X
Py = Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Px = X
Py = Y
End If

End Sub

Private Sub mnuFlee_Click()
BotState = S_FLEE
End Sub

Private Sub mnuRest_Click()
BotState = S_REST
End Sub

Private Sub mnuSeek_Click()
BotState = S_SEEK
End Sub

Private Sub mnuStart_Click()
run = True
End Sub

Private Sub mnuStop_Click()
run = False
End Sub
