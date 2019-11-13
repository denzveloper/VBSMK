VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6240
      Top             =   4680
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   1095
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xspeed As Integer
Dim yspeed As Integer
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'vbKeyF4 115 f4 key
If Shift = vbAltMask Then
If KeyCode = vbKeyF4 Then
End
End If
End If
If KeyCode = vbKeyEscape Then
End
End If
End Sub
Private Sub Form_Load()
xspeed = 10
yspeed = 10
End Sub
Private Sub Timer1_Timer()
Shape1.Left = Shape1.Left + xspeed
Shape1.Top = Shape1.Top + yspeed
If Shape1.Left > (Form1.Width - Shape1.Width - (Form1.Width - Form1.ScaleWidth))
Then
xspeed = xspeed * -1
End If
If Shape1.Top > (Form1.Height - Shape1.Height - (Form1.Height - Form1.Height)) Then
yspeed = yspeed * -1
End If
If Shape1.Left < 0 Then
xspeed = xspeed * -1
End If
If Shape1.Top < 0 Then
yspeed = yspeed * -1
End If
End Sub
