VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Pilih Data Output"
      Height          =   1335
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "&Farenheit"
         Height          =   195
         Left            =   1200
         MousePointer    =   2  'Cross
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "&Reamur"
         Height          =   195
         Left            =   240
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "&Celcius"
         Height          =   195
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilih Data Input"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Far&enheit"
         Height          =   315
         Left            =   1200
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Re&amur"
         Height          =   315
         Left            =   240
         MousePointer    =   2  'Cross
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Celciu&s"
         Height          =   315
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   4200
      MousePointer    =   12  'No Drop
      Picture         =   "Form1.frx":0000
      ToolTipText     =   "Exit Program"
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Program Konversi Suhu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rumus()
If Option1.Value = True And Option4.Value = True Then
Text2.Text = Text1.Text
ElseIf Option1.Value = True And Option5.Value = True Then
Text2.Text = Val(Text1.Text) * 4 / 5
ElseIf Option1.Value = True And Option6.Value = True Then
Text2.Text = (Val(Text1.Text) * 9 / 5) + 32
End If
If Option2.Value = True And Option5.Value = True Then
Text2.Text = Text1.Text
ElseIf Option2.Value = True Then
Text2.Text = Val(Text1.Text) * 5 / 4
ElseIf Option2.Value = True Then
Text2.Text = (9 / 4 * Val(Text1.Text)) + 32
End If
If Option3.Value = True And Option6.Value = True Then
Text2.Text = Text1.Text
ElseIf Option3.Value = True And Option5.Value = True Then
Text2.Text = 5 / 9 * (Val(Text1.Text) - 32)
ElseIf Option3.Value = True And Option4.Value = True Then
Text2.Text = 4 / 9 * (Val(Text1.Text) - 32)
End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Option1_Click()
Call rumus
End Sub

Private Sub Option2_Click()
Call rumus
End Sub

Private Sub Option3_Click()
Call rumus
End Sub

Private Sub Option4_Click()
Call rumus
End Sub

Private Sub Option5_Click()
Call rumus
End Sub

Private Sub Option6_Click()
Call rumus
End Sub

Private Sub Text1_Change()
Call rumus
End Sub
