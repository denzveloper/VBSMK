VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program serangan fajar"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-1"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Reset"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   1785
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      TabIndex        =   6
      Top             =   1440
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mulai"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Ditujukan ke:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Pengulangan  :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Inputan serangan :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim teks As String
Dim jml As Integer
teks = "Hey kau " + Text3.Text + "!! " + Text1.Text
jml = Val(Text2)
For i = 1 To jml
List1.AddItem teks
Next i
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
List1.Clear
Text1.SetFocus
End Sub

Private Sub Command4_Click()
List1.RemoveItem (0)
End Sub
