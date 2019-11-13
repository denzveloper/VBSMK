VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Effek Teks"
      Height          =   2775
      Left            =   4560
      TabIndex        =   5
      Top             =   0
      Width           =   2535
      Begin VB.CheckBox Check2 
         Caption         =   "tebal"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "miring"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Biru"
         Height          =   375
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Merah"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "keluar"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "masukkan nama anda:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Label2.FontItalic = Check1.Value
End Sub

Private Sub Check2_Click()
Label2.FontBold = Check2.Value
End Sub

Private Sub Command1_Click()
Label2.Caption = Text1.Text
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Option1_Click()
Label2.ForeColor = &H888888
End Sub

Private Sub Option2_Click()
Label2.ForeColor = vbBlue
End Sub
