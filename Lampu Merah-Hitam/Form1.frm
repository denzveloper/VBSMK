VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "K&eterangan"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1920
      List            =   "Form1.frx":000D
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   375
      Left            =   600
      Shape           =   2  'Oval
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   375
      Left            =   600
      Shape           =   2  'Oval
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   375
      Left            =   600
      Shape           =   2  'Oval
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Lampu Lalu lintas:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   2055
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   1575
      Left            =   240
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Shape2.BackColor = vbWhite
Shape3.BackColor = vbWhite
Shape4.BackColor = vbWhite
If Combo1 = "Merah" Then
Label1.Caption = "Lampunya Warna Merah :: Berhenti dulu coy!"
Shape2.BackColor = vbRed
ElseIf Combo1 = "Kuning" Then
Label1.Caption = "Lampunya Warna Kuning  :: Bersedia..."
Shape3.BackColor = vbYellow
ElseIf Combo1 = "Hijau" Then
Label1.Caption = "Lampunya Warna Kuning  :: Gage Mangkat Coy!!"
Shape4.BackColor = vbGreen
Else
Label1.Caption = "ERROR!"
Shape2.BackColor = vbRed
Shape3.BackColor = vbYellow
Shape4.BackColor = vbGreen
End If
End Sub

Private Sub Command2_Click()
Shape2.BackColor = vbWhite
Shape3.BackColor = vbWhite
Shape4.BackColor = vbWhite
Label1.Caption = ""
Combo1 = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
