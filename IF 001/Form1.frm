VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Abal-Abal"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton reset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdproses 
      Caption         =   "&Proses"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox nilai 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox nama 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblhasil 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Keterangan:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Nilai Siswa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Nama Siswa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdproses_Click()
If nilai.Text > 100 Then
lblhasil.Caption = "Masyaallah Gue kagak bisa ngitung kalo angkany`e Lebih dari 100."
ElseIf nilai.Text < 1 Then
lblhasil.Caption = "Dengan pertimbangan Tuhan YME, Kasian Deh Loe!! :p"
ElseIf nilai.Text >= 75 Then
lblhasil.Caption = "Dengan pertimbangan Tuhan YME, si: " + nama.Text + " Ternyata: 'Lulus'"
Else
lblhasil.Caption = "Dengan pertimbangan Tuhan YME, si: " + nama.Text + "Ternyata: 'Belum Lulus'"
End If
End Sub

Private Sub reset_Click()
lblhasil.Caption = ""
nama.Text = ""
nilai.Text = ""
nama.SetFocus
End Sub
