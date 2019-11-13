VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Daftar Penduduk"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit 
      Caption         =   "K&eluar"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton send 
      Caption         =   "&Kirim"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hasil:"
      Height          =   4695
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   5415
      Begin VB.Label out 
         Height          =   4335
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.ComboBox jk 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1680
      List            =   "Form1.frx":000A
      TabIndex        =   11
      Text            =   "Jenis Kelamin"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox tgl_c 
      Height          =   285
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox tgl_b 
      Height          =   315
      ItemData        =   "Form1.frx":0024
      Left            =   2280
      List            =   "Form1.frx":004C
      TabIndex        =   7
      Text            =   "Bulan Kelahiran"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox tgl_a 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.ComboBox kelas_c 
      Height          =   315
      ItemData        =   "Form1.frx":00B3
      Left            =   3480
      List            =   "Form1.frx":00C0
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox kelas_b 
      Height          =   315
      ItemData        =   "Form1.frx":00CD
      Left            =   2280
      List            =   "Form1.frx":00E3
      TabIndex        =   4
      Text            =   "Jurusan"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox kelas_a 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox nama 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Jenis Kelamin:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal Lahir:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Kelas:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nama:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
Unload Me
End Sub

Private Sub send_Click()
out.Caption = "Nama:" + nama.Text + ". Jenis kelamin:" + jk.Text + ". Kelas:" + kelas_a.Text + "-" + kelas_b.Text + "-" + kelas_c.Text + ". Tanggal Lahir:" + tgl_a.Text + " " + tgl_b.Text + " " + tgl_c.Text
End Sub
    
