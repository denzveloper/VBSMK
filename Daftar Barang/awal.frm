VERSION 5.00
Begin VB.Form awal 
   Caption         =   "Data Barang"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Kontrol Menu"
      Height          =   2175
      Left            =   6600
      TabIndex        =   24
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Lihat Data"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alamat Pemesan:"
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   10455
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   6960
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   9000
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label12 
         Caption         =   "Kode pos"
         Height          =   255
         Left            =   6960
         TabIndex        =   23
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Kabupaten"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Kecamatan"
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Desa"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "RW"
         Height          =   255
         Left            =   9000
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "RT"
         Height          =   255
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Jalan"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah Barang:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Rp."
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Harga Barang:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Barang:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Barang:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "awal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
look.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
