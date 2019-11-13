VERSION 5.00
Begin VB.Form frmbelanja 
   Caption         =   "Program Belanja Sederhana."
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdulang 
      Caption         =   "&Ulang"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   615
      Left            =   3960
      TabIndex        =   15
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operasi"
      Height          =   735
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
      Begin VB.CommandButton cmdhitung 
         Caption         =   "&Hitung"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtbonus 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox txtbayar 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txtdiskon 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txttotal 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox txtjumlah 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtharga 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtnama 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Bonus                      :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Total Bayar              :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Diskon                     :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Total Harga              :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah Barang          :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Harga Barang           :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Barang           :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmbelanja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhitung_Click()
'deklarasi variabel
Dim harga, jumlah As Integer
Dim total, diskon, bayar As Double
Dim bonus As String

End Sub

Private Sub cmdulang_Click()
'membersihkan textbox
txtnama.Text = ""
txtharga.Text = ""
txtjumlah.Text = ""
txttotal.Text = ""
txtdiskon.Text = ""
txtbayar.Text = ""
txtbonus.Text = ""
txtnama.SetFocus
End Sub
