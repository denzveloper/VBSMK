VERSION 5.00
Begin VB.Form look 
   Caption         =   "Lihat data"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Tutup"
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3840
      TabIndex        =   11
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label10 
      Height          =   3015
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   6615
   End
   Begin VB.Label Label9 
      Caption         =   "Alamat:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah Barang:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Harga Barang:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Barang:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Barang:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "look"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = awal.Text1.Text
Label6.Caption = awal.Text2.Text
Label7.Caption = awal.Text3.Text
Label8.Caption = awal.Text4.Text
Label10.Caption = "Jl. " + awal.Text5.Text + " RT/RW:" + awal.Text6.Text + "/" + awal.Text7.Text + " Ds. " + awal.Text8.Text + " Kec. " + awal.Text9.Text + " - " + awal.Text10.Text + " Kode Pos:" + awal.Text11.Text
End Sub
