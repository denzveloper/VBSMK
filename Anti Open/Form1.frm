VERSION 5.00
Begin VB.Form Anti 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton antihuruhara 
      Caption         =   "Hlp Me Plz!"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "* Virus Bukan Tanggung Jawab Kami!!"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jangan Lagi Gunakan Komputer Lab Ini Untuk Permainan Atau Yang Lainnya!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   6120
      Width           =   10335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Blocking Screen Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maaf Karena Anda Membuat Kesalahan Dengan Mengeksplor Komputer Ini, Anda Kami Blok Dengan Notice Ini!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   10095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "By: Kuncen LAB TKJ"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jangan Bermain Game , Membuka File, Atau Macam-Macam Dengan Komputer Lab Ini!!!"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10575
   End
End
Attribute VB_Name = "Anti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub antihuruhara_Click()
MsgBox ("Jika Ingin Mengakui Kesalahan Anda.. Harap Hubungi Guru Produktif TKJ atau Ke Kuncen LAB TKJnya Langsung!")
End Sub

Private Sub Form_Load()
MsgBox ("Anda Berbahaya Bagi Kelangsungan Lab Tkj Ini!")
MsgBox ("Anda Akan menerima Sebuah Hukuman! Yaitu Di Screen tertampil Pesan Sangat Besar!!!")
End Sub

