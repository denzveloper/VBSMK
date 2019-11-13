VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Penghitung Luas "
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton out 
      Caption         =   "exit"
      Height          =   495
      Left            =   7320
      TabIndex        =   21
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton About 
      Caption         =   "Tentang..."
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   3135
   End
   Begin VB.Frame pp 
      Caption         =   "Persegi Panjang"
      Height          =   1935
      Index           =   1
      Left            =   4200
      TabIndex        =   8
      Top             =   2760
      Width           =   3975
      Begin VB.TextBox lbr 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox hasilpp 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CommandButton hitungpp 
         Caption         =   "Hitung!"
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox pjg 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Hasil:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Lebar:"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Panjang:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame persegi 
      Caption         =   "Luas Persegi"
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   8175
      Begin VB.Frame persegi 
         Caption         =   "Persegi"
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3975
         Begin VB.TextBox hasp 
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CommandButton hitungp 
            Caption         =   "Hitung!"
            Height          =   375
            Left            =   1080
            TabIndex        =   10
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox sisi 
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label6 
            Caption         =   "Hasil:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   615
         End
      End
   End
   Begin VB.Frame lingk 
      Caption         =   "Luas Lingkaran"
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.TextBox hasling 
         Height          =   1575
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton hitungling 
         Caption         =   "Hitung!"
         Height          =   615
         Left            =   3120
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox jari 
         Height          =   1575
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Jari-Jari Lingkaran"
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         Caption         =   "Hasil:"
         Height          =   1935
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
MsgBox ("Penghitung Luas G-TECH BD :: Version 1.0 Beta By: Dendy Octavian")
End Sub

Private Sub hitungling_Click()
hasling = 22 / 7 * (Val(jari) ^ 2)
End Sub

Private Sub hitungp_Click()
hasp = Val(sisi) ^ 2
End Sub

Private Sub hitungpp_Click()
hasilpp = Val(pjg) * Val(lbr)
End Sub

Private Sub out_Click()
Unload Me
End Sub
