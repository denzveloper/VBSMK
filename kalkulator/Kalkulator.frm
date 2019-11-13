VERSION 5.00
Begin VB.Form Kalkulator 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kalkulator :: G-Tech BD version 1.2"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Kalkulator.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Kalkulator.frx":164A
   PaletteMode     =   2  'Custom
   Picture         =   "Kalkulator.frx":2C94
   ScaleHeight     =   4305
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton out 
      Appearance      =   0  'Flat
      Caption         =   "E&xit!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MousePointer    =   12  'No Drop
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Keluar dari Program"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lingkaran"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3000
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
      Begin VB.CommandButton hling 
         Caption         =   "hitung"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Menghitung luas lingkaran."
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Masukkan Jari-Jari Di Form Kiri"
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton hlp 
      Caption         =   "Cara Pakai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cara Pakai Program"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton ttg 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Mengenai Aplikasi Ini."
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cpy 
      Caption         =   "&Copy ke form Kiri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MouseIcon       =   "Kalkulator.frx":582FA
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salin Form Jumlah kekiri untuk menghitung nilai selanjutnya."
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CommandButton reset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Mengkosongkan kembali semua form."
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox b 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1935
      Left            =   6000
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      ToolTipText     =   "Form Isian B"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox jawab 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Hasil Dari Penghitungan."
      Top             =   120
      Width           =   6855
   End
   Begin VB.TextBox a 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1935
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      ToolTipText     =   "Form Isian A"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton pangkat 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Perpangkatan"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton bagi 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pembagian"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Kali 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Perkalian"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton kurang 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Pengurangan"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Tambah 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Penjumlahan"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Build With Visual Basic"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Kalkulator :: By: G-Tech BD(Denzveloper). This Application Is Freeware!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5415
   End
End
Attribute VB_Name = "Kalkulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bagi_Click()
If Val(a) + Val(b) = 0 Then
MsgBox ("Semua Form Tidak terisi atau diisi dengan angka Nol(0)!")
jawab.ForeColor = vbRed
jawab = "Tidak Dapat Dihitung!"
ElseIf Val(a) = 0 Then
jawab = "0"
ElseIf Val(b) = 0 Then
MsgBox ("Tidak Dapat Melakukan Pembagian karena Form ''B'' Berisi Nilai NOL(0)!")
jawab.ForeColor = vbRed
jawab = "Tidak Dapat Dibagi Nol(0)!"
Else
jawab = Val(a) / Val(b)
End If
End Sub

Private Sub cpy_Click()
a = Val(jawab)
jawab = ""
b = ""
End Sub

Private Sub hling_Click()
If Val(a) = 0 Then
jawab.ForeColor = vbRed
jawab = "Nilai Belum Terisi!"
If Val(b) <> 0 Then
jawab.ForeColor = vbRed
b.ForeColor = vbRed
jawab = "Jangan Isi Disitu!"
b = "Bukan Ini"
End If
Else
jawab = 22 / 7 * (Val(a) ^ 2)
End If
End Sub

Private Sub hlp_Click()
jawab.ForeColor = vbGreen
jawab = "Ini Adalah Form Hasil.."
a.ForeColor = vbGray
b.ForeColor = vbGray
a = "Form isi(A)"
b = "Form isi(B)"
MsgBox ("Masukkan Angka Di Kedua Form Lalu Gunakan Tombol seperti +,-,:, atau ^ untuk Menghitung Nilai.")
a = ""
b = ""
jawab = ""
End Sub

Private Sub Kali_Click()
jawab = Val(a) * Val(b)
End Sub

Private Sub kurang_Click()
jawab = Val(a) - Val(b)
End Sub

Private Sub out_Click()
Unload Me
End Sub

Private Sub out_GotFocus()
Label3 = "Please Don't Left Me!!"
End Sub

Private Sub pangkat_Click()
jawab = Val(a) ^ Val(b)
End Sub

Private Sub pi_Click()
a = 22 / 7
End Sub

Private Sub reset_Click()
a = ""
b = ""
jawab = ""
a.ForeColor = vbBlack
b.ForeColor = vbBlack
jawab.ForeColor = vbBlack
End Sub

Private Sub Tambah_Click()
jawab = Val(a) + Val(b)
End Sub

Private Sub ttg_Click()
jawab = "G-TECH BD :: Denzveloper   "
a = "2015"
b = "XI-TKJ-1"
Label1 = "SISWA SMK NEGERI 1 JAMBLANG"
Label3 = "Thanks For Using This Application"
MsgBox ("Kalkulator :: Product By: G-Tech :: Developed by: Denzveloper.  Lisensi: LGPL/GPL versi 2 atau Selanjutnya.")
a = ""
b = ""
jawab = ""
Label1 = "Kalkulator :: By: G-TECH BD(Denzveloper). This Application Is Freeware!"
Label3 = "Build With Visual Basic"
End Sub
