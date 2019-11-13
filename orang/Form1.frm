VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Isian Siswa"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Kontrol Form"
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   8040
      TabIndex        =   30
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command3 
         BackColor       =   &H000080FF&
         Caption         =   "&About"
         Height          =   615
         Left            =   360
         TabIndex        =   39
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000080FF&
         Caption         =   "T&utup"
         Height          =   615
         Left            =   360
         TabIndex        =   32
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "&Tampilkan Data"
         Height          =   615
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Hasil"
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   10695
      Begin VB.Label Label22 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label Label21 
         BackColor       =   &H00808080&
         Caption         =   "Hobby"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2280
         TabIndex        =   36
         Top             =   1800
         Width           =   7935
      End
      Begin VB.Label Label19 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   960
         Width           =   6135
      End
      Begin VB.Label Label18 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label Label16 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label17 
         BackColor       =   &H00808080&
         Caption         =   "Alamat"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00808080&
         Caption         =   "Tempat Tanggal Lahir"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808080&
         Caption         =   "Nama"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         Caption         =   "NIS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text8 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text13 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox Text11 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox Text10 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text9 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text7 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text5 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text4 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   720
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10800
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      Caption         =   "Kodepos"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   "Kecamatan"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808080&
      Caption         =   "Desa/Kelurahan"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7680
      TabIndex        =   22
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      Caption         =   "RW"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Caption         =   "RT"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Caption         =   "Jalan"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "Alamat"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Hobby"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Tanggal-Bulan-Tahun Lahir"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Tempat Lahir"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Nama"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "NIS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label16.Caption = Form1.Text1.Text
Label18.Caption = Form1.Text2.Text
Label19.Caption = Form1.Text3.Text + "," + Form1.Text4.Text + "-" + Form1.Text4.Text + "-" + Form1.Text6.Text
Label22.Caption = Form1.Text7.Text
Label20.Caption = "Jl." + Form1.Text8.Text + " RT/RW:" + Form1.Text9.Text + "/" + Form1.Text10.Text + " Ds." + Form1.Text11.Text + " Kec." + Form1.Text12.Text + " Kodepos:" + Form1.Text13.Text
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Load frmAbout
End Sub
