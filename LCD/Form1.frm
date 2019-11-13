VERSION 5.00
Begin VB.Form main_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Layar LCD"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "&About"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      ToolTipText     =   "Exit Application(alt+x)"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Reset"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      ToolTipText     =   "Reset LED View (alt+r)"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox angka 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4560
      List            =   "Form1.frx":0022
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Exit Application(alt+x)"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&View"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "View what happen (alt+v)"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   360
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   360
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   360
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Masukkan Angka:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Program Nyala Layar LCD(untuk Angka saja)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   1920
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   1920
      Top             =   480
      Width           =   255
   End
   Begin VB.Label err 
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
End
Attribute VB_Name = "main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
err = ""
Shape1.BackColor = vbWhite
Shape2.BackColor = vbWhite
Shape3.BackColor = vbWhite
Shape4.BackColor = vbWhite
Shape5.BackColor = vbWhite
Shape6.BackColor = vbWhite
Shape7.BackColor = vbWhite
Shape5.BorderColor = vbBlack
Shape2.BorderColor = vbBlack
err.Caption = ""
If angka = "0" Then
Shape1.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape5.BackColor = vbBlack
Shape7.BackColor = vbBlack
Shape4.BackColor = vbBlack
Shape3.BackColor = vbBlack
ElseIf angka = "1" Then
Shape2.BackColor = vbBlack
Shape5.BackColor = vbBlack
ElseIf angka = "2" Then
Shape1.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape4.BackColor = vbBlack
Shape7.BackColor = vbBlack
ElseIf angka = "3" Then
Shape1.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape5.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape7.BackColor = vbBlack
ElseIf angka = "4" Then
Shape3.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape5.BackColor = vbBlack
ElseIf angka = "5" Then
Shape1.BackColor = vbBlack
Shape3.BackColor = vbBlack
Shape5.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape7.BackColor = vbBlack
ElseIf angka = "6" Then
Shape1.BackColor = vbBlack
Shape3.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape5.BackColor = vbBlack
Shape7.BackColor = vbBlack
Shape4.BackColor = vbBlack
ElseIf angka = "7" Then
Shape1.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape5.BackColor = vbBlack
ElseIf angka = "8" Then
Shape1.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape3.BackColor = vbBlack
Shape4.BackColor = vbBlack
Shape5.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape7.BackColor = vbBlack
ElseIf angka = "9" Then
Shape1.BackColor = vbBlack
Shape2.BackColor = vbBlack
Shape3.BackColor = vbBlack
Shape5.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape7.BackColor = vbBlack
Else
Shape1.BackColor = vbBlack
Shape3.BackColor = vbBlack
Shape4.BackColor = vbBlack
Shape6.BackColor = vbBlack
Shape7.BackColor = vbBlack
Shape5.BorderColor = vbWhite
Shape2.BorderColor = vbWhite
err.Caption = "ROR"
MsgBox ("Error Masukkan Tidak sah!")
End If
End Sub

Private Sub Command2_Click()
Confirm.Show
End Sub

Private Sub Command3_Click()
angka = "0"
err = ""
Shape1.BackColor = vbWhite
Shape2.BackColor = vbWhite
Shape3.BackColor = vbWhite
Shape4.BackColor = vbWhite
Shape5.BackColor = vbWhite
Shape6.BackColor = vbWhite
Shape7.BackColor = vbWhite
Shape5.BorderColor = vbBlack
Shape2.BorderColor = vbBlack
End Sub

Private Sub Command4_Click()
about.Show
End Sub
