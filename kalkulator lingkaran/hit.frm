VERSION 5.00
Begin VB.Form hit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menghitung Lingkaran GUI::Visual Basic"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "M"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cls 
      Caption         =   "Bersihkan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox jum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Count 
      Caption         =   "Hitung!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox jr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jari-Jari:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   3120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   4935
      Left            =   480
      Shape           =   2  'Oval
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "hit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cls_Click()
jum = ""
jr = ""
End Sub

Private Sub Count_Click()
jum = 22 / 7 * (jr ^ 2)
End Sub
