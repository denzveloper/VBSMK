VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FF9900&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   6300
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   HasDC           =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   14  'Arrow and Question
   Moveable        =   0   'False
   ScaleHeight     =   315
   ScaleMode       =   2  'Point
   ScaleWidth      =   286.5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   465
      Left            =   2280
      MousePointer    =   2  'Cross
      Picture         =   "frmAbout.frx":0ECA
      Top             =   5760
      Width           =   765
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmAbout.frx":21F0
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmAbout.frx":223E
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   282
      X2              =   6
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   6
      X2              =   6
      Y1              =   120
      Y2              =   84
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   6
      X2              =   282
      Y1              =   84
      Y2              =   84
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   282
      X2              =   282
      Y1              =   120
      Y2              =   84
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   120
      Picture         =   "frmAbout.frx":228A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF9900&
      Caption         =   "GNU General Public Licensed Version 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FF9900&
      Caption         =   $"frmAbout.frx":3154
      ForeColor       =   &H00FFFFFF&
      Height          =   2745
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   4.5
      X2              =   282.7
      Y1              =   122.25
      Y2              =   122.25
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FF9900&
      Caption         =   $"frmAbout.frx":33F0
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   4605
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF9900&
      Caption         =   "Kalkulator Denzveloper GNU For Windows"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   5.25
      X2              =   282.7
      Y1              =   123
      Y2              =   123
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FF9900&
      Caption         =   "V 3.0.2"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox ("Build By Denzveloper-TKJ")
End Sub

Private Sub Image2_Click()
Form1.Show
Unload Me
End Sub

