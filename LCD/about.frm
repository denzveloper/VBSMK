VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About LED View"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit About"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"about.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   0
      Picture         =   "about.frx":01F8
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload about
End Sub

