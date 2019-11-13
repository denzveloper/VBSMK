VERSION 5.00
Begin VB.Form Confirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit Application?"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton n 
      Caption         =   "&Cancel, no exit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Thank you. you not to leave me.. :)"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton y 
      Caption         =   "&Yes, Exit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "I verry sad you leave me.. :'("
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Are You really leave me?"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Confirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub n_Click()
Unload Me
End Sub

Private Sub y_Click()
Unload Confirm
Unload main_Menu
End Sub
