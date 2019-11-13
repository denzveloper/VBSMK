VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Run Theme Windows8"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox f 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   28
      Text            =   "255"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox haspang 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox net 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox pngkt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox slsh 
      Height          =   285
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Ip4 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Ip3 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Ip2 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Ip1 
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Hasil:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   9255
      Begin VB.TextBox boardcast 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox range2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   23
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox range1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox subnet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox host 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Boardcast IP:"
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
         TabIndex        =   25
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00EFEFEF&
         Caption         =   "s/d"
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
         Left            =   3840
         TabIndex        =   24
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Range IP      :"
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
         TabIndex        =   21
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Net ID           :"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Subnetmask :"
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
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Jumlah Host :"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF11&
      Caption         =   "Hitung!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   600
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address :"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   8040
      Picture         =   "Form1.frx":0000
      Top             =   5400
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   8760
      Picture         =   "Form1.frx":1776
      Top             =   5400
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aplikasi Penghitung IP Address (IPv4)"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H001111EE&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FF9900&
      BorderWidth     =   4
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF9900&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF9900&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label8_Click()
If Val(slsh) <= 32 Then
pngkt = 32 - Val(slsh)
If Val(pngkt) > 8 Then
pngkt = pngkt - 8
    If Val(pngkt) > 8 Then
    pngkt = pngkt - 8
        If Val(pngkt) > 8 Then
        pngkt = pngkt - 8
        If Val(pngkt) > 8 Then
        subnet = "Error!"
        End If
        ElseIf Val(pngkt) = 8 Then
        subnet = f + "." + "0" + "." + "0" + "." + "0"
        Else
        haspang = 2 ^ Val(pngkt)
        subnet = f + "." + haspang.Text + "." + "0" + "." + "0"
        End If
    ElseIf Val(pngkt) = 8 Then
    subnet = f + "." + f + "." + "0" + "." + "0"
    Else
    haspang = 2 ^ Val(pngkt)
    subnet = f + "." + f + "." + haspang.Text + "." + "0"
    End If
ElseIf Val(pngkt) = 8 Then
subnet = f + "." + f + "." + f + "." + "0"
Else
haspang = 2 ^ Val(pngkt)
subnet = f + "." + f + "." + f + "." + haspang.Text
End If
Else
forerr.Show
End If
End Sub
