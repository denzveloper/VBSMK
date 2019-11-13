VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   1440
   End
   Begin Project1.jcbutton jcbutton2 
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   3000
      Width           =   975
      _extentx        =   1720
      _extenty        =   873
      buttonstyle     =   8
      font            =   "Form1.frx":0000
      backcolor       =   16765357
      caption         =   "Aktifkan"
      picturepushonhover=   -1
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin Project1.jcbutton jcbutton1 
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      buttonstyle     =   8
      font            =   "Form1.frx":0028
      backcolor       =   16765357
      caption         =   "OK"
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aktifkan saat di startup"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   240
      Picture         =   "Form1.frx":0050
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6550
      TabIndex        =   4
      Top             =   50
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dendy Octavian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat Datang"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   0
      Picture         =   "Form1.frx":0C99
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean

Private Sub Check1_Click()
Select Case Check1.value
Case 0
JalanStartUp 0
Case 1
JalanStartUp 1
End Select
End Sub

Private Sub Form_Load()
    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
End Sub

Private Sub Image1_Click()

End Sub

Private Sub jcbutton1_Click()
Naik = False
Timer1.Enabled = True
End Sub

Private Sub jcbutton2_Click()
Check1.value = 1
End Sub

Private Sub Label3_Click()
Naik = False: Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Const s = 80
    Dim v As Single
    v = (GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY
    If Naik = True Then
        If Top - s <= v - Height Then
            Top = Top - (Top - (v - Height))
            Timer1.Enabled = False
        Else
            Top = Top - s
        End If
    Else
        Top = Top + s
        If Top >= v Then
        Timer1.Enabled = False
        Unload Form1
        
    End If
    End If
End Sub
