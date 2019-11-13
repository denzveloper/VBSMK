VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kalkulator :: GNU For Windows"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox sv 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command21 
      Caption         =   "C&E"
      Height          =   615
      Left            =   9360
      TabIndex        =   23
      ToolTipText     =   "Clear Digit Button (alt+e)"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&C"
      Height          =   615
      Left            =   9360
      TabIndex        =   14
      ToolTipText     =   "Clear All (alt+c)"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox res 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   240
      Width           =   9015
   End
   Begin VB.CommandButton Command19 
      Caption         =   "^"
      Height          =   855
      Left            =   5040
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox h 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5040
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox f 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5280
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton Command18 
      Caption         =   "00"
      Height          =   735
      Left            =   3240
      TabIndex        =   17
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "0"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton Command15 
      Caption         =   "+"
      Height          =   855
      Left            =   5040
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "/"
      Height          =   975
      Left            =   5040
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "*"
      Height          =   975
      Left            =   6480
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "-"
      Height          =   855
      Left            =   6480
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "3"
      Height          =   855
      Left            =   3240
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "6"
      Height          =   855
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "9"
      Height          =   855
      Left            =   3240
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "="
      Height          =   1935
      Left            =   7920
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "2"
      Height          =   855
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   855
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "8"
      Height          =   855
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Control Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0033CC33&
      Height          =   3495
      Left            =   4800
      TabIndex        =   21
      Top             =   1440
      Width           =   4695
      Begin VB.Shape iyd 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   615
         Left            =   3840
         Top             =   480
         Width           =   615
      End
      Begin VB.Image Image5 
         Height          =   630
         Left            =   3120
         Picture         =   "Form1.frx":0ECA
         Top             =   480
         Width           =   645
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   1680
         Picture         =   "Form1.frx":24B4
         ToolTipText     =   "Plus/Min"
         Top             =   480
         Width           =   645
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   2400
         Picture         =   "Form1.frx":3A9E
         ToolTipText     =   "Phi Untuk Lingkaran(3,14...)"
         Top             =   480
         Width           =   645
      End
   End
   Begin VB.TextBox ans 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      HideSelection   =   0   'False
      Left            =   240
      MaxLength       =   18
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   9015
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   9600
      MousePointer    =   12  'No Drop
      Picture         =   "Form1.frx":5088
      ToolTipText     =   "Exit Applicaton"
      Top             =   4080
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   9600
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Form1.frx":67FE
      ToolTipText     =   "About Kalkulator GNU"
      Top             =   1560
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command11_Click()
iyd.BackColor = &H808080
If ans = "" Then
    ans = "0"
    f = ans
    sv = Val(sv) + 1
    h = "-"
    res = f + "-"
    ans = ""
    If sv > 1 Then
        res = ""
        sv = 0
    End If
Else
f = ans
sv = Val(sv) + 1
h = "-"
res = f + "-"
ans = ""
If sv > 1 Then
    res = ""
    sv = 0
End If
End If
End Sub

Private Sub Command12_Click()
iyd.BackColor = &H808080
If ans = "" Then
    ans = "0"
    f = ans
    sv = Val(sv) + 1
    h = "*"
    res = f + "*"
    ans = ""
    If sv > 1 Then
        res = ""
        sv = 0
    End If
Else
f = ans
sv = Val(sv) + 1
h = "*"
res = f + "*"
ans = ""
If sv > 1 Then
    res = ""
    sv = 0
End If
End If
End Sub

Private Sub Command13_Click()
iyd.BackColor = &H808080
If ans = "" Then
    ans = "0"
    f = ans
    sv = Val(sv) + 1
    h = "/"
    res = f + "/"
    ans = ""
    If sv > 1 Then
        res = ""
        sv = 0
    End If
Else
f = ans
sv = Val(sv) + 1
h = "/"
res = f + "/"
ans = ""
If sv > 1 Then
    res = ""
    sv = 0
End If
End If
End Sub

Private Sub Command14_Click()
ans = ""
res = ""
iyd.BackColor = &H808080
End Sub

Private Sub Command15_Click()
iyd.BackColor = &H808080
If ans = "" Then
    ans = "0"
    f = ans
    sv = Val(sv) + 1
    h = "p"
    res = f + "+"
    ans = ""
    If sv > 1 Then
        res = ""
        sv = 0
        iyd.BackColor = &H808080
    End If
Else
f = ans
sv = Val(sv) + 1
h = "p"
res = f + "+"
ans = ""
iyd.BackColor = &H808080
If sv > 1 Then
    res = ""
    sv = 0
    iyd.BackColor = &H808080
End If
End If
End Sub
Private Sub Command10_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "3"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "3"
Else
ans = Form1.ans.Text + "3"
End If
End Sub
Private Sub Command16_Click()
iyd.BackColor = vbBlack
Unload Me
End Sub

Private Sub Command1_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "7"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "7"
Else
ans = Form1.ans.Text + "7"
End If
End Sub

Private Sub Command17_Click()
iyd.BackColor = &H808080
If ans = "" Then
ans = ""
ElseIf ans = "0" Then
ans = ""
ElseIf ans = "00" Then
ans = ""
Else
ans = Form1.ans.Text + "0"
End If
End Sub

Private Sub Command18_Click()
iyd.BackColor = &H808080
If ans = "" Then
ans = ""
ElseIf ans = "0" Then
ans = ""
ElseIf ans = "00" Then
ans = ""
Else
ans = Form1.ans.Text + "00"
End If
End Sub

Private Sub Command19_Click()
iyd.BackColor = &H808080
If ans = "" Then
    ans = "0"
    f = ans
    h = "k"
    sv = Val(sv) + 1
    res = Form1.f.Text + "^"
    ans = ""
    If sv > 1 Then
        res = ""
        sv = 0
    End If
Else
f = ans
sv = Val(sv) + 1
h = "k"
res = f + "^"
ans = ""
If sv > 1 Then
    res = ""
    sv = 0
End If
End If
End Sub

Private Sub Command2_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "4"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "4"
Else
ans = Form1.ans.Text + "4"
End If
End Sub

Private Sub Command21_Click()
iyd.BackColor = &H808080
If ans + f = "ERROR!" Then
ans = ""
    If sv >= "2" Then
        res = ""
        sv = 0
    End If
ElseIf ans = "" Then
ans = ""
ElseIf ans = 0 Then
ans = 0
Else
ans = ""
    If sv >= "2" Then
        res = ""
        sv = 0
    End If
End If
End Sub

Private Sub Command3_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "1"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "1"
Else
ans = Form1.ans.Text + "1"
End If
End Sub

Private Sub Command4_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "8"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "8"
Else
ans = Form1.ans.Text + "8"
End If
End Sub

Private Sub Command5_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "5"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "5"
Else
ans = Form1.ans.Text + "5"
End If
End Sub

Private Sub Command6_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "2"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "2"
Else
ans = Form1.ans.Text + "2"
End If
End Sub

Private Sub Command8_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "9"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "9"
Else
ans = Form1.ans.Text + "9"
End If
End Sub

Private Sub Command9_Click()
iyd.BackColor = &H808080
If ans = "0" Then
ans = ""
ans = Form1.ans.Text + "6"
ElseIf ans = "00" Then
ans = ""
ans = Form1.ans.Text + "6"
Else
ans = Form1.ans.Text + "6"
End If
End Sub

Private Sub Command7_Click()
iyd.BackColor = vbGreen
If ans + f = "" Then
    ans = ""
ElseIf ans = "" Then
 ans = 0
    If h = "p" Then
        sv = Val(sv) + 1
        res = f + "+" + ans + "="
        ans = Val(f) + Val(ans)
        f = ""
        If sv > 2 Then
            res = ""
            f = ""
            ans = ""
            sv = ""
        End If
    End If
    If h = "-" Then
    sv = Val(sv) + 1
    res = f + "-" + ans + "="
    ans = Val(f) - Val(ans)
    f = ""
        If sv > 2 Then
        res = ""
        f = ""
        ans = ""
        sv = ""
        End If
    End If
    If h = "*" Then
    sv = Val(sv) + 1
    res = f + "*" + ans + "="
    ans = Val(f) * Val(ans)
    f = ""
        If sv > 2 Then
        res = ""
        f = ""
        ans = ""
        sv = ""
        End If
    End If
    If h = "/" Then
    iyd.BackColor = vbRed
        If ans = "ERROR!" Then
             sv = Val(sv) + 1
             f = ""
             res = ""
             ans = ""
                If sv > 2 Then
                     res = ""
                     f = ""
                     ans = ""
                     sv = ""
                 End If
        ElseIf f + ans = "" Then
        iyd.BackColor = vbRed
        sv = Val(sv) + 1
             f = ""
             res = "All Value is Empty!"
             ans = "ERROR!"
                If sv > 2 Then
                 res = ""
                 f = ""
                 ans = ""
                 sv = ""
                End If
        ElseIf f <> 0 Then
            If ans <> 0 Then
            sv = Val(sv) + 1
            res = f + "/" + ans + "="
            ans = Val(f) / Val(ans)
            f = ""
              If sv > 2 Then
                res = ""
                f = ""
                ans = ""
                sv = ""
              End If
            Else
             sv = Val(sv) + 1
             f = ""
                res = "Cannot devide by zero!"
                ans = "ERROR!"
                iyd.BackColor = vbRed
                If sv > 2 Then
                 res = ""
                 f = ""
                 ans = ""
                 sv = ""
                End If
            End If
        Else
            res = "Cannot devide by zero!"
            ans = "ERROR!"
            iyd.BackColor = vbRed
            f = ""
            If sv > 2 Then
                   res = ""
                   f = ""
                   ans = ""
                   sv = ""
            End If
        End If
    End If
    If h = "k" Then
    sv = Val(sv) + 1
    res = f + "^" + ans + "="
    ans = Val(f) ^ Val(ans)
    f = ""
        If sv > 2 Then
        res = ""
        f = ""
        ans = ""
        sv = ""
        End If
    End If
Else
    If h = "p" Then
        sv = Val(sv) + 1
        res = f + "+" + ans + "="
        ans = Val(f) + Val(ans)
        f = ""
        If sv > 2 Then
            res = ""
            f = ""
            ans = ""
            sv = ""
        End If
    End If
    If h = "-" Then
    sv = Val(sv) + 1
    res = f + "-" + ans + "="
    ans = Val(f) - Val(ans)
    f = ""
        If sv > 2 Then
        res = ""
        f = ""
        ans = ""
        sv = ""
        End If
    End If
    If h = "*" Then
    sv = Val(sv) + 1
    res = f + "*" + ans + "="
    ans = Val(f) * Val(ans)
    f = ""
        If sv > 2 Then
        res = ""
        f = ""
        ans = ""
        sv = ""
        End If
    End If
    If h = "/" Then
        If ans = "ERROR!" Then
             sv = Val(sv) + 1
             f = ""
             res = ""
             ans = ""
                If sv > 2 Then
                     res = ""
                     f = ""
                     ans = ""
                     sv = ""
                 End If
        ElseIf f + ans = "" Then
        sv = Val(sv) + 1
             f = ""
             res = "All Value is Empty!"
             ans = "ERROR!"
             iyd.BackColor = vbRed
                If sv > 2 Then
                 res = ""
                 f = ""
                 ans = ""
                 sv = ""
                End If
        ElseIf f = "" Then
             sv = Val(sv) + 1
             f = ""
                res = "Cannot devide by zero!"
                ans = "ERROR!"
                iyd.BackColor = vbRed
                If sv > 2 Then
                 res = ""
                 f = ""
                 ans = ""
                 sv = ""
                End If
        ElseIf f <> 0 Then
            If ans <> 0 Then
            sv = Val(sv) + 1
            res = f + "/" + ans + "="
            ans = Val(f) / Val(ans)
            f = ""
              If sv > 2 Then
                res = ""
                f = ""
                ans = ""
                sv = ""
              End If
            Else
             sv = Val(sv) + 1
             f = ""
                res = "Cannot devide by zero!"
                ans = "ERROR!"
                iyd.BackColor = vbRed
                If sv > 2 Then
                 res = ""
                 f = ""
                 ans = ""
                 sv = ""
                End If
            End If
        Else
            res = f + "/" + ans + "="
            ans = 0
            f = ""
            If sv > 2 Then
                   res = ""
                   f = ""
                   ans = ""
                   sv = ""
            End If
        End If
    End If
    If h = "k" Then
    sv = Val(sv) + 1
    res = f + "^" + ans + "="
    ans = Val(f) ^ Val(ans)
    f = ""
        If sv > 2 Then
        res = ""
        f = ""
        ans = ""
        sv = ""
        End If
    End If
End If
End Sub

Private Sub Form_Load()
sv = 0
iyd.BackColor = &H808080
End Sub

Private Sub Image1_Click()
iyd.BackColor = vbWhite
Form2.Show
End Sub

Private Sub Image2_Click()
iyd.BackColor = vbYellow
frmAbout.Show
End Sub

Private Sub Image3_Click()
iyd.BackColor = vbGreen
ans = 22 / 7
End Sub

Private Sub Image4_Click()
iyd.BackColor = &HFFFF00
If Val(ans) = "0" Then
ans = 0
Else
ans = ans * (-1)
End If
End Sub

Private Sub Image5_Click()
iyd.BackColor = vbBlue
ans = 3.17
End Sub
