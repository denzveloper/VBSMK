Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	Private Declare Function GetWindowLong Lib "user32.dll"  Alias "GetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
	Private Declare Function SetWindowLong Lib "user32.dll"  Alias "SetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer
	Dim Naik As Boolean
	
	'UPGRADE_WARNING: Event Check1.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
		Select Case Check1.CheckState
			Case 0
				JalanStartUp(0)
			Case 1
				JalanStartUp(1)
		End Select
	End Sub
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Top = VB6.TwipsToPixelsY(((GetSystemMetrics(17) + GetSystemMetrics(4)) * VB6.TwipsPerPixelY))
		Left = VB6.TwipsToPixelsX((GetSystemMetrics(16) * VB6.TwipsPerPixelX) - VB6.PixelsToTwipsX(Width))
		Naik = True
	End Sub
	
	Private Sub Image1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Image1.Click
		
	End Sub
	
	Private Sub jcbutton1_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles jcbutton1.Click
		Naik = False
		Timer1.Enabled = True
	End Sub
	
	Private Sub jcbutton2_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles jcbutton2.Click
		Check1.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub Label3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Label3.Click
		Naik = False : Timer1.Enabled = True
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Const s As Short = 80
		Dim v As Single
		v = (GetSystemMetrics(17) + GetSystemMetrics(4)) * VB6.TwipsPerPixelY
		If Naik = True Then
			If VB6.PixelsToTwipsY(Top) - s <= v - VB6.PixelsToTwipsY(Height) Then
				Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Top) - (VB6.PixelsToTwipsY(Top) - (v - VB6.PixelsToTwipsY(Height))))
				Timer1.Enabled = False
			Else
				Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Top) - s)
			End If
		Else
			Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Top) + s)
			If VB6.PixelsToTwipsY(Top) >= v Then
				Timer1.Enabled = False
				Me.Close()
				
			End If
		End If
	End Sub
End Class