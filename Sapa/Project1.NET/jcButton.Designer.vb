<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated(), ToolboxBitmap(GetType(jcbutton), "jcbutton.ToolboxBitmap")> Partial Class jcbutton
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		UserControl_Initialize()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			UserControl_Terminate()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(jcbutton))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ClientSize = New System.Drawing.Size(89, 25)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Name = "jcbutton"
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
#Region "Upgrade Support"
	<System.Runtime.InteropServices.ProgId("KeyPressEventArgs_NET.KeyPressEventArgs")> Public NotInheritable Class KeyPressEventArgs
		Inherits System.EventArgs
		Public KeyAcsii As Short
		Public Sub New(ByRef KeyAcsii As Short)
			MyBase.New()
			Me.KeyAcsii = KeyAcsii
		End Sub
	End Class
	<System.Runtime.InteropServices.ProgId("KeyUpEventArgs_NET.KeyUpEventArgs")> Public NotInheritable Class KeyUpEventArgs
		Inherits System.EventArgs
		Public KeyCode As Short
		Public Shift As Short
		Public Sub New(ByRef KeyCode As Short, ByRef Shift As Short)
			MyBase.New()
			Me.KeyCode = KeyCode
			Me.Shift = Shift
		End Sub
	End Class
	<System.Runtime.InteropServices.ProgId("KeyDownEventArgs_NET.KeyDownEventArgs")> Public NotInheritable Class KeyDownEventArgs
		Inherits System.EventArgs
		Public KeyCode As Short
		Public Shift As Short
		Public Sub New(ByRef KeyCode As Short, ByRef Shift As Short)
			MyBase.New()
			Me.KeyCode = KeyCode
			Me.Shift = Shift
		End Sub
	End Class
	<System.Runtime.InteropServices.ProgId("MouseDownEventArgs_NET.MouseDownEventArgs")> Public NotInheritable Class MouseDownEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public X As Single
		Public Y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef Y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.X = X
			Me.Y = Y
		End Sub
	End Class
	<System.Runtime.InteropServices.ProgId("MouseUpEventArgs_NET.MouseUpEventArgs")> Public NotInheritable Class MouseUpEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public X As Single
		Public Y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef Y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.X = X
			Me.Y = Y
		End Sub
	End Class
	<System.Runtime.InteropServices.ProgId("MouseMoveEventArgs_NET.MouseMoveEventArgs")> Public NotInheritable Class MouseMoveEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public X As Single
		Public Y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef Y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.X = X
			Me.Y = Y
		End Sub
	End Class
#End Region 
End Class