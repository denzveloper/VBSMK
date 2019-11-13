<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents jcbutton2 As jcbutton
	Public WithEvents jcbutton1 As jcbutton
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents Image2 As System.Windows.Forms.PictureBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.jcbutton2 = New jcbutton
		Me.jcbutton1 = New jcbutton
		Me.Check1 = New System.Windows.Forms.CheckBox
		Me.Image2 = New System.Windows.Forms.PictureBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.White
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(457, 247)
		Me.Location = New System.Drawing.Point(0, 0)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form1"
		Me.Timer1.Interval = 1
		Me.Timer1.Enabled = True
		Me.jcbutton2.Size = New System.Drawing.Size(65, 33)
		Me.jcbutton2.Location = New System.Drawing.Point(376, 200)
		Me.jcbutton2.TabIndex = 5
		Me.jcbutton2.Name = "jcbutton2"
		Me.jcbutton1.Size = New System.Drawing.Size(145, 33)
		Me.jcbutton1.Location = New System.Drawing.Point(232, 200)
		Me.jcbutton1.TabIndex = 3
		Me.jcbutton1.Name = "jcbutton1"
		Me.Check1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.Check1.BackColor = System.Drawing.Color.White
		Me.Check1.Text = "Aktifkan saat di startup"
		Me.Check1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Check1.Size = New System.Drawing.Size(169, 17)
		Me.Check1.Location = New System.Drawing.Point(0, 224)
		Me.Check1.TabIndex = 0
		Me.Check1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Check1.CausesValidation = True
		Me.Check1.Enabled = True
		Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check1.TabStop = True
		Me.Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.Check1.Visible = True
		Me.Check1.Name = "Check1"
		Me.Image2.Size = New System.Drawing.Size(72, 72)
		Me.Image2.Location = New System.Drawing.Point(16, 96)
		Me.Image2.Image = CType(resources.GetObject("Image2.Image"), System.Drawing.Image)
		Me.Image2.Enabled = True
		Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image2.Visible = True
		Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image2.Name = "Image2"
		Me.Label3.Text = "X"
		Me.Label3.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(21, 22)
		Me.Label3.Location = New System.Drawing.Point(437, 4)
		Me.Label3.TabIndex = 4
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.Color.Transparent
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = True
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.Text = "Dendy Octavian"
		Me.Label2.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.Color.White
		Me.Label2.Size = New System.Drawing.Size(217, 33)
		Me.Label2.Location = New System.Drawing.Point(120, 80)
		Me.Label2.TabIndex = 2
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Selamat Datang"
		Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 26.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(457, 42)
		Me.Label1.Location = New System.Drawing.Point(0, 32)
		Me.Label1.TabIndex = 1
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = True
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Image1.Size = New System.Drawing.Size(457, 129)
		Me.Image1.Location = New System.Drawing.Point(0, 0)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.Controls.Add(jcbutton2)
		Me.Controls.Add(jcbutton1)
		Me.Controls.Add(Check1)
		Me.Controls.Add(Image2)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Image1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class