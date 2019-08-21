<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AlignBlocks
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AlignBlocks))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.StAngle = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.WrapRad = New System.Windows.Forms.TextBox()
        Me.BlkOrigin = New System.Windows.Forms.RadioButton()
        Me.Equispace = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Editcheck = New System.Windows.Forms.CheckBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Button1.Location = New System.Drawing.Point(39, 175)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(131, 30)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "ALIGN BLOCKS"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.StAngle)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 37)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(186, 78)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Angles"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(18, 20)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(85, 17)
        Me.RadioButton1.TabIndex = 10
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Start  Angles"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(18, 48)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(93, 17)
        Me.RadioButton2.TabIndex = 11
        Me.RadioButton2.Text = "Auto Center at"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(121, 47)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(44, 20)
        Me.TextBox1.TabIndex = 19
        Me.TextBox1.Text = "90"
        '
        'StAngle
        '
        Me.StAngle.Location = New System.Drawing.Point(121, 19)
        Me.StAngle.Name = "StAngle"
        Me.StAngle.Size = New System.Drawing.Size(44, 20)
        Me.StAngle.TabIndex = 15
        Me.StAngle.Text = "180"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Align Radius"
        '
        'WrapRad
        '
        Me.WrapRad.Location = New System.Drawing.Point(90, 11)
        Me.WrapRad.Name = "WrapRad"
        Me.WrapRad.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.WrapRad.Size = New System.Drawing.Size(100, 20)
        Me.WrapRad.TabIndex = 24
        Me.WrapRad.Text = "180"
        '
        'BlkOrigin
        '
        Me.BlkOrigin.AutoSize = True
        Me.BlkOrigin.Checked = True
        Me.BlkOrigin.Location = New System.Drawing.Point(93, 19)
        Me.BlkOrigin.Name = "BlkOrigin"
        Me.BlkOrigin.Size = New System.Drawing.Size(82, 17)
        Me.BlkOrigin.TabIndex = 27
        Me.BlkOrigin.TabStop = True
        Me.BlkOrigin.Text = "Block Origin"
        Me.BlkOrigin.UseVisualStyleBackColor = True
        '
        'Equispace
        '
        Me.Equispace.AutoSize = True
        Me.Equispace.Enabled = False
        Me.Equispace.Location = New System.Drawing.Point(6, 19)
        Me.Equispace.Name = "Equispace"
        Me.Equispace.Size = New System.Drawing.Size(81, 17)
        Me.Equispace.TabIndex = 26
        Me.Equispace.TabStop = True
        Me.Equispace.Text = "Equispaced"
        Me.Equispace.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Equispace)
        Me.GroupBox2.Controls.Add(Me.BlkOrigin)
        Me.GroupBox2.Location = New System.Drawing.Point(15, 121)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(186, 48)
        Me.GroupBox2.TabIndex = 28
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Place block at Angle"
        '
        'Editcheck
        '
        Me.Editcheck.AutoSize = True
        Me.Editcheck.Location = New System.Drawing.Point(287, 36)
        Me.Editcheck.Name = "Editcheck"
        Me.Editcheck.Size = New System.Drawing.Size(81, 17)
        Me.Editcheck.TabIndex = 29
        Me.Editcheck.Text = "CheckBox1"
        Me.Editcheck.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(287, 68)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 30
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.WrapRad)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(12, 11)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(214, 213)
        Me.Panel1.TabIndex = 31
        '
        'AlignBlocks
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.ClientSize = New System.Drawing.Size(238, 234)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Editcheck)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "AlignBlocks"
        Me.Text = "AlignBlocks "
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents StAngle As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents WrapRad As System.Windows.Forms.TextBox
    Friend WithEvents BlkOrigin As System.Windows.Forms.RadioButton
    Friend WithEvents Equispace As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Editcheck As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
