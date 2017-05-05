<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmRegisterComponents
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRegisterComponents))
        Me.btnRegister = New System.Windows.Forms.Button
        Me.lblcomponents = New System.Windows.Forms.Label
        Me.lblos = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.lblmachine = New System.Windows.Forms.Label
        Me.lblplat = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnRegister
        '
        Me.btnRegister.Image = CType(resources.GetObject("btnRegister.Image"), System.Drawing.Image)
        Me.btnRegister.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRegister.Location = New System.Drawing.Point(0, 3)
        Me.btnRegister.Name = "btnRegister"
        Me.btnRegister.Size = New System.Drawing.Size(152, 23)
        Me.btnRegister.TabIndex = 0
        Me.btnRegister.Text = "Register Components"
        Me.btnRegister.UseVisualStyleBackColor = True
        '
        'lblcomponents
        '
        Me.lblcomponents.AutoSize = True
        Me.lblcomponents.Location = New System.Drawing.Point(5, 28)
        Me.lblcomponents.Name = "lblcomponents"
        Me.lblcomponents.Size = New System.Drawing.Size(222, 13)
        Me.lblcomponents.TabIndex = 1
        Me.lblcomponents.Text = "Components                                                    "
        Me.lblcomponents.Visible = False
        '
        'lblos
        '
        Me.lblos.AutoSize = True
        Me.lblos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblos.Location = New System.Drawing.Point(6, 36)
        Me.lblos.Name = "lblos"
        Me.lblos.Size = New System.Drawing.Size(39, 13)
        Me.lblos.TabIndex = 2
        Me.lblos.Text = "Label2"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(0, 15)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(194, 69)
        Me.ListBox1.TabIndex = 3
        '
        'lblmachine
        '
        Me.lblmachine.AutoSize = True
        Me.lblmachine.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblmachine.Location = New System.Drawing.Point(6, 16)
        Me.lblmachine.Name = "lblmachine"
        Me.lblmachine.Size = New System.Drawing.Size(39, 13)
        Me.lblmachine.TabIndex = 4
        Me.lblmachine.Text = "Label1"
        '
        'lblplat
        '
        Me.lblplat.AutoSize = True
        Me.lblplat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblplat.Location = New System.Drawing.Point(6, 56)
        Me.lblplat.Name = "lblplat"
        Me.lblplat.Size = New System.Drawing.Size(39, 13)
        Me.lblplat.TabIndex = 5
        Me.lblplat.Text = "Label1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(6, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Label2"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ListBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 145)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 100)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Files Not registered ....Do manualy"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblmachine)
        Me.GroupBox2.Controls.Add(Me.lblos)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.lblplat)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 44)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(372, 100)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "System Info"
        '
        'btnExit
        '
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(310, 217)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(74, 27)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Exit"
        '
        'FrmRegisterComponents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(410, 256)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblcomponents)
        Me.Controls.Add(Me.btnRegister)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmRegisterComponents"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Register components (vb6)"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRegister As System.Windows.Forms.Button
    Friend WithEvents lblcomponents As System.Windows.Forms.Label
    Friend WithEvents lblos As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents lblmachine As System.Windows.Forms.Label
    Friend WithEvents lblplat As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button

End Class
