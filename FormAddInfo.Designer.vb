<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormAddInfo
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
        Me.txtPanjang1 = New System.Windows.Forms.TextBox()
        Me.txtReaktansi1 = New System.Windows.Forms.TextBox()
        Me.txtCCC1 = New System.Windows.Forms.TextBox()
        Me.txtResistansi1 = New System.Windows.Forms.TextBox()
        Me.AddInfo1 = New System.Windows.Forms.GroupBox()
        Me.Save = New System.Windows.Forms.Button()
        Me.Back = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.AddInfo1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtPanjang1
        '
        Me.txtPanjang1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPanjang1.Location = New System.Drawing.Point(153, 101)
        Me.txtPanjang1.Name = "txtPanjang1"
        Me.txtPanjang1.Size = New System.Drawing.Size(118, 21)
        Me.txtPanjang1.TabIndex = 98
        Me.txtPanjang1.Text = "Length (Km)"
        Me.txtPanjang1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtReaktansi1
        '
        Me.txtReaktansi1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReaktansi1.Location = New System.Drawing.Point(151, 74)
        Me.txtReaktansi1.Name = "txtReaktansi1"
        Me.txtReaktansi1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtReaktansi1.Size = New System.Drawing.Size(120, 21)
        Me.txtReaktansi1.TabIndex = 97
        Me.txtReaktansi1.Text = "Reactance (Ohm)"
        Me.txtReaktansi1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCCC1
        '
        Me.txtCCC1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCCC1.Location = New System.Drawing.Point(153, 20)
        Me.txtCCC1.Name = "txtCCC1"
        Me.txtCCC1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtCCC1.Size = New System.Drawing.Size(118, 21)
        Me.txtCCC1.TabIndex = 96
        Me.txtCCC1.Text = "CCC (A)"
        Me.txtCCC1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtResistansi1
        '
        Me.txtResistansi1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtResistansi1.Location = New System.Drawing.Point(152, 47)
        Me.txtResistansi1.Name = "txtResistansi1"
        Me.txtResistansi1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtResistansi1.Size = New System.Drawing.Size(119, 21)
        Me.txtResistansi1.TabIndex = 96
        Me.txtResistansi1.Text = "Resistance (Ohm)"
        Me.txtResistansi1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'AddInfo1
        '
        Me.AddInfo1.Controls.Add(Me.Label4)
        Me.AddInfo1.Controls.Add(Me.Label3)
        Me.AddInfo1.Controls.Add(Me.Label2)
        Me.AddInfo1.Controls.Add(Me.Label1)
        Me.AddInfo1.Controls.Add(Me.txtPanjang1)
        Me.AddInfo1.Controls.Add(Me.txtReaktansi1)
        Me.AddInfo1.Controls.Add(Me.txtCCC1)
        Me.AddInfo1.Controls.Add(Me.txtResistansi1)
        Me.AddInfo1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddInfo1.Location = New System.Drawing.Point(12, 12)
        Me.AddInfo1.Name = "AddInfo1"
        Me.AddInfo1.Size = New System.Drawing.Size(288, 132)
        Me.AddInfo1.TabIndex = 94
        Me.AddInfo1.TabStop = False
        '
        'Save
        '
        Me.Save.Location = New System.Drawing.Point(58, 181)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(75, 23)
        Me.Save.TabIndex = 101
        Me.Save.Text = "Save"
        Me.Save.UseVisualStyleBackColor = True
        '
        'Back
        '
        Me.Back.Location = New System.Drawing.Point(193, 181)
        Me.Back.Name = "Back"
        Me.Back.Size = New System.Drawing.Size(75, 23)
        Me.Back.TabIndex = 102
        Me.Back.Text = "Back"
        Me.Back.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 15)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Current Carrying Capacity "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 50)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(143, 15)
        Me.Label2.TabIndex = 100
        Me.Label2.Text = "Resistance of Conductor "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(141, 15)
        Me.Label3.TabIndex = 101
        Me.Label3.Text = "Reactance of Conductor "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(144, 15)
        Me.Label4.TabIndex = 103
        Me.Label4.Text = "Line Length of Conductor"
        '
        'FormAddInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 239)
        Me.Controls.Add(Me.Back)
        Me.Controls.Add(Me.Save)
        Me.Controls.Add(Me.AddInfo1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "FormAddInfo"
        Me.Text = "Form Additional Conductor Type"
        Me.AddInfo1.ResumeLayout(False)
        Me.AddInfo1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtPanjang1 As System.Windows.Forms.TextBox
    Friend WithEvents txtReaktansi1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCCC1 As System.Windows.Forms.TextBox
    Friend WithEvents txtResistansi1 As System.Windows.Forms.TextBox
    Friend WithEvents AddInfo1 As System.Windows.Forms.GroupBox
    Friend WithEvents Save As System.Windows.Forms.Button
    Friend WithEvents Back As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
