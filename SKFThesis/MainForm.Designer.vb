<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.InsertBtn = New System.Windows.Forms.Button()
        Me.ModifyBtn = New System.Windows.Forms.Button()
        Me.LayoutComboBox = New System.Windows.Forms.ComboBox()
        Me.UpdateBtn = New System.Windows.Forms.Button()
        Me.UpdateBlockDef = New System.Windows.Forms.Button()
        Me.ExportExcelCheckBox = New System.Windows.Forms.CheckBox()
        Me.ExcelFileTextBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'InsertBtn
        '
        Me.InsertBtn.Location = New System.Drawing.Point(36, 60)
        Me.InsertBtn.Margin = New System.Windows.Forms.Padding(4)
        Me.InsertBtn.Name = "InsertBtn"
        Me.InsertBtn.Size = New System.Drawing.Size(227, 129)
        Me.InsertBtn.TabIndex = 0
        Me.InsertBtn.Text = "Insert Component"
        Me.InsertBtn.UseVisualStyleBackColor = True
        '
        'ModifyBtn
        '
        Me.ModifyBtn.Location = New System.Drawing.Point(36, 60)
        Me.ModifyBtn.Margin = New System.Windows.Forms.Padding(4)
        Me.ModifyBtn.Name = "ModifyBtn"
        Me.ModifyBtn.Size = New System.Drawing.Size(227, 129)
        Me.ModifyBtn.TabIndex = 1
        Me.ModifyBtn.Text = "Modify Component"
        Me.ModifyBtn.UseVisualStyleBackColor = True
        '
        'LayoutComboBox
        '
        Me.LayoutComboBox.FormattingEnabled = True
        Me.LayoutComboBox.Items.AddRange(New Object() {"A0", "A1", "A2", "A3", "A3L", "A4", "A4SV", "A3SV"})
        Me.LayoutComboBox.Location = New System.Drawing.Point(29, 86)
        Me.LayoutComboBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.LayoutComboBox.Name = "LayoutComboBox"
        Me.LayoutComboBox.Size = New System.Drawing.Size(121, 24)
        Me.LayoutComboBox.TabIndex = 3
        '
        'UpdateBtn
        '
        Me.UpdateBtn.Location = New System.Drawing.Point(29, 146)
        Me.UpdateBtn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.UpdateBtn.Name = "UpdateBtn"
        Me.UpdateBtn.Size = New System.Drawing.Size(225, 129)
        Me.UpdateBtn.TabIndex = 4
        Me.UpdateBtn.Text = "Update BOM and Leaders"
        Me.UpdateBtn.UseVisualStyleBackColor = True
        '
        'UpdateBlockDef
        '
        Me.UpdateBlockDef.Location = New System.Drawing.Point(1072, 559)
        Me.UpdateBlockDef.Margin = New System.Windows.Forms.Padding(4)
        Me.UpdateBlockDef.Name = "UpdateBlockDef"
        Me.UpdateBlockDef.Size = New System.Drawing.Size(200, 65)
        Me.UpdateBlockDef.TabIndex = 5
        Me.UpdateBlockDef.Text = "Update Block Definitions In Drawing"
        Me.UpdateBlockDef.UseVisualStyleBackColor = True
        '
        'ExportExcelCheckBox
        '
        Me.ExportExcelCheckBox.AutoSize = True
        Me.ExportExcelCheckBox.Location = New System.Drawing.Point(317, 180)
        Me.ExportExcelCheckBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ExportExcelCheckBox.Name = "ExportExcelCheckBox"
        Me.ExportExcelCheckBox.Size = New System.Drawing.Size(150, 20)
        Me.ExportExcelCheckBox.TabIndex = 6
        Me.ExportExcelCheckBox.Text = "Export BOM to Excel"
        Me.ExportExcelCheckBox.UseVisualStyleBackColor = True
        '
        'ExcelFileTextBox
        '
        Me.ExcelFileTextBox.Location = New System.Drawing.Point(317, 238)
        Me.ExcelFileTextBox.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ExcelFileTextBox.Name = "ExcelFileTextBox"
        Me.ExcelFileTextBox.Size = New System.Drawing.Size(125, 22)
        Me.ExcelFileTextBox.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(313, 209)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Excel file name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(27, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 16)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Select layout for drawing"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ExcelFileTextBox)
        Me.GroupBox1.Controls.Add(Me.ExportExcelCheckBox)
        Me.GroupBox1.Controls.Add(Me.UpdateBtn)
        Me.GroupBox1.Controls.Add(Me.LayoutComboBox)
        Me.GroupBox1.Location = New System.Drawing.Point(713, 124)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(559, 332)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "BOM"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ModifyBtn)
        Me.GroupBox2.Location = New System.Drawing.Point(385, 210)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(303, 246)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Modify"
        '
        'GroupBox3
        '
        Me.GroupBox3.AutoSize = True
        Me.GroupBox3.Controls.Add(Me.InsertBtn)
        Me.GroupBox3.Location = New System.Drawing.Point(55, 210)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Size = New System.Drawing.Size(303, 246)
        Me.GroupBox3.TabIndex = 12
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Insert"
        '
        'ToolTip1
        '
        Me.ToolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.Rad1.My.Resources.Resources.info
        Me.PictureBox2.Location = New System.Drawing.Point(1033, 576)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(31, 27)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 14
        Me.PictureBox2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox2, "Updates blocks in old drawings to changes made in the source block")
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Rad1.My.Resources.Resources.Picture5
        Me.PictureBox1.Location = New System.Drawing.Point(141, 52)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(359, 84)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 13
        Me.PictureBox1.TabStop = False
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1344, 666)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.UpdateBlockDef)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "MainForm"
        Me.Text = "MainForm"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents InsertBtn As Windows.Forms.Button
    Friend WithEvents ModifyBtn As Windows.Forms.Button
    Friend WithEvents LayoutComboBox As Windows.Forms.ComboBox
    Friend WithEvents UpdateBtn As Windows.Forms.Button
    Friend WithEvents UpdateBlockDef As Windows.Forms.Button
    Friend WithEvents ExportExcelCheckBox As Windows.Forms.CheckBox
    Friend WithEvents ExcelFileTextBox As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
End Class
