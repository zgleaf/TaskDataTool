<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskDataTool
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
        Me.Btn_excel = New System.Windows.Forms.Button
        Me.TB_excel = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TB_db = New System.Windows.Forms.TextBox
        Me.Btn_db = New System.Windows.Forms.Button
        Me.CB_append = New System.Windows.Forms.CheckBox
        Me.Btn_import = New System.Windows.Forms.Button
        Me.Btn_export = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.TB_Log = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Btn_excel
        '
        Me.Btn_excel.Location = New System.Drawing.Point(564, 25)
        Me.Btn_excel.Name = "Btn_excel"
        Me.Btn_excel.Size = New System.Drawing.Size(75, 23)
        Me.Btn_excel.TabIndex = 0
        Me.Btn_excel.Text = "open"
        Me.Btn_excel.UseVisualStyleBackColor = True
        '
        'TB_excel
        '
        Me.TB_excel.Location = New System.Drawing.Point(65, 27)
        Me.TB_excel.Name = "TB_excel"
        Me.TB_excel.Size = New System.Drawing.Size(493, 21)
        Me.TB_excel.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Excel"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(17, 12)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "DB"
        '
        'TB_db
        '
        Me.TB_db.Location = New System.Drawing.Point(65, 65)
        Me.TB_db.Name = "TB_db"
        Me.TB_db.Size = New System.Drawing.Size(493, 21)
        Me.TB_db.TabIndex = 4
        '
        'Btn_db
        '
        Me.Btn_db.Location = New System.Drawing.Point(564, 63)
        Me.Btn_db.Name = "Btn_db"
        Me.Btn_db.Size = New System.Drawing.Size(75, 23)
        Me.Btn_db.TabIndex = 3
        Me.Btn_db.Text = "open"
        Me.Btn_db.UseVisualStyleBackColor = True
        '
        'CB_append
        '
        Me.CB_append.AutoSize = True
        Me.CB_append.Location = New System.Drawing.Point(65, 103)
        Me.CB_append.Name = "CB_append"
        Me.CB_append.Size = New System.Drawing.Size(60, 16)
        Me.CB_append.TabIndex = 6
        Me.CB_append.Text = "Append"
        Me.CB_append.UseVisualStyleBackColor = True
        '
        'Btn_import
        '
        Me.Btn_import.Location = New System.Drawing.Point(186, 148)
        Me.Btn_import.Name = "Btn_import"
        Me.Btn_import.Size = New System.Drawing.Size(75, 23)
        Me.Btn_import.TabIndex = 7
        Me.Btn_import.Text = "import"
        Me.Btn_import.UseVisualStyleBackColor = True
        '
        'Btn_export
        '
        Me.Btn_export.Location = New System.Drawing.Point(367, 148)
        Me.Btn_export.Name = "Btn_export"
        Me.Btn_export.Size = New System.Drawing.Size(75, 23)
        Me.Btn_export.TabIndex = 8
        Me.Btn_export.Text = "export"
        Me.Btn_export.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'TB_Log
        '
        Me.TB_Log.BackColor = System.Drawing.SystemColors.Menu
        Me.TB_Log.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TB_Log.HideSelection = False
        Me.TB_Log.Location = New System.Drawing.Point(26, 177)
        Me.TB_Log.Multiline = True
        Me.TB_Log.Name = "TB_Log"
        Me.TB_Log.ReadOnly = True
        Me.TB_Log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TB_Log.Size = New System.Drawing.Size(613, 121)
        Me.TB_Log.TabIndex = 10
        Me.TB_Log.TabStop = False
        '
        'TaskDataTool
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(665, 310)
        Me.Controls.Add(Me.TB_Log)
        Me.Controls.Add(Me.Btn_export)
        Me.Controls.Add(Me.Btn_import)
        Me.Controls.Add(Me.CB_append)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TB_db)
        Me.Controls.Add(Me.Btn_db)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TB_excel)
        Me.Controls.Add(Me.Btn_excel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TaskDataTool"
        Me.Text = "TaskDataTool"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Btn_excel As System.Windows.Forms.Button
    Friend WithEvents TB_excel As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TB_db As System.Windows.Forms.TextBox
    Friend WithEvents Btn_db As System.Windows.Forms.Button
    Friend WithEvents CB_append As System.Windows.Forms.CheckBox
    Friend WithEvents Btn_import As System.Windows.Forms.Button
    Friend WithEvents Btn_export As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TB_Log As System.Windows.Forms.TextBox

End Class
