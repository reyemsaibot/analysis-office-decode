<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.btn_open = New System.Windows.Forms.Button()
        Me.txt_path = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btn_decompress = New System.Windows.Forms.Button()
        Me.btn_compress = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn_open
        '
        Me.btn_open.Location = New System.Drawing.Point(282, 12)
        Me.btn_open.Name = "btn_open"
        Me.btn_open.Size = New System.Drawing.Size(75, 23)
        Me.btn_open.TabIndex = 0
        Me.btn_open.Text = "Open File"
        Me.btn_open.UseVisualStyleBackColor = True
        '
        'txt_path
        '
        Me.txt_path.Location = New System.Drawing.Point(12, 12)
        Me.txt_path.Name = "txt_path"
        Me.txt_path.Size = New System.Drawing.Size(264, 20)
        Me.txt_path.TabIndex = 1
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog"
        '
        'btn_decompress
        '
        Me.btn_decompress.Location = New System.Drawing.Point(12, 38)
        Me.btn_decompress.Name = "btn_decompress"
        Me.btn_decompress.Size = New System.Drawing.Size(123, 23)
        Me.btn_decompress.TabIndex = 3
        Me.btn_decompress.Text = "Decompress Excelfile"
        Me.btn_decompress.UseVisualStyleBackColor = True
        '
        'btn_compress
        '
        Me.btn_compress.Location = New System.Drawing.Point(259, 38)
        Me.btn_compress.Name = "btn_compress"
        Me.btn_compress.Size = New System.Drawing.Size(98, 23)
        Me.btn_compress.TabIndex = 4
        Me.btn_compress.Text = "Compress XML"
        Me.btn_compress.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(141, 38)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Open Tempfolder"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 123)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btn_compress)
        Me.Controls.Add(Me.btn_decompress)
        Me.Controls.Add(Me.txt_path)
        Me.Controls.Add(Me.btn_open)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(465, 162)
        Me.MinimumSize = New System.Drawing.Size(465, 162)
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manipulate Excel Property"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_open As System.Windows.Forms.Button
    Friend WithEvents txt_path As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btn_decompress As System.Windows.Forms.Button
    Friend WithEvents btn_compress As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button

End Class
