<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Email
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Email))
        Me.Emails = New System.Windows.Forms.DataGridView()
        Me.Extractor = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExeBTN = New System.Windows.Forms.Button()
        Me.ExportBTN = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.E_mails = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.CancBTN = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.Emails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Emails
        '
        Me.Emails.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Extractor})
        Me.Emails.Location = New System.Drawing.Point(12, 41)
        Me.Emails.Name = "Emails"
        Me.Emails.RowHeadersVisible = False
        Me.Emails.Size = New System.Drawing.Size(640, 356)
        Me.Emails.TabIndex = 6
        '
        'Extractor
        '
        Me.Extractor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Extractor.HeaderText = "E-mail Extractor"
        Me.Extractor.Name = "Extractor"
        '
        'ExeBTN
        '
        Me.ExeBTN.Location = New System.Drawing.Point(12, 12)
        Me.ExeBTN.Name = "ExeBTN"
        Me.ExeBTN.Size = New System.Drawing.Size(75, 23)
        Me.ExeBTN.TabIndex = 4
        Me.ExeBTN.Text = "Executar"
        Me.ExeBTN.UseVisualStyleBackColor = True
        '
        'ExportBTN
        '
        Me.ExportBTN.Enabled = False
        Me.ExportBTN.Location = New System.Drawing.Point(554, 12)
        Me.ExportBTN.Name = "ExportBTN"
        Me.ExportBTN.Size = New System.Drawing.Size(98, 23)
        Me.ExportBTN.TabIndex = 5
        Me.ExportBTN.Text = "Exportar CSV"
        Me.ExportBTN.UseVisualStyleBackColor = True
        '
        'E_mails
        '
        Me.E_mails.Name = "E_mails"
        '
        'BackgroundWorker
        '
        Me.BackgroundWorker.WorkerReportsProgress = True
        Me.BackgroundWorker.WorkerSupportsCancellation = True
        '
        'CancBTN
        '
        Me.CancBTN.Location = New System.Drawing.Point(473, 12)
        Me.CancBTN.Name = "CancBTN"
        Me.CancBTN.Size = New System.Drawing.Size(75, 23)
        Me.CancBTN.TabIndex = 8
        Me.CancBTN.Text = "Cancelar"
        Me.CancBTN.UseVisualStyleBackColor = True
        Me.CancBTN.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 400)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(664, 22)
        Me.StatusStrip1.TabIndex = 10
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.ForeColor = System.Drawing.Color.Red
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(200, 16)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.ForeColor = System.Drawing.Color.Red
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(0, 17)
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Email
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(664, 422)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.CancBTN)
        Me.Controls.Add(Me.ExportBTN)
        Me.Controls.Add(Me.ExeBTN)
        Me.Controls.Add(Me.Emails)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Email"
        Me.Text = "E-mail Extractor"
        CType(Me.Emails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Emails As System.Windows.Forms.DataGridView
    Friend WithEvents ExeBTN As System.Windows.Forms.Button
    Friend WithEvents ExportBTN As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents E_mails As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Extractor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents CancBTN As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
