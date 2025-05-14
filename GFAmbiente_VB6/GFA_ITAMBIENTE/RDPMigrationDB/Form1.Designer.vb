<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.lblDataBasePartenza = New System.Windows.Forms.Label
        Me.lblDataBaseDestinazione = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.lblInfo = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(89, 31)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Partenza"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 79)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(89, 31)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Destinazione"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'lblDataBasePartenza
        '
        Me.lblDataBasePartenza.Location = New System.Drawing.Point(12, 46)
        Me.lblDataBasePartenza.Name = "lblDataBasePartenza"
        Me.lblDataBasePartenza.Size = New System.Drawing.Size(358, 30)
        Me.lblDataBasePartenza.TabIndex = 2
        Me.lblDataBasePartenza.Text = "Label1"
        '
        'lblDataBaseDestinazione
        '
        Me.lblDataBaseDestinazione.Location = New System.Drawing.Point(12, 113)
        Me.lblDataBaseDestinazione.Name = "lblDataBaseDestinazione"
        Me.lblDataBaseDestinazione.Size = New System.Drawing.Size(358, 39)
        Me.lblDataBaseDestinazione.TabIndex = 3
        Me.lblDataBaseDestinazione.Text = "Label2"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(107, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(263, 28)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Indicare il database da cui partire per la migrazione dei dati"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(107, 82)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(263, 28)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Indicare il database in cui migrare i dati presi dal DB  ""Partenza"" "
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(19, 220)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(352, 41)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Avvio Migrazione"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(19, 280)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(349, 36)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 7
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'lblInfo
        '
        Me.lblInfo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblInfo.Location = New System.Drawing.Point(20, 264)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(349, 13)
        Me.lblInfo.TabIndex = 8
        Me.lblInfo.Text = "Label1"
        Me.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(16, 173)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(352, 41)
        Me.Button4.TabIndex = 9
        Me.Button4.Text = "Generazione Script"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(17, 126)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(352, 41)
        Me.Button5.TabIndex = 10
        Me.Button5.Text = "Avvio Analisi"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(387, 325)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblDataBaseDestinazione)
        Me.Controls.Add(Me.lblDataBasePartenza)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Form1"
        Me.Text = "Migrazione DataBase RDP"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents lblDataBasePartenza As System.Windows.Forms.Label
    Friend WithEvents lblDataBaseDestinazione As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button

End Class
