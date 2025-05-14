<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Setting
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Setting))
        Me.NotifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.PropertyGrid = New System.Windows.Forms.PropertyGrid
        Me.CtxMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AvvioToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator
        Me.ImpostazioniToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UscitaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FileSystemWatcher = New System.IO.FileSystemWatcher
        Me.Button1 = New System.Windows.Forms.Button
        Me.BWPipe = New System.ComponentModel.BackgroundWorker
        Me.CtxMenuStrip.SuspendLayout()
        CType(Me.FileSystemWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NotifyIcon
        '
        Me.NotifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon.BalloonTipText = "Sincronizzazione dati IsolaEcologica con le rilevazioni effettuate con le unità m" & _
            "obile esterne"
        Me.NotifyIcon.BalloonTipTitle = "Sync ITAmbiente"
        Me.NotifyIcon.Icon = CType(resources.GetObject("NotifyIcon.Icon"), System.Drawing.Icon)
        Me.NotifyIcon.Text = "NotifyIcon1"
        Me.NotifyIcon.Visible = True
        '
        'PropertyGrid
        '
        Me.PropertyGrid.Dock = System.Windows.Forms.DockStyle.Top
        Me.PropertyGrid.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PropertyGrid.Location = New System.Drawing.Point(0, 0)
        Me.PropertyGrid.Name = "PropertyGrid"
        Me.PropertyGrid.Size = New System.Drawing.Size(520, 341)
        Me.PropertyGrid.TabIndex = 0
        '
        'CtxMenuStrip
        '
        Me.CtxMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AvvioToolStripMenuItem, Me.ToolStripMenuItem1, Me.ImpostazioniToolStripMenuItem, Me.UscitaToolStripMenuItem})
        Me.CtxMenuStrip.Name = "ContextMenuStrip"
        Me.CtxMenuStrip.Size = New System.Drawing.Size(143, 76)
        '
        'AvvioToolStripMenuItem
        '
        Me.AvvioToolStripMenuItem.Name = "AvvioToolStripMenuItem"
        Me.AvvioToolStripMenuItem.Size = New System.Drawing.Size(142, 22)
        Me.AvvioToolStripMenuItem.Text = "&Avvio"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(139, 6)
        '
        'ImpostazioniToolStripMenuItem
        '
        Me.ImpostazioniToolStripMenuItem.Image = Global.ITAmbiente.easyrdc.Sync.My.Resources.Resources.Gear
        Me.ImpostazioniToolStripMenuItem.Name = "ImpostazioniToolStripMenuItem"
        Me.ImpostazioniToolStripMenuItem.Size = New System.Drawing.Size(142, 22)
        Me.ImpostazioniToolStripMenuItem.Text = "&Impostazioni"
        '
        'UscitaToolStripMenuItem
        '
        Me.UscitaToolStripMenuItem.Name = "UscitaToolStripMenuItem"
        Me.UscitaToolStripMenuItem.Size = New System.Drawing.Size(142, 22)
        Me.UscitaToolStripMenuItem.Text = "&Uscita"
        '
        'FileSystemWatcher
        '
        Me.FileSystemWatcher.EnableRaisingEvents = True
        Me.FileSystemWatcher.SynchronizingObject = Me
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(211, 347)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(101, 57)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Conferma"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'BWPipe
        '
        Me.BWPipe.WorkerSupportsCancellation = True
        '
        'Setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(520, 416)
        Me.Controls.Add(Me.PropertyGrid)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Setting"
        Me.Text = "ITAmbiente - Server Sync DB"
        Me.CtxMenuStrip.ResumeLayout(False)
        CType(Me.FileSystemWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents NotifyIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents PropertyGrid As System.Windows.Forms.PropertyGrid
    Friend WithEvents CtxMenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AvvioToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ImpostazioniToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UscitaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileSystemWatcher As System.IO.FileSystemWatcher
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BWPipe As System.ComponentModel.BackgroundWorker

End Class
