<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ClientSyncUC
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ClientSyncUC))
        Me.NotifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.PropertyGrid = New System.Windows.Forms.PropertyGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button1 = New System.Windows.Forms.Button
        Me.CtxMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ImpostazioniToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuEsci = New System.Windows.Forms.ToolStripMenuItem
        Me.BWPipe = New System.ComponentModel.BackgroundWorker
        Me.FileSystemWatcher = New System.IO.FileSystemWatcher
        Me.Panel1.SuspendLayout()
        Me.CtxMenuStrip.SuspendLayout()
        CType(Me.FileSystemWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NotifyIcon
        '
        Me.NotifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon.Text = "NotifyIcon"
        Me.NotifyIcon.Visible = True
        '
        'PropertyGrid
        '
        Me.PropertyGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PropertyGrid.Location = New System.Drawing.Point(0, 0)
        Me.PropertyGrid.Name = "PropertyGrid"
        Me.PropertyGrid.Size = New System.Drawing.Size(520, 329)
        Me.PropertyGrid.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 329)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(520, 64)
        Me.Panel1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.Button1.Location = New System.Drawing.Point(203, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(115, 50)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Conferma"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CtxMenuStrip
        '
        Me.CtxMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImpostazioniToolStripMenuItem, Me.mnuEsci})
        Me.CtxMenuStrip.Name = "ContextMenuStrip"
        Me.CtxMenuStrip.Size = New System.Drawing.Size(143, 48)
        '
        'ImpostazioniToolStripMenuItem
        '
        Me.ImpostazioniToolStripMenuItem.Image = Global.ITAmbiente.easyrdc.Sync.My.Resources.Resources.Gear
        Me.ImpostazioniToolStripMenuItem.Name = "ImpostazioniToolStripMenuItem"
        Me.ImpostazioniToolStripMenuItem.Size = New System.Drawing.Size(142, 22)
        Me.ImpostazioniToolStripMenuItem.Text = "&Impostazioni"
        '
        'mnuEsci
        '
        Me.mnuEsci.Name = "mnuEsci"
        Me.mnuEsci.Size = New System.Drawing.Size(142, 22)
        Me.mnuEsci.Text = "&Uscita"
        '
        'BWPipe
        '
        '
        'FileSystemWatcher
        '
        Me.FileSystemWatcher.EnableRaisingEvents = True
        Me.FileSystemWatcher.SynchronizingObject = Me
        '
        'ClientSyncUC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(520, 393)
        Me.Controls.Add(Me.PropertyGrid)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ClientSyncUC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ITAmbiente - Client Sync DB"
        Me.Panel1.ResumeLayout(False)
        Me.CtxMenuStrip.ResumeLayout(False)
        CType(Me.FileSystemWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents NotifyIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents PropertyGrid As System.Windows.Forms.PropertyGrid
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CtxMenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuEsci As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImpostazioniToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BWPipe As System.ComponentModel.BackgroundWorker
    Friend WithEvents FileSystemWatcher As System.IO.FileSystemWatcher

End Class
