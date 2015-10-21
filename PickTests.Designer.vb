<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PickTests
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PickTests))
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.DataGridView200 = New System.Windows.Forms.DataGridView()
        Me.Button200 = New System.Windows.Forms.Button()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.DataGridView200, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Button200)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.DataGridView200)
        Me.SplitContainer2.Size = New System.Drawing.Size(551, 392)
        Me.SplitContainer2.SplitterDistance = 30
        Me.SplitContainer2.TabIndex = 2
        '
        'DataGridView200
        '
        Me.DataGridView200.AllowUserToAddRows = False
        Me.DataGridView200.AllowUserToDeleteRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.Gainsboro
        Me.DataGridView200.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView200.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView200.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView200.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView200.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView200.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView200.Name = "DataGridView200"
        Me.DataGridView200.RowHeadersVisible = False
        Me.DataGridView200.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridView200.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DataGridView200.Size = New System.Drawing.Size(551, 358)
        Me.DataGridView200.TabIndex = 3
        '
        'Button200
        '
        Me.Button200.BackgroundImage = Global.VRMSYS___MiniLab.My.Resources.Resources.SAVER
        Me.Button200.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button200.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button200.FlatAppearance.BorderSize = 0
        Me.Button200.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button200.Location = New System.Drawing.Point(0, 0)
        Me.Button200.Name = "Button200"
        Me.Button200.Size = New System.Drawing.Size(130, 30)
        Me.Button200.TabIndex = 54
        Me.Button200.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button200.UseVisualStyleBackColor = True
        '
        'PickTests
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(551, 392)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "PickTests"
        Me.Text = "PickTests"
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.DataGridView200, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Button200 As System.Windows.Forms.Button
    Friend WithEvents DataGridView200 As System.Windows.Forms.DataGridView
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
End Class
