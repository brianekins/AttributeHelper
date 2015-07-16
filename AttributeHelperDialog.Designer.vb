<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AttributeHelperDialog
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AttributeHelperDialog))
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.treeView = New System.Windows.Forms.TreeView
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.ctxMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AddAttributeSetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteAllAttributesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AddAttributeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EditNameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteAttributeSetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EditAttributeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteAttributeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.btnExpandAll = New System.Windows.Forms.Button
        Me.btnCollapseAll = New System.Windows.Forms.Button
        Me.lblVersion = New System.Windows.Forms.Label
        Me.ctxMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.Location = New System.Drawing.Point(182, 428)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(255, 428)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'treeView
        '
        Me.treeView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.treeView.HideSelection = False
        Me.treeView.ImageIndex = 0
        Me.treeView.ImageList = Me.ImageList
        Me.treeView.Location = New System.Drawing.Point(12, 12)
        Me.treeView.Name = "treeView"
        Me.treeView.SelectedImageIndex = 0
        Me.treeView.Size = New System.Drawing.Size(310, 410)
        Me.treeView.TabIndex = 1
        '
        'ImageList
        '
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList.Images.SetKeyName(0, "ClosedFolder")
        Me.ImageList.Images.SetKeyName(1, "OpenFolder")
        Me.ImageList.Images.SetKeyName(2, "Attribute")
        Me.ImageList.Images.SetKeyName(3, "ArrowWhite")
        Me.ImageList.Images.SetKeyName(4, "ArrowRed")
        '
        'ctxMenu
        '
        Me.ctxMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddAttributeSetToolStripMenuItem, Me.DeleteAllAttributesToolStripMenuItem, Me.AddAttributeToolStripMenuItem, Me.EditNameToolStripMenuItem, Me.DeleteAttributeSetToolStripMenuItem, Me.EditAttributeToolStripMenuItem, Me.DeleteAttributeToolStripMenuItem})
        Me.ctxMenu.Name = "ctxMenu"
        Me.ctxMenu.Size = New System.Drawing.Size(254, 158)
        '
        'AddAttributeSetToolStripMenuItem
        '
        Me.AddAttributeSetToolStripMenuItem.Name = "AddAttributeSetToolStripMenuItem"
        Me.AddAttributeSetToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.AddAttributeSetToolStripMenuItem.Text = "Add Attribute Set"
        '
        'DeleteAllAttributesToolStripMenuItem
        '
        Me.DeleteAllAttributesToolStripMenuItem.Name = "DeleteAllAttributesToolStripMenuItem"
        Me.DeleteAllAttributesToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.DeleteAllAttributesToolStripMenuItem.Text = "Delete All Attributes On this Entity"
        '
        'AddAttributeToolStripMenuItem
        '
        Me.AddAttributeToolStripMenuItem.Name = "AddAttributeToolStripMenuItem"
        Me.AddAttributeToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.AddAttributeToolStripMenuItem.Text = "Add Attribute"
        '
        'EditNameToolStripMenuItem
        '
        Me.EditNameToolStripMenuItem.Name = "EditNameToolStripMenuItem"
        Me.EditNameToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.EditNameToolStripMenuItem.Text = "Edit Name"
        '
        'DeleteAttributeSetToolStripMenuItem
        '
        Me.DeleteAttributeSetToolStripMenuItem.Name = "DeleteAttributeSetToolStripMenuItem"
        Me.DeleteAttributeSetToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.DeleteAttributeSetToolStripMenuItem.Text = "Delete"
        '
        'EditAttributeToolStripMenuItem
        '
        Me.EditAttributeToolStripMenuItem.Name = "EditAttributeToolStripMenuItem"
        Me.EditAttributeToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.EditAttributeToolStripMenuItem.Text = "Edit"
        '
        'DeleteAttributeToolStripMenuItem
        '
        Me.DeleteAttributeToolStripMenuItem.Name = "DeleteAttributeToolStripMenuItem"
        Me.DeleteAttributeToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.DeleteAttributeToolStripMenuItem.Text = "Delete"
        '
        'btnExpandAll
        '
        Me.btnExpandAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExpandAll.Location = New System.Drawing.Point(12, 428)
        Me.btnExpandAll.Name = "btnExpandAll"
        Me.btnExpandAll.Size = New System.Drawing.Size(67, 23)
        Me.btnExpandAll.TabIndex = 0
        Me.btnExpandAll.Text = "Expand All"
        '
        'btnCollapseAll
        '
        Me.btnCollapseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCollapseAll.Location = New System.Drawing.Point(81, 428)
        Me.btnCollapseAll.Name = "btnCollapseAll"
        Me.btnCollapseAll.Size = New System.Drawing.Size(76, 23)
        Me.btnCollapseAll.TabIndex = 0
        Me.btnCollapseAll.Text = "Collapse All"
        '
        'lblVersion
        '
        Me.lblVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.Location = New System.Drawing.Point(270, -1)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(55, 12)
        Me.lblVersion.TabIndex = 2
        Me.lblVersion.Text = "Version: x.x"
        '
        'AttributeHelperDialog
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(334, 464)
        Me.Controls.Add(Me.treeView)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.btnExpandAll)
        Me.Controls.Add(Me.btnCollapseAll)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(350, 500)
        Me.Name = "AttributeHelperDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Attribute Helper"
        Me.ctxMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents treeView As System.Windows.Forms.TreeView
    Friend WithEvents ImageList As System.Windows.Forms.ImageList
    Friend WithEvents ctxMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents DeleteAttributeSetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditAttributeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddAttributeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddAttributeSetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteAllAttributesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditNameToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteAttributeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnExpandAll As System.Windows.Forms.Button
    Friend WithEvents btnCollapseAll As System.Windows.Forms.Button
    Friend WithEvents lblVersion As System.Windows.Forms.Label

End Class
