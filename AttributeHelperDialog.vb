Imports System.Windows.Forms

<System.Runtime.InteropServices.ComVisible(False)> Public Class AttributeHelperDialog
    Private m_InventorDoc As Inventor.Document = Nothing
    Private m_editsMade As Boolean = False
    Private m_Entities As MyEntities
    Private m_HighlightSet As Inventor.HighlightSet
    Private m_NodeCount As Integer = 0
    Private m_ObjectList As New Collection
    Private m_SelectedEntity As Object
    Private m_LastMousePosition As System.Drawing.Point

    Private WithEvents m_DocEvents As Inventor.DocumentEvents
    Private WithEvents m_InputEvents As Inventor.UserInputEvents
    Private WithEvents m_AddAttributeSetButtonDef As Inventor.ButtonDefinition
    Private WithEvents m_FindAttributeInDialogButtonDef As Inventor.ButtonDefinition


    Public Property InventorDocument() As Inventor.Document
        Get
            Return m_InventorDoc
        End Get
        Set(ByVal value As Inventor.Document)
            m_InventorDoc = value
        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Try
            ' Apply the results.
            If m_Entities.HasChanges Then
                For Each entity As MyEntity In m_Entities
                    UpdateEntity(entity)
                Next
            End If

            CleanUp()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox("Unexepected failure exiting the form with OK.")
        End Try
    End Sub

    Private Sub UpdateEntity(ByVal Entity As MyEntity)
        Try
            For Each attribSet As MyAttributeSet In Entity.AttributeSets
                Dim invAttribSet As Inventor.AttributeSet

                ' Check for any new attribute sets.
                If attribSet.IsNew Then
                    invAttribSet = Entity.Entity.AttributeSets.Add(attribSet.Name)

                    ' Add its attributes.
                    For Each attrib As MyAttribute In attribSet
                        invAttribSet.Add(attrib.Name, attrib.AttributeType, attrib.Value)
                    Next
                ElseIf attribSet.IsDeleted Then
                    Entity.Entity.AttributeSets.Item(attribSet.Name).Delete()
                ElseIf attribSet.IsEdited Then
                    invAttribSet = Entity.Entity.AttributeSets.item(attribSet.OriginalName)

                    If invAttribSet.Name <> attribSet.Name Then
                        invAttribSet.Name = attribSet.Name
                    End If

                    ' Update the individual attributes.
                    For Each attrib As MyAttribute In attribSet
                        If attrib.IsDeleted Then
                            ' Delete the attribute.
                            invAttribSet.Item(attrib.Name).Delete()
                        ElseIf attrib.IsNew Then
                            ' Create the attribute.
                            invAttribSet.Add(attrib.Name, attrib.AttributeType, attrib.Value)
                        ElseIf attrib.IsEdited Then
                            Dim invAttrib As Inventor.Attribute = invAttribSet.Item(attrib.OriginalName)

                            If invAttrib.Name <> attrib.Name Then
                                invAttrib.Name = attrib.Name
                            End If

                            invAttrib.Value = attrib.Value
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox("Unexepected Failure while updating the attribute information.")
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Try
            If m_Entities.HasChanges Then
                If MsgBox("Canceling will lose the edits that have been made.  Do you want to continue and loose them?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            CleanUp()
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox("Unexepected failure exiting the form with Cancel.")
        End Try
    End Sub

    Private Sub CleanUp()
        If Not m_HighlightSet Is Nothing Then
            m_HighlightSet.Clear()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(m_HighlightSet)
            m_HighlightSet = Nothing
        End If

        m_DocEvents = Nothing
        m_InputEvents = Nothing
        m_AddAttributeSetButtonDef = Nothing
        m_FindAttributeInDialogButtonDef = Nothing
        Me.InventorDocument = Nothing

        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()

        m_Entities = Nothing
        m_ObjectList = Nothing
    End Sub

    Private Sub AttributeHelperDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim version As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            lblVersion.Text = "Version: " & version.Major & "." & version.Minor

            If Not Me.InventorDocument Is Nothing Then
                Me.Left = Me.InventorDocument.Parent.Left + 25
                Me.Top = Me.InventorDocument.Parent.Top + 25

                Dim commandManager As Inventor.CommandManager = Me.InventorDocument.Parent.CommandManager
                Try
                    m_AddAttributeSetButtonDef = commandManager.ControlDefinitions.Item("ekinsAddAttributeSet")
                    m_FindAttributeInDialogButtonDef = commandManager.ControlDefinitions.Item("ekinsFindInDialog")
                Catch ex As Exception
                    m_AddAttributeSetButtonDef = commandManager.ControlDefinitions.AddButtonDefinition("Add Attribute Set", "ekinsAddAttributeSet", Inventor.CommandTypesEnum.kNonShapeEditCmdType, "{5e98b9dd-6c1a-4ce8-b157-428250cc5e5e}", "Add an attribute set to the selected entity.", "Add Attribute Set")
                    m_FindAttributeInDialogButtonDef = commandManager.ControlDefinitions.AddButtonDefinition("Find Object in Dialog", "ekinsFindInDialog", Inventor.CommandTypesEnum.kNonShapeEditCmdType, "{5e98b9dd-6c1a-4ce8-b157-428250cc5e5e}", "Find this entity in the Attribute Helper dialog.", "Find Object in Dialog")
                    commandManager.CommandCategories.Item("FWxContextMenuCategory").Add(m_AddAttributeSetButtonDef)
                    commandManager.CommandCategories.Item("FWxContextMenuCategory").Add(m_FindAttributeInDialogButtonDef)
                End Try

                Me.InventorDocument.Parent.CommandManager.ControlDefinitions.Item("AppSelectNorthwestArrowCmd").Execute()
                m_DocEvents = Me.InventorDocument.DocumentEvents
                m_InputEvents = Me.InventorDocument.Parent.CommandManager.UserInputEvents

                GetAttributes()

                ' Create the tree to show the attributes.
                Me.treeView.BeginUpdate()
                For Each entity As MyEntity In m_Entities
                    m_ObjectList.Add(entity)

                    m_NodeCount += 1
                    Dim entityNode As TreeNode = Me.treeView.Nodes.Add(m_NodeCount.ToString, GoodEntityName(entity.Entity), 3)
                    entityNode.ContextMenuStrip = ctxMenu
                    entity.TreeNode = entityNode
                    entityNode.ImageIndex = 3
                    entityNode.SelectedImageIndex = 3
                    entityNode.StateImageKey = 3

                    For Each attribSet As MyAttributeSet In entity.AttributeSets
                        m_ObjectList.Add(attribSet)

                        m_NodeCount += 1
                        Dim attribSetNode As TreeNode = entityNode.Nodes.Add(m_NodeCount.ToString, attribSet.Name, 0)
                        attribSetNode.ContextMenuStrip = ctxMenu
                        attribSet.TreeNode = attribSetNode
                        attribSetNode.ImageIndex = 0
                        attribSetNode.SelectedImageIndex = 0
                        attribSetNode.StateImageIndex = 0

                        For Each attrib As MyAttribute In attribSet
                            m_ObjectList.Add(attrib)

                            m_NodeCount += 1
                            Dim attribNode As TreeNode = attribSetNode.Nodes.Add(m_NodeCount.ToString, attrib.Name & " = " & attrib.ValueAsString, 2)
                            attribNode.ContextMenuStrip = ctxMenu
                            attrib.TreeNode = attribNode
                            attribNode.ImageIndex = 2
                            attribNode.SelectedImageIndex = 2
                            attribNode.StateImageIndex = 2
                        Next
                    Next
                Next

                Me.treeView.EndUpdate()

                '' Kick the tree so that everything is displayed.
                'Me.treeView.ExpandAll()
                'Me.treeView.CollapseAll()
                'Me.treeView.Nodes.Item(0).EnsureVisible()

                m_HighlightSet = Me.InventorDocument.CreateHighlightSet
            End If
        Catch ex As Exception
            MsgBox("Unexepected failure while loading the form.")
        End Try
    End Sub

    Private Function GoodEntityName(ByVal Entity As Object) As String
        Dim tempName As String = TypeName(Entity)
        If tempName.Substring(0, 3).ToLower = "irx" Then
            Return tempName.Substring(3)
        Else
            Return tempName
        End If
    End Function

    Private Sub GetAttributes()
        Dim attribManager As Inventor.AttributeManager = Me.InventorDocument.AttributeManager
        Dim attributedEntities As Inventor.ObjectCollection = attribManager.FindObjects("*", "*")

        ' Reinitialize the data.
        m_Entities = New MyEntities

        g_IgnoreDuringLoad = True

        ' Iterate through the entities that have attributes.
        For Each invEntity As Object In attributedEntities
            Dim currentEntity As MyEntity = m_Entities.Add(invEntity)

            Dim attSets As Inventor.AttributeSets
            attSets = invEntity.AttributeSets

            For Each attSet As Inventor.AttributeSet In invEntity.AttributeSets
                Dim currentAttSet As MyAttributeSet = currentEntity.AttributeSets.Add(attSet.Name)

                For Each attrib As Inventor.Attribute In attSet
                    currentAttSet.Add(attrib.Name, attrib.ValueType, attrib.Value)
                Next
            Next
        Next

        ' Get the attribute sets that aren't associated with any entity.  This happens
        ' when the entity is deleted or consumed in some operation.  The attribute
        ' set is not automatically cleaned up.
        Dim detachedAttribSets As Object = Nothing
        attribManager.PurgeAttributeSets("*", True, detachedAttribSets)

        If Not detachedAttribSets Is Nothing Then
            For Each attSet As Inventor.AttributeSet In detachedAttribSets
                Dim currentEntity As MyEntity = m_Entities.Add(Nothing)

                Dim currentAttSet As MyAttributeSet = currentEntity.AttributeSets.Add(attSet.Name)
                For Each attrib As Inventor.Attribute In attSet
                    Call currentAttSet.Add(attrib.Name, attrib.Type, attrib.Value)
                Next
            Next
        End If

        g_IgnoreDuringLoad = False
    End Sub

    Private Sub treeView_AfterCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeView.AfterCollapse
        If Not e.Node.Parent Is Nothing Then
            e.Node.ImageIndex = 0
            e.Node.SelectedImageIndex = 0
            e.Node.StateImageIndex = 0
        End If
    End Sub

    Private Sub treeView_AfterExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeView.AfterExpand
        If Not e.Node.Parent Is Nothing Then
            e.Node.ImageIndex = 1
            e.Node.SelectedImageIndex = 1
            e.Node.StateImageIndex = 1
        End If
    End Sub

    Private Sub treeView_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles treeView.AfterLabelEdit
        ' Check to make sure there aren't any spaces in the attribute set name.
        If e.Label.Trim.Contains(" ") Then
            MsgBox("An attribute set name cannot contain any spaces.")
            e.CancelEdit = True
        End If

        ' Get the entity from the edited tree node.
        Dim ent As MyEntity = m_Entities.Item(e.Node.Parent)

        ' Find the attribute set.
        If Not ent Is Nothing Then
            For Each attSet As MyAttributeSet In ent.AttributeSets
                If e.Node.Text = attSet.Name Then
                    ' Reassign the name.
                    attSet.Name = e.Label.Trim
                End If
            Next
        End If
    End Sub

    Private Sub treeView_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeView.AfterSelect
        Try
            ' Clear the select set.
            Me.InventorDocument.SelectSet.Clear()

            ' Special case based on the type of node.
            Select Case e.Node.ImageIndex
                Case 3, 4  ' Entity node
                    ' Get the selected entity.
                    Dim entity As MyEntity = m_ObjectList.Item(CType(e.Node.Name, Integer))
                    SelectEntityNode(entity)
                Case 0, 1  ' AttributeSet node
                    Dim attribSet As MyAttributeSet = m_ObjectList.Item(CType(e.Node.Name, Integer))
                    SelectEntityNode(attribSet.Parent.Parent)
                Case 2 ' Attribute
                    Dim attrib As MyAttribute = m_ObjectList.Item(CType(e.Node.Name, Integer))
                    SelectEntityNode(attrib.Parent.Parent.Parent)
            End Select
        Catch ex As Exception
            ' Do nothing.  MsgBox("Unexepected failure highlighting the selected node.")
        End Try
    End Sub

    Private Sub SelectEntityNode(ByVal Entity As MyEntity)
        ' Clear the current highlight set.
        m_HighlightSet.Clear()

        ' Highlight the selected entity.
        If Not Entity.Entity Is Nothing Then
            m_HighlightSet.AddItem(Entity.Entity)

            ' Change all entity nodes to be an unselected arrow.
            Dim checkNode As TreeNode = Me.treeView.TopNode
            Do
                checkNode.ImageIndex = 3
                checkNode.SelectedImageIndex = 3
                checkNode.StateImageIndex = 3
                checkNode = checkNode.NextNode
            Loop While Not checkNode Is Nothing

            ' Change the selected node to be a selected arrow.
            Entity.TreeNode.ImageIndex = 4
            Entity.TreeNode.SelectedImageIndex = 4
            Entity.TreeNode.StateImageIndex = 4
        End If
    End Sub

    Private Sub treeView_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles treeView.NodeMouseClick
        Try
            treeView.SelectedNode = e.Node
            m_LastMousePosition = e.Location
            m_LastMousePosition.X = m_LastMousePosition.X + Me.Left
            m_LastMousePosition.Y = m_LastMousePosition.Y + Me.Top
        Catch ex As Exception
            MsgBox("Unexepected failure while selecting the tree node.")
        End Try
    End Sub

    Private Sub treeView_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles treeView.NodeMouseDoubleClick
        Try
            Select Case e.Node.ImageIndex
                Case 0, 1  ' AttributeSet node
                    treeView.LabelEdit = True
                    e.Node.BeginEdit()
                Case 2 ' Attribute
                    treeView.SelectedNode = e.Node
                    m_LastMousePosition = e.Location
                    m_LastMousePosition.X = m_LastMousePosition.X + Me.Left
                    m_LastMousePosition.Y = m_LastMousePosition.Y + Me.Top

                    EditAttributeWithDialog()
            End Select
        Catch ex As Exception
            MsgBox("Unexepected failure while double-clicking the tree node.")
        End Try
    End Sub

    Private Sub m_InputEvents_OnActivateCommand(CommandName As String, Context As Inventor.NameValueMap) Handles m_InputEvents.OnActivateCommand
        ' Another command was started, so terminate this program.  The reason
        ' this is being used instead of InteractionEvents is because I want to
        ' use the more general NW arrow selection which has a wider filter range
        ' and is controlled by the user using the selection options.  One issue is
        ' that I don't get notified is the Escape key is pressed while in the NW arrow
        ' command so I can't terminate based on that.
        If Not (CommandName.Contains("ViewCmd") Or CommandName.Contains("WindowCmd")) Then
            m_HighlightSet.Clear()
            Me.Close()
        End If
    End Sub


    Private Sub ctxMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ctxMenu.Opening
        Try
            ' Determine context.
            Select Case treeView.SelectedNode.ImageIndex
                Case 3, 4  ' Entity node
                    ctxMenu.Items(0).Visible = True
                    ctxMenu.Items(1).Visible = True
                    ctxMenu.Items(2).Visible = False
                    ctxMenu.Items(3).Visible = False
                    ctxMenu.Items(4).Visible = False
                    ctxMenu.Items(5).Visible = False
                    ctxMenu.Items(6).Visible = False

                    e.Cancel = False
                Case 0, 1  ' AttributeSet node
                    ctxMenu.Items(0).Visible = False
                    ctxMenu.Items(1).Visible = False
                    ctxMenu.Items(2).Visible = True
                    ctxMenu.Items(3).Visible = True
                    ctxMenu.Items(4).Visible = True
                    ctxMenu.Items(5).Visible = False
                    ctxMenu.Items(6).Visible = False

                    e.Cancel = False
                Case 2 ' Attribute
                    ctxMenu.Items(0).Visible = False
                    ctxMenu.Items(1).Visible = False
                    ctxMenu.Items(2).Visible = False
                    ctxMenu.Items(3).Visible = False
                    ctxMenu.Items(4).Visible = False
                    ctxMenu.Items(5).Visible = True
                    ctxMenu.Items(6).Visible = True

                    e.Cancel = False
            End Select
        Catch ex As Exception
            MsgBox("Unexepected failure while displaying the treeview context menu.")
        End Try
    End Sub

    Private Function EditAttribute(ByVal AttributeName As String, ByVal AttributeValue As String, ByVal EditNode As TreeNode) As Boolean
        AttributeName = AttributeName.Trim
        If AttributeName.Contains(" ") Then
            MsgBox("Attribute names cannot contain any spaces.")
            Return False
        End If

        ' Get the associated Attribute object.
        Dim attrib As MyAttribute = m_ObjectList.Item(CType(EditNode.Name, Integer))
        Dim attribSet As MyAttributeSet = attrib.Parent
        Dim entity As MyEntity = attribSet.Parent.Parent

        ' Check to see if anything has changed and that the new values are valid.
        If attrib.Name <> AttributeName Then
            ' Check if any other attributes have this name.
            For Each checkAttrib As MyAttribute In attribSet
                If checkAttrib.Name = AttributeName Then
                    MsgBox("The name of the attribute must be unique within this attribute set.")
                    Return False
                End If
            Next

            attrib.Name = AttributeName
        End If

        If attrib.ValueAsFullString <> AttributeValue Then
            ' Check that the specified string is valid for the value type.
            Select Case attrib.AttributeType
                Case Inventor.ValueTypeEnum.kStringType
                    attrib.Value = AttributeValue
                Case Inventor.ValueTypeEnum.kIntegerType
                    Dim intValue As Short
                    Try
                        intValue = CType(AttributeValue, Short)
                    Catch ex As Exception
                        MsgBox("The value entered is not a valid Integer type.")
                        Return False
                    End Try
                    attrib.Value = intValue
                Case Inventor.ValueTypeEnum.kDoubleType
                    Dim dblValue As Double
                    Try
                        dblValue = CType(AttributeValue, Double)
                    Catch ex As Exception
                        MsgBox("The value entered is not a valid Double type.")
                        Return False
                    End Try
                    attrib.Value = dblValue
                Case Inventor.ValueTypeEnum.kByteArrayType
                    Dim values() As String = AttributeValue.Split(","c)
                    Dim newArray() As Byte
                    ReDim newArray(values.Length - 1)
                    Dim i As Integer = 0
                    For Each checkValue As String In values
                        Dim bytValue As Byte
                        Try
                            bytValue = CType(checkValue, Byte)
                        Catch ex As Exception
                            MsgBox("The value """ & checkValue & """ is not a valid Byte value")
                            Return False
                        End Try

                        newArray(i) = CType(checkValue, Byte)
                        i += 1
                    Next

                    attrib.Value = newArray
            End Select
        End If

        EditNode.Text = attrib.Name & " = " & attrib.ValueAsString

        Return True
    End Function

    Private Function CreateAttribute(ByVal AttributeName As String, ByVal AttributeType As Inventor.ValueTypeEnum, ByVal AttributeValue As String, ByVal EditNode As TreeNode) As Boolean
        AttributeName = AttributeName.Trim
        If AttributeName.Contains(" ") Then
            MsgBox("Attribute names cannot contain any spaces.")
            Return False
        End If

        ' Get the associated Attribute set object.
        Dim attribSet As MyAttributeSet = m_ObjectList.Item(CType(Me.treeView.SelectedNode.Name, Integer))
        Dim entity As MyEntity = attribSet.Parent.Parent

        ' Check that the specified string is valid for the value type.
        Dim attrib As MyAttribute = Nothing
        Select Case AttributeType
            Case Inventor.ValueTypeEnum.kStringType
                attrib = attribSet.Add(AttributeName, Inventor.ValueTypeEnum.kStringType, AttributeValue)
            Case Inventor.ValueTypeEnum.kIntegerType
                Dim intValue As Short
                Try
                    intValue = CType(AttributeValue, Short)
                Catch ex As Exception
                    MsgBox("The value entered is not a valid Integer type.")
                    Return False
                End Try

                attrib = attribSet.Add(AttributeName, Inventor.ValueTypeEnum.kIntegerType, intValue)
            Case Inventor.ValueTypeEnum.kDoubleType
                Dim dblValue As Double
                Try
                    dblValue = CType(AttributeValue, Double)
                Catch ex As Exception
                    MsgBox("The value entered is not a valid Double type.")
                    Return False
                End Try

                attrib = attribSet.Add(AttributeName, Inventor.ValueTypeEnum.kDoubleType, dblValue)
            Case Inventor.ValueTypeEnum.kByteArrayType
                Dim values() As String = AttributeValue.Split(","c)
                Dim newArray() As Byte
                ReDim newArray(values.Length - 1)
                Dim i As Integer = 0
                For Each checkValue As String In values
                    Dim bytValue As Byte
                    Try
                        bytValue = CType(checkValue, Byte)
                    Catch ex As Exception
                        MsgBox("The value """ & checkValue & """ is not a valid Byte value")
                        Return False
                    End Try

                    newArray(i) = CType(checkValue, Byte)
                    i += 1
                Next

                attrib = attribSet.Add(AttributeName, Inventor.ValueTypeEnum.kByteArrayType, newArray)
                If attrib Is Nothing Then
                    Return False
                End If
        End Select

        ' Add a node to the tree.
        If Not attrib Is Nothing Then
            m_ObjectList.Add(attrib)

            m_NodeCount += 1
            Dim attribNode As TreeNode = attribSet.TreeNode.Nodes.Add(m_NodeCount.ToString, attrib.Name & " = " & attrib.ValueAsString, 2)
            attribNode.ImageIndex = 2
            attribNode.SelectedImageIndex = 2
            attribNode.StateImageIndex = 2
            attribNode.ContextMenuStrip = ctxMenu
            attribNode.EnsureVisible()
        End If

        Return True
    End Function


    Private Sub m_FindAttributeInDialogButtonDef_OnExecute(Context As Inventor.NameValueMap) Handles m_FindAttributeInDialogButtonDef.OnExecute
        ' Select the entity within the dialog.
        If Not m_SelectedEntity Is Nothing Then
            Dim ent As MyEntity = m_Entities.Item(m_SelectedEntity)
            If Not ent Is Nothing Then
                ent.SelectEntity()
            End If
        End If
    End Sub


    Private Sub m_AddAttributeSetButtonDef_OnExecute(Context As Inventor.NameValueMap) Handles m_AddAttributeSetButtonDef.OnExecute
        Try
            Dim attributeSetName As String = ""
            attributeSetName = InputBox("Enter the name of the attribute set.", "Create Attribute Set", , Me.Left + 25, Me.Top + 50)

            If attributeSetName <> "" Then
                If attributeSetName.Trim.Contains(" ") Then
                    MsgBox("An attribute set name cannot contain any spaces.")
                    Exit Sub
                Else
                    attributeSetName = attributeSetName.Trim
                End If

                ' Check to see if this entity is already in the tree.
                Dim existingEntity As Boolean = False
                For Each ent As MyEntity In m_Entities
                    If ent.Entity Is m_SelectedEntity Then
                        ' Check to see that the name is unique.
                        For Each testAttribSet As MyAttributeSet In ent.AttributeSets
                            If testAttribSet.Name.ToLower = attributeSetName.ToLower Then
                                MsgBox("The specified attribute set name is already used" & vbCr & "by an attribute set on the selected entity.", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation)
                                Exit Sub
                            End If
                        Next

                        ' Add a new attribute set to this entity.
                        Dim attribSet As MyAttributeSet = ent.AttributeSets.Add(attributeSetName)
                        attribSet.IsNew = True

                        ' Add a new node to the tree.
                        m_ObjectList.Add(attribSet)

                        m_NodeCount += 1
                        Dim attribSetNode As TreeNode = ent.TreeNode.Nodes.Add(m_NodeCount.ToString, attributeSetName, 0)
                        attribSetNode.ContextMenuStrip = ctxMenu
                        attribSet.TreeNode = attribSetNode
                        attribSetNode.ImageIndex = 0
                        attribSetNode.SelectedImageIndex = 0
                        attribSetNode.StateImageIndex = 0

                        existingEntity = True
                        Exit For
                    End If
                Next

                If Not existingEntity Then
                    ' Create a new entity.
                    Dim newEntity As MyEntity = m_Entities.Add(m_SelectedEntity)
                    newEntity.IsNew = True

                    ' Add a new entity to the tree.
                    m_ObjectList.Add(newEntity)

                    m_NodeCount += 1
                    Dim entityNode As TreeNode = Me.treeView.Nodes.Add(m_NodeCount.ToString, GoodEntityName(newEntity.Entity), 3)
                    newEntity.TreeNode = entityNode
                    entityNode.ContextMenuStrip = ctxMenu
                    entityNode.ImageIndex = 3
                    entityNode.SelectedImageIndex = 3
                    entityNode.StateImageKey = 3

                    ' Create a new attribute set.
                    ' Add a new attribute set to this entity.
                    Dim attribSet As MyAttributeSet = newEntity.AttributeSets.Add(attributeSetName)
                    attribSet.IsNew = True

                    ' Add the attribute set to the tree
                    m_ObjectList.Add(attribSet)

                    m_NodeCount += 1
                    Dim attribSetNode As TreeNode = entityNode.Nodes.Add(m_NodeCount.ToString, attributeSetName, 0)
                    attribSetNode.ContextMenuStrip = ctxMenu
                    attribSet.TreeNode = attribSetNode
                    attribSetNode.ImageIndex = 0
                    attribSetNode.SelectedImageIndex = 0
                    attribSetNode.StateImageIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox("Unexpected failure while adding the attribute set.")
        End Try
    End Sub

    Private Sub AddAttributeSetToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AddAttributeSetToolStripMenuItem.Click
        Try
            Dim attributeSetName As String = InputBox("Enter the name of the attribute set.", "Create Attribute Set", , Me.Left + 25, m_LastMousePosition.Y - 20)
            If attributeSetName <> "" Then
                If attributeSetName.Trim.Contains(" ") Then
                    MsgBox("An attribute set name cannot contain any spaces.")
                    Exit Sub
                End If

                Dim entity As MyEntity = m_ObjectList.Item(CType(treeView.SelectedNode.Name, Integer))
                Dim inventorEntity As Object = entity.Entity

                ' Check that this name is unique for that attribute sets on this entity.
                For Each testAttribSet As MyAttributeSet In entity.AttributeSets
                    If testAttribSet.Name.ToLower = attributeSetName.ToLower Then
                        MsgBox("The specified attribute set name is already used" & vbCr & "by an attribute set on the selected entity.", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If
                Next

                Try
                    ' Add a new attribute set to this entity.
                    Dim attribSet As MyAttributeSet = entity.AttributeSets.Add(attributeSetName)
                    attribSet.IsNew = True

                    ' Add a new node to the tree.
                    m_ObjectList.Add(attribSet)

                    m_NodeCount += 1
                    Dim attribSetNode As TreeNode = entity.TreeNode.Nodes.Add(m_NodeCount.ToString, attributeSetName, 0)
                    attribSetNode.ContextMenuStrip = ctxMenu
                    attribSet.TreeNode = attribSetNode
                    attribSetNode.ImageIndex = 0
                    attribSetNode.SelectedImageIndex = 0
                    attribSetNode.StateImageIndex = 0
                    attribSetNode.EnsureVisible()
                Catch ex As Exception
                    ' Do nothing
                End Try
            End If
        Catch ex As Exception
            MsgBox("Unexepected failure while adding the attribute set.")
        End Try
    End Sub

    Private Sub DeleteAllAttributesToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteAllAttributesToolStripMenuItem.Click
        Try
            Dim entity As MyEntity = m_ObjectList.Item(CType(treeView.SelectedNode.Name, Integer))

            For Each attribSet As MyAttributeSet In entity.AttributeSets
                attribSet.IsDeleted = True
            Next

            entity.TreeNode.Remove()
            entity.TreeNode = Nothing
        Catch ex As Exception
            MsgBox("Unexepected failure while deleting all attributes.")
        End Try
    End Sub

    Private Sub EditNameToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditNameToolStripMenuItem.Click
        Try
            treeView.LabelEdit = True
            treeView.SelectedNode.BeginEdit()
        Catch ex As Exception
            MsgBox("Unexepected failure while editing the attribute set name.")
        End Try
    End Sub

    Private Sub AddAttributeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AddAttributeToolStripMenuItem.Click
        Try
            ' Get the attribute set being the attribute is being added to.
            Dim attribSet As MyAttributeSet = m_ObjectList.Item(CType(treeView.SelectedNode.Name, Integer))

            ' Create the dialog.
            Dim editDialog As New EditAttribute

            ' Populate the dialog with the current information.
            editDialog.EditMode = False

            Dim badAttribute As Boolean = False
            Do
                editDialog.ShowDialog()
                If editDialog.DialogResult = Windows.Forms.DialogResult.OK Then
                    If CreateAttribute(editDialog.AttributeName, editDialog.AttributeType, editDialog.AttributeValue, treeView.SelectedNode) Then
                        badAttribute = False
                    Else
                        badAttribute = True
                    End If
                Else
                    badAttribute = False
                End If
            Loop While badAttribute
        Catch ex As Exception
            MsgBox("Unexepected failure while adding the attribute.")
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteAttributeSetToolStripMenuItem.Click, DeleteAttributeToolStripMenuItem.Click
        Try
            Dim selectedObject As Object = m_ObjectList(CType(treeView.SelectedNode.Name, Integer))
            selectedObject.IsDeleted = True

            treeView.SelectedNode.Remove()
            selectedObject.TreeNode = Nothing
        Catch ex As Exception
            MsgBox("Unexepected failure while deleting the attribute.")
        End Try
    End Sub

    Private Sub EditAttributeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditAttributeToolStripMenuItem.Click
        EditAttributeWithDialog()
    End Sub

    Private Sub EditAttributeWithDialog()
        Try
            ' Get the attribute being edited.
            Dim attrib As MyAttribute = m_ObjectList.Item(CType(treeView.SelectedNode.Name, Integer))

            ' Create the dialog.
            Dim editDialog As New EditAttribute

            ' Populate the dialog with the current information.
            editDialog.AttributeName = attrib.Name
            editDialog.AttributeType = attrib.AttributeType
            editDialog.AttributeValue = attrib.ValueAsFullString

            editDialog.EditMode = True
            Dim badAttribute As Boolean = False
            Do
                editDialog.ShowDialog()
                If editDialog.DialogResult = Windows.Forms.DialogResult.OK Then
                    If EditAttribute(editDialog.AttributeName, editDialog.AttributeValue, treeView.SelectedNode) Then
                        badAttribute = False
                    Else
                        badAttribute = True
                    End If
                Else
                    badAttribute = False
                End If
            Loop While badAttribute
        Catch ex As Exception
            MsgBox("Unexepected failure while editing the attribute.")
        End Try
    End Sub

    Private Sub m_DocEvents_OnChangeSelectSet(ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum)
        If BeforeOrAfter = Inventor.EventTimingEnum.kAfter Then
            ' Clear any selection in the tree.
            If Not Me.treeView.TopNode Is Nothing Then
                Dim checkNode As TreeNode = Me.treeView.TopNode
                Do
                    checkNode.ImageIndex = 3
                    checkNode.SelectedImageIndex = 3
                    checkNode.StateImageIndex = 3
                    checkNode = checkNode.NextNode
                Loop While Not checkNode Is Nothing

                m_HighlightSet.Clear()
            End If
        End If
    End Sub

    Private Sub btnExpandAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandAll.Click
        treeView.ExpandAll()
    End Sub

    Private Sub btnCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCollapseAll.Click
        treeView.CollapseAll()
    End Sub

    Private Sub m_InputEvents_OnRadialMarkingMenu(SelectedEntities As Inventor.ObjectsEnumerator, SelectionDevice As Inventor.SelectionDeviceEnum, RadialMenu As Inventor.RadialMarkingMenu, AdditionalInfo As Inventor.NameValueMap) Handles m_InputEvents.OnRadialMarkingMenu
        RadialMenu.EastControl = Nothing
        RadialMenu.NorthControl = Nothing
        RadialMenu.NortheastControl = Nothing
        RadialMenu.NorthwestControl = Nothing
        RadialMenu.SouthControl = Nothing
        RadialMenu.SoutheastControl = Nothing
        RadialMenu.SouthwestControl = Nothing
        RadialMenu.WestControl = Nothing
    End Sub

    Private Sub m_InputEvents_OnLinearMarkingMenu(SelectedEntities As Inventor.ObjectsEnumerator, SelectionDevice As Inventor.SelectionDeviceEnum, LinearMenu As Inventor.CommandControls, AdditionalInfo As Inventor.NameValueMap) Handles m_InputEvents.OnLinearMarkingMenu
        Try
            ' Check to see if any of the objects in the select set support attributes.
            Dim entityFound As Boolean = False
            Try
                Dim attribSet As Inventor.AttributeSets = SelectedEntities.Item(1).AttributeSets

                ' If we got here then the entity does support attribute sets.
                m_SelectedEntity = Me.InventorDocument.SelectSet.Item(1)
                entityFound = True
            Catch ex As Exception
                m_SelectedEntity = Nothing
            End Try

            ' Remove all of the existing the controls in the list.
            For i As Integer = LinearMenu.Count To 1 Step -1
                LinearMenu.Item(i).Delete()

                'Dim control As Inventor.CommandControl = LinearMenu.Item(i)
                'If Not (control.InternalName.Contains("ViewCmd") Or control.InternalName.Contains("WindowCmd")) Then
                '    control.Delete()
                'End If
            Next

            ' Add the appropriate command.
            If entityFound Then
                ' Check to see if the selected entity already has an attribute.
                If m_SelectedEntity.AttributeSets.Count = 0 Then
                    ' Check to see if the entity has a pending attribute set added in the dialog.
                    Dim testEnt As MyEntity = m_Entities.Item(m_SelectedEntity)
                    If testEnt Is Nothing Then
                        If LinearMenu.Count > 0 Then
                            LinearMenu.AddButton(m_AddAttributeSetButtonDef, , , LinearMenu.Item(1).InternalName, True)
                        Else
                            LinearMenu.AddButton(m_AddAttributeSetButtonDef)
                        End If
                    Else
                        If LinearMenu.Count > 0 Then
                            LinearMenu.AddButton(m_AddAttributeSetButtonDef, , , LinearMenu.Item(1).InternalName, True)
                            LinearMenu.AddButton(m_FindAttributeInDialogButtonDef, , , LinearMenu.Item(1).InternalName, True)
                        Else
                            LinearMenu.AddButton(m_AddAttributeSetButtonDef)
                            LinearMenu.AddButton(m_FindAttributeInDialogButtonDef)
                        End If
                    End If
                Else
                    If LinearMenu.Count > 0 Then
                        LinearMenu.AddButton(m_AddAttributeSetButtonDef, , , LinearMenu.Item(1).InternalName, True)
                        LinearMenu.AddButton(m_FindAttributeInDialogButtonDef, , , LinearMenu.Item(1).InternalName, True)
                    Else
                        LinearMenu.AddButton(m_AddAttributeSetButtonDef)
                        LinearMenu.AddButton(m_FindAttributeInDialogButtonDef)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Unexepected failure while displaying the Inventor context menu.")
        End Try
    End Sub


    ' Disable the contextual mini toolbar.  This is the toolbar that's displayed when you select an entity.
    ' It's disabled by clearing it's contents just before it's displayed.
    'Private Sub m_InputEvents_OnContextualMiniToolbar(ByVal SelectedEntities As Inventor.ObjectsEnumerator, ByVal DisplayedCommands As Inventor.NameValueMap, ByVal AdditionalInfo As Inventor.NameValueMap)
    '    DisplayedCommands.Clear()
    'End Sub
End Class


Friend Class MyEntities
    Inherits CollectionBase

    Public Function Add(ByVal Entity As Object) As MyEntity
        Try
            ' Create a new object
            Dim newMember As MyEntity
            newMember = New MyEntity

            ' Set the properties passed into the method
            newMember.Entity = Entity

            ' Add the object to the collection.
            MyBase.List.Add(newMember)

            ' Return the object created
            Add = newMember
            newMember = Nothing
        Catch ex As Exception
            ' Failure.
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property HasChanges() As Boolean
        Get
            For Each entity As MyEntity In Me
                If entity.IsEdited Or entity.IsNew Then
                    Return True
                    Exit For
                End If

                For Each attribSet As MyAttributeSet In entity.AttributeSets
                    If attribSet.IsDeleted Or attribSet.IsEdited Or attribSet.IsNew Then
                        Return True
                        Exit For
                    End If

                    For Each attrib As MyAttribute In attribSet
                        If attrib.IsDeleted Or attrib.IsEdited Or attrib.IsNew Then
                            Return True
                            Exit For
                        End If
                    Next
                Next
            Next

            Return False
        End Get
    End Property


    ' Returns the specified item of the collection.
    '
    ' indexKey - Index of the station to return.  The first item in the collection has an index of 1.
    Default Public ReadOnly Property Item(ByVal indexKey As Integer) As MyEntity
        Get
            Return CType(MyBase.List.Item(indexKey - 1), MyEntity)
        End Get
    End Property


    ' InventorEntity - The Inventor entity to returns the corresponding MyEntity object for.
    Default Public ReadOnly Property Item(ByVal InventorEntity As Object) As MyEntity
        Get
            For Each ent As MyEntity In Me
                If ent.Entity Is InventorEntity Then
                    Return ent
                End If
            Next

            Return Nothing
        End Get
    End Property

    ' InventorEntity - The Inventor entity to returns the corresponding MyEntity object for.
    Default Public ReadOnly Property Item(ByVal Node As TreeNode) As MyEntity
        Get
            For Each ent As MyEntity In Me
                If ent.TreeNode Is Node Then
                    Return ent
                End If
            Next

            Return Nothing
        End Get
    End Property

End Class


Friend Class MyEntity
    Private m_Entity As Object
    Private m_AttributeSets As MyAttributeSets
    Private m_Node As TreeNode
    Private m_New As Boolean

    Public Property IsNew() As Boolean
        Get
            Return m_New
        End Get
        Set(ByVal value As Boolean)
            If Not g_IgnoreDuringLoad Then
                m_New = value
            End If
        End Set
    End Property

    Public ReadOnly Property IsEdited() As Boolean
        Get
            For Each attribSet As MyAttributeSet In AttributeSets
                If attribSet.IsEdited Then
                    Return True
                End If
            Next

            Return False
        End Get
    End Property

    Public Property Entity() As Object
        Get
            Return m_Entity
        End Get
        Set(ByVal value As Object)
            m_Entity = value
        End Set
    End Property

    Public ReadOnly Property AttributeSets() As MyAttributeSets
        Get
            Return m_AttributeSets
        End Get
    End Property

    Public Property TreeNode() As TreeNode
        Get
            Return m_Node
        End Get
        Set(ByVal value As TreeNode)
            m_Node = value
        End Set
    End Property

    Public Sub SelectEntity()
        If Not m_Node Is Nothing Then
            m_Node.TreeView.SelectedNode = m_Node
            m_Node.Expand()
        End If
    End Sub

    Public Sub New()
        m_AttributeSets = New MyAttributeSets
        m_AttributeSets.Parent = Me
        m_Entity = Nothing
    End Sub
End Class


Friend Class MyAttributeSets
    Inherits CollectionBase
    Private m_parent As MyEntity
    Private m_isEdited As Boolean = False
    Private m_New As Boolean

    Public Property IsNew() As Boolean
        Get
            Return m_New
        End Get
        Set(ByVal value As Boolean)
            If Not g_IgnoreDuringLoad Then
                m_New = value
            End If
        End Set
    End Property

    Public Function Add(ByVal Name As String) As MyAttributeSet
        Try
            ' Create a new object
            Dim newMember As MyAttributeSet
            newMember = New MyAttributeSet
            newMember.Parent = Me

            ' Set the properties passed into the method
            newMember.Name = Name

            newMember.IsNew = True

            ' Add the object to the collection.
            MyBase.List.Add(newMember)

            ' Return the object created
            Add = newMember
            newMember = Nothing
        Catch ex As Exception
            ' Failure.
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property IsEdited() As Boolean
        Get
            If m_isEdited Then
                Return True
            Else
                For Each attribSet As MyAttribute In Me
                    If attribSet.IsEdited Then
                        Return True
                    End If
                Next
            End If

            Return False
        End Get
    End Property

    Public Property Parent() As MyEntity
        Get
            Return m_parent
        End Get

        Set(ByVal value As MyEntity)
            m_parent = value
        End Set
    End Property


    Default Public ReadOnly Property Item(ByVal indexKey As Integer) As MyAttributeSet
        Get
            Return CType(MyBase.List.Item(indexKey - 1), MyAttributeSet)
        End Get
    End Property
End Class


Friend Class MyAttributeSet
    Inherits CollectionBase

    Private m_originalName As String = ""
    Private m_name As String = ""
    Private m_Parent As MyAttributeSets
    Private m_Node As TreeNode
    Private m_isEdited As Boolean = False
    Private m_Deleted As Boolean = False
    Private m_New As Boolean = False

    Public Property IsNew() As Boolean
        Get
            Return m_New
        End Get
        Set(ByVal value As Boolean)
            If Not g_IgnoreDuringLoad Then
                m_New = value
            End If
        End Set
    End Property

    Public Property IsDeleted() As Boolean
        Get
            Return m_Deleted
        End Get
        Set(ByVal value As Boolean)
            If Not g_IgnoreDuringLoad Then
                m_Deleted = value
            End If
        End Set
    End Property

    Public Function Add(ByVal Name As String, ByVal AttribType As Inventor.ValueTypeEnum, ByVal Value As Object) As MyAttribute
        Try
            ' Create a new object
            Dim newMember As MyAttribute
            newMember = New MyAttribute

            ' Check that the name is unique.
            For Each attrib As MyAttribute In Me
                If attrib.Name.ToUpper = Name.ToUpper Then
                    MsgBox("The specified parameter name already exists.")
                    Return Nothing
                End If
            Next
            ' Set the properties passed into the method

            newMember.Name = Name
            newMember.AttributeType = AttribType
            newMember.Value = Value
            newMember.IsNew = True
            newMember.Parent = Me

            ' Add the object to the collection.
            MyBase.List.Add(newMember)

            ' Return the object created
            Add = newMember
            newMember = Nothing
        Catch ex As Exception
            ' Failure.
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property IsEdited() As Boolean
        Get
            If m_isEdited Then
                Return True
            Else
                For Each attrib As MyAttribute In Me
                    If attrib.IsEdited Or attrib.IsNew Or attrib.IsDeleted Then
                        Return True
                    End If
                Next
            End If

            Return False
        End Get
    End Property

    Public Property Parent() As MyAttributeSets
        Get
            Return m_Parent
        End Get
        Set(ByVal value As MyAttributeSets)
            m_Parent = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return m_name
        End Get
        Set(ByVal value As String)
            If m_name = "" Then
                m_originalName = value
            End If

            m_name = value

            If Not g_IgnoreDuringLoad Then
                m_isEdited = True
            End If
        End Set
    End Property

    Public ReadOnly Property OriginalName As String
        Get
            Return m_originalName
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal indexKey As Integer) As MyAttribute
        Get
            Return CType(MyBase.List.Item(indexKey - 1), MyAttribute)
        End Get
    End Property

    Public Property TreeNode() As TreeNode
        Get
            Return m_Node
        End Get
        Set(ByVal value As TreeNode)
            m_Node = value
        End Set
    End Property
End Class


Friend Class MyAttribute
    Private m_Name As String = ""
    Private m_OriginalName As String
    Private m_Type As Inventor.ValueTypeEnum
    Private m_Value As Object
    Private m_parent As MyAttributeSet
    Private m_Node As TreeNode
    Private m_IsEdited As Boolean = False
    Private m_Deleted As Boolean = False
    Private m_New As Boolean = False

    Public Property IsNew() As Boolean
        Get
            Return m_New
        End Get
        Set(ByVal value As Boolean)
            If Not g_IgnoreDuringLoad Then
                m_New = value
            End If
        End Set
    End Property

    Public Property IsDeleted() As Boolean
        Get
            Return m_Deleted
        End Get
        Set(ByVal value As Boolean)
            If Not g_IgnoreDuringLoad Then
                m_Deleted = value
            End If
        End Set
    End Property

    Public Property TreeNode() As TreeNode
        Get
            Return m_Node
        End Get
        Set(ByVal value As TreeNode)
            m_Node = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return m_name
        End Get
        Set(ByVal value As String)
            If m_name = "" Then
                m_originalName = value
            End If

            m_name = value

            If Not g_IgnoreDuringLoad Then
                m_isEdited = True
            End If
        End Set
    End Property

    Public ReadOnly Property OriginalName As String
        Get
            Return m_originalName
        End Get
    End Property

    Public Property AttributeType() As Inventor.ValueTypeEnum
        Get
            Return m_Type
        End Get
        Set(ByVal value As Inventor.ValueTypeEnum)
            m_Type = value

            If Not g_IgnoreDuringLoad Then
                m_IsEdited = True
            End If
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return m_Value
        End Get
        Set(ByVal value As Object)
            m_Value = value

            If Not g_IgnoreDuringLoad Then
                m_IsEdited = True
            End If
        End Set
    End Property

    Public ReadOnly Property ValueAsString() As String
        Get
            If AttributeType = Inventor.ValueTypeEnum.kByteArrayType Then
                Return "(Byte Array)"
            Else
                Return m_Value.ToString
            End If
        End Get
    End Property

    Public ReadOnly Property ValueAsFullString() As String
        Get
            Dim arrayString As String = ""
            If AttributeType = Inventor.ValueTypeEnum.kByteArrayType Then
                For Each byteValue As Byte In m_Value
                    If arrayString = "" Then
                        arrayString = byteValue
                    Else
                        arrayString = arrayString & "," & byteValue
                    End If
                Next
                Return arrayString
            Else
                Return m_Value.ToString
            End If
        End Get
    End Property

    Public ReadOnly Property IsEdited() As Boolean
        Get
            Return m_IsEdited
        End Get
    End Property

    Public Property Parent() As MyAttributeSet
        Get
            Return m_parent
        End Get
        Set(ByVal value As MyAttributeSet)
            m_parent = value
        End Set
    End Property
End Class


Public Module Globals
    Public g_IgnoreDuringLoad As Boolean = False
End Module