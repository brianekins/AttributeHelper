Imports System.Windows.Forms

<System.Runtime.InteropServices.ComVisible(False)> Public Class EditAttribute
    'Public Sub New()
    '    ' This call is required by the designer.
    '    InitializeComponent()

    '    ' Add any initialization after the InitializeComponent() call.

    '    ' Set "String" as the initial default type.
    '    cboAttributeType.Text = "String"
    'End Sub

    Public Property AttributeName() As String
        Get
            Return Me.txtAttributeName.Text
        End Get
        Set(ByVal value As String)
            Me.txtAttributeName.Text = value
        End Set
    End Property

    Public Property AttributeValue() As String
        Get
            Return Me.txtAttributeValue.Text
        End Get
        Set(ByVal value As String)
            Me.txtAttributeValue.Text = value
        End Set
    End Property

    Public Property AttributeType() As Inventor.ValueTypeEnum
        Get
            Select Case cboAttributeType.SelectedIndex
                Case 0
                    Return Inventor.ValueTypeEnum.kStringType
                Case 1
                    Return Inventor.ValueTypeEnum.kIntegerType
                Case 2
                    Return Inventor.ValueTypeEnum.kDoubleType
                Case 3
                    Return Inventor.ValueTypeEnum.kByteArrayType
            End Select
        End Get

        Set(ByVal value As Inventor.ValueTypeEnum)
            Select Case value
                Case Inventor.ValueTypeEnum.kStringType
                    cboAttributeType.Text = "String"
                Case Inventor.ValueTypeEnum.kIntegerType
                    cboAttributeType.Text = "Integer"
                Case Inventor.ValueTypeEnum.kDoubleType
                    cboAttributeType.Text = "Double"
                Case Inventor.ValueTypeEnum.kByteArrayType
                    cboAttributeType.Text = "Byte Array"
            End Select
        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public WriteOnly Property EditMode() As Boolean
        Set(ByVal value As Boolean)
            If value = True Then
                Me.Text = "Edit Attribute"
                Me.cboAttributeType.Enabled = False
            Else
                Me.Text = "Create Attribute"
                Me.txtAttributeName.Text = ""
                Me.txtAttributeValue.Text = ""
                Me.cboAttributeType.Text = "String"
                Me.cboAttributeType.Enabled = True
            End If
        End Set
    End Property

    Private Sub EditAttribute_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

    End Sub

    Private Sub EditAttribute_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If Me.Text = "Create Attribute" Then
            Me.txtAttributeName.Focus()
            Me.txtAttributeName.SelectAll()
        End If
    End Sub


    Private Sub EditAttribute_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub EditAttribute_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub cboAttributeType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAttributeType.SelectedIndexChanged

    End Sub
End Class
