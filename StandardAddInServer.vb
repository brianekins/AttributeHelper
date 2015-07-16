Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32


Namespace AttributeHelper
    <ProgIdAttribute("AttributeHelper.StandardAddInServer"), _
    GuidAttribute("9c93ff52-e2fa-4ec4-8c4a-5747ad0bafef")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Private WithEvents m_attributeButtonDef As ButtonDefinition
        Private WithEvents m_UIEvents As UserInterfaceEvents

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Get the button images.
            Dim smallIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Icon16)
            Dim largeIcon As IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.Icon32)

            ' Create the button definition.
            m_attributeButtonDef = g_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition("Attribute" & vbCr & "Helper", "ekinsAttributeHelper", CommandTypesEnum.kNonShapeEditCmdType, AddInGuid(Me.GetType), "View, edit, and create attributes.", "Attribute Helper", smallIcon, largeIcon)

            'If firstTime Then
            AddToRibbon()
            'End If
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ' Release objects.
            Marshal.FinalReleaseComObject(g_inventorApplication)
            g_inventorApplication = Nothing

            If Not m_attributeButtonDef Is Nothing Then
                Marshal.FinalReleaseComObject(m_attributeButtonDef)
            End If

            If Not m_UIEvents Is Nothing Then
                Marshal.FinalReleaseComObject(m_UIEvents)
                m_UIEvents = Nothing
            End If

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub
#End Region

#Region "Ribbon related Members"
        Private Sub m_UIEvents_OnResetRibbonInterface(ByVal Context As Inventor.NameValueMap) Handles m_UIEvents.OnResetRibbonInterface
            AddToRibbon()
        End Sub


        Public Sub AddToRibbon()
            Try
                ' Create a panel on the Tools tab in all ribbons except zero doc and unknown.
                For Each ribbon As Inventor.Ribbon In g_inventorApplication.UserInterfaceManager.Ribbons
                    Select Case ribbon.InternalName
                        Case "Part", "Assembly", "Drawing", "Presentation", "iFeatures"
                            Dim tab As RibbonTab = ribbon.RibbonTabs.Item("id_TabTools")

                            Dim newPanel As RibbonPanel = tab.RibbonPanels.Add("Attributes", "ekinsAttributes", AddInGuid(Me.GetType))

                            Call newPanel.CommandControls.AddButton(m_attributeButtonDef, True, True)
                    End Select
                Next
            Catch ex As Exception
                MsgBox("Unexpected error adding Attribute Manager to the ribbon.")
            End Try
        End Sub
#End Region

        Private Sub m_attributeButtonDef_OnExecute(ByVal Context As Inventor.NameValueMap) Handles m_attributeButtonDef.OnExecute
            If Not g_inventorApplication.ActiveEditDocument Is Nothing Then
                Dim m_dialog As AttributeHelperDialog
                m_dialog = New AttributeHelperDialog
                m_dialog.InventorDocument = g_inventorApplication.ActiveEditDocument
                m_dialog.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
            End If
        End Sub


        ' This property uses reflection to get the value for the GuidAttribute attached to the class.
        Public Shared ReadOnly Property AddInGuid(ByVal t As Type) As String
            Get
                Dim guid As String = ""
                Try
                    Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
                    Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
                    guid = "{" + guidAttribute.Value.ToString() + "}"
                Finally
                    AddInGuid = guid
                End Try
            End Get
        End Property

    End Class

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(m_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

    Public Module Globals
        ' Inventor application object.
        Public g_inventorApplication As Inventor.Application

        Public NotInheritable Class PictureDispConverter
            <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
            Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
            End Function

            Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

            Private NotInheritable Class PICTDESC
                Private Sub New()
                End Sub

                'Picture Types
                Public Const PICTYPE_UNINITIALIZED As Short = -1
                Public Const PICTYPE_NONE As Short = 0
                Public Const PICTYPE_BITMAP As Short = 1
                Public Const PICTYPE_METAFILE As Short = 2
                Public Const PICTYPE_ICON As Short = 3
                Public Const PICTYPE_ENHMETAFILE As Short = 4

                <StructLayout(LayoutKind.Sequential)> _
                Public Class Icon
                    Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                    Friend picType As Integer = PICTDESC.PICTYPE_ICON
                    Friend hicon As IntPtr = IntPtr.Zero
                    Friend unused1 As Integer
                    Friend unused2 As Integer

                    Friend Sub New(ByVal icon As System.Drawing.Icon)
                        Me.hicon = icon.ToBitmap().GetHicon()
                    End Sub
                End Class

                <StructLayout(LayoutKind.Sequential)> _
                Public Class Bitmap
                    Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                    Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                    Friend hbitmap As IntPtr = IntPtr.Zero
                    Friend hpal As IntPtr = IntPtr.Zero
                    Friend unused As Integer

                    Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                        Me.hbitmap = bitmap.GetHbitmap()
                    End Sub
                End Class
            End Class

            Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
                Dim pictIcon As New PICTDESC.Icon(icon)
                Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
            End Function

            Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
                Dim pictBmp As New PICTDESC.Bitmap(bmp)
                Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
            End Function
        End Class

    End Module
End Namespace

