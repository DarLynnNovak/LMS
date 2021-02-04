Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.Application
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.ExceptionManagement

Public Class ACSLMSBundledProductLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private bAdded As Boolean = False
    Private lGridID As Long = -1
    Dim userid As Long
    Dim courseCreatorGroupSQL As String
    Private WithEvents IsBundle As AptifyCheckBox
    Private WithEvents BundledProductTab As FormTemplateTab
    Private WithEvents CourseCreationTab As FormTemplateTab
    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try

            Me.AutoScroll = True

            FindControls()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e)
    End Sub
    Protected Overridable Sub FindControls()
        Try
            If IsBundle Is Nothing OrElse IsBundle.IsDisposed = True Then
                IsBundle = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product Info.IsBundledProduct"), AptifyCheckBox)
            End If
            If BundledProductTab Is Nothing OrElse BundledProductTab.IsDisposed Then
                BundledProductTab = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorAppBundled.Tabs"), FormTemplateTab)
            End If
            If CourseCreationTab Is Nothing OrElse CourseCreationTab.IsDisposed Then
                CourseCreationTab = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Tabs"), FormTemplateTab)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub IsBundle_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles IsBundle.ValueChanged
        'Dim lTypeID As Long = -1
        If (NewValue) = True Then
            If IsBundle.Value = 0 Then
                MsgBox("Made it")
                BundledProductTab.Show()
                CourseCreationTab.Hide()
            Else
                CourseCreationTab.Show()
                BundledProductTab.Hide()
            End If
        End If
    End Sub


    Private Sub InitializeComponent()
        Me.SuspendLayout()

        Me.Name = "ACSLMSBundledProductLC"
        Me.ResumeLayout(False)

    End Sub

End Class
