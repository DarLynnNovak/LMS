
'Option Explicit On
'Option Strict On

Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices

Public Class ACSLMSEventTabLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction

    Private WithEvents EventCreateButton As AptifyActiveButton
    Private WithEvents EventUpdateButton As AptifyActiveButton
    Private WithEvents EventIdLB As AptifyLinkBox

    Dim getEventIdSql As String
    Dim EventId As Integer
    Dim da As New DataAction



    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try

            'Me.AutoScroll = True
            'Dim newFormTemplateid As Integer = 27183

            FindControls()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e) 
    End Sub
    Protected Overridable Sub FindControls()
        Try
            If EventCreateButton Is Nothing OrElse EventCreateButton.IsDisposed = True Then
                EventCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.Active Button.1"), AptifyActiveButton)
            End If
            If EventUpdateButton Is Nothing OrElse EventUpdateButton.IsDisposed = True Then
                EventUpdateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.Active Button.2"), AptifyActiveButton)
            End If

            If EventIdLB Is Nothing OrElse EventIdLB.IsDisposed = True Then
                EventIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.EventId"), AptifyLinkBox)
            End If


            If EventIdLB.Value > 0 Then

                EventCreateButton.Visible = False
                EventUpdateButton.Visible = True
            Else

                EventCreateButton.Visible = True
                EventUpdateButton.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub EventCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EventCreateButton.Click
        Try
            Dim CourseCreatorAppGE As Aptify.Framework.BusinessLogic.GenericEntity.AptifyGenericEntity
            CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreatorAppGE.SetValue("CourseCreationStatus", 4)
            CourseCreatorAppGE.Save()
            m_oAppObj.DisplayEntityRecord("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
        Catch ex As Exception

        End Try

    End Sub


    Private Sub EventUpdateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EventUpdateButton.Click
        Try
            Dim CourseCreatorAppGE As Aptify.Framework.BusinessLogic.GenericEntity.AptifyGenericEntity
            CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreatorAppGE.SetValue("CourseCreationStatus", 14)
            CourseCreatorAppGE.Save()
            m_oAppObj.DisplayEntityRecord("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
        Catch ex As Exception

        End Try

    End Sub
End Class

