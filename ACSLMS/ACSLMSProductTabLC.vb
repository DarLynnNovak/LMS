
'Option Explicit On
'Option Strict On

Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices

Public Class ACSLMSProductTabLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction

    Private WithEvents ProductCreateButton As AptifyActiveButton
    Private WithEvents ProductUpdateButton As AptifyActiveButton
    Private WithEvents GLCreateButton As AptifyActiveButton
    Private WithEvents ProductIdLB As AptifyLinkBox
    Private WithEvents SalesGLTB As AptifyTextBox


    Dim da As New DataAction



    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try

            Me.AutoScroll = True
            'Dim newFormTemplateid As Integer = 27183

            FindControls()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e) 
    End Sub
    Protected Overridable Sub FindControls()
        Try
            If ProductCreateButton Is Nothing OrElse ProductCreateButton.IsDisposed = True Then
                ProductCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.Active Button.1"), AptifyActiveButton)
            End If
            If ProductUpdateButton Is Nothing OrElse ProductUpdateButton.IsDisposed = True Then
                ProductUpdateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.Active Button.2"), AptifyActiveButton)
            End If

            If GLCreateButton Is Nothing OrElse GLCreateButton.IsDisposed = True Then
                GLCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.Active Button.3"), AptifyActiveButton)
            End If
            If ProductIdLB Is Nothing OrElse ProductIdLB.IsDisposed = True Then
                ProductIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.ProductId"), AptifyLinkBox)
            End If
            If SalesGLTB Is Nothing OrElse ProductIdLB.IsDisposed = True Then
                SalesGLTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.SalesGL"), AptifyTextBox)
            End If


            If ProductIdLB.Value > 0 Then

                ProductCreateButton.Visible = False
                ProductUpdateButton.Visible = True
                If CStr(SalesGLTB.Value) IsNot Nothing Then
                    GLCreateButton.Visible = True
                    GLCreateButton.Enabled = True
                Else
                    GLCreateButton.Visible = True
                    GLCreateButton.Enabled = False
                End If


            Else

                ProductCreateButton.Visible = True
                ProductUpdateButton.Visible = False
                GLCreateButton.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ProductCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProductCreateButton.Click
        Try
            Dim CourseCreatorAppGE As Aptify.Framework.BusinessLogic.GenericEntity.AptifyGenericEntity
            CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreatorAppGE.SetValue("CourseCreationStatus", 6)
            CourseCreatorAppGE.Save()
            m_oAppObj.DisplayEntityRecord("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ProductUpdateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProductUpdateButton.Click
        Try
            Dim CourseCreatorAppGE As Aptify.Framework.BusinessLogic.GenericEntity.AptifyGenericEntity
            CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreatorAppGE.SetValue("CourseCreationStatus", 16)
            CourseCreatorAppGE.Save()
            m_oAppObj.DisplayEntityRecord("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub GLCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GLCreateButton.Click
        Try
            Dim CourseCreatorAppGE As Aptify.Framework.BusinessLogic.GenericEntity.AptifyGenericEntity
            CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreatorAppGE.SetValue("CourseCreationStatus", 8)
            CourseCreatorAppGE.Save()
            m_oAppObj.DisplayEntityRecord("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
        Catch ex As Exception

        End Try

    End Sub

End Class

