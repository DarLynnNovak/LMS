'Option Explicit On
'Option Strict On

Imports System.Drawing
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.WindowsControls
Imports System.Windows.Forms
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.DataServices

Public Class ACSLMSSuvinaTabLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private WithEvents SuvinaTab As FormTemplateTab

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


            If SuvinaTab Is Nothing OrElse SuvinaTab.IsDisposed = True Then
                SuvinaTab = TryCast(Me.GetFormComponentByLayoutKey(Me, "ACSLMSCourseCreatorApp Form - Event: Step 2: Admin Tab"), FormTemplateTab)
            End If




        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
End Class
