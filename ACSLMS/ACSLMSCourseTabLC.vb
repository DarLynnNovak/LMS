
'Option Explicit On
'Option Strict On

Imports Aptify.Framework.Application
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.ExceptionManagement
Imports System.Windows.Forms
'Imports Aptify.Framework.WindowsControls.StepList
Imports Aptify.Framework.WindowsControls
Imports System.Text
Imports System.Globalization
Imports System
Imports System.IO
Imports System.Object
Imports System.Data.OleDb
Imports System.Data
Imports Microsoft.VisualBasic
Imports Microsoft.Office.Interop
Imports Aptify.Framework.BusinessLogic

Public Class ACSLMSCourseTabLC
    Inherits FormTemplateLayout
    Private m_oApp As New AptifyApplication
    Private m_oDA As New DataAction
    Private m_oProps As New AptifyProperties

    Dim CourseCreationGE As AptifyGenericEntityBase
    Dim CourseGE As AptifyGenericEntityBase
    Dim result As String = "Failed"
    Private WithEvents CourseCreateButton As AptifyActiveButton
    Private WithEvents CourseUpdateButton As AptifyActiveButton
    Private WithEvents CourseIdLB As AptifyLinkBox

    Dim da As New DataAction
    Dim ID As Long
    Dim Status As Long
    Dim CourseCreationCourseId As Long
    Dim InstructorId As Long
    Dim SchoolId As Long
    Dim CourseCreatorAppGE As AptifyGenericEntityBase


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
            If CourseCreateButton Is Nothing OrElse CourseCreateButton.IsDisposed = True Then
                CourseCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.Active Button.1"), AptifyActiveButton)
            End If
            If CourseUpdateButton Is Nothing OrElse CourseUpdateButton.IsDisposed = True Then
                CourseUpdateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.Active Button.2"), AptifyActiveButton)
            End If
            If CourseIdLB Is Nothing OrElse CourseIdLB.IsDisposed = True Then
                CourseIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseId"), AptifyLinkBox)
            End If

            If CourseIdLB.Value > 0 Then

                CourseCreateButton.Visible = False
                CourseUpdateButton.Visible = True
            Else

                CourseCreateButton.Visible = True
                CourseUpdateButton.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub


    Private Sub CourseCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CourseCreateButton.Click
        Try

            CourseCreationGE = m_oApp.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreationGE.SetValue("CourseCreationStatus", 2)
            CourseCreationGE.Save(True)

            'm_oAppObj.DisplayEntityRecord("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
        Catch ex As Exception

        End Try

    End Sub


    Private Sub CourseUpdateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CourseUpdateButton.Click
        Try
            CourseCreationGE = m_oApp.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            CourseCreationGE.SetValue("CourseCreationStatus", 12)
            CourseCreationGE.Save(True)
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try

    End Sub


End Class
