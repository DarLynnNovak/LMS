﻿Imports Aptify.Framework.Application
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices

Imports System.Text.RegularExpressions



Public Class ACSLMSNightlyNotificationPC
    Implements IProcessComponent

    Private m_oApp As New AptifyApplication
    Private m_oProps As New AptifyProperties
    Private m_oGE As AptifyGenericEntityBase
    Dim CourseCreatorAppGE As AptifyGenericEntity
    Private m_oda As DataAction
    Private m_sResult As String = "SUCCESS"
    Dim EmailAddress
    Dim emailgroup As String
    Dim ccEmail As String
    Dim SubjectText As String
    Dim HTMLText As String
    Dim lMessageTemplateID As Integer
    Dim thisMessageTemplateId As Integer
    Dim lMessageSourceID As Integer

    Dim sql As String
    Dim result As String = "Failed"
    Dim applicantID As Long
    Dim sSQL1 As String
    Dim recordId As Integer
    Dim DescriptionText As String
    Dim PrevRun As DataTable
    Dim MRId As Integer
    Public Overridable ReadOnly Property DataAction() As DataAction
        Get
            If m_oda Is Nothing Then
                m_oda = New DataAction(m_oApp.UserCredentials)
            End If
            Return m_oda
        End Get
    End Property
    Public Sub Config(ByVal ApplicationObject As Aptify.Framework.Application.AptifyApplication) Implements Aptify.Framework.BusinessLogic.ProcessPipeline.IProcessComponent.Config
        Try
            m_oApp = ApplicationObject
            Me.m_oda = New Aptify.Framework.DataServices.DataAction(Me.m_oApp.UserCredentials)
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Public ReadOnly Property Properties() As Aptify.Framework.Application.AptifyProperties Implements Aptify.Framework.BusinessLogic.ProcessPipeline.IProcessComponent.Properties
        Get
            If m_oProps Is Nothing Then
                m_oProps = New Aptify.Framework.Application.AptifyProperties
            End If
            Return m_oProps
        End Get
    End Property

    Public Function Run() As String Implements IProcessComponent.Run
        'm_sResult = "SUCCESS"
        Dim m_sDatabase As String = "APTIFY"

        Dim da As New DataAction
        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim dt2 As DataTable
        Dim sql2 As String
        Dim PersonMailName As String
        Dim EventTemplateId As Integer
        Dim GLTemplateId As Integer
        Dim NavInclusionId As Integer
        Dim ICTemplateId As Integer
        Dim RequestCompleteTemplateId As Integer
        Dim Dept As String
        Dim theDate As Date = Now()


        'CourseCreatorAppGE = CType(Me.m_oProps.GetProperty("CourseCreatorAppGE"), AptifyGenericEntity)
        'recordId = CourseCreatorAppGE.RecordID
        If m_oda.UserCredentials.Server.ToLower = "aptify" Then
            EventTemplateId = 1585
            GLTemplateId = 1586
            NavInclusionId = 1587
            ICTemplateId = 1589
            RequestCompleteTemplateId = 1588
            'thisMessageTemplateId = 1391 
            lMessageSourceID = 118
        End If
        If m_oda.UserCredentials.Server.ToLower = "stagingaptify2" Then
            'staging
            ' thisMessageTemplateId = 1274
            EventTemplateId = 1292
            GLTemplateId = 1293
            NavInclusionId = 1294
            ICTemplateId = 1296
            RequestCompleteTemplateId = 1295
            lMessageSourceID = 122
        End If

        If m_oda.UserCredentials.Server.ToLower = "testaptifydb" Then
            'staging
            ' thisMessageTemplateId = 1266
            'lMessageSourceID = 118

        End If
        If m_oda.UserCredentials.Server.ToLower = "testaptify610" Then
            EventTemplateId = 1585
            GLTemplateId = 1586
            NavInclusionId = 1587
            ICTemplateId = 1589
            RequestCompleteTemplateId = 1588
            thisMessageTemplateId = 1475
            lMessageSourceID = 118
        End If
        Try


            emailgroup = ""

            'sql = "select * from acslmscoursecreatorapp where id = " & CourseCreatorAppGE.RecordID
            sql = "select * from acslmscoursecreatorapp where GLCreated = 1 and CourseIdCreated = 1 and SalesGL is not null and ProductIdCreated = 1 and convert(date, GLCreatedDate, 120) = convert(date, getdate(), 120)"

            dt = m_oda.GetDataTable(sql)

            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows

                    lMessageTemplateID = NavInclusionId

                    emailgroup = "jbodnar@facs.org"
                        applicantID = "3241471"
                    SubjectText = "LMS Request for Nav Inclusion" & dr.Item("ID")

                    HTMLText = "The course setup has been completed.&nbsp; Please include the following new product in the Nav:</p>
                        <p>
                        <br />ProductName<br />:" & CStr(dr.Item("ProductName")) & "<br></br>" &
                          "SalesGL:  " & CStr(dr.Item("SalesGL"))
                        PreviousRunCheck()
                    If PrevRun.Rows.Count = 0 Then
                        'SendEmail()
                        createMessageRun()
                    End If
                Next
            End If
            Return result
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            Return "FAILED"
        End Try

    End Function


    Public Sub SendEmail()
        Dim pfresult As ProcessFlowResult = Nothing
        Dim lProcessFlowID As Long = -1
        Dim boolSuccess As Boolean = False
        Dim procFlowSql As String

        Dim da As New DataAction
        procFlowSql = ("Select ID FROM aptify..vwProcessFlows WHERE Name='LMSCourseCreationMessage'")
        lProcessFlowID = CLng(da.ExecuteScalar(procFlowSql))
        If lProcessFlowID > 0 Then
            Dim context As New AptifyContext
            context.Properties.AddProperty("email", emailgroup)
            context.Properties.AddProperty("ccEmail", ccEmail)
            context.Properties.AddProperty("lMessageTemplateID", lMessageTemplateID)
            context.Properties.AddProperty("recordid", recordId)
            context.Properties.AddProperty("SubjectText", SubjectText)
            context.Properties.AddProperty("HTMLText", HTMLText)
            pfresult = ProcessFlowEngine.ExecuteProcessFlow(m_oApp, lProcessFlowID, context)  'execute the process flow 
            result = "SUCCESS"
            CreateContactLog()
        Else
            result = "NORECORD"
        End If

    End Sub

    Private Sub createMessageRun()
        Try
            Dim messageRunGe As AptifyGenericEntityBase
            messageRunGe = m_oApp.GetEntityObject("Message Runs", -1)
            With messageRunGe
                .SetValue("MessageSystemID", 6)
                .SetValue("MessageSourceID", lMessageSourceID)
                .SetValue("MessageTemplateID", lMessageTemplateID)
                .SetValue("ApprovalStatus", "Approved")
                .SetValue("Status", "Pending")
                .SetValue("ScheduledStartDate", Now())
                .SetValue("Priority", "Normal")
                .SetValue("ToType", "Static")
                .SetValue("ToValue", emailgroup)
                .SetValue("CCType", "Static")
                '.SetValue("Comments", PersonID)
                .SetValue("RecipientCount", 0)
                .SetValue("SourceType", "ID String")
                .SetValue("IDString", recordId)
                .SetValue("HTMLBody", HTMLText)
            End With



            If messageRunGe.IsDirty Then
                If Not messageRunGe.Save(False) Then
                    Throw New Exception("Problem Saving Course Record:" & messageRunGe.RecordID)
                    result = "Error"
                Else
                    messageRunGe.Save(True)
                    result = "Success"
                    'm_sResult = "SUCCESS"
                    MRId = messageRunGe.RecordID
                    CreateContactLog()
                End If

            End If


        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            'Return False
        End Try

    End Sub
    Private Function CreateContactLog()

        '  Private Function CreateContactLog(ByVal pCompaniesID As Long, ByVal pOrderID As Long, ByVal pMessageTemplateID As Long) As Boolean
        'Dim bResult As Boolean = False
        Try


            Dim contactLogGe As AptifyGenericEntityBase

            contactLogGe = m_oApp.GetEntityObject("Contact Log", -1)
            With contactLogGe
                .SetValue("Description", SubjectText)
                .SetValue("Details", HTMLText)
                .SetValue("CategoryID", 7)
                .SetValue("TypeID", 5)
                .SetValue("NextContactStatus", "Complete")
                .SetValue("DefaultPersonLinkID", applicantID)
            End With

            'If CourseCreatorAppGE.RecordID > 0 Then
            '    With contactLogGe.SubTypes("ContactLogLinks").Add()
            '        .SetValue("EntityID", m_oApp.GetEntityID("ACSLMSCourseCreatorApp"))
            '        .SetValue("AltID", CourseCreatorAppGE.RecordID)
            '    End With
            'End If

            If applicantID > 0 Then
                With contactLogGe.SubTypes("ContactLogLinks").Add()
                    .SetValue("EntityID", m_oApp.GetEntityID("Persons"))
                    .SetValue("AltID", applicantID)
                End With
            End If
            If MRId > 0 Then
                With contactLogGe.SubTypes("ContactLogLinks").Add()
                    .SetValue("EntityID", m_oApp.GetEntityID("Message Runs"))
                    .SetValue("AltID", MRId)
                End With
            End If
            If contactLogGe.IsDirty Then
                If Not contactLogGe.Save(False) Then
                    Throw New Exception("Problem Saving Course Record:" & contactLogGe.RecordID)
                    result = "Error"
                Else
                    contactLogGe.Save(True)
                    result = "Success"
                End If

            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            'Return False
        End Try

        'Return pfresult
    End Function
    Private Function CheckEmailforSemicolon(ByVal email As String) As String
        Dim lResult As String
        If Right(email, 1) = ";" Then
            lResult = Left(email, Len(email) - 1)
        Else
            lResult = email
        End If
        Return lResult
    End Function

    Private Function IsValidEmail(ByVal email As String) As Boolean
        Dim lResult As Boolean = True
        Try
            lResult = Regex.IsMatch(email, "^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" &
                      "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$", RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250))

        Catch ex As Exception
            lResult = False
        End Try

        Return lResult
    End Function
    Private Function PreviousRunCheck()

        Dim previousrun As String
        DescriptionText = SubjectText
        previousrun = "Select * From contactlog Where DefaultPersonLinkID = " & applicantID & " And Description = '" & DescriptionText & "'"
        PrevRun = Me.DataAction.GetDataTable(previousrun)

    End Function
End Class
