Imports Aptify.Framework.Application
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices

Imports System.Text.RegularExpressions



Public Class ACSLMSNotificationsPCv2
    Implements IProcessComponent


    Private m_oApp As New AptifyApplication
    Private m_oProps As New AptifyProperties
    Private m_oGE As AptifyGenericEntityBase
    Dim CourseCreatorAppGE As AptifyGenericEntityBase
    Dim CCAGE As AptifyGenericEntityBase
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
    Dim RecordId As Integer
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
        m_sResult = "SUCCESS"
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
            'dt = m_oda.GetDataTable(sql)
            RecordId = CInt(m_oProps.GetProperty("RecordId"))
            If CInt(RecordId) > 0 Then

                CourseCreatorAppGE = m_oApp.GetEntityObject("acslmscoursecreatorapp", RecordId)

            Else

                CourseCreatorAppGE = CType(m_oProps.GetProperty("CourseCreatorAppGE"), AptifyGenericEntityBase)
                RecordId = CInt(CourseCreatorAppGE.RecordID)

            End If

            Dept = CourseCreatorAppGE.GetValue("ContactDepartment")

            sql2 = "select acsmailname from vwpersons (nolock) where id = " & CInt(CourseCreatorAppGE.GetValue("courseowner"))
            PersonMailName = m_oda.ExecuteScalar(sql2)

            If (CourseCreatorAppGE.GetValue("CourseIdCreated") = True Or CourseCreatorAppGE.GetValue("CourseIdCreated") = 1) Then

                lMessageTemplateID = EventTemplateId

                emailgroup = "ssallan@facs.org"
                applicantID = "3255965"
                SubjectText = "LMS Request for Event for CCA Record " & CourseCreatorAppGE.RecordID
                HTMLText = "A new course has been created on: " & CourseCreatorAppGE.GetValue("courseidcreateddate") & " for due date: " & CourseCreatorAppGE.GetValue("requestedduedate") & " by: " & PersonMailName & ".  Please go to " & m_oda.UserCredentials.Server.ToLower & " and create an Event For this course. ID: " & CourseCreatorAppGE.RecordID
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If

                emailgroup = "opetinaux@facs.org"
                applicantID = "3115385"
                SubjectText = "LMS Request for Event for CCA Record " & CourseCreatorAppGE.RecordID
                HTMLText = "A request for Event ID has been submitted for the CCA Request Id: " & CourseCreatorAppGE.RecordID & " in " & m_oda.UserCredentials.Server.ToLower & " on " & theDate
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If

            End If
            If (CourseCreatorAppGE.GetValue("CourseIdCreated") = True Or CourseCreatorAppGE.GetValue("CourseIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("ProductIdCreated") = True Or CourseCreatorAppGE.GetValue("ProductIdCreated") = 1) Then

                thisMessageTemplateId = GLTemplateId

                emailgroup = "dnovak@facs.org"
                applicantID = "3096875"
                SubjectText = "LMS Request for GL for CCA Record " & CourseCreatorAppGE.RecordID
                HTMLText = "The CCA Request Id: " & CourseCreatorAppGE.RecordID & " in " & m_oda.UserCredentials.Server.ToLower & " needs an GL.  Please create an GL for this request."
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If


            End If
            Dim cid = CourseCreatorAppGE.GetValue("CourseIdCreated")
            Dim pid = CourseCreatorAppGE.GetValue("ProductIdCreated")
            Dim glc = CourseCreatorAppGE.GetValue("GLCreated")
            If (CourseCreatorAppGE.GetValue("CourseIdCreated") = True Or CourseCreatorAppGE.GetValue("CourseIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("ProductIdCreated") = True Or CourseCreatorAppGE.GetValue("ProductIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("GLCreated") = True Or CourseCreatorAppGE.GetValue("GLCreated") = 1) Then

                lMessageTemplateID = NavInclusionId

                emailgroup = "jbodnar@facs.org"
                applicantID = "3241471"
                SubjectText = "LMS Request for Nav Inclusion" & CourseCreatorAppGE.RecordID

                HTMLText = "The course setup has been completed.&nbsp; Please include the following new product in the Nav:</p>
                        <p>
                        <br />ProductName<br />:" & CStr(CourseCreatorAppGE.GetValue("ProductName")) & "<br></br>" &
                          "SalesGL:  " & CStr(CourseCreatorAppGE.GetValue("SalesGL"))
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    'SendEmail()
                    createMessageRun()
                End If
            End If

            If (CourseCreatorAppGE.GetValue("CourseIdCreated") = True Or CourseCreatorAppGE.GetValue("CourseIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("ProductIdCreated") = True Or CourseCreatorAppGE.GetValue("ProductIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("GLCreated") = True Or CourseCreatorAppGE.GetValue("GLCreated") = 1) Then
                lMessageTemplateID = ICTemplateId

                emailgroup = "ahastings@facs.org"
                applicantID = "3241507"
                SubjectText = "LMS Request for IC Check " & CourseCreatorAppGE.RecordID
                HTMLText = "A Request needs a IC Check the aptify instance: " & m_oda.UserCredentials.Server.ToLower & "The LMS is:  ProductName<br />:" & CStr(CourseCreatorAppGE.GetValue("ProductName")) & " The id is: " & CourseCreatorAppGE.RecordID & "  Please Do IC Check For this request."
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If
            End If

            If (CourseCreatorAppGE.GetValue("CourseIdCreated") = True Or CourseCreatorAppGE.GetValue("CourseIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("EventIdCreated") = True Or CourseCreatorAppGE.GetValue("EventIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("ProductIdCreated") = True Or CourseCreatorAppGE.GetValue("ProductIdCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("GLCreated") = True Or CourseCreatorAppGE.GetValue("GLCreated") = 1) AndAlso (CourseCreatorAppGE.GetValue("CourseSetupComplete") = True Or CourseCreatorAppGE.GetValue("CourseSetupComplete") = 1) Then
                lMessageTemplateID = RequestCompleteTemplateId
                emailgroup = "mfield@facs.org"
                applicantID = "3267257"
                SubjectText = "EthosCE Request Complete for CCA Record " & CourseCreatorAppGE.RecordID
                'SubjectText = "EthosCE Request Complete For " & Dept
                HTMLText = "A New LMS course setup has been complete. The ID For this request Is: " & CourseCreatorAppGE.RecordID & " This request was created in the " & m_oda.UserCredentials.Server.ToLower & " instance of Aptify."
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If

                'emailgroup = "opetinaux@facs.org"
                'applicantID = "3115385"


                'SubjectText = "EthosCE Request Complete for CCA Record " & CourseCreatorAppGE.RecordID
                'HTMLText = "A new LMS course setup has been complete. The ID for this request is: " & CourseCreatorAppGE.RecordID & " This request was created in the " & m_oda.UserCredentials.Server.ToLower & " instance of Aptify."
                'PreviousRunCheck()
                'If PrevRun.Rows.Count = 0 Then
                '    'SendEmail()
                '    createMessageRun()
                'End If

                'emailgroup = "ajames@facs.org"
                'applicantID = "3096875"

                'SubjectText = "EthosCE Request Complete for CCA Record " & CourseCreatorAppGE.RecordID
                'HTMLText = "A new LMS course setup has been complete. The ID for this request is: " & CourseCreatorAppGE.RecordID & " This request was created in the " & m_oda.UserCredentials.Server.ToLower & " instance of Aptify."
                'PreviousRunCheck()
                'If PrevRun.Rows.Count = 0 Then
                '    'SendEmail()
                '    createMessageRun()
                'End If

                emailgroup = "sratsavong@facs.org"
                applicantID = "88175987"

                SubjectText = "EthosCE Request Complete for CCA Record " & CourseCreatorAppGE.RecordID
                HTMLText = "A new LMS course setup has been complete. The ID for this request is: " & CourseCreatorAppGE.RecordID & " This request was created in the " & m_oda.UserCredentials.Server.ToLower & " instance of Aptify."
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If

                emailgroup = CourseCreatorAppGE.GetValue("ContactEmail")
                applicantID = CourseCreatorAppGE.GetValue("CourseOwner")

                SubjectText = "EthosCE Request Complete for CCA Record " & CourseCreatorAppGE.RecordID
                HTMLText = "A new LMS course setup has been complete. The ID for this request is: " & CourseCreatorAppGE.RecordID & " This request was created in the " & m_oda.UserCredentials.Server.ToLower & " instance of Aptify."
                PreviousRunCheck()
                If PrevRun.Rows.Count = 0 Then
                    createMessageRun()
                End If
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            Return "FAILED"
        End Try

    End Function


    Private Sub createMessageRun()
        Dim bResult As Boolean = False
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
                .SetValue("RecipientCount", 0)
                .SetValue("SourceType", "ID String")
                .SetValue("IDString", recordId)
                .SetValue("Subject", SubjectText)
                .SetValue("HTMLBody", HTMLText)
            End With



            If messageRunGe.IsDirty Then
                If Not messageRunGe.Save(False) Then
                    Throw New Exception("Problem Saving Course Record:" & messageRunGe.RecordID)
                    result = "Error"
                Else
                    messageRunGe.Save(True)
                    result = "Success"
                    m_sResult = "SUCCESS"
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
        Dim bResult As Boolean = False
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

            If CourseCreatorAppGE.RecordID > 0 Then
                With contactLogGe.SubTypes("ContactLogLinks").Add()
                    .SetValue("EntityID", m_oApp.GetEntityID("ACSLMSCourseCreatorApp"))
                    .SetValue("AltID", CourseCreatorAppGE.RecordID)
                End With
            End If

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
