
Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.BusinessLogic.GenericEntity

Public Class ACSLMSEventUserTab
    Inherits FormTemplateLayout

    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Dim CourseCreatorAppGE As AptifyGenericEntityBase

    Dim ID As Integer


    Private WithEvents CourseSponsorTB As AptifyTextBox
    Private WithEvents CourseFormatDCB As AptifyDataComboBox
    Private WithEvents CMECreditAmountTB As AptifyTextBox
    Private WithEvents CECreditAmountTB As AptifyTextBox
    Private WithEvents CACreditAmountTB As AptifyTextBox
    Private WithEvents SACreditsAmountTB As AptifyTextBox
    Private WithEvents SACECreditsAmountTB As AptifyTextBox
    Private WithEvents IsRegMandate As AptifyCheckBox

    Private WithEvents RequestedEventStartDateTB As AptifyTextBox
    Private WithEvents RequestedEventEndDateTB As AptifyTextBox
    Private WithEvents RequestedClaimingExpDateTB As AptifyTextBox
    Private WithEvents RequestedRegMandateTypeLB As AptifyLinkBox
    Private WithEvents RequestedRegMandateAmount As AptifyTextBox
    Private WithEvents CMECertificateLB As AptifyLinkBox
    Private WithEvents CECertificateLB As AptifyLinkBox
    Private WithEvents CACertificateLB As AptifyLinkBox
    Dim courseCreatorGroupSQL As String
    Dim courseOwnerSQL As String
    Dim CourseOwnerId As String
    Dim CourseOwner As Integer
    Dim UserCreatedId As String
    Dim CourseOwnerPersonId As Integer
    Private WithEvents IsCME As AptifyCheckBox
    Private WithEvents IsCE As AptifyCheckBox
    Private WithEvents IsCOP As AptifyCheckBox

    Dim COPCertId As Integer = 26
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

            If CourseSponsorTB Is Nothing OrElse CourseSponsorTB.IsDisposed = True Then
                CourseSponsorTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CourseSponsoringAssociation"), AptifyTextBox)
            End If
            If CourseFormatDCB Is Nothing OrElse CourseFormatDCB.IsDisposed = True Then
                CourseFormatDCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CourseFormat"), AptifyDataComboBox)
            End If
            If RequestedEventStartDateTB Is Nothing OrElse RequestedEventStartDateTB.IsDisposed = True Then
                RequestedEventStartDateTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.RequestEventStartDate"), AptifyTextBox)
            End If
            If RequestedEventEndDateTB Is Nothing OrElse RequestedEventEndDateTB.IsDisposed = True Then
                RequestedEventEndDateTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.RequestEventEndDate"), AptifyTextBox)
            End If
            If RequestedClaimingExpDateTB Is Nothing OrElse RequestedClaimingExpDateTB.IsDisposed = True Then
                RequestedClaimingExpDateTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.RequestedClaimingExpDate"), AptifyTextBox)
            End If
            If CMECertificateLB Is Nothing OrElse CMECertificateLB.IsDisposed = True Then
                CMECertificateLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CMECertificate"), AptifyLinkBox)
            End If
            If CMECreditAmountTB Is Nothing OrElse CMECreditAmountTB.IsDisposed = True Then
                CMECreditAmountTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CMECreditAmount"), AptifyTextBox)
            End If

            If CECertificateLB Is Nothing OrElse CECertificateLB.IsDisposed = True Then
                CECertificateLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CECertificateId"), AptifyLinkBox)
            End If
            If CECreditAmountTB Is Nothing OrElse CECreditAmountTB.IsDisposed = True Then
                CECreditAmountTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CECreditAmount"), AptifyTextBox)
            End If

            If CACertificateLB Is Nothing OrElse CACertificateLB.IsDisposed = True Then
                CACertificateLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CACertificate"), AptifyLinkBox)
            End If
            If CACreditAmountTB Is Nothing OrElse CACreditAmountTB.IsDisposed = True Then
                CACreditAmountTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.CACreditAmount"), AptifyTextBox)
            End If
            If SACreditsAmountTB Is Nothing OrElse SACreditsAmountTB.IsDisposed = True Then
                SACreditsAmountTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.SACreditAmount"), AptifyTextBox)
            End If
            If SACECreditsAmountTB Is Nothing OrElse SACECreditsAmountTB.IsDisposed = True Then
                SACECreditsAmountTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.SACECreditAmount"), AptifyTextBox)
            End If
            If RequestedRegMandateTypeLB Is Nothing OrElse RequestedRegMandateTypeLB.IsDisposed = True Then
                RequestedRegMandateTypeLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.RequestedRegMandateType"), AptifyLinkBox)
            End If

            If RequestedRegMandateAmount Is Nothing OrElse RequestedRegMandateAmount.IsDisposed = True Then
                RequestedRegMandateAmount = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.RequestedRegMandateAmount"), AptifyTextBox)
            End If


            If IsCME Is Nothing OrElse IsCME.IsDisposed = True Then
                IsCME = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.IsCMEActivity"), AptifyCheckBox)
            End If

            If IsCE Is Nothing OrElse IsCE.IsDisposed = True Then
                IsCE = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.IsCEActivity"), AptifyCheckBox)
            End If

            If IsCOP Is Nothing OrElse IsCE.IsDisposed = True Then
                IsCOP = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event.IsCAActivity"), AptifyCheckBox)
            End If
            'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
            RequestedEventStartDateTB.Value = FormTemplateContext.GE.GetValue("CourseStartDate")
            RequestedEventEndDateTB.Value = FormTemplateContext.GE.GetValue("CourseEndDate")
            If IsCME.Value = True Then
                CMECertificateLB.Visible = True
                CMECreditAmountTB.Visible = True
                SACreditsAmountTB.Visible = True
            Else
                CMECertificateLB.Visible = False
                CMECreditAmountTB.Visible = False
                SACreditsAmountTB.Visible = True


            End If

            If IsCOP.Value = True Then
                CMECertificateLB.Value = COPCertId
                CECertificateLB.Value = COPCertId
            End If

            If IsCE.Value = 1 Then
                CECertificateLB.Visible = True
                CECreditAmountTB.Visible = True

            Else
                CECertificateLB.Visible = False
                CECreditAmountTB.Visible = False

            End If
            ID = Me.FormTemplateContext.GE.RecordID
            'RequestedClaimingExpDateTB.Value = CourseCreatorAppGE.GetValue("CourseEndDate") 
            CheckCourseOwnerLB()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub
    Private Sub isCME_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles IsCME.ValueChanged
        'Dim lTypeID As Long = -1
        If (NewValue) = True Then

            CMECertificateLB.Visible = True
            CMECreditAmountTB.Visible = True
            SACreditsAmountTB.Visible = True

        Else
            CMECertificateLB.Visible = False
            CMECreditAmountTB.Visible = False
            SACreditsAmountTB.Visible = False
        End If
    End Sub

    Private Sub isCE_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles IsCE.ValueChanged
        'Dim lTypeID As Long = -1
        If (NewValue) = True Then

            CECertificateLB.Visible = True
            CECreditAmountTB.Visible = True

        Else
            CECertificateLB.Visible = False
            CECreditAmountTB.Visible = False
        End If
    End Sub
    Private Sub CourseFormatDCB_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles CourseFormatDCB.ValueChanged

        If OldValue <> NewValue AndAlso NewValue > 0 Then


            'CCAGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
            'CCAGE.SetValue("EventType", NewValue)
            'CCAGE.Save(True)

            FormTemplateContext.GE.SetValue("EventType", NewValue)
            FormTemplateContext.GE.Save()


        End If
    End Sub
    Public Sub CheckCourseOwnerLB()
        Dim da As New DataAction
        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim FTP As Integer = FormTemplateID


        Dim userid As Long = m_oAppObj.UserCredentials.AptifyUserID
        Dim Uservalue As String = m_oAppObj.UserCredentials.GetUserRelatedRecordID(userid)

        Dim thisEntityID As Integer

        If Me.DataAction.UserCredentials.Server.ToLower = "aptify" Then
            'production
            thisEntityID = 2788
        End If
        If Me.DataAction.UserCredentials.Server.ToLower = "stagingaptify2" Then
            'staging
            thisEntityID = 2776
        End If

        If Me.DataAction.UserCredentials.Server.ToLower = "testaptifydb" Then
            'staging 
            thisEntityID = 2863
        End If

        If Me.DataAction.UserCredentials.Server.ToLower = "testaptify610" Then
            'staging 
            thisEntityID = 2788
        End If

        courseCreatorGroupSQL = "select UserID from GroupMember where groupid = (select id FROM Groups where name = 'LMS Course Creators') and UserID = " & userid
        dt = m_oDA.GetDataTable(courseCreatorGroupSQL)

        UserCreatedId = "select * from vwEntityRecordHistory where EntityID = " & thisEntityID & " and Changes = 'Record created' and (LTRIM(RTRIM(WhoUpdated)) like '%' + (Select UserID from vwUsers where ID = " & userid & ") +  '%' or (ltrim(RTRIM(WhoUpdated)) = 'sa'))  and RecordID = " & Me.FormTemplateContext.GE.RecordID
        dt1 = m_oDA.GetDataTable(UserCreatedId)

        'courseOwnerSQL = "select e.linkedpersonid from vwUserEntityRelations uer join vwemployees e on e.id = uer.EntityRecordID join vwusers u on u.id = uer.userid where u.id = " & userid
        'CourseOwnerPersonId = CLng(da.ExecuteScalar(courseOwnerSQL))

        Dim courseSql As String = "select courseowner from vwacslmscoursecreatorapp where ID = " & Me.FormTemplateContext.GE.RecordID
        Dim coursesqlid As Integer = da.ExecuteScalar(courseSql)
        If userid <> 11 Then

            courseOwnerSQL = "select e.linkedpersonid from vwUserEntityRelations uer join vwemployees e on e.id = uer.EntityRecordID join vwusers u on u.id = uer.userid where u.id = " & userid
            CourseOwnerPersonId = CLng(da.ExecuteScalar(courseOwnerSQL))
        End If


        If Me.FormTemplateContext.GE.RecordID > 0 Then

            If Not FormTemplateContext.GE.Save(False) Then
                Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)

            Else

                FormTemplateContext.GE.Save(True)
                'CourseCreatorAppGE.CommitTransaction()
                'UpdateCourseCreator()
            End If
        End If

    End Sub

End Class
