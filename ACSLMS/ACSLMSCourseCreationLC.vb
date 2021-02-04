'Option Explicit On
'Option Strict On

Imports System.Drawing
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.WindowsControls
Imports System.Windows.Forms
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.Application

Public Class ACSLMSCourseCreationLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private bAdded As Boolean = False
    Private lGridID As Long = -1
    Dim userid As Long
    Dim courseCreatorGroupSQL As String
    Dim courseOwnerSQL As String
    Dim CourseOwnerId As String
    Dim CourseOwner As Integer
    Dim CourseId As String
    Dim EventId As String
    Dim ProductId As String
    Dim GLId As String
    Dim UserCreatedId As String
    Dim CourseOwnerPersonId As Integer
    Private WithEvents CourseCreationTab As FormTemplateTab
    Private WithEvents CourseRequestTab As FormTemplateTab
    Private WithEvents CourseOwnerIdLinkbox As AptifyLinkBox
    Private WithEvents CourseCreationStatus As AptifyDataComboBox
    Private WithEvents CourseCreationId As AptifyLinkBox
    Private WithEvents CourseCreatorApp As ApplicationsForm
    Private WithEvents ContactDepartment As AptifyTextBox
    Private WithEvents CostCenter As AptifyDataComboBox
    Private WithEvents ContactEmail As AptifyTextBox
    Private WithEvents ContactPhone As AptifyTextBox
    Private WithEvents RequestDecription As AptifyTextBox
    Private WithEvents IsCMEActivity As AptifyCheckBox
    Private WithEvents CMECertificateId As AptifyLinkBox
    Private WithEvents CMECreditAmount As AptifyTextBox
    Private WithEvents IsCEActivity As AptifyCheckBox
    Private WithEvents CECertificateId As AptifyLinkBox
    Private WithEvents CECreditAmount As AptifyTextBox
    Private WithEvents IsCAActivity As AptifyCheckBox
    Private WithEvents CACertificateId As AptifyLinkBox
    Private WithEvents CACreditAmount As AptifyTextBox
    Private WithEvents EventStartDate As AptifyTextBox
    Private WithEvents EventEndDate As AptifyTextBox
    Private WithEvents CourseSponsoringAssociation As AptifyTextBox
    Private WithEvents CourseFormat As AptifyComboBox
    Private WithEvents EnrollmentType As AptifyComboBox
    Private WithEvents CourseAdminNotes As AptifyTextBox
    Private WithEvents RequestedDueDate As AptifyTextBox
    Private WithEvents CreditClaimingExpirationDate As AptifyTextBox
    Private WithEvents UserMessage As CultureLabel
    Private WithEvents RequestedCourseName As AptifyTextBox
    Private WithEvents CopyNameButton As FormComponent
    Private WithEvents CopyDescButton As FormComponent
    Private WithEvents isClonedCourse As AptifyCheckBox
    Private WithEvents EthosNodeId As AptifyTextBox
    Private WithEvents PriceTableLabel As CultureLabel
    Private WithEvents AdminEthosSetupComplete As AptifyCheckBox
    Private WithEvents IsBundle As AptifyCheckBox
    Private WithEvents BundledProductTab As FormTemplateTab

    Dim currentDate As DateTime = Now

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


            If CourseOwnerIdLinkbox Is Nothing OrElse CourseOwnerIdLinkbox.IsDisposed = True Then
                CourseOwnerIdLinkbox = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseOwnerId"), AptifyLinkBox)
            End If

            If CourseCreationId Is Nothing OrElse CourseCreationId.IsDisposed = True Then
                CourseCreationId = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseId"), AptifyLinkBox)
            End If

            If ContactDepartment Is Nothing OrElse ContactDepartment.IsDisposed = True Then
                ContactDepartment = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.ContactDepartment"), AptifyTextBox)
            End If

            If CostCenter Is Nothing OrElse CostCenter.IsDisposed Then
                CostCenter = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CostCenter"), AptifyDataComboBox)
            End If
            If ContactEmail Is Nothing OrElse ContactEmail.IsDisposed = True Then
                ContactEmail = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.ContactEmail"), AptifyTextBox)
            End If
            If ContactPhone Is Nothing OrElse ContactPhone.IsDisposed = True Then
                ContactPhone = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.ContactPhone"), AptifyTextBox)
            End If

            If RequestDecription Is Nothing OrElse RequestDecription.IsDisposed = True Then
                RequestDecription = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.RequestDescription"), AptifyTextBox)
            End If
            If IsCMEActivity Is Nothing OrElse IsCMEActivity.IsDisposed = True Then
                IsCMEActivity = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.IsCMEActivity"), AptifyCheckBox)
            End If
            If CMECertificateId Is Nothing OrElse CMECertificateId.IsDisposed = True Then
                CMECertificateId = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CMECertificateId"), AptifyLinkBox)
            End If
            If CMECreditAmount Is Nothing OrElse CMECreditAmount.IsDisposed = True Then
                CMECreditAmount = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CMECreditAmount"), AptifyTextBox)
            End If
            If IsCEActivity Is Nothing OrElse IsCEActivity.IsDisposed = True Then
                IsCEActivity = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.IsCEActivity"), AptifyCheckBox)
            End If
            If CECertificateId Is Nothing OrElse CECertificateId.IsDisposed = True Then
                CECertificateId = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CECertificateId"), AptifyLinkBox)
            End If
            If CECreditAmount Is Nothing OrElse CECreditAmount.IsDisposed = True Then
                CECreditAmount = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CECreditAmount"), AptifyTextBox)
            End If

            If IsCAActivity Is Nothing OrElse IsCAActivity.IsDisposed = True Then
                IsCAActivity = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.IsCAActivity"), AptifyCheckBox)
            End If
            If CACertificateId Is Nothing OrElse CACertificateId.IsDisposed = True Then
                CACertificateId = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CACertificateId"), AptifyLinkBox)
            End If
            If CACreditAmount Is Nothing OrElse CACreditAmount.IsDisposed = True Then
                CACreditAmount = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CACreditAmount"), AptifyTextBox)
            End If
            If EventStartDate Is Nothing OrElse EventStartDate.IsDisposed = True Then
                EventStartDate = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.EventStartDate"), AptifyTextBox)
            End If
            If EventEndDate Is Nothing OrElse EventEndDate.IsDisposed = True Then
                EventEndDate = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.EventEndDate"), AptifyTextBox)
            End If

            If CourseSponsoringAssociation Is Nothing OrElse CourseSponsoringAssociation.IsDisposed = True Then
                CourseSponsoringAssociation = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseSponsoringAssociation"), AptifyTextBox)
            End If

            If CourseFormat Is Nothing OrElse CourseFormat.IsDisposed = True Then
                CourseFormat = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseFormat"), AptifyComboBox)
            End If

            If EnrollmentType Is Nothing OrElse EnrollmentType.IsDisposed = True Then
                EnrollmentType = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.EnrollmentType"), AptifyComboBox)
            End If

            If CourseAdminNotes Is Nothing OrElse CourseAdminNotes.IsDisposed = True Then
                CourseAdminNotes = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseAdminNotes"), AptifyTextBox)
            End If
            If RequestedDueDate Is Nothing OrElse RequestedDueDate.IsDisposed = True Then
                RequestedDueDate = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.RequestedDueDate"), AptifyTextBox)
            End If
            If CreditClaimingExpirationDate Is Nothing OrElse CreditClaimingExpirationDate.IsDisposed = True Then
                CreditClaimingExpirationDate = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CreditClaimingExpirationDate"), AptifyTextBox)
            End If
            If CourseCreationTab Is Nothing OrElse CourseCreationTab.IsDisposed Then
                CourseCreationTab = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Tabs"), FormTemplateTab)
            End If

            If CourseCreationStatus Is Nothing OrElse CourseCreationStatus.IsDisposed Then
                CourseCreationStatus = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseCreatorStatus"), AptifyDataComboBox)
            End If
            If UserMessage Is Nothing OrElse UserMessage.IsDisposed Then
                UserMessage = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.Culture Label.1"), CultureLabel)
            End If
            If RequestedCourseName Is Nothing OrElse RequestedCourseName.IsDisposed = True Then
                RequestedCourseName = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.RequestedCourseName"), AptifyTextBox)
            End If
            If CopyDescButton Is Nothing OrElse CopyDescButton.IsDisposed = True Then
                CopyDescButton = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.RequestDescription"), FormComponent)
            End If
            If CopyNameButton Is Nothing OrElse CopyNameButton.IsDisposed = True Then
                CopyNameButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course Info.RequestedCourseName"), FormComponent)
            End If

            If isClonedCourse Is Nothing OrElse isClonedCourse.IsDisposed = True Then
                isClonedCourse = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.isclonedcourse"), AptifyCheckBox)
            End If
            If EthosNodeId Is Nothing OrElse EthosNodeId.IsDisposed = True Then
                EthosNodeId = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.EthosNodeId"), AptifyTextBox)
            End If
            If PriceTableLabel Is Nothing OrElse PriceTableLabel.IsDisposed Then
                PriceTableLabel = TryCast(Me.GetFormComponent(Me, "Product Info.Culture Label.1"), CultureLabel)
            End If
            If CourseRequestTab Is Nothing OrElse CourseRequestTab.IsDisposed Then
                CourseRequestTab = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp4Admins.Tabs"), FormTemplateTab)
            End If

            If AdminEthosSetupComplete Is Nothing OrElse AdminEthosSetupComplete.IsDisposed = True Then
                AdminEthosSetupComplete = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Progress Checklist.CourseAdminEthosSetupComplete"), AptifyCheckBox)
            End If
            'If IsBundle Is Nothing OrElse IsBundle.IsDisposed = True Then
            '    IsBundle = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.IsBundledProduct"), AptifyCheckBox)
            'End If
            'If BundledProductTab Is Nothing OrElse BundledProductTab.IsDisposed Then
            '    BundledProductTab = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorAppBundled.Tabs"), FormTemplateTab)
            'End If

            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CourseId = FormTemplateContext.GE.GetValue("CourseId").ToString
                CourseOwnerId = FormTemplateContext.GE.GetValue("CourseOwner").ToString
                CourseOwner = CInt(FormTemplateContext.GE.GetValue("CourseOwner"))
                EventId = FormTemplateContext.GE.GetValue("EventId").ToString
                ProductId = FormTemplateContext.GE.GetValue("ProductId").ToString
                GLId = FormTemplateContext.GE.GetValue("SalesGL").ToString

                CheckCourseOwnerLB()
            Else
                ShowUsersFields()
            End If

            If isClonedCourse.Value = 0 Then
                EthosNodeId.Hide()
            Else
                EthosNodeId.Show()
            End If
            'If IsBundle.Value = 0 Then
            '    BundledProductTab.Show()
            '    CourseCreationTab.Hide()
            'Else
            '    CourseCreationTab.Show()
            '    BundledProductTab.Hide()
            'End If

            'CourseCreationTab.Hide()

            'CheckStatusList()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub


    Private Sub isClonedCourse_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles isClonedCourse.ValueChanged
        'Dim lTypeID As Long = -1
        If (NewValue) = True Then
            EthosNodeId.Show()
            If EthosNodeId.Value = 0 Then
                MsgBox("If this is a cloned course, you must enter the node id of the course from the lms.")
            End If
        Else
            EthosNodeId.Hide()
        End If
    End Sub

    Private Sub EthosNodeId_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles EthosNodeId.ValueChanged
        'Dim lTypeID As Long = -1
        If Not NewValue Is Nothing Then
            'Dim NodeId As Integer = CInt(NewValue)
            If NewValue IsNot "" Then
                Dim NodeId As Integer = CInt(NewValue)
                If NodeId > 0 Then
                Else
                    MsgBox("If this is a cloned course, you must enter the node id of the course from the lms.")
                End If
            Else
                MsgBox("If this is a cloned course, you must enter the node id of the course from the lms.")
            End If

        End If
    End Sub
    Private Sub CourseOwnerIdLinkbox_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles CourseOwnerIdLinkbox.ValueChanged
        Dim lTypeID As Long = -1
        If FormTemplateContext.GE.RecordID = -1 Then
            Me.CourseCreationStatus.Value = 1
        End If

        CheckCourseOwnerLB()



    End Sub

    Public Sub CheckCourseOwnerLB()
        Dim da As New DataAction
        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim dt2 As DataTable
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
        'If Me.FormTemplateContext.GE.RecordID > 0 AndAlso (dt1.Rows.Count > 0 AndAlso dt.Rows.Count = 0) Then
        If Me.FormTemplateContext.GE.RecordID > 0 Then
            If dt.Rows.Count > 0 Then
                ShowAdminFields()
            ElseIf dt1.Rows.Count > 0 Or CourseOwnerPersonId = coursesqlid Then
                ShowUsersFields()
            Else
                HideFields()
            End If
        Else
            MsgBox("This record has not been created yet.  Please save the form to create the record")

        End If

    End Sub


    Public Sub CheckStatusList()
        Dim da As New DataAction
        If FormTemplateContext.GE.RecordID = -1 Then
            CourseCreationStatus.DisplaySQL = "select ID, Name from ACSLMSCourseCreatorStatus where id <= 2"
        End If
        If CInt(CourseId) > 0 Then

            CourseCreationStatus.DisplaySQL = "select ID, Name from ACSLMSCourseCreatorStatus where id in (3,4,12,13)"
            Me.CourseCreationStatus.Value = 3

        End If
        If CInt(EventId) > 0 Then

            CourseCreationStatus.DisplaySQL = "select ID, Name from ACSLMSCourseCreatorStatus where id in (5,6,12,13,14,15)"
            Me.CourseCreationStatus.Value = 5
        End If
        If CInt(ProductId) > 0 Then
            CourseCreationStatus.DisplaySQL = "select ID, Name from ACSLMSCourseCreatorStatus where id in (7,8,12,13,16,17)"
            Me.CourseCreationStatus.Value = 7
        End If
        'If Not GLId Is "" Then
        '    courseCreationStatus.DisplaySQL = "select ID, Name from ACSLMSCourseCreatorStatus where id in (9,10,11)"

        'End If
    End Sub
    Public Sub HideFields()
        CourseRequestTab.Hide()
        CourseCreationTab.Hide()
        UserMessage.Visible = True
        CourseCreationStatus.Hide()
        CourseOwnerIdLinkbox.Hide()
        ContactDepartment.Hide()
        CostCenter.Hide()
        ContactEmail.Hide()
        ContactPhone.Hide()
        RequestedDueDate.Hide()

    End Sub
    Public Sub ShowUsersFields()
        If Not CourseRequestTab Is Nothing Then
            CourseRequestTab.Visible = True
        End If
        If Not CourseCreationTab Is Nothing Then
            CourseCreationTab.Visible = False
        End If
        If Not UserMessage Is Nothing Then
            UserMessage.Visible = False
        End If
        If Not CourseCreationStatus Is Nothing Then
            CourseCreationStatus.Enabled = False
        End If


    End Sub
    Public Sub ShowAdminFields()

        If Not UserMessage Is Nothing Then
            UserMessage.Visible = False
        End If
        'If CourseOwnerIdLinkbox.Value IsNot Nothing AndAlso CInt(CourseOwnerIdLinkbox.Value) > 0 Or userid = 11 Then
        'CourseCreationTab.Enabled = True
        If Not CourseCreationTab Is Nothing Then
            CourseCreationTab.Visible = True
        End If
        If Not CourseRequestTab Is Nothing Then
            CourseRequestTab.Visible = True
        End If

        'Else
        'CourseCreationTab.Visible = False
        'CourseRequestTab.Visible = False
        ' End If

        'RequestedCourseName.Show()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        Me.Name = "ACSLMSCourseCreationLC"
        Me.ResumeLayout(False)

    End Sub

    Private Sub ACSLMSCourseCreationLC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'BackColor = Color.White
        'ForeColor = Color.FromArgb(27, 125, 154)
        'CourseCreationTab.BackColor = Color.White
        'CourseCreationTab.ForeColor = Color.FromArgb(27, 125, 154)



    End Sub

End Class