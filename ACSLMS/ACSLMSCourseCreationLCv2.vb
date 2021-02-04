'Option Explicit On
'Option Strict On

Imports System.Drawing
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.WindowsControls
Imports System.Windows.Forms
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.Application

Public Class ACSLMSCourseCreationLCv2
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private bAdded As Boolean = False
    Private lGridID As Long = -1
    Dim userid As Long = m_oAppObj.UserCredentials.AptifyUserID
    Dim Uservalue As String = m_oAppObj.UserCredentials.GetUserRelatedRecordID(userid)
    Dim courseCreatorGroupSQL As String
    Dim CourseOwnerPersonId As Integer
    Dim courseOwnerSQL As String
    Dim CourseOwnerId As String
    Dim CourseOwner As Integer
    Dim CourseId As String
    Dim EventId As String
    Dim ProductId As String
    Dim GLId As String
    Dim UserCreatedId As String
    Dim loggedInPersonSQL As String
    Dim LoggedInPersonId As Integer

    Private WithEvents CourseOwnerIdLinkbox As AptifyLinkBox
    Private WithEvents CourseCreationStatus As AptifyDataComboBox
    Private WithEvents CourseCreationId As AptifyLinkBox
    Private WithEvents CourseCreatorApp As ApplicationsForm
    Private WithEvents ContactDepartment As AptifyDataComboBox
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
    Private WithEvents EventAdminTab As FormTemplateTab
    Private WithEvents SuvinaTab As FormTemplateTab

    Dim currentDate As DateTime = Now

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



            If CourseOwnerIdLinkbox Is Nothing OrElse CourseOwnerIdLinkbox.IsDisposed = True Then
                CourseOwnerIdLinkbox = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseOwnerId"), AptifyLinkBox)
            End If

            If CourseCreationId Is Nothing OrElse CourseCreationId.IsDisposed = True Then
                CourseCreationId = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseId"), AptifyLinkBox)
            End If

            If ContactDepartment Is Nothing OrElse ContactDepartment.IsDisposed = True Then
                ContactDepartment = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.ContactDepartment"), AptifyDataComboBox)
            End If

            If CostCenter Is Nothing OrElse CostCenter.IsDisposed = True Then
                CostCenter = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CostCenter"), AptifyDataComboBox)
            End If
            If ContactEmail Is Nothing OrElse ContactEmail.IsDisposed = True Then
                ContactEmail = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.ContactEmail"), AptifyTextBox)
            End If
            If ContactPhone Is Nothing OrElse ContactPhone.IsDisposed = True Then
                ContactPhone = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.ContactPhone"), AptifyTextBox)
            End If

            If CourseAdminNotes Is Nothing OrElse CourseAdminNotes.IsDisposed = True Then
                CourseAdminNotes = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.CourseAdminNotes"), AptifyTextBox)
            End If
            If RequestedDueDate Is Nothing OrElse RequestedDueDate.IsDisposed = True Then
                RequestedDueDate = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.RequestedDueDate"), AptifyTextBox)
            End If

            If UserMessage Is Nothing OrElse UserMessage.IsDisposed = True Then
                UserMessage = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Form.Culture Label.1"), CultureLabel)
            End If

            If EventAdminTab Is Nothing OrElse EventAdminTab.IsDisposed = True Then
                EventAdminTab = TryCast(Me.GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp.Tabs2"), FormTemplateTab)
            End If

            If SuvinaTab Is Nothing OrElse SuvinaTab.IsDisposed = True Then
                SuvinaTab = TryCast(Me.GetFormComponentByLayoutKey(Me, "ACSLMSCourseCreatorApp Form - Event: Step 2: Admin Tab"), FormTemplateTab)
            End If

            If userid <> 11 Then

                loggedInPersonSQL = "select e.linkedpersonid from vwUserEntityRelations uer join vwemployees e on e.id = uer.EntityRecordID join vwusers u on u.id = uer.userid where u.id = " & userid
                LoggedInPersonId = CLng(m_oDA.ExecuteScalar(loggedInPersonSQL))
            End If


            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CourseId = FormTemplateContext.GE.GetValue("CourseId").ToString
                CourseOwnerId = FormTemplateContext.GE.GetValue("CourseOwner").ToString
                CourseOwner = CInt(FormTemplateContext.GE.GetValue("CourseOwner"))


            Else
                If Not IsDBNull(LoggedInPersonId) AndAlso LoggedInPersonId > 0 Then
                    CourseOwnerIdLinkbox.Value = LoggedInPersonId
                End If

                If Not UserMessage Is Nothing Then
                    UserMessage.Visible = False
                End If
            End If
            CheckCourseOwnerLB()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Public Sub CheckCourseOwnerLB()
        Dim da As New DataAction
        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim FTP As Integer = FormTemplateID

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

        Dim courseSql As String = "select courseowner from vwacslmscoursecreatorapp where ID = " & Me.FormTemplateContext.GE.RecordID
        Dim coursesqlid As Integer = da.ExecuteScalar(courseSql)

        If userid <> 11 Then
            courseOwnerSQL = "select e.linkedpersonid from vwUserEntityRelations uer join vwemployees e on e.id = uer.EntityRecordID join vwusers u on u.id = uer.userid where u.id = " & userid
            CourseOwnerPersonId = CLng(da.ExecuteScalar(courseOwnerSQL))
        End If

        If Me.FormTemplateContext.GE.RecordID > 0 Then
            If dt.Rows.Count > 0 Or dt1.Rows.Count > 0 Or CourseOwnerPersonId = coursesqlid Then
                ShowAdminFields()
            Else
                HideFields()
            End If
        Else
            'MsgBox("This record has not been created yet.  Please save the form to create the record")

        End If



    End Sub

    Public Sub HideFields()
        'CourseRequestTab.Hide() 
        If Not UserMessage Is Nothing Then
            UserMessage.Visible = True
        End If
        If Not CourseOwnerIdLinkbox Is Nothing Then
            CourseOwnerIdLinkbox.Visible = False
        End If
        If Not ContactDepartment Is Nothing Then
            ContactDepartment.Visible = False
        End If
        If Not CostCenter Is Nothing Then
            CostCenter.Visible = False
        End If
        If Not ContactEmail Is Nothing Then
            ContactEmail.Visible = False
        End If
        If Not ContactPhone Is Nothing Then
            ContactPhone.Visible = False
        End If
        If Not RequestedDueDate Is Nothing Then
            RequestedDueDate.Visible = False
        End If
        'CourseOwnerIdLinkbox.Hide()
        'ContactDepartment.Hide()
        'CostCenter.Hide()
        'ContactEmail.Hide()
        'ContactPhone.Hide()
        'RequestedDueDate.Hide()
        If Not EventAdminTab Is Nothing Then
            EventAdminTab.Hide()
        End If
        'If Not CourseRequestTab Is Nothing Then
        '    CourseRequestTab.Hide()
        'End If
        'CourseCreationTab.Hide()
        'CourseRequestTab.Hide()
    End Sub

    Public Sub ShowAdminFields()

        If Not UserMessage Is Nothing Then
            UserMessage.Visible = False
        End If


        'If CourseOwnerIdLinkbox.Value IsNot Nothing AndAlso CInt(CourseOwnerIdLinkbox.Value) > 0 Or userid = 11 Then
        'CourseCreationTab.Enabled = True
        'If Not CourseCreationTab Is Nothing Then
        '    CourseCreationTab.Visible = True
        'End If
        'If Not CourseRequestTab Is Nothing Then
        '    CourseRequestTab.Visible = True
        'End If

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