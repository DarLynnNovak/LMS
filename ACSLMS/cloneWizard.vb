Option Explicit On
Option Strict On

Imports Aptify.Framework.Application
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.ExceptionManagement
Imports Aptify.Framework.WindowsControls
Imports System.Windows.Forms

Public Class cloneWizard
    Private m_oApp As AptifyApplication
    Private m_oDA As DataAction
    Private WithEvents lnkVersion As New AptifyLinkBox
    Private WithEvents toolTip As New ToolTip
    Dim m_bCancel As Boolean = False
    Private m_lViewID As Long
    Private m_sViewSQL As String
    Private m_lEntityID As Long
    Private m_oSelectedItems() As Object
    Private m_oUserCredential As UserCredentials
    Private bClose As Boolean = False
    Private ReadOnly Property AppObj() As AptifyApplication
        Get
            Return m_oApp
        End Get
    End Property

    Private ReadOnly Property DA() As DataAction
        Get
            If m_oDA Is Nothing Then
                m_oDA = New DataAction(AppObj.UserCredentials)
            End If
            Return m_oDA
        End Get
    End Property
    Public Sub Config(ByVal ApplicationObject As AptifyApplication, ByVal ViewID As Long, ByVal ParamArray SelectedItems() As Object)
        Try
            m_oApp = ApplicationObject
            m_oDA = DA
            Dim pnl As Panel
            For Each c As Control In Me.Controls
                If TypeOf c Is Panel Then
                    pnl = TryCast(c, Panel)
                    If pnl IsNot Nothing Then
                        pnl.Size = New System.Drawing.Size(726, 343)
                        pnl.Location = New System.Drawing.Point(3, 3)

                    End If
                End If
            Next
            SetupToolTip()
            SetupVersionLinkBox()
            Panel1.Visible = True
            Panel1.Controls.Add(lnkVersion)

            m_lViewID = ViewID
            If Me.m_oApp.Command.Parameters.IndexOf("ViewSQL") >= 0 Then
                m_sViewSQL = CStr(Me.m_oApp.Command.Parameters.Item("ViewSQL"))
            Else
                m_sViewSQL = Me.m_oApp.View(m_lViewID).ViewSQL
            End If
            m_lEntityID = CLng(Me.m_oApp.GetEntityID("acslmscoursecreatorapp"))

            m_oSelectedItems = SelectedItems

            If SelectedItems IsNot Nothing AndAlso
                    SelectedItems.Length > 0 Then
                ' Me.radSelected.Enabled = True
                ' Me.radSelected.Text = Me.radSelected.Text & " (" & SelectedItems.Length & ")"
            Else
                ' Me.radSelected.Enabled = False
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub SetupToolTip()
        toolTip.SetToolTip(Me.PictureBox1, "This wizard will copy CCA fields from the entered cca record id \n to a new request record, it will not copy the Course Id, Event Id, Pro" &
        "duct Id or GL Code.")
    End Sub
    Private Sub SetupVersionLinkBox()
        '
        'lnkProduct
        '
        lnkVersion.BackColor = System.Drawing.SystemColors.Control
        lnkVersion.DataControl = Nothing
        lnkVersion.DataField = Nothing
        lnkVersion.DisabledLinkColor = System.Drawing.SystemColors.ControlText
        lnkVersion.DisabledLinkFont = New System.Drawing.Font("Tahoma", 8.0!)
        lnkVersion.DisableGeDataTransfer = False
        lnkVersion.EntityID = CType(m_oApp.GetEntityID("acslmscoursecreatorapp").ToString, Long)
        lnkVersion.EntityName = "ACSLMSCourseCreatorApp"
        lnkVersion.HiddenFilter = ""
        lnkVersion.HyperlinkEnabled = True
        lnkVersion.LabelAutoSize = False
        lnkVersion.LabelFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Underline)
        lnkVersion.LabelRequiredFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Underline)
        'lnkProduct.LabelSize = New System.Drawing.Size(50, 17)
        lnkVersion.LabelAutoSize = True
        lnkVersion.LabelTextAlign = System.Drawing.ContentAlignment.TopRight
        lnkVersion.LinkBoxEnabled = True
        lnkVersion.LinkColor = System.Drawing.SystemColors.HotTrack
        lnkVersion.LinkCursor = System.Windows.Forms.Cursors.Arrow
        lnkVersion.LinkFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Underline)
        lnkVersion.Location = New System.Drawing.Point(340, 180)
        lnkVersion.LookupCursor = System.Windows.Forms.Cursors.WaitCursor
        lnkVersion.Margin = New System.Windows.Forms.Padding(2)
        lnkVersion.Name = "lnkVersion"
        lnkVersion.NewRecordParameterString = Nothing
        lnkVersion.RecordID = CType(-1, Long)
        lnkVersion.RecordName = Nothing
        lnkVersion.ShowLabel = True
        lnkVersion.Size = New System.Drawing.Size(255, 20)
        lnkVersion.TabIndex = 1
        lnkVersion.Value = CType(-1, Long)
        lnkVersion.ValueField = Aptify.Framework.WindowsControls.BoundListBoxValueField.RecordID
        lnkVersion.Visible = True

        'lnkProduct.BoundControlConfig()
        lnkVersion.SetUserCredential(m_oDA.UserCredentials)


    End Sub

    Private Sub cmdProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProcess.Click
        Try
            Dim dt As DataTable
            Dim dt1 As DataTable
            Dim dt2 As DataTable
            Dim bOK As Boolean = True
            Dim sSql As String
            Dim sSql1 As String
            Dim nextCloneName As String

            Dim priceSequenceSql As String
            Dim priceSequenceId As Integer
            Dim CostCenter As String
            Dim CourseSponsoringAssociation As String
            Dim EventTypeId As String
            Dim IsCAActivity As Boolean
            Dim IsCMEActivity As Boolean
            Dim CMECertificate As Integer
            Dim CMECreditAmount As Decimal
            Dim CECertificate As Integer
            Dim CECreditAmount As Decimal
            Dim SACreditAmount As Decimal
            Dim IsRegMandate As Boolean
            Dim RequestedRegMandateType As Integer
            Dim RequestedRegMandateAmount As Decimal
            Dim IsCEActivity As Boolean
            Dim ProductCategoryID As Integer
            Dim ProductEmailTemplateId As Integer
            Dim CreditClaimingExpDate As String
            Dim EventProgramId As Integer

            If lnkVersion.RecordID < 1 Then
                MsgBox("Please enter a CCA request id!", MsgBoxStyle.Information, "ACS Copy Wizard")
                Exit Sub
            End If
            Dim sError As String = ""
            Dim RecordID As Long = lnkVersion.RecordID
            nextCloneName = NewCloneName.Text

            If nextCloneName = "" Then
                MsgBox("You must enter a New Course Name")
            Else
                If MsgBox("The processor will now create a new copy from " & lnkVersion.RecordName & " record. Do you want to continue?", MsgBoxStyle.YesNo, "ACS Version Processing") = MsgBoxResult.Yes Then
                    Cursor = Cursors.WaitCursor
                    sSql = "select * from aptify..acslmscoursecreatorapp where id=" & RecordID
                    dt = m_oDA.GetDataTable(sSql)



                    If dt.Rows.Count > 0 Then
                        For Each dr As DataRow In dt.Rows
                            CostCenter = CStr(dr.Item("CostCenter"))
                            Dim CourseOwner As Integer = CInt(dr.Item("CourseOwner"))
                            Dim ContactPhone As String = CStr(dr.Item("ContactPhone"))
                            Dim ContactDepartment As String = CStr(dr.Item("ContactDepartment"))
                            'Dim ContactEmail As String = CStr(dr.Item("ContactEmail"))
                            Dim CourseDescription As String = CStr(dr.Item("CourseDescription"))
                            Dim CourseCategoryId As Integer = CInt(dr.Item("CourseCategoryId"))
                            Dim CourseStartDate As String = CStr(dr.Item("CourseStartDate"))
                            Dim CourseEndDate As String = CStr(dr.Item("CourseEndDate"))
                            Dim CourseInstructorId As Integer = CInt(dr.Item("CourseInstructorId"))
                            Dim CourseSchoolId As Integer = CInt(dr.Item("CourseSchoolId"))
                            If Not IsDBNull(dr.Item("CourseSponsoringAssociation")) Then
                                CourseSponsoringAssociation = CStr(dr.Item("CourseSponsoringAssociation"))
                            End If
                            If Not IsDBNull(dr.Item("EventType")) Then
                                EventTypeId = CStr(dr.Item("EventType"))
                            End If
                            If Not IsDBNull(dr.Item("IsCEActivity")) Then
                                IsCAActivity = CBool(dr.Item("IsCEActivity"))
                            End If
                            If Not IsDBNull(dr.Item("IsCMEActivity")) Then
                                IsCMEActivity = CBool(dr.Item("IsCMEActivity"))
                            End If
                            If Not IsDBNull(dr.Item("CMECertificate")) Then
                                CMECertificate = CInt(dr.Item("CMECertificate"))
                            End If
                            If Not IsDBNull(dr.Item("CMECreditAmount")) Then
                                CMECreditAmount = CDec(dr.Item("CMECreditAmount"))
                            End If

                            If Not IsDBNull(dr.Item("CECertificate")) Then
                                CECertificate = CInt(dr.Item("CECertificate"))
                            End If
                            If Not IsDBNull(dr.Item("CECreditAmount")) Then
                                CECreditAmount = CDec(dr.Item("CECreditAmount"))
                            End If
                            If Not IsDBNull(dr.Item("SACreditAmount")) Then
                                SACreditAmount = CDec(dr.Item("SACreditAmount"))
                            End If
                            If Not IsDBNull(dr.Item("IsRegMandate")) Then
                                IsRegMandate = CBool(dr.Item("IsRegMandate"))
                            End If
                            If Not IsDBNull(dr.Item("RequestedRegMandateType")) Then
                                RequestedRegMandateType = CInt(dr.Item("RequestedRegMandateType"))
                            End If
                            If Not IsDBNull(dr.Item("RequestedRegMandateAmount")) Then
                                RequestedRegMandateAmount = CDec(dr.Item("RequestedRegMandateAmount"))
                            End If
                            If Not IsDBNull(dr.Item("IsCEActivity")) Then
                                IsCEActivity = CBool(dr.Item("IsCEActivity"))
                            End If
                            If Not IsDBNull(dr.Item("ProductCategoryID")) Then
                                ProductCategoryID = CInt(dr.Item("ProductCategoryID"))
                            End If
                            If Not IsDBNull(dr.Item("ProductEmailTemplateId")) Then
                                ProductEmailTemplateId = CInt(dr.Item("ProductEmailTemplateId"))
                            End If
                            If Not IsDBNull(dr.Item("CreditClaimingExpirationDate")) Then
                                CreditClaimingExpDate = CStr(dr.Item("CreditClaimingExpirationDate"))
                            End If
                            If Not IsDBNull(dr.Item("EventProgramId")) Then
                                EventProgramId = CInt(dr.Item("EventProgramId"))
                            End If

                            Dim oACSLMSCourseCreatorAppGE As AptifyGenericEntityBase
                            oACSLMSCourseCreatorAppGE = m_oApp.GetEntityObject("acslmscoursecreatorapp", -1)
                            With oACSLMSCourseCreatorAppGE
                                .SetValue("CostCenter", CostCenter)
                                .SetValue("CourseOwner", CourseOwner)
                                .SetValue("CourseName", nextCloneName)
                                .SetValue("CourseDescription", CourseDescription)
                                .SetValue("ContactDepartment", ContactDepartment)
                                .SetValue("ContactPhone", ContactPhone)
                                '.SetValue("ContactEmail", ContactEmail)
                                .SetValue("CourseCategoryId", CourseCategoryId)
                                .SetValue("CourseStartDate", CourseStartDate)
                                .SetValue("CourseEndDate", CourseEndDate)
                                .SetValue("CourseInstructorId", CourseInstructorId)
                                .SetValue("CourseSchoolId", CourseSchoolId)
                                .SetValue("CourseSponsoringAssociation", CourseSponsoringAssociation)
                                .SetValue("EventType", EventTypeId)
                                .SetValue("IsCAActivity", IsCAActivity)
                                .SetValue("IsCMEActivity", IsCMEActivity)
                                .SetValue("CMECertificate", CMECertificate)
                                .SetValue("CMECreditAmount", CMECreditAmount)
                                .SetValue("CECertificate", CECertificate)
                                .SetValue("CECreditAmount", CECreditAmount)
                                .SetValue("SACreditAmount", SACreditAmount)
                                .SetValue("IsRegMandate", IsRegMandate)
                                .SetValue("RequestedRegMandateType", RequestedRegMandateType)
                                .SetValue("RequestedRegMandateAmount", RequestedRegMandateAmount)
                                .SetValue("IsCEActivity", IsCEActivity)
                                .SetValue("ProductCategoryID", ProductCategoryID)
                                .SetValue("ProductEmailTemplateId", ProductEmailTemplateId)
                                .SetValue("CreditClaimingExpirationDate", CreditClaimingExpDate)
                                .SetValue("EventProgramId", EventProgramId)
                            End With

                            sSql1 = "select * from ACSLMSCourseCreatorProductPrices where ACSLMSCourseCreatorAppID = " & RecordID
                            dt1 = m_oDA.GetDataTable(sSql1)

                            If dt1.Rows.Count > 0 Then
                                For Each dr1 As DataRow In dt1.Rows
                                    Dim ProductPriceName As String = CStr(dr1.Item("ProductPriceName"))
                                    Dim ProductFilterMemberType As String = CStr(dr1.Item("ProductFilterMemberType"))
                                    Dim ProductFilterRule As String = CStr(dr1.Item("ProductFilterRule"))
                                    Dim ProductFilterRuleValue As Integer = CInt(dr1.Item("ProductFilterRuleValue"))
                                    Dim ProductFilterRulePrice As Decimal = CDec(dr1.Item("ProductFilterRulePrice"))
                                    Dim Sequence As Integer = CInt(dr1.Item(Sequence))
                                    'priceSequenceSql = "select case when (select max(pp.Sequence) from acslmscoursecreatorproductprices pp where ACSLMSCourseCreatorAppID = " & RecordID & " ) is null then 1 else  (select max(pp.Sequence) from acslmscoursecreatorproductprices pp where ID = " & RecordID & ") + 1 end"
                                    'priceSequenceId = CInt(DA.ExecuteScalar(priceSequenceSql))


                                    With oACSLMSCourseCreatorAppGE.SubTypes("acslmscoursecreatorproductprices").Add()
                                        .SetValue("Sequence", Sequence)
                                        .SetValue("ProductPriceName", ProductPriceName)
                                        .SetValue("ProductFilterMemberType", ProductFilterMemberType)
                                        .SetValue("ProductFilterRule", ProductFilterRule)
                                        .SetValue("ProductFilterRuleValue", ProductFilterRuleValue)
                                        .SetValue("ProductFilterRulePrice", ProductFilterRulePrice)

                                    End With
                                Next
                            End If


                            If oACSLMSCourseCreatorAppGE.Save(False) = False Then
                                If sError = "" Then
                                    sError = "copying failed to save for record " & nextCloneName & ". Error - " & oACSLMSCourseCreatorAppGE.LastUserError.ToString
                                Else
                                    sError = sError + Environment.NewLine + Environment.NewLine + "copying failed to save for record " & nextCloneName & ". Error - " & oACSLMSCourseCreatorAppGE.LastUserError.ToString
                                End If
                            Else

                                Me.lblMessage.Text = "Generating copy: " & lnkVersion.RecordName.ToString

                            End If

                            System.Windows.Forms.Application.DoEvents()

                            Cursor = Cursors.Default

                            If sError <> "" Then
                                Panel1.Visible = False
                                Panel2.Visible = True
                                txtError.Text = sError
                            Else
                                If bClose = False Then
                                    MsgBox("The copy processing is complete.", MsgBoxStyle.Information, "ACS Copy Processing")
                                    Me.Close()
                                Else
                                    Me.Close()
                                End If
                            End If
                        Next
                    End If

                End If
            End If



        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        bClose = True
        Me.Close()
    End Sub


End Class