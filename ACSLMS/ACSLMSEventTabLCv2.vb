
'Option Explicit On
'Option Strict On

Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.BusinessLogic.GenericEntity

Public Class ACSLMSEventTabLCv2
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Dim CourseCreatorAppGE As AptifyGenericEntityBase
    Dim EventGE As AptifyGenericEntityBase
    Dim ChildEventGE As AptifyGenericEntityBase
    Dim ChildEventIdSql As String
    Dim ChildEvent As DataTable
    Private WithEvents EventCreateButton As AptifyActiveButton
    Private WithEvents EventUpdateButton As AptifyActiveButton
    Private WithEvents EventIdLB As AptifyLinkBox
    Private WithEvents EventNameTB As AptifyTextBox
    Private WithEvents EventProgramIdLB As AptifyLinkBox
    Private WithEvents CMEProgramTB As AptifyTextBox
    Private WithEvents CMELocationTB As AptifyTextBox
    Private WithEvents CMENameOrderTB As AptifyTextBox
    Private WithEvents EventTypeDCB As AptifyDataComboBox
    Private WithEvents JointSponsorTB As AptifyTextBox
    Private WithEvents CertificateLine1TB As AptifyTextBox
    Private WithEvents CertificateLine2TB As AptifyTextBox
    Private WithEvents CertificateLine3TB As AptifyTextBox
    Private WithEvents CMECertificateIdLB As AptifyLinkBox
    Private WithEvents CECertificateIdLB As AptifyLinkBox
    Private WithEvents CACertificateIdLB As AptifyLinkBox
    Private WithEvents DatePrintCB As AptifyComboBox
    Private WithEvents LocationPrintCB As AptifyComboBox
    Private WithEvents CertificateVersionCB As AptifyComboBox
    Private WithEvents CMECreditsTB As AptifyTextBox
    Private WithEvents CMESACreditsTB As AptifyTextBox
    Private WithEvents CECreditsTB As AptifyTextBox
    Private WithEvents CESACreditsTB As AptifyTextBox
    Private WithEvents CACreditsTB As AptifyTextBox
    Private WithEvents RMCreditsTB As AptifyTextBox
    Private WithEvents StateMandateCB As AptifyCheckBox
    Private WithEvents EventStartDateTB As AptifyTextBox
    Private WithEvents EventEndDateTB As AptifyTextBox
    Private WithEvents ClaimingExpDateTB As AptifyTextBox
    Private WithEvents Paragraph1CB As AptifyComboBox
    Private WithEvents IsRegMandateCB As AptifyCheckBox
    Private WithEvents RegMandateTypeCB As AptifyLinkBox
    Private WithEvents SuvinaTab As FormTemplateTab
    Dim ID As Long

    Dim CourseCreationEventId As Long
    Dim result As String = "Failed"
    Dim getEventIdSql As String
    Dim EventId As Integer
    Dim getEventTypeNameSql As String
    Dim EventTypeName As String
    Dim da As New DataAction

    Dim courseCreatorGroupSQL As String
    Dim courseOwnerSQL As String
    Dim CourseOwnerId As String
    Dim CourseOwner As Integer
    Dim UserCreatedId As String
    Dim CourseOwnerPersonId As Integer
    Dim IsCOPSQL As String
    Dim IsCop As Boolean
    Dim copCertIdSql As String
    Dim copCertId As Integer


    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try

            'Me.AutoScroll = True
            'Dim newFormTemplateid As Integer = 27183
            ID = FormTemplateContext.GE.GetValue("ID")
            FindControls()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e) 
    End Sub
    Protected Overridable Sub FindControls()
        Try
            Me.Hide()
            If EventCreateButton Is Nothing OrElse EventCreateButton.IsDisposed = True Then
                EventCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.Active Button.1"), AptifyActiveButton)
            End If
            If EventUpdateButton Is Nothing OrElse EventUpdateButton.IsDisposed = True Then
                EventUpdateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.Active Button.2"), AptifyActiveButton)
            End If

            If EventIdLB Is Nothing OrElse EventIdLB.IsDisposed = True Then
                EventIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.EventId"), AptifyLinkBox)
            End If
            If EventNameTB Is Nothing OrElse EventNameTB.IsDisposed = True Then
                EventNameTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.EventName"), AptifyTextBox)
            End If
            If EventProgramIdLB Is Nothing OrElse EventProgramIdLB.IsDisposed = True Then
                EventProgramIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.EventProgramId"), AptifyLinkBox)
            End If
            If CMEProgramTB Is Nothing OrElse CMEProgramTB.IsDisposed = True Then
                CMEProgramTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CMEProgram"), AptifyTextBox)
            End If
            If CMELocationTB Is Nothing OrElse CMELocationTB.IsDisposed = True Then
                CMELocationTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CMELocation"), AptifyTextBox)
            End If

            If CMENameOrderTB Is Nothing OrElse CMENameOrderTB.IsDisposed = True Then
                CMENameOrderTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.NameOrder"), AptifyTextBox)
            End If
            If EventTypeDCB Is Nothing OrElse EventTypeDCB.IsDisposed = True Then
                EventTypeDCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.EventType"), AptifyDataComboBox)
            End If
            If JointSponsorTB Is Nothing OrElse JointSponsorTB.IsDisposed = True Then
                JointSponsorTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.JointSponsorSociety"), AptifyTextBox)
            End If
            If CertificateLine1TB Is Nothing OrElse CertificateLine1TB.IsDisposed = True Then
                CertificateLine1TB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CertificateLine1"), AptifyTextBox)
            End If
            If CertificateLine2TB Is Nothing OrElse CertificateLine2TB.IsDisposed = True Then
                CertificateLine2TB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CertificateLine2"), AptifyTextBox)
            End If
            If CertificateLine3TB Is Nothing OrElse CertificateLine3TB.IsDisposed = True Then
                CertificateLine3TB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CertificateLine3"), AptifyTextBox)
            End If
            If CMECertificateIdLB Is Nothing OrElse CMECertificateIdLB.IsDisposed = True Then
                CMECertificateIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CMECertificateId"), AptifyLinkBox)
            End If
            If CECertificateIdLB Is Nothing OrElse CECertificateIdLB.IsDisposed = True Then
                CECertificateIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CECertificateId"), AptifyLinkBox)
            End If
            If CACertificateIdLB Is Nothing OrElse CACertificateIdLB.IsDisposed = True Then
                CACertificateIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CACertificateId"), AptifyLinkBox)
            End If

            If DatePrintCB Is Nothing OrElse DatePrintCB.IsDisposed = True Then
                DatePrintCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.DatePrint"), AptifyComboBox)
            End If
            If LocationPrintCB Is Nothing OrElse LocationPrintCB.IsDisposed = True Then
                LocationPrintCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.LocationPrint"), AptifyComboBox)
            End If
            If CertificateVersionCB Is Nothing OrElse CertificateVersionCB.IsDisposed = True Then
                CertificateVersionCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CertificateVersion"), AptifyComboBox)
            End If
            If CMECreditsTB Is Nothing OrElse CMECreditsTB.IsDisposed = True Then
                CMECreditsTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CMEMaxCredits"), AptifyTextBox)
            End If
            If CMESACreditsTB Is Nothing OrElse CMESACreditsTB.IsDisposed = True Then
                CMESACreditsTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.SACMEMaxCredits"), AptifyTextBox)
            End If
            If CECreditsTB Is Nothing OrElse CECreditsTB.IsDisposed = True Then
                CECreditsTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CEMaxCredits"), AptifyTextBox)
            End If
            If CESACreditsTB Is Nothing OrElse CESACreditsTB.IsDisposed = True Then
                CESACreditsTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.SACEMaxCredits"), AptifyTextBox)
            End If

            If CACreditsTB Is Nothing OrElse CACreditsTB.IsDisposed = True Then
                CACreditsTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CAMaxCredits"), AptifyTextBox)
            End If
            If RMCreditsTB Is Nothing OrElse RMCreditsTB.IsDisposed = True Then
                RMCreditsTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.RegMandateAmount"), AptifyTextBox)
            End If
            If StateMandateCB Is Nothing OrElse StateMandateCB.IsDisposed = True Then
                StateMandateCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.isRegMandate"), AptifyCheckBox)
            End If
            If EventStartDateTB Is Nothing OrElse EventStartDateTB.IsDisposed = True Then
                EventStartDateTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.EventStartDate"), AptifyTextBox)
            End If
            If EventEndDateTB Is Nothing OrElse EventEndDateTB.IsDisposed = True Then
                EventEndDateTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.EventEndDate"), AptifyTextBox)
            End If
            If ClaimingExpDateTB Is Nothing OrElse ClaimingExpDateTB.IsDisposed = True Then
                ClaimingExpDateTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.CreditClaimingExpirationDate"), AptifyTextBox)
            End If
            If Paragraph1CB Is Nothing OrElse Paragraph1CB.IsDisposed = True Then
                Paragraph1CB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.Paragraph1"), AptifyComboBox)
            End If

            If IsRegMandateCB Is Nothing OrElse IsRegMandateCB.IsDisposed = True Then
                IsRegMandateCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.isRegMandate"), AptifyCheckBox)
            End If
            If RegMandateTypeCB Is Nothing OrElse RegMandateTypeCB.IsDisposed = True Then
                RegMandateTypeCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Event: Step 2 Admin.RegMandateType"), AptifyLinkBox)
            End If

            If CInt(EventIdLB.Value) > 0 Then


            Else

                'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", Me.FormTemplateContext.GE.RecordID)
                JointSponsorTB.Value = FormTemplateContext.GE.GetValue("CourseSponsoringAssociation")
                EventStartDateTB.Value = FormTemplateContext.GE.GetValue("RequestEventStartDate")
                EventEndDateTB.Value = FormTemplateContext.GE.GetValue("RequestEventEndDate")
                ClaimingExpDateTB.Value = FormTemplateContext.GE.GetValue("RequestedClaimingExpDate")
                StateMandateCB.Value = FormTemplateContext.GE.GetValue("IsRegMandate")
                CMEProgramTB.Value = FormTemplateContext.GE.GetValue("CourseName")
                CMECreditsTB.Value = FormTemplateContext.GE.GetValue("CMECreditAmount")
                CECreditsTB.Value = FormTemplateContext.GE.GetValue("CECreditAmount")
                CACreditsTB.Value = FormTemplateContext.GE.GetValue("CACreditAmount")
                CMESACreditsTB.Value = FormTemplateContext.GE.GetValue("SACreditAmount")
                CESACreditsTB.Value = FormTemplateContext.GE.GetValue("SACECreditAmount")
                RMCreditsTB.Value = FormTemplateContext.GE.GetValue("RequestedRegMandateAmount")
                EventTypeDCB.Value = FormTemplateContext.GE.GetValue("CourseFormat")

                RegMandateTypeCB.Value = FormTemplateContext.GE.GetValue("RequestedRegMandateType")
                'EventTypeDCB.Value = FormTemplateContext.GE.GetValue("CourseFormat")
                ' If EventNameTB.Value Is Nothing Then
                If Not IsDBNull(EventTypeDCB.Value) Then
                    getEventTypeNameSql = "select name from acscmeeventtype where id = " & EventTypeDCB.Value
                    EventTypeName = m_oDA.ExecuteScalar(getEventTypeNameSql)
                End If


                EventNameTB.Value = FormTemplateContext.GE.GetValue("ContactDepartment") & "_" & EventTypeName & "_" & FormTemplateContext.GE.GetValue("CourseName") & "_" & Year(Now())
                    ' End If
                    CMENameOrderTB.Value = FormTemplateContext.GE.GetValue("ContactDepartment") & "_" & EventTypeName & "_" & FormTemplateContext.GE.GetValue("CourseName") & "_" & Year(Now())
                    DatePrintCB.Value = "D"
                    LocationPrintCB.Value = "E"
                    CertificateVersionCB.Value = "V02"

                End If
                CheckCourseOwnerLB()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub checkCop()
        IsCOPSQL = "select IsCAActivity from acslmscoursecreatorapp where id = " & ID
        IsCop = CInt(da.ExecuteScalar(IsCOPSQL))


        If IsCop = True Then
            'copCertIdSql = "select acscmecerttemplate from acslmscoursecreatorapp where id = " & ID
            copCertId = 26
            CMECertificateIdLB.Value = copCertId
            CECertificateIdLB.Value = copCertId
            CACertificateIdLB.Value = copCertId
        Else
            CMECertificateIdLB.Value = FormTemplateContext.GE.GetValue("CMECertificate")
            CECertificateIdLB.Value = FormTemplateContext.GE.GetValue("CECertificate")
            CACertificateIdLB.Value = FormTemplateContext.GE.GetValue("CACertificate")
        End If

    End Sub

    Private Sub EventCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EventCreateButton.Click
        Try

            If Me.FormTemplateContext.GE.RecordID > 0 Then
                'ParentForm.Close()
                CreateEventId()
            Else
                'Select Case MsgBox("This record has not been created yet.  Please save the form to create the record and populate the fields on this tab from the data you have already entered. Save the form now?", MsgBoxStyle.YesNo, "CCA")
                '    Case MsgBoxResult.Yes
                '        GE


                '    Case Else
                'End Select
                MsgBox("This record has not been created yet.  Please save the form to create the record and populate the fields on this tab from the data you have already entered.")
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub


    Private Sub EventUpdateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EventUpdateButton.Click
        Try

            UpdateEventId()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Public Function CreateEventId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        CourseCreationEventId = Me.FormTemplateContext.GE.GetValue("EventId")
        Try
            ' UpdateCourseCreator()

            EventGE = m_oAppObj.GetEntityObject("ACSCMEEvent", -1)
            If CourseCreationEventId <= 0 Then
                If Len(EventNameTB.Value) > 100 Then
                    result = "Failed"
                    MsgBox("The Event name must be 100 characters or less.  The Event Name is: " & Len(EventNameTB.Value) & " characters")
                Else
                    EventGE.SetValue("Name", EventNameTB.Value)
                End If
                If EventProgramIdLB.Value > 0 Then
                    EventGE.SetValue("ProgramID", EventProgramIdLB.Value)
                Else
                    MsgBox("The program for the event must be filled out, leaving it blank will result in error.")
                End If

                EventGE.SetValue("CME_Program", CMEProgramTB.Value)

                EventGE.SetValue("CME_Location", CMELocationTB.Value)
                If Len(CMENameOrderTB.Value) > 100 Then
                    result = "Failed"
                    MsgBox("The Name Order name must be 100 characters or less.  The Name Order is: " & Len(CMENameOrderTB.Value) & " characters")
                Else
                    EventGE.SetValue("NameOrder", CMENameOrderTB.Value)
                    End If
                    EventGE.SetValue("CME_Start_Date", EventStartDateTB.Value)
                    EventGE.SetValue("CME_End_Date", EventEndDateTB.Value)
                    EventGE.SetValue("CME_Max_Credits", CMECreditsTB.Value)
                    EventGE.SetValue("SACME_Max_Credits", CMESACreditsTB.Value)
                    EventGE.SetValue("CE_Max_Credits", CECreditsTB.Value)
                    EventGE.SetValue("SACE_Max_Credits", CESACreditsTB.Value)
                    EventGE.SetValue("CA_Max_Credits", CACreditsTB.Value)
                    EventGE.SetValue("AwardStateMandated", StateMandateCB.Value)
                    EventGE.SetValue("CMETypeId", RegMandateTypeCB.Value)
                    EventGE.SetValue("CertLine1", CertificateLine1TB.Value)
                    EventGE.SetValue("CertLine2", CertificateLine2TB.Value)
                    EventGE.SetValue("CertLine3", CertificateLine3TB.Value)
                    EventGE.SetValue("Paragraph1", Paragraph1CB.Value)
                    EventGE.SetValue("DatePrint", DatePrintCB.Value)
                    EventGE.SetValue("LocationPrint", LocationPrintCB.Value)
                    EventGE.SetValue("EventType", EventTypeDCB.Value)
                    EventGE.SetValue("jspsociety", JointSponsorTB.Value)


                EventGE.SetValue("CertificateVersion", CertificateVersionCB.Value)

                    If EventGE.IsDirty Then 'if the ge has changed then save
                        If Not EventGE.Save(False) Then
                            Throw New Exception("Problem Saving Event Record:" & EventGE.RecordID)
                            result = "Error"
                        Else
                            EventGE.Save(True)
                            result = "Success"
                            If StateMandateCB.Value = True Then
                                CreateChildEventId()
                            End If

                        End If
                    End If
                    If result = "Success" Then

                        EventIdLB.Value = EventGE.RecordID
                        MsgBox("Success.  Please be sure to save your changes when closing this form.")
                        SetCourseCreatorDate()
                        DisplayEntity()
                        'Select Case MsgBox("Success, Please save this form.  Save? ", MsgBoxStyle.YesNo, "CCA")
                        '    Case MsgBoxResult.Yes
                        '        FormTemplateContext.GE.Save()

                        '    Case Else
                        'End Select

                    End If

                End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Function


    Public Function CreateChildEventId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction


        Try

            ChildEventGE = m_oAppObj.GetEntityObject("ACSCMEEvent", -1)

            If Len(EventNameTB.Value) > 100 Then
                result = "Failed"
                MsgBox("The Event name must be 100 characters or less.  The Event Name is: " & Len(EventNameTB.Value) & " characters")
            Else
                ChildEventGE.SetValue("Name", EventNameTB.Value)
            End If
            ChildEventGE.SetValue("ParentId", EventGE.RecordID)
            ChildEventGE.SetValue("ProgramID", EventProgramIdLB.Value)
            ChildEventGE.SetValue("CME_Location", LocationPrintCB.Value)
            ChildEventGE.SetValue("CMETypeID", RegMandateTypeCB.Value)
            If Len(CMENameOrderTB.Value) > 100 Then
                result = "Failed"
                MsgBox("The Name Order name must be 100 characters or less.  The Name Order is: " & Len(CMENameOrderTB.Value) & " characters")
            Else
                ChildEventGE.SetValue("NameOrder", CMENameOrderTB.Value)
            End If

            ChildEventGE.SetValue("CME_Start_Date", EventStartDateTB.Value)
            ChildEventGE.SetValue("CME_End_Date", EventEndDateTB.Value)
            ChildEventGE.SetValue("CME_Max_Credits", CMECreditsTB.Value)
            ChildEventGE.SetValue("SACME_Max_Credits", CMESACreditsTB.Value)
            ChildEventGE.SetValue("EventType", EventTypeDCB.Value)
            ChildEventGE.SetValue("CMETypeId", RegMandateTypeCB.Value)
            ChildEventGE.SetValue("CertLine1", CertificateLine1TB.Value)
            ChildEventGE.SetValue("CertLine2", CertificateLine2TB.Value)
            ChildEventGE.SetValue("CertLine3", CertificateLine3TB.Value)
            ChildEventGE.SetValue("Paragraph1", Paragraph1CB.Value)
            ChildEventGE.SetValue("DatePrint", DatePrintCB.Value)
            ChildEventGE.SetValue("LocationPrint", LocationPrintCB.Value)
            ChildEventGE.SetValue("EventType", EventTypeDCB.Value)
            ChildEventGE.SetValue("jspsociety", JointSponsorTB.Value)

            ChildEventGE.SetValue("CertificateVersion", CertificateVersionCB.Value)

            If ChildEventGE.IsDirty Then 'if the ge has changed then save
                If Not ChildEventGE.Save(False) Then
                    Throw New Exception("Problem Saving Event Record:" & ChildEventGE.RecordID)
                    result = "Error"
                Else
                    ChildEventGE.Save(True)
                    result = "Success"

                End If
            End If



        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Function
    Public Function UpdateEventId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        CourseCreationEventId = Me.FormTemplateContext.GE.GetValue("EventId")
        Try


            Me.EventIdLB.ClearData()
            ' UpdateCourseCreator()
            EventGE = m_oAppObj.GetEntityObject("ACSCMEEvent", CourseCreationEventId)
            If CourseCreationEventId > 0 Then
                If Len(EventNameTB.Value) > 100 Then
                    result = "Failed"
                    MsgBox("The Event name must be 100 characters or less.  The Event Name is: " & Len(EventNameTB.Value) & " characters")
                Else
                    EventGE.SetValue("Name", EventNameTB.Value)
                End If
                If Not IsDBNull(EventProgramIdLB.Value) Then
                    EventGE.SetValue("ProgramID", EventProgramIdLB.Value)
                Else
                    MsgBox("The program for the event must be filled out, leaving it blank will result in error.")
                End If
                EventGE.SetValue("CME_Program", CMEProgramTB.Value)
                EventGE.SetValue("CME_Location", CMELocationTB.Value)
                If Len(CMENameOrderTB.Value) > 100 Then
                    result = "Failed"
                    MsgBox("The Name Order name must be 100 characters or less.  The Name Order is: " & Len(CMENameOrderTB.Value) & " characters")
                Else
                    EventGE.SetValue("NameOrder", CMENameOrderTB.Value)
                End If
                EventGE.SetValue("CME_Start_Date", EventStartDateTB.Value)
                EventGE.SetValue("CME_End_Date", EventEndDateTB.Value)
                EventGE.SetValue("CME_Max_Credits", CMECreditsTB.Value)
                EventGE.SetValue("SACME_Max_Credits", CMESACreditsTB.Value)
                EventGE.SetValue("CE_Max_Credits", CECreditsTB.Value)
                EventGE.SetValue("SACE_Max_Credits", CESACreditsTB.Value)
                EventGE.SetValue("CA_Max_Credits", CACreditsTB.Value)
                EventGE.SetValue("AwardStateMandated", StateMandateCB.Value)
                EventGE.SetValue("CertLine1", CertificateLine1TB.Value)
                EventGE.SetValue("CertLine2", CertificateLine2TB.Value)
                EventGE.SetValue("CertLine3", CertificateLine3TB.Value)
                EventGE.SetValue("Paragraph1", Paragraph1CB.Value)
                EventGE.SetValue("DatePrint", DatePrintCB.Value)
                EventGE.SetValue("LocationPrint", LocationPrintCB.Value)
                EventGE.SetValue("EventType", EventTypeDCB.Value)
                EventGE.SetValue("jspsociety", JointSponsorTB.Value)

                EventGE.SetValue("CertificateVersion", CertificateVersionCB.Value)

                If EventGE.IsDirty Then 'if the ge has changed then save
                    If Not EventGE.Save(False) Then
                        Throw New Exception("Problem Saving Event Record:" & EventGE.RecordID)
                        result = "Error"
                    Else
                        EventGE.Save(True)
                        result = "Success"
                        If StateMandateCB.Value = True Then
                            ChildEventIdSql = "select * from ACSCMEEvent where parentid = " & EventGE.RecordID

                            ChildEvent = m_oDA.GetDataTable(ChildEventIdSql)
                            If ChildEvent.Rows.Count > 0 Then
                                UpdateChildEventId()
                            Else
                                CreateChildEventId()
                            End If

                        End If
                    End If
                End If
                If result = "Success" Then

                    EventIdLB.Value = EventGE.RecordID
                    MsgBox("Success.  Please be sure to save your changes when closing this form.")
                    'Select Case MsgBox("Success, Please save this form.  Save? ", MsgBoxStyle.YesNo, "Course Creator")
                    '    Case MsgBoxResult.Yes
                    '        ParentForm.Close()
                    UpdateCourseCreator()
                            DisplayEntity()
                    '    Case MsgBoxResult.No

                    '    Case Else
                    'End Select
                End If

            End If



        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Function

    Public Function UpdateChildEventId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction


        ChildEventIdSql = "select * from ACSCMEEvent where parentid = " & EventGE.RecordID

        ChildEvent = m_oDA.GetDataTable(ChildEventIdSql)

        Try
            If ChildEvent.Rows.Count > 0 Then
                For Each dr As DataRow In ChildEvent.Rows
                    ChildEventGE = m_oAppObj.GetEntityObject("ACSCMEEvent", dr.Item("Id"))

                    If Len(EventNameTB.Value) > 100 Then
                        result = "Failed"
                        MsgBox("The Event name must be 100 characters or less.  The Event Name is: " & Len(EventNameTB.Value) & " characters")
                    Else
                        ChildEventGE.SetValue("Name", EventNameTB.Value)
                    End If
                    'ChildEventGE.SetValue("ParentId", EventGE.RecordID)
                    ChildEventGE.SetValue("ProgramID", EventProgramIdLB.Value)
                    ChildEventGE.SetValue("CME_Location", LocationPrintCB.Value)
                    ChildEventGE.SetValue("CMETypeID", RegMandateTypeCB.Value)
                    If Len(CMENameOrderTB.Value) > 100 Then
                        result = "Failed"
                        MsgBox("The Name Order name must be 100 characters or less.  The Name Order is: " & Len(CMENameOrderTB.Value) & " characters")
                    Else
                        ChildEventGE.SetValue("NameOrder", CMENameOrderTB.Value)
                    End If

                    ChildEventGE.SetValue("CME_Start_Date", EventStartDateTB.Value)
                    ChildEventGE.SetValue("CME_End_Date", EventEndDateTB.Value)
                    ChildEventGE.SetValue("CME_Max_Credits", CMECreditsTB.Value)
                    ChildEventGE.SetValue("SACME_Max_Credits", CMESACreditsTB.Value)
                    ChildEventGE.SetValue("EventType", EventTypeDCB.Value)
                    ChildEventGE.SetValue("CMETypeId", RegMandateTypeCB.Value)
                    ChildEventGE.SetValue("CertLine1", CertificateLine1TB.Value)
                    ChildEventGE.SetValue("CertLine2", CertificateLine2TB.Value)
                    ChildEventGE.SetValue("CertLine3", CertificateLine3TB.Value)
                    ChildEventGE.SetValue("Paragraph1", Paragraph1CB.Value)
                    ChildEventGE.SetValue("DatePrint", DatePrintCB.Value)
                    ChildEventGE.SetValue("LocationPrint", LocationPrintCB.Value)
                    ChildEventGE.SetValue("EventType", EventTypeDCB.Value)
                    ChildEventGE.SetValue("jspsociety", JointSponsorTB.Value)

                    ChildEventGE.SetValue("CertificateVersion", CertificateVersionCB.Value)

                    If ChildEventGE.IsDirty Then 'if the ge has changed then save
                        If Not ChildEventGE.Save(False) Then
                            Throw New Exception("Problem Saving Event Record:" & ChildEventGE.RecordID)
                            result = "Error"
                        Else
                            ChildEventGE.Save(True)
                            result = "Success"

                        End If
                    End If
                Next
            End If






        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Function



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
        If m_oDA.UserCredentials.Server.ToLower = "testaptify610" Then
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
        checkCop()
        'If Me.FormTemplateContext.GE.RecordID > 0 AndAlso (dt1.Rows.Count > 0 AndAlso dt.Rows.Count = 0) Then
        If FormTemplateContext.GE.RecordID > 0 Then
            If dt.Rows.Count > 0 AndAlso CInt(EventIdLB.Value) > 0 Then
                EventCreateButton.Visible = False
                EventUpdateButton.Visible = True

            ElseIf dt.Rows.Count > 0 AndAlso CInt(EventIdLB.Value) < 0 Then
                EventCreateButton.Visible = True
                EventUpdateButton.Visible = False
            ElseIf dt.Rows.Count < 0 Then
                EventCreateButton.Visible = False
                EventUpdateButton.Visible = False
            End If
        Else

            EventCreateButton.Visible = False
            EventUpdateButton.Visible = False
        End If

    End Sub

    Public Function SetCourseCreatorDate() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.FormTemplateContext.GE.RecordID
        Try
            'With CourseCreatorAppGE
            'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
            FormTemplateContext.GE.SetValue("EventIdCreated", 1)
            FormTemplateContext.GE.SetValue("EventIdCreatedDate", Now())
            FormTemplateContext.GE.SetValue("EventName", EventNameTB.Value)
            FormTemplateContext.GE.SetValue("EventProgramId", EventProgramIdLB.Value)
            FormTemplateContext.GE.SetValue("CMEProgram", CMEProgramTB.Value)
            FormTemplateContext.GE.SetValue("CMELocation", CMELocationTB.Value)
            FormTemplateContext.GE.SetValue("NameOrder", CMENameOrderTB.Value)
            FormTemplateContext.GE.SetValue("EventType", EventTypeDCB.Value)
            FormTemplateContext.GE.SetValue("JointSponsorSociety", JointSponsorTB.Value)
            FormTemplateContext.GE.SetValue("CertificateLine1", CertificateLine1TB.Value)
            FormTemplateContext.GE.SetValue("CertificateLine2", CertificateLine2TB.Value)
            FormTemplateContext.GE.SetValue("CertificateLine2", CertificateLine3TB.Value)
            FormTemplateContext.GE.SetValue("Paragraph1", Paragraph1CB.Value)
            FormTemplateContext.GE.SetValue("CMECertificate", CMECertificateIdLB.Value)
            FormTemplateContext.GE.SetValue("CECertificate", CECertificateIdLB.Value)
            FormTemplateContext.GE.SetValue("CACertificate", CACertificateIdLB.Value)
            FormTemplateContext.GE.SetValue("DatePrint", DatePrintCB.Value)
            FormTemplateContext.GE.SetValue("LocationPrint", LocationPrintCB.Value)
            FormTemplateContext.GE.SetValue("CertificateVersion", CertificateVersionCB.Value)
            FormTemplateContext.GE.SetValue("EventStartDate", EventStartDateTB.Value)
            FormTemplateContext.GE.SetValue("EventEndDate", EventEndDateTB.Value)
            FormTemplateContext.GE.SetValue("CMEMaxCredits", CMECreditsTB.Value)
            FormTemplateContext.GE.SetValue("SACMEMaxCredits", CMESACreditsTB.Value)
            FormTemplateContext.GE.SetValue("CEMaxCredits", CECreditsTB.Value)
            FormTemplateContext.GE.SetValue("SACEMaxCredits", CESACreditsTB.Value)
            FormTemplateContext.GE.SetValue("CAMaxCredits", CACreditsTB.Value)
            FormTemplateContext.GE.SetValue("StateMandated", StateMandateCB.Value)
            If Not FormTemplateContext.GE.Save(False) Then
                Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                result = "Error"
            Else
                result = "Success"
                FormTemplateContext.GE.Save(True)
                'CourseCreatorAppGE.CommitTransaction()
                'UpdateCourseCreator()
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Function
    Public Function UpdateCourseCreator() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.FormTemplateContext.GE.RecordID
        Try
            If ID < 0 Then
                MsgBox("Please save this record before proceeding")
            Else

                'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
                FormTemplateContext.GE.SetValue("EventName", EventNameTB.Value)
                FormTemplateContext.GE.SetValue("EventProgramId", EventProgramIdLB.Value)
                FormTemplateContext.GE.SetValue("CMEProgram", CMEProgramTB.Value)
                FormTemplateContext.GE.SetValue("CMELocation", CMELocationTB.Value)
                FormTemplateContext.GE.SetValue("NameOrder", CMENameOrderTB.Value)
                FormTemplateContext.GE.SetValue("EventType", EventTypeDCB.Value)
                FormTemplateContext.GE.SetValue("JointSponsorSociety", JointSponsorTB.Value)
                FormTemplateContext.GE.SetValue("CertificateLine1", CertificateLine1TB.Value)
                FormTemplateContext.GE.SetValue("CertificateLine2", CertificateLine2TB.Value)
                FormTemplateContext.GE.SetValue("CertificateLine2", CertificateLine3TB.Value)
                FormTemplateContext.GE.SetValue("Paragraph1", Paragraph1CB.Value)
                FormTemplateContext.GE.SetValue("CMECertificate", CMECertificateIdLB.Value)
                FormTemplateContext.GE.SetValue("CECertificate", CECertificateIdLB.Value)
                FormTemplateContext.GE.SetValue("CACertificate", CACertificateIdLB.Value)
                FormTemplateContext.GE.SetValue("DatePrint", DatePrintCB.Value)
                FormTemplateContext.GE.SetValue("LocationPrint", LocationPrintCB.Value)
                FormTemplateContext.GE.SetValue("CertificateVersion", CertificateVersionCB.Value)
                FormTemplateContext.GE.SetValue("EventStartDate", EventStartDateTB.Value)
                FormTemplateContext.GE.SetValue("EventEndDate", EventEndDateTB.Value)
                FormTemplateContext.GE.SetValue("CMEMaxCredits", CMECreditsTB.Value)
                FormTemplateContext.GE.SetValue("SACMEMaxCredits", CMESACreditsTB.Value)
                FormTemplateContext.GE.SetValue("CEMaxCredits", CECreditsTB.Value)
                FormTemplateContext.GE.SetValue("SACEMaxCredits", CESACreditsTB.Value)
                FormTemplateContext.GE.SetValue("CAMaxCredits", CACreditsTB.Value)
                FormTemplateContext.GE.SetValue("StateMandated", StateMandateCB.Value)


                If Not FormTemplateContext.GE.Save(False) Then
                    Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                    result = "Error"
                Else

                    result = "Success"
                    FormTemplateContext.GE.Save(True)

                End If
                '  End If
            End If
            'If CourseCreatorAppGE.IsDirty Then 'if the ge has changed then save

            '  End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Function


    Public Function DisplayEntity() As String

        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.FormTemplateContext.GE.RecordID
        Try
            If ID > 0 Then

                m_oAppObj.DisplayEntityRecord("acslmscoursecreatorapp", ID)
            End If

            'If CourseCreatorAppGE.IsDirty Then 'if the ge has changed then save

            '  End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Function

End Class

