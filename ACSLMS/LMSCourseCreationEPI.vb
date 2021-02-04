'Option Strict On
'Option Explicit On
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.Application
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls
Imports System.Windows.Forms
Imports Aptify.Framework.FormTemplate

Public Class LMSCourseCreationEPI
    Inherits AptifyGenericEntity
    Private m_oApp As New AptifyApplication
    Private m_oDA As New DataAction
    Private m_oProps As New AptifyProperties
    Private errors As System.Text.StringBuilder = New System.Text.StringBuilder
    Public exception As String = ""
    Dim dt As DataTable
    Dim dt1 As DataTable
    Dim dt2 As DataTable
    Dim result As String = "Failed"
    Dim presult As ProcessFlowResult = Nothing
    Dim lResult As Boolean = False
    Dim CourseGE As AptifyGenericEntityBase
    Dim CourseCreationGE As AptifyGenericEntityBase
    Dim CourseCreationPriceGE As AptifyGenericEntityBase
    Dim EventGE As AptifyGenericEntityBase
    Dim ProductGE As AptifyGenericEntityBase
    Dim ClassGE As AptifyGenericEntityBase
    Dim GLAccountGE As AptifyGenericEntityBase
    Dim ProductGLAccountsGE As AptifyGenericEntityBase
    Dim ProductPriceGE As AptifyGenericEntityBase
    Dim FilterRuleGE As AptifyGenericEntityBase
    Dim FilterRuleItemGE As AptifyGenericEntityBase
    Dim AccountingGLGE As AptifyGenericEntityBase
    Dim CourseCreatorAppGE As AptifyGenericEntityBase
    Dim Status As Integer
    Dim AdminCompletedEthose As Integer
    Dim ID As Long
    Dim CourseCreationCourseId As Long
    Dim CourseName As String
    Dim CourseCreationEventId As Long
    Dim CourseCreationProductId As Long
    Dim CourseCreationProductGL As Long
    Dim lMessageTemplateID As Integer
    Dim productEmailTemplateId As Integer
    Dim lProcessFlowID As Long = -1
    Dim processFlowResult As String = String.Empty
    Dim procFlowSql As String
    Dim Email As String
    Dim ccEmail As String
    Dim SubjectText As String
    Dim HTMLText As String
    Dim InstructorId As Long
    Dim SchoolId As Long
    Dim ProductSalesGLID As String
    Dim newGl As String
    Dim EthosNodeId As Integer
    Dim thisCategoryId As Integer
    Dim emailgroup As String = ""
    Dim ccemailgroup As String = ""
    Dim accountingemailgroup As String = ""
    Dim icemailgroup As String = ""
    Dim courseStatusSQL As String = ""
    Dim courseStatusName As String = ""
    Dim RequestName As String = ""



    Public Overridable ReadOnly Property Application() As AptifyApplication
        Get
            Return m_oApp
        End Get
    End Property
    Public Overrides Function Save(AllowGUI As Boolean, ByRef ErrorString As String, TransactionID As String) As Boolean
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.RecordID

        Status = CInt(Me.GetValue("CourseCreationStatus"))
        courseStatusSQL = "select name from acslmscoursecreatorstatus where id = " & Status
        courseStatusName = CStr(da.ExecuteScalar(courseStatusSQL))

        Try
            If Not ID > 0 Then
                Validate()
                'Include additional Save here
            End If

            Return MyBase.Save(AllowGUI, ErrorString, TransactionID)


        Catch ex As Exception

            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            Return False
        End Try
    End Function

    Public Overrides Function Validate(ByRef ErrorString As String) As Boolean

        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.RecordID
        Status = CInt(Me.GetValue("CourseCreationStatus"))
        AdminCompletedEthose = CInt(Me.GetValue("AdminEthosSetupComplete"))

        Dim ProductName As String = CStr("ProductName")
        Dim SalesGL As String = CStr("ProductName")

        Dim thisMessageTemplateId As Integer
        Dim DueDate As Date = CDate(Me.GetValue("RequestedDueDate"))

        Dim LMSURL As String
        Dim creatoremail As String = CStr(GetValue("ContactEmail"))
        EthosNodeId = CInt(GetValue("EthosNodeId"))

        If Me.DataAction.UserCredentials.Server.ToLower = "aptify" Then
            'production
            thisMessageTemplateId = 1391
            LMSURL = "Https://learning.facs.org"
            thisCategoryId = 48
            emailgroup = "opetinaux@facs.org, ssallan@facs.org,sratsavong@facs.org, dnovak@facs.org" & ", " & CStr(Me.GetValue("ContactEmail"))
            accountingemailgroup = "mfield@facs.org,jbodnar@facs.org"
            icemailgroup = "ahastings@facs.org"
        End If
        If Me.DataAction.UserCredentials.Server.ToLower = "stagingaptify2" Then
            'staging
            thisMessageTemplateId = 1274
            LMSURL = "Https://stage-learning.facs.org"
            thisCategoryId = 47
            emailgroup = "opetinaux@facs.org, ssallan@facs.org,sratsavong@facs.org, dnovak@facs.org" & ", " & CStr(Me.GetValue("ContactEmail"))
            accountingemailgroup = "mfield@facs.org,jbodnar@facs.org"
            icemailgroup = "ahastings@facs.org"
        End If

        If Me.DataAction.UserCredentials.Server.ToLower = "testaptifydb" Then
            'staging
            thisMessageTemplateId = 1266
            LMSURL = "Https://dev-learning.facs.org"
            thisCategoryId = 47
            emailgroup = "dnovak@facs.org"
            accountingemailgroup = "dnovak@facs.org"
            icemailgroup = "dnovak@facs.org"
        End If


        If Me.DataAction.UserCredentials.Server.ToLower = "testaptify610" Then
            'staging
            thisMessageTemplateId = 1266
            LMSURL = "Https://dev-learning.facs.org"
            thisCategoryId = 47
            emailgroup = "dnovak@facs.org"
            accountingemailgroup = "dnovak@facs.org"
            icemailgroup = "dnovak@facs.org"
        End If
        RequestName = CStr(Me.GetValue("RequestedName"))
        Try
            If Status = 1 Then
                lMessageTemplateID = thisMessageTemplateId
                Email = emailgroup
                SubjectText = "Application Created"
                HTMLText = "A new course creation application titled <b>" & RequestName & "</b> has been created please log into " & Me.DataAction.UserCredentials.Server.ToLower & " and create a new course.  The Requested due date is:  " & DueDate & " the request status is: " & courseStatusName
                SendEmail()
            Else
                bResult = True
            End If

            If Status = 2 Then
                If CInt(Me.GetValue("CourseId")) <= 0 Then
                    Select Case MsgBox("Ready to create course?  Proceeding will create a new course and send a message to the Event Creator requesting a new event for this course.  Continue? ", MsgBoxStyle.YesNo, "Course Creator")
                        Case MsgBoxResult.Yes

                            CreateCourseId()
                            'MsgBox("The result is " & result)
                            If result = "Success" Then
                                Me.SetValue("CourseId", CourseGE.RecordID)
                                Me.SetValue("CourseCreationStatus", 3)
                                Me.SetValue("CourseIdCreated", 1)
                                Me.SetValue("CourseIdCreatedDate", Now())
                                lMessageTemplateID = thisMessageTemplateId
                                Email = emailgroup
                                ccEmail = ccemailgroup
                                SubjectText = "Course Created"
                                HTMLText = "A new course has been created please log into " & Me.DataAction.UserCredentials.Server.ToLower & ", go to the course creator app and create an event for course request id: " & Me.RecordID & ", course name: <u>" & CourseName & "</u>.  The Requested due date is:  " & DueDate & " the request status is: " & courseStatusName & " The request name is: <b>" & RequestName & "</b>"


                                SendEmail()

                            End If
                            MsgBox("The process has " & result)
                        Case MsgBoxResult.No
                            Return MyBase.Validate(ErrorString)
                        Case Else
                    End Select
                Else
                    MsgBox("This Request already has a course id, please update the course.  The status will now be Set To Ready To Update Course")
                    Me.SetValue("CourseCreationStatus", 12)
                End If
            Else
                bResult = True
            End If

            If Status = 4 Then
                If CInt(Me.GetValue("EventId")) <= 0 Then
                    Select Case MsgBox("Ready To create Event?  Proceeding will create a New Event And send a message To the Product Creator requesting a New product.  Continue? ", MsgBoxStyle.YesNo, "Event Creator")
                        Case MsgBoxResult.Yes

                            CreateEventId()
                            'MsgBox("The result Is " & result)
                            If result = "Success" Then
                                Me.SetValue("EventId", EventGE.RecordID)
                                Me.SetValue("CourseCreationStatus", 5)
                                Me.SetValue("EventIdCreated", 1)
                                Me.SetValue("EventIdCreatedDate", Now())
                                lMessageTemplateID = thisMessageTemplateId
                                Email = emailgroup
                                ccEmail = ccemailgroup
                                SubjectText = "Event Created"
                                HTMLText = "A New Event has been created please log into " & Me.DataAction.UserCredentials.Server.ToLower & ", go To the course creator app And And create a New product For course request id: " & Me.RecordID & ".  The Requested due date is:  " & DueDate & " the request status is: " & courseStatusName & " The request name is: <b>" & RequestName & "</b>"
                                SendEmail()
                            End If
                            MsgBox("The process has " & result)
                        Case MsgBoxResult.No
                            Return MyBase.Validate(ErrorString)
                        Case Else
                    End Select
                Else
                    MsgBox("This Request already has a event id, please update the event.  The status will now be set to Ready to Update Event")
                    Me.SetValue("CourseCreationStatus", 14)
                End If

            Else
                bResult = True
            End If

            If Status = 6 Then
                If CInt(Me.GetValue("ProductId")) <= 0 Then
                    Select Case MsgBox("Ready to create product?  Proceeding will create a new product and send a message to the GL Creator requesting a new GL for this Product.  Continue? ", MsgBoxStyle.YesNo, "Product Creator")
                        Case MsgBoxResult.Yes

                            CreateProductId()
                            'MsgBox("The result is " & result)
                            If result = "Success" Then
                                Me.SetValue("ProductId", ProductGE.RecordID)
                                Me.SetValue("CourseCreationStatus", 7)
                                Me.SetValue("ProductIdCreated", 1)
                                Me.SetValue("ProductIdCreatedDate", Now())
                                lMessageTemplateID = thisMessageTemplateId
                                Email = emailgroup
                                ccEmail = ccemailgroup
                                SubjectText = "Product Created"
                                HTMLText = "A new product has been created please log into " & Me.DataAction.UserCredentials.Server.ToLower & ", go to the course creator app and create a new GL for the product for course request id: " & Me.RecordID & ".  The Requested due date is:  " & DueDate & " the request status is: " & courseStatusName & " The request name is: <b>" & RequestName & "</b>"
                                SendEmail()
                            End If
                            MsgBox("The process has " & result)

                        Case MsgBoxResult.No
                            Return MyBase.Validate(ErrorString)
                        Case Else
                    End Select
                Else

                    MsgBox("This Request already has a product id, please update the product.  The status will now be set to Ready to Update Product")
                    Me.SetValue("CourseCreationStatus", 16)
                End If


            Else
                bResult = True
            End If

            If Status = 8 AndAlso CInt(Me.GetValue("ProductId")) > 0 Then
                Select Case MsgBox("Ready to create GL?  Proceeding will create a new GL for the Product and replace the sales GL currently assigned.  Continue? ", MsgBoxStyle.YesNo, "GL Creator")
                    Case MsgBoxResult.Yes

                        CreateCourseGL()
                        'MsgBox("The result is " & result)
                        If result = "Success" Then
                            Me.SetValue("CourseCreationStatus", 9)
                            Me.SetValue("SalesGL", newGl)
                            Me.SetValue("GLCreated", 1)
                            Me.SetValue("GLCreatedDate", Now())
                            lMessageTemplateID = thisMessageTemplateId
                            Email = emailgroup
                            ccEmail = ccemailgroup
                            SubjectText = "GL Setup Complete"
                            HTMLText = "The GL setup has been completed.  When ready please log into " & Me.DataAction.UserCredentials.Server.ToLower & " and complete the course setup for course creator request id: " & Me.RecordID & ".  The Requested due date is:  " & DueDate & " the request status is: " & courseStatusName & " The request name is: <b>" & RequestName & "</b>"
                            SendEmail()
                            Select Case MsgBox("This will send an email to the accounting department to include the new product GL in the Nav?  Continue? ", MsgBoxStyle.YesNo, "Product Creator")
                                Case MsgBoxResult.Yes

                                    'If result = "Success" Then
                                    'Me.SetValue("CourseCreationStatus", "Setup Complete")
                                    lMessageTemplateID = thisMessageTemplateId
                                    Email = accountingemailgroup
                                    SubjectText = "New Product"
                                    HTMLText = "The course setup has been completed.&nbsp; Please include the following new product in the Nav:</p>
                                                <p>
                                                <br />ProductName<br />:" & CStr(Me.GetValue("ProductName")) & "<br></br>" &
                                                   "SalesGL:  " & CStr(Me.GetValue("SalesGL"))

                                    SendEmail()
                        'End If
                                Case MsgBoxResult.No
                                    Return MyBase.Validate(ErrorString)
                                Case Else
                            End Select
                        End If
                        MsgBox("The process has " & result)
                    Case MsgBoxResult.No
                        Return MyBase.Validate(ErrorString)
                    Case Else
                End Select

            Else
                bResult = True
            End If

            If Status = 10 AndAlso CInt(Me.GetValue("ProductId")) > 0 Then
                Select Case MsgBox("Ready to complete course setup?  Proceeding will complete the course setup and create a new course in the LMS.  Continue? ", MsgBoxStyle.YesNo, "Product Creator")
                    Case MsgBoxResult.Yes

                        CompleteCourseSetup()
                        'MsgBox("The result is " & result)
                        If result = "Success" Then
                            Me.SetValue("CourseCreationStatus", 11)
                            Me.SetValue("CourseSetupComplete", 1)
                            Me.SetValue("CourseSetupCompleteDate", Now())
                            lMessageTemplateID = thisMessageTemplateId
                            Email = emailgroup & "," & creatoremail & "," & icemailgroup
                            ccEmail = ccemailgroup
                            SubjectText = "Course Setup Complete"
                            HTMLText = "The course setup for <b>" & RequestName & "</b> has been completed.  Please log in to the LMS to view the new course <a href='" & LMSURL & "'>Go to the LMS<a>"
                            SendEmail()

                        End If
                    Case MsgBoxResult.No
                        Return MyBase.Validate(ErrorString)
                    Case Else
                End Select

            Else
                bResult = True
            End If

            If Status = 12 AndAlso CInt(Me.GetValue("CourseId")) > 0 Then

                Select Case MsgBox("This will save the course information you have updated.  Continue? ", MsgBoxStyle.YesNo, "Course Creator")
                    Case MsgBoxResult.Yes
                        UpdateCourseId()
                        If result = "Success" Then
                            'Me.SetValue("CourseCreationStatus", "Setup Complete")
                            Me.SetValue("CourseCreationStatus", 13)
                        End If
                        MsgBox("The process has " & result)
                    Case MsgBoxResult.No
                        Return MyBase.Validate(ErrorString)
                    Case Else
                End Select

            Else
                bResult = True
            End If
            'If Status = 13 Then
            '    m_oApp.DisplayEntityRecord("ACSLMSCourseCreatorApp", RecordID)
            'End If

            If Status = 14 AndAlso CInt(Me.GetValue("EventId")) > 0 Then
                Select Case MsgBox("This will save the event information you have updated.  Continue? ", MsgBoxStyle.YesNo, "Event Creator")
                    Case MsgBoxResult.Yes
                        UpdateEventId()
                        If result = "Success" Then
                            'Me.SetValue("CourseCreationStatus", "Setup Complete")
                            Me.SetValue("CourseCreationStatus", 15)
                        End If
                        MsgBox("The process has " & result)
                    Case MsgBoxResult.No
                        Return MyBase.Validate(ErrorString)
                    Case Else
                End Select

            Else
                bResult = True
            End If

            If Status = 16 AndAlso CInt(Me.GetValue("ProductId")) > 0 Then
                Select Case MsgBox("This will save the product information you have updated.  Continue? ", MsgBoxStyle.YesNo, "Product Creator")
                    Case MsgBoxResult.Yes
                        UpdateProductId()
                        If result = "Success" Then
                            'Me.SetValue("CourseCreationStatus", "Setup Complete")
                            Me.SetValue("CourseCreationStatus", 17)
                        End If
                        MsgBox("The process has " & result)
                    Case MsgBoxResult.No
                        Return MyBase.Validate(ErrorString)
                    Case Else
                End Select

            Else
                bResult = True
            End If

            If Me.GetField("CourseAdminEthosSetupComplete").IsDirty And CInt(Me.GetValue("CourseAdminEthosSetupComplete")) = 1 And Status >= 11 Then
                Select Case MsgBox("This send an email to IC to have your course QA checked.  Continue? ", MsgBoxStyle.YesNo, "Course Creator")
                    Case MsgBoxResult.Yes
                        Me.SetValue("CourseAdminSetupCompleteDate", Now())
                        lMessageTemplateID = thisMessageTemplateId
                        Email = icemailgroup
                        SubjectText = "Course Ready for QA"
                        HTMLText = "The course setup has been completed in Ethose.&nbsp; 
                            <br />Course Name<br />:" & CStr(Me.GetValue("CourseName")) & "<br> 
                            Please log in to the LMS to view the new course <a href='" & LMSURL & "'>Go to the LMS<a>"
                        SendEmail()
                    Case MsgBoxResult.No
                        Return MyBase.Validate(ErrorString)
                    Case Else
                End Select

            Else
                bResult = True
            End If

            If bResult = True Then
                Return MyBase.Validate(ErrorString)
            End If

        Catch ex As Exception
            result = "Error"
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            Return False
        End Try
    End Function


    Public Function CreateCourseId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.RecordID
        Status = CInt(Me.GetValue("CourseCreationStatus"))
        CourseCreationCourseId = CLng(Me.GetValue("CourseId"))

        InstructorId = CLng(Me.GetValue("CourseInstructorId"))
        SchoolId = CLng(Me.GetValue("CourseSchoolId"))
        Dim getInstructorSequenceSql As String
        Dim lInstructorSequenceId As Long
        Dim getSchoolSequenceSql As String
        Dim lSchoolSequenceId As Long

        Try


            CourseGE = m_oApp.GetEntityObject("Courses", -1)
            If CourseCreationCourseId <= 0 Then
                CourseGE.SetValue("Name", Me.GetValue("CourseName"))
                If Len(Me.GetValue("CourseDescription")) > 255 Then
                    result = "Failed"
                    CourseGE.SetValue("Description", Me.GetValue("CourseDescription"))
                    MsgBox("The Course Description must be 255 characters or less.  The Course Description is: " & Len(Me.GetValue("CourseDescription")) & " characters")
                Else

                    CourseGE.SetValue("Description", Me.GetValue("CourseDescription"))
                End If
                CourseGE.SetValue("CategoryID", Me.GetValue("CourseCategoryId"))
                If CInt(Me.GetValue("IsBundledProduct")) = 1 Then
                    CourseGE.SetValue("ProductTypeID", 5)
                Else
                    CourseGE.SetValue("ProductTypeID", 7)
                End If

                CourseGE.SetValue("Status", "Available")
                CourseGE.SetValue("Units", 0)
                If Not Me.GetValue("EthosNodeId") Is Nothing Then
                    CourseGE.SetValue("acsExternalId", EthosNodeId)
                End If
                CourseName = CStr(Me.GetValue("CourseName"))
                If CourseGE.IsDirty Then 'if the ge has changed then save
                    If Not CourseGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & CourseGE.RecordID)
                        result = "Error"
                    Else
                        CourseGE.Save(True)
                        result = "Success"

                    End If
                End If
                getInstructorSequenceSql = "select case when (select max(ci.Sequence) from courseinstructor ci where courseid = " & CourseGE.RecordID & " ) is null then 1 else  (select max(ci.Sequence) from courseinstructor ci where courseid = " & CourseGE.RecordID & ") + 1 end"
                lInstructorSequenceId = CLng(da.ExecuteScalar(getInstructorSequenceSql))

                Dim courseInstructorGE As AptifyGenericEntityBase = m_oApp.GetEntityObject("courseinstructors", -1)

                courseInstructorGE.SetValue("CourseID", CourseGE.RecordID)
                courseInstructorGE.SetValue("Sequence", lInstructorSequenceId)
                courseInstructorGE.SetValue("InstructorID", InstructorId)
                courseInstructorGE.SetValue("Status", "Active")
                If courseInstructorGE.IsDirty Then
                    If Not courseInstructorGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & courseInstructorGE.RecordID)
                        result = "Error"
                    Else
                        courseInstructorGE.Save(True)
                        result = "Success"
                    End If

                End If
                getSchoolSequenceSql = "select case when (select max(cs.Sequence) from courseschool cs where courseid = " & CourseGE.RecordID & " ) is null then 1 else  (select max(cs.Sequence) from courseschool cs where courseid = " & CourseGE.RecordID & ") + 1 end"
                lSchoolSequenceId = CLng(da.ExecuteScalar(getSchoolSequenceSql))

                Dim courseSchoolGE As AptifyGenericEntityBase = m_oApp.GetEntityObject("courseschools", -1)

                courseSchoolGE.SetValue("CourseID", CourseGE.RecordID)
                courseSchoolGE.SetValue("Sequence", lSchoolSequenceId)
                courseSchoolGE.SetValue("SchoolID", SchoolId)
                courseSchoolGE.SetValue("Status", "Active")
                courseSchoolGE.SetValue("Rank", 1)
                If courseSchoolGE.IsDirty Then
                    If Not courseSchoolGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & courseSchoolGE.RecordID)
                        result = "Error"
                    Else
                        courseSchoolGE.Save(True)
                        result = "Success"
                    End If

                End If


            End If



        Catch ex As Exception

        End Try

    End Function
    Public Function UpdateCourseId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.RecordID
        Status = CInt(Me.GetValue("CourseCreationStatus"))
        CourseCreationCourseId = CLng(Me.GetValue("CourseId"))
        InstructorId = CLng(Me.GetValue("CourseInstructorId"))
        SchoolId = CLng(Me.GetValue("CourseSchoolId"))
        Dim getInstructorRecordSql As String
        Dim lInstructorRecordId As Long
        Dim getSchoolRecordSql As String
        Dim lSchoolRecordId As Long
        Try


            CourseGE = m_oApp.GetEntityObject("Courses", CourseCreationCourseId)
            If CourseCreationCourseId > 0 Then
                CourseGE.SetValue("Name", Me.GetValue("CourseName"))
                CourseGE.SetValue("Description", Me.GetValue("CourseDescription"))
                CourseGE.SetValue("CategoryID", Me.GetValue("CourseCategoryId"))
                If CInt(Me.GetValue("IsBundledProduct")) = 1 Then
                    CourseGE.SetValue("ProductTypeID", 5)
                Else
                    CourseGE.SetValue("ProductTypeID", 7)
                End If
                CourseGE.SetValue("Status", "Available")
                CourseGE.SetValue("Units", 0)
                If Not Me.GetValue("EthosNodeId") Is Nothing Then
                    CourseGE.SetValue("acsExternalId", EthosNodeId)
                End If
                If CourseGE.IsDirty Then 'if the ge has changed then save
                    If Not CourseGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & CourseGE.RecordID)
                        result = "Error"
                    Else
                        CourseGE.Save(True)
                        result = "Success"


                    End If
                End If
                getInstructorRecordSql = "select id from courseinstructor ci where courseid = " & CourseGE.RecordID
                lInstructorRecordId = CLng(da.ExecuteScalar(getInstructorRecordSql))

                Dim courseInstructorGE As AptifyGenericEntityBase = m_oApp.GetEntityObject("courseinstructors", lInstructorRecordId)
                courseInstructorGE.SetValue("InstructorID", InstructorId)
                courseInstructorGE.SetValue("Status", "Active")
                If courseInstructorGE.IsDirty Then
                    If Not courseInstructorGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & courseInstructorGE.RecordID)
                        result = "Error"
                    Else
                        courseInstructorGE.Save(True)
                        result = "Success"
                    End If

                End If
                getSchoolRecordSql = "select id from courseschool cs where courseid = " & CourseGE.RecordID
                lSchoolRecordId = CLng(da.ExecuteScalar(getSchoolRecordSql))

                Dim courseSchoolGE As AptifyGenericEntityBase = m_oApp.GetEntityObject("courseschools", lSchoolRecordId)
                courseSchoolGE.SetValue("SchoolID", SchoolId)
                courseSchoolGE.SetValue("Status", "Active")
                courseSchoolGE.SetValue("Rank", 1)
                If courseSchoolGE.IsDirty Then
                    If Not courseSchoolGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & courseSchoolGE.RecordID)
                        result = "Error"
                    Else
                        courseSchoolGE.Save(True)
                        result = "Success"
                    End If

                End If


            End If

        Catch ex As Exception

        End Try

    End Function

    Public Function CreateEventId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim ID As Long = Me.RecordID
        CourseCreationEventId = CLng(Me.GetValue("EventId"))
        Try


            EventGE = m_oApp.GetEntityObject("ACSCMEEvent", -1)
            If CourseCreationEventId <= 0 Then
                If Len(Me.GetValue("EventName")) > 50 Then
                    result = "Failed"
                    MsgBox("The Event name must be 50 characters or less.  The Event Name is: " & Len(Me.GetValue("EventName")) & " characters")
                Else
                    EventGE.SetValue("Name", Me.GetValue("EventName"))
                End If

                EventGE.SetValue("ProgramID", Me.GetValue("EventProgramId"))
                EventGE.SetValue("CME_Program", Me.GetValue("CMEProgram"))
                EventGE.SetValue("CME_Location", Me.GetValue("CMELocation"))
                If Len(Me.GetValue("NameOrder")) > 50 Then
                    result = "Failed"
                    MsgBox("The Name Order name must be 50 characters or less.  The Name Order is: " & Len(Me.GetValue("NameOrder")) & " characters")
                Else
                    EventGE.SetValue("NameOrder", Me.GetValue("NameOrder"))
                End If

                EventGE.SetValue("CME_Start_Date", Me.GetValue("CMEStartDate"))
                EventGE.SetValue("CME_End_Date", Me.GetValue("CMEEndDate"))
                EventGE.SetValue("CME_Max_Credits", Me.GetValue("CMEMaxCredits"))
                EventGE.SetValue("SACME_Max_Credits", Me.GetValue("SACMEMaxCredits"))
                EventGE.SetValue("CE_Max_Credits", Me.GetValue("CEMaxCredits"))
                EventGE.SetValue("SACE_Max_Credits", Me.GetValue("SACEMaxCredits"))
                EventGE.SetValue("CA_Max_Credits", Me.GetValue("CAMaxCredits"))
                EventGE.SetValue("AwardStateMandated", Me.GetValue("StateMandated"))
                EventGE.SetValue("CertLine1", Me.GetValue("CertificateLine1"))
                EventGE.SetValue("CertLine2", Me.GetValue("CertificateLine2"))
                EventGE.SetValue("CertLine3", Me.GetValue("CertificateLine3"))
                EventGE.SetValue("Paragraph1", Me.GetValue("Paragraph1"))
                EventGE.SetValue("DatePrint", Me.GetValue("DatePrint"))
                EventGE.SetValue("LocationPrint", Me.GetValue("LocationPrint"))
                EventGE.SetValue("EventType", Me.GetValue("EventType"))
                EventGE.SetValue("jspsociety", Me.GetValue("JointSponsorSociety"))
                EventGE.SetValue("ACSCmeCertTemplate_Id", Me.GetValue("CMECertificate"))
                EventGE.SetValue("ACSCMECertTemplate_IDNonMD", Me.GetValue("CACertificate"))
                EventGE.SetValue("ACSCECertTemplateId", Me.GetValue("CECertificate"))
                EventGE.SetValue("CertificateVersion", Me.GetValue("CertificateVersion"))

                If EventGE.IsDirty Then 'if the ge has changed then save
                    If Not EventGE.Save(False) Then
                        Throw New Exception("Problem Saving Event Record:" & EventGE.RecordID)
                        result = "Error"
                    Else
                        EventGE.Save(True)
                        result = "Success"

                    End If
                End If
            End If

        Catch ex As Exception

        End Try

    End Function
    Public Function UpdateEventId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim ID As Long = Me.RecordID
        CourseCreationEventId = CLng(Me.GetValue("EventId"))
        Try


            EventGE = m_oApp.GetEntityObject("ACSCMEEvent", CourseCreationEventId)
            If CourseCreationEventId > 0 Then
                If Len(Me.GetValue("EventName")) > 50 Then
                    result = "Failed"
                    MsgBox("The Event name must be 50 characters or less.  The Event Name is: " & Len(Me.GetValue("EventName")))

                Else
                    EventGE.SetValue("Name", Me.GetValue("EventName"))
                End If
                EventGE.SetValue("ProgramID", Me.GetValue("EventProgramId"))
                EventGE.SetValue("CME_Program", Me.GetValue("CMEProgram"))
                EventGE.SetValue("CME_Location", Me.GetValue("CMELocation"))
                If Len(Me.GetValue("NameOrder")) > 50 Then
                    result = "Failed"
                    MsgBox("The Name Order name must be 50 characters or less.  The Name Order is: " & Len(Me.GetValue("NameOrder")))
                Else
                    EventGE.SetValue("NameOrder", Me.GetValue("NameOrder"))
                End If
                EventGE.SetValue("CME_Start_Date", Me.GetValue("CMEStartDate"))
                EventGE.SetValue("CME_End_Date", Me.GetValue("CMEEndDate"))
                EventGE.SetValue("CME_Max_Credits", Me.GetValue("CMEMaxCredits"))
                EventGE.SetValue("SACME_Max_Credits", Me.GetValue("SACMEMaxCredits"))
                EventGE.SetValue("CE_Max_Credits", Me.GetValue("CEMaxCredits"))
                EventGE.SetValue("SACE_Max_Credits", Me.GetValue("SACEMaxCredits"))
                EventGE.SetValue("CA_Max_Credits", Me.GetValue("CAMaxCredits"))
                EventGE.SetValue("AwardStateMandated", Me.GetValue("StateMandated"))
                EventGE.SetValue("CertLine1", Me.GetValue("CertificateLine1"))
                EventGE.SetValue("CertLine2", Me.GetValue("CertificateLine2"))
                EventGE.SetValue("CertLine3", Me.GetValue("CertificateLine3"))
                EventGE.SetValue("Paragraph1", Me.GetValue("Paragraph1"))
                EventGE.SetValue("DatePrint", Me.GetValue("DatePrint"))
                EventGE.SetValue("LocationPrint", Me.GetValue("LocationPrint"))
                EventGE.SetValue("EventType", Me.GetValue("EventType"))
                EventGE.SetValue("jspsociety", Me.GetValue("JointSponsorSociety"))
                EventGE.SetValue("ACSCmeCertTemplate_Id", Me.GetValue("CMECertificate"))
                EventGE.SetValue("ACSCMECertTemplate_IDNonMD", Me.GetValue("CACertificate"))
                EventGE.SetValue("ACSCECertTemplateId", Me.GetValue("CECertificate"))
                EventGE.SetValue("CertificateVersion", Me.GetValue("CertificateVersion"))

                If EventGE.IsDirty Then 'if the ge has changed then save
                    If Not EventGE.Save(False) Then
                        Throw New Exception("Problem Saving Event Record:" & EventGE.RecordID)
                        result = "Failed"
                    Else
                        EventGE.Save(True)
                        result = "Success"

                    End If
                End If
            End If

        Catch ex As Exception

        End Try

    End Function

    Public Function CreateProductId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim ID As Long = Me.RecordID

        Dim sSQL As String
        Dim getProductPriceSequenceSql As String
        Dim lProductPriceSequenceId As Long
        Dim getFilterRuleItemSequenceSql As String
        Dim lFilterRuleItemSequenceId As Long
        Dim getAccountingGLItemSequenceSql As String
        Dim lAccountingGLItemSequenceSql As Long
        CourseCreationProductId = CLng(Me.GetValue("ProductId"))
        Dim CourseCreationProductName As String = CStr(Me.GetValue("ProductName"))
        Try

            CourseCreationGE = m_oApp.GetEntityObject("ACSLMSCourseCreatorApp", Me.RecordID)

            'If CourseCreationGE.IsDirty Then 'if the ge has changed then save
            'If Not CourseCreationGE.Save(False) Then
            'Throw New Exception("Problem Saving Product Record:" & CourseCreationGE.RecordID)
            '     result = "Error"
            'Else
            CourseCreationGE.Save(True)
            result = "Success"


            'End If
            'End If
        Catch ex As Exception

        End Try


        Try

            ProductGE = m_oApp.GetEntityObject("Products", -1)
            If CourseCreationProductId <= 0 Then
                ProductGE.SetValue("Name", CStr(Me.GetValue("ProductName")))
                ProductGE.SetValue("CategoryID", CInt(thisCategoryId))
                If CInt(Me.GetValue("IsBundledProduct")) = 1 Then
                    ProductGE.SetValue("ProductTypeID", 5)
                Else
                    ProductGE.SetValue("ProductTypeID", 7)
                End If

                ProductGE.SetValue("ACSHasCustomMessage", 1)
                ProductGE.SetValue("ACSMessageTemplateID", CInt(Me.GetValue("ProductEmailTemplateId")))
                ProductGE.SetValue("ACSNoStandardOrderConfirmMessage", 1)
                If ProductGE.IsDirty Then 'if the ge has changed then save
                    If Not ProductGE.Save(False) Then
                        Throw New Exception("Problem Saving Product Record:" & ProductGE.RecordID)
                        result = "Error"
                    Else
                        ProductGE.Save(True)
                        result = "Success"

                    End If
                End If


                sSQL = "select * from ACSLMSCourseCreatorProductPrices where ACSLMSCourseCreatorAppID = " & Me.RecordID
                dt = m_oDA.GetDataTable(sSQL)
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows

                        getProductPriceSequenceSql = "select case when (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & " ) is null then 1 else  (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & ") + 1 end"
                        lProductPriceSequenceId = CLng(da.ExecuteScalar(getProductPriceSequenceSql))

                        ProductPriceGE = m_oApp.GetEntityObject("ProductPrices", -1)
                        If ProductGE.RecordID > 0 Then
                            ProductPriceGE.SetValue("ProductId", ProductGE.RecordID)
                            ProductPriceGE.SetValue("Name", dr.Item("ProductPriceName"))
                            ProductPriceGE.SetValue("MemberTypeID", dr.Item("ProductFilterMemberType"))
                            ProductPriceGE.SetValue("Price", dr.Item("ProductFilterRulePrice"))
                            ProductPriceGE.SetValue("CurrencyTypeID", 2)
                            ProductPriceGE.SetValue("Sequence", lProductPriceSequenceId)
                        End If

                        If CStr(dr.Item("ProductFilterRule")) = "ACSMemberClass" Then
                            ProductPriceGE.SetValue("ApplyFilterRule", 1)
                        End If

                        If ProductPriceGE.IsDirty Then 'if the ge has changed then save
                            If Not ProductPriceGE.Save(False) Then
                                Throw New Exception("Problem Saving Product Price Record:" & ProductPriceGE.RecordID)
                                result = "Error"
                            Else
                                ProductPriceGE.Save(True)
                                result = "Success"

                            End If
                        End If

                        FilterRuleGE = m_oApp.GetEntityObject("Filter Rules", -1)

                        FilterRuleGE.SetValue("Name", "Price: " & CStr(dr.Item("ProductPriceName")))
                        'FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(Me.GetValue("ProductName")) & ") , Price Name: " & CStr(dr.Item("ProductPriceName")) & ")")
                        FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(dr.Item("ProductPriceName")) & ")")
                        FilterRuleGE.SetValue("FilterRuleTypeID", 1)
                        FilterRuleGE.SetValue("LogicString", 1)
                        FilterRuleGE.SetValue("LogicStringGenerated", 0)

                        getFilterRuleItemSequenceSql = "select case when (select max(fri.Sequence) from FilterRuleItem fri where FilterRuleId = (select max(filterruleid) from Productprice where productid = " & ProductGE.RecordID & ") ) is null then 1 else  (select max(fri.Sequence) from FilterRuleItem fri where FilterRuleid = (select max(filterruleid) from Productprice where productid = " & ProductGE.RecordID & ")) + 1 end"
                        lFilterRuleItemSequenceId = CLng(da.ExecuteScalar(getFilterRuleItemSequenceSql))

                        With FilterRuleGE.SubTypes("FilterRuleItems").Add()
                            .SetValue("Sequence", lFilterRuleItemSequenceId)
                            .SetValue("ItemNumber", lFilterRuleItemSequenceId)
                            .SetValue("EntityInstance", "Bill To Person")
                            .SetValue("Field", "ACSMemberClassID")
                            .SetValue("Value", dr.Item("ProductFilterRuleValue"))
                            .SetValue("OperatorID", 1)
                        End With
                        If FilterRuleGE.IsDirty Then 'if the ge has changed then save
                            If Not FilterRuleGE.Save(False) Then
                                Throw New Exception("Problem Saving Filter Rule Record:" & FilterRuleGE.RecordID)
                                result = "Error"
                            Else
                                FilterRuleGE.Save(True)
                                result = "Success"

                            End If
                        End If


                        ProductPriceGE = m_oApp.GetEntityObject("ProductPrices", ProductPriceGE.RecordID)
                        ProductPriceGE.SetValue("FilterRuleID", FilterRuleGE.RecordID)
                        If ProductPriceGE.IsDirty Then 'if the ge has changed then save
                            If Not ProductPriceGE.Save(False) Then
                                Throw New Exception("Problem Saving Product Price Record:" & ProductPriceGE.RecordID)
                                result = "Error"
                            Else
                                ProductPriceGE.Save(True)
                                result = "Success"

                            End If
                        End If

                    Next

                End If

            End If


        Catch ex As Exception

        End Try

    End Function
    Public Function UpdateProductId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim ID As Long = Me.RecordID
        Dim sSQL As String
        Dim sSQL1 As String
        Dim sSQL2 As String
        Dim lProductPriceId As Long
        Dim lProductPriceRecordId As Long
        Dim getProductPriceSequenceSql As String
        Dim getCurrentPriceFilterRecordSql As String
        Dim lCurrentPriceFilterRecordId As Long
        Dim getFilterRuleRecordSql As String
        Dim lFilterRuleRecordId As Long
        Dim lProductPriceSequence As Long
        Dim lProductPriceSequenceId As Long
        Dim getFilterRuleItemSequenceSql As String
        Dim lFilterRuleItemSequenceId As Long
        CourseCreationProductId = CLng(Me.GetValue("ProductId"))
        Dim getFilterRuleItemRecordSql As String
        Dim lFilterRuleItemRecordId As Long
        Dim CourseCreationProductName As String = CStr(Me.GetValue("ProductName"))
        Try
            CourseCreationGE = m_oApp.GetEntityObject("ACSLMSCourseCreatorApp", Me.RecordID)
            'If CourseCreationGE.IsDirty Then 'if the ge has changed then save
            'If Not CourseCreationGE.Save(False) Then
            'Throw New Exception("Problem Saving Product Record:" & CourseCreationGE.RecordID)
            'result = "Error"
            'Else
            CourseCreationGE.Save(True)
            result = "Success"

            'End If
            'End If
        Catch ex As Exception

        End Try

        Try

            ProductGE = m_oApp.GetEntityObject("Products", CourseCreationProductId)
            If CourseCreationProductId > 0 Then
                ProductGE.SetValue("Name", Me.GetValue("ProductName"))
                ProductGE.SetValue("CategoryID", thisCategoryId)
                ProductGE.SetValue("ProductTypeID", 7)
                ProductGE.SetValue("ACSHasCustomMessage", 1)
                ProductGE.SetValue("ACSMessageTemplateID", Me.GetValue("ProductEmailTemplateId"))
                ProductGE.SetValue("ACSNoStandardOrderConfirmMessage", 1)
                If ProductGE.IsDirty Then 'if the ge has changed then save
                    If Not ProductGE.Save(False) Then
                        Throw New Exception("Problem Saving Product Record:" & ProductGE.RecordID)
                        result = "Error"
                    Else
                        ProductGE.Save(True)
                        result = "Success"

                    End If
                End If

                'Get all the prices from the app pricing table
                sSQL = "select * from ACSLMSCourseCreatorProductPrices where ACSLMSCourseCreatorAppID = " & Me.RecordID
                dt = m_oDA.GetDataTable(sSQL)

                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        lProductPriceSequence = CLng(dr.Item("sequence"))
                        getCurrentPriceFilterRecordSql = "select ID from productprice pp where productid = " & CourseCreationProductId & " and Sequence = " & lProductPriceSequence
                        lCurrentPriceFilterRecordId = CLng(da.ExecuteScalar(getCurrentPriceFilterRecordSql))

                        If lCurrentPriceFilterRecordId > 0 Then 'Updating existing prices and prices filters

                            ProductPriceGE = m_oApp.GetEntityObject("ProductPrices", lCurrentPriceFilterRecordId)
                            ProductPriceGE.SetValue("ProductId", ProductGE.RecordID)
                            ProductPriceGE.SetValue("Name", dr.Item("ProductPriceName"))
                            ProductPriceGE.SetValue("MemberTypeID", dr.Item("ProductFilterMemberType"))
                            ProductPriceGE.SetValue("Price", dr.Item("ProductFilterRulePrice"))
                            ProductPriceGE.SetValue("CurrencyTypeID", 2)

                        Else
                            getProductPriceSequenceSql = "select case when (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & " ) is null then 1 else  (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & ") + 1 end"
                            lProductPriceSequenceId = CLng(da.ExecuteScalar(getProductPriceSequenceSql))

                            ProductPriceGE = m_oApp.GetEntityObject("ProductPrices", -1) 'Creating new prices and price filters.
                            ProductPriceGE.SetValue("ProductId", ProductGE.RecordID)
                            ProductPriceGE.SetValue("Name", dr.Item("ProductPriceName"))
                            ProductPriceGE.SetValue("MemberTypeID", dr.Item("ProductFilterMemberType"))
                            ProductPriceGE.SetValue("Price", dr.Item("ProductFilterRulePrice"))
                            ProductPriceGE.SetValue("CurrencyTypeID", 2)
                            ProductPriceGE.SetValue("Sequence", lProductPriceSequenceId)


                        End If


                        If CStr(dr.Item("ProductFilterRule")) = "ACSMemberClass" Then
                            ProductPriceGE.SetValue("ApplyFilterRule", 1)
                        End If

                        If CLng(ProductPriceGE.GetValue("ApplyFilterRule")) = 1 Then
                            getFilterRuleRecordSql = "select ID from FilterRule fr where ID in (select filterruleid from Productprice where productid = " & ProductGE.RecordID & ") "
                            lFilterRuleRecordId = CLng(da.ExecuteScalar(getFilterRuleRecordSql))
                            'dt1 = m_oDA.GetDataTable(getFilterRuleRecordSql)
                            If lFilterRuleRecordId > 0 Then
                                FilterRuleGE = m_oApp.GetEntityObject("Filter Rules", lFilterRuleRecordId)

                                FilterRuleGE.SetValue("Name", "Price: " & CStr(dr.Item("ProductPriceName")))
                                'FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(Me.GetValue("ProductName")) & ") , Price Name: " & CStr(dr.Item("ProductPriceName")) & ")")
                                FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(dr.Item("ProductPriceName")) & ")")
                                FilterRuleGE.SetValue("FilterRuleTypeID", 1)
                                FilterRuleGE.SetValue("LogicString", 1)
                                FilterRuleGE.SetValue("LogicStringGenerated", 0)
                                If FilterRuleGE.IsDirty Then 'if the ge has changed then save
                                    If Not FilterRuleGE.Save(False) Then
                                        Throw New Exception("Problem Saving Filter Rule Record:" & FilterRuleGE.RecordID)
                                        result = "Error"
                                    Else
                                        FilterRuleGE.Save(True)
                                        result = "Success"

                                        ProductPriceGE.SetValue("FilterRuleId", FilterRuleGE.RecordID)

                                    End If
                                End If

                                getFilterRuleItemRecordSql = "select ID from FilterRuleItem fri where FilterRuleId = " & FilterRuleGE.RecordID & " and Sequence = " & lProductPriceSequence
                                lFilterRuleItemRecordId = CLng(da.ExecuteScalar(getFilterRuleItemRecordSql))
                                If lFilterRuleItemRecordId > 0 Then
                                    FilterRuleItemGE = m_oApp.GetEntityObject("FilterRuleItems", lFilterRuleItemRecordId)
                                    FilterRuleItemGE.SetValue("EntityInstance", "Bill To Person")
                                    FilterRuleItemGE.SetValue("Field", "ACSMemberClassID")
                                    FilterRuleItemGE.SetValue("Value", dr.Item("ProductFilterRuleValue"))
                                    FilterRuleItemGE.SetValue("OperatorID", 1)
                                    If FilterRuleItemGE.IsDirty Then
                                        Throw New Exception("Problem Saving Filter Rule Record:" & FilterRuleItemGE.RecordID)
                                        result = "Error"
                                    Else
                                        FilterRuleItemGE.Save(True)
                                        result = "Success"
                                    End If
                                Else
                                    FilterRuleGE = m_oApp.GetEntityObject("Filter Rules", -1)

                                    FilterRuleGE.SetValue("Name", "Price: " & CStr(dr.Item("ProductPriceName")))
                                    'FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(Me.GetValue("ProductName")) & ") , Price Name: " & CStr(dr.Item("ProductPriceName")) & ")")
                                    FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(dr.Item("ProductPriceName")) & ")")
                                    FilterRuleGE.SetValue("FilterRuleTypeID", 1)
                                    FilterRuleGE.SetValue("LogicString", 1)
                                    FilterRuleGE.SetValue("LogicStringGenerated", 0)

                                    getFilterRuleItemSequenceSql = "select case when (select max(fri.Sequence) from FilterRuleItem fri where FilterRuleId = (select max(filterruleid) from Productprice where productid = " & ProductGE.RecordID & ") ) is null then 1 else  (select max(fri.Sequence) from FilterRuleItem fri where FilterRuleid = (select max(filterruleid) from Productprice where productid = " & ProductGE.RecordID & ")) + 1 end"
                                    lFilterRuleItemSequenceId = CLng(da.ExecuteScalar(getFilterRuleItemSequenceSql))

                                    With FilterRuleGE.SubTypes("FilterRuleItems").Add()
                                        .SetValue("Sequence", lFilterRuleItemSequenceId)
                                        .SetValue("ItemNumber", lFilterRuleItemSequenceId)
                                        .SetValue("EntityInstance", "Bill To Person")
                                        .SetValue("Field", "ACSMemberClassID")
                                        .SetValue("Value", dr.Item("ProductFilterRuleValue"))
                                        .SetValue("OperatorID", 1)
                                    End With
                                    If FilterRuleGE.IsDirty Then 'if the ge has changed then save
                                        If Not FilterRuleGE.Save(False) Then
                                            Throw New Exception("Problem Saving Filter Rule Record:" & FilterRuleGE.RecordID)
                                            result = "Error"
                                        Else
                                            FilterRuleGE.Save(True)
                                            result = "Success"

                                            ProductPriceGE.SetValue("FilterRuleId", FilterRuleGE.RecordID)

                                        End If
                                    End If
                                End If

                            Else
                                FilterRuleGE = m_oApp.GetEntityObject("Filter Rules", -1)

                                FilterRuleGE.SetValue("Name", "Price: " & CStr(dr.Item("ProductPriceName")))
                                'FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(Me.GetValue("ProductName")) & ") , Price Name: " & CStr(dr.Item("ProductPriceName")) & ")")
                                FilterRuleGE.SetValue("Description", "Pricing Filter Rule (Product:" & CStr(dr.Item("ProductPriceName")) & ")")
                                FilterRuleGE.SetValue("FilterRuleTypeID", 1)
                                FilterRuleGE.SetValue("LogicString", 1)
                                FilterRuleGE.SetValue("LogicStringGenerated", 0)

                                getFilterRuleItemSequenceSql = "select case when (select max(fri.Sequence) from FilterRuleItem fri where FilterRuleId = (select max(filterruleid) from Productprice where productid = " & ProductGE.RecordID & ") ) is null then 1 else  (select max(fri.Sequence) from FilterRuleItem fri where FilterRuleid = (select max(filterruleid) from Productprice where productid = " & ProductGE.RecordID & ")) + 1 end"
                                lFilterRuleItemSequenceId = CLng(da.ExecuteScalar(getFilterRuleItemSequenceSql))

                                With FilterRuleGE.SubTypes("FilterRuleItems").Add()
                                    .SetValue("Sequence", lFilterRuleItemSequenceId)
                                    .SetValue("ItemNumber", lFilterRuleItemSequenceId)
                                    .SetValue("EntityInstance", "Bill To Person")
                                    .SetValue("Field", "ACSMemberClassID")
                                    .SetValue("Value", dr.Item("ProductFilterRuleValue"))
                                    .SetValue("OperatorID", 1)
                                End With
                                If FilterRuleGE.IsDirty Then 'if the ge has changed then save
                                    If Not FilterRuleGE.Save(False) Then
                                        Throw New Exception("Problem Saving Filter Rule Record:" & FilterRuleGE.RecordID)
                                        result = "Error"
                                    Else
                                        FilterRuleGE.Save(True)
                                        result = "Success"

                                        ProductPriceGE.SetValue("FilterRuleId", FilterRuleGE.RecordID)

                                    End If
                                End If
                            End If


                        End If
                        'If CLng(FilterRuleGE.RecordID) > 0 Then
                        '    ProductPriceGE.SetValue("FilterRuleId", FilterRuleGE.RecordID)
                        'End If

                        If ProductPriceGE.IsDirty Then 'if the ge has changed then save
                            If Not ProductPriceGE.Save(False) Then
                                Throw New Exception("Problem Saving Product Price Record:" & ProductPriceGE.RecordID)
                                result = "Error"
                            Else
                                ProductPriceGE.Save(True)
                                result = "Success"

                            End If
                        End If

                    Next

                End If

            End If


        Catch ex As Exception

        End Try

    End Function

    Public Function CreateCourseGL() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim ID As Long = Me.RecordID
        Dim productGLDetailsSQL As String
        CourseCreationProductId = CLng(Me.GetValue("ProductId"))
        Dim GLCodePrefix As String = "42025.0101."
        Dim CourseCostCenter As Integer = CInt(Me.GetValue("CostCenter"))
        Dim lProductGLAccountDetail As Long
        Dim lSalesGLAccountId As Long

        productGLDetailsSQL = "select case when (select max(ACSNavProduct) from vwGLAccounts) > 80999 then 8" & CourseCreationProductId & " else 80" & CourseCreationProductId & " end"
        lProductGLAccountDetail = CLng(da.ExecuteScalar(productGLDetailsSQL))

        newGl = GLCodePrefix & CourseCostCenter & ".00000." & lProductGLAccountDetail
        Dim GLName As String = "LMS Course Revenue - " & CStr(Me.GetValue("ProductName"))
        GLAccountGE = m_oApp.GetEntityObject("GL Accounts", -1)
        If CourseCreationProductGL <= 0 Then
            GLAccountGE.SetValue("AccountNumber", CStr(newGl))
            GLAccountGE.SetValue("Name", "LMS Course Revenue - " & CStr(Me.GetValue("ProductName")))
            GLAccountGE.SetValue("Type", "Credit")
            GLAccountGE.SetValue("OrganizationID", 1)
            GLAccountGE.SetValue("CurrencyTypeID", 2)
            GLAccountGE.SetValue("IsActive", 1)
            GLAccountGE.SetValue("DefaultARAccount", 0)
            If GLAccountGE.IsDirty Then 'if the ge has changed then save
                If Not GLAccountGE.Save(False) Then
                    Throw New Exception("Problem Saving GL Record:" & GLAccountGE.RecordID)
                    result = "Error"
                Else
                    GLAccountGE.Save(True)

                    result = "Success"

                End If
            End If


            ProductSalesGLID = "select id from ProductGLAccount where productid =" & CourseCreationProductId & " and Type = 'Sales'"
            lSalesGLAccountId = CLng(da.ExecuteScalar(ProductSalesGLID))

            ProductGLAccountsGE = m_oApp.GetEntityObject("ProductGLAccounts", lSalesGLAccountId)

            If lSalesGLAccountId > 0 Then
                ProductGLAccountsGE.SetValue("GLAccountNumber", CStr(newGl))

                If ProductGLAccountsGE.IsDirty Then 'if the ge has changed then save
                    If Not ProductGLAccountsGE.Save(False) Then
                        Throw New Exception("Problem Saving GL Record:" & ProductGLAccountsGE.RecordID)
                        result = "Error"
                    Else
                        ProductGLAccountsGE.Save(True)

                        result = "Success"
                    End If
                End If

            End If

        End If

    End Function

    Public Function CompleteCourseSetup() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim ID As Long = Me.RecordID
        CourseCreationCourseId = CLng(Me.GetValue("CourseId"))
        CourseCreationProductId = CLng(Me.GetValue("ProductId"))
        InstructorId = CLng(Me.GetValue("CourseInstructorId"))
        SchoolId = CLng(Me.GetValue("CourseSchoolId"))
        CourseCreationEventId = CLng(Me.GetValue("EventId"))

        Try
            ClassGE = m_oApp.GetEntityObject("Classes", -1)



            If CourseCreationCourseId > 0 Then
                ClassGE.SetValue("ClassTitle", Me.GetValue("CourseName"))
                ClassGE.SetValue("CourseID", CourseCreationCourseId)
                ClassGE.SetValue("ProductID", CourseCreationProductId)
                ClassGE.SetValue("Status", "Approved")
                ClassGE.SetValue("Type", "Classroom")
                ClassGE.SetValue("SchoolID", SchoolId)
                ClassGE.SetValue("InstructorId", InstructorId)

                If ClassGE.IsDirty Then 'if the ge has changed then save
                    If Not ClassGE.Save(False) Then
                        Throw New Exception("Problem Saving Class Record:" & ClassGE.RecordID)
                        result = "Error"
                    Else
                        ClassGE.Save(True)

                        result = "Success"

                    End If
                End If

                CourseGE = m_oApp.GetEntityObject("Courses", CourseCreationCourseId)
                CourseGE.SetValue("acsIsLms", 1)
                CourseGE.SetValue("ProductID", CourseCreationProductId)
                CourseGE.SetValue("acsCMEEventId", CourseCreationEventId)
                If CourseGE.IsDirty Then 'if the ge has changed then save
                    If Not CourseGE.Save(False) Then
                        Throw New Exception("Problem Saving Course Record:" & CourseGE.RecordID)
                        result = "Error"
                    Else
                        CourseGE.Save(True)

                        result = "Success"

                    End If
                End If
            End If


        Catch ex As Exception

        End Try

    End Function

    Public Sub SendEmail()
        Dim da As New DataAction
        procFlowSql = ("SELECT ID FROM aptify..vwProcessFlows WHERE Name='LMSCourseCreationMessage'")
        lProcessFlowID = CLng(da.ExecuteScalar(procFlowSql))
        If lProcessFlowID > 0 Then
            Dim context As New AptifyContext
            context.Properties.AddProperty("email", Email)
            context.Properties.AddProperty("ccEmail", ccEmail)
            context.Properties.AddProperty("MessageTemplateID", lMessageTemplateID)
            context.Properties.AddProperty("SubjectText", SubjectText)
            context.Properties.AddProperty("HTMLText", HTMLText)
            presult = ProcessFlowEngine.ExecuteProcessFlow(m_oApp, lProcessFlowID, context)  'execute the process flow 
            processFlowResult = "SUCCESS"
        Else
            processFlowResult = "NORECORD"
        End If

    End Sub

End Class
