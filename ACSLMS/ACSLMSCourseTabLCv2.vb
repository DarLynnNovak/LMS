
'Option Explicit On
'Option Strict On
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls

Public Class ACSLMSCourseTabLCv2
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Dim CourseCreatorAppGE As AptifyGenericEntityBase
    Private WithEvents CourseCreateButton As AptifyActiveButton
    Private WithEvents CourseUpdateButton As AptifyActiveButton
    Private WithEvents CourseIdLB As AptifyLinkBox
    Private WithEvents CourseNameTB As AptifyTextBox
    Private WithEvents CourseDescTB As AptifyTextBox
    Private WithEvents CourseCategoryLB As AptifyLinkBox
    Private WithEvents CourseInstructorLB As AptifyLinkBox
    Private WithEvents CourseSchoolLB As AptifyLinkBox
    Private WithEvents isClonedCourse As AptifyCheckBox
    Private WithEvents EthosNodeId As AptifyTextBox
    Private WithEvents CourseStartDate As AptifyTextBox
    Private WithEvents CourseEndDate As AptifyTextBox
    Private WithEvents CourseStartTime As AptifyTimeControl
    Private WithEvents CourseEndTime As AptifyTimeControl
    Private WithEvents isBundledProduct As AptifyCheckBox

    Dim CourseGE As AptifyGenericEntityBase
    Dim ID As Long
    Dim CourseCreationCourseId As Long
    Dim AcsExtId As Long

    Dim result As String = "Failed"



    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try
            'If m_oDA.UserCredentials.Server.ToLower = "aptify" Then
            '    InstructorId = 3285702
            '    SchoolId = 17464
            'End If
            'If m_oDA.UserCredentials.Server.ToLower = "stagingaptify2" Then

            'End If

            ''If m_oDA.UserCredentials.Server.ToLower = "testaptifydb" Then
            ''    InstructorId = 3285702
            ''    SchoolId = 17305
            ''End If
            ''UpdateCourseCreator()
            'If m_oDA.UserCredentials.Server.ToLower = "testaptify610" Then
            '    InstructorId = 3285702
            '    SchoolId = 17464
            'End If
            'InstructorId = 3285702
            'SchoolId = 17464
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
            If CourseCreateButton Is Nothing OrElse CourseCreateButton.IsDisposed = True Then
                CourseCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.Active Button.1"), AptifyActiveButton)
            End If
            If CourseUpdateButton Is Nothing OrElse CourseUpdateButton.IsDisposed = True Then
                CourseUpdateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.Active Button.2"), AptifyActiveButton)
            End If
            If CourseNameTB Is Nothing OrElse CourseNameTB.IsDisposed = True Then
                CourseNameTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseName"), AptifyTextBox)
            End If
            If CourseDescTB Is Nothing OrElse CourseDescTB.IsDisposed = True Then
                CourseDescTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseDescription"), AptifyTextBox)
            End If

            If CourseIdLB Is Nothing OrElse CourseIdLB.IsDisposed = True Then
                CourseIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseId"), AptifyLinkBox)
            End If

            If CourseCategoryLB Is Nothing OrElse CourseCategoryLB.IsDisposed = True Then
                CourseCategoryLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseCategoryId"), AptifyLinkBox)
            End If
            If CourseInstructorLB Is Nothing OrElse CourseInstructorLB.IsDisposed = True Then
                CourseInstructorLB = TryCast(GetFormComponent(Me, "ACS.ACSLMSCourseCreatorApp Course.CourseInstructorId"), AptifyLinkBox)
            End If
            If CourseSchoolLB Is Nothing OrElse CourseSchoolLB.IsDisposed = True Then
                CourseSchoolLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseSchoolId"), AptifyLinkBox)
            End If

            If EthosNodeId Is Nothing OrElse EthosNodeId.IsDisposed = True Then
                EthosNodeId = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course: Step 1.EthosNodeId"), AptifyTextBox)
            End If

            If isClonedCourse Is Nothing OrElse isClonedCourse.IsDisposed = True Then
                isClonedCourse = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course: Step 1.isclonedcourse"), AptifyCheckBox)
            End If

            If CourseStartDate Is Nothing OrElse CourseStartDate.IsDisposed = True Then
                CourseStartDate = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseStartDate"), AptifyTextBox)
            End If
            If CourseEndDate Is Nothing OrElse CourseEndDate.IsDisposed = True Then
                CourseEndDate = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.CourseEndDate"), AptifyTextBox)
            End If
            If CourseStartTime Is Nothing OrElse CourseStartTime.IsDisposed = True Then
                CourseStartTime = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.Time Control.1"), AptifyTimeControl)
            End If
            If CourseEndTime Is Nothing OrElse CourseEndTime.IsDisposed = True Then
                CourseEndTime = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Course.Time Control.2"), AptifyTimeControl)
            End If
            If isBundledProduct Is Nothing OrElse isBundledProduct.IsDisposed = True Then
                isBundledProduct = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.IsBundledProduct"), AptifyCheckBox)
            End If

            If FormTemplateContext.GE.RecordID > 0 Then
                FormTemplateContext.GE.Save()
                If CInt(CourseIdLB.Value) > 0 Then

                    CourseCreateButton.Visible = False
                    CourseUpdateButton.Visible = True
                Else

                    CourseCreateButton.Visible = True
                    CourseUpdateButton.Visible = False

                End If
                If isClonedCourse.Value = 0 Then
                    EthosNodeId.Hide()
                Else
                    EthosNodeId.Show()
                    AcsExtId = EthosNodeId.Value
                End If
            Else
                CourseCreateButton.Visible = False
                CourseUpdateButton.Visible = False
            End If
            InstructorId = 3285702
            SchoolId = 17464
            If CourseInstructorLB.Value Is "" Or IsDBNull(CourseInstructorLB.Value) Then
                CourseInstructorLB.Value = InstructorId

            End If
            If CourseSchoolLB.Value Is "" Or IsDBNull(CourseSchoolLB.Value) Then
                CourseSchoolLB.Value = SchoolId

            End If

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

    Private Sub CourseCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CourseCreateButton.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then

                CreateCourseId()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If



        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub CourseUpdateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CourseUpdateButton.Click

        'Me.CourseIdLB.ClearData()
        ID = Me.FormTemplateContext.GE.RecordID
            Try

            'ParentForm.Close()
            UpdateCourseId()


        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Public Function CreateCourseId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        'ID = Me.FormTemplateContext.GE.RecordID
        CourseCreationCourseId = CourseIdLB.Value
        Dim getInstructorSequenceSql As String
        Dim lInstructorSequenceId As Long
        Dim getSchoolSequenceSql As String
        Dim lSchoolSequenceId As Long

        Try
            UpdateCourseCreator()
            CourseGE = m_oAppObj.GetEntityObject("Courses", -1)
            If CourseCreationCourseId <= 0 Then
                'CourseGE.SetValue("Name", CourseNameTB.Value)
                If Len(CourseNameTB.Value) > 100 Then
                    result = "Failed"
                    MsgBox("The course name must be 100 characters or less.  The course name is: " & Len(CourseNameTB.Value) & " characters")
                Else
                    CourseGE.SetValue("Name", CourseNameTB.Value)
                End If

                If Len(CourseDescTB.Value) > 255 Then
                    result = "Failed"
                    CourseGE.SetValue("Description", CourseDescTB.Value)
                    MsgBox("The Course Description must be 255 characters or less.  The Course Description is: " & Len(CourseDescTB.Value) & " characters")
                Else

                    CourseGE.SetValue("Description", CourseDescTB.Value)
                End If
                CourseGE.SetValue("CategoryID", CourseCategoryLB.Value)
                If isBundledProduct.Value = True Then
                    CourseGE.SetValue("ProductTypeID", 5)
                Else
                    CourseGE.SetValue("ProductTypeID", 7)
                End If

                CourseGE.SetValue("Status", "Available")
                CourseGE.SetValue("Units", 0)

                If IsDBNull(AcsExtId) Then
                Else
                    CourseGE.SetValue("acsExternalId", AcsExtId)
                End If
                CourseGE.SetValue("acsIsLms", 0)
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

                Dim courseInstructorGE As AptifyGenericEntityBase = m_oAppObj.GetEntityObject("courseinstructors", -1)

                courseInstructorGE.SetValue("CourseID", CourseGE.RecordID)
                courseInstructorGE.SetValue("Sequence", lInstructorSequenceId)
                courseInstructorGE.SetValue("InstructorID", CourseInstructorLB.Value)
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

                Dim courseSchoolGE As AptifyGenericEntityBase = m_oAppObj.GetEntityObject("courseschools", -1)

                courseSchoolGE.SetValue("CourseID", CourseGE.RecordID)
                courseSchoolGE.SetValue("Sequence", lSchoolSequenceId)
                courseSchoolGE.SetValue("SchoolID", CourseSchoolLB.Value)
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
                If result = "Success" Then
                    CourseIdLB.Value = CourseGE.RecordID
                    MsgBox("Success.  Please be sure to save your changes when closing this form.")
                    SetCourseCreatorDate()
                    DisplayEntity()
                End If

            End If



        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Function


    Public Function UpdateCourseId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        'ID = Me.FormTemplateContext.GE.RecordID
        CourseCreationCourseId = Me.FormTemplateContext.GE.GetValue("CourseId")
        Dim getInstructorRecordSql As String
        Dim lInstructorRecordId As Long
        Dim getSchoolRecordSql As String
        Dim lSchoolRecordId As Long

        CourseGE = m_oAppObj.GetEntityObject("Courses", CourseCreationCourseId)

        Try
            'Me.CourseIdLB.ClearData()
            If CourseCreationCourseId > 0 Then

                If Len(CourseNameTB.Value) > 100 Then
                    result = "Failed"
                    MsgBox("The course name must be 100 characters or less.  The course name is: " & Len(CourseNameTB.Value) & " characters")
                Else
                    CourseGE.SetValue("Name", CourseNameTB.Value)
                End If
                If Len(CourseDescTB.Value) > 255 Then
                    result = "Failed"
                    CourseGE.SetValue("Description", CourseDescTB.Value)
                    MsgBox("The Course Description must be 255 characters or less.  The Course Description is: " & Len(CourseDescTB.Value) & " characters")
                Else

                    CourseGE.SetValue("Description", CourseDescTB.Value)
                End If
                CourseGE.SetValue("CategoryID", CourseCategoryLB.Value)
                If isBundledProduct.Value = True Then
                    CourseGE.SetValue("ProductTypeID", 5)
                Else
                    CourseGE.SetValue("ProductTypeID", 7)
                End If
                CourseGE.SetValue("Status", "Available")
                CourseGE.SetValue("Units", 0)
                If IsDBNull(AcsExtId) Then
                Else
                    CourseGE.SetValue("acsExternalId", AcsExtId)
                End If
                'If CourseGE.IsDirty Then 'if the ge has changed then save
                If Not CourseGE.Save(False) Then
                    Throw New Exception("Problem Saving Course Record:" & CourseGE.RecordID)
                    result = "Error"
                Else
                    CourseGE.Save(True)
                    result = "Success"

                End If
                ' End If
                getInstructorRecordSql = "select id from courseinstructor ci where courseid = " & CourseGE.RecordID
                lInstructorRecordId = CLng(da.ExecuteScalar(getInstructorRecordSql))

                Dim courseInstructorGE As AptifyGenericEntityBase = m_oAppObj.GetEntityObject("courseinstructors", lInstructorRecordId)
                courseInstructorGE.SetValue("InstructorID", CourseInstructorLB.Value)
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

                Dim courseSchoolGE As AptifyGenericEntityBase = m_oAppObj.GetEntityObject("courseschools", lSchoolRecordId)
                courseSchoolGE.SetValue("SchoolID", CourseSchoolLB.Value)
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
                If result = "Success" Then
                    CourseIdLB.Value = CourseGE.RecordID
                    MsgBox("Success.  Please be sure to save your changes when closing this form.")
                    UpdateCourseCreator()
                    DisplayEntity()
                End If


            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Function

    Public Function SetCourseCreatorDate() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.FormTemplateContext.GE.RecordID
        Try
            'With CourseCreatorAppGE
            ' CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
            FormTemplateContext.GE.SetValue("CourseIdCreated", 1)
            FormTemplateContext.GE.SetValue("CourseIdCreatedDate", Now())
            FormTemplateContext.GE.SetValue("CourseName", CourseNameTB.Value)
            FormTemplateContext.GE.SetValue("CourseDescription", CourseDescTB.Value)
            FormTemplateContext.GE.SetValue("CourseCategoryID", CourseCategoryLB.Value)
            FormTemplateContext.GE.SetValue("CourseStartDate", CourseStartDate.Value)
            FormTemplateContext.GE.SetValue("CourseEndDate", CourseEndDate.Value)
            FormTemplateContext.GE.SetValue("CourseStartTime", CourseStartTime.Value)
            FormTemplateContext.GE.SetValue("CourseEndTime", CourseEndTime.Value)
            'FormTemplateContext.GE.SetValue("CourseInstructorId", CourseInstuctorLB.Value)
            'FormTemplateContext.GE.SetValue("CourseSchoolId", CourseSchoolLB.Value)
            If Not FormTemplateContext.GE.Save(False) Then
                Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                result = "Error"
            Else
                result = "Success"
                FormTemplateContext.GE.Save()
                'CourseCreatorAppGE.Save(True)
                'CourseCreatorAppGE.CommitTransaction()
                'UpdateCourseCreator()
                'ParentForm.
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
                FormTemplateContext.GE.SetValue("CourseName", CourseNameTB.Value)
                FormTemplateContext.GE.SetValue("CourseDescription", CourseDescTB.Value)
                FormTemplateContext.GE.SetValue("CourseCategoryID", CourseCategoryLB.Value)
                FormTemplateContext.GE.SetValue("CourseStartDate", CourseStartDate.Value)
                FormTemplateContext.GE.SetValue("CourseEndDate", CourseEndDate.Value)
                FormTemplateContext.GE.SetValue("CourseStartTime", CourseStartTime.Value)
                FormTemplateContext.GE.SetValue("CourseEndTime", CourseEndTime.Value)
                'FormTemplateContext.GE.SetValue("CourseInstructorId", CourseInstuctorLB.Value)
                'FormTemplateContext.GE.SetValue("CourseSchoolId", CourseSchoolLB.Value)


                If Not FormTemplateContext.GE.Save(False) Then
                    Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                    result = "Error"
                Else
                    result = "Success"
                    FormTemplateContext.GE.Save(True)
                    'CourseCreatorAppGE.CommitTransaction()

                End If
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
