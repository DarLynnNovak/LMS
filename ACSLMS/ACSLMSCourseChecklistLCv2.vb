'Option Explicit On
'Option Strict On

Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.BusinessLogic.GenericEntity

Public Class ACSLMSCourseChecklistLCv2
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Dim CourseCreatorAppGE As AptifyGenericEntityBase
    Private WithEvents CompleteSetupButton As AptifyActiveButton
    Private WithEvents CourseIdCB As AptifyCheckBox
    Private WithEvents ProductIdCB As AptifyCheckBox
    Private WithEvents CourseCompleteCB As AptifyCheckBox
    Private WithEvents CourseCompleteDate As AptifyTextBox
    Private WithEvents CourseOwnerIdLB As AptifyLinkBox
    Dim CourseGE As AptifyGenericEntityBase
    Dim ClassGE As AptifyGenericEntityBase
    Dim CourseIdCreated As Integer
    Dim EventIdCreated As Integer
    Dim ProductIdCreated As Integer
    Dim ProductId As Integer
    Dim GLCreated As Integer
    Dim CourseSetupComplete As Integer
    Dim userid As Long
    Dim courseCreatorGroupSQL As String
    Dim courseOwnerSQL As String
    Dim CourseOwnerId As String
    Dim CourseOwner As Integer
    Dim CourseId As Integer
    Dim UserCreatedId As String
    Dim CourseOwnerPersonId As Integer
    Dim ID As Long
    Dim CourseCreationCourseId As Long
    Dim CourseCreationEventId As Long
    Dim CourseCreationProductId As Long
    Dim CourseCreationProductGL As Long
    Dim result As String = "Failed"

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
            If CompleteSetupButton Is Nothing OrElse CompleteSetupButton.IsDisposed = True Then
                CompleteSetupButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp CheckList: Step 4.Active Button.1"), AptifyActiveButton)
            End If
            If CourseCompleteCB Is Nothing OrElse CourseCompleteCB.IsDisposed = True Then
                CourseCompleteCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp CheckList: Step 4.CourseSetupComplete"), AptifyCheckBox)
            End If
            If CourseCompleteDate Is Nothing OrElse CourseCompleteDate.IsDisposed = True Then
                CourseCompleteDate = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp CheckList: Step 4.CourseSetupCompleteDate"), AptifyTextBox)
            End If

            'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)

            CourseId = FormTemplateContext.GE.GetValue("CourseId")
            EventIdCreated = FormTemplateContext.GE.GetValue("EventIdCreated")
            ProductIdCreated = FormTemplateContext.GE.GetValue("ProductIdCreated")
            ProductId = FormTemplateContext.GE.GetValue("ProductId")
            GLCreated = FormTemplateContext.GE.GetValue("GLCreated")
            CourseSetupComplete = FormTemplateContext.GE.GetValue("CourseSetupComplete")
            CheckCourseOwnerLB()


        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub


    Private Sub CompleteSetupButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CompleteSetupButton.Click
        Try

            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CompleteCourseSetup()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

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
        'If Me.FormTemplateContext.GE.RecordID > 0 AndAlso (dt1.Rows.Count > 0 AndAlso dt.Rows.Count = 0) Then
        If Me.FormTemplateContext.GE.RecordID > 0 Then
            FormTemplateContext.GE.Save(True)
            If CourseId > 0 AndAlso ProductId > 0 Then
                CompleteSetupButton.Visible = True
            Else
                CompleteSetupButton.Visible = False
            End If
        Else
            'MsgBox("This record has not been created yet.  Please save the form to create the record")
            CompleteSetupButton.Visible = False
        End If

    End Sub

    Public Function CompleteCourseSetup() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.FormTemplateContext.GE.RecordID
        Dim sql As String
        Dim dt As DataTable

        Try




            If ID < 0 Then
                MsgBox("Please save this record before proceeding")
            Else
                sql = "select * from acslmscoursecreatorapp where id = " & ID
                dt = m_oDA.GetDataTable(sql)
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows

                        If IsDBNull(dr.Item("CourseId")) Then
                        Else
                            CourseCreationCourseId = dr.Item("CourseId")
                        End If
                        If IsDBNull(dr.Item("ProductId")) Then
                        Else

                            CourseCreationProductId = dr.Item("ProductId")
                        End If
                        If IsDBNull(dr.Item("EventId")) Then
                        Else

                            CourseCreationEventId = dr.Item("EventId")
                        End If

                        Dim InstructorId As Long = dr.Item("CourseInstructorId")
                        Dim SchoolId As Long = dr.Item("CourseSchoolId")


                        ClassGE = m_oAppObj.GetEntityObject("Classes", -1)



                        If ID > 0 Then
                            ClassGE.SetValue("ClassTitle", dr.Item("CourseName"))
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

                            CourseGE = m_oAppObj.GetEntityObject("Courses", CourseCreationCourseId)
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



                    Next
                End If
                If result = "Success" Then
                    MsgBox("Success.  Please be sure to save your changes when closing this form.")
                    SetCourseCreatorDate()
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
            'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
            FormTemplateContext.GE.SetValue("CourseSetupComplete", 1)
            FormTemplateContext.GE.SetValue("CourseSetupCompleteDate", Now())
            CourseCompleteCB.Value = True
            CourseCompleteDate.Value = Now()

            If Not FormTemplateContext.GE.Save(False) Then
                Throw New Exception("Problem Saving Product Record:" & CourseCreatorAppGE.RecordID)
                result = "Error"
            Else
                result = "Success"
                FormTemplateContext.GE.Save(True)
                'CourseCreatorAppGE.CommitTransaction()
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Function


End Class
