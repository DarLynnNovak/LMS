
'Option Explicit On
'Option Strict On

Imports Aptify.Framework.WindowsControls
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.BusinessLogic.GenericEntity

Public Class ACSLMSProductTabLCv2
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Dim CourseCreatorAppGE As AptifyGenericEntityBase
    Dim ProductGE As AptifyGenericEntityBase

    Private WithEvents ProductCreateButton As AptifyActiveButton
    Private WithEvents ProductUpdateButton As AptifyActiveButton
    Private WithEvents GLCreateButton As AptifyActiveButton
    Private WithEvents ProductIdLB As AptifyLinkBox
    Private WithEvents SalesGLTB As AptifyTextBox
    Private WithEvents ProductNameTB As AptifyTextBox
    Private WithEvents ProductTemplateLB As AptifyLinkBox
    Private WithEvents isBundledProduct As AptifyCheckBox
    Private WithEvents ProductIdDCB As AptifyDataComboBox
    Private WithEvents AccreditationProgCategory As AptifyLinkBox
    Private _parentForm As System.Windows.Forms.Form
    Dim da As New DataAction
    Dim courseCreatorGroupSQL As String
    Dim courseOwnerSQL As String
    Dim CourseOwnerId As String
    Dim CourseOwner As Integer
    Dim UserCreatedId As String
    Dim CourseOwnerPersonId As Integer
    Dim CourseCreationProductId As Long
    Dim CourseCreationProductGL As Long
    Dim GLAccountGE As AptifyGenericEntityBase
    Dim ProductGLAccountsGE As AptifyGenericEntityBase
    Dim ProductPriceGE As AptifyGenericEntityBase
    Dim FilterRuleGE As AptifyGenericEntityBase
    Dim FilterRuleItemGE As AptifyGenericEntityBase
    Dim AccountingGLGE As AptifyGenericEntityBase
    Dim SalesGl As String
    Dim ARGL As String
    Dim DefGL As String
    Dim result As String = "Failed"
    Dim ProductSalesGLID As String
    Dim thisCategoryId As Integer
    Dim ID As Long
    Dim ProductAccredSubCatPrefixSql As String
    Dim ProductAccredSubCatPrefix As String

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
            If ProductCreateButton Is Nothing OrElse ProductCreateButton.IsDisposed = True Then
                ProductCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.Active Button.1"), AptifyActiveButton)
            End If
            If ProductUpdateButton Is Nothing OrElse ProductUpdateButton.IsDisposed = True Then
                ProductUpdateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.Active Button.2"), AptifyActiveButton)
            End If

            If GLCreateButton Is Nothing OrElse GLCreateButton.IsDisposed = True Then
                GLCreateButton = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.Active Button.3"), AptifyActiveButton)
            End If
            If ProductIdLB Is Nothing OrElse ProductIdLB.IsDisposed = True Then
                ProductIdLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.ProductId"), AptifyLinkBox)
            End If
            If SalesGLTB Is Nothing OrElse ProductIdLB.IsDisposed = True Then
                SalesGLTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.SalesGL"), AptifyTextBox)
            End If
            If ProductNameTB Is Nothing OrElse ProductNameTB.IsDisposed = True Then
                ProductNameTB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.ProductName"), AptifyTextBox)
            End If
            If ProductTemplateLB Is Nothing OrElse ProductTemplateLB.IsDisposed = True Then
                ProductTemplateLB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.ProductEmailTemplateId"), AptifyLinkBox)
            End If
            If isBundledProduct Is Nothing OrElse isBundledProduct.IsDisposed = True Then
                isBundledProduct = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.IsBundledProduct"), AptifyCheckBox)
            End If
            If ProductIdDCB Is Nothing OrElse ProductIdDCB.IsDisposed = True Then
                ProductIdDCB = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.ProductCategoryID"), AptifyDataComboBox)
            End If

            If AccreditationProgCategory Is Nothing OrElse AccreditationProgCategory.IsDisposed = True Then
                AccreditationProgCategory = TryCast(GetFormComponent(Me, "ACSLMSCourseCreatorApp Product.AccreditationProgCategory"), AptifyLinkBox)
            End If
            _parentForm = ParentForm

            If FormTemplateContext.GE.RecordID > 0 Then
                ID = FormTemplateContext.GE.RecordID
                'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
                ProductNameTB.Value = FormTemplateContext.GE.GetValue("CourseName")


            End If
            CheckCourseOwnerLB()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub ProductCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProductCreateButton.Click
        Try

            If Me.FormTemplateContext.GE.RecordID > 0 Then
                ' ParentForm.Close()
                CreateProductId()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Private Sub ProductUpdateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProductUpdateButton.Click
        Try
            'ParentForm.Close()
            UpdateProductId()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Private Sub GLCreateButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GLCreateButton.Click
        Try
            CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
            Dim CostCenterId As String = CourseCreatorAppGE.GetValue("CostCenter")
            If CostCenterId = "" Then
                MsgBox("Cost Center cannot be blank.")

            Else
                CreateCourseGL()
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Public Function CreateProductId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim getProductPriceSequenceSql As String
        Dim lProductPriceSequenceId As Long
        Dim getFilterRuleItemSequenceSql As String
        Dim lFilterRuleItemSequenceId As Long
        Dim getAccountingGLItemSequenceSql As String
        Dim lAccountingGLItemSequenceSql As Long
        Dim sSQL As String
        Dim dt As DataTable
        CourseCreationProductId = Me.FormTemplateContext.GE.GetValue("ProductId")
        Dim CourseCreationProductName As String = ProductNameTB.Value


        Try

            UpdateCourseCreator()
            ProductGE = m_oAppObj.GetEntityObject("Products", -1)
            If CourseCreationProductId <= 0 Then
                If Len(CourseCreationProductName) > 100 Then
                    result = "Failed"
                    MsgBox("The product name must be 100 characters or less.  The product name is: " & Len(CourseCreationProductName) & " characters")
                Else
                    ProductGE.SetValue("Name", CourseCreationProductName)
                End If

                'ProductGE.SetValue("CategoryID", CInt(thisCategoryId))
                ProductGE.SetValue("CategoryID", ProductIdDCB.Value)
                If FormTemplateContext.GE.GetValue("IsBundledProduct") = True Then
                    ProductGE.SetValue("ProductTypeID", 5)
                Else
                    ProductGE.SetValue("ProductTypeID", 7)
                End If
                'ProductGE.SetValue("ProductCategoryID", ProductIdDCB.Value)
                ProductGE.SetValue("ACSHasCustomMessage", 1)
                ProductGE.SetValue("ACSMessageTemplateID", ProductTemplateLB.Value)
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


                sSQL = "select * from ACSLMSCourseCreatorProductPrices where ACSLMSCourseCreatorAppID = " & Me.FormTemplateContext.GE.RecordID
                dt = m_oDA.GetDataTable(sSQL)
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows

                        getProductPriceSequenceSql = "select case when (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & " ) is null then 1 else  (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & ") + 1 end"
                        lProductPriceSequenceId = CLng(da.ExecuteScalar(getProductPriceSequenceSql))

                        ProductPriceGE = m_oAppObj.GetEntityObject("ProductPrices", -1)
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

                        FilterRuleGE = m_oAppObj.GetEntityObject("Filter Rules", -1)

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


                        ProductPriceGE = m_oAppObj.GetEntityObject("ProductPrices", ProductPriceGE.RecordID)
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
                If result = "Success" Then

                    ProductIdLB.Value = ProductGE.RecordID
                    MsgBox("Success.  Please be sure to save your changes when closing this form.")
                    'Select Case MsgBox("Success, Please save this form.  Save? ", MsgBoxStyle.YesNo, "Course Creator")
                    '    Case MsgBoxResult.Yes
                    '        ParentForm.Close()
                    SetCourseCreatorDate()
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

    Public Function UpdateProductId() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim getProductPriceSequenceSql As String

        Dim getCurrentPriceFilterRecordSql As String
        Dim getFilterRuleItemSequenceSql As String
        Dim lFilterRuleItemSequenceId As Long
        Dim lProductPriceSequence As Long
        Dim lProductPriceSequenceId As Long
        Dim lCurrentPriceFilterRecordId As Long
        Dim getFilterRuleRecordSql As String
        Dim lFilterRuleRecordId As Long
        Dim getFilterRuleItemRecordSql As String
        Dim lFilterRuleItemRecordId As Long

        Dim sSQL As String
            Dim dt As DataTable
            CourseCreationProductId = Me.FormTemplateContext.GE.GetValue("ProductId")
            Dim CourseCreationProductName As String = ProductNameTB.Value


        Try
            Me.ProductIdLB.ClearData()
            'UpdateCourseCreator()
            ProductGE = m_oAppObj.GetEntityObject("Products", CourseCreationProductId)
            If CourseCreationProductId > 0 Then
                'ProductGE.SetValue("Name", CourseCreationProductName)
                If Len(CourseCreationProductName) > 100 Then
                    result = "Failed"
                    MsgBox("The product name must be 100 characters or less.  The product name is: " & Len(CourseCreationProductName) & " characters")
                Else
                    ProductGE.SetValue("Name", CourseCreationProductName)
                End If
                'ProductGE.SetValue("CategoryID", CInt(thisCategoryId))
                ProductGE.SetValue("CategoryID", ProductIdDCB.Value)
                If FormTemplateContext.GE.GetValue("IsBundledProduct") = True Then
                    ProductGE.SetValue("ProductTypeID", 5)
                Else
                    ProductGE.SetValue("ProductTypeID", 7)
                End If
                'ProductGE.SetValue("ProductCategoryID", ProductIdDCB.Value)
                ProductGE.SetValue("ACSHasCustomMessage", 1)
                ProductGE.SetValue("ACSMessageTemplateID", ProductTemplateLB.Value)
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


                sSQL = "select * from ACSLMSCourseCreatorProductPrices where ACSLMSCourseCreatorAppID = " & FormTemplateContext.GE.RecordID

                dt = m_oDA.GetDataTable(sSQL)

                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        lProductPriceSequence = CLng(dr.Item("sequence"))
                        getCurrentPriceFilterRecordSql = "select ID from productprice pp where productid = " & CourseCreationProductId & " and Sequence = " & lProductPriceSequence
                        lCurrentPriceFilterRecordId = CLng(da.ExecuteScalar(getCurrentPriceFilterRecordSql))

                        If lCurrentPriceFilterRecordId > 0 Then 'Updating existing prices and prices filters

                            ProductPriceGE = m_oAppObj.GetEntityObject("ProductPrices", lCurrentPriceFilterRecordId)
                            ProductPriceGE.SetValue("ProductId", ProductGE.RecordID)
                            ProductPriceGE.SetValue("Name", dr.Item("ProductPriceName"))
                            ProductPriceGE.SetValue("MemberTypeID", dr.Item("ProductFilterMemberType"))
                            ProductPriceGE.SetValue("Price", dr.Item("ProductFilterRulePrice"))
                            ProductPriceGE.SetValue("CurrencyTypeID", 2)

                        Else
                            getProductPriceSequenceSql = "select case when (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & " ) is null then 1 else  (select max(pp.Sequence) from productprice pp where productid = " & CourseCreationProductId & ") + 1 end"
                            lProductPriceSequenceId = CLng(da.ExecuteScalar(getProductPriceSequenceSql))

                            ProductPriceGE = m_oAppObj.GetEntityObject("ProductPrices", -1) 'Creating new prices and price filters.
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
                                FilterRuleGE = m_oAppObj.GetEntityObject("Filter Rules", lFilterRuleRecordId)

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
                                    FilterRuleItemGE = m_oAppObj.GetEntityObject("FilterRuleItems", lFilterRuleItemRecordId)
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
                                    FilterRuleGE = m_oAppObj.GetEntityObject("Filter Rules", -1)

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
                                FilterRuleGE = m_oAppObj.GetEntityObject("Filter Rules", -1)

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

                If result = "Success" Then

                    ProductIdLB.Value = ProductGE.RecordID
                    MsgBox("Success.  Please be sure to save your changes when closing this form.")
                    SetCourseCreatorDate()
                    UpdateCourseCreator()
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


    Public Function CreateCourseGL() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        Dim productGLDetailsSQL As String
        CourseCreationProductId = Me.FormTemplateContext.GE.GetValue("ProductId")
        Dim GLSalesCodePrefix As String = "42110.0101"
        Dim CourseCostCenter As String
        Dim CourseCostCenterSQL As String
        Dim lProductGLAccountDetail As Long
        Dim lSalesGLAccountId As Long
        Dim ProductCodeFirstNumb As Long
        Dim ProductCodeSecNumb As Long
        Dim ProductCodeThirdNumb As Long




        Try
            'ProductCodeFirstNumb = "SELECT substring(max(acsnavproduct), 1, 1) from vwglaccounts"
            'ProductCodeSecNumb = "SELECT substring(max(acsnavproduct), 2, 1) from vwglaccounts"
            'ProductCodeThirdNumb = CourseCreationProductId

            'ProductAccredSubCatPrefixSql = "select GLAccountReceivable from acsaccreditationprogramcategorymap where id = " & CInt(AccreditationProgCategory.Value)
            'ProductAccredSubCatPrefix = m_oDA.ExecuteScalar(ProductAccredSubCatPrefixSql)

            productGLDetailsSQL = "8" & CourseCreationProductId
            'lProductGLAccountDetail = CLng(da.ExecuteScalar(productGLDetailsSQL))
            lProductGLAccountDetail = productGLDetailsSQL

            CourseCostCenterSQL = "select costcenter from vwacslmscoursecreatorapp where id  = " & ID
            CourseCostCenter = da.ExecuteScalar(CourseCostCenterSQL)

            SalesGl = GLSalesCodePrefix & "." & CourseCostCenter & ".00000." & lProductGLAccountDetail
            'ARGL = ProductAccredSubCatPrefix & "." & CourseCostCenter & ".00000." & lProductGLAccountDetail

            Dim GLName As String = "LMS Course Revenue - " & ProductNameTB.Value
            GLAccountGE = m_oAppObj.GetEntityObject("GL Accounts", -1)

            If CourseCreationProductGL <= 0 Then
                GLAccountGE.SetValue("AccountNumber", CStr(SalesGl))
                GLAccountGE.SetValue("Name", "LMS Course Revenue - " & ProductNameTB.Value)
                GLAccountGE.SetValue("Type", "Credit")
                GLAccountGE.SetValue("OrganizationID", 1)
                GLAccountGE.SetValue("CurrencyTypeID", 2)
                GLAccountGE.SetValue("IsActive", 1)
                'GLAccountGE.SetValue("DefaultARAccount", ARGL)
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

                ProductGLAccountsGE = m_oAppObj.GetEntityObject("ProductGLAccounts", lSalesGLAccountId)

                If lSalesGLAccountId > 0 Then
                    ProductGLAccountsGE.SetValue("GLAccountNumber", CStr(SalesGl))

                    If ProductGLAccountsGE.IsDirty Then 'if the ge has changed then save
                        If Not ProductGLAccountsGE.Save(False) Then
                            Throw New Exception("Problem Saving GL Record:" & ProductGLAccountsGE.RecordID)
                            result = "Error"
                        Else
                            ProductGLAccountsGE.Save(True)

                            result = "Success"
                        End If
                    End If
                Else

                End If
                If result = "Success" Then

                    FormTemplateContext.GE.SetValue("SalesGL", SalesGl)
                    SalesGLTB.Value = SalesGl
                    'With CourseCreatorAppGE


                    If Not FormTemplateContext.GE.Save(False) Then
                        Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                        result = "Error"
                    Else

                        CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
                        CourseCreatorAppGE.SetValue("SalesGL", SalesGl)
                        CourseCreatorAppGE.SetValue("GLCreated", 1)
                        CourseCreatorAppGE.SetValue("GLCreatedDate", Now())
                        CourseCreatorAppGE.Save(True)
                        _parentForm.Refresh()
                        FormTemplateContext.GE.Save(True)

                        UpdateCourseCreator()
                        MsgBox("Success.  Please be sure to save your changes when closing this form.")
                        'CourseCreatorAppGE.CommitTransaction()

                    End If

                End If

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
            thisCategoryId = 48
            thisEntityID = 2788
        End If
        If Me.DataAction.UserCredentials.Server.ToLower = "stagingaptify2" Then
            'staging
            thisCategoryId = 47
            thisEntityID = 2776
        End If

        If Me.DataAction.UserCredentials.Server.ToLower = "testaptifydb" Then
            'staging 
            thisCategoryId = 47
            thisEntityID = 2863
        End If

        If m_oDA.UserCredentials.Server.ToLower = "testaptify610" Then
            thisCategoryId = 48
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
            FormTemplateContext.GE.Save(True)
            If CInt(ProductIdLB.Value) > 0 Then
                ProductCreateButton.Visible = False
                ProductUpdateButton.Visible = True
                GLCreateButton.Visible = False
            Else
                ProductCreateButton.Visible = True
                ProductUpdateButton.Visible = False
                GLCreateButton.Visible = False
            End If

        Else
            'MsgBox("This record has not been created yet.  Please save the form to create the record")
            ProductCreateButton.Visible = False
            ProductUpdateButton.Visible = False
            GLCreateButton.Visible = False
        End If
        If dt.Rows.Count() > 0 AndAlso CInt(ProductIdLB.Value) > 0 AndAlso CStr(SalesGLTB.Value) = "" Then
            GLCreateButton.Visible = True
        Else
            GLCreateButton.Visible = False
        End If
    End Sub

    Public Function SetCourseCreatorDate() As String
        Dim bResult As Boolean = False
        Dim da As New DataAction
        ID = Me.FormTemplateContext.GE.RecordID
        Try
            'With CourseCreatorAppGE
            'CourseCreatorAppGE = m_oAppObj.GetEntityObject("ACSLMSCourseCreatorApp", ID)
            FormTemplateContext.GE.SetValue("ProductIdCreated", 1)
            FormTemplateContext.GE.SetValue("ProductIdCreatedDate", Now())
            FormTemplateContext.GE.SetValue("ProductName", ProductNameTB.Value)
            FormTemplateContext.GE.SetValue("ProductEmailTemplateId", ProductTemplateLB.Value)
            'FormTemplateContext.GE.SetValue("IsBundledProduct", isBundledProduct.Value)
            If Not FormTemplateContext.GE.Save(False) Then
                Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                result = "Error"
            Else
                result = "Success"
                FormTemplateContext.GE.Save(True)

                'CourseCreatorAppGE.CommitTransaction()
                ' UpdateCourseCreator()
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
                FormTemplateContext.GE.SetValue("ProductName", ProductNameTB.Value)
                FormTemplateContext.GE.SetValue("ProductEmailTemplateId", ProductTemplateLB.Value)
                'FormTemplateContext.GE.SetValue("IsBundledProduct", isBundledProduct.Value)
                If Not FormTemplateContext.GE.Save(False) Then
                    Throw New Exception("Problem Saving Product Record:" & FormTemplateContext.GE.RecordID)
                    result = "Error"
                Else
                    result = "Success"
                    FormTemplateContext.GE.Save(True)
                    'CourseCreatorAppGE.CommitTransaction()

                End If
            End If
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

