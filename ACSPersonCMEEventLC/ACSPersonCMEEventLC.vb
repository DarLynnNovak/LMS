Option Explicit On
'Option Strict On

Imports Aptify.Framework.WindowsControls


Public Class ACSPersonCMEEventLC
    Inherits FormTemplateLayout

    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New Aptify.Framework.DataServices.DataAction
    Private bAdded As Boolean = False
    Protected lGridID As Long = -1
    Public WithEvents btnCMEPrint As AptifyActiveButton
    Protected WithEvents lPersonLinkbox As AptifyLinkBox
    Protected WithEvents lEligibilityLinkbox As AptifyLinkBox
    Protected WithEvents lEventAssocLinkbox As AptifyDataComboBox
    Protected WithEvents lPersonEventLinkbox As AptifyLinkBox
    Dim personEligibility As Integer

    'Protected WithEvents grdVRCCDSearch As DataGridView

    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As Aptify.Framework.WindowsControls.FormTemplateLayout.FormTemplateLoadedEventArgs)
        Try
            'grdVRCCDSearch = CreateGrid()'
            MyBase.OnFormTemplateLoaded(e)
            FindControls()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e)
    End Sub
    Protected Overridable Sub FindControls()
        Try
            If lPersonLinkbox Is Nothing OrElse lPersonLinkbox.IsDisposed = True Then
                lPersonLinkbox = TryCast(Me.GetFormComponent(Me, "ACS.ACSPersonCME.PersonID"), AptifyLinkBox)
            End If
            If lPersonEventLinkbox Is Nothing OrElse lPersonEventLinkbox.IsDisposed = True Then
                lPersonEventLinkbox = TryCast(Me.GetFormComponent(Me, "ACS.ACSPersonCME.ACSCMEEventID"), AptifyLinkBox)
            End If
            If lEligibilityLinkbox Is Nothing OrElse lEligibilityLinkbox.IsDisposed = True Then
                lEligibilityLinkbox = TryCast(Me.GetFormComponent(Me, "ACS.ACSPersonCME.Tabs.General.ACSCmeEligibilityId"), AptifyLinkBox)
            End If
            If lEventAssocLinkbox Is Nothing OrElse lEventAssocLinkbox.IsDisposed = True Then
                lEventAssocLinkbox = TryCast(Me.GetFormComponent(Me, "ACS.ACSPersonCME.Tabs.General.CmeEventAssociationId"), AptifyDataComboBox)
            End If
            If Me.btnCMEPrint Is Nothing OrElse Me.btnCMEPrint.IsDisposed Then
                Me.btnCMEPrint = TryCast(Me.GetFormComponent(Me, "ACS.ACSPersonCME.Tabs.General.Print Button"), AptifyActiveButton)
            End If

            personEligibility = System.Convert.ToString(Me.DataAction.ExecuteScalar("select acscmeeligibilityid from aptify..vwPersons (nolock) where personid=" & CStr(lPersonLinkbox.Value.ToString)))
            If Me.lEligibilityLinkbox.Value < 0 Then
                Me.lEligibilityLinkbox.Value = personEligibility
            End If

            If lEligibilityLinkbox.Value = 1 Then
                lEventAssocLinkbox.DisplaySQL = CStr("select ID,CMEAssociationId_Name + ' | ' + ApprovalCategory from aptify..vwAcsCmeEventAssociation where cmeeventid=" & lPersonEventLinkbox.Value & " and ApprovalCategory = 'CME'")
            Else
                lEventAssocLinkbox.DisplaySQL = CStr("select ID,CMEAssociationId_Name + ' | ' + ApprovalCategory from aptify..vwAcsCmeEventAssociation where cmeeventid=" & lPersonEventLinkbox.Value & " and ApprovalCategory <> 'CME'")
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub lPersonEventLinkbox_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles lPersonEventLinkbox.ValueChanged
        If Not NewValue Is Nothing Then
            If IsNumeric(NewValue) AndAlso CLng(NewValue) > 0 Then
                Dim dt As DataTable

                If lEligibilityLinkbox.Value = 1 Then
                    lEventAssocLinkbox.DisplaySQL = CStr("select ID,CMEAssociationId_Name + ' | ' + ApprovalCategory from aptify..vwAcsCmeEventAssociation where cmeeventid=" & lPersonEventLinkbox.Value & " and ApprovalCategory = 'CME'")
                Else
                    lEventAssocLinkbox.DisplaySQL = CStr("select ID,CMEAssociationId_Name + ' | ' + ApprovalCategory from aptify..vwAcsCmeEventAssociation where cmeeventid=" & lPersonEventLinkbox.Value & " and ApprovalCategory <> 'CME'")
                End If

                'End If
                'If IsNumeric(NewValue) AndAlso CLng(NewValue) = 6 Then

            End If
        End If

    End Sub
    Private Sub lPersonLinkbox_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles lPersonLinkbox.ValueChanged
        If Not NewValue Is Nothing Then
            If IsNumeric(NewValue) AndAlso CLng(NewValue) > 0 Then
                Dim dt As DataTable

                personEligibility = System.Convert.ToString(Me.DataAction.ExecuteScalar("select acscmeeligibilityid from aptify..vwPersons (nolock) where personid=" & CStr(lPersonLinkbox.Value.ToString)))
                If Me.lEligibilityLinkbox.Value < 0 Then
                    Me.lEligibilityLinkbox.Value = personEligibility
                End If

                'End If
                'If IsNumeric(NewValue) AndAlso CLng(NewValue) = 6 Then

            End If
        End If

    End Sub

    Private Sub btnCMEPrint_Click(sender As Object, e As EventArgs) Handles btnCMEPrint.Click
        Try
            If Me.FormTemplateContext.GE.RecordID < 1 Then
                MsgBox("Please save the record prior to running this process.", MsgBoxStyle.Information, "CME Print")
                Exit Sub
            End If

            Dim acsUniqueId As String = Me.FormTemplateContext.GE.GetValue("ACSUniqueId").ToString()

            Dim sURL As String = ""

            If Me.DataAction.UserCredentials.Server.ToLower = "aptify" Then
                'production
                sURL = "https://cmeapps.facs.org/cmeprint/printcmepdf.aspx?cmeid=" & acsUniqueId
            End If

            If Me.DataAction.UserCredentials.Server.ToLower = "stagingaptify2" Then
                'staging
                sURL = "https://test2.facs.org/cmeprint/PrintCMEPdf.aspx?cmeid=" & acsUniqueId
            End If

            If Me.DataAction.UserCredentials.Server.ToLower = "testaptifydb" Then
                'staging
                sURL = "https://test1.facs.org/cmeprint/PrintCMEPdf.aspx?cmeid=" & acsUniqueId
            End If

            Process.Start(sURL)

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
End Class