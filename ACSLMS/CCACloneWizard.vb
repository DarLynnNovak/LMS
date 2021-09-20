Option Explicit On
Option Strict On
Imports Aptify.Framework.ExceptionManagement


Public Class CCACloneWizard
    Implements Aptify.Framework.Application.IAptifyAction



    Public Sub DoAction(ByVal ApplicationObject As Aptify.Framework.Application.AptifyApplication, ByVal EntityID As Long) Implements Aptify.Framework.Application.IAptifyAction.DoAction
        DoAction(ApplicationObject, EntityID, -1, Nothing)
    End Sub

    Public Sub DoAction(ByVal ApplicationObject As Aptify.Framework.Application.AptifyApplication, ByVal EntityID As Long, ByVal ViewID As Long) Implements Aptify.Framework.Application.IAptifyAction.DoAction
        DoAction(ApplicationObject, EntityID, ViewID, Nothing)
    End Sub

    Public Sub DoAction(ByVal ApplicationObject As Aptify.Framework.Application.AptifyApplication, ByVal EntityID As Long, ByVal ViewID As Long, ByVal ParamArray SelectedItems() As Object) Implements Aptify.Framework.Application.IAptifyAction.DoAction
        Try
            Dim f As New cloneWizard
            f.Config(ApplicationObject, ViewID, SelectedItems)
            f.ShowDialog()

        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Sub

    Public Sub DoAction(ByVal ApplicationObject As Aptify.Framework.Application.AptifyApplication, ByVal EntityID As Long, ByVal ParamArray SelectedItems() As String) Implements Aptify.Framework.Application.IAptifyAction.DoAction
        DoAction(ApplicationObject, EntityID, -1, SelectedItems)

    End Sub
End Class
