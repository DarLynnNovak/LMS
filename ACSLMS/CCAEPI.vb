'Option Strict On
'Option Explicit On
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.BusinessLogic.ProcessPipeline
Imports Aptify.Framework.Application
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls
Imports System.Windows.Forms
Imports Aptify.Framework.FormTemplate



Public Class CCAEPI
    Inherits AptifyGenericEntity
    Private m_oApp As New AptifyApplication
    Private m_oDA As New DataAction
    Private m_oProps As New AptifyProperties
    Private errors As System.Text.StringBuilder = New System.Text.StringBuilder
    Public exception As String = ""
    Dim ID As Long
    Dim emailgroup As String = ""
    Dim ccemailgroup As String = ""
    Dim accountingemailgroup As String = ""
    Dim icemailgroup As String = ""


    Public Overridable ReadOnly Property Application() As AptifyApplication
        Get
            Return m_oApp
        End Get
    End Property
    Public Overrides Function Save(AllowGUI As Boolean, ByRef ErrorString As String, TransactionID As String) As Boolean

        ID = Me.RecordID


        Try
            If ID > 0 Then
                Return MyBase.Save(AllowGUI, ErrorString, TransactionID)
                'Include additional Save here
            End If

            'Return MyBase.Save(AllowGUI, ErrorString, TransactionID)


        Catch ex As Exception

            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            Return False
        End Try
    End Function


    'Public Overrides Function Validate(ByRef ErrorString As String) As Boolean

    '    Dim bResult As Boolean = False
    '    Dim da As New DataAction
    '    ID = Me.RecordID


    '    Try

    '    Catch ex As Exception
    '        Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
    '    End Try
    'End Function


End Class
