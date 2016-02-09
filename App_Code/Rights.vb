Public Class Rights
    Private cCommon As Common
    Private cEnviro As Enviro
    Private cRightsColl As Collection

#Region " Shared variables "
    Public Shared CallbackView As String = "CBV"
    Public Shared CallbackEdit As String = "CBE"
#End Region

    Public Enum AccessLevelEnum
        AllAccess = 1
        SupervisorAccess = 2
        EnrollerAccess = 3
    End Enum
    Public ReadOnly Property RightsColl() As Collection
        Get
            Return cRightsColl
        End Get
    End Property
    Public Sub New(ByRef Enviro As Enviro, ByRef CurPage As Page)
        Dim dt As Data.DataTable

        Try

            cEnviro = Enviro
            cCommon = New Common
            cRightsColl = New Collection

            dt = Common.GetDT("SELECT CompanyID, Role, LocationID FROM Users WHERE UserID ='" & cEnviro.LoggedInUserID & "'", Enviro.DBHost, "UserManagement")
            If dt.Rows.Count = 0 Then
                CurPage.Response.Redirect("InsufficientRights.aspx")
            End If

            If dt.Rows(0)("CompanyID") = "BVI" Then
                If Common.InList(dt.Rows(0)("Role"), "ENROLLER, SUPERVISOR, IT, ADMIN, ADMIN LIC", True) Then
                    cRightsColl.Add(CallbackView)
                    cRightsColl.Add(CallbackEdit)
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #1770: RightsClass New. " & ex.Message, ex)
        End Try
    End Sub
    Public Shared Sub GetSecurityFlds(ByRef AccessLevel As AccessLevelEnum, ByRef LoginRole As RoleCatgyEnum)
        Try

            Select Case LoginRole
                Case RoleCatgyEnum.Other
                    AccessLevel = AccessLevelEnum.AllAccess
                Case RoleCatgyEnum.Supervisor
                    AccessLevel = AccessLevelEnum.SupervisorAccess
                Case RoleCatgyEnum.Enroller
                    AccessLevel = AccessLevelEnum.EnrollerAccess
            End Select

        Catch ex As Exception
            Throw New Exception("Error #1771: RightsClass GetSecurityFlds. " & ex.Message, ex)
        End Try
    End Sub
    Public Function HasSufficientRights(ByRef RightsRqd As String(), ByVal RedirectOnError As Boolean, ByRef CurPage As System.Web.UI.Page) As Boolean
        Dim i, j As Integer
        Dim Passed As Boolean
        Dim Result As Boolean

        Try

            For i = 0 To RightsRqd.GetUpperBound(0)
                For j = 1 To cRightsColl.Count
                    If cRightsColl(j) = RightsRqd(i) Then
                        Passed = True
                        Exit For
                    End If
                Next
                If Passed Then
                    Exit For
                End If
            Next

            If Passed Then
                Result = True
            Else
                If RedirectOnError Then
                    CurPage.Response.Redirect("InsufficientRights.aspx")
                End If
            End If

            Return Result

        Catch ex As Exception
            Throw New Exception("Error #1772: RightsClass HasSufficientRights. " & ex.Message, ex)
        End Try
    End Function
    Public Function HasThisRight(ByVal RightCd As String) As Boolean
        Dim i As Integer
        For i = 1 To cRightsColl.Count
            If cRightsColl(i) = RightCd Then
                Return True
            End If
        Next
        Return False
    End Function
End Class