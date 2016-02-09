Public Class ErrorObj

    ' //////////////////////////////////////////////////
    '
    ' Microsoft bug alert
    ' Error handling uses a session object, called
    ' ErrorArgs. If the error handling is written such
    ' that Global.Application_Error is used at all,
    ' it causes the loss of the session objects. This
    ' application processes errors in the ErrorObj 
    ' as an alternative.
    ' //////////////////////////////////////////////////

#Region " Methods "
    Public Sub New(ByRef exMessage As String)
        Dim ErrorMessage As String
        Dim Box As String()
        Try
            Box = exMessage.ToString.Split("Error #")
            ErrorMessage = "E" & Box(Box.GetUpperBound(0))
            If InStr(ErrorMessage, "Thread was being aborted.") > 0 Then
                Exit Sub
            End If
            HandleError(ErrorMessage, False)
        Catch ex As Exception
        End Try
    End Sub

    Public Sub HandleError(ByVal RawMessage As String, ByVal Shutdown As Boolean)
        Dim mCommon As Common
        Dim mEnviro As Enviro
        Dim HeaderMessage As String
        Dim ErrorArgs As ErrorArgs
        Dim ErrorMessage As String
        Dim WriteToLog As Boolean
        Dim SendEmail As Boolean
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim Report As New Report
        Dim Response As System.Web.HttpResponse
        Dim Sql As New System.Text.StringBuilder

        Try

            ' ___ Get the response object
            Response = HttpContext.Current.Response

            ' ___ Get Enviro from session
            mEnviro = SessionObj("Enviro")
            mCommon = New Common

            ' ___ Get the ErrorArgs from session
            ErrorArgs = SessionObj("ErrorArgs")

            'If mEnviro.SessionID = Nothing Then
            '    mEnviro.SessionID = Guid.NewGuid.ToString  ' Date.Now.ToUniversalTime.AddHours(-5).ToString("ddHHmmss")
            'End If

            '' ___ Write to ErrorLog table
            'Sql.Append("exec usp_WriteToErrorLog " & mEnviro.SessionID & ", ")
            'Sql.Append("'" & Enviro.BuildNum & "', ")
            'Sql.Append("'" & Enviro.DBHost & "', ")
            'Sql.Append("'" & mEnviro.LoggedInUserID & "', ")
            'RawMessage = Left(RawMessage, 1000).Replace("'", "*")
            'Sql.Append("'" & RawMessage & "'")
            'Common.ExecuteNonQuery(Sql.ToString)

            ' ___ Get the ErrorType and ErrorMessage
            If InStr(RawMessage, "UnableToConnect", CompareMethod.Text) > 0 Then
                ErrorMessage = "Callback is unable to establish a connection with " & Enviro.DBHost & ". Please report this problem to your supervisor or IT support person. If no one is available, submit an issue to the HelpDesk as a last resort. You will not be able to work with the Callback until the database connection is restored."
                WriteToLog = True
                SendEmail = False
                Shutdown = True
            ElseIf InStr(RawMessage, "timeout1", CompareMethod.Text) > 0 Then
                ErrorMessage = "Callback has either timed out or you have tried to launch from an incorrect page. Please close the application and attempt to launch using this link: " & GetLink()
                WriteToLog = False
                SendEmail = False
                Shutdown = True
            ElseIf InStr(RawMessage, "timeout2", CompareMethod.Text) > 0 Then
                ErrorMessage = "Callback has timed out. Please close the application and attempt to launch using this link: " & GetLink()
                WriteToLog = False
                SendEmail = False
                Shutdown = True
            Else
                ErrorMessage = RawMessage
                WriteToLog = True
                SendEmail = True
                Shutdown = True
            End If

            ' ___ Send the email
            HeaderMessage = Report.Report(ErrorMessage, WriteToLog, SendEmail, Shutdown)
            ErrorArgs.HeaderMessage = HeaderMessage
            ErrorArgs.ErrorMessage = ErrorMessage

            ' ___ Handle possible redirect to error page and application shutdown.
            If Shutdown Then
                ErrorMessage = Replace(ErrorMessage, "#", "[sharp]")
                ErrorMessage = Replace(ErrorMessage, vbCrLf, "~")
                ErrorMessage = Replace(ErrorMessage, Chr(10), "~")
                If ErrorMessage.Length + HeaderMessage.Length > 2000 Then
                    ErrorMessage = ErrorMessage.Substring(0, (2000 - HeaderMessage.Length))
                End If

                Response.Redirect("ErrorPage.aspx?ErrorMessage=" & ErrorMessage & "&HeaderMessage=" & HeaderMessage)
            End If

        Catch ex As Exception
            'Throw New Exception("Error #2303: ErrorObj HandleError. " & ex.Message, ex)
        End Try
    End Sub

    Private Function GetLink() As String
        Dim Results As String = Nothing
        Select Case ConfigurationManager.AppSettings("DBHost").ToString.ToLower
            Case "hbg-sql"
                Results = "http://netserver.benefitvision.com/Callback/"
            Case "hbg-tst"
                Results = "http://test.benefitvision.com/Callback/"
            Case "training"
                Results = "http://train.benefitvision.com/Callback/"
        End Select
        Return Results
    End Function
#End Region
End Class

'Public Sub New(ByRef Exception As Exception, ByVal ErrorTopLevel As String)
'    Dim CurException As Exception
'    Dim Coll As New Collection
'    Dim ExceptionMessage As String

'    Try

'        ' ___ Extract the error
'        If Exception.InnerException Is Nothing Then
'            'Coll.Add(ErrorTopLevel)
'            ExceptionMessage = ErrorTopLevel
'        Else
'            CurException = Exception
'            While Not (CurException Is Nothing)
'                Coll.Add(CurException.Message)
'                CurException = CurException.InnerException
'            End While
'            ExceptionMessage = Coll(Coll.Count - 1)
'        End If


'        If InStr(ExceptionMessage, "Thread was being aborted.") > 0 Then
'            Exit Sub
'        End If

'        HandleError(ExceptionMessage, True)
'    Catch ex As Exception
'        'Throw New Exception("Error #2302: ErrorObj New. " & ex.Message, ex)
'    End Try
'End Sub

'Public Sub New(ByRef ErrorMessage As String)
'    Try
'        HandleError(ErrorMessage, False)
'    Catch ex As Exception
'    End Try
'End Sub