Public Class Enviro
#Region " Declarations "
    Private cAppTimedOut As Boolean
    Private cLoggedInUserID As String
    Private cLoggedInUserName As String
    Private cComputerId As String
    Private cLoginLocationID As String
    Private cLoginRoleCatgy As RoleCatgyEnum
    Private cInit As Boolean
    Private cLastPageLoad As DateTime
    Private cLoginIP As String
    Private cClientUTCOffsetHours As Integer
    Private cSessionID As String
    Private cBrowserName As String
    Private cIsAuthenticated As Boolean
    Private cLogColl17 As New Collection
#End Region

#Region " Shared variables "
    'Private Shared cVersionNumber As String = "1.00.004"
    'Private Shared cVersionNumber As String = "1.00.004 spec"
    'Private Shared cVersionNumber As String = "1.00.005 spec"  ' Added hack to compensate for incorrect EnrollWinCode for ASCOM. Added emp name to escalation email.
    'Private Shared cVersionNumber As String = "1.00.006"  ' Added temporary IsPurged indicator to EmpIDLookup.
    Private Shared cVersionNumber As String = "1.00.007"  ' Allow checkout if currently checkedout to myself (needs improvement when time allows).
    Private Shared cBuildNum As String = "20160205.2"
    Private Shared cVSPlatform As String = "VS2015"


    Private Shared cDBHost As String = ConfigurationManager.AppSettings("DBHost")
    Private Shared cDefaultDatabase As String = "Callback"
    Private Shared cDBTimeout As Integer = 10
    Private Shared cAppTimeout As Integer = 3600   ' 60 minutes

    'Private Shared cMaxDisplayRecords As Integer = 500
    'Private Shared cExcessiveRecordAmount As Integer = 1000
    'Private Shared cRecordMaximum As Integer = 5000

    Private Shared cMaxDisplayRecords As Integer = 50000
    Private Shared cExcessiveRecordAmount As Integer = 100000
    Private Shared cRecordMaximum As Integer = 500000

    'Private ccc As New Class2
    Private Shared cLogRetentionDays As Integer = 30
    Private Shared cApplicationPath As String
    Private Shared cAcctNotation As String = "YepdU5s+CTSCHBYd9Kb/Sw=="
    Private Shared cF6Sequencer As String = "o+0RC/MOgd7Yy6+VeU/xRg=="
    Private Shared cSRACode As String = "/ngOQI6g0mNaT35vriINpkM5hNm/ckMzUgIx1HFeyQg="
    Private Shared cExcelTemplateFileName As String = "Rpt_SupvMaster_Template.xls"
    Private Shared cConnectionStringTemplate As String = "User ID=" & AESEncryption.AES_Decrypt(cF6Sequencer, cSRACode) & ";Password=" & AESEncryption.AES_Decrypt(cAcctNotation, cSRACode) & ";database=|;server="
#End Region

    'Public Sub New(ByVal ApplicationPath As String)
    '    If ApplicationPath.Substring(ApplicationPath.Length - 1) = "\" Then
    '        ApplicationPath = ApplicationPath.Substring(0, ApplicationPath.Length - 1)
    '    End If
    '    cApplicationPath = ApplicationPath
    'End Sub

#Region " Instance properties"
    Public Property Init() As Boolean
        Get
            Return cInit
        End Get
        Set(ByVal Value As Boolean)
            cInit = Value
        End Set
    End Property
    Public Property SessionID() As String
        Get
            Return cSessionID
        End Get
        Set(ByVal Value As String)
            cSessionID = Value
        End Set
    End Property
    Public Property LastPageLoad() As DateTime
        Get
            Return cLastPageLoad
        End Get
        Set(ByVal Value As DateTime)
            cLastPageLoad = Value
        End Set
    End Property
    Public Property LoggedInUserID() As String
        Get
            If cLoggedInUserID = Nothing Then
                Return GetAlternateLoggedInUserID()
            Else
                Return cLoggedInUserID
            End If
        End Get
        Set(ByVal Value As String)
            cLoggedInUserID = Value
        End Set
    End Property
    Private Function GetAlternateLoggedInUserID() As String
        Dim UserID As String
        UserID = HttpContext.Current.User.Identity.Name.ToString
        UserID = UserID.Substring(InStr(UserID, "\", CompareMethod.Binary))
        Return UserID
    End Function

    Public Property LoggedInUserName() As String
        Get
            Return cLoggedInUserName
        End Get
        Set(ByVal Value As String)
            cLoggedInUserName = Value
        End Set
    End Property
    Public Property LoginIP() As String
        Get
            Return cLoginIP
        End Get
        Set(ByVal Value As String)
            cLoginIP = Value
        End Set
    End Property

    Public Property LoginLocationID() As String
        Get
            Return cLoginLocationID
        End Get
        Set(ByVal Value As String)
            cLoginLocationID = Value
        End Set
    End Property

    Public Property LogInRoleCatgy() As RoleCatgyEnum
        Get
            Return cLoginRoleCatgy
        End Get
        Set(ByVal Value As RoleCatgyEnum)
            cLoginRoleCatgy = Value
        End Set
    End Property

    'Public Property TestProd() As TestProdEnum
    '    Get
    '        Return cTestProd
    '    End Get
    '    Set(ByVal Value As TestProdEnum)
    '        cTestProd = Value
    '    End Set
    'End Property
    Public Property AppTimedOut() As Boolean
        Get
            Return cAppTimedOut
        End Get
        Set(ByVal Value As Boolean)
            cAppTimedOut = Value
        End Set
    End Property

    Public Property ClientUTCOffsetHours() As Integer
        Get
            Return cClientUTCOffsetHours
        End Get
        Set(ByVal Value As Integer)
            cClientUTCOffsetHours = Value
        End Set
    End Property
    Public Property ApplicationPath() As String
        Set(ByVal value As String)
            If value.Substring(value.Length - 1) = "\" Then
                value = value.Substring(0, value.Length - 1)
            End If
            cApplicationPath = value
        End Set
        Get
            Return cApplicationPath
        End Get
    End Property

    Public ReadOnly Property LogColl17() As Collection
        Get
            Return cLogColl17
        End Get
    End Property
#End Region

#Region " Shared Properties "
    Public Shared ReadOnly Property VersionNumber() As String
        Get
            Return cVersionNumber
        End Get
    End Property
    Public Shared ReadOnly Property BuildNum() As String
        Get
            Return cBuildNum
        End Get
    End Property
    Public Shared ReadOnly Property MaxDisplayRecords() As Integer
        Get
            Return cMaxDisplayRecords
        End Get
    End Property
    Public Shared ReadOnly Property ExcessiveRecordAmount() As Integer
        Get
            Return cExcessiveRecordAmount
        End Get
    End Property
    Public Shared ReadOnly Property RecordMaximum() As Integer
        Get
            Return cRecordMaximum
        End Get
    End Property
    Public Shared ReadOnly Property LogRetentionDays() As Integer
        Get
            Return cLogRetentionDays
        End Get
    End Property
    Public Shared ReadOnly Property ExcelTemplateFileName() As String
        Get
            Return cExcelTemplateFileName
        End Get
    End Property
    Public Shared ReadOnly Property DBHost() As String
        Get
            Return ConfigurationManager.AppSettings("DBHost")
        End Get
    End Property
    Public Shared ReadOnly Property AppTimeout() As Integer
        Get
            Return cAppTimeout
        End Get
    End Property
    Public Shared ReadOnly Property DefaultDatabase() As String
        Get
            Return cDefaultDatabase
        End Get
    End Property
    Public Shared ReadOnly Property LogFileFullPath() As String
        Get
            Return cApplicationPath & "\CallbackLogFile.txt"
        End Get
    End Property

    Public Shared ReadOnly Property QueryDownloadDir() As String
        Get
            Return cApplicationPath & "\Temp"
        End Get
    End Property
    Public Property BrowserName As String
        Get
            Return cBrowserName
        End Get
        Set(value As String)
            cBrowserName = value
        End Set
    End Property
    Public Property IsAuthenticated As Boolean
        Get
            Return cIsAuthenticated
        End Get
        Set(value As Boolean)
            cIsAuthenticated = value
        End Set
    End Property
#Region " ConnectionString "
    Public Shared ReadOnly Property ConnectionStringTemplate() As String
        Get
            Return cConnectionStringTemplate
        End Get
    End Property
    Public Shared Function GetConnectionString(ByVal DBHost As String, ByVal Database As String) As String
        Return Replace(cConnectionStringTemplate, "|", Database) & DBHost
    End Function

    Public Shared Function GetConnectionString() As String
        Return Replace(cConnectionStringTemplate, "|", cDefaultDatabase) & DBHost
    End Function
#End Region
#End Region

    'Public Function TestIDSelect(ByVal ClientID As String, ByVal EmpID As String, ByVal ClientAddSideQuotes As Boolean) As String
    '    Dim Results As String
    '    Select Case cTestProd
    '        Case TestProdEnum.Test
    '            Return String.Empty
    '        Case TestProdEnum.Production
    '            If ClientAddSideQuotes Then
    '                Results = " AND dbo.ufn_IsTestID('" & ClientID & "', " & EmpID & ") = 0 "
    '            Else
    '                Results = " AND dbo.ufn_IsTestID(" & ClientID & ", " & EmpID & ") = 0 "
    '            End If
    '    End Select
    '    Return Results
    'End Function

    Public Sub MakeCookie(ByVal Page As Page)
        Dim SessionID As String
        Dim SessionCookie As HttpCookie
        'Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session

        SessionID = cSessionID
        SessionCookie = New HttpCookie(SessionID)
        Page.Response.Cookies.Add(SessionCookie)
    End Sub

    Public Sub AuthenticateRequest(ByRef Page As System.Web.UI.Page)
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        'Dim ts As TimeSpan

        Try

            ' ___ Keep this handy
            'If Page.ToString = "ASP.Trn_CallMonitorMaintain_aspx" Then
            '    Timeout = cTrn_MonitorMaintainTimeout
            'Else
            '    Timeout = cAppTimeout
            'End If



            ' ___ Page entry
            If cSessionID = Nothing Then
                Throw New Exception("timeout1")
            End If

            '' ___ Cookie
            'If IsNothing(Page.Request.Cookies.Item(cSessionID)) Then
            '    Throw New Exception("No Cookie Present.")
            'End If

            '' ___ Timeout
            'ts = Date.Now.Subtract(cLastPageLoad)
            'If ts.TotalSeconds > cAppTimeout Then
            '    Throw New Exception("timeout2")
            'End If


            ' ___ If we made it this far, we're good!
            cLastPageLoad = Date.Now

        Catch ex As Exception
            Throw New Exception("Error #820: Enviro AuthenticateRequest. " & ex.Message)
        End Try
    End Sub
End Class