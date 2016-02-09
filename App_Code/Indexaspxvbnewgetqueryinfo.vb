'Option Explicit On
'Imports System.IO
'Imports System.Data

''DELETE callbacknote
''DELETE callbackphone
''DELETE callbackattempt
''DELETE callbackmaster

'' http://www.netdenizen.com/buttonmill/glassy.php
'' font: lucida console 7

'' Javascript: add and remove elements
'' http://www.dustindiaz.com/add-and-remove-html-elements-dynamically-with-javascript/

''Icons
''http://www.iconfinder.com/search/?q=iconset%3Aapp_iconset_creative_nerds

'Partial Class Index
'    Inherits System.Web.UI.Page

'#Region " Declarations "
'    Private cEnviro As Enviro
'    Private cCommon As Common
'    Private cRights As Rights
'    Private cIndexSess As IndexSession
'    Private cDGCallAttempt As DG
'#End Region

'    'Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
'    '    Dim LoggedInUserID

'    '    LoggedInUserID = HttpContext.Current.User.Identity.Name.ToString
'    '    litMessage.Text = "<script  type='text/javascript'>alert('" & LoggedInUserID & "')</script>"

'    'End Sub

'#Region " Page load "
'    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

'        '*********************************************************************************************************************
'        ' *** Periodicaly run the query at the top of helpers.vb that checks for multiple nonpurged callbackmaster records ***
'        '*********************************************************************************************************************

'        Dim AltLoginInd As Boolean
'        Dim AltLogin As String
'        Dim Message As String = Nothing

'        Try

'            ' ___ Get Enviro from Session
'            cEnviro = CType(Session("Enviro"), Enviro)

'            ' ___ Instantiate Common
'            cCommon = New Common

'            ' ___ Get the enroller session object
'            cIndexSess = CType(Session("IndexSession"), IndexSession)

'            If Page.IsPostBack Then
'                If Request.Form("hdAction") = "AltLoginAuthenticated" Then
'                    AltLoginInd = True
'                    AltLogin = Common.GetBootsAPuppy(Request.Form("txtAltLogin"))
'                    LoadEnviro(True, AltLogin)
'                    If Not cEnviro.IsAuthenticated Then
'                        litEnviro.Text = "<Input type=""hidden"" id=""hdAltLogin"" value=""" & Message & """ />"
'                        'litJumpToAnchor.Text = "function OnLoadEvents() {alert(" & Message & "}"
'                    End If
'                End If
'            End If

'            ' ___ Load the application settings/restore session
'            If Not cEnviro.Init Then
'                LoadEnviro(False, Nothing)
'            Else
'                cEnviro.AuthenticateRequest(Me)
'            End If


'            If cEnviro.IsAuthenticated Then
'                PageLoadDetails(AltLoginInd)
'            End If

'        Catch ex As Exception
'            Dim ErrorObj As New ErrorObj(ex, "Error #102: Index Page_Load. " & ex.Message)
'        End Try
'    End Sub

'    Private Sub PageLoadDetails(ByVal AltLoginInd As Boolean)
'        Dim DGCall As DG
'        Dim Action As String
'        Dim Results As Results
'        Dim PageMode As PageMode
'        Dim sbHiddens As New System.Text.StringBuilder

'        Try

'            ' ___ Right check
'            cRights = New Rights(cEnviro, Page)
'            Dim RightsRqd(0) As String
'            RightsRqd.SetValue(Rights.CallbackView, 0)
'            cRights.HasSufficientRights(RightsRqd, True, Page)
'            lblCurrentRights.Text = cCommon.GetCurRightsAndTopicsHidden(cRights.RightsColl)


'            ' ___ Get the page mode 
'            If AltLoginInd Then
'                PageMode = PageMode.Initial
'            Else
'                PageMode = Common.GetPageMode(Page, cIndexSess)
'            End If

'            ' ___ ClearLog
'            ClearLog(PageMode)

'            ' ___ Load the page session variables
'            LoadVariables(PageMode)

'            ' ___ Initialize the datagrid
'            DGCall = DefineDataGrid("ticket")
'            AddHandler DGCall.ChildDTRequest, AddressOf HandleChildDTRequest
'            cDGCallAttempt = DefineDataGrid("CallAttempt")

'            ' ___ Execute action
'            Select Case PageMode
'                Case PageMode.Initial
'                    DisplayPage(PageMode, DGCall, DG.OrderByType.Initial)

'                Case PageMode.Postback
'                    Action = Request.Form("hdAction")

'                    Select Case Action
'                        Case "Sort"
'                            DisplayPage(PageMode, DGCall, DG.OrderByType.Field, Request.Form("hdSortField"))

'                        Case "ApplyFilter"
'                            DisplayPage(PageMode, DGCall, DG.OrderByType.Recurring)

'                        Case "DeleteCallback"
'                            Results = DeleteCallbackRecord()
'                            DisplayPage(PageMode, DGCall, DG.OrderByType.Recurring)
'                            litMessage.Text = "<script language='javascript'>alert('" & Common.ToJSAlert(Results.Message) & "')</script>"

'                        Case "DeleteCallbackAttempt"
'                            Results = DeleteCallbackAttemptRecord()
'                            DisplayPage(PageMode, DGCall, DG.OrderByType.Recurring)
'                            litMessage.Text = "<script language='javascript'>alert('" & Common.ToJSAlert(Results.Message) & "')</script>"

'                        Case "ReqSubTblOpen", "ReqSubTblClose"
'                            DisplayPage(PageMode, DGCall, DG.OrderByType.Recurring)

'                    End Select

'                Case PageMode.ReturnFromChild, PageMode.CalledByOther
'                    DisplayPage(PageMode, DGCall, DG.OrderByType.ReturnToPage)
'                    'If cIndexSess.PageReturnOnLoadMessage <> Nothing Then
'                    '    litMessage.Text = "<script language='javascript'>alert('" & cIndexSess.PageReturnOnLoadMessage & "')</script>"
'                    '    'litMessage.Text = LittleMessage()
'                    '    cIndexSess.PageReturnOnLoadMessage = Nothing
'                    'End If
'            End Select

'            If cIndexSess.PageReturnOnLoadMessage <> Nothing Then
'                litMessage.Text = "<script language='javascript'>alert('" & cIndexSess.PageReturnOnLoadMessage & "')</script>"
'                cIndexSess.PageReturnOnLoadMessage = Nothing
'            End If

'            ' ___ Display enviroment
'            PageCaption.Text = Common.GetPageCaption
'            sbHiddens.Append("<input type='hidden' id='hdBrowserName' value='" & cEnviro.BrowserName & "'/><input type='hidden' id='hdLoggedInUserID' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' name='hdDBHost'  value='" & Enviro.DBHost & "'/><input type='hidden' id='hdBuildNum'  value='" & Enviro.BuildNum & "'/>")


'            If cIndexSess.DGSess.JumpToRecID = String.Empty Then
'                'litJumpToAnchor.Text = "function OnLoadEvents() {var j=4}"
'                sbHiddens.Append("<input type=""hidden"" id=""hdJumpToRecInd"" value=""0""/>")

'            Else
'                sbHiddens.Append("<input type=""hidden"" id=""hdJumpToRecInd"" value=""1""/>")
'                sbHiddens.Append("<input type=""hidden"" id=""hdJumpToRecID"" value=""" & cIndexSess.DGSess.JumpToRecID & """/>")
'                'litJumpToAnchor.Text = "function OnLoadEvents() {window.location.hash='#" & cIndexSess.DGSess.JumpToRecID & "'}"
'            End If


'            'If cIndexSess.DGSess.JumpToRecID <> String.Empty Then
'            '    'litJumpToAnchor.Text = "window.location.hash='#" & cIndexSess.DGSess.JumpToRecID & "'"
'            '    litJumpToAnchor.Text = "function OnLoadEvents() {window.location.hash='#" & cIndexSess.DGSess.JumpToRecID & "'}"
'            'End If

'            'If cIndexSess.DGSess.JumpToRecID <> String.Empty Then
'            '    sbHiddens.Append("<input type=""hidden"" id=""hdJumpToRecInd"" value=""1""/>")
'            '    sbHiddens.Append("<input type=""hidden"" id=""hdJumpToRecID"" value=""" & cIndexSess.DGSess.JumpToRecID & """/>")
'            'Else
'            '    sbHiddens.Append("<input type=""hidden"" id=""hdJumpToRecInd"" value=""0""/>")
'            'End If

'            litEnviro.Text = sbHiddens.ToString
'        Catch ex As Exception
'            Throw New Exception("Error #116: Index PageLoadDetails. " & ex.Message, ex)
'        End Try
'    End Sub

'    Private Function DeleteCallbackRecord() As Results
'        Dim QueryPack As DBase.QueryPack
'        Dim Sql As String
'        Dim MyResults As New Results

'        Try

'            Sql = "exec usp_DeleteCallback @CallbackID = '" & cIndexSess.CallbackID & "'"
'            QueryPack = Common.GetDTWithQueryPack(Sql)

'            If QueryPack.Success Then
'                If QueryPack.dt.Rows(0)("ErrorInd") = 0 Then
'                    MyResults.Success = True
'                Else
'                    MyResults.Success = False
'                End If
'                MyResults.Message = QueryPack.dt.Rows(0)(1)
'            Else
'                MyResults.Success = False
'                MyResults.Message = QueryPack.TechErrMsg
'            End If
'            Return MyResults

'        Catch ex As Exception
'            Throw New Exception("Error #105: Index DeleteCallbackRecord. " & ex.Message, ex)
'        End Try
'    End Function

'    Private Function DeleteCallbackAttemptRecord() As Results
'        Dim QueryPack As DBase.QueryPack
'        Dim Sql As String
'        Dim MyResults As New Results

'        Try

'            'Sql = "DELETE CallbackAttempt WHERE CallbackAttemptID = " & cIndexSess.CallbackAttemptID
'            Sql = "UPDATE CallbackAttempt SET LogicalDelete = 1, ChangeDate = '" & Common.GetServerDateTime & "' WHERE CallbackAttemptID = " & cIndexSess.CallbackAttemptID
'            QueryPack = Common.GetDTWithQueryPack(Sql)
'            If QueryPack.Success Then
'                MyResults.Success = True
'                MyResults.Message = "Record deleted."
'            Else
'                MyResults.Success = False
'                MyResults.Message = QueryPack.TechErrMsg
'            End If

'            Return MyResults

'        Catch ex As Exception
'            Throw New Exception("Error #106: Index DeleteCallbackAttemptRecord. " & ex.Message, ex)
'        End Try
'    End Function

'    Private Sub LoadEnviro(ByVal AltLoginInd As Boolean, ByVal AltLogin As String)
'        Dim LoggedInUserID As String = Nothing
'        'Dim HTTPHost As String
'        Dim dt As DataTable
'        Dim Querypack As DBase.QueryPack
'        Dim SessionID As String

'        Try

'            cEnviro.Init = True
'            SessionID = Guid.NewGuid.ToString
'            cEnviro.SessionID = SessionID
'            cEnviro.LastPageLoad = Date.Now

'            ' ___ LoggedInUserID
'            If AltLoginInd Then
'                LoggedInUserID = AltLogin
'            Else
'                If System.Environment.MachineName.ToUpper = "WADEV" Then
'                    LoggedInUserID = "rbluestein"
'                Else
'                    'LoggedInUserID = HttpContext.Current.User.Identity.Name.ToString
'                    'LoggedInUserID = LoggedInUserID.Substring(InStr(LoggedInUserID, "\", CompareMethod.Binary))
'                    'LoggedInUserID = LoggedInUserID.ToLower
'                    '' LoggedInUserID = "lwirrick"
'                    LoggedInUserID = Request.ServerVariables("REMOTE_USER")
'                    LoggedInUserID = LoggedInUserID.Substring(InStr(LoggedInUserID, "\", CompareMethod.Binary))
'                    LoggedInUserID = LoggedInUserID.ToLower
'                End If
'            End If

'            cEnviro.LoggedInUserID = LoggedInUserID

'            ' ___ LoginIP, app path
'            cEnviro.LoginIP = Page.Request.ServerVariables("REMOTE_ADDR")
'            cEnviro.ApplicationPath = Page.Server.MapPath(Page.Request.ApplicationPath)

'            ' ___ Test database connection
'            Querypack = Common.GetDTWithQueryPack("SELECT Count (*) FROM Callback..CallbackMaster")
'            If Not Querypack.Success Then
'                Throw New Exception("Error #114a: LoadEnviro. UnableToConnect")
'            End If

'            ' ___ Get user data
'            dt = Common.GetDT("SELECT Name = LastName + ', ' + FirstName, Role, LocationID FROM UserManagement..Users WHERE UserID = '" & LoggedInUserID & "'")
'            If dt.Rows.Count = 0 Then
'                Throw New Exception("Error #114b: LoadEnviro. Unable to authenticate user in UserManagement database")
'            Else
'                cEnviro.IsAuthenticated = True
'            End If

'            cEnviro.LoggedInUserName = CType(dt.Rows(0)("Name"), System.String)
'            cEnviro.LoginLocationID = CType(dt.Rows(0)("LocationID"), System.String)

'            cEnviro.LogColl17.Add("SessionID: " & SessionID)
'            cEnviro.LogColl17.Add("LoggedInUserID: " & cEnviro.LoggedInUserName)

'            Select Case CType(dt.Rows(0)("Role"), System.String).ToUpper
'                Case "IT", "ADMIN", "ADMIN LIC"
'                    cEnviro.LogInRoleCatgy = RoleCatgyEnum.Other
'                Case "ENROLLER"
'                    cEnviro.LogInRoleCatgy = RoleCatgyEnum.Enroller
'                Case "SUPERVISOR"
'                    cEnviro.LogInRoleCatgy = RoleCatgyEnum.Supervisor
'            End Select

'            Select Case Request.Browser.Browser.ToLower
'                Case "ie", "internet explorer", "internetexplorer"
'                    cEnviro.BrowserName = "IE"
'                Case "chrome"
'                    cEnviro.BrowserName = "Chrome"
'                Case "firefox"
'                    cEnviro.BrowserName = "Firefox"
'                Case Else
'                    cEnviro.BrowserName = Request.Browser.Browser
'            End Select

'            ' ___ Cookie
'            cEnviro.MakeCookie(Me)

'        Catch ex As Exception
'            Throw New Exception("Error #114c: Index LoadEnviro. " & ex.Message, ex)
'        End Try
'    End Sub

'    Private Sub LoadVariables(ByVal PageMode As PageMode)

'        Try

'            Select Case PageMode
'                Case PageMode.Initial
'                    ' cIndexSess.Filter.DateRange = "|||||"

'                    ' ___ Initialize SubTable
'                    cIndexSess.DGSess.SubTableInd = "0"

'                Case PageMode.CalledByOther

'                Case PageMode.ReturnFromChild

'                Case PageMode.Postback

'                    Select Case Request.Form("hdAction")
'                        Case "ReqSubTblOpen"
'                            cIndexSess.DGSess.SubTableInd = "1"
'                            cIndexSess.DGSess.ActiveSubTableRecID = Request.Form("hdActiveSubTableRecID")
'                            cIndexSess.DGSess.JumpToRecID = Request.Form("hdActiveSubTableRecID")

'                        Case "ReqSubTblClose"
'                            cIndexSess.DGSess.SubTableInd = "0"
'                            cIndexSess.DGSess.ActiveSubTableRecID = String.Empty

'                        Case "ApplyFilter"
'                            cIndexSess.Filter.DateRange = Request.Form("hdDateRange")
'                            cIndexSess.Filter.EmpName = Request.Form("txtEmpName").Trim
'                            cIndexSess.Filter.ClientID = Request.Form("ddClientID")
'                            cIndexSess.Filter.State = Request.Form("ddState")
'                            cIndexSess.Filter.CallPurposeCode = Request.Form("ddccp.CallPurposeCode")
'                            cIndexSess.Filter.StatusCodeAdjDetail = Request.Form("ddStatusCodeAdjDetail")
'                            cIndexSess.Filter.PriorityTagInd = Request.Form("ddPriorityTagInd")
'                            cIndexSess.Filter.PreferSpanishInd = Request.Form("ddPreferSpanishInd")

'                        'Case "ExistingCallback"
'                        '    cIndexSess.CallbackID = Request.Form("hdCallbackID")
'                        '    cIndexSess.DGSess.JumpToRecID = Request.Form("hdCallbackID")

'                        Case "DeleteCallback"
'                            cIndexSess.CallbackID = Request.Form("hdCallbackID")

'                        Case "DeleteCallbackAttempt"
'                            cIndexSess.CallbackAttemptID = Request.Form("hdCallbackAttemptID")

'                        'Case "ExistingCallbackAttempt"
'                        '    cIndexSess.CallbackAttemptID = Request.Form("hdCallbackAttemptID")

'                        Case Else
'                            ' no action

'                    End Select

'                    cEnviro.LogColl17.Add("LoadVariablesPageMode|DateRange: " & PageMode.ToString & "|" & cIndexSess.Filter.DateRange)

'            End Select

'        Catch ex As Exception
'            Throw New Exception("Error #103: Index LoadVariables. " & ex.Message, ex)
'        End Try
'    End Sub

'    Private Sub ClearLog(ByVal PageMode As PageMode)
'        Dim FileInfo As FileInfo
'        Dim FileDays As Integer
'        Dim NowDays As Integer

'        Try

'            If PageMode <> PageMode.Initial Then
'                Exit Sub
'            End If

'            ' ___ Delete the log file if it is more than 10 days old..
'            If System.IO.File.Exists(Enviro.LogFileFullPath) Then

'                FileInfo = New System.IO.FileInfo(Enviro.LogFileFullPath)
'                FileDays = FileInfo.CreationTimeUtc.Subtract(#1/1/2007#).Days()
'                NowDays = Date.Now.Subtract(#1/1/2007#).Days()

'                If NowDays - FileDays > (Enviro.LogRetentionDays - 1) Then
'                    Dim FileStream As New System.IO.FileStream(Enviro.LogFileFullPath, FileMode.Create, FileAccess.Write, FileShare.None)
'                    FileStream.Close()
'                    FileInfo.CreationTime = Date.Now.ToUniversalTime.AddHours(-5)
'                End If

'            End If

'        Catch ex As Exception
'            Throw New Exception("Error #108: Index ClearLog. " & ex.Message, ex)
'        End Try
'    End Sub

'    Private Function DefineDataGrid(ByVal Entity As String) As DG
'        Dim DG As DG = Nothing

'        Try

'            If Entity = "ticket" Then

'                'DG = New DG("CallbackID", cCommon, cRights, True, "EmbeddedTableDef", "dbo.ufn_GetLastActivityDate(CallbackID)", DG.DefaultSortDirectionEnum.Descending)
'                DG = New DG("CallbackID", cCommon, cRights, True, "EmbeddedTableDef", "CreationDate", DG.DefaultSortDirectionEnum.Descending)
'                If cRights.HasThisRight(Rights.CallbackEdit) Then
'                    DG.AddNewButton(Nothing)
'                End If

'                DG.AddDataBoundColumn("TicketNumber", "TicketNumber", "Ticket<br />Num", "TicketNumber", True, Nothing, Nothing, "align='left'")

'                DG.AddDateColumn("CreationDate", "CreationDate", "Call<br />Date", "CreationDate", True, "SpecialFormat3", Nothing, "width= '120px' align='left'")
'                DG.AddDataBoundColumn("EmpName", "EmpName", "Employee", "EmpName", True, Nothing, Nothing, "width='100px' align='left'")
'                DG.AddDataBoundColumn("ClientID", "ClientID", "Client", "ClientID", True, Nothing, Nothing, "align='left'")
'                DG.AddDataBoundColumn("State", "State", "State", "State", True, Nothing, Nothing, "align='left'")
'                DG.AddDataBoundColumn("CallPurposeCode", "CallPurposeDescription", "Call<br />Type", "ccp.CallPurposeDescription", True, Nothing, Nothing, "align='left'")

'                DG.AddDataBoundColumn("StatusCodeAdjDetail", "StatusCodeAdjDetail", "Status", "dbo.ufn_GetStatusCodeAdj(1, cm.CallbackID, getDate())", True, Nothing, Nothing, "align='left'")

'                DG.AddBooleanColumn("PriorityTagInd", "PriorityTagInd", "Priority<br>Tag", "PriorityTagInd", True, "1", "Yes", "No", Nothing, "align='left'")
'                DG.AddDataBoundColumn("BestTime", "BestTime", "Best Time<br />To Call", "cp.BestTime", True, Nothing, Nothing, "width='150px' align='left'")
'                DG.AddBooleanColumn("PreferSpanishInd", "PreferSpanishInd", "Prefer<br>Spanish", "PreferSpanishInd", True, "1", "Yes", "No", Nothing, "align='left'")
'                DG.AddDataBoundColumn("NumEmployeeCalls", "NumEmployeeCalls", "Num<br />Empl<br />Calls", "NumEmployeeCalls", True, Nothing, Nothing, "align='center'")

'                DG.AddDataBoundColumn("NumAttempts", "NumAttempts", "Num<br />Tries", "NumAttempts", True, Nothing, Nothing, "align='center'")
'                DG.AddDateColumn("LastActivityDate", "LastActivityDate", "Last<br />Activity<br />Date", "dbo.ufn_GetLastActivityDate(cm.CallbackID)", True, "SpecialFormat3", Nothing, "width='110px' align='left'")
'                DG.AddDataBoundColumn("DaysRemaining", "DaysRemaining", "Days<br />Left", "DaysRemaining", True, Nothing, Nothing, "align='center'")

'                ' ___ Build the filter
'                Dim Filter As DG.Filter
'                Filter = DG.AttachFilter(DG.FilterOperationModeEnum.FilterAlwaysOn, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialHide)

'                Filter.AddDateCtlYMD("CreationDate", "CreationDate", "Get Date", "GetDateRange", "align='left' class='cellL'")
'                Filter.AddNameTextbox("EmpName", "EmpName", "EmpLastName", "EmpFirstName", 48, Nothing)
'                Filter.AddDropdown("ClientID", "ClientID")
'                Filter.AddDropdown("State", "State")
'                Filter.AddDropdown("CallPurposeCode", "ccp.CallPurposeCode")
'                Filter.AddExtendedDropdown("StatusCodeAdjDetail", "StatusCodeAdjDetail")
'                Filter.AddExtendedDropdown("PriorityTagInd", "PriorityTagInd")
'                Filter.AddExtendedDropdown("PreferSpanishInd", "PreferSpanishInd")

'                Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, Nothing, True, Nothing, True)
'                TemplateCol.AddDefaultIconItem("ViewCallback", "ExistingRequirement", "Magnify", "Setup Detail", Rights.CallbackView, Nothing)
'                TemplateCol.AddDefaultIconItem("NewAttempt", "NewAttempt", "NewIcon", "New Callback Attempt", Rights.CallbackEdit, "NewAttemptInd")
'                TemplateCol.AddDefaultIconItem("DeleteCallback", "DeleteCallback", "DeleteStrong", "Delete", Rights.CallbackEdit, Nothing)
'                DG.AddTemplateCol(TemplateCol)

'                Dim ChildTable As DG.ChildTableClass
'                ChildTable = DG.AttachChildTable("CallbackAttempt")
'                DG.AddChildTableSelectColumn("CallbackAttempt", Nothing, Nothing, Nothing, "CallbackID")

'            ElseIf Entity = "CallAttempt" Then
'                DG = New DG("CallbackAttemptID", cCommon, cRights, True, "EmbeddedTableDef", "AttemptDate", DG.DefaultSortDirectionEnum.Ascending)
'                DG.AddDateColumn("AttemptDate", "AttemptDate", "Date", Nothing, True, "SpecialFormat2", Nothing, "align='left'")
'                DG.AddDataBoundColumn("EnrollerName", "EnrollerName", "Enroller", Nothing, True, Nothing, Nothing, "align='left'")
'                DG.AddDataBoundColumn("Action", "Action", "Action", Nothing, True, Nothing, Nothing, "align='left'")

'                Dim TemplateCol As New DG.TemplateColumn("Icons", Nothing, Nothing, True, Nothing, True)
'                'TemplateCol.AddDefaultIconItem("ViewCallbackAttempt", "ExistingCallbackAttempt", "StandardView", "Setup Detail", Rights.CallbackView, Nothing, "CallbackID")
'                TemplateCol.AddDefaultIconItem("ViewCallbackAttempt", "ExistingCallbackAttempt", "Magnify", "Setup Detail", Rights.CallbackView, Nothing, "CallbackID")
'                'TemplateCol.AddDefaultIconItem("DeleteCallbackAttempt", "DeleteCallbackAttempt", "StandardDelete", "Delete", Rights.CallbackEdit, Nothing)
'                TemplateCol.AddDefaultIconItem("DeleteCallbackAttempt", "DeleteCallbackAttempt", "DeleteStrong", "Delete", Rights.CallbackEdit, Nothing, "CallbackID")
'                DG.AddTemplateCol(TemplateCol)

'                DG.FormatAsSubTable = True
'            End If

'            Return DG

'        Catch ex As Exception
'            Throw New Exception("Error #117: Index DefineDataGrid. " & ex.Message, ex)
'        End Try
'    End Function
'#End Region

'#Region " Display Page "
'    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
'        Dim ViewOrDownload As String = "View"
'        Dim sbAltLogin As New System.Text.StringBuilder

'        Try

'            ' ___ Heading/UserID
'            litHeading.Text = "Callback Worklist"

'            ' ___ Handle the filter
'            HandleFilter(DG, PageMode)

'            ' ___ Handle the sort
'            If cIndexSess.DGSess.SortReference <> Nothing Then
'                DG.UpdateSortReference(cIndexSess.DGSess.SortReference)
'            End If
'            DG.SetSortElements(OrderByField, OrderByType)

'            ' ___ Datagrid subtable
'            If cIndexSess.DGSess.SubTableInd = "1" Then
'                DG.GetChildTableSelectColumn.ParmColl("DataFldName").Value = cIndexSess.DGSess.ActiveSubTableRecID
'            End If

'            ' ___ Set the FilterOnOffState
'            cIndexSess.DGSess.FilterOnOffState = "on"

'            ' ___ Handle the data
'            HandleData(DG, PageMode, OrderByType, ViewOrDownload)

'            ' ___ Set the last field sorted and sort direction in the sort reference
'            cIndexSess.DGSess.SortReference = DG.GetSortReference

'            litFilterHiddens.Text = "<input type='hidden' name='hdDateRange' id='hdDateRange' value='" & cIndexSess.Filter.DateRange & "'>"

'            sbAltLogin.Append("<table style='background:#eeeedd;PADDING-LEFT: 12px;FONT:8pt Arial, Helvetica, sans-serif;POSITION:absolute;TOP:14px;left:170px' cellSpacing='0' cellPadding='0' width='250px' border='0'>")
'            sbAltLogin.Append("<colgroup><col width ='40px'/><col width='190px' /></colgroup>")
'            sbAltLogin.Append("<tr><td align ='center' colspan='2'>App Not cooperating?</td></tr>")
'            sbAltLogin.Append("<tr><td>UserID : </td>")
'            sbAltLogin.Append("<td><input type ='text' id='txtAltLogin' name='txtAltlogin' />")
'            sbAltLogin.Append("&nbsp;&nbsp;<input type='button' onclick='fnAltLogin()'/></td></tr></table>")
'            litAltLogin.Text = sbAltLogin.ToString

'        Catch ex As Exception
'            Throw New Exception("Error #110: Index DisplayPage. " & ex.Message, ex)
'        End Try
'    End Sub

'    'Private Sub HandleAltLogin()
'    '    Dim sb As New System.Text.StringBuilder

'    '    Try

'    '        sb.Append("<div class=""altlogin"">)



'    '    Catch ex As Exception
'    '        Throw New Exception("Error #116: Index.HandleAltLogin. " & ex.Message, ex)
'    '    End Try
'    'End Sub

'    Private Sub HandleFilter(ByRef DG As DG, ByVal PageMode As PageMode)
'        Dim i As Integer
'        Dim dt As DataTable

'        Try

'            ' ___ Get a filter reference
'            Dim Filter As DG.Filter
'            Filter = DG.GetFilter

'            ' '' ___ ClientID
'            ''If PageMode <> PageMode.Initial Then
'            ''    Filter.Coll("UserID").SetFilterValue(cCallSess.UserIDFilter)
'            ''    Filter.Coll("FullName").SetFilterValue(cCallSess.FullNameFilter)
'            ''End If

'            ' ___ Call date range
'            Filter.Coll("CreationDate").SetFilterValue(cIndexSess.Filter.DateRange)


'            ' ___ EnrollerID, FullName
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("EmpName").SetFilterValue(cIndexSess.Filter.EmpName)
'            End If


'            ' ___ ClientID
'            dt = Common.GetDT("SELECT DISTINCT ClientID from CallbackMaster WHERE LogicalDelete = 0 ORDER BY ClientID")
'            Filter("ClientID").AddDropdownItem("", "", True)
'            For i = 0 To dt.Rows.Count - 1
'                Filter("ClientID").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
'            Next
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("ClientID").SetFilterValue(cIndexSess.Filter.ClientID)
'            End If

'            ' ___ State
'            dt = Common.GetDT("SELECT StateCode, StateName FROM UserManagement..Codes_State WHERE StateCode <> 'XX' ORDER BY StateName")
'            Filter("State").AddDropdownItem("", "", True)
'            For i = 0 To dt.Rows.Count - 1
'                Filter("State").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(0), False)
'            Next
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("State").SetFilterValue(cIndexSess.Filter.State)
'            End If

'            ' ___ Call purpose
'            dt = Common.GetDT("SELECT CallPurposeCode, CallPurposeDescription FROM Codes_CallPurpose ORDER BY Seq")
'            Filter("CallPurposeCode").AddDropdownItem("", "", True)
'            For i = 0 To dt.Rows.Count - 1
'                Filter("CallPurposeCode").AddDropdownItem(dt.Rows(i)(0), dt.Rows(i)(1), False)
'            Next
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("CallPurposeCode").SetFilterValue(cIndexSess.Filter.CallPurposeCode)
'            End If

'            Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("", "", False)
'            'Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("0", "Initialize", "dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate())='INIT'", False)
'            Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("1", "Callback", "dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate())='CB'", True)
'            Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("2", "Ticket closed", "dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate())='CLOS'", False)
'            Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("3", "Win exp", "dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate())='WE'", False)
'            Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("4", "Trouble call", "dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate())='TC'", False)
'            'Filter("StatusCodeAdjDetail").AddExtendedDropdownItem("5", "Wrong num ver", "dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate())='WNV'", False)
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("StatusCodeAdjDetail").SetFilterValue(cIndexSess.Filter.StatusCodeAdjDetail)
'            End If


'            ' ___ PriorityTagInd
'            Filter("PriorityTagInd").AddExtendedDropdownItem("", "", "", True)
'            Filter("PriorityTagInd").AddExtendedDropdownItem("1", "Yes", " PriorityTagInd = 1 ", False)
'            Filter("PriorityTagInd").AddExtendedDropdownItem("0", "No", " (PriorityTagInd = 0 OR PriorityTagInd IS NULL)", False)
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("PriorityTagInd").SetFilterValue(cIndexSess.Filter.PriorityTagInd)
'            End If

'            ' ___ PreferSpanishInd
'            Filter("PreferSpanishInd").AddExtendedDropdownItem("", "", "", True)
'            Filter("PreferSpanishInd").AddExtendedDropdownItem("1", "Yes", " PreferSpanishInd = 1 ", False)
'            Filter("PreferSpanishInd").AddExtendedDropdownItem("0", "No", " (PreferSpanishInd = 0 OR PreferSpanishInd IS NULL)", False)
'            If PageMode <> PageMode.Initial Then
'                Filter.Coll("PreferSpanishInd").SetFilterValue(cIndexSess.Filter.PreferSpanishInd)
'            End If

'        Catch ex As Exception
'            Throw New Exception("Error #111: Index HandleFilter. " & ex.Message, ex)
'        End Try
'    End Sub

'    Private Function GetFilterColl() As CollX
'        Dim FilterColl As New CollX
'        Dim FilterValue As String
'        Dim Box() As String

'        Try

'            FilterColl.Assign("@EmpLastName", DBNull.Value)
'            FilterColl.Assign("@EmpFirstName", DBNull.Value)
'            FilterColl.Assign("@CreationDateFrom", DBNull.Value)
'            FilterColl.Assign("@CreationDateTo", DBNull.Value)
'            FilterColl.Assign("@ClientID", DBNull.Value)
'            FilterColl.Assign("@StatusCodeAdj", DBNull.Value)
'            FilterColl.Assign("@CallPurposeCode", DBNull.Value)
'            FilterColl.Assign("@State", DBNull.Value)
'            FilterColl.Assign("@PriorityTagInd", DBNull.Value)
'            FilterColl.Assign("@PreferSpanishInd", DBNull.Value)

'            FilterValue = cIndexSess.Filter.EmpName
'            If Common.IsNotBlank(FilterValue) Then
'                Box = FilterValue.Split(",")
'                If Box.GetUpperBound(0) = 0 Then
'                    FilterColl.Assign("@EmpLastName", Box(0))
'                Else
'                    FilterColl.Assign("@EmpLastName", Box(0))
'                    FilterColl.Assign("@EmpFirstName", Box(1))
'                End If
'            End If

'            FilterValue = cIndexSess.Filter.DateRange
'            Box = FilterValue.Split("|")
'            If IsNumeric(Box(0)) AndAlso IsNumeric(Box(1)) AndAlso IsNumeric(Box(2)) Then
'                FilterColl.Assign("@CreationDateFrom", Box(1) & "/" & Box(2) & "/" & Box(0))
'            End If
'            If IsNumeric(Box(3)) AndAlso IsNumeric(Box(4)) AndAlso IsNumeric(Box(5)) Then
'                FilterColl.Assign("@CreationDateTo", Box(4) & "/" & Box(5) & "/" & Box(3))
'            End If

'            FilterValue = cIndexSess.Filter.ClientID
'            If Common.IsNotBlank(FilterValue) Then
'                FilterColl.Assign("@ClientID", FilterValue)
'            End If

'            FilterValue = cIndexSess.Filter.State
'            If Common.IsNotBlank(FilterValue) Then
'                FilterColl.Assign("@State", FilterValue)
'            End If

'            Select Case cIndexSess.Filter.StatusCodeAdjDetail
'                Case "1"
'                    FilterValue = "CB"
'                Case "2"
'                    FilterValue = "CLOS"
'                Case "3"
'                    FilterValue = "WE"
'                Case "4"
'                    FilterValue = "TC"
'            End Select
'            If Common.IsNotBlank(FilterValue) Then
'                FilterColl.Assign("@StatusCodeAdjDetail", FilterValue)
'            End If

'            FilterValue = cIndexSess.Filter.PriorityTagInd
'            If Common.IsNotBlank(FilterValue) Then
'                FilterColl.Assign("@PriorityTagInd", FilterValue)
'            End If

'            FilterValue = cIndexSess.Filter.PreferSpanishInd
'            If Common.IsNotBlank(FilterValue) Then
'                FilterColl.Assign("@PreferSpanishInd", FilterValue)
'            End If

'            Return FilterColl

'        Catch ex As Exception
'            Throw New Exception("Error #119: Index GetFilterColl. " & ex.Message, ex)
'        End Try
'    End Function

'    Private Sub HandleData(ByRef DG As DG, ByVal PageMode As PageMode, ByVal OrderByType As DG.OrderByType, ByVal ViewOrDownload As String)
'        Dim dt As DataTable = Nothing
'        Dim RecordCount As Integer
'        Dim Coll As Collection
'        Dim SuppressDisplayData As Boolean
'        Dim EmbeddedMessage As String = Nothing
'        Dim sb As New System.Text.StringBuilder
'        Dim DownloadPathColl As Collection = Nothing
'        Dim HTMLReport As Boolean
'        Dim DownloadReport As Boolean
'        Dim ExceedsRecordMaximum As Boolean
'        Dim IgnoreExcessiveRecords As Boolean
'        Dim NewQuery As Boolean
'        Dim PerformNextTest As Boolean

'        Try

'            ' ___ #1: FIGURE OUT WHAT'S GOING ON

'            ' ___ Get the recordcount
'            Coll = GetQueryInfo("RecordCount", DG, OrderByType)
'            RecordCount = Coll("RecordCount")


'            ' ___ Test #1: InitialReportDataSuppress
'            ' User opens this page. The datagrid suppresses the initial display of the data. 
'            ' If the user navigates to a different page and then returns to this page, PageMode is no longer set
'            ' to initial and the initial data is permitted to display. The Sql attempts to return all of the records the sql statement
'            ' with no restrictions. A postback is required to enable the display of the data.

'            If PageMode = PageMode.Initial AndAlso DG.RecordsInitialShowHide = DG.RecordsInitialShowHideEnum.RecordsInitialHide Then
'                cIndexSess.DGSess.InitialReportDataSuppressInEffect = True
'            ElseIf PageMode = PageMode.Postback AndAlso cIndexSess.DGSess.InitialReportDataSuppressInEffect Then
'                cIndexSess.DGSess.InitialReportDataSuppressInEffect = False
'            End If
'            If Not cIndexSess.DGSess.InitialReportDataSuppressInEffect Then
'                PerformNextTest = True
'            End If


'            ' ___ Test #2: Exceeds record maximum test
'            If PerformNextTest Then
'                If RecordCount > Enviro.RecordMaximum Then
'                    ExceedsRecordMaximum = True
'                    PerformNextTest = False
'                End If
'            End If

'            ' ___ Test #3: Ignore excessive records warning
'            ' If an excessive records warning was put into effect and the user has chosen to proceed with the same query, ignore and reset the warning.
'            If PerformNextTest Then
'                If cIndexSess.DGSess.ExcessiveRecordsWarningInEffect Then
'                    Coll = GetQueryInfo("Sql", DG, OrderByType)
'                    If StrComp(Coll("Sql"), cIndexSess.DGSess.Sql, CompareMethod.Text) = 0 Then
'                        IgnoreExcessiveRecords = True
'                        PerformNextTest = False
'                    Else
'                        NewQuery = True
'                    End If
'                Else
'                    NewQuery = True
'                End If

'                ' ___ Reset the excessive record warning properties
'                If cIndexSess.DGSess.ExcessiveRecordsWarningInEffect Then
'                    cIndexSess.DGSess.ExcessiveRecordsWarningInEffect = False
'                    cIndexSess.DGSess.Sql = String.Empty
'                End If
'            End If

'            ' ___ Test #4: Excessive records test
'            If PerformNextTest Then
'                If RecordCount > Enviro.ExcessiveRecordAmount Then
'                    Coll = GetQueryInfo("Sql", DG, OrderByType)
'                    cIndexSess.DGSess.ExcessiveRecordsWarningInEffect = True
'                    cIndexSess.DGSess.Sql = Coll("Sql")
'                End If
'            End If

'            ' ___ DIRECT THE REPORT
'            If ViewOrDownload = "Download" Then
'                DownloadReport = True
'            Else
'                HTMLReport = True
'            End If
'            If cIndexSess.DGSess.InitialReportDataSuppressInEffect Then
'                HTMLReport = True
'                DownloadReport = False
'                SuppressDisplayData = True
'            ElseIf ExceedsRecordMaximum Then
'                HTMLReport = True
'                DownloadReport = False
'                SuppressDisplayData = True
'                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & RecordCount & " records. Respecify report.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
'            ElseIf IgnoreExcessiveRecords Then
'                HTMLReport = False
'                DownloadReport = True
'            ElseIf cIndexSess.DGSess.ExcessiveRecordsWarningInEffect Then
'                HTMLReport = True
'                DownloadReport = False
'                SuppressDisplayData = True
'                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & RecordCount & " records. Proceed or respecify report.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
'            End If
'            If DownloadReport Then
'                DownloadPathColl = cCommon.GetDownloadPath(Page)
'                SuppressDisplayData = True
'                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif"">&nbsp;&nbsp;Click here to <a href=""" & DownloadPathColl("RelPath") & """>download</a> your CSV file.</td>"
'            End If


'            ' ___ EXECUTE THE REPORT

'            ' ___ Get the data
'            If (HTMLReport And (Not SuppressDisplayData)) OrElse DownloadReport Then
'                Coll = GetQueryInfo("Data", DG, OrderByType)
'                dt = Coll("Data")
'            End If

'            ' ___ Process the download
'            If DownloadReport Then
'                cCommon.PrintCSVVersionLocal(dt, DownloadPathColl("AbsPath"), Nothing)
'            End If

'            ' ___ Process the html
'            If SuppressDisplayData Then
'                dt = Nothing
'            End If
'            litDG.Text = DG.GetHTML(dt, Request, EmbeddedMessage)

'        Catch ex As Exception
'            Throw New Exception("Error #112: Index DisplayPage. " & ex.Message, ex)
'        End Try
'    End Sub

'    'Private Function GetQueryInfo(ByVal InfoType As String, ByVal DG As DG, ByVal OrderByType As DG.OrderByType) As Collection
'    '    Dim i As Integer
'    '    Dim sb As New System.Text.StringBuilder
'    '    Dim Sql As String
'    '    Dim dt As DataTable = Nothing
'    '    Dim ShowFilter As Boolean
'    '    Dim Coll As New Collection
'    '    Dim Pos As Integer
'    '    Dim QueryPack As DBase.QueryPack
'    '    Dim Recordcount As Integer
'    '    Dim PremiumGrandTotal As Decimal
'    '    Dim NonFilterWhereClause As String = Nothing
'    '    Dim QueryColl As New CollX
'    '    Dim FilterColl As CollX


'    '    Try

'    '        FilterColl = GetFilterColl()

'    '        ' ___ This method is requested to return one of three items:
'    '        ' (1) a recordcount, (2) a datatable, or (3) a sql statement.
'    '        If InfoType = "RecordCount" Then
'    '            sb.Append("SELECT Count (*) ")
'    '        ElseIf InfoType = "Data" Or InfoType = "Sql" Then
'    '            sb.Append("SELECT cm.CallbackID, cm.CreationDate, TicketNumber = " & Common.GetTicketNumberSyntax("cm") & ", ")
'    '            sb.Append("EmpName = EmpLastName + ', ' + EmpFirstName + ' ' + EmpMI , ")
'    '            sb.Append("EmpLastName, EmpFirstName, ")
'    '            sb.Append("ClientID = cm.ClientID, State = cm.State, ")
'    '            sb.Append("CallPurposeCode = cm.CallPurposeCode, CallPurposeDescription = ccp.CallPurposeDescription, ")
'    '            sb.Append("PriorityTagInd, cp.BestTime, cm.PreferSpanishInd, ")
'    '            sb.Append("NumAttempts = (SELECT Count (*) FROM CallbackAttempt ca WHERE ca.CallbackID = cm.CallbackID AND ca.LogicalDelete = 0), ")
'    '            sb.Append("LastActivityDate =  dbo.ufn_GetLastActivityDate(cm.CallbackID), ")
'    '            sb.Append("cm.NumEmployeeCalls, ")
'    '            sb.Append("DaysRemaining = case ")
'    '            'sb.Append("when cm.EnrollWinEndDate IS NULL then 'No win' ")
'    '            sb.Append("when dbo.ufn_GetEnrollWinEndDate(cm.EmpID, cm.ClientID) IS NULL then 'No win' ")

'    '            'sb.Append("else Cast((SELECT DateDiff(d, getDate(), cm.EnrollWinEndDate)) as varchar(10)) ")
'    '            sb.Append("else Cast((SELECT DateDiff(d, getDate(), dbo.ufn_GetEnrollWinEndDate(cm.EmpID, cm.ClientID))) as varchar(10)) ")
'    '            sb.Append("end, ")
'    '            sb.Append("NewAttemptInd = case ")
'    '            sb.Append("when dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate()) <> 'TC' then '1' ")
'    '            sb.Append("else '0' ")
'    '            sb.Append("end, ")
'    '            sb.Append("StatusCodeAdjDetail = dbo.ufn_GetStatusCodeAdj(1, cm.CallbackID, getDate()) ")
'    '        End If

'    '        sb.Append("FROM CallbackMaster cm ")

'    '        '11/04/2013

'    '        'sb.Append("INNER JOIN Codes_CallPurpose ccp ON cm.CallPurposeCode = ccp.CallPurposeCode ")
'    '        sb.Append("INNER JOIN Codes_CallPurpose ccp ON cm.CallPurposeCode = ccp.CallPurposeCode AND cm.LogicalDelete = 0 And ISNULL(cm.IsPurged, 0) = 0")

'    '        sb.Append("LEFT JOIN CallbackPhone cp ON cm.CallbackID = cp.CallbackID AND cp.Seq = 1 ")
'    '        Sql = sb.ToString

'    '        'NonFilterWhereClause = "cm.LogicalDelete = 0"
'    '        'NonFilterWhereClause = "cm.LogicalDelete = 0 AND ISNULL(cm.IsPurged, 0) = 0"

'    '        DG.GenerateSQL(Sql, ShowFilter, NonFilterWhereClause, OrderByType, Request, cIndexSess.DGSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"), QueryColl)
'    '        'DG.GenerateSQLUnion(Sql1, Sql2, ShowFilter, NonFilterWhereClause1, NonFilterWhereClause2, OrderByType, Request, cCallSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))

'    '        ' ___ Eliminate order by clause from recordcount query
'    '        If InfoType = "RecordCount" Then
'    '            Pos = InStr(Sql, "ORDER BY", CompareMethod.Binary)
'    '            If Pos > 0 Then
'    '                Sql = Sql.Substring(0, Pos - 1)
'    '            End If
'    '        End If

'    '        If InfoType <> "Sql" Then
'    '            QueryPack = Common.GetDTExtendedWithQueryPack(Sql)
'    '            If QueryPack.Success Then
'    '                dt = QueryPack.dt
'    '            Else
'    '                Throw New Exception("Error #118a: CallWorklist GetQueryInfo Info Type: " & InfoType & QueryPack.TechErrMsg)
'    '            End If
'    '        End If

'    '        If InfoType = "RecordCount" Then
'    '            Recordcount = dt.Rows(0)(0).Value
'    '            Coll.Add(Recordcount, "RecordCount")
'    '        ElseIf InfoType = "Data" Then
'    '            Coll.Add(dt, "Data")
'    '            If DG.GetTotal IsNot Nothing Then
'    '                For i = 0 To dt.Rows.Count - 1
'    '                    If IsNumeric(dt.Rows(i)("TotalPremium").Value) Then
'    '                        PremiumGrandTotal += dt.Rows(i)("TotalPremium").Value
'    '                    End If
'    '                Next
'    '                DG.GetTotal.Coll("TotalPremium").Value = PremiumGrandTotal
'    '            End If


'    '        ElseIf InfoType = "Sql" Then
'    '            Coll.Add(Sql, "Sql")
'    '        End If
'    '        Return Coll

'    '    Catch ex As Exception
'    '        Throw New Exception("Error #118b: Index GetQueryInfo Info Type: " & InfoType & ". " & ex.Message, ex)
'    '    End Try
'    'End Function

'    Private Function GetQueryInfo(ByVal InfoType As String, ByRef DG As DG, ByVal OrderByType As DG.OrderByType) As Collection
'        Dim i As Integer
'        'Dim sb As New System.Text.StringBuilder
'        Dim Sql As String
'        Dim dt As DataTable = Nothing
'        Dim ShowFilter As Boolean
'        Dim Coll As New Collection
'        'Dim Pos As Integer
'        Dim QueryPack As DBase.QueryPack
'        Dim Recordcount As Integer
'        Dim PremiumGrandTotal As Decimal
'        Dim NonFilterWhereClause As String = Nothing
'        Dim QueryColl As New CollX
'        Dim FilterColl As CollX
'        Dim ParamStr As New System.Text.StringBuilder
'        Dim OrderByStr As String = Nothing

'        Try

'            FilterColl = GetFilterColl()
'            For i = 1 To FilterColl.Count
'                If Not IsDBNull(FilterColl(i)) Then
'                    If i < FilterColl.Count Then
'                        ParamStr.Append(FilterColl.Key(i), FilterColl(i) & ", ")
'                    Else
'                        ParamStr.Append(FilterColl.Key(i), FilterColl(i))
'                    End If
'                End If
'            Next

'            '' ___ This method is requested to return one of three items:
'            '' (1) a recordcount, (2) a datatable, or (3) a sql statement.
'            'If InfoType = "RecordCount" Then
'            '    sb.Append("SELECT Count (*) ")
'            'ElseIf InfoType = "Data" Or InfoType = "Sql" Then
'            '    sb.Append("SELECT cm.CallbackID, cm.CreationDate, TicketNumber = " & Common.GetTicketNumberSyntax("cm") & ", ")
'            '    sb.Append("EmpName = EmpLastName + ', ' + EmpFirstName + ' ' + EmpMI , ")
'            '    sb.Append("EmpLastName, EmpFirstName, ")
'            '    sb.Append("ClientID = cm.ClientID, State = cm.State, ")
'            '    sb.Append("CallPurposeCode = cm.CallPurposeCode, CallPurposeDescription = ccp.CallPurposeDescription, ")
'            '    sb.Append("PriorityTagInd, cp.BestTime, cm.PreferSpanishInd, ")
'            '    sb.Append("NumAttempts = (SELECT Count (*) FROM CallbackAttempt ca WHERE ca.CallbackID = cm.CallbackID AND ca.LogicalDelete = 0), ")
'            '    sb.Append("LastActivityDate =  dbo.ufn_GetLastActivityDate(cm.CallbackID), ")
'            '    sb.Append("cm.NumEmployeeCalls, ")
'            '    sb.Append("DaysRemaining = case ")
'            '    sb.Append("when dbo.ufn_GetEnrollWinEndDate(cm.EmpID, cm.ClientID) IS NULL then 'No win' ")
'            '    sb.Append("else Cast((SELECT DateDiff(d, getDate(), dbo.ufn_GetEnrollWinEndDate(cm.EmpID, cm.ClientID))) as varchar(10)) ")
'            '    sb.Append("end, ")
'            '    sb.Append("NewAttemptInd = case ")
'            '    sb.Append("when dbo.ufn_GetStatusCodeAdj(0, cm.CallbackID, getDate()) <> 'TC' then '1' ")
'            '    sb.Append("else '0' ")
'            '    sb.Append("end, ")
'            '    sb.Append("StatusCodeAdjDetail = dbo.ufn_GetStatusCodeAdj(1, cm.CallbackID, getDate()) ")
'            'End If

'            'sb.Append("FROM CallbackMaster cm ")
'            'sb.Append("INNER JOIN Codes_CallPurpose ccp ON cm.CallPurposeCode = ccp.CallPurposeCode AND cm.LogicalDelete = 0 And ISNULL(cm.IsPurged, 0) = 0")
'            'sb.Append("LEFT JOIN CallbackPhone cp ON cm.CallbackID = cp.CallbackID AND cp.Seq = 1 ")
'            'Sql = sb.ToString

'            DG.GenerateSQL("Deprecated", ShowFilter, OrderByStr, NonFilterWhereClause, OrderByType, Request, cIndexSess.DGSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"), QueryColl)
'            'DG.GenerateSQLUnion(Sql1, Sql2, ShowFilter, NonFilterWhereClause1, NonFilterWhereClause2, OrderByType, Request, cCallSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))

'            ' ___ Eliminate order by clause from recordcount query
'            'If InfoType = "RecordCount" Then
'            '    Pos = InStr(Sql, "ORDER BY", CompareMethod.Binary)
'            '    If Pos > 0 Then
'            '        Sql = Sql.Substring(0, Pos - 1)
'            '    End If
'            'End If

'            If InfoType <> "Sql" Then
'                QueryPack = Common.GetDTExtendedWithQueryPack(Sql)
'                If QueryPack.Success Then
'                    dt = QueryPack.dt
'                Else
'                    Throw New Exception("Error #118a: CallWorklist GetQueryInfo Info Type: " & InfoType & QueryPack.TechErrMsg)
'                End If
'            End If

'            If InfoType = "RecordCount" Then
'                Recordcount = dt.Rows(0)(0).Value
'                Coll.Add(Recordcount, "RecordCount")
'            ElseIf InfoType = "Data" Then
'                Coll.Add(dt, "Data")
'                If DG.GetTotal IsNot Nothing Then
'                    For i = 0 To dt.Rows.Count - 1
'                        If IsNumeric(dt.Rows(i)("TotalPremium").Value) Then
'                            PremiumGrandTotal += dt.Rows(i)("TotalPremium").Value
'                        End If
'                    Next
'                    DG.GetTotal.Coll("TotalPremium").Value = PremiumGrandTotal
'                End If
'            ElseIf InfoType = "Sql" Then
'                Coll.Add(Sql, "Sql")
'            End If





'            Select Case "InfoType"
'                Case "Recordcount"
'                    QueryPack = Common.GetDTWithQueryPack("exec usp_TempTemp @InfoType='RecCount' " & ParamStr.ToString)
'                    Recordcount = QueryPack.dt.Rows(0)(0)
'                    Coll.Add(Recordcount, "RecordCount")
'                Case "Data"
'                    QueryPack = Common.GetDTWithQueryPack("exec usp_TempTemp @InfoType='Data' " & ParamStr.ToString)
'                    dt = QueryPack.dt
'                    Coll.Add(dt, "Data")
'                    If DG.GetTotal IsNot Nothing Then
'                        For i = 0 To dt.Rows.Count - 1
'                            If IsNumeric(dt.Rows(i)("TotalPremium")) Then
'                                PremiumGrandTotal += dt.Rows(i)("TotalPremium")
'                            End If
'                        Next
'                        DG.GetTotal.Coll("TotalPremium").value = PremiumGrandTotal
'                    End If
'                Case "Sql"
'                    QueryPack = Common.GetDTWithQueryPack("exec usp_TempTemp @InfoType='Sql' " & ParamStr.ToString)
'            End Select

'            Return Coll

'        Catch ex As Exception
'            Throw New Exception("Error #118b: Index GetQueryInfo Info Type: " & InfoType & ". " & ex.Message, ex)
'        End Try
'    End Function


'#End Region

'#Region " Event raised by the data grid "
'    Public Sub HandleChildDTRequest(ByVal ParmColl As Collection, ByRef ChildText As String)
'        Dim Sql As New System.Text.StringBuilder
'        Dim dt As DataTable

'        Try
'            Sql.Append("SELECT ca.CallbackAttemptID, ca.CallbackID, ca.AttemptDate, ")
'            Sql.Append("EnrollerName = u.LastName + ', ' + u.FirstName, ")
'            Sql.Append("Action = case ")
'            Sql.Append("when  ca.CallbackAttemptStatusCode  = 'LMWP'  then 'Left message with ' + ca.LeftMessageWith ")
'            Sql.Append("else stat.AttemptDescription ")
'            Sql.Append("end ")
'            Sql.Append("FROM CallbackAttempt ca ")
'            Sql.Append("INNER JOIN UserManagement..Users u ON ca.EnrollerID = u.UserID ")
'            Sql.Append("INNER JOIN Codes_CallbackAttemptStatus stat ON ca.CallbackAttemptStatusCode = stat.CallbackAttemptStatusCode ")
'            Sql.Append("WHERE ca.CallbackID = " & ParmColl("DataFldName").Value & " AND ")
'            Sql.Append("ca.LogicalDelete = 0 ")
'            Sql.Append("ORDER BY ca.AttemptDate")
'            dt = Common.GetDTExtended(Sql.ToString)
'            ChildText = cDGCallAttempt.GetHTML(dt, Request, Nothing)

'        Catch ex As Exception
'            Throw New Exception("Error #115b: Index HandleChildDTRequest. " & ex.Message, ex)
'        End Try

'    End Sub
'#End Region

'    <System.Web.Services.WebMethod()>
'    Public Shared Function CheckoutRecordIfAvailable(ByVal LoggedInUserID As String, ByVal CallbackID As Int64, ByVal CallbackAttemptID As Int64, ByVal PageCode As String, ByVal MaintDelete As String, ByVal NewExisting As String, ByVal MaintOpen As String, ByVal AttemptOpen As String) As String
'        Dim mEnviro As Enviro
'        Dim Sql As New System.Text.StringBuilder
'        Dim Querypack As DBase.QueryPack
'        Dim dr As DataRow = Nothing
'        Dim Results As New System.Text.StringBuilder
'        Dim IsAlreadyCheckedOut As Char = "0"
'        Dim CheckoutUser As String = String.Empty
'        Dim CheckoutUserID As String = String.Empty
'        Dim ErrorInd As String = String.Empty
'        Dim ErrorMessage As String = String.Empty
'        Dim NewCheckoutSuccessfulInd As String = "0"
'        Dim TimeSpan As TimeSpan
'        Dim OKToCheckout As Boolean

'        Try

'            mEnviro = Common.GetSessionObj("Enviro")

'            ' this is good
'            'Sql.Append("SELECT NumSecondsElapsed = DateDiff(ss ,  DateAdd(hh, -5, getutcdate()) , CheckoutTime, ")

'            Sql.Append("SELECT cm.CheckoutTime, ")
'            Sql.Append("cm.CheckoutUserID, ")
'            Sql.Append("CurrentCheckoutUser = u.FirstName + ' ' + u.LastName ")
'            Sql.Append("FROM CallbackMaster cm ")
'            Sql.Append("INNER JOIN UserManagement..Users u ON cm.CheckoutUserID = u.UserID And cm.CallbackID = " & CallbackID)
'            Querypack = Common.GetDTWithQueryPack(Sql.ToString)

'            ' ___ Error condition
'            If Querypack.Success Then
'                ErrorInd = "0"
'                If Querypack.dt.Rows.Count > 0 Then
'                    dr = Querypack.dt.Rows(0)
'                End If
'                OKToCheckout = True
'            Else
'                ErrorInd = "1"
'                ErrorMessage = Common.ToJSAlert(Querypack.TechErrMsg)
'                OKToCheckout = False
'            End If

'            If OKToCheckout AndAlso Querypack.dt.Rows.Count > 0 AndAlso dr("CheckoutUserID") <> mEnviro.LoggedInUserID Then
'                TimeSpan = Common.GetServerDateTime().Subtract(dr("CheckoutTime"))
'                If TimeSpan.TotalSeconds < 25 Then
'                    IsAlreadyCheckedOut = "1"
'                    CheckoutUserID = dr("CheckoutUserID")
'                    CheckoutUser = dr("CurrentCheckoutUser")
'                    OKToCheckout = False
'                End If
'            End If




'            'If OKToCheckout Then
'            '    ' ___ Never been checked out
'            '    If Querypack.dt.Rows.Count = 0 Then
'            '        IsAlreadyCheckedOutInd = "0"
'            '    Else

'            '        dr = Querypack.dt.Rows(0)

'            '        ' ___ Last checkout write less than 25 seconds ago?
'            '        TimeSpan = Common.GetServerDateTime().Subtract(dr("CheckoutTime"))
'            '        If TimeSpan.TotalSeconds < 25 Then
'            '            If dr("CheckoutUserID") = mEnviro.LoggedInUserID Then
'            '                IsAlreadyCheckedOutInd = "0"
'            '            Else
'            '                IsAlreadyCheckedOutInd = "1"
'            '                CheckoutUserID = dr("CheckoutUserID")
'            '                CheckoutUser = dr("CurrentCheckoutUser")
'            '                OKToProceedInd = False
'            '            End If
'            '        Else
'            '            IsAlreadyCheckedOutInd = "0"
'            '        End If
'            '    End If
'            'End If


'            If OKToCheckout Then
'                Sql.Length = 0
'                Sql.Append("UPDATE CallbackMaster SET ")
'                Sql.Append("CheckoutUserID = '" & LoggedInUserID & "',  ")
'                Sql.Append("CheckoutTime = '" & Common.GetServerDateTime & "' ")
'                Sql.Append("WHERE CallbackID = " & CallbackID)
'                Querypack = Common.ExecuteNonQueryWithQuerypack(Sql.ToString)

'                If Querypack.Success Then
'                    NewCheckoutSuccessfulInd = "1"
'                Else
'                    ErrorInd = "1"
'                    ErrorMessage = Common.ToJSAlert(Querypack.TechErrMsg)
'                End If
'            End If

'            ' 0: CallbackID
'            ' 1: IsAlreadyCheckedOut
'            ' 2: CheckoutUser
'            ' 3: ErrorInd
'            ' 4: ErrorMessage
'            ' 5: PageCode
'            ' 6: NewCheckoutSuccessfulInd
'            ' 7: CallbackAttemptID
'            ' 8: MaintDelete
'            ' 9: CheckoutUserID
'            '10: NewExisting
'            '11: MaintOpen
'            '12: AttemptOpen

'            Results.Append(CallbackID.ToString & "|")
'            Results.Append(IsAlreadyCheckedOut & "|")
'            Results.Append(CheckoutUser & "|")
'            Results.Append(ErrorInd & "|")
'            Results.Append(ErrorMessage & "|")
'            Results.Append(PageCode & "|")
'            Results.Append(NewCheckoutSuccessfulInd & "|")
'            Results.Append(CallbackAttemptID & "|")
'            Results.Append(MaintDelete & "|")
'            Results.Append(CheckoutUserID & "|")
'            Results.Append(NewExisting & "|")
'            Results.Append(MaintOpen & "|")
'            Results.Append(AttemptOpen)

'            Return Results.ToString

'        Catch ex As Exception
'            Results.Length = 0
'            Results.Append(CallbackID.ToString & "|||1|" & Common.ToJSAlert(ex.Message) & "|" & PageCode & "|0|")
'            Return Results.ToString
'        End Try
'    End Function

'    <System.Web.Services.WebMethod()>
'    Public Shared Function CheckInRecord(ByVal CallbackID As Int64) As String
'        Dim Sql As New System.Text.StringBuilder
'        Dim Querypack As DBase.QueryPack
'        Dim Results As New System.Text.StringBuilder
'        Dim ErrorInd As String
'        Dim ErrorMessage As String = Nothing

'        ' 0: ErrorInd
'        ' 1: ErrorMessage

'        Sql.Append("UPDATE CallbackMaster SET ")
'        Sql.Append("CheckoutUserID = null,  ")
'        Sql.Append("CheckoutTime = null ")
'        Sql.Append("WHERE CallbackID = " & CallbackID)
'        Querypack = Common.ExecuteNonQueryWithQuerypack(Sql.ToString)

'        If Querypack.Success Then
'            ErrorInd = "0"
'        Else
'            ErrorInd = "1"
'            ErrorMessage = Common.ToJSAlert(Querypack.TechErrMsg)
'        End If

'        Results.Append(ErrorInd & "|")
'        Results.Append(ErrorMessage)

'        Return Results.ToString
'    End Function

'    <System.Web.Services.WebMethod()>
'    Public Shared Function IsRecordCheckedOut(ByVal CallbackID As Int64, ByVal CallbackAttemptID As Int64) As String
'        Dim Sql As String
'        Dim Querypack As DBase.QueryPack
'        Dim Results As New System.Text.StringBuilder
'        Dim ErrorInd As String
'        Dim ErrorMessage As String = Nothing
'        Dim IsCurrentlyCheckedOutInd As String
'        Dim ts As TimeSpan

'        ' 0: CallbackID
'        ' 1: CallbackAttemptID
'        ' 2: IsCurrentlyCheckedOutInd
'        ' 3: ErrorInd
'        ' 4: ErrorMessage

'        Sql = "SELECT CheckoutUserID, CheckoutTime FROM CallbackMaster WHERE CallbackID = '" & CallbackID & "'"
'        Querypack = Common.GetDTWithQueryPack(Sql)

'        If Querypack.Success Then
'            ErrorInd = "0"
'            If IsDBNull(Querypack.dt.Rows(0)(0)) Then
'                IsCurrentlyCheckedOutInd = "0"
'            Else
'                ts = Common.GetServerDateTime.Subtract(Querypack.dt.Rows(0)(1))
'                If ts.TotalSeconds > 29 Then
'                    IsCurrentlyCheckedOutInd = "0"
'                Else
'                    IsCurrentlyCheckedOutInd = "1"
'                End If
'            End If
'        Else
'            IsCurrentlyCheckedOutInd = "0"
'            ErrorInd = "1"
'            ErrorMessage = Common.ToJSAlert(Querypack.TechErrMsg)
'        End If

'        Results.Append(CallbackID.ToString & "|")
'        Results.Append(CallbackAttemptID.ToString & "|")
'        Results.Append(IsCurrentlyCheckedOutInd & "|")
'        Results.Append(ErrorInd & "|")
'        Results.Append(ErrorMessage)

'        Return Results.ToString
'    End Function

'    <System.Web.Services.WebMethod()>
'    Public Shared Function AuthenticateEnroller(ByVal UserID As String) As String
'        Dim mEnviro As Enviro
'        Dim Sql As String
'        Dim Querypack As DBase.QueryPack
'        Dim Results As New System.Text.StringBuilder
'        Dim IsAuthenticated As String = "0"
'        Dim ConnectionError As String = "0"
'        Dim InsufficentRights As String = "0"
'        Dim NotFound As String = "0"
'        Dim Message As String = String.Empty

'        ' 0: IsAuthenticated
'        ' 1: NotFound
'        ' 2: ConnectionError
'        ' 3: InsufficentRights
'        ' 4: Message

'        mEnviro = Common.GetSessionObj("Enviro")

'        Sql = "SELECT CompanyID, Role, LocationID FROM UserManagement..Users WHERE UserID ='" & UserID & "'"
'        Querypack = Common.GetDTWithQueryPack(Sql)

'        If Querypack.dt Is Nothing OrElse Not Querypack.Success Then
'            ConnectionError = "1"
'            Message = "A connection error has occurred. Please try again."
'        End If

'        If Querypack.dt.Rows.Count = 0 Then
'            NotFound = "1"
'            Message = "Unable to find " & UserID & " in UserManagement database."
'        End If

'        If ConnectionError = "0" AndAlso NotFound = "0" Then

'            If Querypack.dt.Rows(0)("CompanyID") = "BVI" Then
'                If Common.InList(Querypack.dt.Rows(0)("Role"), "ENROLLER, SUPERVISOR, IT, ADMIN, ADMIN LIC", True) Then
'                    'mRights.RightsColl.Add(Rights.CallbackView)
'                    'mRights.RightsColl.Add(Rights.CallbackEdit)
'                    IsAuthenticated = "1"
'                    Message = "You are authenticated!"
'                Else
'                    InsufficentRights = "1"
'                    Message = "You do not have sufficient rights."
'                End If
'            End If

'        End If

'        Results.Append(IsAuthenticated & "|")
'        Results.Append(NotFound & "|")
'        Results.Append(ConnectionError & "|")
'        Results.Append(InsufficentRights & "|")
'        Results.Append(Message)

'        Return Results.ToString
'    End Function

'    '    <System.Web.Services.WebMethod()> _
'    'Public Shared Function GetCheckoutAgent(ByVal CallbackID As Int64, ByVal PageCode As String) As String
'    '        Return Common.GetCheckoutAgent(CallbackID, PageCode)
'    '    End Function
'    '#End Region
'End Class