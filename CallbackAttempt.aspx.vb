
Imports System.Data

Partial Class CallbackAttempt
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As Rights
    Private cIndexSess As IndexSession
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim RequestActionResults As RequestActionResults
        Dim RequestAction As RequestActionEnum
        Dim ResponseAction As ResponseActionEnum
        Dim BehaviorOnSuccessfulSave As BehaviorOnSuccessfulSaveEnum
        Dim BehaviorOnCancelButtonClick As BehaviorOnCancelButtonClickEnum

        Try

            ' ___ Get Enviro from Session
            cEnviro = CType(Session("Enviro"), Enviro)

            ' ___ Instantiate objects
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Get the index session object
            cIndexSess = Session("IndexSession")


            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New Rights(cEnviro, Page)
            RightsRqd.SetValue(Rights.CallbackView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsAndTopicsHidden(cRights.RightsColl)

            ' ___ Successful save and cancel button behavior
            BehaviorOnSuccessfulSave = BehaviorOnSuccessfulSaveEnum.ReturnToCallerAfterRender
            BehaviorOnCancelButtonClick = BehaviorOnCancelButtonClickEnum.DisplayPage

            ' ___ Get RequestAction
            RequestAction = Common.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            RequestActionResults = ExecuteRequestAction(RequestAction, BehaviorOnSuccessfulSave, BehaviorOnCancelButtonClick)
            ResponseAction = RequestActionResults.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseActionEnum.ReturnToCallingPageDirectly Then
                cIndexSess.PageReturnOnLoadMessage = RequestActionResults.Message
                Response.Redirect("Index.aspx?CalledBy=Child")
            ElseIf ResponseAction = ResponseActionEnum.ReturnToCallingPageAfterRender Then
                litHiddens.Text = "<input type=""hidden"" id=""hdNotifyParentThenClose"" value=""1"" /><input type=""hidden"" id=""hdLoggedInUserID"" name=""hdLoggedInUserID"" value=""" & cEnviro.LoggedInUserID & """ /><input type=""hidden"" id=""hdCallbackID"" name=""hdCallbackID"" value=""" & Request.QueryString("CallbackID") & """ />"
            Else
                DisplayPage(ResponseAction)
                If Not RequestActionResults.Message = Nothing Then
                    litMessage.Text = "<script language='javascript'>alert('" & Common.ToJSAlert(RequestActionResults.Message) & "')</script>"
                End If
            End If

            ' ___ Display enviroment
            PageCaption.Text = Common.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & Enviro.DBHost & "'>"

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj("Error #750: CallbackAttempt Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function ExecuteRequestAction(ByVal RequestAction As RequestActionEnum, ByVal BehaviorOnSuccessfulSave As BehaviorOnSuccessfulSaveEnum, ByVal BehaviorOnCancelButtonClick As BehaviorOnCancelButtonClickEnum) As RequestActionResults
        Dim ValidationResults As New Results
        Dim SaveResults As Results
        Dim MyResults As New RequestActionResults

        Try

            Select Case RequestAction
                Case RequestActionEnum.ChangeTab
                    MyResults.ResponseAction = ResponseActionEnum.ChangeTab


                Case RequestActionEnum.CancelChanges
                    MyResults.ResponseAction = ExecuteRequestAction_BehaviorOnCancelButtonClick(MyResults, BehaviorOnCancelButtonClick)

                Case RequestActionEnum.ReturnToCallingPage
                    MyResults.ResponseAction = ExecuteRequestAction_BehaviorOnCancelButtonClick(MyResults, BehaviorOnCancelButtonClick)

                Case RequestActionEnum.CreateNew
                    MyResults.ResponseAction = ResponseActionEnum.DisplayBlank

                Case RequestActionEnum.SaveNew
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ExecuteRequestAction_BehaviorOnSuccessfulSave(MyResults, BehaviorOnSuccessfulSave)
                            MyResults.Message = SaveResults.Message
                        Else
                            MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew
                            MyResults.Message = SaveResults.Message
                        End If

                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew
                        MyResults.ObtainConfirm = True
                        MyResults.Message = ValidationResults.Message
                    Else
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew
                        MyResults.Message = ValidationResults.Message
                    End If

                Case RequestActionEnum.LoadExisting
                    MyResults.ResponseAction = ResponseActionEnum.DisplayExisting

                Case RequestActionEnum.SaveExisting
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ExecuteRequestAction_BehaviorOnSuccessfulSave(MyResults, BehaviorOnSuccessfulSave)
                            MyResults.Message = SaveResults.Message
                        Else
                            MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                            MyResults.Message = SaveResults.Message
                        End If

                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                        MyResults.ObtainConfirm = True
                        MyResults.Message = ValidationResults.Message
                    Else
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                        MyResults.Message = ValidationResults.Message
                    End If

                Case RequestActionEnum.Other
                    MyResults.ResponseAction = ResponseActionEnum.Other

            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #751: CallbackAtttempt ExecuteRequestAction. " & ex.Message, ex)
        End Try
    End Function

    Private Function ExecuteRequestAction_BehaviorOnCancelButtonClick(ByRef MyResults As RequestActionResults, ByVal BehaviorOnCancelButtonClick As BehaviorOnCancelButtonClickEnum) As ResponseActionEnum
        Dim ResponseAction As ResponseActionEnum

        Select Case BehaviorOnCancelButtonClick
            Case BehaviorOnCancelButtonClickEnum.DisplayPage
                Select Case cIndexSess.CallbackID
                    Case String.Empty
                        ResponseAction = ResponseActionEnum.DisplayUserInputNew
                    Case Else
                        ResponseAction = ResponseActionEnum.DisplayExisting
                End Select
            Case BehaviorOnCancelButtonClickEnum.ReturnToCallerDirectly
                ResponseAction = ResponseActionEnum.ReturnToCallingPageDirectly
            Case BehaviorOnCancelButtonClickEnum.ReturnToCallerAfterRender
                ResponseAction = ResponseActionEnum.ReturnToCallingPageAfterRender
        End Select
        Return ResponseAction
    End Function


    Private Function ExecuteRequestAction_BehaviorOnSuccessfulSave(ByRef MyResults As RequestActionResults, ByVal BehaviorOnSuccessfulSave As BehaviorOnSuccessfulSaveEnum) As ResponseActionEnum
        Dim ResponseAction As ResponseActionEnum

        Select Case BehaviorOnSuccessfulSave
            Case BehaviorOnSuccessfulSaveEnum.DisplayPage
                Select Case cIndexSess.CallbackID
                    Case String.Empty
                        ResponseAction = ResponseActionEnum.DisplayUserInputNew
                    Case Else
                        ResponseAction = ResponseActionEnum.DisplayExisting
                End Select
            Case BehaviorOnSuccessfulSaveEnum.ReturnToCallerDirectly
                ResponseAction = ResponseActionEnum.ReturnToCallingPageDirectly
            Case BehaviorOnSuccessfulSaveEnum.ReturnToCallerAfterRender
                ResponseAction = ResponseActionEnum.ReturnToCallingPageAfterRender
        End Select
        Return ResponseAction
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestActionEnum) As Results
        Dim MyResults As New Results
        Dim ErrColl As New Collection

        Try

            ' ___ Trim the textbox input
            txtLeftMessageWith.Text = txtLeftMessageWith.Text.Trim

            Common.ValidateDropDown(ErrColl, ddEnroller, 1, "enroller not provided")
            Common.ValidateDropDown(ErrColl, ddAction, 1, "action not provided")

            Select Case Common.DropdownGetTextFromValue(ddAction, ddAction.SelectedValue).ToLower
                Case "left message with person"
                    If txtLeftMessageWith.Text.Length = 0 Then
                        Common.ValidateErrorOnly(ErrColl, "name of person with whom message left not provided")
                    End If
                Case Else
                    If txtLeftMessageWith.Text.Length > 0 Then
                        Common.ValidateErrorOnly(ErrColl, "name of person with whom message left not consistent with action")
                    End If
            End Select

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Message = Common.ErrCollToString(ErrColl, "Not saved. Please correct the following:", True)
            End If

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #752: CallbackAtttempt IsDataValid. " & ex.Message, ex)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestActionEnum) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results
        Dim RecID As Integer
        Dim EST As DateTime

        Try

            EST = Common.GetServerDateTime

            If RequestAction = RequestActionEnum.SaveNew Then
                RecID = Common.GetNewRecordID("CallbackAttempt", "CallbackAttemptID")

                Sql.Append("INSERT INTO CallbackAttempt ")
                Sql.Append("(CallbackAttemptID, CallbackID, AttemptDate, EnrollerID, CallbackAttemptStatusCode, LeftMessageWith, LogicalDelete, AddDate, ChangeDate")
                Sql.Append(") VALUES (")
                Sql.Append(Common.NumOutHandler(RecID, False, False) & ", ")
                Sql.Append(Common.NumOutHandler(Request.QueryString("CallbackID").ToString, False, False) & ", ")
                Sql.Append(Common.DateOutHandler(EST, False, True) & ", ")
                Sql.Append(Common.StrOutHandler(ddEnroller.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append(Common.StrOutHandler(ddAction.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append(Common.StrOutHandler(txtLeftMessageWith.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("0, ")
                Sql.Append(Common.DateOutHandler(EST, False, True) & ", ")
                Sql.Append(Common.DateOutHandler(EST, False, True) & ")")

            Else

                Sql.Append("UPDATE CallbackAttempt SET ")
                Sql.Append("EnrollerID = " & Common.StrOutHandler(ddEnroller.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append("CallbackAttemptStatusCode = " & Common.StrOutHandler(ddAction.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append("LeftMessageWith = " & Common.StrOutHandler(txtLeftMessageWith.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("ChangeDate = " & Common.DateOutHandler(EST, False, True) & " ")
                Sql.Append("WHERE CallbackAttemptID = " & Request.QueryString("CallbackAttemptID"))

            End If

            ' ___ Save CallbackMaster record
            Common.ExecuteNonQuery(Sql.ToString)

            MyResults.Success = True
            MyResults.Message = "Record saved."
            If RequestAction = RequestActionEnum.SaveNew Then
                cIndexSess.CallbackID = RecID.ToString
            End If

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #758: CallbackAtttempt PerformSave_CallbackPhone. " & ex.Message, ex)
        End Try
    End Function

    Private Sub DisplayPage(ByVal ResponseAction As ResponseActionEnum)
        Try

            If Request.QueryString("CallbackID") <> Nothing AndAlso Request.QueryString("CallbackID") <> String.Empty Then
                cIndexSess.CallbackID = Request.QueryString("CallbackID")
            End If

            If Not Page.IsPostBack Then
                DisplayPage_PopulateDropdowns()
            End If

            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank
                    DisplayPage_DisplayBlank()

                Case ResponseActionEnum.DisplayUserInputNew, ResponseActionEnum.DisplayUserInputExisting
                    '   No action

                Case ResponseActionEnum.DisplayExisting
                    DisplayPage_DisplayExisting()

            End Select

            DisplayPage_HandleHiddens(ResponseAction)

            DisplayPage_FormatControls()

        Catch ex As Exception
            Throw New Exception("Error #770: CallbackAtttempt DisplayPage. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_PopulateDropdowns()
        Dim Sql As String
        Dim dt As System.Data.DataTable

        Try

            ' ___ Enroller
            Sql = "SELECT EnrollerID = UserID, EnrollerName = LastName + ', ' + FirstName FROM UserManagement..Users WHERE StatusCode = 'ACTIVE' AND LocationID = 'HBG' AND Role IN ('ENROLLER', 'SUPERVISOR') ORDER BY LastName + FirstName"
            dt = Common.GetDT(Sql)
            dt.Rows.InsertAt(dt.NewRow, 0)
            dt.Rows(0)(0) = "0"
            dt.Rows(0)(1) = String.Empty
            ddEnroller.DataValueField = "EnrollerID"
            ddEnroller.DataTextField = "EnrollerName"
            ddEnroller.DataSource = dt
            ddEnroller.DataBind()

            ' ___ Action
            Sql = "SELECT CallbackAttemptStatusCode, AttemptDescription FROM Codes_CallbackAttemptStatus ORDER BY Seq"
            dt = Common.GetDT(Sql)
            dt.Rows.InsertAt(dt.NewRow, 0)
            dt.Rows(0)(0) = "0"
            dt.Rows(0)(1) = String.Empty
            ddAction.DataValueField = "CallbackAttemptStatusCode"
            ddAction.DataTextField = "AttemptDescription"
            ddAction.DataSource = dt
            ddAction.DataBind()

        Catch ex As Exception
            Throw New Exception("Error #772: CallbackAtttempt DisplayPage_PopulateDropdowns. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_DisplayBlank()
        Dim sb As New System.Text.StringBuilder
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try

            'sb.Append("SELECT cm.TicketNumber, cs.StateName, ClientName = cc.Name, ")
            sb.Append("SELECT TicketNumber = " & Common.GetTicketNumberSyntax("mast") & ", state.StateName, ClientName = client.Name, ")
            sb.Append("EmpName = rtrim(mast.EmpLastName + ', ' + mast.EmpFirstName + ' ' + mast.EmpMI), ")
            sb.Append("stat.StatusDescription ")
            sb.Append("FROM CallbackMaster mast ")
            sb.Append("INNER JOIN Codes_Status stat ON mast.StatusCode = stat.StatusCode ")
            sb.Append("INNER JOIN ProjectReports..Codes_ClientID client ON mast.ClientID = client.ClientID ")
            sb.Append("INNER JOIN UserManagement..Codes_State state ON mast.State = state.StateCode ")
            sb.Append("WHERE CallbackID = " & Request.QueryString("CallbackID"))
            dt = Common.GetDT(sb.ToString)
            dr = dt.Rows(0)

            lblTicketNumber.Text = "Ticket #: " & dr("TicketNumber")
            lblStatus.Text = "Status: " & dr("StatusDescription")
            txtDate.Text = Common.FormatDate(Common.GetServerDateTime, FormatDateEnum.SpecialFormat2)
            txtState.Text = dr("StateName")
            txtClient.Text = dr("ClientName")
            txtEmpName.Text = dr("EmpName")

        Catch ex As Exception
            Throw New Exception("Error #774: CallbackAtttempt DisplayPage_DisplayBlank. " & ex.Message, ex)
        End Try
    End Sub

    'Private Sub DisplayPage_DisplayExisting()
    '    Dim sb As New System.Text.StringBuilder
    '    Dim dt As DataTable
    '    Dim dr As DataRow

    '    Try

    '        ' ___ Clear the dropdowns
    '        ddEnroller.ClearSelection()
    '        ddAction.ClearSelection()

    '        sb.Append("SELECT att.AttemptDate, att.EnrollerID, att.CallbackAttemptStatusCode, ")
    '        sb.Append("EnrollerName = u.LastName + ', ' + u.FirstName, ")
    '        sb.Append("maststatus.StatusDescription, ")
    '        sb.Append("attstatus.AttemptDescription, att.LeftMessageWith, ")
    '        sb.Append("TicketNumber = " & Common.GetTicketNumberSyntax("mast") & ", state.StateName, ClientName = client.Name, ")
    '        sb.Append("EmpName = mast.EmpLastName + ', ' + mast.EmpFirstName + ' ' + mast.EmpMI ")
    '        sb.Append("FROM CallbackAttempt att ")
    '        sb.Append("INNER JOIN CallbackMaster mast ON att.CallbackID = mast.CallbackID ")
    '        sb.Append("INNER JOIN Codes_Status maststatus ON mast.StatusCode = maststatus.StatusCode ")
    '        sb.Append("INNER JOIN ProjectReports..Codes_ClientID client ON mast.ClientID = client.ClientID ")
    '        sb.Append("INNER JOIN UserManagement..Users u ON att.EnrollerID = u.UserID ")
    '        sb.Append("INNER JOIN UserManagement..Codes_State state ON mast.State = state.StateCode ")
    '        sb.Append("INNER JOIN Codes_CallbackAttemptStatus attstatus ON att.CallbackAttemptStatusCode = attstatus.CallbackAttemptStatusCode ")
    '        sb.Append("WHERE att.CallbackAttemptID = " & Request.QueryString("CallbackAttemptID"))
    '        dt = Common.GetDT(sb.ToString)
    '        dr = dt.Rows(0)

    '        lblTicketNumber.Text = "Ticket #: " & dr("TicketNumber")
    '        lblStatus.Text = "Status: " & dr("StatusDescription")


    '        txtDate.Text = Common.FormatDate(dr("AttemptDate"), FormatDateEnum.SpecialFormat2)
    '        txtEmpName.Text = dr("EmpName")
    '        txtClient.Text = dr("ClientName")
    '        txtState.Text = dr("StateName")

    '        ddEnroller.Items.FindByValue(dr("EnrollerID")).Selected = True
    '        txtEnroller.Text = dr("EnrollerName")
    '        ddAction.Items.FindByValue(dr("CallbackAttemptStatusCode")).Selected = True
    '        txtAction.Text = dr("AttemptDescription")
    '        txtLeftMessageWith.Text = dr("LeftMessageWith")

    '    Catch ex As Exception
    '        Throw New Exception("Error #777: CallbackAttempt DisplayPage_DisplayExisting. " & ex.Message, ex)
    '    End Try
    'End Sub

    Private Sub DisplayPage_DisplayExisting()
        Dim sb As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim dr As DataRow

        Try

            ' ___ Clear the dropdowns
            ddEnroller.ClearSelection()
            ddAction.ClearSelection()



            If Request.QueryString("CallbackAttemptID") = Nothing Then


            Else
                sb.Append("SELECT att.AttemptDate, att.EnrollerID, att.CallbackAttemptStatusCode, ")
                sb.Append("EnrollerName = u.LastName + ', ' + u.FirstName, ")
                sb.Append("maststatus.StatusDescription, ")
                sb.Append("attstatus.AttemptDescription, att.LeftMessageWith, ")
                sb.Append("TicketNumber = " & Common.GetTicketNumberSyntax("mast") & ", state.StateName, ClientName = client.Name, ")
                sb.Append("EmpName = mast.EmpLastName + ', ' + mast.EmpFirstName + ' ' + mast.EmpMI ")
                sb.Append("FROM CallbackAttempt att ")
                sb.Append("INNER JOIN CallbackMaster mast ON att.CallbackID = mast.CallbackID ")
                sb.Append("INNER JOIN Codes_Status maststatus ON mast.StatusCode = maststatus.StatusCode ")
                sb.Append("INNER JOIN ProjectReports..Codes_ClientID client ON mast.ClientID = client.ClientID ")
                sb.Append("INNER JOIN UserManagement..Users u ON att.EnrollerID = u.UserID ")
                sb.Append("INNER JOIN UserManagement..Codes_State state ON mast.State = state.StateCode ")
                sb.Append("INNER JOIN Codes_CallbackAttemptStatus attstatus ON att.CallbackAttemptStatusCode = attstatus.CallbackAttemptStatusCode ")
                sb.Append("WHERE att.CallbackAttemptID = " & Request.QueryString("CallbackAttemptID"))
                dt = Common.GetDT(sb.ToString)
                dr = dt.Rows(0)


                lblTicketNumber.Text = "Ticket #: " & dr("TicketNumber")
                lblStatus.Text = "Status: " & dr("StatusDescription")

                txtDate.Text = Common.FormatDate(dr("AttemptDate"), FormatDateEnum.SpecialFormat2)
                txtEmpName.Text = dr("EmpName")
                txtClient.Text = dr("ClientName")
                txtState.Text = dr("StateName")


                ddEnroller.Items.FindByValue(dr("EnrollerID")).Selected = True
                txtEnroller.Text = dr("EnrollerName")
                ddAction.Items.FindByValue(dr("CallbackAttemptStatusCode")).Selected = True
                txtAction.Text = dr("AttemptDescription")
                txtLeftMessageWith.Text = dr("LeftMessageWith")
            End If






        Catch ex As Exception
            Throw New Exception("Error #777: CallbackAttempt DisplayPage_DisplayExisting. " & ex.Message, ex)
        End Try
    End Sub


    Private Sub DisplayPage_HandleHiddens(ByVal ResponseAction As ResponseActionEnum)
        Dim sb As New System.Text.StringBuilder
        Dim ExistingRecordInd As Integer

        Try

            ' ___ Existing record indicator
            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank, ResponseActionEnum.DisplayUserInputNew
                    ExistingRecordInd = 0
                Case ResponseActionEnum.DisplayExisting, ResponseActionEnum.DisplayUserInputExisting
                    ExistingRecordInd = 1
                Case Else
                    ExistingRecordInd = CType(Request.Form("hdExistingRecordInd"), System.Int64)
            End Select

            sb.Append("<input type=""hidden"" id=""hdNotifyParentThenClose"" value=""0"" />")
            sb.Append("<input type=""hidden"" id=""hdExistingRecordInd"" name=""hdExistingRecordInd"" value=""" & ExistingRecordInd.ToString & """ />")
            sb.Append("<input type=""hidden"" id=""hdLoggedInUserID"" name=""hdLoggedInUserID"" value=""" & cEnviro.LoggedInUserID & """ />")
            sb.Append("<input type=""hidden"" id=""hdCallbackID"" name=""hdCallbackID"" value=""" & cIndexSess.CallbackID & """ />")
            litHiddens.Text = sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #779: CallbackAttempt DisplayPage_HandleHiddens. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_FormatControls()
        Dim sb As New System.Text.StringBuilder

        Try

            Style.AddStyle(txtDate, Style.StyleType.NoneditableGrayed, 180)
            Style.AddStyle(txtEmpName, Style.StyleType.NoneditableGrayed, 230)
            Style.AddStyle(txtState, Style.StyleType.NoneditableGrayed, 140)
            Style.AddStyle(txtClient, Style.StyleType.NoneditableGrayed, 180)

            If cRights.HasThisRight(Rights.CallbackEdit) Then
                ddEnroller.Visible = True
                txtEnroller.Visible = False

                ddAction.Visible = True
                txtAction.Visible = False

                Style.AddStyle(txtLeftMessageWith, Style.StyleType.NormalEditable, 185)

                sb.Append(Common.BuildButton("Save", Nothing, 24, 255, "Save"))
                sb.Append(Common.BuildButton("Cancel", Nothing, 136, 255, "Cancel"))
                sb.Append(Common.BuildButton("Close", Nothing, 248, 255, "Close"))
                litSaveCancelReturn.Text = sb.ToString

            Else

                ddEnroller.Visible = False
                txtEnroller.Visible = True
                Style.AddStyle(txtEnroller, Style.StyleType.NoneditableGrayed, 230)

                ddAction.Visible = False
                txtAction.Visible = True
                Style.AddStyle(txtAction, Style.StyleType.NoneditableGrayed, 200)
                Style.AddStyle(txtLeftMessageWith, Style.StyleType.NoneditableGrayed, 185)

                sb.Append(Common.BuildButton("Close", Nothing, 24, 255, "Close"))
                litSaveCancelReturn.Text = sb.ToString

            End If

        Catch ex As Exception
            Throw New Exception("Error #778: CallbackAttempt FormatControls. " & ex.Message, ex)
        End Try
    End Sub

#Region " Ajax calls "
    <System.Web.Services.WebMethod()>
    Public Shared Function RefreshCheckout(ByVal LoggedInUserID As String, ByVal CallbackID As Int64) As String
        Return Common.RefreshCheckout(LoggedInUserID, CallbackID)
    End Function

    '    <System.Web.Services.WebMethod()> _
    'Public Shared Function GetCheckoutAgent(ByVal CallbackID As Int64, ByVal PageCode As String) As String
    '        Return Common.GetCheckoutAgent(CallbackID, PageCode)
    '    End Function
#End Region
End Class
