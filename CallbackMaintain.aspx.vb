Imports System.Data

Partial Class CallbackMaintain
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As Rights
    Private cIndexSess As IndexSession
    Private cDiag As New System.Text.StringBuilder
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        PageLoad()
    End Sub

    Private Sub PageLoad()
        Dim sb As New System.Text.StringBuilder
        Dim RequestActionResults As RequestActionResults
        Dim RequestAction As RequestActionEnum
        Dim ResponseAction As ResponseActionEnum
        Dim BehaviorOnSuccessfulSave As BehaviorOnSuccessfulSaveEnum
        Dim BehaviorOnCancelButtonClick As BehaviorOnCancelButtonClickEnum

        Try

            cDiag.Append("PageLoad | ")

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

            ' Load the session variables
            LoadVariables()

            ' ___ Get RequestAction
            RequestAction = Common.GetRequestAction(Page, cDiag)

            ' ___ Execute the RequestAction
            RequestActionResults = ExecuteRequestAction(RequestAction, BehaviorOnSuccessfulSave, BehaviorOnCancelButtonClick)
            ResponseAction = RequestActionResults.ResponseAction

            cDiag.Append("RespAct: " & ResponseAction.ToString & "| ")

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseActionEnum.ReturnToCallingPageDirectly Then
                cIndexSess.PageReturnOnLoadMessage = RequestActionResults.Message
                Response.Redirect("Index.aspx?CalledBy=Child")
            ElseIf ResponseAction = ResponseActionEnum.ReturnToCallingPageAfterRender Then
                sb.Append("<input type=""hidden"" id=""hdNotifyParentThenClose"" value=""1"" />")
                sb.Append("<input type=""hidden"" id=""hdCallbackID"" name=""hdCallbackID"" value=""" & Request.QueryString("CallbackID") & """ />")
                sb.Append("<input type=""hidden"" id=""hdLoggedInUserID"" name=""hdLoggedInUserID"" value=""" & cEnviro.LoggedInUserID & """ />")
                litHiddens.Text = sb.ToString
            Else
                DisplayPage(ResponseAction)
                If Not RequestActionResults.Message = Nothing Then
                    litMessage.Text = "<script language='javascript'>alert('" & Common.ToJSAlert(RequestActionResults.Message) & "')</script>"
                End If
            End If


            ' ___ Display enviroment
            PageCaption.Text = Common.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & Enviro.DBHost & "'>"

            txtDiagnostic.Text = cDiag.ToString

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj("Error #650: CallMaintain Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Sub LoadVariables()
        If Not Page.IsPostBack Then
            Select Case Request.QueryString("CallType").ToLower
                Case "new"
                    cIndexSess.CallbackID = String.Empty
                Case "existing"
                    cIndexSess.CallbackID = Request.QueryString("CallbackID")
            End Select
        End If
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
            Throw New Exception("Error #651: ProductMaintain ExecuteRequestAction. " & ex.Message, ex)
        End Try
    End Function

    Private Function ExecuteRequestAction_BehaviorOnCancelButtonClick(ByRef MyResults As RequestActionResults, ByVal BehaviorOnCancelButtonClick As BehaviorOnCancelButtonClickEnum) As ResponseActionEnum
        Dim ResponseAction As ResponseActionEnum

        Select Case BehaviorOnCancelButtonClick
            Case BehaviorOnCancelButtonClickEnum.DisplayPage
                Select Case cIndexSess.CallbackID
                    Case String.Empty
                        ResponseAction = ResponseActionEnum.ClearAll
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
            txtEmpLastName.Text = txtEmpLastName.Text.Trim
            txtEmpFirstName.Text = txtEmpFirstName.Text.Trim
            txtEmpMI.Text = txtEmpMI.Text.Trim
            txtNumEmployeeCalls.Text = txtNumEmployeeCalls.Text.Trim
            txtContactPhoneNumber1.Text = txtContactPhoneNumber1.Text.Trim
            txtContactPhoneNumber2.Text = txtContactPhoneNumber2.Text.Trim
            txtContactPhoneNumber3.Text = txtContactPhoneNumber3.Text.Trim
            txtContactPhoneNumber1Ext.Text = txtContactPhoneNumber1Ext.Text.Trim
            txtContactPhoneNumber2Ext.Text = txtContactPhoneNumber2Ext.Text.Trim
            txtContactPhoneNumber3Ext.Text = txtContactPhoneNumber3Ext.Text.Trim
            txtContactType1.Text = txtContactType1.Text.Trim
            txtContactType2.Text = txtContactType2.Text.Trim
            txtContactType3.Text = txtContactType3.Text.Trim
            txtContactBestTime1.Text = txtContactBestTime1.Text.Trim
            txtContactBestTime2.Text = txtContactBestTime2.Text.Trim
            txtContactBestTime3.Text = txtContactBestTime3.Text.Trim
            'txtAuthPerson.Text = txtAuthPerson.Text.Trim
            'txtAuthRelationship.Text = txtAuthRelationship.Text.Trim
            txtNotes.Text = txtNotes.Text.Trim


            ' ____ Home tab
            Common.ValidateDropDown(ErrColl, ddClientID, 1, "client not provided")
            Common.ValidateDropDown(ErrColl, ddState, 1, "state not provided")
            Common.ValidateStringField(ErrColl, txtEmpLastName.Text, 2, "emp last name not provided")
            Common.ValidateStringField(ErrColl, txtEmpFirstName.Text, 1, "emp first name not provided")
            Common.ValidateNumericField(ErrColl, txtNumEmployeeCalls.Text, False, "number of employee calls not provided")

            ' ___ Setup tab
            Common.ValidateDropDown(ErrColl, ddOverflowAgent, 1, "overflow agent not provided")
            Common.ValidateDropDown(ErrColl, ddEnrollWin, 1, "enrollment window code not provided")
            Common.ValidateDropDown(ErrColl, ddCallPurpose, 1, "call type not provided")

            ' ___ Contact tab
            Common.ValidateStringField(ErrColl, Common.DBFormatPhone(txtContactPhoneNumber1.Text), 10, "phone #1 number not provided")
            Common.ValidateStringField(ErrColl, txtContactBestTime1.Text, 5, "phone #1 best time to call not provided (minimum 5 characters)")
            If txtContactPhoneNumber2.Text.Length < 7 AndAlso (txtContactType2.Text.Length > 0 Or txtContactBestTime2.Text.Length > 0) Then
                Common.ValidateErrorOnly(ErrColl, "contact #2 info provided without phone number")
            End If
            If txtContactPhoneNumber3.Text.Length < 7 AndAlso (txtContactType3.Text.Length > 0 Or txtContactBestTime3.Text.Length > 0) Then
                Common.ValidateErrorOnly(ErrColl, "contact #3 info provided without phone number")
            End If

            ' ___ Authorized person/relationship
            'If txtAuthPerson.Text.Length = 0 AndAlso txtAuthRelationship.Text.Length > 0 Then
            '    Common.ValidateErrorOnly(ErrColl, "authorization relationship provided but person not provided")
            'ElseIf txtAuthPerson.Text.Length > 0 AndAlso txtAuthRelationship.Text.Length = 0 Then
            '    Common.ValidateErrorOnly(ErrColl, "authorization person provided but relationship not provided")
            'End If

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Message = Common.ErrCollToString(ErrColl, "Not saved. Please correct the following:", True)
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #652: CallbackMaintain IsDataValid. " & ex.Message, ex)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestActionEnum) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim SqlNote As New System.Text.StringBuilder
        Dim MyResults As New Results
        Dim RecID As Integer
        Dim ChangeDate As DateTime
        Dim StatusCode As String
        Dim EnrollWinCode As String
        Dim dt As DataTable
        Dim dtEnrollWin As DataTable = Nothing
        Dim dr As DataRow
        Dim AcesInd As Boolean
        Dim EmpEnrollWinRecordFound As Boolean
        Dim CallbackNoteRecID As Integer


        Try

            AcesInd = Common.IsAcesClient(ddClientID.SelectedValue)
            ChangeDate = Common.GetServerDateTime
            EnrollWinCode = Request.Form("hdEnrollWinCode")

            ' ___ Status code
            If AcesInd Then
                If txtEmpID.Text.Length = 0 Then
                    StatusCode = "TCA"
                Else
                    StatusCode = "CB"
                End If
            Else
                StatusCode = "CB"
            End If


            ' ___ Get enroll win data
            'If AcesInd Then
            '    Sql.Append("SELECT TOP 1 ActivityID, StartDate, EndDate FROM " & ddClientID.SelectedValue & "..EmpEnrollWin WHERE EmpID = '" & txtEmpID.Text & "' AND EnrollWinCode = '" & EnrollWinCode & "'")
            '    dtEnrollWin = Common.GetDT(Sql.ToString)
            '    If dtEnrollWin.Rows.Count > 0 Then
            '        EmpEnrollWinRecordFound = True
            '    End If
            '    Sql.Length = 0
            'End If



            If AcesInd Then
                Sql.Append("SELECT TOP 1 ActivityID, StartDate, EndDate FROM " & ddClientID.SelectedValue & "..EmpEnrollWin ")
                Sql.Append("WHERE EmpID = '" & txtEmpID.Text & "' AND EnrollWinCode = '" & EnrollWinCode & "' AND ")
                Sql.Append("ProjectReports.dbo.ufn_IsDateBetween('" & Common.GetServerDateTime.ToString & "', StartDate, EndDate) = 1")
                dtEnrollWin = Common.GetDT(Sql.ToString)
                ' Sql.Append("SELECT TOP 1 ActivityID, StartDate, EndDate FROM " & ddClientID.SelectedValue & "..EmpEnrollWin WHERE EmpID = '" & txtEmpID.Text & "' AND EnrollWinCode = '" & EnrollWinCode & "'")
                If dtEnrollWin.Rows.Count > 0 Then
                    EmpEnrollWinRecordFound = True
                End If
                Sql.Length = 0
            End If

            If RequestAction = RequestActionEnum.SaveNew Then
                RecID = Common.GetNewRecordID("CallbackMaster", "CallbackID")

                Sql.Append("INSERT INTO CallbackMaster ")
                Sql.Append("(")
                Sql.Append("CallbackID, CreationDate, TicketNumber, ClientID, State, ")
                Sql.Append("EmpLastName, EmpFirstName, EmpMI, EmpID, ")
                Sql.Append("PreferSpanishInd, OverflowAgentID, EnrollWinCode, ")
                Sql.Append("PriorityTagInd, CallPurposeCode, StatusCode, ")
                Sql.Append("Notes, NumEmployeeCalls, EnrollWinActivityID, EnrollWinStartDate, EnrollWinEndDate, LogicalDelete, ")
                Sql.Append("AddDate, ChangeDate")
                Sql.Append(") ")

                Sql.Append(" Values ")

                Sql.Append("(")
                Sql.Append(Common.NumOutHandler(RecID, False, False) & ", ")
                Sql.Append(Common.DateOutHandler(txtDate.Text, False, True) & ", ")
                Sql.Append(Common.StrOutHandler(Replace(txtTicketNumber.Text, "-", ""), False, StringTreatEnum.SideQts) & ", ")
                Sql.Append(Common.StrOutHandler(ddClientID.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append(Common.StrOutHandler(ddState.SelectedValue, False, StringTreatEnum.SideQts) & ",")

                Sql.Append(Common.StrOutHandler(txtEmpLastName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(Common.StrOutHandler(txtEmpFirstName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(Common.StrOutHandler(txtEmpMI.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(Common.StrOutHandler(txtEmpID.Text, False, StringTreatEnum.SideQts) & ", ")

                Sql.Append(Common.BitOutHandler(ddPreferSpanishInd.SelectedValue, False, False) & ", ")
                Sql.Append(Common.StrOutHandler(ddOverflowAgent.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append(Common.StrOutHandler(EnrollWinCode, False, StringTreatEnum.SideQts) & ", ")

                Sql.Append(Common.BitOutHandler(ddPriorityTagInd.SelectedValue, False, False) & ", ")
                Sql.Append(Common.StrOutHandler(ddCallPurpose.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append(Common.StrOutHandler(StatusCode, False, StringTreatEnum.SideQts) & ", ")
                'Sql.Append(Common.StrOutHandler(txtAuthPerson.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")
                'Sql.Append(Common.StrOutHandler(txtAuthRelationship.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")




                Sql.Append(Common.StrOutHandler(String.Empty, False, StringTreatEnum.SideQts_SecApost) & ", ")

                Sql.Append(Common.NumOutHandler(txtNumEmployeeCalls.Text, False, False) & ", ")

                If EmpEnrollWinRecordFound Then
                    dr = dtEnrollWin.Rows(0)
                    Sql.Append(Common.StrOutHandler(dr("ActivityID"), True, StringTreatEnum.SideQts) & ", ")
                    Sql.Append(Common.DateOutHandler(dr("StartDate"), True, True) & ", ")
                    Sql.Append(Common.DateOutHandler(dr("EndDate"), True, True) & ", ")
                Else
                    Sql.Append("null, ")
                    Sql.Append("null, ")
                    Sql.Append("null, ")
                End If

                Sql.Append("0 , ")
                Sql.Append(Common.DateOutHandler(ChangeDate, False, True) & ", ")
                Sql.Append(Common.DateOutHandler(ChangeDate, False, True))
                Sql.Append(") ")

            Else

                Sql.Append("UPDATE CallbackMaster SET ")

                Sql.Append("ClientID = " & Common.StrOutHandler(ddClientID.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append("State = " & Common.StrOutHandler(ddState.SelectedValue, False, StringTreatEnum.SideQts) & ", ")

                Sql.Append("EmpLastName = " & Common.StrOutHandler(txtEmpLastName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("EmpFirstName = " & Common.StrOutHandler(txtEmpFirstName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("EmpMI = " & Common.StrOutHandler(txtEmpMI.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("EmpID = " & Common.StrOutHandler(txtEmpID.Text, False, StringTreatEnum.SideQts) & ", ")

                Sql.Append("PreferSpanishInd = " & Common.BitOutHandler(ddPreferSpanishInd.SelectedValue, False, False) & ", ")
                Sql.Append("OverflowAgentID = " & Common.StrOutHandler(ddOverflowAgent.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append("EnrollWinCode = " & Common.StrOutHandler(EnrollWinCode, False, StringTreatEnum.SideQts) & ", ")

                If EmpEnrollWinRecordFound Then
                    dr = dtEnrollWin.Rows(0)
                    Sql.Append("EnrollWinStartDate = " & Common.DateOutHandler(dr("StartDate"), True, True) & ", ")
                    Sql.Append("EnrollWinEndDate = " & Common.DateOutHandler(dr("EndDate"), True, True) & ", ")
                End If

                Sql.Append("PriorityTagInd = " & Common.BitOutHandler(ddPriorityTagInd.SelectedValue, False, False) & ", ")
                Sql.Append("CallPurposeCode = " & Common.StrOutHandler(ddCallPurpose.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                Sql.Append("StatusCode = " & Common.StrOutHandler(StatusCode, False, StringTreatEnum.SideQts) & ", ")
                'Sql.Append("AuthPerson = " & Common.StrOutHandler(txtAuthPerson.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")
                'Sql.Append("AuthRelationship = " & Common.StrOutHandler(txtAuthRelationship.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")

                Sql.Append("Notes = '', ")

                Sql.Append("NumEmployeeCalls = " & Common.NumOutHandler(txtNumEmployeeCalls.Text, False, False) & ", ")
                Sql.Append("ChangeDate = " & Common.DateOutHandler(ChangeDate, False, True) & " ")
                Sql.Append("WHERE CallbackID = " & cIndexSess.CallbackID)

            End If

            ' ___ Save CallbackMaster record
            Common.ExecuteNonQuery(Sql.ToString)

            ' ___ Load the new CallbackID
            If RequestAction = RequestActionEnum.SaveNew Then
                cIndexSess.CallbackID = RecID.ToString
            End If


            ' ___ CallbackNote

            ' ___ Is there an existing note record"
            dt = Common.GetDT("SELECT CallbackNoteID FROM CallbackNote WHERE CallbackID = " & cIndexSess.CallbackID.ToString)
            If dt.Rows.Count = 0 Then
                If txtNotes.Text.Length > 0 Then
                    CallbackNoteRecID = Common.GetNewRecordID("CallbackNote", "CallbackNoteID", 10001, 99999)
                    SqlNote.Append("INSERT INTO CallbackNote (CallbackNoteID, CallbackID, EnrollerID, Note, LogicalDelete, AddDate, ChangeDate) ")
                    SqlNote.Append("VALUES (")
                    SqlNote.Append(CType(CallbackNoteRecID, System.String) & ", ")
                    SqlNote.Append(CType(cIndexSess.CallbackID, System.String) & ", ")
                    SqlNote.Append(Common.StrOutHandler(ddOverflowAgent.SelectedValue, False, StringTreatEnum.SideQts) & ", ")
                    SqlNote.Append(Common.StrOutHandler(txtNotes.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                    SqlNote.Append("0, ")
                    SqlNote.Append(Common.DateOutHandler(ChangeDate, False, True) & ", ")
                    SqlNote.Append(Common.DateOutHandler(ChangeDate, False, True) & ")")
                    Common.ExecuteNonQuery(SqlNote.ToString)
                End If
            Else
                SqlNote.Append("UPDATE CallbackNote SET ")
                SqlNote.Append("Note = " & Common.StrOutHandler(txtNotes.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                SqlNote.Append("ChangeDate = " & Common.DateOutHandler(ChangeDate, False, True) & " ")
                SqlNote.Append("WHERE CallbackNoteID = " & dt.Rows(0)(0))
                Common.ExecuteNonQuery(SqlNote.ToString)
            End If







            ' ___ CallbackPhone
            Common.ExecuteNonQuery("DELETE CallbackPhone WHERE CallbackID = " & cIndexSess.CallbackID)

            If txtContactPhoneNumber1.Text.Length >= 7 Then
                PerformSave_CallbackPhone(txtContactPhoneNumber1, txtContactPhoneNumber1Ext, txtContactType1, txtContactBestTime1, 1)
            End If

            If txtContactPhoneNumber2.Text.Length >= 7 Then
                PerformSave_CallbackPhone(txtContactPhoneNumber2, txtContactPhoneNumber2Ext, txtContactType2, txtContactBestTime2, 2)
            End If

            If txtContactPhoneNumber3.Text.Length >= 7 Then
                PerformSave_CallbackPhone(txtContactPhoneNumber3, txtContactPhoneNumber3Ext, txtContactType3, txtContactBestTime3, 3)
            End If

            If RequestAction = RequestActionEnum.SaveNew AndAlso ddPriorityTagInd.SelectedValue = "1" Then
                PerformSave_PriorityTagNotify()
                cIndexSess.PageReturnOnLoadMessage = "Notification of new priority callback emailed to supervisors"
            End If

            MyResults.Success = True
            MyResults.Message = "Record saved."

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #653: CallbackMaintain PerformSave. " & ex.Message, ex)
        End Try
    End Function

    Public Sub PerformSave_PriorityTagNotify()
        Dim DesignatedDateTime As String
        Dim dt As DataTable
        Dim OverflowAgentText As String
        Dim sb As New System.Text.StringBuilder


        Try

            dt = Common.GetDT("SELECT FirstName + ' ' + LastName From UserManagement..Users WHERE UserID = '" & ddOverflowAgent.SelectedValue & "'")
            OverflowAgentText = dt.Rows(0)(0)
            DesignatedDateTime = Common.GetServerDateTime.ToString("MM/dd/yyyy hh:mm tt")

            sb.Append("A new callback has been created with its priority indicator set." & Chr(10) & Chr(10))
            sb.Append("Callback ticket: " & txtTicketNumber.Text & Chr(10))
            sb.Append("Time of notice: " & DesignatedDateTime & Chr(10))
            sb.Append("Overflow agent: " & OverflowAgentText & Chr(10))

            'Common.SendEmail("rbluestein@benefitvision.com", ddOverflowAgent.SelectedValue & "@benefitvision.com", "rbluestein@benefitvision.com", "Notice of priorty callback creation", sb.ToString)
            Common.SendEmail("HBG_SUPERVISORS@benefitvision.com", ddOverflowAgent.SelectedValue & "@benefitvision.com", Nothing, "Notice of priorty callback creation (" & Enviro.DBHost & ")", sb.ToString)
        Catch ex As Exception
            Throw New Exception("Error #654: CallbackMaintain PerformSave_PriorityTagNotify. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub PerformSave_CallbackPhone(ByRef txtPhoneNumber As TextBox, ByRef txtPhoneExtension As TextBox, ByRef txtType As TextBox, ByRef txtBestTime As TextBox, ByVal Seq As Integer)
        Dim RecID As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim ChangeDate As DateTime

        Try

            ChangeDate = Common.GetServerDateTime

            RecID = Common.GetNewRecordID("CallbackPhone", "CallbackPhoneID")
            Sql.Append("INSERT INTO CallbackPhone (CallbackPhoneID, CallbackID, PhoneNumber, PhoneExtension, Type, BestTime, Seq, LogicalDelete, AddDate, ChangeDate) ")
            Sql.Append(" VALUES ")
            Sql.Append("(")
            Sql.Append(RecID.ToString & ", ")
            Sql.Append(cIndexSess.CallbackID & ", ")
            'Sql.Append(Common.StrOutHandler(txtPhoneNumber.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")

            Sql.Append(Common.StrOutHandler(Common.DBFormatPhone(txtPhoneNumber.Text), False, StringTreatEnum.SideQts) & ", ")
            Sql.Append(Common.StrOutHandler(txtPhoneExtension.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")

            Sql.Append(Common.StrOutHandler(txtType.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(Common.StrOutHandler(txtBestTime.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(Seq.ToString & ", ")
            Sql.Append("0 ,")
            Sql.Append(Common.DateOutHandler(ChangeDate, False, True) & ", ")
            Sql.Append(Common.DateOutHandler(ChangeDate, False, True))
            Sql.Append(") ")
            Common.ExecuteNonQuery(Sql.ToString)

        Catch ex As Exception
            Throw New Exception("Error #658: CallbackMaintain PerformSave_CallbackPhone. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage(ByVal ResponseAction As ResponseActionEnum)
        Dim SelectedTab As String = Nothing
        Dim VIPInd As Integer
        Dim ExistingRecordInd As Integer
        Dim ClientID As String = Nothing
        Dim EmpIDInd As Integer
        Dim EnrollWinCode As String = Nothing

        Try

            cDiag.Append("DisplayPage | ")

            DisplayPage_AddAttributes()

            If Not Page.IsPostBack Then
                DisplayPage_PopulateDropdowns()
            End If

            DisplayPage_SelectTab(SelectedTab)

            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank
                    DisplayPage_DisplayBlank(ClientID, EnrollWinCode, ExistingRecordInd, EmpIDInd)

                Case ResponseActionEnum.ClearAll
                    DisplayPage_ClearAll(ClientID, EnrollWinCode, ExistingRecordInd, EmpIDInd)

                Case ResponseActionEnum.DisplayUserInputNew, ResponseActionEnum.DisplayUserInputExisting, ResponseActionEnum.ChangeTab
                    DisplayPage_Mult(ClientID, EnrollWinCode, ExistingRecordInd, EmpIDInd)

                Case ResponseActionEnum.DisplayExisting
                    DisplayPage_DisplayExisting(ClientID, EnrollWinCode, ExistingRecordInd, EmpIDInd)

                Case ResponseActionEnum.ChangeTab

                Case ResponseActionEnum.Other
                    DisplayPage_DisplayOther(VIPInd, ClientID, EnrollWinCode, ExistingRecordInd, EmpIDInd)

            End Select

            DisplayPage_HandleEnrollWin(ClientID, EnrollWinCode)

            DisplayPage_HandleHiddens(VIPInd, SelectedTab, ResponseAction, ClientID, EnrollWinCode, ExistingRecordInd, EmpIDInd)

            DisplayPage_FormatControls(ClientID)

        Catch ex As Exception
            Throw New Exception("Error #670: CallbackMaintain DisplayPage. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_AddAttributes()
        Try

            lnkSendEmail.Attributes.Add("href", "javascript:EscalateCallback()")
            lnkEmpIDClear.Attributes.Add("href", "javascript:EmpIDClear()")
            lnkLookupEmpID.Attributes.Add("href", "javascript:LookupEmpID()")
            'txtEmpID.Attributes.Add("onpropertychange", "EmpIDInserted()")
            btnEmpIDLookupNotifier.Attributes.Add("onclick", "EmpIDInserted()")

            'ddClientID.Attributes.Add("onpropertychange", "ClientIDChanged()")
            ddClientID.Attributes.Add("onchange", "ClientIDChanged()")

            ddEnrollWin.Attributes.Add("onchange", "HandleEnrollWinChange()")

            txtContactPhoneNumber1.Attributes.Add("onfocus", "NumericOnly(this)")
            txtContactPhoneNumber2.Attributes.Add("onfocus", "NumericOnly(this)")
            txtContactPhoneNumber3.Attributes.Add("onfocus", "NumericOnly(this)")

            txtContactPhoneNumber1.Attributes.Add("onblur", "PhoneBoxBlur(this)")
            txtContactPhoneNumber2.Attributes.Add("onblur", "PhoneBoxBlur(this)")
            txtContactPhoneNumber3.Attributes.Add("onblur", "PhoneBoxBlur(this)")

            txtContactPhoneNumber1.Attributes.Add("onkeyup", "PhoneBoxChange(this)")
            txtContactPhoneNumber2.Attributes.Add("onkeyup", "PhoneBoxChange(this)")
            txtContactPhoneNumber3.Attributes.Add("onkeyup", "PhoneBoxChange(this)")

            txtContactPhoneNumber1.Attributes.Add("onkeypress", "return KeypressNumeric(this)")
            txtContactPhoneNumber2.Attributes.Add("onkeypress", "return KeypressNumeric(this)")
            txtContactPhoneNumber3.Attributes.Add("onkeypress", "return KeypressNumeric(this)")

            txtNotes.Attributes.Add("onkeyup", "NotesMax(this)")
            txtNotes.Attributes.Add("onblur", "NotesMax(this)")

            txtNumEmployeeCalls.Attributes.Add("onblur", "NumericOnly(this)")
            'txtNumEmployeeCalls.Attributes.Add("onkeypress", "if (event.keyCode < 48 || event.keyCode > 57) {return false}")

            txtNumEmployeeCalls.Attributes.Add("onkeypress", "return KeypressNumeric(this)")

        Catch ex As Exception
            Throw New Exception("Error #671: CallbackMaintain DisplayPage_AddAttributes. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_PopulateDropdowns()
        Dim Sql As String
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try

            ' ___ Client
            Sql = "SELECT ClientID, ClientName = Name, AcesInd FROM ProjectReports..Codes_ClientID WHERE CallbackInd = 1 AND LogicalDelete = 0 ORDER BY ClientID"
            dt = Common.GetDT(Sql)

            dr = dt.NewRow
            dr(0) = String.Empty
            dr(1) = "0"
            dt.Rows.InsertAt(dr, 0)

            'dt.Rows.InsertAt(dt.NewRow, 0)
            'dt.Rows(0)(0) = "0"
            'dt.Rows(0)(1) = String.Empty
            ddClientID.DataValueField = "ClientID"
            ddClientID.DataTextField = "ClientID"
            ddClientID.DataSource = dt
            ddClientID.DataBind()


            ' ___ State
            Sql = "SELECT StateCode, StateName FROM UserManagement..Codes_State WHERE StateCode <> 'XX' ORDER BY StateName"
            dt = Common.GetDT(Sql)
            dt.Rows.InsertAt(dt.NewRow, 0)
            dt.Rows(0)(0) = "0"
            dt.Rows(0)(1) = String.Empty



            'dt.Rows.InsertAt(dt.NewRow, dt.Rows.Count)
            dr = dt.NewRow
            dr(0) = "PR"
            dr(1) = "Puerto Rico"
            dt.Rows.Add(dr)

            dr = dt.NewRow
            dr(0) = "ON"
            dr(1) = "Ontario"
            dt.Rows.Add(dr)


            ddState.DataValueField = "StateCode"
            ddState.DataTextField = "StateName"
            ddState.DataSource = dt
            ddState.DataBind()


            ' ___ OverflowAgent / Designated enroller
            'Sql = "SELECT AgentID = UserID, AgentName = LastName + ', ' + FirstName FROM UserManagement..Users WHERE StatusCode = 'ACTIVE' AND LocationID = 'HBG' AND Role IN ('ENROLLER', 'SUPERVISOR') ORDER BY LastName + FirstName"
            Sql = "SELECT AgentID = UserID, AgentName = LastName + ', ' + FirstName FROM UserManagement..Users WHERE CompanyID = 'BVI' AND StatusCode = 'ACTIVE' ORDER BY LastName + FirstName"
            dt = Common.GetDT(Sql)
            dt.Rows.InsertAt(dt.NewRow, 0)
            dt.Rows(0)(0) = "0"
            dt.Rows(0)(1) = String.Empty
            ddOverflowAgent.DataValueField = "AgentID"
            ddOverflowAgent.DataTextField = "AgentName"
            ddOverflowAgent.DataSource = dt
            ddOverflowAgent.DataBind()

            ddDesignatedEnroller.DataValueField = "AgentID"
            ddDesignatedEnroller.DataTextField = "AgentName"
            ddDesignatedEnroller.DataSource = dt
            ddDesignatedEnroller.DataBind()

            ' ___ EnrollWinCode
            ddEnrollWin.Items.Add(New ListItem("Client Required", "0"))

            ' ___ Priority tag
            ddPriorityTagInd.Items.Add(New ListItem("No", "0"))
            ddPriorityTagInd.Items.Add(New ListItem("Yes", "1"))

            ' ___ Prefer Spanish
            ddPreferSpanishInd.Items.Add(New ListItem("No", "0"))
            ddPreferSpanishInd.Items.Add(New ListItem("Yes", "1"))


            ' ___ Call purpose
            Sql = "SELECT CallPurposeCode, CallPurposeDescription FROM Codes_CallPurpose ORDER BY Seq"
            dt = Common.GetDT(Sql)
            dt.Rows.InsertAt(dt.NewRow, 0)
            dt.Rows(0)(0) = "0"
            dt.Rows(0)(1) = String.Empty
            ddCallPurpose.DataValueField = "CallPurposeCode"
            ddCallPurpose.DataTextField = "CallPurposeDescription"
            ddCallPurpose.DataSource = dt
            ddCallPurpose.DataBind()

        Catch ex As Exception
            Throw New Exception("Error #672: CallbackMaintain DisplayPage_PopulateDropdowns. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_SelectTab(ByRef SelectedTab As String)

        Try

            pnHome.Visible = False
            pnSetup.Visible = False
            pnContact.Visible = False
            pnAuthorize.Visible = False
            pnNotes.Visible = False
            pnEscalate.Visible = False

            If Page.IsPostBack Then
                SelectedTab = Request.Form("hdSelectedTab")
            Else
                SelectedTab = "Home"
            End If
            Select Case SelectedTab
                Case "Home"
                    pnHome.Visible = True
                Case "Setup"
                    pnSetup.Visible = True
                Case "Contact"
                    pnContact.Visible = True
                Case "Authorize"
                    pnAuthorize.Visible = True
                Case "Notes"
                    pnNotes.Visible = True
                Case "Escalate"
                    pnEscalate.Visible = True
            End Select

            cDiag.Append("DispPgSelTab: " & SelectedTab & " | ")

        Catch ex As Exception
            Throw New Exception("Error #673: CallbackMaintain DisplayPage_SelectTab. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_DisplayBlank(ByRef ClientID As String, ByRef EnrollWinCode As String, ByRef ExistingRecordInd As Integer, ByRef EmpIDInd As String)
        Dim TicketNumber As String
        Dim EST As DateTime

        Try

            ' ___ Date
            EST = Common.GetServerDateTime
            txtDate.Text = EST.ToString("MMM d yyyy h:mm") & EST.ToString("tt").ToLower

            ' ___ ClientID, ExistingRecordInd, EmpIDInd
            ClientID = "0"
            EnrollWinCode = "0"
            ExistingRecordInd = 0
            EmpIDInd = 0

            ' ___ Ticket Number
            TicketNumber = Common.GetNewRecordID("CallbackMaster", "TicketNumber", 10001, 99999)
            TicketNumber = TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(2)
            txtTicketNumber.Text = TicketNumber
            lblTicketNumber.Text = "Ticket #: " & TicketNumber

            ' ___ Status
            lblStatus.Text = GetStatus("NEW")

        Catch ex As Exception
            Throw New Exception("Error #674: CallbackMaintain DisplayPage_DisplayBlank. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_ClearAll(ByRef ClientID As String, ByRef EnrollWinCode As String, ByRef ExistingRecordInd As Integer, ByRef EmpIDInd As String)
        Try

            ' ___ ClientID, ExistingRecordInd, EmpIDInd
            ClientID = "0"
            EnrollWinCode = "0"
            ExistingRecordInd = 0
            EmpIDInd = 0

            ddState.ClearSelection()
            txtState.Text = String.Empty
            ddClientID.ClearSelection()
            txtClientID.Text = String.Empty
            txtEmpLastName.Text = String.Empty
            txtEmpFirstName.Text = String.Empty
            txtEmpMI.Text = String.Empty
            txtEmpID.Text = String.Empty
            txtNumEmployeeCalls.Text = String.Empty

            ddOverflowAgent.ClearSelection()
            txtOverflowAgent.Text = String.Empty
            ddEnrollWin.ClearSelection()
            txtEnrollWin.Text = String.Empty
            ddPriorityTagInd.ClearSelection()
            txtPriorityTagInd.Text = String.Empty
            ddCallPurpose.ClearSelection()
            txtCallPurpose.Text = String.Empty
            ddPreferSpanishInd.ClearSelection()
            txtPreferSpanishInd.Text = String.Empty

            txtContactPhoneNumber1.Text = String.Empty
            txtContactPhoneNumber2.Text = String.Empty
            txtContactPhoneNumber3.Text = String.Empty

            txtContactPhoneNumber1Ext.Text = String.Empty
            txtContactPhoneNumber2Ext.Text = String.Empty
            txtContactPhoneNumber3Ext.Text = String.Empty

            txtContactType1.Text = String.Empty
            txtContactType2.Text = String.Empty
            txtContactType3.Text = String.Empty

            txtContactBestTime1.Text = String.Empty
            txtContactBestTime2.Text = String.Empty
            txtContactBestTime3.Text = String.Empty

            'txtAuthPerson.Text = String.Empty
            'txtAuthRelationship.Text = String.Empty

            txtNotes.Text = String.Empty

            ddDesignatedEnroller.ClearSelection()
            txtEmailComments.Text = String.Empty


        Catch ex As Exception
            Throw New Exception("Error #680: CallbackMaintain DisplayPage_ClearAll. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_Mult(ByRef ClientID As String, ByRef EnrollWinCode As String, ByRef ExistingRecordInd As Integer, ByRef EmpIDInd As String)
        Try

            ClientID = Request.Form("hdClientID")
            EnrollWinCode = Request.Form("hdEnrollWinCode")
            ExistingRecordInd = 0
            EmpIDInd = CType(Request.Form("hdEmpIDInd"), System.Int64)

        Catch ex As Exception
            Throw New Exception("Error #675: CallbackMaintain DisplayPage_Mult. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_DisplayExisting(ByRef ClientID As String, ByRef EnrollWinCode As String, ByRef ExistingRecordInd As Integer, ByRef EmpIDInd As String)
        Dim sb As New System.Text.StringBuilder
        Dim Querypack As DBase.QueryPack
        Dim TicketNumber As String
        Dim dr As DataRow
        Dim dtAcesEmp As DataTable
        Dim dtNotes As DataTable
        Dim TagNum As Integer

        Try

            ' ___ Clear the dropdowns
            ddOverflowAgent.ClearSelection()
            ddDesignatedEnroller.ClearSelection()
            ddClientID.ClearSelection()
            ddEnrollWin.ClearSelection()
            ddPreferSpanishInd.ClearSelection()
            ddPriorityTagInd.ClearSelection()
            ddState.ClearSelection()
            ddCallPurpose.ClearSelection()

            TagNum = 1

            ' ___ Get the master record
            sb.Length = 0
            sb.Append("SELECT cbm.*, cs.StatusCode FROM CallbackMaster cbm ")
            sb.Append("INNER JOIN Codes_Status cs ON cbm.StatusCode = cs.StatusCode ")
            sb.Append("WHERE cbm.CallbackID = " & CType(cIndexSess.CallbackID, System.Int64))

            Querypack = Common.GetDTExtendedWithQueryPack(sb.ToString)
            If Not Querypack.Success Then
                Throw New Exception("Error #654a: DisplayPage. Unable to retrieve record")
            End If

            dr = Querypack.dt.Rows(0)

            TagNum = 2

            ' ___ Status
            lblStatus.Text = GetStatus(dr("StatusCode").ToText)

            TagNum = 3

            ' ___ ExistingRecordInd, EmpIDInd, EnrollWinCode
            ExistingRecordInd = 1
            EmpIDInd = CType(Request.Form("hdEmpIDInd"), System.Int64)
            EnrollWinCode = dr("EnrollWinCode").ToText

            TagNum = 4

            ' ___ Lookup EmpID panel
            ClientID = dr("ClientID").ToText
            ddClientID.Items.FindByValue(ClientID).Selected = True
            txtClientID.Text = ClientID

            If Common.IsAcesClient(ddClientID.SelectedValue) AndAlso dr("EmpID").ToText <> String.Empty Then

                TagNum = 5

                dtAcesEmp = Common.GetDT("SELECT FirstName, LastName, MI, State FROM " & dr("ClientID").ToText & "..Employee WHERE EmpID = '" & dr("EmpID").ToText & "'")

                TagNum = 6

                ddState.Items.FindByValue(dtAcesEmp.Rows(0)("State")).Selected = True

                TagNum = 7

                txtState.Text = dtAcesEmp.Rows(0)("State")

                TagNum = 8

                txtEmpLastName.Text = dtAcesEmp.Rows(0)("LastName")

                TagNum = 9
                txtEmpFirstName.Text = dtAcesEmp.Rows(0)("FirstName")
                TagNum = 10

                txtEmpMI.Text = dtAcesEmp.Rows(0)("MI")
            Else
                ddState.Items.FindByValue(dr("State").ToText).Selected = True
                txtState.Text = dr("State").ToText
                txtEmpLastName.Text = dr("EmpLastName").ToText
                txtEmpFirstName.Text = dr("EmpFirstName").ToText
                txtEmpMI.Text = dr("EmpMI").ToText
            End If

            TagNum = 11

            txtEmpID.Text = dr("EmpID").ToText
            txtNumEmployeeCalls.Text = dr("NumEmployeeCalls").ToText

            TagNum = 12

            ' ___ Setup panel
            TicketNumber = CType(dr("TicketNumber").ToText, System.String)
            txtTicketNumber.Text = TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(2)
            lblTicketNumber.Text = "Ticket #: " & TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(2)
            txtDate.Text = dr("CreationDate").ToSpecialFormat2 ' String("MMM d yyyy h:mm") & EST.ToString("tt").ToLower
            ddOverflowAgent.Items.FindByValue(dr("OverflowAgentID").ToText).Selected = True
            txtOverflowAgent.Text = dr("OverflowAgentID").ToText

            ddPriorityTagInd.Items.FindByValue(dr("PriorityTagInd").ToDropdown3State).Selected = True
            txtPriorityTagInd.Text = Common.DropdownGetTextFromValue(ddPriorityTagInd, dr("PriorityTagInd").ToDropdown3State)

            ddCallPurpose.Items.FindByValue(dr("CallPurposeCode").ToText).Selected = True
            txtCallPurpose.Text = Common.DropdownGetTextFromValue(ddCallPurpose, dr("CallPurposeCode").ToText)

            ddPreferSpanishInd.Items.FindByValue(dr("PreferSpanishInd").ToDropdown3State).Selected = True
            txtPreferSpanishInd.Text = Common.DropdownGetTextFromValue(ddPreferSpanishInd, dr("PreferSpanishInd").ToDropdown3State)

            ' ___ Notes panel
            dtNotes = Common.GetDT("SELECT TOP 1 Note FROM CallbackNote WHERE CallbackID = " & cIndexSess.CallbackID)
            If dtNotes.Rows.Count > 0 Then
                txtNotes.Text = dtNotes.Rows(0)(0)
            Else
                txtNotes.Text = String.Empty
            End If


            ' ___ Contact panel
            Querypack = Common.GetDTExtendedWithQueryPack("SELECT Seq, PhoneNumber, PhoneExtension, Type, BestTime FROM CallbackPhone WHERE CallbackID = " & CType(cIndexSess.CallbackID, System.Int64) & " ORDER BY Seq")
            If Not Querypack.Success Then
                Throw New Exception("Error #654b: DisplayPage. Unable to contact data. " & Querypack.TechErrMsg)
            End If

            For i = 0 To Querypack.dt.Rows.Count - 1
                dr = Querypack.dt.Rows(i)
                Select Case i
                    Case 0
                        txtContactPhoneNumber1.Text = Common.PageFormatPhone(dr("PhoneNumber").Value)
                        txtContactPhoneNumber1Ext.Text = dr("PhoneExtension").ToText
                        txtContactType1.Text = dr("Type").ToText
                        txtContactBestTime1.Text = dr("BestTime").ToText

                    Case 1
                        txtContactPhoneNumber2.Text = Common.PageFormatPhone(dr("PhoneNumber").Value)
                        txtContactPhoneNumber2Ext.Text = dr("PhoneExtension").ToText
                        txtContactType2.Text = dr("Type").ToText
                        txtContactBestTime2.Text = dr("BestTime").ToText
                    Case 2
                        txtContactPhoneNumber3.Text = Common.PageFormatPhone(dr("PhoneNumber").Value)
                        txtContactPhoneNumber3Ext.Text = dr("PhoneExtension").ToText
                        txtContactType3.Text = dr("Type").ToText
                        txtContactBestTime3.Text = dr("BestTime").ToText
                End Select
            Next

            ExistingRecordInd = 1
            If txtEmpID.Text.Length > 0 Then
                EmpIDInd = 1
            End If

        Catch ex As Exception
            Throw New Exception("Error #677: CallbackMaintain DisplayPage_DisplayExisting. CallbackID: " & cIndexSess.CallbackID.ToString & " . " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_DisplayOther(ByVal VIPInd As Integer, ByRef ClientID As String, ByRef EnrollWinCode As String, ByRef ExistingRecordInd As Integer, ByRef EmpIDInd As String)
        Dim dt As System.Data.DataTable

        Try

            cDiag.Append("DispPgOther hdSubAction: " & Request.Form("hdSubAction") & " | ")

            Select Case Request.Form("hdSubAction")
                Case "EmpIDInserted"
                    dt = Common.GetDT("SELECT FirstName, LastName, MI = IsNull(MI, ''), State, AnnualPayRate FROM " & ddClientID.SelectedValue & "..Employee WHERE EmpID = '" & Request.Form("txtEmpID") & "'")

                    txtClientID.Text = Common.DropdownGetTextFromValue(ddClientID, ddClientID.SelectedValue)
                    ddState.ClearSelection()
                    ddState.Items.FindByValue(dt.Rows(0)("State")).Selected = True
                    txtState.Text = Common.DropdownGetTextFromValue(ddState, ddState.SelectedValue)
                    txtEmpLastName.Text = dt.Rows(0)("LastName")
                    txtEmpFirstName.Text = dt.Rows(0)("FirstName")
                    txtEmpMI.Text = dt.Rows(0)("MI")
                    txtEmpID.Text = Request.Form("txtEmpID")
                    If dt.Rows(0)("AnnualPayRate") IsNot DBNull.Value AndAlso dt.Rows(0)("AnnualPayRate") > 99999 Then
                        VIPInd = 1
                    End If

                    ' txtDiagnostic.Text = "NumRows: " & dt.Rows.Count & ". DBHost: " & Enviro.DBHost & ". Sql: SELECT FirstName, LastName, MI = IsNull(MI, ''), State, AnnualPayRate FROM " & ddClientID.SelectedValue & "..Employee WHERE EmpID = '" & Request.Form("txtEmpID") & "'"

                    ClientID = ddClientID.SelectedValue
                    EnrollWinCode = Request.Form("hdEnrollWinCode")
                    EmpIDInd = 1
                    Select Case Request.Form("hdExistingRecordInd")
                        Case "0"
                            ExistingRecordInd = 0
                        Case "1"
                            ExistingRecordInd = 1
                    End Select


                Case "EmpIDClear"
                    ddState.ClearSelection()
                    ddClientID.ClearSelection()
                    txtState.Text = String.Empty
                    txtEmpLastName.Text = String.Empty
                    txtEmpFirstName.Text = String.Empty
                    txtEmpMI.Text = String.Empty
                    txtEmpID.Text = String.Empty

                    ClientID = "0"
                    EnrollWinCode = 0
                    EmpIDInd = 0

                    Select Case Request.Form("hdExistingRecordInd")
                        Case "0"
                            ExistingRecordInd = 0
                        Case "1"
                            ExistingRecordInd = 1
                    End Select

            End Select


        Catch ex As Exception
            Throw New Exception("Error #678: CallbackMaintain DisplayPage_DisplayOther. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_HandleHiddens(ByVal VIPInd As Integer, ByVal SelectedTab As String, ByVal ResponseAction As ResponseActionEnum, ByRef ClientID As String, ByRef EnrollWinCode As String, ByRef ExistingRecordInd As Integer, ByRef EmpIDInd As String)
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        Dim dt As System.Data.DataTable
        Dim AcesClients As String

        Try

            ' ___ Build Aces string
            dt = Common.GetDT("SELECT ClientID FROM Codes_ClientID WHERE AcesInd = 1", Enviro.DBHost, "ProjectReports")
            For i = 0 To dt.Rows.Count - 1
                sb.Append(dt.Rows(i)(0) & "|")
            Next
            AcesClients = sb.ToString


            ' ___ Existing record indicator
            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank, ResponseActionEnum.ClearAll, ResponseActionEnum.DisplayUserInputNew
                    ExistingRecordInd = 0
                Case ResponseActionEnum.DisplayExisting, ResponseActionEnum.DisplayUserInputExisting
                    ExistingRecordInd = 1
                Case Else
                    ExistingRecordInd = CType(Request.Form("hdExistingRecordInd"), System.Int64)
            End Select

            sb.Length = 0
            sb.Append("<input type=""hidden"" id=""hdNotifyParentThenClose"" value=""0"" />")
            sb.Append("<input type=""hidden"" id=""hdTicketNumber"" value=""" & txtTicketNumber.Text & """ />")
            sb.Append("<input type=""hidden"" id=""hdOverflowAgent"" value=""" & ddOverflowAgent.SelectedValue & """ />")
            sb.Append("<input type=""hidden"" id=""hdPriorityTag"" value=""" & ddPriorityTagInd.SelectedItem.Text & """ />")
            sb.Append("<input type=""hidden"" id=""hdAcesClients"" value=""" & AcesClients & """ />")
            sb.Append("<input type=""hidden"" id=""hdVIPInd"" value=""" & VIPInd.ToString & """ />")
            sb.Append("<input type=""hidden"" id=""hdSelectedTab"" name=""hdSelectedTab"" value=""" & SelectedTab & """/>")
            sb.Append("<input type=""hidden"" id=""hdExistingRecordInd"" name=""hdExistingRecordInd"" value=""" & ExistingRecordInd.ToString & """ />")
            sb.Append("<input type=""hidden"" id=""hdClientID"" name=""hdClientID"" value=""" & ClientID.ToString & """ />")
            sb.Append("<input type=""hidden"" id=""hdEmpIDInd"" name=""hdEmpIDInd"" value=""" & EmpIDInd.ToString & """ />")
            sb.Append("<input type=""hidden"" id=""hdEnrollWinCode"" name=""hdEnrollWinCode"" value=""" & EnrollWinCode & """ />")
            sb.Append("<input type=""hidden"" id=""hdCallbackID"" name=""hdCallbackID"" value=""" & cIndexSess.CallbackID & """ />")
            sb.Append("<input type=""hidden"" id=""hdLoggedInUserID"" name=""hdLoggedInUserID"" value=""" & cEnviro.LoggedInUserID & """ />")
            litHiddens.Text = sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #679: CallbackMaintain DisplayPage_HandleHiddens. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub DisplayPage_HandleEnrollWin(ByVal ClientID As String, ByVal EnrollWinCode As String)
        Dim AcesInd As Boolean
        Dim dt As System.Data.DataTable
        Dim WhereClause As String = Nothing

        Try


            ' // 1. Populate dropdown

            ddEnrollWin.Items.Clear()

            If ClientID = "0" Then
                ddEnrollWin.Items.Add(New ListItem("Client required", "Client required"))

            Else

                AcesInd = Common.IsAcesClient(ddClientID.SelectedValue)

                If AcesInd Then

                    ' ___ Aces client: Get codes from client Codes_EnrollWin table
                    dt = Common.GetDT("SELECT EnrollWinCode, ShortEnrollWinDesc = EnrollWinDesc FROM " & ClientID & "..Codes_EnrollWin WHERE LogicalDelete = 0 ORDER BY EnrollWinCode")

                Else

                    ' ___ Non-Aces client: Get codes from ProjectReports..Config_ClientEnrollWinCodes
                    dt = Common.GetDT("SELECT * FROM ProjectReports..Config_ClientEnrollWinCodes WHERE ClientID = '" & ddClientID.SelectedValue & "'")
                    For i = 1 To dt.Columns.Count - 1
                        If dt.Rows(0)(i) Then
                            WhereClause &= "'" & dt.Columns(i).ColumnName & "', "
                        End If
                    Next
                    If WhereClause.Length > 0 Then
                        WhereClause = WhereClause.Substring(0, WhereClause.Length - 2)
                        WhereClause = "WHERE EnrollWinCode IN (" & WhereClause & ")"
                    End If
                    dt = Common.GetDT("SELECT EnrollWinCode, ShortEnrollWinDesc FROM ProjectReports..Codes_EnrollWin " & WhereClause & " Order By OrderBy")

                End If

                ' ___ Add blank row; then databind dropdown and datatable
                dt.Rows.InsertAt(dt.NewRow, 0)
                dt.Rows(0)(0) = "0"
                dt.Rows(0)(1) = String.Empty
                ddEnrollWin.DataValueField = "EnrollWinCode"
                ddEnrollWin.DataTextField = "ShortEnrollWinDesc"

                ddEnrollWin.DataSource = dt
                ddEnrollWin.DataBind()

            End If


            ' // 2. Select value

            If ClientID <> "0" AndAlso EnrollWinCode <> "0" Then
                Try
                    ddEnrollWin.Items.FindByValue(EnrollWinCode).Selected = True
                    txtEnrollWin.Text = Common.DropdownGetTextFromValue(ddEnrollWin, EnrollWinCode)
                Catch
                End Try
            End If

        Catch ex As Exception
            Throw New Exception("Error #658: CallbackMaintain HandleEnrollWin. " & ex.Message, ex)
        End Try
    End Sub

    Private Function GetEnrollWinDesc(ByVal EnrollWinCode As String, ByVal ClientID As String) As String
        Dim dt As System.Data.DataTable
        If Common.IsAcesClient(ddClientID.SelectedValue) Then
            dt = Common.GetDT("SELECT EnrollWinDesc FROM " & ddClientID.SelectedValue & "..Codes_EnrollWin WHERE EnrollWinCode = '" & EnrollWinCode & "' AND LogicalDelete = 0")
        Else
            dt = Common.GetDT("SELECT EnrollWinDesc FROM ProjectReports..Codes_EnrollWin WHERE EnrollWinCode = '" & EnrollWinCode & "' AND LogicalDelete = 0")
        End If
        Return dt.Rows(0)(0)
    End Function

    Private Sub DisplayPage_FormatControls(ByVal ClientID As String)
        Dim sb As New System.Text.StringBuilder

        Try


            Style.AddStyle(txtDiagnostic, Style.StyleType.NoneditableGrayed, 260)


            Style.AddStyle(txtTicketNumber, Style.StyleType.NoneditableGrayed, 150)
            Style.AddStyle(txtDate, Style.StyleType.NoneditableGrayed, 180)
            Style.AddStyle(txtEmpID, Style.StyleType.NoneditableGrayed, 190)

            If cRights.HasThisRight(Rights.CallbackEdit) Then

                If txtEmpID.Text.Length = 0 Then
                    ddState.Visible = True
                    txtState.Visible = False

                    ddClientID.Visible = True
                    txtClientID.Visible = False

                    Style.AddStyle(txtEmpLastName, Style.StyleType.NormalEditable, 220)
                    Style.AddStyle(txtEmpFirstName, Style.StyleType.NormalEditable, 170)
                    Style.AddStyle(txtEmpMI, Style.StyleType.NormalEditable, 30)

                    lnkEmpIDClear.Visible = False
                    lnkLookupEmpID.Enabled = True

                Else
                    ddState.Visible = False
                    txtState.Visible = True
                    Style.AddStyle(txtState, Style.StyleType.NoneditableGrayed, 140)

                    ddClientID.Visible = False
                    txtClientID.Visible = True
                    Style.AddStyle(txtClientID, Style.StyleType.NoneditableGrayed, 180)

                    Style.AddStyle(txtEmpLastName, Style.StyleType.NoneditableGrayed, 220)
                    Style.AddStyle(txtEmpFirstName, Style.StyleType.NoneditableGrayed, 170)
                    Style.AddStyle(txtEmpMI, Style.StyleType.NoneditableGrayed, 30)

                    If Common.IsAcesClient(ClientID) AndAlso cIndexSess.CallbackID <> String.Empty Then
                        lnkEmpIDClear.Visible = False
                    Else
                        lnkEmpIDClear.Visible = True
                    End If

                    lnkLookupEmpID.Enabled = False
                End If
                Style.AddStyle(txtNumEmployeeCalls, Style.StyleType.NormalEditable, 30)


                ddOverflowAgent.Visible = True
                txtOverflowAgent.Visible = False


                ddEnrollWin.Visible = True
                txtEnrollWin.Visible = False


                ddPriorityTagInd.Visible = True
                txtPriorityTagInd.Visible = False

                ddCallPurpose.Visible = True
                txtCallPurpose.Visible = False

                ddPreferSpanishInd.Visible = True
                txtPreferSpanishInd.Visible = False

                Style.AddStyle(txtContactPhoneNumber1, Style.StyleType.NormalEditable, 100)
                Style.AddStyle(txtContactPhoneNumber1Ext, Style.StyleType.NormalEditable, 80)
                Style.AddStyle(txtContactType1, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtContactBestTime1, Style.StyleType.NormalEditable, 270)

                Style.AddStyle(txtContactPhoneNumber2, Style.StyleType.NormalEditable, 100)
                Style.AddStyle(txtContactPhoneNumber2Ext, Style.StyleType.NormalEditable, 80)
                Style.AddStyle(txtContactType2, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtContactBestTime2, Style.StyleType.NormalEditable, 270)

                Style.AddStyle(txtContactPhoneNumber3, Style.StyleType.NormalEditable, 100)
                Style.AddStyle(txtContactPhoneNumber3Ext, Style.StyleType.NormalEditable, 80)
                Style.AddStyle(txtContactType3, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtContactBestTime3, Style.StyleType.NormalEditable, 300)

                'Style.AddStyle(txtAuthPerson, Style.StyleType.NormalEditable, 260)
                'Style.AddStyle(txtAuthRelationship, Style.StyleType.NormalEditable, 280)

                Style.AddStyle(txtNotes, Style.StyleType.NormalEditable, 461, True)

                lnkSendEmail.Visible = True
                Style.AddStyle(txtEmailComments, Style.StyleType.NormalEditable, 461, True)
                litSaveCancelReturn.Visible = True


                'sb.Append("<a  href=""javascript:Save()""    style=""position:absolute;left:24px;top:392px""  class=""pairedbuttons4"">Save&nbsp;&nbsp;</a>")
                'sb.Append("<a  href=""javascript:Cancel()""  style=""position:absolute;left:136px;top:392px"" class=""pairedbuttons4"">Cancel&nbsp;&nbsp;</a>")
                'sb.Append("<a  href=""javascript:Return()""  style=""position:absolute;left:248px;top:392px"" class=""pairedbuttons4"">Return&nbsp;&nbsp;</a>")
                'litSaveCancelReturn.Text = sb.ToString

                'sb.Append(Common.BuildButton("Save", Nothing, "position:absolute;left:24px;top:392px", "Save"))
                'sb.Append(Common.BuildButton("Cancel", Nothing, "position:absolute;left:136px;top:392px", "Cancel"))
                'sb.Append(Common.BuildButton("Return", Nothing, "position:absolute;left:248px;top:392px", "Return"))

                sb.Append(Common.BuildButton("Save", Nothing, 24, 392, "Save"))
                sb.Append(Common.BuildButton("Cancel", Nothing, 136, 392, "Cancel"))
                sb.Append(Common.BuildButton("Close", Nothing, 248, 392, "Close"))

                sb.Append("<a href=""javascript:Previous()"" style=""position: absolute;left: 440px;top:388px;"" class=""previousbutton""></a>")
                sb.Append("<a href=""javascript:Next()"" style=""position: absolute;left: 480px;top:388px;"" class=""nextbutton""></a>")

                litSaveCancelReturn.Text = sb.ToString

            Else

                ddClientID.Visible = False
                txtClientID.Visible = True
                Style.AddStyle(txtClientID, Style.StyleType.NoneditableGrayed, 180)

                ddState.Visible = False
                txtState.Visible = True
                Style.AddStyle(txtState, Style.StyleType.NoneditableGrayed, 140)

                Style.AddStyle(txtEmpLastName, Style.StyleType.NoneditableGrayed, 220)
                Style.AddStyle(txtEmpFirstName, Style.StyleType.NoneditableGrayed, 190)
                Style.AddStyle(txtEmpMI, Style.StyleType.NoneditableGrayed, 30)
                lnkEmpIDClear.Visible = False
                lnkLookupEmpID.Enabled = False
                Style.AddStyle(txtNumEmployeeCalls, Style.StyleType.NoneditableGrayed, 30)

                ddOverflowAgent.Visible = False
                txtOverflowAgent.Visible = True
                Style.AddStyle(txtOverflowAgent, Style.StyleType.NoneditableGrayed, 220)

                ddEnrollWin.Visible = False
                txtEnrollWin.Visible = True
                Style.AddStyle(txtEnrollWin, Style.StyleType.NoneditableGrayed, 235)

                ddPriorityTagInd.Visible = False
                txtPriorityTagInd.Visible = True
                Style.AddStyle(txtPriorityTagInd, Style.StyleType.NoneditableGrayed, 60)

                ddCallPurpose.Visible = False
                txtCallPurpose.Visible = True
                Style.AddStyle(txtCallPurpose, Style.StyleType.NoneditableGrayed, 235)

                ddPreferSpanishInd.Visible = False
                txtPreferSpanishInd.Visible = True
                Style.AddStyle(txtPreferSpanishInd, Style.StyleType.NoneditableGrayed, 60)

                Style.AddStyle(txtContactPhoneNumber1, Style.StyleType.NoneditableGrayed, 100)
                Style.AddStyle(txtContactPhoneNumber1Ext, Style.StyleType.NoneditableGrayed, 80)
                Style.AddStyle(txtContactType1, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtContactBestTime1, Style.StyleType.NoneditableGrayed, 300)

                Style.AddStyle(txtContactPhoneNumber2, Style.StyleType.NoneditableGrayed, 100)
                Style.AddStyle(txtContactPhoneNumber2Ext, Style.StyleType.NoneditableGrayed, 80)
                Style.AddStyle(txtContactType2, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtContactBestTime2, Style.StyleType.NoneditableGrayed, 300)

                Style.AddStyle(txtContactPhoneNumber3, Style.StyleType.NoneditableGrayed, 100)
                Style.AddStyle(txtContactPhoneNumber2Ext, Style.StyleType.NoneditableGrayed, 80)
                Style.AddStyle(txtContactType3, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtContactBestTime3, Style.StyleType.NoneditableGrayed, 300)

                'Style.AddStyle(txtAuthPerson, Style.StyleType.NoneditableGrayed, 300)
                'Style.AddStyle(txtAuthRelationship, Style.StyleType.NoneditableGrayed, 300)

                Style.AddStyle(txtNotes, Style.StyleType.NoneditableGrayed, 300, True)

                lnkSendEmail.Visible = False
                Style.AddStyle(txtEmailComments, Style.StyleType.NoneditableGrayed, 300, True)

                sb.Append(Common.BuildButton("Return", Nothing, 24, 392, "Close"))
                litSaveCancelReturn.Text = sb.ToString
            End If

        Catch ex As Exception
            Throw New Exception("Error #662: CallbackMaintain DisplayPage_FormatControls. " & ex.Message, ex)
        End Try
    End Sub

#Region " Ajax calls "
    <System.Web.Services.WebMethod()>
    Public Shared Function EscalateCallback(ByVal TicketNumber As String, ByVal OverflowAgent As String, ByVal DesignatedEnroller As String, ByVal PriorityTag As String, ByVal EmailComments As String) As String
        Dim DesignatedDateTime As String
        Dim dt As DataTable
        Dim OverflowAgentText As String
        Dim DesignatedEnrollerText As String
        Dim ClientID As String
        Dim EmpFullName As String
        Dim sb As New System.Text.StringBuilder

        ' SELECT ClientID, EmpFullNam= RTrim(EmpLastName + ', ' + EmpFirstName + ' ' + EmpMI) FROM CallbackMaster WHERE TicketNumber = 70988

        'dt = Common.GetDT("SELECT ClientID FROM CallbackMaster WHERE TicketNumber = " & Replace(TicketNumber, "-", ""))
        dt = Common.GetDT("SELECT ClientID, EmpFullNam = RTrim(EmpLastName + ', ' + EmpFirstName + ' ' + EmpMI) FROM CallbackMaster WHERE TicketNumber = " & Replace(TicketNumber, "-", ""))

        ClientID = dt.Rows(0)(0)
        EmpFullName = dt.Rows(0)(1)
        dt = Common.GetDT("SELECT FirstName + ' ' + LastName From UserManagement..Users WHERE UserID = '" & OverflowAgent & "'")
        OverflowAgentText = dt.Rows(0)(0)
        dt = Common.GetDT("SELECT FirstName + ' ' + LastName From UserManagement..Users WHERE UserID = '" & DesignatedEnroller & "'")
        DesignatedEnrollerText = dt.Rows(0)(0)
        DesignatedDateTime = Common.GetServerDateTime.ToString("MM/dd/yyyy hh:mm tt")


        sb.Append("You've been designated as callback agent." & Chr(10) & Chr(10))
        sb.Append("Callback ticket: " & TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(3) & Chr(10))
        sb.Append("Time of notice: " & DesignatedDateTime & Chr(10))
        sb.Append("Client: " & ClientID & Chr(10))
        sb.Append("Employee: " & EmpFullName & Chr(10))
        sb.Append("Overflow agent: " & OverflowAgentText & Chr(10))
        sb.Append("Designated enroller: " & DesignatedEnrollerText & Chr(10))
        sb.Append("Priority tag: " & PriorityTag & Chr(10))

        If EmailComments.Length > 0 Then
            sb.Append(Chr(10) & Chr(10) & EmailComments)
        End If

        'TextBody = "Overflow agent : & OverflowAgentText & chr(10) & DesignatedEnroller: " & DesignatedEnrollerText & ", to handle the callback for Ticket# " & TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(3) & " as of " & DesignatedDateTime & " Eastern Standard Time. Priority tag: " & PriorityTag & Chr(10) & Chr(10) & EmailComments
        Common.SendEmail(DesignatedEnroller & "@benefitvision.com", OverflowAgent & "@benefitvision.com", "HBG_SUPERVISORS@benefitvision.com", "Notice of callback escalation (" & Enviro.DBHost & ")", sb.ToString)

        Return DesignatedDateTime
    End Function

    <System.Web.Services.WebMethod()>
    Public Shared Function EscalateCallback2(ByVal TicketNumber As String, ByVal OverflowAgent As String, ByVal DesignatedEnroller As String, ByVal PriorityTag As String, ByVal EmailComments As String) As String
        Dim DesignatedDateTime As String
        Dim dt As DataTable
        Dim OverflowAgentText As String
        Dim DesignatedEnrollerText As String
        Dim sb As New System.Text.StringBuilder

        dt = Common.GetDT("SELECT FirstName + ' ' + LastName From UserManagement..Users WHERE UserID = '" & OverflowAgent & "'")
        OverflowAgentText = dt.Rows(0)(0)
        dt = Common.GetDT("SELECT FirstName + ' ' + LastName From UserManagement..Users WHERE UserID = '" & DesignatedEnroller & "'")
        DesignatedEnrollerText = dt.Rows(0)(0)
        DesignatedDateTime = Common.GetServerDateTime.ToString("MM/dd/yyyy hh:mm tt")

        sb.Append("You've been designated as callback agent." & Chr(10) & Chr(10))
        sb.Append("Callback ticket: " & TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(3) & Chr(10))
        sb.Append("Time of notice: " & DesignatedDateTime & Chr(10))
        sb.Append("Overflow agent: " & OverflowAgentText & Chr(10))
        sb.Append("Designated enroller: " & DesignatedEnrollerText & Chr(10))
        sb.Append("Priority tag: " & PriorityTag & Chr(10))

        If EmailComments.Length > 0 Then
            sb.Append(Chr(10) & Chr(10) & EmailComments)
        End If

        'TextBody = "Overflow agent : & OverflowAgentText & chr(10) & DesignatedEnroller: " & DesignatedEnrollerText & ", to handle the callback for Ticket# " & TicketNumber.Substring(0, 2) & "-" & TicketNumber.Substring(3) & " as of " & DesignatedDateTime & " Eastern Standard Time. Priority tag: " & PriorityTag & Chr(10) & Chr(10) & EmailComments
        Common.SendEmail(DesignatedEnroller & "@benefitvision.com", OverflowAgent & "@benefitvision.com", "HBG_SUPERVISORS@benefitvision.com", "Notice of callback escalation (" & Enviro.DBHost & ")", sb.ToString)

        Return DesignatedDateTime
    End Function

    <System.Web.Services.WebMethod()>
    Public Shared Function RefreshCheckout(ByVal LoggedInUserID As String, ByVal CallbackID As Int64) As String
        Return Common.RefreshCheckout(LoggedInUserID, CallbackID)
    End Function
#End Region

#Region " Helpers "
    Private Function GetStatus(ByVal StatusCode As String) As String
        Dim dt As New System.Data.DataTable
        dt = Common.GetDT("SELECT StatusDescription FROM Codes_Status WHERE StatusCode = '" & StatusCode & "'")
        Return "Status: " & dt.Rows(0)(0)
    End Function
#End Region
End Class