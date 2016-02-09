
Partial Class EmpIDLookup
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DG As DG
        Dim Common As New Common
        Dim sb As New System.Text.StringBuilder
        Dim dt As System.Data.DataTable
        Dim MainClause As New System.Text.StringBuilder
        Dim EmbeddedMessage As String = Nothing

        Try

            ' ___ Datagrid
            DG = New DG("EmpID", Common, Nothing, True, "EmbeddedTableDef", "Name", DG.DefaultSortDirectionEnum.Ascending)
            DG.AddLinkColumn("Name", "Name", "javascript:EmpIDClick", "Name", Nothing, True, Nothing, Nothing, "align='left'", Nothing, Nothing, Nothing)
            DG.AddDataBoundColumn("State", "State", "State", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("EmpID", "EmpID", "Emp ID", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("Status", "EmployeeStatusDesc", "Status", Nothing, True, Nothing, Nothing, "align='left'")

            ' ___ Main clause
            MainClause.Append("FROM " & Request.QueryString("ClientID") & "..Employee e ")
            MainClause.Append("LEFT JOIN " & Request.QueryString("ClientID") & "..Codes_EmployeeStatus sc ")
            MainClause.Append("ON e.StatusCode = sc.EmployeeStatusCode ")

            MainClause.Append("LEFT JOIN " & Request.QueryString("ClientID") & "..TestIDs tid ")
            MainClause.Append("ON e.EmpID = tid.EmpID ")


            MainClause.Append("WHERE tid.EmpID IS NULL AND ")
            MainClause.Append("e.LastName LIKE '" & Replace(Request.QueryString("EmpLastName"), "'", "''") & "%' ")
            If Request.QueryString("EmpFirstName").Length > 0 Then
                MainClause.Append("AND e.FirstName LIKE '" & Replace(Request.QueryString("EmpFirstName"), "'", "''") & "%' ")
            End If
            If Request.QueryString("EmpMI").Length > 0 Then
                MainClause.Append("AND e.MI = '" & Request.QueryString("EmpMI") & "' ")
            End If
            If Request.QueryString("State") <> "0" Then
                MainClause.Append("AND e.State = '" & Request.QueryString("State") & "' ")
            End If

            ' ____ Recordcount
            sb.Append("SELECT Count (*) ")
            sb.Append(MainClause.ToString)
            dt = Common.GetDT(sb.ToString)

            If dt.Rows(0)(0) > 300 Then
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & dt.Rows(0)(0).ToString & " records. Please narrow search.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
                dt = Nothing
            Else
                sb.Length = 0
                'sb.Append("SELECT Name = case ")
                'sb.Append("when tid.EmpID IS NULL then e.LastName + ', ' + e.FirstName + ' ' + e.MI ")
                'sb.Append("else e.LastName + ', ' + e.FirstName + ' ' + e.MI + ' (test)'")
                'sb.Append("end, ")

                sb.Append("SELECT Name = rtrim(e.LastName + ', ' + e.FirstName + ' ' + IsNull(e.MI, '')), ")
                sb.Append("e.State, e.EmpID, sc.EmployeeStatusDesc ")
                sb.Append(MainClause.ToString)
                sb.Append("ORDER BY CAST(sc.OrderBy as varchar(2)), lower(e.LastName), lower(e.FirstName), lower(IsNull(e.MI, '')) ")
                dt = Common.GetDTExtended(sb.ToString)
            End If
            litDG.Text = DG.GetHTML(dt, Nothing, EmbeddedMessage)
            litHiddens.Text = "<input type=""hidden"" id=""hdClientID"" value=""" & Request.QueryString("ClientID") & """ />"

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj("Error #902: EmpIDLookup Page_Load. " & ex.Message)
        End Try
    End Sub


    <System.Web.Services.WebMethod()>
    Public Shared Function EmpIDClick(ByVal EmpID As String, ByVal ClientID As String) As String
        Dim Sql As String
        Dim Querypack As DBase.QueryPack
        Dim Results As New System.Text.StringBuilder
        Dim ErrorInd As String
        Dim ErrorMessage As String = Nothing
        Dim TicketExistsInd As String = Nothing

        ' 0: ErrorInd
        ' 1: ErrorMessage
        ' 2: EmpID
        ' 3: TicketExistsInd

        'Sql = "SELECT Count (*) FROM CallbackMaster WHERE EmpID = '" & EmpID & "' AND LogicalDelete = 0"
        Sql = "SELECT Count (*) FROM CallbackMaster WHERE ClientID = '" & ClientID & "' AND EmpID = '" & EmpID & "' AND LogicalDelete = 0 AND ISNULL(IsPurged, 0) = 0"


        Querypack = Common.GetDTWithQueryPack(Sql)

        If Querypack.Success Then
            ErrorInd = "0"
            If Querypack.dt.Rows(0)(0) = 0 Then
                TicketExistsInd = "0"
            Else
                TicketExistsInd = "1"
            End If
        Else
            ErrorInd = "1"
            ErrorMessage = Common.ToJSAlert(Querypack.TechErrMsg)
        End If

        Results.Append(ErrorInd & "|")
        Results.Append(ErrorMessage & "|")
        Results.Append(EmpID & "|")
        Results.Append(TicketExistsInd)

        Return Results.ToString
    End Function
End Class
