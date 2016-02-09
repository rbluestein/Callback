Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

#Region " Enums "
Public Enum PageMode
    Initial = 1
    Postback = 2
    ReturnFromChild = 3
    CalledByOther = 4
End Enum

Public Enum RequestActionEnum
    CreateNew = 1
    LoadExisting = 2
    SaveNew = 3
    SaveExisting = 4
    Other = 5
    ChangeTab = 6
    CancelChanges = 7
    ReturnToCallingPage = 8
End Enum

Public Enum ResponseActionEnum
    DisplayBlank = 1
    DisplayUserInputNew = 2
    DisplayUserInputExisting = 3
    DisplayExisting = 4
    ChangeTab = 5
    Other = 6
    ReturnToCallingPageDirectly = 7
    ReturnToCallingPageAfterRender = 8
    ClearAll = 9
End Enum
Public Enum BehaviorOnSuccessfulSaveEnum
    DisplayPage = 1
    ReturnToCallerDirectly = 2
    ReturnToCallerAfterRender = 3
End Enum
Public Enum BehaviorOnCancelButtonClickEnum
    DisplayPage = 1
    ReturnToCallerDirectly = 2
    ReturnToCallerAfterRender = 3
End Enum
Public Enum StringTreatEnum
    AsIs = 1
    SideQts = 2
    SecApost = 3
    SideQts_SecApost = 4
End Enum
Public Enum FormatDateEnum
    SpecialFormat1 = 1
    SpecialFormat2 = 2
    SpecialFormat3 = 3
End Enum
Public Enum RoleCatgyEnum
    Other = 1
    Supervisor = 2
    Enroller = 3
End Enum
Public Enum DataTypeEnum
    [String] = 1
    [Integer] = 2
    [Decimal] = 3
    [Boolean] = 4
    [Date] = 5
End Enum
#End Region

Public Class Common
#Region " Declarations "
    Private cEnviro As Enviro
    Private cDBase As DBase
#End Region

    Public Shared Function GetBootsAPuppy(ByVal Text As String) As String
        'only allow letters and spaces and numerals and decimal place
        'and # and @ and . and -
        Dim i As Integer
        Dim Result As String = Nothing
        Dim ThisChar As Char
        Dim ThisAsc As Integer

        Try
            If IsBlank(Text) Then
                Result = String.Empty
            Else
                'Text = Convert.ToString(Text)
                For i = 0 To Text.Length - 1
                    ThisChar = Text.Substring(i, 1)
                    ThisAsc = Asc(ThisChar)
                    'If (Asc(ThisChar) >= 65 And Asc(ThisChar) <= 90) Or (Asc(ThisChar) >= 97 And Asc(ThisChar) <= 122) Or Asc(ThisChar) = 32 Or Asc(ThisChar) = 45 Or Asc(ThisChar) = 35 Or Asc(ThisChar) = 46 Or Asc(ThisChar) = 64 Then
                    '    Result = Result & ThisChar
                    'End If
                    If (ThisAsc = 39) Or (ThisAsc >= 65 AndAlso ThisAsc <= 90) Or (ThisAsc >= 97 And ThisAsc <= 122) Or (ThisAsc >= 48 AndAlso ThisAsc <= 57) Or ThisAsc = 32 Or ThisAsc = 45 Or ThisAsc = 35 Or ThisAsc = 46 Or ThisAsc = 64 Then
                        Result = Result & ThisChar
                    End If

                Next
            End If

            Return Result

        Catch ex As Exception
            Throw New Exception("Error #2298. Common.GetBootsAPuppy. " & ex.Message)
        End Try
    End Function

    Public Shared Function FormatDate(ByVal Input As DateTime, ByVal Format As FormatDateEnum) As String
        Dim Results As String = Nothing
        Select Case Format
            Case FormatDateEnum.SpecialFormat1
                Results = Input.ToString("MMM d yyyy<br />h:mm") & Input.ToString("tt").ToLower
            Case FormatDateEnum.SpecialFormat2
                Results = Input.ToString("MMM d yyyy h:mm") & Input.ToString("tt").ToLower
            Case FormatDateEnum.SpecialFormat3
                Results = Input.ToString("M/dd/yyyy<br>h:mm") & Input.ToString("tt").ToLower
        End Select
        Return Results
    End Function

    Public Shared Function RefreshCheckout(ByVal LoggedInUserID As String, ByVal CallbackID As Int64) As String
        Dim Sql As New System.Text.StringBuilder
        Dim Querypack As DBase.QueryPack
        Dim Results As New System.Text.StringBuilder
        Dim ErrorInd As String
        Dim ErrorMessage As String = Nothing

        ' 0: CallbackID
        ' 1: ErrorInd
        ' 2: ErrorMessage
        ' 3: PageCode

        Sql.Append("UPDATE CallbackMaster SET ")
        Sql.Append("CheckoutUserID = '" & LoggedInUserID & "',  ")
        Sql.Append("CheckoutTime = '" & Common.GetServerDateTime & "', ")
        Sql.Append("SuppressLogWriteInd = 1 ")
        Sql.Append("WHERE CallbackID = " & CallbackID)
        Querypack = ExecuteNonQueryWithQuerypack(Sql.ToString)

        If Querypack.Success Then
            ErrorInd = "0"
        Else
            ErrorInd = "1"
            ErrorMessage = ToJSAlert(Querypack.TechErrMsg)
        End If

        Results.Append(CallbackID.ToString & "|")
        Results.Append(ErrorInd & "|")
        Results.Append(ErrorMessage)

        Return Results.ToString
    End Function

    'Public Shared Function GetCheckoutAgent(ByVal CallbackID As Int64, ByVal PageCode As String) As String
    '    Dim Sql As New System.Text.StringBuilder
    '    Dim Querypack As DBase.QueryPack
    '    Dim dr As DataRow
    '    Dim Results As New System.Text.StringBuilder
    '    Dim IsCheckedOutInd As String = String.Empty
    '    Dim CheckoutAgent As String = String.Empty
    '    Dim ErrorInd As String
    '    Dim ErrorMessage As String
    '    Dim TimeSpan As TimeSpan

    '    ' 0: CallbackID
    '    ' 1: Is checked out
    '    ' 2: Checkout agent
    '    ' 3: ErrorInd
    '    ' 4: Error message
    '    ' 5: PageCode

    '    ' this is good
    '    'Sql.Append("SELECT NumSecondsElapsed = DateDiff(ss ,  DateAdd(hh, -5, getutcdate()) , CheckoutTime, ")

    '    Sql.Append("SELECT CheckoutTime, ")
    '    Sql.Append("CheckoutUserID, ")
    '    Sql.Append("CheckoutAgent = u.FirstName + ' ' + u.LastName ")
    '    Sql.Append("FROM CallbackMaster mast ")
    '    Sql.Append("INNER JOIN UserManagement..Users u ON mast.CheckoutUserID = u.UserID ")
    '    Sql.Append("WHERE mast.CallbackID = " & CallbackID & " ")
    '    Querypack = GetDTWithQueryPack(Sql.ToString)

    '    ' ___ Error condition
    '    If Querypack.Success Then
    '        ErrorInd = "0"
    '        ErrorMessage = String.Empty
    '    Else
    '        ErrorInd = "1"
    '        ErrorMessage = ToJSAlert(Querypack.TechErrMsg)
    '    End If

    '    ' ___ Never been checked out
    '    If Querypack.dt.Rows.Count = 0 Then
    '        IsCheckedOutInd = "0"
    '    Else

    '        dr = Querypack.dt.Rows(0)

    '        ' ___ Last checkout write less than 45 seconds ago?
    '        'TimeSpan = dr("CheckoutTime").Subtract(GetServerDateTime())
    '        TimeSpan = GetServerDateTime().Subtract(dr("CheckoutTime"))
    '        If TimeSpan.TotalSeconds < 46 Then
    '            IsCheckedOutInd = "1"
    '            CheckoutAgent = dr("CheckoutAgent")
    '        Else
    '            IsCheckedOutInd = "0"
    '        End If
    '    End If

    '    Results.Append(CallbackID.ToString & "|")
    '    Results.Append(IsCheckedOutInd & "|")
    '    Results.Append(CheckoutAgent & "|")
    '    Results.Append(ErrorInd & "|")
    '    Results.Append(ErrorMessage & "|")
    '    Results.Append(PageCode)

    '    Return Results.ToString
    'End Function


    Public Shared Function GetTicketNumberSyntax(ByVal Prefix As String) As String
        If Prefix = Nothing Then
            Prefix = String.Empty
        End If
        If Prefix.Length > 0 Then
            Prefix &= "."
        End If
        Return " Substring(Cast(" & Prefix & "TicketNumber as varchar(6)), 1, 2) + '-' + Substring(Cast(" & Prefix & "TicketNumber as varchar(6)), 3, 4) "
    End Function

    Public Shared Function BuildButton(ByVal MethodName As String, ByVal Args As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Label As String) As String
        If Args = Nothing Then
            Args = String.Empty
        End If
        Return "<a href=""javascript:" & MethodName & "(" & IIf(Args.Length > 0, Args, "") & ")""  style=""position: absolute; left:" & Left.ToString & "px; top: " & Top.ToString & "px;""  class=""pairedbuttons4"">" & Label & "&nbsp;&nbsp;</a>"
    End Function
    Public Shared Function BuildButton(ByVal MethodName As String, ByVal Args As String, ByVal Style As String, ByVal Label As String) As String
        If Args = Nothing Then
            Args = String.Empty
        End If

        '   Return "<a href=""javascript:" & MethodName & "(" & IIf(Args.Length > 0, Args, "") & ")""  style=""" & Style & """ class=""pairedbuttons4"">" & Label & "&nbsp;&nbsp;</a>"

        If Style = Nothing Then
            Return "<a href=""javascript:" & MethodName & "(" & IIf(Args.Length > 0, Args, "") & ")""  class=""pairedbuttons4"">" & Label & "&nbsp;&nbsp;</a>"
        Else
            Return "<a href=""javascript:" & MethodName & "(" & IIf(Args.Length > 0, Args, "") & ")""  style=""" & Style & """ class=""pairedbuttons4"">" & Label & "&nbsp;&nbsp;</a>"
        End If

    End Function

#Region " Constructors "
    Public Sub New()
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        cEnviro = SessionObj("Enviro")
        cDBase = New DBase
    End Sub

    Public Sub New(ByVal Enviro As Enviro)
        cEnviro = Enviro
        cDBase = New DBase
    End Sub
#End Region

#Region " String manipulation "
    Public Shared Function Left(ByVal Value As Object, ByVal Length As Integer) As String
        Dim Results As String = String.Empty

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            If Value.length <= Length Then
                Return Value.ToString
            Else
                Return Value.Substring(0, Length)
            End If
        End If
    End Function

    Public Shared Function ToProper(ByVal Value As Object) As String
        Dim i As Integer
        Dim Box As Object
        Dim Results As String = String.Empty
        Dim NewValue As String

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            Box = Split(Value, " ")
            For i = 0 To Box.GetUpperBound(0)
                ' NewValue = Box(i).Substring(0, 1).ToUpper
                NewValue = Box(i).Substring(0, 1).ToUpper
                If Box(i).Length > 1 Then
                    NewValue &= Box(i).Substring(1).ToLower
                End If
                Results &= " " & NewValue
            Next
            Results = Trim(Results)
            Return Results
        End If
    End Function

    Public Shared Function Right(ByVal Str As String, ByVal Len As Integer) As String
        Return Str.Substring(Str.Length - Len)
    End Function

    Public Shared Function InsertAt(ByVal Value As String, ByVal InsChar As String, ByVal Pos As Integer) As String
        Dim ValuePos As Integer = 1
        Dim Output As String = String.Empty
        Dim OutputPos As Integer = 1
        Do
            If OutputPos = Pos Then
                Output &= InsChar
                OutputPos += 1
            Else
                Output &= Value.Substring(ValuePos - 1, 1)
                ValuePos += 1
                OutputPos += 1
                If ValuePos > Value.Length Then
                    Exit Do
                End If
            End If
        Loop
        Return Output
    End Function

    Public Shared Function ToJSAlert(ByVal Value As String) As String
        If Value = Nothing Then
            Value = String.Empty
        End If
        If Value.Length > 0 Then
            Value = Replace(Value, "'", "")
            Value = Replace(Value, "(", "")
            Value = Replace(Value, ")", "")
            Value = Replace(Value, """", "")
            Value = Replace(Value, vbCrLf, " ")
            Value = Replace(Value, Chr(10), " ")
            Value = Replace(Value, Chr(13), " ")
        End If
        Return Value
    End Function

    Public Shared Function SBRight(ByRef sb As System.Text.StringBuilder, ByVal Length As Integer) As String
        Return Right(sb.ToString, Length)
    End Function

    Public Shared Function PadDecimal(ByVal Input As Object, ByVal Precision As Integer) As String
        Dim DecimalPlace As Integer
        Dim InputStr As String
        Dim Box As String()

        If IsDBNull(Input) Then
            Input = PadDecimal(0, Precision)
        End If

        If Not IsNumeric(Input) Then
            If Input = String.Empty Then
                Input = 0
            Else
                Return Input
            End If
        End If

        InputStr = CType(Input, System.String)
        If InStr(InputStr, ".") = 0 Then
            InputStr &= "."
        End If
        DecimalPlace = InStr(InputStr, ".")
        Box = Split(InputStr, ".")
        If Box(1).Length > Precision Then
            Box(1) = Box(1).Substring(0, Precision)
        ElseIf Box(1).Length < Precision Then
            Box(1) = Box(1).PadRight(Precision, "0")
        End If

        If Box(0) = "0" Then
            Box(0) = String.Empty
        End If

        Return Box(0) & "." & Box(1)

    End Function
#End Region

#Region " Data: General "
    Public Function DoesTableExist(ByVal TableName As String) As Boolean
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable

        Sql.Append("USE " & Enviro.DefaultDatabase & " ")
        Sql.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        Sql.Append(TableName)
        Sql.Append("') ")
        Sql.Append("SELECT Results = 1 ")
        Sql.Append("ELSE ")
        Sql.Append("SELECT Results = 0")

        dt = GetDT(Sql.ToString, Enviro.DBHost, Enviro.DefaultDatabase)
        If dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DoesTableExist(ByVal DBHost As String, ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable

        Sql.Append("USE " & DBName & " ")
        Sql.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        Sql.Append(TableName)
        Sql.Append("') ")
        Sql.Append("SELECT Results = 1 ")
        Sql.Append("ELSE ")
        Sql.Append("SELECT Results = 0")

        dt = GetDT(Sql.ToString, DBHost, DBName)
        If dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DoesDatabaseAndTableExist(ByVal DBName As String, ByVal TableName As String) As Boolean
        Return DoesDatabaseAndTableExist(Enviro.DBHost, DBName, TableName)
    End Function

    Public Function DoesDatabaseAndTableExist(ByVal DBHost As String, ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim Sql As String
        Dim Querypack As DBase.QueryPack

        Sql = DoesTableExistSql(DBName, TableName)
        Querypack = GetDTWithQueryPack(Sql, DBHost, DBName)
        If Querypack.Success Then
            If Querypack.dt.Rows(0)(0) = 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

        'Sql = DoesTableExistSql(DBName, TableName)
        'Dim CmdAsst As New CmdAsst(cAppSession, CommandType.Text, Sql)
        'Dim QueryPack As DBase.QueryPack
        'QueryPack = CmdAsst.Execute
        'If QueryPack.Success Then
        '    If QueryPack.dt.rows(0)(0) = 1 Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'Else
        '    Return False
        'End If
    End Function

    Public Shared Function DoesTableExistSql(ByVal DBName As String, ByVal TableName As String) As String
        Dim Sql As New System.Text.StringBuilder

        Sql.Append("USE " & DBName & " ")
        Sql.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        Sql.Append(TableName)
        Sql.Append("') ")
        Sql.Append("SELECT Results = 1 ")
        Sql.Append("ELSE ")
        Sql.Append("SELECT Results = 0")
        Return Sql.ToString
    End Function

    Public Function GetFieldList(ByVal TableName As String) As DataTable
        Return GetFieldList(Enviro.DBHost, Enviro.DefaultDatabase, TableName)
    End Function

    Public Function GetFieldList(ByVal DBHost As String, ByVal Database As String, ByVal TableName As String) As DataTable
        Dim Sql As New System.Text.StringBuilder

        'Sql.append("SELECT * from information_schema.tables where table_type='BASE TABLE'")

        Sql.Append("SELECT column_name,ordinal_position,column_default,data_type, ")
        Sql.Append("Is_nullable from information_schema.columns ")
        Sql.Append("WHERE table_name='" & TableName & "'")
        Return GetDT(Sql.ToString, DBHost, Database)
    End Function

    Public Function ConvertToExtendedTable(ByRef dt As DataTable) As DataTable
        Return DBase.GetDTExtended(dt)
    End Function

    'Public Function GetDT(ByVal Sql As String, ByVal ExtendedTbl As Boolean) As DataTable
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString)
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    DataAdapter.Fill(dt)
    '    If ExtendedTbl Then
    '        dt = cDBase.GetDTExtended(dt)
    '    End If
    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return dt
    'End Function

    'Public Function GetDT(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal ExtendedTbl As Boolean) As DataTable
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    DataAdapter.Fill(dt)
    '    If ExtendedTbl Then
    '        dt = cDBase.GetDTExtended(dt)
    '    End If
    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return dt
    'End Function

    'Public Function GetDTWithQueryPack(ByVal Sql As String, ByVal ExtendedTbl As Boolean) As DBase.QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim QueryPack As New DBase.QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString())
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    Try
    '        DataAdapter.Fill(dt)
    '        If ExtendedTbl Then
    '            dt = cDBase.GetDTExtended(dt)
    '        End If
    '        QueryPack.Success = True
    '        QueryPack.dt = dt
    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function

    'Public Function GetDTWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal ExtendedTbl As Boolean) As DBase.QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim QueryPack As New DBase.QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    Try
    '        DataAdapter.Fill(dt)
    '        If ExtendedTbl Then
    '            dt = cDBase.GetDTExtended(dt)
    '        End If
    '        QueryPack.Success = True
    '        QueryPack.dt = dt
    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function

#End Region

#Region " Data: Query "
    Public Shared Sub ExecuteNonQuery(ByVal Sql As String)
        ExecuteNonQueryMaster(Sql, Enviro.DBHost, Enviro.DefaultDatabase, False)
    End Sub

    Public Shared Sub ExecuteNonQuery(ByVal DBHost As String, ByVal Database As String, ByVal Sql As String)
        ExecuteNonQueryMaster(Sql, DBHost, Database, False)
    End Sub

    Public Shared Function ExecuteNonQueryWithQuerypack(ByVal Sql As String) As DBase.QueryPack
        Dim Querypack As DBase.QueryPack
        Querypack = ExecuteNonQueryMaster(Sql, Enviro.DBHost, Enviro.DefaultDatabase, True)
        Return Querypack
    End Function

    Public Shared Function ExecuteNonQueryWithQuerypack(ByVal DBHost As String, ByVal Database As String, ByVal Sql As String) As DBase.QueryPack
        Dim Querypack As New DBase.QueryPack
        Querypack = ExecuteNonQueryMaster(Sql, DBHost, Database, True)
        Return Querypack
    End Function

    Public Shared Function ExecuteNonQueryMaster(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal WithQuerypack As Boolean) As DBase.QueryPack
        Dim Querypack As New DBase.QueryPack

        Try
            Dim SqlConnection1 As New SqlClient.SqlConnection(Enviro.GetConnectionString(DBHost, Database))
            SqlConnection1.Open()
            Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
            SqlCmd.CommandType = System.Data.CommandType.Text
            SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
            SqlConnection1.Close()
            Querypack.Success = True
        Catch ex As Exception

            If Not WithQuerypack Then
                Throw New Exception("Error #CM2203: Common ExecuteNonQueryMaster.~Sql: " & Sql & "~DBHost: " & DBHost & "~Database: " & Database & "~Error message: " & ex.Message)
            End If

            Querypack.Success = False
            Querypack.TechErrMsg = ex.Message
        End Try
        Return Querypack
    End Function

    ' ___ GetDT
    Public Shared Function GetDT(ByVal Sql As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, Enviro.DBHost, Enviro.DefaultDatabase, False, False)
        Return QueryPack.dt
    End Function

    Public Shared Function GetDT(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, False, False)
        Return QueryPack.dt
    End Function

    Public Shared Function GetDTWithQueryPack(ByVal Sql As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, Enviro.DBHost, Enviro.DefaultDatabase, False, True)
        Return QueryPack
    End Function

    Public Shared Function GetDTWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, False, True)
        Return QueryPack
    End Function

    ' ___ GetDTExtended
    Public Shared Function GetDTExtended(ByVal Sql As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, Enviro.DBHost, Enviro.DefaultDatabase, True, False)
        Return QueryPack.dt
    End Function

    Public Shared Function GetDTExtended(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, True, False)
        Return QueryPack.dt
    End Function

    Public Shared Function GetDTExtendedWithQueryPack(ByVal Sql As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, Enviro.DBHost, Enviro.DefaultDatabase, True, True)
        Return QueryPack
    End Function

    Public Shared Function GetDTExtendedWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, True, True)
        Return QueryPack
    End Function

    Public Shared Function GetDTMaster(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal ExtendedTbl As Boolean, ByVal WithQuerypack As Boolean) As DBase.QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim QueryPack As New DBase.QueryPack

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.CommandTimeout = 90
        SqlCmd.Connection = New SqlConnection(Enviro.GetConnectionString(DBHost, Database))
        DataAdapter = New SqlDataAdapter(SqlCmd)
        Try
            DataAdapter.Fill(dt)
            If ExtendedTbl Then
                dt = DBase.GetDTExtended(dt)
            End If
            QueryPack.Success = True
            QueryPack.dt = dt
        Catch ex As Exception

            If Not WithQuerypack Then
                Throw New Exception("Error #CM2202: Common GetDTMaster.~Sql: " & Sql & "~DBHost: " & DBHost & "~Database: " & Database & "~Error message: " & ex.Message)
            End If

            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
        End Try

        DataAdapter.Dispose()
        SqlCmd.Dispose()
        SqlCmd.Connection.Close()
        Return QueryPack
    End Function

    'Public Function GetDTExtendedWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal ExtendedTbl As Boolean) As DBase.QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim QueryPack As New DBase.QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    Try
    '        DataAdapter.Fill(dt)
    '        If ExtendedTbl Then
    '            dt = cDBase.GetDTExtended(dt)
    '        End If
    '        QueryPack.Success = True
    '        QueryPack.dt = dt
    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function
#End Region

#Region " Page/Controls handling "
    Public Shared Function GetPageControls(ByRef ControlCollection As ControlCollection) As ArrayList
        Dim ControlList As New ArrayList

        GetPageControls2(ControlCollection, ControlList)
        Return ControlList

        'For i = 0 To ControlList.Count - 1
        '    System.Diagnostics.Debug.WriteLine(ControlList.Count & "  " & ControlList.Item(i).ID)
        'Next
    End Function

    'Private Function GetPageControls() As ArrayList
    '    Dim i As Integer
    '    Dim ControlList As New ArrayList

    '    GetPageControls2(Page.Controls, ControlList)
    '    Return ControlList

    '    'For i = 0 To ControlList.Count - 1
    '    '    System.Diagnostics.Debug.WriteLine(ControlList.Count & "  " & ControlList.Item(i).ID)
    '    'Next
    'End Function
    Private Shared Sub GetPageControls2(ByRef ControlCollection As ControlCollection, ByRef ControlList As ArrayList)
        Dim ctl As System.Web.UI.Control

        For Each ctl In ControlCollection
            If ctl.ID <> Nothing Then
                ControlList.Add(ctl)
            End If
            '   System.Diagnostics.Debug.WriteLine(ControlList.Count & "  " & ctl.ID)
            If ctl.HasControls Then
                GetPageControls2(ctl.Controls, ControlList)
            End If
        Next
    End Sub

    Public Shared Function GetPageMode(ByVal Page As Page, ByVal Sess As PageSession) As PageMode
        Dim PageMode As PageMode

        If Page.IsPostBack AndAlso Page.Request.Form("__EVENTTARGET") = "" Then
            PageMode = PageMode.Postback
        Else
            Select Case Page.Request.QueryString("CalledBy")
                Case "Child"
                    PageMode = PageMode.ReturnFromChild
                Case "Other"
                    If Sess.PageInitiallyLoaded Then
                        PageMode = PageMode.CalledByOther
                    Else
                        PageMode = PageMode.Initial
                        Sess.PageInitiallyLoaded = True
                    End If
                Case Else
                    PageMode = PageMode.Initial
                    Sess.PageInitiallyLoaded = True
            End Select
        End If
        Return PageMode
    End Function

    Public Shared Function GetRequestAction(ByVal Page As Page, ByRef mDiag As System.Text.StringBuilder) As RequestActionEnum
        If Page.IsPostBack Then
            mDiag.Append("GetReqAct hdAction: " & Page.Request.Form("hdAction") & " | ")
        End If
        Return GetRequestAction(Page)
    End Function

    Public Shared Function GetRequestAction(ByVal Page As Page) As RequestActionEnum
        Dim RequestAction As RequestActionEnum

        Try
            If Not Page.IsPostBack Then

                If Page.Request.QueryString("CallType") = Nothing OrElse Page.Request.QueryString("CallType") = String.Empty Then
                    Throw New Exception("Error #CM2248a Common GetRequestAction.Failed querystring.")
                End If
                Select Case Page.Request.QueryString("CallType").ToLower
                    Case "new"
                        RequestAction = RequestActionEnum.CreateNew
                    Case "existing"
                        RequestAction = RequestActionEnum.LoadExisting
                    Case Else
                        Throw New Exception("Error #CM2248b Common GetRequestAction.Failed querystring.")
                End Select

            Else

                If Page.Request.Form("hdAction") = Nothing OrElse Page.Request.Form("hdAction") = String.Empty Then
                    Throw New Exception("Error #CM2248c: Common GetRequestAction.Failed hdAction.")
                End If
                Select Case Page.Request("hdAction").ToLower
                    Case "savenew"
                        RequestAction = RequestActionEnum.SaveNew
                    Case "saveexisting"
                        RequestAction = RequestActionEnum.SaveExisting
                    Case "changetab"
                        RequestAction = RequestActionEnum.ChangeTab
                    Case "cancelchanges"
                        RequestAction = RequestActionEnum.CancelChanges
                    Case "return"
                        RequestAction = RequestActionEnum.ReturnToCallingPage
                    Case "other"
                        RequestAction = RequestActionEnum.Other
                    Case Else
                        Throw New Exception("Error #CM2248d: Common GetRequestAction.Failed hdAction.")
                End Select

            End If

            Return RequestAction

        Catch ex As Exception
            Throw New Exception("Error #CM2248e: Common GetRequestAction " & ex.Message)
        End Try
    End Function


    'Public Shared Function GetResponseActionFromRequestActionOther(ByRef Page As Page) As ResponseAction
    '    Select Case Page.Request("hdResponseAction").ToString
    '        Case ResponseAction.DisplayBlank.ToString
    '            Return ResponseAction.DisplayUserInputNew
    '        Case ResponseAction.DisplayUserInputNew.ToString
    '            Return ResponseAction.DisplayUserInputNew
    '        Case ResponseAction.DisplayUserInputExisting.ToString
    '            Return ResponseAction.DisplayUserInputExisting
    '        Case ResponseAction.DisplayUserInputNewOrExisting.ToString
    '            Return ResponseAction.DisplayUserInputNewOrExisting
    '        Case ResponseAction.DisplayExisting.ToString
    '            Return ResponseAction.DisplayUserInputExisting
    '        Case ResponseAction.ReturnToCallingPage.ToString
    '            Return ResponseAction.ReturnToCallingPage
    '    End Select
    'End Function

    Public Shared Sub DropdownFindByValueSelect(ByRef dd As DropDownList, ByVal Value As Object)
        Dim i As Integer
        Dim TestValue As String
        If IsDBNull(Value) OrElse Value = Nothing Then
            TestValue = String.Empty
        Else
            TestValue = Value
        End If
        TestValue = TestValue.ToLower

        For i = 0 To dd.Items.Count - 1
            If dd.Items(i).Value.ToLower = TestValue Then
                dd.Items(i).Selected = True
                Exit For
            End If
        Next
    End Sub

    Public Shared Function DropdownGetTextFromValue(ByRef dd As DropDownList, ByVal Value As Object) As String
        Dim i As Integer
        Dim TestValue As String
        Dim Results As String = String.Empty

        If IsBlank(Value) Then
            TestValue = String.Empty
        Else
            TestValue = Value
        End If
        TestValue = TestValue.ToLower

        For i = 0 To dd.Items.Count - 1
            If dd.Items(i).Value.ToLower = TestValue Then
                Results = dd.Items(i).Text
                Exit For
            End If
        Next
        Return Results
    End Function
#End Region

#Region " In handlers "
    Public Shared Function StrInHandler(ByVal Input As Object) As Object
        Dim Output As Object

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            Return String.Empty
            'ElseIf (Not IsDate(Input)) AndAlso Input.length = 0 Then
            '    Return String.Empty
        Else
            Output = Input
            If Input = Nothing Then
                Return String.Empty
            End If
            Return Output
        End If
    End Function

    Public Shared Function DateInHandler(ByVal Input As Object) As Object
        ' 12/31/2399
        Dim Output As Object
        Output = Input

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf Input = "01/01/1900" Then
            Return String.Empty
        ElseIf Input = "01/01/1950" Then
            Return String.Empty
        Else
            Return Output
        End If
    End Function

    Public Shared Function DateInHandler(ByVal Input As Object, ByVal FormatString As String) As Object
        ' 12/31/2399
        Dim Output As Object
        Output = Input

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf Input = "01/01/1900" Then
            Return String.Empty
        ElseIf Input = "01/01/1950" Then
            Return String.Empty
        ElseIf Not IsDate(Input) Then
            Return String.Empty
        Else
            'Return Output
            Return CType(Output, System.DateTime).ToString(FormatString)
        End If
    End Function

    Public Shared Function NumInHandler(ByVal Input As Object, ByVal NullAsZero As Boolean) As Object
        If IsDBNull(Input) Then
            If NullAsZero Then
                Return 0
            Else
                Return String.Empty
            End If
        Else
            Return Input
        End If
    End Function

    Public Shared Function GuidInHandler(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return String.Empty
        Else
            Return Input.ToString
        End If
    End Function

    Public Shared Function StrXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        Dim Output As Object = DBNull.Value
        Dim ReturnNull As Boolean

        If IsDBNull(Input) Then
            ReturnNull = True
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            ReturnNull = True
        Else
            Output = Replace(Input, "~", "'")
            If Output = Nothing Then
                ReturnNull = True
            End If
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return String.Empty
            End If
        Else
            Return Output
        End If
    End Function

    Public Shared Function DateXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        ' 12/31/2399
        Dim Output As Object
        Dim ReturnNull As Boolean
        Output = Input

        If IsDBNull(Input) OrElse Input = Nothing Then
            ReturnNull = True
        Else
            Output = Input
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return "1/1/1950"
            End If
        Else
            Return Output
        End If
    End Function

    Public Shared Function NumXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        Dim Output As Object = DBNull.Value
        Dim ReturnNull As Boolean

        If IsDBNull(Input) Then
            ReturnNull = True
        Else
            Output = Input
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return 0
            End If
        Else
            Return Output
        End If

    End Function
#End Region

#Region " Out handlers"
    'Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
    '    Dim ReturnNull As Boolean
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        Output = Replace(Input, "'", "~")
    '    Else
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    End If

    '    If ReturnNull Then
    '        Return "null"
    '    Else
    '        If AddSingleQuotes Then
    '            Return "'" & Output & "'"
    '        Else
    '            Return Output
    '        End If
    '    End If
    'End Function

    'Public Function OrigStrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
    '    Dim ReturnNull As Boolean
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        'Output = Replace(Input, "'", "~")

    '        If AddSingleQuotes Then
    '            Output = Replace(Input, "'", "''")
    '        Else
    '            Output = Input
    '        End If

    '    Else
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    End If

    '    If ReturnNull Then
    '        Return "null"
    '    Else
    '        If AddSingleQuotes Then
    '            Return "'" & Output & "'"
    '        Else
    '            Return Output
    '        End If
    '    End If
    'End Function

    'Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As StringTreatEnum) As Object
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        Output = Input
    '        If (StringTreat = StringTreatEnum.SecApost) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
    '            Output = Replace(Output, "'", "''")
    '        End If
    '        If (StringTreat = StringTreatEnum.SideQts) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
    '            Output = "'" & Output & "'"
    '        End If
    '    Else
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = String.Empty
    '            If (StringTreat = StringTreatEnum.SideQts) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
    '                Output = "'" & Output & "'"
    '            End If
    '        End If
    '    End If
    '    Return Output
    'End Function

    Public Shared Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As StringTreatEnum) As Object
        Dim Output As String

        Try

            ' ___ Output, adjusting for AllowNull
            If IsDBNull(Input) Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            Else
                Try
                    Output = Input
                Catch
                    If AllowNull Then
                        Output = "null"
                    Else
                        Output = String.Empty
                    End If
                End Try
            End If

            ' ___ Apply string treatment
            If Output <> "null" Then
                Select Case StringTreat
                    Case StringTreatEnum.AsIs
                        ' no action
                    Case StringTreatEnum.SecApost
                        Output = Replace(Output, "'", "''")
                    Case StringTreatEnum.SideQts
                        Output = "'" & Output & "'"
                    Case StringTreatEnum.SideQts_SecApost
                        Output = Replace(Output, "'", "''")
                        Output = "'" & Output & "'"
                End Select
            End If

            Return Output

        Catch ex As Exception
            Throw New Exception("Error #CM2220: Common StrOutHandler. " & ex.Message, ex)
        End Try
    End Function

    'Public Function NumOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean) As String
    '    Dim i As Integer
    '    Dim Output As String = String.Empty
    '    Dim Allowable As String
    '    Dim Working As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = "0"
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = "0"
    '        End If
    '    Else
    '        Try
    '            Working = Input
    '            Allowable = "-.0123456789"
    '            For i = 0 To Working.Length - 1
    '                If InStr(Allowable, Working.Substring(i, 1)) > 0 Then
    '                    Output &= Working.Substring(i, 1)
    '                End If
    '            Next
    '        Catch ex As Exception
    '            If AllowNull Then
    '                Output = "null"
    '            Else
    '                Output = "0"
    '            End If
    '        End Try
    '    End If

    '    Return Output
    'End Function

    'Public Shared Function NumOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal IntegerOnly As Boolean, ByVal ApplyFilter As Boolean) As String
    '    Dim i As Integer
    '    Dim Output As String = String.Empty
    '    Dim Allowable As String
    '    Dim Working As String
    '    Dim re As String
    '    Dim Match As Match

    '    Dim IsValidInd As Boolean



    '    'UNFINISHED

    '    If IntegerOnly Then
    '        re = "[^0-9]"
    '    Else
    '        re = "[^0-9-.]"
    '    End If

    '    Try

    '        If IsDBNull(Input) Then
    '            If AllowNull Then
    '                Output = "null"
    '            Else
    '                Output = "0"
    '            End If
    '        ElseIf Input Is Nothing Then
    '            If AllowNull Then
    '                Output = "null"
    '            Else
    '                Output = "0"
    '            End If
    '        Else


    '            Match = System.Text.RegularExpressions.Regex.Match(Input, re)
    '            If Match.Success Then
    '                IsValidInd = False
    '            Else
    '                IsValidInd = True
    '            End If


    '            Try
    '                Working = Input
    '                Allowable = "-.0123456789"
    '                For i = 0 To Working.Length - 1
    '                    If InStr(Allowable, Working.Substring(i, 1)) > 0 Then
    '                        Output &= Working.Substring(i, 1)
    '                    End If
    '                Next
    '                If Output.Length = 0 Then
    '                    If AllowNull Then
    '                        Output = "null"
    '                    Else
    '                        Output = "0"
    '                    End If
    '                End If
    '            Catch
    '                If AllowNull Then
    '                    Output = "null"
    '                Else
    '                    Output = "0"
    '                End If
    '            End Try
    '        End If

    '    Catch
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = "0"
    '        End If
    '    End Try

    '    Return Output
    'End Function


    Public Shared Function NumOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal AllowMinusAndNeg As Boolean) As String
        Dim i As Integer
        Dim Output As String = String.Empty
        Dim Allowable As String
        Dim Working As String

        Try

            If IsDBNull(Input) Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = "0"
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = "0"
                End If
            Else

                Try
                    Working = Input
                    If AllowMinusAndNeg Then
                        Allowable = "-.0123456789"
                    Else
                        Allowable = "0123456789"
                    End If

                    For i = 0 To Working.Length - 1
                        If InStr(Allowable, Working.Substring(i, 1)) > 0 Then
                            Output &= Working.Substring(i, 1)
                        End If
                    Next
                    If Output.Length = 0 Then
                        If AllowNull Then
                            Output = "null"
                        Else
                            Output = "0"
                        End If
                    End If
                Catch
                    If AllowNull Then
                        Output = "null"
                    Else
                        Output = "0"
                    End If
                End Try
            End If

        Catch
            If AllowNull Then
                Output = "null"
            Else
                Output = "0"
            End If
        End Try

        Return Output
    End Function

    Public Shared Function DateOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim ReturnNull As Boolean
        Dim Output As Object = DBNull.Value

        If IsDBNull(Input) OrElse Input = Nothing Then
            If AllowNull Then
                ReturnNull = True
            Else
                Output = "01/01/1950"
            End If
        Else
            Output = Input
        End If

        If ReturnNull Then
            Return "null"
        Else
            If AddSingleQuotes Then
                Return "'" & Output & "'"
            Else
                Return Output
            End If
        End If
    End Function

    Public Shared Function DateOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, ByVal FormatString As String, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim ReturnNull As Boolean
        Dim Output As Object = DBNull.Value

        If IsDBNull(Input) OrElse Input = Nothing Then
            If AllowNull Then
                ReturnNull = True
            Else
                Output = "01/01/1950"
            End If
        Else
            Output = Input
        End If

        If (Not ReturnNull) And IsDate(Output) Then
            Output = CType(Output, System.DateTime).ToString(FormatString)
        End If

        If ReturnNull Then
            Return "null"
        Else
            If AddSingleQuotes Then
                Return "'" & Output & "'"
            Else
                Return Output
            End If
        End If
    End Function

    Public Shared Function PhoneOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim i As Integer
        Dim Output As String = String.Empty
        Dim Working As String
        Working = StrOutHandler(Input, AllowNull, StringTreatEnum.SideQts)

        If Working = "null" Or Working = String.Empty Then
        Else
            If Working.Length >= 10 Then
                For i = 0 To Working.Length - 1
                    If IsNumeric(Working.Substring(i, 1)) Then
                        Output &= Working.Substring(i, 1)
                    End If
                Next
            End If
        End If

        If Output.Length = 10 Then
            Output = InsertAt(Output, "(", 1)
            Output = InsertAt(Output, ") ", 5)
            Output = InsertAt(Output, "-", 10)
        Else
            Output = Input
        End If

        If AddSingleQuotes Then
            Output = "'" & Output & "'"
        End If

        Return Output
    End Function

    Public Shared Function PageFormatPhone(ByVal Value As Object) As String
        Dim Results As String
        Dim Working As String

        If IsDBNull(Value) Then
            Results = String.Empty
        Else
            Working = CType(Value, System.String).PadLeft(10, "0")
            Results = "(" & Working.Substring(0, 3) & ") " & Working.Substring(3, 3) & "-" & Working.Substring(6, 4)
        End If
        Return Results
    End Function

    Public Shared Function DBFormatPhone(ByVal Value As String) As String
        Dim i As Integer
        Dim Results As New System.Text.StringBuilder
        Dim Match As Match

        ' Dim myMatch As Match = System.Text.RegularExpressions.Regex.Match(Value, "^[A-Z][a-zA-Z]*$")

        If Value.Length = 0 Then
            Results.Append(String.Empty)
        Else
            For i = 0 To Value.Length - 1
                Match = System.Text.RegularExpressions.Regex.Match(Value.Substring(i, 1), "[0-9]")
                If Match.Success Then
                    Results.Append(Value.Substring(i, 1))
                End If
            Next
        End If

        Return Results.ToString
    End Function



    'Public Function BitOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            Return "null"
    '        Else
    '            Return 0
    '        End If
    '    Else
    '        If CType(Input, Boolean) Then
    '            Return 1
    '        Else
    '            Return 0
    '        End If
    '    End If
    'End Function

    Public Shared Function BitOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, ByVal AddSingleQuoteToNull As Boolean) As Object
        Try
            If IsDBNull(Input) Then
                If AllowNull Then
                    Return "null"
                Else
                    Return 0
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Return "null"
                Else
                    Return 0
                End If
            ElseIf Input = String.Empty Then
                If AllowNull Then
                    Return "null"
                Else
                    Return 0
                End If

            Else
                If CType(Input, Boolean) Then
                    Return 1
                Else
                    Return 0
                End If
            End If
        Catch
            If AllowNull Then
                Return "null"
            Else
                Return 0
            End If
        End Try
    End Function


#End Region

#Region " Compare "
    Public Shared Function IsDateBetween(ByVal FromDate As Object, ByVal ToDate As Object, ByVal SubjectDate As DateTime) As Boolean
        Dim AfterStart As Boolean
        Dim BeforeEnd As Boolean
        Dim Results As Boolean

        If Not IsDate(FromDate) Then
            AfterStart = True
        Else
            If DateCompare(SubjectDate, FromDate, False) > -1 Then
                AfterStart = True
            End If
        End If

        If Not IsDate(ToDate) Then
            BeforeEnd = True
        Else
            If DateCompare(SubjectDate, ToDate, False) < 1 Then
                BeforeEnd = True
            End If
        End If

        If AfterStart And BeforeEnd Then
            Results = True
        End If

        Return Results
    End Function

    Public Shared Function DateCompare(ByVal Date1 As DateTime, ByVal Date2 As DateTime, ByVal CompareDatePartOnly As Boolean) As Single
        Try
            If CompareDatePartOnly Then
                Date1 = CType(Date1.ToString("MM/dd/yyyy") & " 00:00 AM", System.DateTime)
                Date2 = CType(Date2.ToString("MM/dd/yyyy") & " 00:00 AM", System.DateTime)
            End If
            Return Date.Compare(Date1, Date2)
        Catch ex As Exception
            Throw New Exception("Error #CM2210: Common DateCompare " & ex.Message, ex)
        End Try
    End Function

    Public Shared Sub CompareStrings(ByVal FldName As String, ByRef Coll As Collection, ByVal TableValue As Object, ByVal PageValue As String)
        If Not IsStrEqual(TableValue.ToText, PageValue) Then
            Coll.Add(New DataChangedPack(FldName, TableValue.ToText, PageValue))
        End If
    End Sub

    Public Shared Sub CompareDates(ByVal FldName As String, ByRef Coll As Collection, ByVal TableValue As Object, ByVal PageValue As Object, ByVal CompareFormatString As String, ByVal LeewayMinutes As Integer)
        Dim TableValueIsDate As Boolean
        Dim PageValueIsDate As Boolean
        Dim TimeSpan As TimeSpan

        Try

            TableValueIsDate = IsDate(TableValue.Value)
            PageValueIsDate = IsDate(PageValue)

            If TableValueIsDate Then
                If PageValueIsDate Then
                    If Not CType(TableValue.Value, System.DateTime).ToString(CompareFormatString) = CType(PageValue, System.DateTime).ToString(CompareFormatString) Then

                        If LeewayMinutes > 0 Then
                            TimeSpan = CType(TableValue.Value, System.DateTime).Subtract(CType(PageValue, System.DateTime))
                            If Math.Abs(TimeSpan.TotalMinutes) > LeewayMinutes Then
                                Coll.Add(New DataChangedPack(FldName, CType(TableValue.Value, System.DateTime).ToString(CompareFormatString), CType(PageValue, System.DateTime).ToString(CompareFormatString)))
                            End If
                        Else
                            Coll.Add(New DataChangedPack(FldName, CType(TableValue.Value, System.DateTime).ToString(CompareFormatString), CType(PageValue, System.DateTime).ToString(CompareFormatString)))
                        End If

                    End If
                Else
                    Coll.Add(New DataChangedPack(FldName, CType(TableValue.Value, System.DateTime).ToString(CompareFormatString), String.Empty))
                End If
            Else
                If PageValueIsDate Then
                    Coll.Add(New DataChangedPack(FldName, String.Empty, CType(PageValue, System.DateTime).ToString(CompareFormatString)))
                Else
                    ' no action
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #CM2205: Common CompareDates " & ex.Message, ex)
        End Try
    End Sub

    'Public Function CompareDecimals(ByVal FldName As String, ByRef Coll As Collection, ByVal TableValue As Object, ByVal PageValue As Object)
    '    Dim SameValue As Boolean
    '    Dim IsNumeric As Boolean = True

    '    Try

    '        If IsDBNull(TableValue.Value) Or PageValue.length = 0 Then
    '            IsNumeric = False
    '        Else
    '            If TableValue.Value = CType(Replace(PageValue, ",", ""), System.Decimal) Then
    '                SameValue = True
    '            End If
    '        End If

    '        If Not SameValue Then
    '            Coll.Add(New DataChangedPack(FldName, FormatNumber(CType(TableValue.Value, System.Decimal), 2), FormatNumber(CType(PageValue, System.Decimal), 2)))
    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("Error #CM2206: Common CompareDecimals " & ex.Message, ex)
    '    End Try
    'End Function

    Public Shared Sub CompareMoney(ByVal FldName As String, ByRef Coll As Collection, ByVal TableValue As Object, ByVal PageValue As Object, ByVal NumDecPlaces As Integer)
        Dim SameValue As Boolean
        Dim Working1 As Decimal
        Dim Working2 As Decimal

        Try

            If IsDBNull(TableValue) Or PageValue.length = 0 Then
            Else
                Working1 = Math.Round(TableValue, NumDecPlaces)
                Working2 = CType(Replace(PageValue, ",", ""), System.Decimal)
                Working2 = Math.Round(Working2, NumDecPlaces)
                If Working1 = Working2 Then
                    SameValue = True
                End If
            End If

            If Not SameValue Then
                'Coll.Add(New DataChangedPack(FldName, IIf(TableValue Is DBNull.Value, "null", Working1), FormatNumber(CType(PageValue, System.Decimal), 2)))
                Coll.Add(New DataChangedPack(FldName, IIf(TableValue Is DBNull.Value, "null", Working1), IIf(IsNumeric(PageValue), Working2, "")))
            End If

        Catch ex As Exception
            Throw New Exception("Error #CM2206: Common CompareDecimals " & ex.Message, ex)
        End Try
    End Sub
#End Region

#Region " Validate "
    Public Shared Function IsStrEqual(ByVal FirstValue As String, ByVal SecondValue As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        Dim Output As Integer
        Dim Results As Boolean
        Output = String.Compare(Trim(FirstValue), Trim(SecondValue), IgnoreCase)
        If Output = 0 Then
            Results = True
        End If
        Return Results
    End Function

    'Public Shared Function InList(ByVal SearchFor As String, ByVal SearchInList As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean
    '    Dim i As Integer
    '    Dim Box() As String

    '    If SearchInList Is Nothing OrElse SearchInList.Length = 0 Then
    '        Return False
    '    Else
    '        Box = Split(SearchInList, ",")
    '        For i = 0 To Box.GetUpperBound(0)
    '            If IsStrEqual(SearchFor, Trim(Box(i)), IgnoreCase) Then
    '                Return True
    '            End If
    '        Next
    '        Return False
    '    End If
    'End Function

    Public Shared Function InList(ByVal SearchFor As String, ByVal SearchInList As String, ByVal ExactMatch As Boolean, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        Dim i As Integer
        Dim Box() As String

        If SearchInList Is Nothing OrElse SearchInList.Length = 0 Then
            Return False
        Else
            Box = Split(SearchInList, ",")

            If ExactMatch Then

                For i = 0 To Box.GetUpperBound(0)
                    If IsStrEqual(SearchFor, Trim(Box(i)), IgnoreCase) Then
                        Return True
                    End If
                Next

            Else

                ' is 2 in 1?
                '? "jwatroba" like "*w*"
                'True

                For i = 0 To Box.GetUpperBound(0)

                    If InStr(SearchFor.ToLower, Trim(Box(i)).ToLower, CompareMethod.Text) > 0 Then
                        Return True
                    End If

                Next

            End If

            Return False
        End If
    End Function

    Public Shared Function IsBlank(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return True
        ElseIf Value = Nothing Then
            Return True
        Else
            If Value.length = 0 Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Shared Function IsNotBlank(ByVal Value As Object) As Boolean
        Return Not IsBlank(Value)
    End Function

    Public Shared Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal MaxLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Or Value.length > MaxLength Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateApostrophe(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If InStr(Value, "'") > 0 Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateNumericField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                PassTest = True
            Else
                PassTest = False
            End If
        End If

        If Not PassTest Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateDateTime(ByRef ErrColl As Collection, ByVal Input As Object, ByVal ValidateDatePart As Boolean, ByVal ValidateTimePart As Boolean, ByVal ErrMsg As String)
        Dim DatePart As String
        Dim TimePart As String
        Dim DateIsValid As Boolean = True
        Dim TimeIsValid As Boolean = True
        Dim ReportError As Boolean

        If ValidateDatePart Then
            Try
                DatePart = CType(Input, System.DateTime).ToString("MM/dd/yyyy")
                If Not IsDate(DatePart) Then
                    DateIsValid = False
                End If
            Catch
                DateIsValid = False
            End Try
        End If

        If ValidateTimePart Then
            Try
                TimePart = CType(Input, System.DateTime).ToString("hh:mm tt")
                If Not IsDate(CType(Date.Now.ToString("MM/dd/yyyy") & " " & TimePart, System.DateTime)) Then
                    TimeIsValid = False
                End If
            Catch
                TimeIsValid = False
            End Try
        End If

        If ValidateDatePart And Not ValidateTimePart Then
            If Not DateIsValid Then
                ReportError = True
            End If
        ElseIf Not ValidateDatePart And ValidateTimePart Then
            If Not TimeIsValid Then
                ReportError = True
            End If
            If ValidateDatePart And ValidateTimePart Then
                If DateIsValid And Not TimeIsValid Then
                    ReportError = True
                ElseIf Not DateIsValid And TimeIsValid Then
                    ReportError = True
                ElseIf Not DateIsValid And Not TimeIsValid Then
                    ReportError = True
                End If
            End If
        End If

        If ReportError Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateRadio(ByRef ErrColl As Collection, ByVal SelectedIndex As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        If (SelectedIndex < 0) AndAlso (Not AllowNull) Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateNumericRange(ByRef ErrColl As Collection, ByVal Value As Object, ByVal Min As Integer, ByVal Max As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                If Value >= Min AndAlso Value <= Max Then
                    PassTest = True
                Else
                    PassTest = False
                End If
            End If
        End If

        If Not PassTest Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateDateField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim Valid As Boolean
        If IsDBNull(Value) OrElse Value = Nothing Then
            If AllowNull Then
                Valid = True
            Else
                Valid = False
            End If
        ElseIf IsDate(Value) Then
            Valid = True
        Else
            Valid = False
        End If
        If Not Valid Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Function IsValidPhoneNumber(ByVal Value As Object) As Boolean
        Dim i As Integer
        Dim NumCount As Integer

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return False
        End If

        If Value.length >= 10 Then
            For i = 0 To Value.Length - 1
                If IsNumeric(Value.Substring(i, 1)) Then
                    NumCount += 1
                End If
            Next
        End If

        If NumCount = 10 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Sub ValidatePhoneNumber(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidPhoneNumber(Value) Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateEmailAddress(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidEmailAddress(Value) Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Function IsValidEmailAddress(ByVal Value As Object) As Boolean
        Dim OKSoFar As Boolean = True
        Const InvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Integer
        Dim Num As Integer
        Dim DotPos As Integer
        Dim Part2 As String = Nothing

        ' ___ Check for null or empty value
        If IsDBNull(Value) OrElse Value = Nothing Then
            OKSoFar = False
        End If

        ' ___ Check for minimum length
        If OKSoFar Then
            If Value.Length < 5 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for a double quote
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, Chr(34)) > 0  'Check to see if there is a double quote
        End If

        ' ___ Check for consecutive dots
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, "..") > 0
        End If

        ' ___ Check for invalid characters
        If OKSoFar Then
            For i = 0 To InvalidChars.Length - 1
                If InStr(1, Value, InvalidChars.Substring(i, 1)) > 0 Then
                    OKSoFar = False
                    Exit For
                End If
            Next
        End If

        ' ___ Check for number of @ symbols
        If OKSoFar Then
            For i = 0 To Value.Length - 1
                If InStr(Value.Substring(i, 1), "@") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for the @ symbol in starting before the third position
        If OKSoFar Then
            If InStr(Value, "@") < 2 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for number of dots
        If OKSoFar Then
            Num = 0
            Part2 = Value.substring(InStr(Value, "@"))
            For i = 0 To Part2.Length - 1
                If InStr(Part2.Substring(i, 1), ".") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Dot is present and not immediately after ampersand and not at end. 
        '___  Dot separated from ampersand by at least one character
        If OKSoFar Then
            DotPos = InStr(Part2, ".")
            If DotPos < 2 Or DotPos = Part2.Length Then
                OKSoFar = False
            End If
        End If

        Return OKSoFar
    End Function

    Public Shared Sub ValidateDropDown(ByRef ErrColl As Collection, ByRef dd As DropDownList, ByVal MinSelectedIndex As Integer, ByVal ErrMsg As String)
        If dd.SelectedIndex < MinSelectedIndex Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateDropDownSelect0(ByRef ErrColl As Collection, ByRef dd As DropDownList, ByVal ErrMsg As String)
        If dd.SelectedIndex > 0 Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateCheckbox(ByRef ErrColl As Collection, ByRef chkBox As CheckBox, ByVal ValidState As Integer, ByVal ErrMsg As String)
        Dim IsValid As Boolean = True
        If ValidState = 0 AndAlso Not chkBox.Checked Then
            IsValid = False
        ElseIf ValidState = 1 AndAlso Not chkBox.Checked Then
            IsValid = False
        End If
        If Not IsValid Then
            ErrColl.Add(ErrMsg)
        End If
    End Sub

    Public Shared Sub ValidateErrorOnly(ByRef ErrColl As Collection, ByVal ErrMsg As String)
        ErrColl.Add(ErrMsg)
    End Sub

    Public Shared Function ErrCollToString(ByRef ErrColl As Collection, ByVal Intro As String) As String
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer
        If ErrColl.Count > 0 Then
            For i = 1 To ErrColl.Count
                sb.Append(ErrColl(i))
            Next
        End If
        Return Intro & " " & sb.ToString & "."
    End Function

    Public Shared Function ErrCollToString(ByRef ErrColl As Collection, ByVal Intro As String, ByVal NewLine As Boolean) As String
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer
        If ErrColl.Count > 0 Then

            If NewLine Then
                sb.Append(Intro)
                For i = 1 To ErrColl.Count
                    sb.Append("\n" & ErrColl(i))
                Next

            Else
                sb.Append(Intro & " ")
                For i = 1 To ErrColl.Count
                    If i < ErrColl.Count Then
                        sb.Append(ErrColl(i) & ", ")
                    Else
                        sb.Append(ErrColl(i) & ".")
                    End If
                Next
            End If
        End If

        Return sb.ToString
    End Function
#End Region

#Region " This to that "
    Public Shared Function BoolToBitString(ByVal Value As Object, ByVal AllowNull As Boolean) As String
        If IsDBNull(Value) Then
            If AllowNull Then
                Return "null"
            Else
                Return "0"
            End If
        End If
        If Value Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Public Shared Function BitToRadio(ByVal Value As Object, ByVal TrueIndex As Integer, ByVal AllowNoneSelected As Boolean) As Integer
        Dim FalseIndex As Integer
        FalseIndex = System.Math.Abs(TrueIndex - 1)

        If IsDBNull(Value) Then
            If AllowNoneSelected Then
                Return -1
            Else
                Return FalseIndex
            End If
        Else
            If Value Then
                Return TrueIndex
            Else
                Return FalseIndex
            End If
        End If
    End Function

    Public Shared Function BitToString(ByVal Value As Object, ByVal TrueString As String, ByVal FalseString As String, ByVal AllowNull As Boolean) As String
        If IsDBNull(Value) Then
            If AllowNull Then
                Return String.Empty
            Else
                Return FalseString
            End If
        End If
        If Value Then
            Return TrueString
        Else
            Return FalseString
        End If
    End Function

    Public Shared Function BitToInt(ByVal Value As Object) As Integer
        If IsDBNull(Value) Then
            Return 0
        Else
            If Value Then
                Return 1
            Else
                Return 0
            End If
        End If
    End Function

    Public Shared Function BitToBool(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return False
        Else
            If Value Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Shared Function ChkToInd(ByVal chkBox As CheckBox) As Integer
        If chkBox.Checked Then
            Return 1
        Else
            Return 0
        End If
    End Function

    Public Shared Sub IndToChk(ByVal Ind As Object, ByVal chkBox As CheckBox)
        If IsDBNull(Ind) Then
            chkBox.Checked = False
        Else
            If Ind Then
                chkBox.Checked = True
            Else
                chkBox.Checked = False
            End If
        End If
    End Sub
#End Region

#Region " Everything else "
    'Public Function FormatCalendar(ByRef Calendar1 As System.Web.UI.WebControls.Calendar)
    '    With Calendar1
    '        .Font.Name = "Verdana;Arial"
    '        .NextPrevFormat = NextPrevFormat.ShortMonth
    '        .Font.Size = System.Web.UI.WebControls.FontUnit.Point(8)
    '        .DayHeaderStyle.BackColor = Color.LightSkyBlue
    '        .DayStyle.BackColor = Color.White
    '        .OtherMonthDayStyle.ForeColor = Color.Gray
    '        .OtherMonthDayStyle.BackColor = Color.White
    '        .TitleStyle.BackColor = Color.LightSkyBlue
    '        .TitleStyle.ForeColor = Color.Black
    '        .TitleStyle.Font.Bold = True
    '        .SelectedDayStyle.BackColor = Color.Navy
    '        .SelectedDayStyle.Font.Bold = False
    '        ' .SelectedDate = Common.GetServerDateTime
    '    End With
    'End Function

    Public Shared Function GetServerDateTime() As DateTime
        Return Date.Now.ToUniversalTime.AddHours(-5)
        'Dim CmdAsst As New CmdAsst(CommandType.StoredProcedure, "ServerDateTime")
        'Dim QueryPack As CmdAsst.QueryPack
        'Try
        '    QueryPack = CmdAsst.Execute
        'Catch
        '    QueryPack.Success = False
        'End Try
        'Return QueryPack.dt.rows(0)("ServerDateTime")
    End Function

    'Public Function NumPart(ByVal Value As String) As Integer
    '    Dim i As Integer
    '    Dim Result As String
    '    For i = 0 To Value.Length - 1
    '        If Asc(Value.Substring(i, 1)) > 47 And Asc(Value.Substring(i, 1)) < 58 Then
    '            Result &= Value.Substring(i, 1)
    '        End If
    '    Next
    '    If Result = Nothing Then
    '        Return -1
    '    Else
    '        Return CType(Result, System.Int64)
    '    End If
    'End Function

    Public Function ConvertLocalTimeToServerTime(ByVal LocalDateTime As Object) As Object
        ' Public Function ConvertLocalTimeToEasternStandardTime(ByVal LocalDateTime As DateTime) As DateTime
        Dim WorkingDateTime As DateTime
        Dim DaylightSavingsInd As Boolean

        ' ___ Return null value if null passed in
        If IsDBNull(LocalDateTime) Then
            Return LocalDateTime
        End If

        ' ___ Is daylight savings time in effect for the local datetime?
        Dim ServerTimeZone As TimeZone = TimeZone.CurrentTimeZone
        DaylightSavingsInd = ServerTimeZone.IsDaylightSavingTime(LocalDateTime)

        ' ___ If so, convert to local standard time by subtracting one hour
        If DaylightSavingsInd Then
            LocalDateTime = LocalDateTime.AddHours(-1)
        End If

        ' ___ Convert local standard time to UTC time
        WorkingDateTime = LocalDateTime.AddHours(cEnviro.ClientUTCOffsetHours)

        ' ___ Convert the UTC time to server time
        WorkingDateTime = WorkingDateTime.AddHours(-5)

        Return WorkingDateTime
    End Function

    'Public Sub RecordLoggedInUserData(ByVal LoggedInUserID As String, ByVal SessionID As String, ByVal LastLoginIP As String)
    '    Dim dt As DataTable
    '    Dim Sql As String

    '    dt = GetDT("SELECT LastSessionID FROM Users WHERE UserID = '" & LoggedInUserID & "'", cEnviro.DBHost, cEnviro.DefaultDatabase)
    '    If dt.Rows(0)("LastSessionID") <> SessionID Then
    '        Sql = "UPDATE Users SET LastSessionID = '" & SessionID & "', LastLoginDate = '" & GetServerDateTime() & "', LastLoginIP = '" & LastLoginIP & "' WHERE UserID = '" & LoggedInUserID & "'"
    '        ExecuteNonQuery(Sql)
    '    End If
    'End Sub

    Public Function ConvertServerTimeToLocalTime(ByVal ServerDateTime As Object) As Object
        'Public Function ConvertEasternStandardTimeToLocalTime(ByVal ServerDateTime As DateTime) As DateTime
        Dim WorkingDateTime As DateTime
        Dim DaylightSavingsInd As Boolean

        ' ___ Return null value if null passed in
        If IsDBNull(ServerDateTime) Then
            Return ServerDateTime
        End If

        ' ___ Convert server time to UTC time
        WorkingDateTime = ServerDateTime.AddHours(5)

        ' ___ Convert the UTC to local standard time
        WorkingDateTime = WorkingDateTime.AddHours(cEnviro.ClientUTCOffsetHours * (-1))

        ' ___ Is daylight savings time in effect for the server datetime?
        Dim ServerTimeZone As TimeZone = TimeZone.CurrentTimeZone
        DaylightSavingsInd = ServerTimeZone.IsDaylightSavingTime(WorkingDateTime)

        ' ___ If so, convert to local daylight savings time by adding an hour
        If DaylightSavingsInd Then
            WorkingDateTime = WorkingDateTime.AddHours(1)
        End If

        Return WorkingDateTime
    End Function

    'Public Function ToPacificTime(ByVal SubjDateTime As DateTime) As DateTime
    '    Dim PacificTime As DateTime
    '    Dim FirstOfMonthDOW As Integer
    '    Dim FirstOfMonth As DateTime
    '    Dim Offset As Integer
    '    Dim DSStartDt As DateTime
    '    Dim DSEndDt As DateTime

    '    FirstOfMonth = CType("#3/1/" & DateTime.Now.ToString("yyyy") & "#", DateTime)
    '    FirstOfMonthDOW = CType(FirstOfMonth.DayOfWeek, Integer) + 1
    '    Offset = 16 - FirstOfMonthDOW
    '    DSStartDt = CType("#3/" & Offset & "/" & SubjDateTime.ToString("yyyy") & "#", DateTime)

    '    FirstOfMonth = CType("#11/1/" & DateTime.Now.ToString("yyyy") & "#", DateTime)
    '    FirstOfMonthDOW = CType(FirstOfMonth.DayOfWeek, Integer) + 1
    '    Offset = 9 - FirstOfMonthDOW
    '    DSEndDt = CType("#11/" & Offset & "/" & SubjDateTime.ToString("yyyy") & "#", DateTime)

    '    If (SubjDateTime.ToString("yyyyMMdd") < DSStartDt.ToString("yyyyMMdd")) OrElse (SubjDateTime.ToString("yyyyMMdd") >= DSEndDt.ToString("yyyyMMdd")) Then
    '        PacificTime = SubjDateTime.AddHours(-8)    ' Standard time
    '    Else
    '        PacificTime = SubjDateTime.AddHours(-7)    ' Daylight savings time
    '    End If
    '    Return PacificTime
    'End Function

    Public Shared Function GetRightsStr(ByRef dt As DataTable) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        If dt.Rows.Count = 0 Then
            Return String.Empty
        Else
            For i = 0 To dt.Rows.Count - 1
                sb.Append("|" & StrInHandler(dt.Rows(i)("RightCd")))
            Next
            Return sb.ToString.Substring(1)
        End If
    End Function

    'Public Function ConditionStringForHTML(ByVal Value As Object) As String
    '    Dim Results As String
    '    If IsDBNull(Value) Then
    '        Results = String.Empty
    '    Else
    '        Results = Value.ToString
    '    End If
    '    Results = Replace(Results, Chr(10).ToString, "<br />")
    '    Return Results
    'End Function

    Public Shared Function ToNull(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return DBNull.Value
        ElseIf Input Is Nothing Then
            Return DBNull.Value
        ElseIf Input.length = 0 Then
            Return DBNull.Value
        Else
            Return Input
        End If
    End Function

    'Public Function DateSqlWhere(ByRef Input As Object) As String
    '    If IsDBNull(Input) Then
    '        Return " is null "
    '    ElseIf Input = Nothing Then
    '        Return " is null "
    '    ElseIf Trim(Input).Length = 0 Then
    '        Return " is null "
    '    Else
    '        Return " = '" & Input & "' "
    '    End If
    'End Function

    'Public Function DateSqlWhereNoNull(ByRef Input As Object) As String
    '    If IsDBNull(Input) Then
    '        Return " = '01/01/1900' "
    '    ElseIf Input = Nothing Then
    '        Return " = '01/01/1900' "
    '    ElseIf Trim(Input).Length = 0 Then
    '        Return " = '01/01/1900' "
    '    Else
    '        Return " = '" & Input & "' "
    '    End If
    'End Function

    Public Shared Function IsBVIDate(ByVal Input As Object) As Boolean
        If IsDBNull(Input) Then
            Return False
        ElseIf Input = Nothing Then
            Return False
        ElseIf Input = "01/01/1950" Then
            Return False
        ElseIf Input.ToString = String.Empty Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function GetRandomlyGeneratedPassword(ByVal Length As Integer) As String
        Dim i As Integer
        Dim RndValue As Integer
        Dim sb As New System.Text.StringBuilder

        ' ___ Generate random password 8 digits long
        For i = 1 To Length
            Randomize()
            RndValue = CInt(Int(62 * Rnd() + 1))
            Select Case RndValue
                Case 1 To 10
                    sb.Append(Chr(RndValue + 47))
                Case 11 To 36
                    sb.Append(Chr(RndValue + 54))
                Case 37 To 62
                    sb.Append(Chr(RndValue + 28))
            End Select
        Next
        Return sb.ToString.ToLower

    End Function

    Public Shared Function GetNewRecordID(ByVal TableName As String, ByVal KeyFldName As String) As Integer
        Dim RandValue As Integer
        Dim dt As DataTable

        Try

            Do
                Do
                    Randomize()
                    RandValue = CType(Rnd() * 1000000, System.Int64)
                Loop Until RandValue > 99999
                dt = GetDT("SELECT Count (*) FROM " & TableName & " WHERE " & KeyFldName & " = " & RandValue)
                If dt.Rows(0)(0) = 0 Then
                    Exit Do
                End If
            Loop

            Return RandValue

        Catch ex As Exception
            Throw New Exception("Error #2210: Common GetNewRecordID. " & ex.Message, ex)
        End Try

    End Function

    Public Shared Function GetNewRecordID(ByVal TableName As String, ByVal KeyFldName As String, ByVal MinValue As Integer, ByVal MaxValue As Integer) As Integer
        Dim MaxNumDecPlaces As Integer
        Dim Factor As Integer
        Dim RandValue As Integer
        Dim dt As DataTable

        ' Rnd() return a number less than 1 and equal to or greater than 0.
        ' If was want a maximum number of 48,000, we need to allow for 5 decimal places. 
        ' To do this, we use the number 100,000 because it it 1 greater than the maximum 
        ' number having the number of decimal places we want.

        MaxNumDecPlaces = MaxValue.ToString.Length + 1
        Factor = "1".PadRight(MaxNumDecPlaces, "0")

        Do
            Do
                Randomize()
                RandValue = CType(Rnd() * CInt(Factor), System.Int64)
            Loop Until (RandValue >= MinValue) AndAlso (RandValue <= MaxValue)
            dt = Common.GetDT("SELECT Count (*) FROM " & TableName & " WHERE " & KeyFldName & " = " & RandValue)
            If dt.Rows(0)(0) = 0 Then
                Exit Do
            End If
        Loop

        Return RandValue
    End Function

    Public Shared Function GetNewRecordID(ByVal TableName As String, ByVal KeyFldName As String, ByVal NumDigits As Integer) As Integer
        Dim RandValue As Integer
        Dim dt As DataTable
        Dim MinValue As String
        Dim Factor As String

        Try

            MinValue = String.Empty.PadRight(NumDigits - 1, "9")
            Factor = "1"
            Factor = Factor.PadRight(NumDigits + 1, "0")

            Do
                Do
                    Randomize()
                    RandValue = CType(Rnd() * CInt(Factor), System.Int64)
                Loop Until RandValue > CInt(MinValue)
                dt = Common.GetDT("SELECT Count (*) FROM " & TableName & " WHERE " & KeyFldName & " = " & RandValue)
                If dt.Rows(0)(0) = 0 Then
                    Exit Do
                End If
            Loop

            Return RandValue

        Catch ex As Exception
            Throw New Exception("Error #2210: Common GetNewRecordID. " & ex.Message, ex)
        End Try
    End Function


    Public Shared Function GetRandomValue(ByVal Length As Integer, ByVal NumOnly As Boolean, ByVal AlphaOnly As Boolean) As String
        Dim RndValue As Integer
        Dim sb As New System.Text.StringBuilder

        If NumOnly And AlphaOnly Then
            Return "-1"
        End If

        Do
            Randomize()
            RndValue = CInt(Int(62 * Rnd() + 1))
            Select Case RndValue
                Case 1 To 10
                    If Not AlphaOnly Then
                        sb.Append(Chr(RndValue + 47))
                    End If
                Case 11 To 36
                    If Not NumOnly Then
                        sb.Append(Chr(RndValue + 54))
                    End If
                Case 37 To 62
                    If Not NumOnly Then
                        sb.Append(Chr(RndValue + 28))
                    End If
            End Select
        Loop Until sb.ToString.Length = Length
        Return sb.ToString.ToLower

    End Function

    Public Shared Function GetCurRightsHidden(ByVal RightsColl As Collection) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        For i = 1 To RightsColl.Count
            sb.Append(RightsColl(i) & "|")
        Next
        sb.Length -= 1
        Return "<input type='hidden' id='currentrights' name='currentrights' value=""" & sb.ToString & """>"
    End Function

    Public Shared Function GetCurRightsHidden(ByVal RightsStr As String) As String
        Return "<input type='hidden'id='currentrights' name='currentrights' value=""" & RightsStr & """ > "
    End Function

    Public Function GetCurRightsAndTopicsHidden(ByVal RightsColl As Collection) As String
        Dim i As Integer

        Dim sb As New System.Text.StringBuilder
        For i = 1 To RightsColl.Count
            sb.Append(RightsColl(i) & "|")
        Next
        sb.Length -= 1
        Return "<input type='hidden' id='currentrights' name='currentrights' value=""" & sb.ToString & """ /><input type='hidden' id='currenttopics' name = 'currenttopics' value=""PUB|" & cEnviro.LogInRoleCatgy.ToString & " ""/>"
    End Function

    Public Shared Function GetCurRightsAndTopicsHidden(ByVal RightsStr As String) As String
        Return "<input type='hidden'id='currentrights' name='currentrights' value=""" & RightsStr & """ > "
    End Function

    Public Shared Function DTSort(ByRef dt As DataTable, ByVal Filter As String, ByVal Sort As String, ByVal Ascending As Boolean) As DataTable
        'strExpr = "id > 5"
        ' Sort descending by CompanyName column.
        'strSort = "name DESC"
        ' Use the Select method to find all rows matching the filter.

        ' To filter out the null first row in a DS DT:
        ' Filter = "CUST_ID > ''"
        '.FilterOn("Holiday = 'xmas'")

        Dim Row As Integer
        'Dim Col As Integer
        Dim SortedRows() As DataRow
        Dim NewDT As New DataTable
        Dim dr As DataRow
        Dim CompoundSort As Boolean

        Try

            ' ___ Use the Select method to find all rows matching the filter.
            If Sort = Nothing Then
                SortedRows = dt.Select(Filter)
            Else
                If InStr(Sort, " ") > 0 Then
                    CompoundSort = True
                End If
                If CompoundSort Then
                    SortedRows = dt.Select(Filter, Sort)
                Else
                    If Ascending Then
                        SortedRows = dt.Select(Filter, Sort & " ASC")
                    Else
                        SortedRows = dt.Select(Filter, Sort & " DESC")
                    End If
                End If
            End If

            NewDT = dt.Copy
            NewDT.Rows.Clear()
            For Row = 0 To SortedRows.GetUpperBound(0)
                Dim ItemArray(dt.Columns.Count - 1) As Object
                ItemArray = SortedRows(Row).ItemArray
                dr = NewDT.NewRow()
                NewDT.Rows.Add(dr)
                NewDT.Rows(Row).ItemArray = ItemArray
            Next

            Return NewDT

        Catch ex As Exception
            Throw New Exception("Error #CM2208: Common DTSort. " & ex.Message, ex)
        End Try
    End Function

    Public Shared Function CloneColumn(ByVal SourceColumn As DataColumn) As DataColumn
        Dim ColumnName As String
        Dim ColumnType As Type
        Dim NewColumn As DataColumn

        ColumnName = SourceColumn.ColumnName
        ColumnType = SourceColumn.DataType
        NewColumn = New DataColumn(ColumnName, ColumnType)
        Return NewColumn
    End Function
#End Region

#Region " Generic "
    Public Shared Sub GenerateError()
        Dim a, b, c As Integer
        a = b / c
    End Sub

    Public Shared Function Rounder(ByVal Value As Object, ByVal NumDecPlaces As Integer) As Decimal
        Dim i As Integer
        Dim InputStr As String
        Dim SubmittedWithDec As Boolean
        Dim IsWholeNumber As Boolean
        Dim HasMinusSign As Boolean
        Dim DecimalPlace As Integer
        Dim OutputStr As String = Nothing
        Dim Truncator() As String
        Dim Carry As Boolean

        Try

            InputStr = CType(Value, System.String)
            If InStr(InputStr, "-") > 0 Then
                HasMinusSign = True
            End If
            InputStr = Replace(InputStr, ",", "")
            InputStr = Replace(InputStr, "$", "")
            InputStr = Replace(InputStr, "-", "")

            If (InStr(InputStr, ".") = 0) Or (InStr(InputStr, ".") = InputStr.Length) Then
                IsWholeNumber = True
                NumDecPlaces = 0
                DecimalPlace = 0
            Else
                DecimalPlace = InStr(InputStr, ".") - 1
                SubmittedWithDec = True
            End If

            If SubmittedWithDec Then

                ' ___ Shorten the NumDecPlaces if fewer than the number submitted
                If (InputStr.Length - InStr(InputStr, ".")) < NumDecPlaces Then
                    NumDecPlaces = InputStr.Length - InStr(InputStr, ".")
                End If

                InputStr = Replace(InputStr, ".", "")

                ' ___ Write the members of the input object to an array
                Dim Arr(InputStr.Length - 1) As Integer
                For i = 0 To InputStr.Length - 1
                    Arr(i) = CType(InputStr.Substring(i, 1), System.Int16)
                Next

                ' ___ Perform the right side of the rounding
                If SubmittedWithDec Then
                    For i = Arr.GetUpperBound(0) To (DecimalPlace + NumDecPlaces) Step -1
                        If Arr(i) >= 0 And Arr(i) <= 4 Then
                            Arr(i) = 0
                        ElseIf Arr(i) >= 5 And Arr(i) <= 10 Then
                            Arr(i) = Arr(i) = 0
                            Arr(i - 1) = Arr(i - 1) + 1
                        End If
                    Next

                    ' ____ Perform the left side of the rounding
                    If (DecimalPlace + NumDecPlaces - 1) > -1 Then
                        For i = (DecimalPlace + NumDecPlaces - 1) To 0 Step -1
                            If Arr(i) = 10 Then
                                Arr(i) = 0
                                If i = 0 Then
                                    Carry = True
                                Else
                                    Arr(i - 1) = Arr(i - 1) + 1
                                End If
                            End If
                        Next
                    End If

                End If

                ' ___ Start assembling the output
                For i = 0 To Arr.GetUpperBound(0)
                    OutputStr += Arr(i).ToString
                Next

                ' ___ Reinsert decimal point and truncate the number
                Select Case NumDecPlaces
                    Case 0

                        ' ___ Return input without decimal point
                        OutputStr = OutputStr.Substring(0, DecimalPlace - 1)
                    Case Else

                        ' ___ Reinsert the decimal point and truncate
                        OutputStr = OutputStr.Insert(DecimalPlace, ".")
                        Truncator = Split(OutputStr, ".")
                        If Truncator(1).Length > NumDecPlaces Then
                            Truncator(1) = Truncator(1).Substring(0, NumDecPlaces)
                        End If
                        OutputStr = Truncator(0) & "." & Truncator(1)
                End Select

                If Carry Then
                    OutputStr = "1" & OutputStr
                End If

            End If

            If IsWholeNumber Then
                OutputStr = InputStr
            End If

            If HasMinusSign Then
                OutputStr = "-" & OutputStr
            End If

            Return CType(OutputStr, System.Decimal)

        Catch ex As Exception
            Throw New Exception("Error #CM2212: Common.Rounder " & ex.Message)
        End Try
    End Function

    Public Shared Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String)
        Dim Report As New Report
        Report.SendEmail(SendTo, From, cc, Subject, TextBody)
    End Sub

    Public Sub PrintCSVVersionLocal(ByRef dt As DataTable, ByVal FileFullPath As String, ByRef TotalColl As Collection)
        Dim i, j As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim ColNum As Integer
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim FirstRow As Boolean = True
        Dim RecCount As Integer = 0
        Dim Found As Boolean


        For i = 0 To dt.Rows.Count - 1

            ' ___ Header row.
            If FirstRow Then
                For ColNum = 0 To dt.Columns.Count - 1
                    If ColNum > 0 Then
                        Sql.Append(",")
                    End If
                    Sql.Append("""")
                    Sql.Append(dt.Columns(ColNum).ColumnName)
                    Sql.Append("""")
                Next
                Sql.Append(vbCrLf)
                FirstRow = False
            End If

            ' ___ Data rows.
            For ColNum = 0 To dt.Columns.Count - 1
                If ColNum > 0 Then
                    Sql.Append(",")
                End If
                Sql.Append("""")
                Try
                    Sql.Append(dt.Rows(i)(ColNum).ToJSParm)
                Catch
                    Sql.Append(dt.Rows(i)(ColNum))
                End Try

                Sql.Append("""")
            Next
            Sql.Append(vbCrLf)

        Next

        ' ___ Totals
        If Not TotalColl Is Nothing Then
            For ColNum = 0 To dt.Columns.Count - 1
                If ColNum = 0 Then
                    Sql.Append("""")
                    Sql.Append("TOTAL")
                    Sql.Append("""")
                Else
                    Sql.Append(",")
                    For j = 1 To TotalColl.Count
                        Found = False
                        If TotalColl(j).ItemName = dt.Columns(ColNum).ColumnName Then
                            Sql.Append("""")
                            Sql.Append(TotalColl(j).Value)
                            Sql.Append("""")
                            Found = True
                            Exit For
                        End If
                    Next
                End If
            Next
            Sql.Append(vbCrLf)
        End If


        If FileFullPath = Nothing Then
            FileFullPath = cEnviro.ApplicationPath & "\csv.csv"
        End If

        If File.Exists(FileFullPath) Then
            File.Delete(FileFullPath)
        End If
        fs = File.Create(FileFullPath)
        fs.Close()

        sw = New StreamWriter(FileFullPath)
        sw.Write(Sql.ToString)
        sw.Close()
        sw.Close()
        fs.Close()

    End Sub

    Public Function GetDownloadPath(ByVal Page As Page) As Collection
        Dim PathsColl As New Collection
        Dim DirPath As String
        Dim FileName As String
        Dim FullPath As String

        DirPath = Page.Request.ServerVariables("APPL_PHYSICAL_PATH") & "TempData\"
        'FileName = "Rpt" & DateDiff("s", "1/1/2000", Now) & "a" & CInt(Rnd() * 1000) & ".csv"
        FileName = "Rpt_" & cEnviro.LoggedInUserID & "_" & Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyyMMdd_HHmmss_fff") & ".csv"
        FullPath = DirPath & FileName
        PathsColl.Add(DirPath, "DirPath")
        PathsColl.Add(FullPath, "AbsPath")
        PathsColl.Add("./TempData/" & FileName, "RelPath")
        Return PathsColl
    End Function
#End Region

#Region " Callback-specific methods "
    Public Property EnviroExists() As Boolean
        Get
            If Not IsNothing(cEnviro) Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)

        End Set
    End Property

    Public Shared Function GetPageCaption() As String
        If ConfigurationManager.AppSettings("ShowBuildNum") = "1" Then
            Return "Callback Mgr v" & Enviro.VersionNumber & " BuildNum: " & Enviro.BuildNum
        Else
            Return "Callback Mgr v" & Enviro.VersionNumber
        End If
    End Function

    Public Shared Function IsAcesClient(ByVal ClientID As String) As Boolean
        Dim dt As System.Data.DataTable
        dt = Common.GetDT("SELECT Count (*) FROM Codes_ClientID WHERE AcesInd = 1 AND ClientID = '" & ClientID & "'", Enviro.DBHost, "ProjectReports")
        If dt.Rows(0)(0) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function GetSessionObj(ByVal Name As String) As Object
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Return SessionObj(Name)
    End Function
#End Region
End Class