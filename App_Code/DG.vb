'Imports System.Data.SqlClient
Imports System.Data

Public Class DG
    Inherits DataGrid

#Region " Declarations "
    Public Event ChildDTRequest(ByVal ParmColl As Collection, ByRef ChildText As String)

    'Public Event ChildDTRequest2(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String, ByVal Parm2 As String)
    'Public Event ChildDTRequest3(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String, ByVal Parm2 As String, ByVal Parm3 As String)
    'Public Event ChildDTRequest4(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
    Private cChildTable As ChildTableClass
    Private cInternalFilter As DG.Filter
    Private cTotal As DG.Total
    Private cExternalFilter As DG.ExternalFilter
    Private cMenu As DG.Menu
    Private cMenuLevel2 As DG.MenuLevel2Class
    Private cAnchorObject As AnchorObject
    Private cColumnColl As New Collection
    Private cSortReferenceColl As New Collection
    Private cTemplateColl As New Collection
    Private cCommon As Common
    Private cRights As Rights
    Private cKeyFieldName As String
    Private cTableDef As String
    Private cUseDefaultTableDef As Boolean
    Private cDefaultColl As New Collection
    Private cCheckboxToggleColl As Collection
    Private cHiddenColumnColl As Collection
    Private cLeftPadding As String
    Private cDefaultSortField As String
    Private cDefaultSortDirection As DefaultSortDirectionEnum
    Private cInternalFilterOperationMode As FilterOperationModeEnum = FilterOperationModeEnum.NoFilter
    Private cInternalFilterInitialShowHide As FilterInitialShowHideEnum
    Private cRecordsInitialShowHide As RecordsInitialShowHideEnum = RecordsInitialShowHideEnum.RecordsInitialShow
    Private cShowInternalFilter As Boolean

    'Private cExternalFilterOperationMode As FilterOperationModeEnum = FilterOperationMode.NoFilter
    'Private cExternalFilterInitialShowHide As FilterInitialShowHideEnum
    'Private cExternalInitialShowHide As RecordsInitialShowHideEnum = RecordsInitialShowHideEnum.RecordsInitialShow
    Private cExternalInternalFilter As Boolean

    Private cAttachNewButton As Boolean
    Private cNewButtonRight As String
    Private cFormatAsSubTable As Boolean
    Private cNewOrderByField As String
    Private cNewSortDirection As String
#End Region

#Region " Enums "
    Public Enum Justify
        none = 0
        left = 1
        center = 2
        right = 3
    End Enum

    Public Enum ColumnType
        Databound = 1
        DataboundSpecial = 2
        Template = 3
        Link = 4
        CheckboxToggle = 5
        Hidden = 6
        FreeForm = 7
        ChildTableSelect = 8
        [Boolean] = 9
        [Date] = 10
        Icon = 11
        CompoundDatabound = 12
        CheckboxRowSelect = 13
    End Enum

    Public Enum OrderByType
        Initial = 1
        Recurring = 2
        Field = 3
        ReturnToPage = 4
    End Enum

    Public Enum FilterOperationModeEnum
        NoFilter = 1
        FilterAlwaysOn = 2
        FilterAlwaysOff = 3
        FilterSwitchable = 4
    End Enum

    Public Enum FilterInitialShowHideEnum
        FilterInitialShow = 1
        FilterInitalHide = 2
    End Enum

    Public Enum RecordsInitialShowHideEnum
        RecordsInitialShow = 1
        RecordsInitialHide = 2
    End Enum
    Public Enum DefaultSortDirectionEnum
        Ascending = 1
        Descending = 2
    End Enum

    Public Enum FilterTypeEnum
        Textbox = 1
        NameTextbox = 2
        Dropdown = 3
        DateCtlYM = 4
        DateCtlYMD = 5
        CheckboxRowSelect = 6
    End Enum
#End Region

#Region " Properties "
    Public ReadOnly Property NewOrderByField() As String
        Get
            Return cNewOrderByField
        End Get
    End Property
    Public ReadOnly Property NewSortDirection() As String
        Get
            Return cNewSortDirection
        End Get
    End Property
    Public WriteOnly Property FormatAsSubTable() As Boolean
        Set(ByVal Value As Boolean)
            cFormatAsSubTable = Value
        End Set
    End Property
    Public Property ChildTable() As ChildTableClass
        Get
            Return cChildTable
        End Get
        Set(ByVal value As ChildTableClass)
            cChildTable = value
        End Set
    End Property
    Public ReadOnly Property FilterOperationMode() As FilterOperationModeEnum
        Get
            Return cInternalFilterOperationMode
        End Get
    End Property
    Public ReadOnly Property FilterInitialShowHide() As FilterInitialShowHideEnum
        Get
            Return cInternalFilterInitialShowHide
        End Get
    End Property
    Public ReadOnly Property RecordsInitialShowHide() As RecordsInitialShowHideEnum
        Get
            Return cRecordsInitialShowHide
        End Get
    End Property
    Public ReadOnly Property MenuLevel2() As DG.MenuLevel2Class
        Get
            Return cMenuLevel2
        End Get
    End Property
#End Region

#Region " Constructors "
    Public Sub New()
        Dim dt As New DataTable
    End Sub

    Public Sub New(ByVal KeyFieldName As String, ByRef Common As Common, ByRef Rights As Rights, ByVal UseDefaultTableDef As Boolean, ByVal TableDef As String, ByVal DefaultSortField As String, ByVal DefaultSortDirection As DefaultSortDirectionEnum)
        cKeyFieldName = KeyFieldName
        cCommon = Common
        cUseDefaultTableDef = UseDefaultTableDef
        cTableDef = TableDef
        cRights = Rights
        cDefaultColl.Add("<table class='DG' style='LEFT: 150px; POSITION: absolute; TOP: 214px' cellSpacing='0' cellPadding='0' width='650' border='0'>", "StandardTableDef")
        cDefaultColl.Add("<table class='DGEmbedded' cellSpacing='0' cellPadding='0' width='100%' border='0'>", "EmbeddedTableDef")
        cLeftPadding = ""
        cDefaultSortField = DefaultSortField
        cDefaultSortDirection = DefaultSortDirection
    End Sub
#End Region

#Region " GetFilterWhereClauseForMenuLevel2 "
    Public Function GetFilterWhereClauseForMenuLevel2(ByVal OrderByType As DG.OrderByType, ByVal FilterOnOffState As String, ByVal FilterShowHideToggle As String) As String
        Dim FilterWhereClause As String
        Dim ShowFilter As Boolean
        Dim ShowRecords As Boolean

        ShowFilter = GetShowInternalFilter(OrderByType, FilterOnOffState, FilterShowHideToggle)
        RecordFilterValues(ShowFilter, FilterOnOffState)
        ShowRecords = GetShowRecords(OrderByType)
        If ShowRecords And FilterOperationMode <> DG.FilterOperationModeEnum.NoFilter Then
            FilterWhereClause = GetFilterWhereClause()
        Else
            FilterWhereClause = String.Empty
        End If
        Return FilterWhereClause
    End Function
#End Region

#Region " Generate SQL "
    'Public Sub GenerateSQL(ByRef Sql As String, ByRef ShowFilter As Boolean, ByVal NonFilterWhereClause As String, ByVal OrderByType As DG.OrderByType, ByRef Request As HttpRequest, ByVal FilterOnOffState As String, ByVal FilterShowHideToggle As String, Optional ByVal OmitWhere As Boolean = False)
    'Public Sub GenerateSQL(ByRef Sql As String, ByRef OrderByStr As String, ByRef ShowFilter As Boolean, ByVal NonFilterWhereClause As String, ByVal OrderByType As DG.OrderByType, ByRef Request As HttpRequest, ByVal FilterOnOffState As String, ByVal FilterShowHideToggle As String, ByRef QueryColl As CollX, Optional ByVal OmitWhere As Boolean = False)
    '    Dim ShowRecords As Boolean
    '    Dim FilterWhereClause As String = Nothing
    '    Dim WorkingSql As String = String.Empty
    '    Dim i As Integer

    '    If cInternalFilter Is Nothing Then
    '        If NonFilterWhereClause <> Nothing Then

    '            If OmitWhere Then
    '                Sql &= " AND " & NonFilterWhereClause
    '            Else
    '                Sql &= " WHERE " & NonFilterWhereClause
    '            End If

    '        End If
    '        Sql &= " ORDER BY " & Replace(cNewOrderByField, "~", "'")
    '        OrderByStr = Replace(cNewOrderByField, "~", "'")

    '        If cNewSortDirection = "A" Then
    '            Sql &= " ASC"
    '        Else
    '            Sql &= " DESC"
    '        End If
    '    Else

    '        ' ___ Show/hide the filter. Record filter values
    '        ShowFilter = GetShowInternalFilter(OrderByType, FilterOnOffState, FilterShowHideToggle)
    '        ShowRecords = GetShowRecords(OrderByType)
    '        RecordFilterValues(ShowFilter, FilterOnOffState)

    '        If ShowRecords And FilterOperationMode <> DG.FilterOperationModeEnum.NoFilter Then
    '            FilterWhereClause = GetFilterWhereClause()
    '        End If
    '        If Not ShowRecords Then
    '            FilterWhereClause = "  0 = 1 "
    '        End If

    '        If NonFilterWhereClause = Nothing Then
    '            NonFilterWhereClause = String.Empty
    '        End If

    '        If NonFilterWhereClause.Length = 0 Then
    '            If FilterWhereClause.Length = 0 Then
    '                ' No action
    '            Else
    '                WorkingSql = FilterWhereClause
    '            End If
    '        Else
    '            If FilterWhereClause.Length = 0 Then
    '                WorkingSql = NonFilterWhereClause
    '            Else
    '                WorkingSql = NonFilterWhereClause & " AND " & FilterWhereClause
    '            End If
    '        End If

    '        If WorkingSql.Length > 0 Then
    '            If OmitWhere Then
    '                Sql &= " AND " & WorkingSql
    '            Else
    '                Sql &= " WHERE " & WorkingSql
    '            End If
    '        End If

    '        Sql &= " ORDER BY " & cNewOrderByField
    '        If cNewSortDirection = "A" Then
    '            Sql &= " ASC"
    '        Else
    '            Sql &= " DESC"
    '        End If
    '    End If


    '    If Not cExternalFilter Is Nothing Then
    '        ' Record filter values
    '        For i = 1 To cExternalFilter.Coll.Count
    '            If Not cExternalFilter.Coll(i).IsLink Then
    '                If cExternalFilter.Coll(i).GetOverrideValue = Nothing Then
    '                    cExternalFilter.Coll(i).SetValue(Request.Form(cExternalFilter.Coll(i).CtlName))
    '                Else
    '                    cExternalFilter.Coll(i).SetValue(cExternalFilter.Coll(i).GetOverrideValue)
    '                End If
    '            Else
    '            End If
    '        Next
    '    End If

    'End Sub

    Public Function SetDisplayProperties(ByVal OrderByType As DG.OrderByType, ByVal FilterOnOffState As String, ByVal FilterShowHideToggle As String) As String
        Dim ShowRecords As Boolean
        Dim OrderByStr As String

        Try

            ' ___ Show/hide the filter. Record filter values
            'ShowFilter = GetShowInternalFilter(OrderByType, FilterOnOffState, FilterShowHideToggle)
            GetShowInternalFilter(OrderByType, FilterOnOffState, FilterShowHideToggle)
            ShowRecords = GetShowRecords(OrderByType)
            RecordFilterValues(True, FilterOnOffState)

            If cNewSortDirection = "A" Then
                OrderByStr = cNewOrderByField & " ASC"
            Else
                OrderByStr = cNewOrderByField & " DESC"
            End If
            Return OrderByStr

        Catch ex As Exception
            Throw New Exception("Error #DG429: DG SetDisplayProperties. " & ex.Message, ex)
        End Try
    End Function


    Public Sub GenerateSQLUnion(ByRef Sql1 As String, ByVal Sql2 As String, ByRef ShowFilter As Boolean, ByVal NonFilterWhereClause1 As String, ByVal NonFilterWhereClause2 As String, ByVal OrderByType As DG.OrderByType, ByRef Request As HttpRequest, ByVal FilterOnOffState As String, ByVal FilterShowHideToggle As String, Optional ByVal OmitWhere As Boolean = False)
        Dim ShowRecords As Boolean
        Dim FilterWhereClause1 As String = Nothing
        Dim FilterWhereClause2 As String = Nothing
        Dim WorkingSql1 As String = String.Empty
        Dim WorkingSql2 As String = String.Empty
        Dim i As Integer

        If cInternalFilter Is Nothing Then
            If NonFilterWhereClause1 <> Nothing Then
                If OmitWhere Then
                    Sql1 &= " AND " & NonFilterWhereClause1
                Else
                    Sql1 &= " WHERE " & NonFilterWhereClause1
                End If
            End If

            If NonFilterWhereClause2 <> Nothing Then
                If OmitWhere Then
                    Sql2 &= " AND " & NonFilterWhereClause2
                Else
                    Sql2 &= " WHERE " & NonFilterWhereClause2
                End If
            End If

            Sql2 &= " ORDER BY " & Replace(cNewOrderByField, "~", "'")
            If cNewSortDirection = "A" Then
                Sql2 &= " ASC"
            Else
                Sql2 &= " DESC"
            End If
        Else

            ' ___ Show/hide the filter. Record filter values
            ShowFilter = GetShowInternalFilter(OrderByType, FilterOnOffState, FilterShowHideToggle)
            ShowRecords = GetShowRecords(OrderByType)
            RecordFilterValues(ShowFilter, FilterOnOffState)

            ' ___ Sql1
            If ShowRecords And FilterOperationMode <> DG.FilterOperationModeEnum.NoFilter Then
                FilterWhereClause1 = GetFilterWhereClause()
            End If
            If Not ShowRecords Then
                FilterWhereClause1 = "  0 = 1 "
            End If

            If NonFilterWhereClause1 = Nothing Then
                NonFilterWhereClause1 = String.Empty
            End If

            If NonFilterWhereClause1.Length = 0 Then
                If FilterWhereClause1.Length = 0 Then
                    ' No action
                Else
                    WorkingSql1 = FilterWhereClause1
                End If
            Else
                If FilterWhereClause1.Length = 0 Then
                    WorkingSql1 = NonFilterWhereClause1
                Else
                    WorkingSql1 = NonFilterWhereClause1 & " AND " & FilterWhereClause1
                End If
            End If

            If WorkingSql1.Length > 0 Then
                If OmitWhere Then
                    Sql1 &= " AND " & WorkingSql1
                Else
                    Sql1 &= " WHERE " & WorkingSql1
                End If
            End If

            ' ___ Sql2
            If ShowRecords And FilterOperationMode <> DG.FilterOperationModeEnum.NoFilter Then
                FilterWhereClause2 = GetFilterWhereClause()
            End If
            If Not ShowRecords Then
                FilterWhereClause2 = "  0 = 1 "
            End If

            If NonFilterWhereClause2 = Nothing Then
                NonFilterWhereClause2 = String.Empty
            End If

            If NonFilterWhereClause2.Length = 0 Then
                If FilterWhereClause2.Length = 0 Then
                    ' No action
                Else
                    WorkingSql2 = FilterWhereClause2
                End If
            Else
                If FilterWhereClause2.Length = 0 Then
                    WorkingSql2 = NonFilterWhereClause2
                Else
                    WorkingSql2 = NonFilterWhereClause2 & " AND " & FilterWhereClause2
                End If
            End If

            If WorkingSql2.Length > 0 Then
                If OmitWhere Then
                    Sql2 &= " AND " & WorkingSql2
                Else
                    Sql2 &= " WHERE " & WorkingSql2
                End If
            End If



            Sql1 = Sql1 & " UNION " & Sql2



            Sql1 &= " ORDER BY " & cNewOrderByField
            If cNewSortDirection = "A" Then
                Sql1 &= " ASC"
            Else
                Sql1 &= " DESC"
            End If
        End If


        If Not cExternalFilter Is Nothing Then
            ' Record filter values
            For i = 1 To cExternalFilter.Coll.count
                If Not cExternalFilter.Coll(i).IsLink Then
                    If cExternalFilter.Coll(i).GetOverrideValue = Nothing Then
                        cExternalFilter.Coll(i).SetValue(Request.Form(cExternalFilter.Coll(i).CtlName))
                    Else
                        cExternalFilter.Coll(i).SetValue(cExternalFilter.Coll(i).GetOverrideValue)
                    End If
                Else
                End If
            Next
        End If

    End Sub
#End Region

#Region " Handle sorting "
    Public Sub SetSortElements(ByVal OrderByField As String, ByVal OrderByType As OrderByType)
        Dim CurSortDirection As String
        Dim NewOrderByField As String = Nothing
        Dim NewSortDirection As String = Nothing

        Try

            If OrderByType = OrderByType.Initial Then
                NewOrderByField = cDefaultSortField
                If cDefaultSortDirection = DefaultSortDirectionEnum.Ascending Then
                    NewSortDirection = "A"
                Else
                    NewSortDirection = "D"
                End If

            ElseIf OrderByType = OrderByType.Recurring Or OrderByType = OrderByType.ReturnToPage Then
                NewOrderByField = GetLastFieldSorted()
                If NewOrderByField = "" Then
                    NewOrderByField = cDefaultSortField
                    NewSortDirection = "A"
                Else
                    NewSortDirection = GetSortDirection(NewOrderByField)
                End If
            ElseIf OrderByType = OrderByType.Field Then
                NewOrderByField = OrderByField
                CurSortDirection = GetSortDirection(OrderByField)
                Select Case CurSortDirection
                    Case "A"
                        NewSortDirection = "D"
                    Case "N", "D"
                        NewSortDirection = "A"
                End Select
            End If
            SetLastFieldSorted(NewOrderByField)
            SetSortDirection(NewOrderByField, NewSortDirection)
            cNewOrderByField = NewOrderByField
            cNewSortDirection = NewSortDirection

        Catch ex As Exception
            Throw New Exception("Error #DG410: DG SetSortElements. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub AppendSortReference(ByVal SortExpression As String)
        Dim Position As String
        If SortExpression <> Nothing AndAlso SortExpression.Length > 0 Then
            Position = CStr(101 + cSortReferenceColl.Count).ToString
            cSortReferenceColl.Add(New SortItem(SortExpression, Position), SortExpression)
        End If
    End Sub

    Public Sub UpdateSortReference(ByVal SortReference As String)
        Dim i As Integer
        For i = 1 To cSortReferenceColl.Count
            cSortReferenceColl(i).SortDirection = SortReference.Substring((5 * (i - 1)) + 3, 1)
            cSortReferenceColl(i).LastFieldSorted = SortReference.Substring((5 * (i - 1)) + 4, 1)
        Next
    End Sub

    Public Function GetSortReference() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        For i = 1 To cSortReferenceColl.Count
            sb.Append(cSortReferenceColl(i).Position & cSortReferenceColl(i).SortDirection & cSortReferenceColl(i).LastFieldSorted)
        Next
        Return sb.ToString
    End Function

    Private Function GetSortDirection(ByVal FldName As String) As String
        If FldName = String.Empty Then
            Return String.Empty
        Else
            Return cSortReferenceColl(FldName).SortDirection
        End If
    End Function

    Private Sub SetSortDirection(ByVal FldName As String, ByVal Value As String)
        If cSortReferenceColl.Count > 0 Then
            cSortReferenceColl(FldName).SortDirection = Value
        End If
    End Sub

    Public Function GetLastFieldSorted() As String
        Dim i As Integer
        Dim Results As String = String.Empty

        For i = 1 To cSortReferenceColl.Count
            If cSortReferenceColl(i).LastFieldSorted = "T" Then
                Results = cSortReferenceColl(i).SortExpression
            End If
        Next
        Return Results
    End Function

    Public Sub SetLastFieldSorted(ByVal FldName As String)
        Dim i As Integer

        Try

            If cSortReferenceColl.Count > 0 Then
                For i = 1 To cSortReferenceColl.Count
                    cSortReferenceColl(i).LastFieldSorted = "F"
                Next
                If FldName <> String.Empty Then
                    cSortReferenceColl(FldName).LastFieldSorted = "T"
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #DG411: DG SetLastFieldSorted. FldName: " & FldName & " " & ex.Message, ex)
        End Try
    End Sub
#End Region

#Region " Handle the menu "
    Public Function AttachMenu(ByVal CellWidthPercent As Integer) As DG.Menu
        cMenu = New DG.Menu(CellWidthPercent)
        Return cMenu
    End Function
#End Region

#Region " Handle MenuLevel2 "
    Public Function AttachMenuLevel2() As DG.MenuLevel2Class
        cMenuLevel2 = New DG.MenuLevel2Class
        Return cMenuLevel2
    End Function
#End Region

#Region " Handle the total "
    Public Function GetTotal() As DG.Total
        Return cTotal
    End Function
    Public Function AttachTotal() As DG.Total
        cTotal = New DG.Total(cColumnColl)
        Return cTotal
    End Function
#End Region

#Region " Handle the internal filter "
    Public Function GetFilter() As DG.Filter
        Return cInternalFilter
    End Function

    Public Function AttachFilter(ByVal FilterOperationMode As FilterOperationModeEnum, ByVal FilterInitialShowHide As FilterInitialShowHideEnum, ByVal RecordsInitialShowHide As RecordsInitialShowHideEnum) As DG.Filter
        cInternalFilterOperationMode = FilterOperationMode
        cInternalFilterInitialShowHide = FilterInitialShowHide
        cRecordsInitialShowHide = RecordsInitialShowHide
        cInternalFilter = New DG.Filter
        Return cInternalFilter
    End Function

    Private Function GetShowInternalFilter(ByVal OrderByType As OrderByType, ByVal FilterOnOffState As String, ByVal FilterShowHideToggle As String) As Boolean
        Dim ShowFilter As Boolean

        Select Case cInternalFilterOperationMode
            Case DG.FilterOperationModeEnum.NoFilter, DG.FilterOperationModeEnum.FilterAlwaysOff
                ShowFilter = False
            Case DG.FilterOperationModeEnum.FilterAlwaysOn
                ShowFilter = True
            Case DG.FilterOperationModeEnum.FilterSwitchable
                Select Case OrderByType
                    Case DG.OrderByType.Initial
                        Select Case cInternalFilterInitialShowHide
                            Case DG.FilterInitialShowHideEnum.FilterInitialShow
                                ShowFilter = True
                            Case DG.FilterInitialShowHideEnum.FilterInitalHide
                                ShowFilter = False
                        End Select

                    Case OrderByType.ReturnToPage
                        Select Case FilterOnOffState
                            Case "on"
                                ShowFilter = True
                            Case Else
                                ShowFilter = False
                        End Select

                    Case Else

                        Select Case FilterOnOffState
                            Case "on"
                                Select Case FilterShowHideToggle
                                    Case "0"
                                        ShowFilter = True
                                    Case "1"
                                        ShowFilter = False
                                End Select
                            Case "off"
                                Select Case FilterShowHideToggle
                                    Case "0"
                                        ShowFilter = False
                                    Case "1"
                                        ShowFilter = True
                                End Select
                        End Select
                End Select
        End Select
        cShowInternalFilter = ShowFilter
        Return ShowFilter
    End Function

    Public Function GetShowRecords(ByVal OrderByType As DG.OrderByType) As Boolean
        Dim ShowRecords As Boolean
        Select Case OrderByType
            Case DG.OrderByType.Initial
                Select Case cRecordsInitialShowHide
                    Case DG.FilterInitialShowHideEnum.FilterInitialShow
                        ShowRecords = True
                    Case DG.FilterInitialShowHideEnum.FilterInitalHide
                        ShowRecords = False
                End Select
            Case Else
                ShowRecords = True
        End Select
        Return ShowRecords
    End Function

    Public Sub RecordFilterValues(ByVal ShowFilter As Boolean, ByVal FilterOn As String)
        Dim i As Integer

        ' ___ Internal filter
        Select Case FilterOn
            Case "on"
                Select Case ShowFilter
                    Case True
                        ' Write the values
                        For i = 1 To cInternalFilter.Coll.count
                            cInternalFilter.Coll(i).SetValue(cInternalFilter.Coll(i).GetFilterValue)
                        Next
                    Case False
                        ' Clear it out
                        For i = 1 To cInternalFilter.Coll.count
                            cInternalFilter.Coll(i).SetValue("")
                        Next
                End Select
            Case "off"
                ' No action
        End Select
    End Sub

    Public Function GetFilterWhereClause() As String
        Dim Coll As New Collection
        Dim i As Integer
        Dim FilterWhereClause As New System.Text.StringBuilder
        Dim SelectedValue As String
        Dim DateCtlDateRangeColl As Collection
        Dim FilterValue As String
        Dim Box() As String

        For i = 1 To cInternalFilter.Coll.count
            If cInternalFilter.Coll(i).GetValue <> String.Empty Then
                If cInternalFilter.Coll(i).FilterType = FilterTypeEnum.Textbox Then
                    If cInternalFilter.Coll(i).FilterField = Nothing Then
                        Coll.Add(Common.StrInHandler(cInternalFilter.Coll(i).DataFldName) & " LIKE  '" & Common.StrOutHandler(cInternalFilter.Coll(i).GetValue, False, StringTreatEnum.SecApost) & "%'")
                    Else
                        Coll.Add(Common.StrInHandler(cInternalFilter.Coll(i).FilterField) & " LIKE  '" & Common.StrOutHandler(cInternalFilter.Coll(i).GetValue, False, StringTreatEnum.SecApost) & "%'")
                    End If

                ElseIf cInternalFilter.Coll(i).FilterType = FilterTypeEnum.NameTextbox Then
                    FilterValue = Common.StrOutHandler(cInternalFilter.Coll(i).GetValue, False, StringTreatEnum.SecApost)
                    Box = Split(FilterValue, ",")
                    If Box.GetUpperBound(0) = 0 Then
                        Coll.Add(Common.StrInHandler(cInternalFilter.Coll(i).LastNameFld) & " LIKE  '" & Box(0) & "%'")
                    Else
                        Coll.Add(Common.StrInHandler(cInternalFilter.Coll(i).LastNameFld) & " LIKE  '" & Box(0) & "%' AND " & Common.StrInHandler(cInternalFilter.Coll(i).FirstNameFld) & " LIKE  '" & Box(1).Trim & "%'")
                    End If

                ElseIf cInternalFilter.Coll(i).FilterType = FilterTypeEnum.DateCtlYM Then
                    DateCtlDateRangeColl = GetDateCtlDateRange(cInternalFilter.Coll(i))
                    If DateCtlDateRangeColl(1) Then
                        Coll.Add(DateCtlDateRangeColl(2))
                    End If

                ElseIf cInternalFilter.Coll(i).FilterType = FilterTypeEnum.DateCtlYMD Then
                    DateCtlDateRangeColl = GetDateCtlDateRange(cInternalFilter.Coll(i))
                    If DateCtlDateRangeColl(1) Then
                        Coll.Add(DateCtlDateRangeColl(2))
                    End If
                Else
                    If cInternalFilter.Coll(i).IsStandard Then
                        Coll.Add(cInternalFilter.Coll(i).DataFldName & " = '" & cInternalFilter.Coll(i).GetValue & "'")
                    ElseIf cInternalFilter.Coll(i).IsExtended Then
                        SelectedValue = cInternalFilter.Coll(i).GetValue
                        Coll.Add(cInternalFilter.Coll(i).Coll(SelectedValue).Sql)
                    End If
                End If
            End If
        Next

        For i = 1 To Coll.Count
            If i = 1 Then
                FilterWhereClause.Append(Coll(i))
            Else
                FilterWhereClause.Append(" and " & Coll(i))
            End If
        Next
        Return FilterWhereClause.ToString
    End Function

    Private Shared Function GetDateCtlDateRange(ByRef FilterItem As Filter.FilterItem) As Collection
        Dim Coll As New Collection
        Dim Results As String
        Dim FromDate As String = Nothing
        Dim ToDate As String = Nothing
        Dim Dates As Object
        Dim Value As String
        Dim DataFldName As String

        DataFldName = FilterItem.DataFldName
        Value = FilterItem.GetValue
        Dates = Split(Value, "|")

        If FilterItem.FilterType = FilterTypeEnum.DateCtlYM Then
            If IsNumeric(Dates(0)) AndAlso IsNumeric(Dates(1)) Then
                FromDate = DataFldName & " >= '" & Dates(0) & "/01/" & Dates(1) & "'"
            End If
            If IsNumeric(Dates(2)) AndAlso IsNumeric(Dates(3)) Then
                ToDate = DataFldName & " <= '" & Dates(2) & "/" & GetDateCtlLastDayOfMonth(Dates(2), Dates(3)) & "/" & Dates(3) & "'"
            End If

        ElseIf FilterItem.FilterType = FilterTypeEnum.DateCtlYMD Then
            If IsNumeric(Dates(0)) AndAlso IsNumeric(Dates(1)) AndAlso IsNumeric(Dates(2)) Then
                FromDate = DataFldName & " >= '" & Dates(1) & "/" & Dates(2) & "/" & Dates(0) & "'"

            End If
            If IsNumeric(Dates(3)) AndAlso IsNumeric(Dates(4)) AndAlso IsNumeric(Dates(5)) Then
                ToDate = DataFldName & " < '" & GetNextDay(Dates(4), Dates(5), Dates(3)) & "'"
            End If
        End If

        '2008-11-08 23:59:59.999

        If FromDate = Nothing Then
            Results = Nothing
        Else
            If ToDate = Nothing Then
                Results = FromDate
            Else
                Results = FromDate & " AND " & ToDate
            End If
        End If
        If Results = Nothing Then
            Coll.Add(False)
        Else
            Coll.Add(True)
            Coll.Add(Results)
        End If
        Return Coll
    End Function

    Private Shared Function GetNextDay(ByVal Month As Integer, ByVal Day As Integer, ByVal Year As Integer) As String
        Dim LastDayOfMonth As Integer
        Dim Results As String

        LastDayOfMonth = GetDateCtlLastDayOfMonth(Month, Year)
        If Day = LastDayOfMonth Then
            If Month = 12 Then
                Results = "1/1/" & Year + 1
            Else
                Results = Month + 1 & "/" & 1 & "/" & Year
            End If
        Else
            Results = Month & "/" & Day + 1 & "/" & Year
        End If

        Return Results
    End Function

    Private Shared Function GetDateCtlLastDayOfMonth(ByVal Month As Integer, ByVal Year As Integer) As Integer
        Month = Month - 1
        If ((Year Mod 4 = 0) And (Year Mod 100 <> 0) Or (Year Mod 400 = 0)) Then
            Dim Days() As String = {31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
            Return Days(Month)
        Else
            Dim Days() As String = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
            Return Days(Month)
        End If
    End Function

#End Region

#Region " Handle ChildTableSelectColumn "
    Public Function GetChildTableSelectColumn() As ChildTableSelectColumn
        Dim i As Integer
        For i = 1 To cColumnColl.Count
            If cColumnColl(i).ColumnType = ColumnType.ChildTableSelect Then
                Exit For
            End If
        Next
        Return cColumnColl(i)
    End Function
#End Region

#Region " Handle the external filter "
    Public Function GetExternalFilter() As DG.ExternalFilter
        Return cExternalFilter
    End Function

    Public Function AttachExternalFilter() As DG.ExternalFilter
        cExternalFilter = New DG.ExternalFilter
        Return cExternalFilter
    End Function
#End Region

#Region " Attach components "
    Public Sub AddNewButton(ByVal Right As String)
        cAttachNewButton = True
        cNewButtonRight = Right
    End Sub

    Public Sub AddAnchorObject(ByVal Key2 As String, ByVal Key3 As String, ByVal Key4 As String)
        cAnchorObject = New AnchorObject(Key2, Key3, Key4)
    End Sub

    Public Sub AddDataBoundColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New DataBoundColumn(ColumnType.Databound, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendSortReference(SortExpression)
    End Sub

    Public Sub AddDataBoundColumnSpecial(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal DisplayFldName As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New DataBoundColumnSpecial(ColumnType.DataboundSpecial, ItemName, DataFldName, HeaderText, SortExpression, DisplayFldName, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendSortReference(SortExpression)
    End Sub

    Public Sub AddCompoundDataBoundColumn(ByVal ItemName As String, ByVal DataFldName1 As String, ByVal LineBreak1 As Boolean, ByVal DataFldName2 As String, ByVal LineBreak2 As Boolean, ByVal DataFldName3 As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New CompoundDataboundColumn(ColumnType.CompoundDatabound, ItemName, DataFldName1, LineBreak1, DataFldName2, LineBreak2, DataFldName3, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendSortReference(SortExpression)
    End Sub

    Public Sub AddDateColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New DateColumn(ColumnType.Date, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendSortReference(SortExpression)
    End Sub

    'Public Sub AddFreeFormColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal Text As String)
    '    cColumnColl.Add(New FreeFormColumn(ColumnType.FreeForm, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes, Text), ItemName)
    '    AppendSortReference(SortExpression)
    'End Sub

    '        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal EnableFldName As String, ByVal DataFormatString As String, ByVal Title As String, ByVal Attributes As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal PermissionFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal DataFldName1 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(ColumnType.ChildTableSelect, ItemName, DataFldName1, Nothing, Nothing, True, PermissionFldName, Nothing, Title, Attributes, Nothing, Nothing, Nothing)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTable.ChildTableSelectColumn = ChildTableSelectColumn
    End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal PermissionFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal DataFldName1 As String, ByVal DataFldName2 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(ColumnType.ChildTableSelect, ItemName, DataFldName1, Nothing, Nothing, True, PermissionFldName, Nothing, Title, Attributes, DataFldName2, Nothing, Nothing)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTable.ChildTableSelectColumn = ChildTableSelectColumn
    End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal PermissionFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal DataFldName1 As String, ByVal DataFldName2 As String, ByVal DataFldName3 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(ColumnType.ChildTableSelect, ItemName, DataFldName1, Nothing, Nothing, True, PermissionFldName, Nothing, Title, Attributes, DataFldName2, DataFldName3, Nothing)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTable.ChildTableSelectColumn = ChildTableSelectColumn
    End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal PermissionFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal DataFldName1 As String, ByVal DataFldName2 As String, ByVal DataFldName3 As String, ByVal DataFldName4 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(ColumnType.ChildTableSelect, ItemName, DataFldName1, Nothing, Nothing, True, PermissionFldName, Nothing, Title, Attributes, DataFldName2, DataFldName3, DataFldName4)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTable.ChildTableSelectColumn = ChildTableSelectColumn
    End Sub

    Public Sub AddBooleanColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal TrueValue As String, ByVal TrueText As String, ByVal FalseText As String, ByVal Title As String, ByVal Attributes As String)
        cColumnColl.Add(New BooleanColumn(ColumnType.Boolean, ItemName, DataFldName, HeaderText, SortExpression, Visible, TrueValue, TrueText, FalseText, Title, Attributes), ItemName)
        AppendSortReference(SortExpression)
    End Sub

    Public Sub AddFreeFormColumn(ByVal ItemName As String, ByVal CellText As String, ByVal Header As String, ByVal Title As String, ByVal Visible As Boolean, ByVal Attributes As String)
        cColumnColl.Add(New FreeFormColumn(ColumnType.FreeForm, ItemName, CellText, Header, Title, Visible, Attributes), ItemName)
    End Sub

    Public Sub AddHiddenColumn(ByVal ItemName As String, ByVal DataFldName As String)
        cColumnColl.Add(New HiddenItems(ColumnType.Hidden, ItemName, DataFldName), ItemName)
        If cHiddenColumnColl Is Nothing Then
            cHiddenColumnColl = New Collection
        End If
        cHiddenColumnColl.Add(New Collection, ItemName)
    End Sub

    Public Sub AddLinkColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HRef As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal AddlParm As String, ByVal Right As String, ByVal PermissionFldName As String)
        cColumnColl.Add(New LinkColumnItems(ColumnType.Link, HRef, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes, AddlParm, Right, PermissionFldName), ItemName)
        AppendSortReference(SortExpression)
    End Sub

    Public Sub AddCheckboxToggleColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal Right As String, ByVal PermissionFldName As String, ByVal TestFld As String, ByVal TrueText As String, ByVal FalseText As String, ByVal TitleFldName As String, ByVal Attributes As String)
        If cRights Is Nothing OrElse cRights.HasThisRight(Right) Then
            cColumnColl.Add(New CheckboxToggleColumnItems(ColumnType.CheckboxToggle, ItemName, DataFldName, HeaderText, SortExpression, Visible, PermissionFldName, TitleFldName, Attributes, TestFld, TrueText, FalseText), ItemName)
            If cCheckboxToggleColl Is Nothing Then
                cCheckboxToggleColl = New Collection
            End If
            cCheckboxToggleColl.Add(New System.Text.StringBuilder, ItemName)
            AppendSortReference(SortExpression)
        End If
    End Sub

    Public Sub AddCheckboxRowSelectColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal Right As String, ByVal PermissionFldName As String, ByVal TestFld As String, ByVal TrueText As String, ByVal FalseText As String, ByVal TitleFldName As String, ByVal Attributes As String)
        If cRights Is Nothing OrElse cRights.HasThisRight(Right) Then
            cColumnColl.Add(New CheckboxRowSelectColumnItems(ColumnType.CheckboxRowSelect, ItemName, DataFldName, HeaderText, SortExpression, Visible, PermissionFldName, TitleFldName, Attributes, TestFld, TrueText, FalseText), ItemName)
            AppendSortReference(SortExpression)
        End If
    End Sub

    'Public Function GetTemplateColumn(ByVal ItemName As String, ByVal HeaderText As String, ByVal Wrap As Boolean, ByVal Visible As Boolean) As TemplateColumn
    '    cColumnColl.Add(New TemplateColumn(ItemName, HeaderText, Wrap, Visible), ItemName)
    '    Return cColumnColl(cColumnColl.Count)
    'End Function

    Public Sub AddTemplateCol(ByRef TemplateCol As TemplateColumn)
        cColumnColl.Add(TemplateCol, TemplateCol.ItemName)
        AppendSortReference(TemplateCol.SortExpression)
    End Sub

    Public Function GetCheckboxToggleColl(ByVal ItemName As String) As String
        Dim Value As String
        Value = cCheckboxToggleColl(ItemName).ToString
        Value = Value.Substring(0, Value.Length - 1)
        Return Value
    End Function

    Public Function GetColumnColl(ByVal ItemName As String) As String
        Dim i As Integer
        Dim Coll As Collection
        Coll = cHiddenColumnColl(ItemName)
        Dim Value As String = Nothing

        For i = 1 To Coll.Count
            Value &= Coll(i) & "|"
        Next
        Return Value
        Value = Value.Substring(0, Value.Length - 1)
        Return Value
    End Function

    'Public Function AttachChildTables(ByVal ItemName As String, ByVal DataFldName As String, ByVal PermissionFldName As String) As DG.ChildTablesClass
    '    cChildTables = New ChildTablesClass(ItemName, DataFldName, PermissionFldName)
    '    Return cChildTables
    'End Function

    Public Function AttachChildTable(ByVal ItemName As String) As DG.ChildTableClass
        cChildTable = New ChildTableClass(ItemName)
        Return cChildTable
    End Function
#End Region

#Region " GetText et al "
    Public Function GetHTML(ByRef dt As Data.DataTable, ByRef Request As HttpRequest, ByVal EmbeddedMessage As String) As String
        Dim sb As New System.Text.StringBuilder
        Dim OddRow As Boolean = True
        Dim RowNum As Integer
        Dim ChildText As String = Nothing
        Dim ShowSubTable As Boolean

        Try

            ' ___ Level 1 table: Defined in page

            ' ___ Level 2 table
            If cUseDefaultTableDef Then
                sb.Append(cDefaultColl(cTableDef))
            Else
                sb.Append(cTableDef)
            End If

            sb.Append(AffixMenuBand(EmbeddedMessage))

            sb.Append(AffixMenuLevel2)

            sb.Append(AffixExternalFilterBand)

            ' ___ Level 3 table
            sb.Append("<tr class=""dgh""><td><table  cellSpacing='0' cellPadding='0' width='100%' border='0'>")

            sb.Append(AffixHeaderRow)

            sb.Append(AffixInternalFilter)

            If Not IsNothing(dt) Then
                For RowNum = 0 To dt.Rows.Count - 1
                    ShowSubTable = AffixDataRow(dt, RowNum, OddRow, sb, Request)
                    If ShowSubTable Then
                        ' RaiseEvent ChildDTRequest(ChildText)
                        '   RaiseEvent ChildDTRequest(dt.Rows(RowNum)(cKeyFieldName).ToText, ChildText)
                        RaiseEvent ChildDTRequest(cChildTable.ChildTableSelectColumn.ParmColl, ChildText)
                        sb.Append("<tr><td width='20px'>&nbsp;</td><td colspan='" & cColumnColl.Count - 1 & "'>" & ChildText & "</td></tr>")
                        sb.Append("<td>&nbsp;</td></tr>")
                    Else
                        If Not cFormatAsSubTable Then
                            OddRow = Not OddRow
                        End If
                    End If
                Next

                If dt.Rows.Count > 0 AndAlso (Not IsNothing(cTotal)) Then
                    sb.Append(AffixTotalRow(OddRow))
                End If

            End If
            ' ___ Close level 3 table
            sb.Append("</table>")

            ' ___ Close level 2 table
            sb.Append("</table>")

            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG412: DG GetText. " & ex.Message, ex)
        End Try
    End Function

    Private Function AffixTotalRow(ByVal OddRow As Boolean) As String
        Dim i As Integer
        Dim ColNum As Integer
        Dim sb As New System.Text.StringBuilder
        Dim LabelInd As Boolean
        Dim TotalColl As Collection
        Dim TotalItem As Total.TotalItem

        Try

            If OddRow Then
                sb.Append("<tr class=""DGOdd"">" & vbCrLf)
            Else
                sb.Append("<tr class=""DGEven"">" & vbCrLf)
            End If

            TotalColl = cTotal.Coll
            For ColNum = 1 To cColumnColl.Count
                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                    If cColumnColl(ColNum).Visible Then
                        If Not LabelInd Then
                            sb.Append("<td " & cColumnColl(ColNum).Attributes & "><b>TOTAL</b></td>")
                            LabelInd = True
                        Else
                            For i = 1 To TotalColl.Count
                                TotalItem = TotalColl(i)
                                If TotalItem.ItemName = cColumnColl(ColNum).ItemName Then
                                    sb.Append("<td " & cColumnColl(ColNum).Attributes & "><b>" & TotalItem.Value & "</b></td>")
                                Else
                                    sb.Append("<td " & cColumnColl(ColNum).Attributes & ">&nbsp;</td>")
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            sb.Append("</tr>")
            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG430: DG AffixTotalRow. " & ex.Message, ex)
        End Try
    End Function

    Private Function AffixMenuBand(ByVal EmbeddedMessage As String) As String
        Dim i As Integer
        Dim ShowHideFilter As String = Nothing
        Dim DisplayNewButton As Boolean
        Dim DisplayEmbeddedMessage As Boolean
        Dim sbMenu As New System.Text.StringBuilder
        Dim InternalFilterExists As Boolean
        Dim MenuExists As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim DisplayGoButton As Boolean
        Dim DisplayShowHideFilter As Boolean
        Dim Width As Integer = 10

        Try

            ' ___ Note: Eliminating external filter


            ' ___ How this works
            ' If a filter exists, the menu displays the Go button.
            ' If a filter exists and is switchable , the menu displays the Show/Hide filter link.
            ' Then the menu displays the New button.
            ' Then, if the menu object exists, the menu displays the items in the menu collection.
            ' Then the menu displays the embedded message.
            ' Whether or not any menu items appear, this methods adds a 20px row.

            ' ___ GATHER THE INFORMATION

            ' ___ Filter
            If Not cInternalFilter Is Nothing Then
                InternalFilterExists = True
                DisplayGoButton = True
                If cInternalFilterOperationMode = FilterOperationModeEnum.FilterSwitchable Then
                    DisplayShowHideFilter = True
                    If cShowInternalFilter Then
                        ShowHideFilter = "Hide Filter"
                    Else
                        ShowHideFilter = "Show Filter"
                    End If
                End If
            End If

            ' ___ New button
            If cAttachNewButton Then
                If cNewButtonRight = Nothing Then
                    DisplayNewButton = True
                Else
                    If cRights Is Nothing OrElse cRights.HasThisRight(cNewButtonRight) Then
                        DisplayNewButton = True
                    End If
                End If
            End If

            If Not cMenu Is Nothing Then
                MenuExists = True
                Width = cMenu.CellWidthPercent
            End If

            If Not EmbeddedMessage = Nothing Then
                DisplayEmbeddedMessage = True
            End If

            ' ___ BUILD THE MENU STRING
            If DisplayGoButton Then
                sbMenu.Append("<td  align='center' width='" & Width & "%'  class=""DGMenuItem""><a href=""javascript:ApplyFilter()"">Go</a></td>")
                'sbMenu.Append("<td>" & Common.BuildButton("ApplyFilter", Nothing, Nothing, "Go") & "</td>")
            End If
            If DisplayShowHideFilter Then
                sbMenu.Append("<td  align='center' width='" & Width & "%'  class=""DGMenuItem""><a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a></td>")
            End If
            If DisplayNewButton Then
                sbMenu.Append("<td  align='center' width='" & Width & "%'  class=""DGMenuItem""><a href='javascript:NewRecord()'>New</a></td>")
                'sbMenu.Append("<td>" & Common.BuildButton("NewRecord", Nothing, Nothing, "New") & "</td>")
            End If



            If MenuExists Then
                For i = 1 To cMenu.Coll.count
                    If cMenu.Coll(i).IsVisible And (cRights Is Nothing OrElse cRights.HasThisRight(cMenu.Coll(i).Right)) Then
                        If cMenu.Coll(i).IsLink Then
                            sbMenu.Append("<td  align='center' width='" & Width & "%'  class=""DGMenuItem""><a href='javascript:" & cMenu.Coll(i).OnClickMethod & "()'>" & cMenu.Coll(i).Text & "</a></td>")
                        ElseIf cMenu.Coll(i).IsButton Then
                            sbMenu.Append("<td  align='center' width='" & Width & "%'  class=""DGMenuItem""><a href=""javascript:" & cMenu.Coll(i).OnClickMethod & "()"">" & cMenu.Coll(i).Text & "</a></td>")
                        ElseIf cMenu.Coll(i).IsText Then
                            sbMenu.Append(cMenu.Coll(i).Text)
                        End If
                    End If
                Next
            End If

            If DisplayEmbeddedMessage Then
                sbMenu.Append(EmbeddedMessage)
            Else
                sbMenu.Append("<td>&nbsp;</td>")
            End If


            ' ___ INSERT THE MENU STRING INTO THE GRID STRING
            sb.Append("<tr height='20px' align=""left""><td>")
            If sbMenu.Length > 0 Then
                sb.Append("<table cellspacing=""3"" cellpadding=""0"" padding=""0"" width=""100%"">")
                sb.Append("<tr align=""left"">" & sbMenu.ToString & "</tr></table>")
            Else
                sb.Append("<table  cellSpacing='0' cellPadding='0' width='100%' border='0'><tr><td>&nbsp;</td></tr></table>")
            End If
            sb.Append("</td></tr>")

            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG413: DG AffixMenuBand. " & ex.Message, ex)
        End Try
    End Function

    Private Function AffixMenuLevel2() As String
        Dim sbMenuLevel2 As New System.Text.StringBuilder

        Try

            If Not cMenuLevel2 Is Nothing AndAlso cMenuLevel2.IsVisible Then
                sbMenuLevel2.Append("<tr height=""20px"" align=""left""><td>")
                sbMenuLevel2.Append("<table cellspacing=""3"" cellpadding=""0"" padding=""0"" width=""100%"">")
                sbMenuLevel2.Append("<tr align=""left"">")

                If cMenuLevel2.Attributes = Nothing Then
                    sbMenuLevel2.Append("<td>")
                Else
                    sbMenuLevel2.Append("<td " & cMenuLevel2.Attributes & ">")
                End If
                sbMenuLevel2.Append(cMenuLevel2.Item & "</td></tr></table>")
                sbMenuLevel2.Append("</td></tr>")
            End If
            Return sbMenuLevel2.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG414: DG AffixMenuLevel2. " & ex.Message, ex)
        End Try
    End Function


    Private Function AffixHeaderRow() As String
        Dim ColNum As Integer
        Dim HeaderText As String = String.Empty
        Dim sb As New System.Text.StringBuilder

        Try

            sb.Append("<tr class=""dgh"">")

            For ColNum = 1 To cColumnColl.Count

                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then

                    If cColumnColl(ColNum).Visible Then
                        If cColumnColl(ColNum).HeaderText = Nothing Then
                            HeaderText = "&nbsp;"
                        Else
                            HeaderText = cColumnColl(ColNum).HeaderText
                        End If

                        If cColumnColl(ColNum).SortExpression = Nothing Then
                            sb.Append("<td " & cColumnColl(ColNum).Attributes & ">" & HeaderText & "</td>")
                        Else
                            sb.Append("<td " & cColumnColl(ColNum).Attributes & "><a href=""javascript:Sort('" & cColumnColl(ColNum).SortExpression & "')"">" & HeaderText & "</a></td>")
                        End If
                    End If

                End If

            Next

            sb.Append("</tr>")

            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG415: DG AffixHeaderRow. " & ex.Message, ex)
        End Try
    End Function

    Private Function REAL_AffixInternalFilter() As String
        Dim i, j As Integer
        Dim ColNum As Integer
        Dim ItemName As String
        Dim ColCount As Integer
        Dim ColumnHasFilter As Boolean
        Dim DropdownColl As Collection
        Dim SelectedValue As String
        Dim sb As New System.Text.StringBuilder
        Dim FromDate As String
        Dim ToDate As String
        Dim Dates As Object
        Dim Working As String
        Dim Value As String

        Try

            ' ___ Is there an internal filter?
            If cInternalFilterOperationMode = FilterOperationModeEnum.NoFilter Then
                Return String.Empty
            End If

            ' ___ Are there any filter items?
            If cInternalFilter.Coll.Count = 0 Then
                Return String.Empty
            End If

            ' ___ Get the column count
            For ColNum = 1 To cColumnColl.Count
                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                    If cColumnColl(ColNum).Visible Then
                        If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
                            ColCount += 1
                        End If
                    End If
                End If
            Next

            If cShowInternalFilter Then
                sb.Append("<tr class=""dgh"">")
                For ColNum = 1 To cColumnColl.Count
                    If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                        If cColumnColl(ColNum).Visible Then
                            If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then

                                ItemName = cColumnColl(ColNum).ItemName.ToUpper
                                ColumnHasFilter = False
                                For i = 1 To cInternalFilter.Coll.Count
                                    If cInternalFilter.Coll(i).ItemName.ToUpper = ItemName Then
                                        ColumnHasFilter = True
                                        Exit For
                                    End If
                                Next

                                If ColumnHasFilter Then
                                    If cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.Textbox Then
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><input style='FONT: 8pt Arial, Helvetica, sans-serif;' type='text' name='" & cInternalFilter.Coll(ItemName).CtlName & "' id='" & cInternalFilter.Coll(ItemName).CtlName & "' value=""" & cInternalFilter.Coll(ItemName).GetValue() & """ maxlength='" & cInternalFilter.Coll(ItemName).MaxLength & "' onkeypress='return SubmitOnEnterKey(event)'></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.NameTextbox Then
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><input style='FONT: 8pt Arial, Helvetica, sans-serif;' type='text' name='" & cInternalFilter.Coll(ItemName).CtlName & "' id='" & cInternalFilter.Coll(ItemName).CtlName & "' value=""" & cInternalFilter.Coll(ItemName).GetValue() & """ maxlength='" & cInternalFilter.Coll(ItemName).MaxLength & "' onkeypress='return SubmitOnEnterKey(event)'></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.Dropdown Then
                                        DropdownColl = cInternalFilter.Coll(ItemName).Coll
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><select style='FONT: 8pt Arial, Helvetica, sans-serif;' name='" & cInternalFilter.Coll(ItemName).CtlName & "' id='" & cInternalFilter.Coll(ItemName).CtlName & "' value=""" & cInternalFilter.Coll(ItemName).GetValue() & """>")
                                        SelectedValue = cInternalFilter.Coll(ItemName).GetValue()
                                        For j = 1 To DropdownColl.Count
                                            If DropdownColl(j).Value = SelectedValue Then
                                                sb.Append("<option selected value=""" & DropdownColl(j).Value & """>" & DropdownColl(j).Text & "</option>")
                                            Else
                                                sb.Append("<option value=""" & DropdownColl(j).Value & """>" & DropdownColl(j).Text & "</option>")
                                            End If
                                        Next
                                        sb.Append("</select>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.DateCtlYM Then
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><a id=""" & cInternalFilter.Coll(ItemName).EventString & """ href=""javascript:" & cInternalFilter.Coll(ItemName).EventString & "()"">" & cInternalFilter.Coll(ItemName).Caption & "</a></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.DateCtlYMD Then
                                        FromDate = Nothing
                                        ToDate = Nothing
                                        Dates = Split(cInternalFilter.Coll(ItemName).GetFilterValue(), "|")

                                        If IsNumeric(Dates(0)) AndAlso IsNumeric(Dates(1)) AndAlso IsNumeric(Dates(2)) Then
                                            Working = Dates(1).ToString & "/" & Dates(2).ToString & "/" & Dates(0).ToString
                                            FromDate = CType(Working, DateTime).ToString("MMM-dd-yyyy")
                                        End If
                                        If IsNumeric(Dates(3)) AndAlso IsNumeric(Dates(4)) AndAlso IsNumeric(Dates(5)) Then
                                            Working = Dates(4).ToString & "/" & Dates(5).ToString & "/" & Dates(3).ToString
                                            ToDate = CType(Working, DateTime).ToString("MMM-dd-yyyy")
                                        End If

                                        If FromDate = Nothing Then
                                            Value = "Get Date"
                                        Else
                                            If ToDate Is Nothing Then
                                                Value = FromDate
                                            Else
                                                Value = FromDate & "<br>" & ToDate
                                            End If
                                        End If

                                        ' Use the attributes of the filter rather than than the data.
                                        sb.Append("<td " & cInternalFilter.Coll(ItemName).Attributes & "><a id=""" & cInternalFilter.Coll(ItemName).EventString & """ href=""javascript:" & cInternalFilter.Coll(ItemName).EventString & "()"">" & Value & "</a></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).filtertype = FilterTypeEnum.CheckboxRowSelect Then
                                        sb.Append("<td>&nbsp;</td>")
                                    End If

                                Else
                                    sb.Append("<td>&nbsp;</td>")
                                End If

                            End If
                        End If
                    End If
                Next
                sb.Append("</tr>")
            End If

            sb.Append("<tr class=""dgh"">")

            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG416: DG AffixInternalFilter. " & ex.Message, ex)
        End Try
    End Function







    Private Function AffixInternalFilter() As String
        Dim i, j As Integer
        Dim ColNum As Integer
        Dim ItemName As String
        Dim ColCount As Integer
        Dim ColumnHasFilter As Boolean
        Dim DropdownColl As Collection
        Dim SelectedValue As String
        Dim sb As New System.Text.StringBuilder
        Dim FromDate As String
        Dim ToDate As String
        Dim Dates As Object
        Dim Working As String
        Dim Value As String
        Dim TagNum As Integer

        Try

            TagNum = 1

            ' ___ Is there an internal filter?
            If cInternalFilterOperationMode = FilterOperationModeEnum.NoFilter Then
                Return String.Empty
            End If

            TagNum = 2

            ' ___ Are there any filter items?
            If cInternalFilter.Coll.Count = 0 Then
                Return String.Empty
            End If

            TagNum = 3

            ' ___ Get the column count
            For ColNum = 1 To cColumnColl.Count
                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                    If cColumnColl(ColNum).Visible Then
                        If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
                            ColCount += 1
                        End If
                    End If
                End If
            Next

            TagNum = 4

            If cShowInternalFilter Then
                sb.Append("<tr class=""dgh"">")
                For ColNum = 1 To cColumnColl.Count
                    If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then

                        TagNum = 5

                        If cColumnColl(ColNum).Visible Then

                            TagNum = 6

                            If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then

                                TagNum = 7

                                ItemName = cColumnColl(ColNum).ItemName.ToUpper


                                TagNum = 8

                                ColumnHasFilter = False
                                For i = 1 To cInternalFilter.Coll.Count
                                    If cInternalFilter.Coll(i).ItemName.ToUpper = ItemName Then
                                        ColumnHasFilter = True
                                        Exit For
                                    End If
                                Next

                                TagNum = 9


                                If ColumnHasFilter Then
                                    If cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.Textbox Then
                                        TagNum = 10

                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><input style='FONT: 8pt Arial, Helvetica, sans-serif;' type='text' name='" & cInternalFilter.Coll(ItemName).CtlName & "' id='" & cInternalFilter.Coll(ItemName).CtlName & "' value=""" & cInternalFilter.Coll(ItemName).GetValue() & """ maxlength='" & cInternalFilter.Coll(ItemName).MaxLength & "' onkeypress='return SubmitOnEnterKey(event)'></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.NameTextbox Then
                                        TagNum = 11

                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><input style='FONT: 8pt Arial, Helvetica, sans-serif;' type='text' name='" & cInternalFilter.Coll(ItemName).CtlName & "' id='" & cInternalFilter.Coll(ItemName).CtlName & "' value=""" & cInternalFilter.Coll(ItemName).GetValue() & """ maxlength='" & cInternalFilter.Coll(ItemName).MaxLength & "' onkeypress='return SubmitOnEnterKey(event)'></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.Dropdown Then
                                        TagNum = 12

                                        DropdownColl = cInternalFilter.Coll(ItemName).Coll
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><select style='FONT: 8pt Arial, Helvetica, sans-serif;' name='" & cInternalFilter.Coll(ItemName).CtlName & "' id='" & cInternalFilter.Coll(ItemName).CtlName & "' value=""" & cInternalFilter.Coll(ItemName).GetValue() & """>")
                                        SelectedValue = cInternalFilter.Coll(ItemName).GetValue()


                                        TagNum = 13

                                        For j = 1 To DropdownColl.Count
                                            If DropdownColl(j).Value = SelectedValue Then
                                                sb.Append("<option selected value=""" & DropdownColl(j).Value & """>" & DropdownColl(j).Text & "</option>")
                                            Else
                                                sb.Append("<option value=""" & DropdownColl(j).Value & """>" & DropdownColl(j).Text & "</option>")
                                            End If
                                        Next
                                        sb.Append("</select>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.DateCtlYM Then
                                        TagNum = 14

                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><a id=""" & cInternalFilter.Coll(ItemName).EventString & """ href=""javascript:" & cInternalFilter.Coll(ItemName).EventString & "()"">" & cInternalFilter.Coll(ItemName).Caption & "</a></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).FilterType = FilterTypeEnum.DateCtlYMD Then

                                        TagNum = 15

                                        FromDate = Nothing
                                        ToDate = Nothing
                                        Dates = Split(cInternalFilter.Coll(ItemName).GetFilterValue(), "|")


                                        TagNum = 16

                                        If IsNumeric(Dates(0)) AndAlso IsNumeric(Dates(1)) AndAlso IsNumeric(Dates(2)) Then
                                            Working = Dates(1).ToString & "/" & Dates(2).ToString & "/" & Dates(0).ToString
                                            FromDate = CType(Working, DateTime).ToString("MMM-dd-yyyy")
                                        End If


                                        TagNum = 17

                                        Try

                                            If IsNumeric(Dates(3)) AndAlso IsNumeric(Dates(4)) AndAlso IsNumeric(Dates(5)) Then
                                                Working = Dates(4).ToString & "/" & Dates(5).ToString & "/" & Dates(3).ToString
                                                ToDate = CType(Working, DateTime).ToString("MMM-dd-yyyy")
                                            End If

                                        Catch ex As Exception
                                            Value = "Get Date"
                                            'Common.SendEmail("rbluestein@benefitvision.com", "automail@benfitvision.com", Nothing, "Tag17", cInternalFilter.Coll(ItemName).GetFilterValue())
                                            Dim k As Integer
                                            Dim sb17 As New System.Text.StringBuilder
                                            Dim cEnviro As Enviro
                                            Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
                                            cEnviro = SessionObj("Enviro")

                                            For k = 1 To cEnviro.LogColl17.Count
                                                sb17.Append(cEnviro.LogColl17(k) & Chr(10))
                                            Next

                                            Common.SendEmail("rbluestein@benefitvision.com", "automail@benefitvision.com", Nothing, "Tag17", sb17.ToString)
                                        End Try



                                        Try

                                            TagNum = 18

                                            If FromDate = Nothing Then
                                                Value = "Get Date"
                                            Else
                                                If ToDate Is Nothing Then
                                                    Value = FromDate
                                                Else
                                                    Value = FromDate & "<br>" & ToDate
                                                End If
                                            End If

                                        Catch ex As Exception
                                            Value = "Get Date"
                                            Common.SendEmail("rbluestein@benefitvision.com", "automail@benefitvision.com", Nothing, "Tag18", cInternalFilter.Coll(ItemName).GetFilterValue())
                                        End Try


                                        TagNum = 19

                                        ' Use the attributes of the filter rather than than the data.
                                        sb.Append("<td " & cInternalFilter.Coll(ItemName).Attributes & "><a id=""" & cInternalFilter.Coll(ItemName).EventString & """ href=""javascript:" & cInternalFilter.Coll(ItemName).EventString & "()"">" & Value & "</a></td>")

                                    ElseIf cInternalFilter.Coll(ItemName).filtertype = FilterTypeEnum.CheckboxRowSelect Then
                                            TagNum = 20

                                            sb.Append("<td>&nbsp;</td>")
                                    End If

                                Else
                                    TagNum = 21

                                    sb.Append("<td>&nbsp;</td>")
                                End If

                            End If
                        End If
                    End If
                Next
                sb.Append("</tr>")
            End If

            TagNum = 22

            sb.Append("<tr class=""dgh"">")

            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG416: DG AffixInternalFilter. ColNum: " & ColNum.ToString & ", TagNum: " & TagNum.ToString & ".   " & ex.Message, ex)
        End Try
    End Function









    Private Function AffixExternalFilterBand() As String
        Dim i, j As Integer
        Dim ColNum As Integer
        Dim ItemName As String = Nothing
        Dim ColCount As Integer
        Dim ColumnHasFilter As Boolean
        Dim DropdownColl As Collection
        Dim SelectedValue As String
        Dim sb As New System.Text.StringBuilder

        Try

            ' ___ Exit if no external filter
            If cExternalFilter Is Nothing Then
                Return String.Empty
            End If

            ' ___ Exit if there are not any filter items
            If cExternalFilter.Coll.Count = 0 Then
                Return String.Empty
            End If

            ' ___ Get the column count
            For ColNum = 1 To cColumnColl.Count
                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                    If cColumnColl(ColNum).Visible Then
                        If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
                            ColCount += 1
                        End If
                    End If
                End If
            Next

            sb.Append("<table  cellSpacing='0' cellPadding='0' width='100%' border='0'>")
            sb.Append("<tr class=""dgh"" height=""30px"" align=""bottom""><td align=""left"">")

            For ColNum = 1 To cColumnColl.Count
                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                    If cColumnColl(ColNum).Visible Then
                        If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then

                            ColumnHasFilter = False
                            For i = 1 To cExternalFilter.Coll.count
                                If cExternalFilter.Coll(i).Position = ColNum Then
                                    ColumnHasFilter = True
                                    ItemName = cExternalFilter.Coll(i).ItemName
                                    Exit For
                                End If
                            Next

                            If ColumnHasFilter Then
                                If cExternalFilter.Coll(ItemName).IsTextbox Then
                                    sb.Append("<input style='FONT: 8pt Arial, Helvetica, sans-serif;' type='text' name='" & cExternalFilter.Coll(ItemName).CtlName & "' id='" & cExternalFilter.Coll(ItemName).CtlName & "' value=""" & cExternalFilter.Coll(ItemName).GetValue() & """ maxlength='" & cExternalFilter.Coll(ItemName).MaxLength & "' onkeypress='return SubmitOnEnterKey(event)'>&nbsp;&nbsp;&nbsp;&nbsp;")


                                ElseIf cExternalFilter.Coll(ItemName).IsDropdown Then
                                    DropdownColl = cExternalFilter.Coll(ItemName).Coll
                                    sb.Append("<select style='FONT: 8pt Arial, Helvetica, sans-serif;' name='" & cExternalFilter.Coll(ItemName).CtlName & "' id='" & cExternalFilter.Coll(ItemName).CtlName & "'")
                                    If cExternalFilter.Coll(ItemName).EventString = Nothing Then
                                        sb.Append(">")
                                    Else
                                        sb.Append(" " & cExternalFilter.Coll(ItemName).EventString & ">")
                                    End If

                                    For j = 1 To DropdownColl.Count
                                        SelectedValue = cExternalFilter.Coll(ItemName).GetValue()
                                        If DropdownColl(j).Value = SelectedValue Then
                                            sb.Append("<option selected value=""" & DropdownColl(j).Value & """>" & DropdownColl(j).Text & "</option>")
                                        Else
                                            sb.Append("<option value=""" & DropdownColl(j).Value & """>" & DropdownColl(j).Text & "</option>")
                                        End If
                                    Next
                                    sb.Append("</select>&nbsp;&nbsp;&nbsp;&nbsp;")

                                ElseIf cExternalFilter.Coll(ItemName).IsLink Then
                                    sb.Append("<a href=""javascript:" & cExternalFilter.Coll(ItemName).EventString & "()"">" & cExternalFilter.Coll(ItemName).Text & "</a>&nbsp;&nbsp;&nbsp;&nbsp;")
                                End If

                            Else
                                sb.Append("&nbsp;")
                            End If

                        End If
                    End If
                End If
            Next
            sb.Append("</td></tr></table>")

            Return sb.ToString

        Catch ex As Exception
            Throw New Exception("Error #DG417: DG AffixExternalMenuBand. " & ex.Message, ex)
        End Try
    End Function

    Private Function AffixDataRow(ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByVal OddRow As Boolean, ByRef sb As System.Text.StringBuilder, ByRef Request As HttpRequest) As Boolean
        Dim ColNum As Integer
        Dim ShowSubTable As Boolean

        Try

            If OddRow Then
                sb.Append("<tr class=""DGOdd"">" & vbCrLf)
            Else
                sb.Append("<tr class=""DGEven"">" & vbCrLf)
            End If

            ' ___ Handle Anchor
            If Not cAnchorObject Is Nothing Then
                HandleAnchor(dt, RowNum, sb)
            End If

            For ColNum = 1 To cColumnColl.Count

                If cColumnColl(ColNum).ColumnType = DG.ColumnType.Hidden Then
                    HandleHiddenColumn(ColNum, dt, RowNum, sb)
                Else
                    If cColumnColl(ColNum).Visible Then
                        Select Case cColumnColl(ColNum).ColumnType
                            Case DG.ColumnType.Databound, DG.ColumnType.DataboundSpecial
                                HandleDataboundColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.CompoundDatabound
                                HandleCompoundDataboundColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.Link
                                HandleLinkColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.CheckboxRowSelect
                                HandleCheckboxRowSelectColumn(ColNum, dt, RowNum, sb, Request)
                            Case DG.ColumnType.CheckboxToggle
                                HandleCheckboxToggleColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.Template
                                HandleTemplateColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.ChildTableSelect
                                ShowSubTable = HandleChildTableSelectColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.Boolean
                                HandleBooleanColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.Date
                                HandleDateColumn(ColNum, dt, RowNum, sb)
                            Case DG.ColumnType.FreeForm
                                HandleFreeFormColumn(ColNum, dt, RowNum, sb)
                        End Select
                    End If
                End If

            Next
            sb.Append("</tr>")

            Return ShowSubTable

        Catch ex As Exception
            Throw New Exception("Error #DG418: DG AffixDataRow. " & ex.Message, ex)
        End Try
    End Function

    Private Sub HandleAnchor(ByRef dt As DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        '   sb.Append("<a name='" & dt.Rows(RowNum)(cKeyFieldName).ToText & "'></a>")
        sb.Append("<a name='" & dt.Rows(RowNum)(cKeyFieldName).ToText)

        If Not cAnchorObject.Key2 = Nothing Then
            sb.Append(cAnchorObject.Key2)
        End If
        If Not cAnchorObject.Key3 = Nothing Then
            sb.Append(cAnchorObject.Key3)
        End If
        If Not cAnchorObject.Key4 = Nothing Then
            sb.Append(cAnchorObject.Key4)
        End If
        sb.Append("'></a>")
    End Sub

    Private Sub HandleCheckboxToggleColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String
        Dim CheckboxLabel As String = Nothing

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).TitleFldName <> Nothing Then
                ColumnAttributes &= " title='" & dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName).ToText & "' "
            End If

            If dt.Rows(RowNum)(cColumnColl(ColNum).PermissionFldName).ToText = "1" Then
                If cColumnColl(ColNum).TestFld <> Nothing Then
                    If dt.Rows(RowNum)(cColumnColl(ColNum).TestFld).ToText = "1" Then
                        CheckboxLabel = cColumnColl(ColNum).TrueText
                    Else
                        CheckboxLabel = cColumnColl(ColNum).FalseText
                    End If
                End If

                Value = "<input type='checkbox' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText & "' value='on'>&nbsp;" & CheckboxLabel

            Else
                Value = "&nbsp;"
            End If

            ' ___ This is provided as an alternative way of iterating through the checkbox items.
            'cCheckboxToggleColl(cColumnColl(ColNum).ItemName).append(dt.Rows(RowNum)(cKeyFieldName).ToText & "~" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText & "|")
            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG419: DG HandleCheckboxToggleColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleCheckboxRowSelectColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder, ByVal Request As HttpRequest)
        Dim Value As String
        Dim ColumnAttributes As String
        Dim CheckboxLabel As String = Nothing
        Dim ItemName As String

        Try

            ' ___ Get ItemName
            ItemName = cColumnColl(ColNum).ItemName

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).TitleFldName <> Nothing Then
                ColumnAttributes &= " title='" & dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName).ToText & "' "
            End If

            If dt.Rows(RowNum)(cColumnColl(ColNum).PermissionFldName).ToText = "1" Then
                If cInternalFilter.Coll(ItemName).ClearAll Then
                    Value = "<input type='checkbox' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText & "'>&nbsp;" & CheckboxLabel
                Else
                    If Request.Form(cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText) = Nothing Then
                        Value = "<input type='checkbox' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText & "'>&nbsp;" & CheckboxLabel
                    Else
                        Value = "<input type='checkbox' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText & "' checked>&nbsp;" & CheckboxLabel
                    End If
                End If
            Else
                Value = "&nbsp;"
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG429: DG HandleCheckboxRowSelectColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleLinkColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String = String.Empty
        Dim ColumnAttributes As String
        Dim Box As Object
        Dim DoNotDisplay As Boolean
        Dim DisplayLinkAsLink As Boolean
        Dim DisplayLinkAsText As Boolean
        Dim DataFldRawValue As String = String.Empty
        Dim DisplayValue As String = String.Empty

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).TitleFldName <> Nothing Then
                ColumnAttributes &= " title='" & dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName).ToText & "' "
            End If

            ' HRef = "javascript:fnToggle"
            ' DataFldName "ReadyForPrint|1"       Display as link
            ' DataFldName "ReadyForPrint|0"       Display as text
            ' DataFldName "ReadyForPrint"          Display as link

            If IsDBNull(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).Value) Then
                DoNotDisplay = True
            Else
                DataFldRawValue = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText
                If InStr(DataFldRawValue, "|") = 0 Then
                    DisplayLinkAsText = False
                    DisplayLinkAsLink = True
                    DisplayValue = DataFldRawValue
                Else
                    Box = Split(DataFldRawValue, "|")
                    If Box(1) = "0" Then
                        DisplayLinkAsText = True
                        DisplayLinkAsLink = False
                        DisplayValue = Box(0)
                    Else
                        DisplayLinkAsText = False
                        DisplayLinkAsLink = True
                        DisplayValue = Box(0)
                    End If
                End If
            End If

            If DisplayLinkAsLink Then
                Value = cColumnColl(ColNum).Href
                If cColumnColl(ColNum).AddlParm = Nothing Then
                    Value = "<a href=""javascript:" & Value & "('" & dt.Rows(RowNum)(cKeyFieldName).ToText & "')"">" & DisplayValue & "</a>"
                Else
                    Value = "<a href=""javascript:" & Value & "('" & dt.Rows(RowNum)(cKeyFieldName).ToText & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).AddlParm).ToText & "')"">" & DisplayValue & "</a>"
                End If
            ElseIf DisplayLinkAsText Then
                Value = DisplayValue
                If Value.Length = 0 Then
                    Value = "&nbsp;"
                End If
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG420: DG HandleLinkColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleBooleanColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).Title <> Nothing Then
                ColumnAttributes &= " title='" & Common.StrInHandler(cColumnColl(ColNum).Title) & "' "
            End If

            ' ___ Determine the value
            ' If Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)) = cColumnColl(ColNum).TrueValue Then
            If dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText = cColumnColl(ColNum).TrueValue Then
                Value = cColumnColl(ColNum).TrueText
            Else
                Value = cColumnColl(ColNum).FalseText
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG421: DG HandleBooleanColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleDateColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim OrigValue As DateTime
        Dim Value As String
        Dim ColumnAttributes As String

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).Title <> Nothing Then
                ColumnAttributes &= " title='" & Common.StrInHandler(cColumnColl(ColNum).Title) & "' "
            End If

            ' ___ Get and format value
            Value = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText
            If Value = String.Empty Then
                Value = "&nbsp;"
            Else
                OrigValue = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText
                If cColumnColl(ColNum).DataFormatString = Nothing Then
                    Value = OrigValue.ToString
                ElseIf cColumnColl(ColNum).DataFormatString = "SpecialFormat2" Then
                    Value = Common.FormatDate(OrigValue, FormatDateEnum.SpecialFormat2)
                ElseIf cColumnColl(ColNum).DataFormatString = "SpecialFormat3" Then
                    Value = Common.FormatDate(OrigValue, FormatDateEnum.SpecialFormat3)
                Else
                    Value = OrigValue.ToString(cColumnColl(ColNum).DataFormatString)
                End If
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG422: DG HandleDateColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleFreeFormColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).Title <> Nothing Then
                ColumnAttributes &= " title='" & dt.Rows(RowNum)(cColumnColl(ColNum).Title).ToText & "' "
            End If

            ' ___ Get the value
            Value = Common.StrInHandler(cColumnColl(ColNum).CellText)

            ' ___ Format value
            If Value = Nothing Then
                Value = "&nbsp;"
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG423: DG HandleFreeFormColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleDataboundColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String = Nothing
        Dim ColumnAttributes As String
        Dim MethodName As String

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).TitleFldName <> Nothing Then
                ColumnAttributes &= " title='" & dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName).ToText & "' "
            End If

            '' ___ Get the value
            'Value = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText

            Select Case cColumnColl(ColNum).ColumnType
                Case DG.ColumnType.Databound
                    Value = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText
                Case DG.ColumnType.DataboundSpecial
                    Value = dt.Rows(RowNum)(cColumnColl(ColNum).DisplayFldName).ToText
            End Select


            ' ___ Format value
            If Value.Length = 0 Then
                Value = "&nbsp;"

            Else

                If Not cColumnColl(ColNum).DataFormatString = Nothing Then
                    If cColumnColl(ColNum).DataFormatString.substring(0, 1) = "|" Then
                        MethodName = cColumnColl(ColNum).DataFormatString.Substring(1)
                        Select Case MethodName
                            Case "ToProper"
                                Value = Common.ToProper(Value)
                        End Select
                    Else
                        Value = CDbl(Value).ToString(cColumnColl(ColNum).DataFormatString)
                    End If
                End If

            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG424: DG HandleDataboundColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleCompoundDataboundColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        Try

            ' __ Handle column attributes and tooltip
            ColumnAttributes = cColumnColl(ColNum).Attributes
            If cColumnColl(ColNum).TitleFldName <> Nothing Then
                ColumnAttributes &= " title='" & dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName).ToText & "' "
            End If

            ' ___ Get the value
            Value = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText

            ' ___ Format value
            If Value.Length = 0 Then
                Value = "&nbsp;"
            Else
                If Not cColumnColl(ColNum).DataFormatString = Nothing Then
                    Value = CDbl(Value).ToString(cColumnColl(ColNum).DataFormatString)
                End If
            End If

            If cColumnColl(ColNum).DataFldName2 <> Nothing AndAlso (Not IsDBNull(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName2).Value)) Then
                If cColumnColl(ColNum).LineBreak1 Then
                    Value &= "<br>" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName2).ToText
                Else
                    Value &= " " & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName2).ToText
                End If
            End If

            If cColumnColl(ColNum).DataFldName3 <> Nothing AndAlso (Not IsDBNull(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName3).Value)) Then
                If cColumnColl(ColNum).LineBreak2 Then
                    Value &= "<br>" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName3).ToText
                Else
                    Value &= " " & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName3).ToText
                End If
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG425: DG HandleCompoundDataboundColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleHiddenColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String

        Try

            Value = "<input type='hidden' id='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText & "' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName).ToText & "' value=""" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText & """>"
            cHiddenColumnColl(cColumnColl(ColNum).ItemName).add(dt.Rows(RowNum)(cKeyFieldName).ToText & "~" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName).ToText)
            sb.Append(Value)

        Catch ex As Exception
            Throw New Exception("Error #DG426: DG HandleHiddenColumn. " & ex.Message, ex)
        End Try
    End Sub

    Private Function HandleChildTableSelectColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder) As Boolean
        Dim Value As String = Nothing
        Dim ColumnAttributes As String
        Dim i As Integer
        Dim ShowSubTable As Boolean
        Dim ShowPlusMinus As Boolean
        Dim FldName As String
        Dim ChildTableSelectColumn As DG.ChildTableSelectColumn
        Dim ParmColl As Collection

        Try

            ChildTableSelectColumn = cColumnColl(ColNum)
            ParmColl = ChildTableSelectColumn.ParmColl

            ' __ Handle column attributes and tooltip
            ColumnAttributes = "width= '20px' "
            ColumnAttributes &= ChildTableSelectColumn.Attributes

            ' ___ With proper permissions show the plus or minus sign
            If (Not ChildTableSelectColumn.PermissionFldName = Nothing) AndAlso (dt.Rows(RowNum)(ChildTableSelectColumn.PermissionFldName).ToText = 0) Then
                ShowPlusMinus = False
            Else
                ShowPlusMinus = True
            End If


            If ShowPlusMinus Then

                If ChildTableSelectColumn.Title <> Nothing Then
                    ColumnAttributes &= " title='" & ChildTableSelectColumn.Title & "' "
                End If

                ' ___ Determine whether to display the subtable.
                ShowSubTable = True
                For i = 1 To ParmColl.Count
                    FldName = ParmColl(i).FldName
                    If (Not dt.Rows(RowNum)(FldName).ToText = ParmColl(i).Value) Then
                        ShowSubTable = False
                        Exit For
                    End If
                Next

            End If


            If ShowPlusMinus Then

                If ShowSubTable Then

                    ' ___ Show minus sign
                    Value = "<a href=""javascript:ShowHideSubTable('ReqSubTblClose', '')""><strong>&nbsp;&nbsp;-&nbsp;&nbsp;</strong></a>"

                Else

                    ' ___ Show plus sign and build the link
                    Select Case ParmColl.Count
                        Case 1
                            Value = "<a href=""javascript:ShowHideSubTable('ReqSubTblOpen', '" & dt.Rows(RowNum)(ParmColl(1).FldName).ToJSParm & "')""><strong>&nbsp;&nbsp;+&nbsp;&nbsp;</strong></a>"
                        Case 2
                            Value = "<a href=""javascript:ShowHideSubTable('ReqSubTblOpen', '" & dt.Rows(RowNum)(ParmColl(1).FldName).ToJSParm & "', '" & dt.Rows(RowNum)(ParmColl(2).FldName).ToJSParm & "')""><strong>&nbsp;&nbsp;+&nbsp;&nbsp;</strong></a>"
                        Case 3
                            Value = "<a href=""javascript:ShowHideSubTable('ReqSubTblOpen', '" & dt.Rows(RowNum)(ParmColl(1).FldName).ToJSParm & "', '" & dt.Rows(RowNum)(ParmColl(2).FldName).ToJSParm & "', '" & dt.Rows(RowNum)(ParmColl(3).FldName).ToJSParm & "')""><strong>&nbsp;&nbsp;+&nbsp;&nbsp;</strong></a>"
                        Case 4
                            Value = "<a href=""javascript:ShowHideSubTable('ReqSubTblOpen', '" & dt.Rows(RowNum)(ParmColl(1).FldName).ToJSParm & "', '" & dt.Rows(RowNum)(ParmColl(2).FldName).ToJSParm & "', '" & dt.Rows(RowNum)(ParmColl(3).FldName).ToJSParm & "', '" & dt.Rows(RowNum)(ParmColl(4).FldName).ToJSParm & "')""><strong>&nbsp;&nbsp;+&nbsp;&nbsp;</strong></a>"
                    End Select

                End If

            Else

                Value = "&nbsp;"
            End If

            sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
            Return ShowSubTable

        Catch ex As Exception
            Throw New Exception("Error #DG427: DG HandleChildTableSelectColumn. " & ex.Message, ex)
        End Try
    End Function

    Private Sub HandleTemplateColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim TemplateCol As DG.TemplateColumn
        Dim IconColumnItem As DG.TemplateColumn.IconColumnItem
        Dim IsVisible As Boolean
        Dim i, j As Integer
        Dim PassPermissionFldTest As Boolean
        Dim PassRightTest As Boolean
        Dim DataFldRawValue As String = String.Empty
        Dim DisplayValue As String = String.Empty
        Dim LinkValue As String = String.Empty
        Dim ParmColl As New Collection

        Try

            TemplateCol = cColumnColl(ColNum)

            ' __ Handle column attributes and tooltip
            sb.Append("<td ")
            If cColumnColl(ColNum).Attributes <> Nothing Then
                sb.Append(cColumnColl(ColNum).Attributes)
            End If
            If Not TemplateCol.Wrap Then
                sb.Append(" nowrap ")
            End If
            sb.Append(">")

            For j = 1 To TemplateCol.Count
                IconColumnItem = TemplateCol.Coll(j)
                ParmColl = New Collection
                IsVisible = False
                PassPermissionFldTest = False
                PassRightTest = False

                ' ___ Permission field test
                If IconColumnItem.PermissionFldName = Nothing Then
                    PassPermissionFldTest = True
                Else
                    If (Not IsDBNull(dt.Rows(RowNum)(IconColumnItem.PermissionFldName).Value)) AndAlso (dt.Rows(RowNum)(IconColumnItem.PermissionFldName).ToText = "1") Then
                        PassPermissionFldTest = True
                    End If
                End If

                ' ___ Rights test
                If IconColumnItem.Right = Nothing Then
                    PassRightTest = True
                Else
                    If cRights Is Nothing OrElse cRights.HasThisRight(IconColumnItem.Right) Then
                        PassRightTest = True
                    End If
                End If

                If PassPermissionFldTest AndAlso PassRightTest Then
                    IsVisible = True
                End If

                If IsVisible Then


                    ' ___ Build anchor tag
                    ParmColl.Add(IconColumnItem.DataFldName)
                    If IconColumnItem.Parm2 <> Nothing Then
                        ParmColl.Add(IconColumnItem.Parm2)
                    End If
                    If IconColumnItem.Parm3 <> Nothing Then
                        ParmColl.Add(IconColumnItem.Parm3)
                    End If
                    If IconColumnItem.Parm4 <> Nothing Then
                        ParmColl.Add(IconColumnItem.Parm4)
                    End If

                    sb.Append("<a ")
                    sb.Append("id="""" ")
                    sb.Append("onclick=""" & IconColumnItem.OnClickMethod & "(")

                    For i = 1 To ParmColl.Count
                        If i = 1 Then
                            sb.Append("'" & dt.Rows(RowNum)(cKeyFieldName).ToJSParm & "'")
                        Else
                            If ParmColl(i).Length > 2 AndAlso ParmColl(i).Substring(0, 1) = "[" AndAlso ParmColl(i).Substring(ParmColl(i).Length - 1) = "]" Then
                                sb.Append("'" & ParmColl(i).Substring(1, ParmColl(i).Length - 2) & "'")
                            Else
                                sb.Append("'" & dt.Rows(RowNum)(ParmColl(i)).ToText & "'")
                            End If
                        End If
                        If i < ParmColl.Count Then
                            sb.Append(", ")
                        End If
                    Next
                    sb.Append(")"">")


                    If IconColumnItem.UseDefaultImage Then
                        sb.Append(cColumnColl(ColNum).DefaultImageColl(IconColumnItem.DefaultImage))
                    Else
                        sb.Append("<img src='" & IconColumnItem.ImagePath & "' ")
                        sb.Append(IconColumnItem.ImageAttributes)
                    End If

                    If IconColumnItem.Title <> Nothing Then
                        sb.Append(" title=""" & IconColumnItem.Title & """")
                    End If

                    sb.Append(">")

                    sb.Append("</a>")

                End If

            Next
            sb.Append("</td>")

        Catch ex As Exception
            Throw New Exception("Error #DG428: DG HandleTemplateColumn. " & ex.Message, ex)
        End Try
    End Sub
#End Region

#Region " Column classes "
    Public Class FreeFormColumn
        Inherits DataBoundColumn
        Private cCellText As String
        Private cTitle As String
        Public ReadOnly Property Title() As String
            Get
                Return cTitle
            End Get
        End Property

        Public ReadOnly Property CellText() As String
            Get
                Return cCellText
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal CellText As String, ByVal HeaderText As String, ByVal Title As String, ByVal Visible As Boolean, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, Nothing, HeaderText, Nothing, Visible, Nothing, Nothing, Attributes)
            cCellText = CellText
            cTitle = Title
        End Sub
    End Class

    'Public Class ChildTableSelectColumn
    '    Inherits DataBoundColumn
    '    Private cTitle As String
    '    Private cPermissionFldName As String
    '    Private cParmColl As Collection

    '    Public ReadOnly Property ParmColl()
    '        Get
    '            Return cParmColl
    '        End Get
    '    End Property
    '    Public ReadOnly Property DataFldColl() As Collection
    '        Get
    '            Return cDataFldColl
    '        End Get
    '    End Property
    '    Public ReadOnly Property PermissionFldName() As String
    '        Get
    '            Return cPermissionFldName
    '        End Get
    '    End Property
    '    Public ReadOnly Property Title() As String
    '        Get
    '            Return cTitle
    '        End Get
    '    End Property

    '    Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal PermissionFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal DataFldName1 As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
    '        MyBase.New(ColumnType, ItemName, DataFldName1, Nothing, Nothing, True, Nothing, Nothing, Attributes)
    '        cPermissionFldName = PermissionFldName
    '        cTitle = Title
    '        cParmColl = New Collection
    '        cParmColl.Add(New Items("DataFldName1", DataFldName1), "DataFldName1")
    '        If Not DataFldName2 = Nothing Then
    '            cParmColl.Add(New Items("DataFldName2", DataFldName2), "DataFldName2")
    '        End If
    '        If Not DataFldName3 = Nothing Then
    '            cParmColl.Add(New Items("DataFldName3", DataFldName3), "DataFldName3")
    '        End If
    '        If Not DataFldName4 = Nothing Then
    '            cParmColl.Add(New Items("DataFldName4", DataFldName4), "DataFldName4")
    '        End If
    '    End Sub

    '    Public Class Items
    '        Private cItemName As String
    '        Private cFldName As String
    '        Private cValue As String
    '        Public ReadOnly Property ItemName() As String
    '            Get
    '                Return cItemName
    '            End Get
    '        End Property
    '        Public ReadOnly Property FldName() As String
    '            Get
    '                Return cFldName
    '            End Get
    '        End Property
    '        Public Property Value() As String
    '            Get
    '                Return cValue
    '            End Get
    '            Set(ByVal Value As String)
    '                cValue = Value
    '            End Set
    '        End Property

    '        Public Sub New(ByVal ItemName As String, ByVal FldName As String)
    '            cItemName = ItemName
    '            cFldName = FldName
    '        End Sub
    '    End Class
    'End Class






    Public Class ChildTableSelectColumn
        Inherits DataBoundColumn
        Private cTitle As String
        Private cPermissionFldName As String
        Private cParmColl As Collection
        ' Private cChildTables As DG.ChildTablesClass

        Public ReadOnly Property ParmColl() As Collection
            Get
                Return cParmColl
            End Get
        End Property
        Public ReadOnly Property PermissionFldName() As String
            Get
                Return cPermissionFldName
            End Get
        End Property
        Public ReadOnly Property Title() As String
            Get
                Return cTitle
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal PermissionFldName As String, ByVal DataFormatString As String, ByVal Title As String, ByVal Attributes As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)

            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, Nothing, Attributes)

            cPermissionFldName = PermissionFldName
            cTitle = Title
            'cChildTables = ChildTables
            cParmColl = New Collection
            cParmColl.Add(New Items("DataFldName", DataFldName), "DataFldName")
            If Not Parm2 = Nothing Then
                cParmColl.Add(New Items("Parm2", Parm2), "Parm2")
            End If
            If Not Parm3 = Nothing Then
                cParmColl.Add(New Items("Parm3", Parm3), "Parm3")
            End If
            If Not Parm4 = Nothing Then
                cParmColl.Add(New Items("Parm4", Parm4), "Parm4")
            End If
        End Sub

        Public Class Items
            Private cItemName As String
            Private cFldName As String
            Private cValue As String
            Public ReadOnly Property ItemName() As String
                Get
                    Return cItemName
                End Get
            End Property
            Public ReadOnly Property FldName() As String
                Get
                    Return cFldName
                End Get
            End Property
            Public Property Value() As String
                Get
                    Return cValue
                End Get
                Set(ByVal Value As String)
                    cValue = Value
                End Set
            End Property

            Public Sub New(ByVal ItemName As String, ByVal FldName As String)
                cItemName = ItemName
                cFldName = FldName
            End Sub
        End Class
    End Class

    Public Class TemplateColumn
        Private cColl As New Collection
        Private cHeaderText As String
        Private cSortExpression As String
        Private cWrap As Boolean
        Private cVisible As Boolean
        Private cDefaultImageColl As New Collection
        Private cItemName As String
        Private cAttributes As String

        Public ReadOnly Property Coll() As Collection
            Get
                Return cColl
            End Get
        End Property
        Public ReadOnly Property Attributes() As String
            Get
                Return cAttributes
            End Get
        End Property
        Public ReadOnly Property ItemName() As String
            Get
                Return cItemName
            End Get
        End Property
        Public ReadOnly Property DefaultImageColl() As Collection
            Get
                Return cDefaultImageColl
            End Get
        End Property
        Public ReadOnly Property ColumnType() As DG.ColumnType
            Get
                Return DG.ColumnType.Template
            End Get
        End Property
        Public ReadOnly Property HeaderText() As String
            Get
                Return cHeaderText
            End Get
        End Property
        Public ReadOnly Property SortExpression() As String
            Get
                Return cSortExpression
            End Get
        End Property
        Public ReadOnly Property Wrap() As Boolean
            Get
                Return cWrap
            End Get
        End Property
        Public ReadOnly Property Visible() As Boolean
            Get
                Return cVisible
            End Get
        End Property
        Public ReadOnly Property Count() As Integer
            Get
                Return cColl.Count
            End Get
        End Property

        Public Sub New(ByVal ItemName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Wrap As Boolean, ByVal Attributes As String, ByVal Visible As Boolean)
            cItemName = ItemName
            cHeaderText = HeaderText
            cSortExpression = SortExpression
            cWrap = Wrap
            cVisible = Visible
            cAttributes = Attributes
            cDefaultImageColl.Add("<img src='img/view.gif'  border='0' width='14' height='14' vspace='0' hspace='0'", "StandardView")
            cDefaultImageColl.Add("<img src='img/delete.gif'  border='0' width='14' height='14' vspace='0' hspace='0'", "StandardDelete")
            cDefaultImageColl.Add("<img src='img/Delete_Strong.png'  border='0' width='14' height='14' vspace='0' hspace='0'", "DeleteStrong")
            cDefaultImageColl.Add("<img src='img/Clip.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardClip")
            cDefaultImageColl.Add("<img src='img/Magnify.png'  border='0' width='16' height='16' vspace='0' hspace='0'", "Magnify")
            cDefaultImageColl.Add("<img src='img/New.gif'  border='0'  vspace='0' hspace='0'", "NewIcon")
            cDefaultImageColl.Add("<img src='img/BriefCse.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardBriefcase")
            cDefaultImageColl.Add("<img src='img/Cardfil1.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardCardfile")
            cDefaultImageColl.Add("<img src='img/23_6.png'  border='0' width='16' height='16' vspace='0' hspace='0'", "Standard23_6")
            cDefaultImageColl.Add("<img src='img/133.png'  border='0' width='16' height='16' vspace='0' hspace='0'", "Standard133")
            cDefaultImageColl.Add("<img src='img/Notes.gif'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardNotes")
            cDefaultImageColl.Add("<img src='img/Exclam3.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "Exclamation")
        End Sub

        Public Sub AddLinkItem(ByVal HRef As String, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal AddlParm As String, ByVal Right As String, ByVal PermissionFldName As String)
            cColl.Add(New LinkColumnItems(DG.ColumnType.Link, HRef, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes, AddlParm, Right, PermissionFldName), ItemName)
        End Sub

        Public Sub AddIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Nothing, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Parm2, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Parm2, Parm3, Nothing), ItemName)
        End Sub

        Public Sub AddIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Parm2, Parm3, Parm4), ItemName)
        End Sub

        Public Sub AddDefaultIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Nothing, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddDefaultIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Parm2, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddDefaultIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Parm2, Parm3, Nothing), ItemName)
        End Sub

        Public Sub AddDefaultIconItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
            cColl.Add(New IconColumnItem(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Parm2, Parm3, Parm4), ItemName)
        End Sub

        Public Class IconColumnItem
            Inherits DataBoundColumn
            Private cOnClickMethod As String
            Private cImagePath As String
            Private cImageAttributes As String
            Private cTitle As String
            Private cRight As String
            Private cPermissionFldName As String
            Private cParm2 As String
            Private cParm3 As String
            Private cParm4 As String
            Private cDefaultImage As String
            Private cUseDefaultImage As Boolean
            Private cIsImageType As Boolean
            Private cIsFreeForm As Boolean
            Private cCellText As String
            Private cColumnType As DG.ColumnType

            Public ReadOnly Property IsImageType() As Boolean
                Get
                    Return cIsImageType
                End Get
            End Property
            Public ReadOnly Property IsFreeForm() As Boolean
                Get
                    Return cIsFreeForm
                End Get
            End Property
            Public ReadOnly Property CellText() As String
                Get
                    Return cCellText
                End Get
            End Property
            Public ReadOnly Property DefaultImage() As String
                Get
                    Return cDefaultImage
                End Get
            End Property
            Public ReadOnly Property ImageAttributes() As String
                Get
                    Return cImageAttributes
                End Get
            End Property
            Public ReadOnly Property UseDefaultImage() As Boolean
                Get
                    Return cUseDefaultImage
                End Get
            End Property
            Public ReadOnly Property OnClickMethod() As String
                Get
                    Return cOnClickMethod
                End Get
            End Property
            Public ReadOnly Property ImagePath() As String
                Get
                    Return cImagePath
                End Get
            End Property
            Public ReadOnly Property Title() As String
                Get
                    Return cTitle
                End Get
            End Property
            Public ReadOnly Property Right() As String
                Get
                    Return cRight
                End Get
            End Property
            Public ReadOnly Property PermissionFldName() As String
                Get
                    Return cPermissionFldName
                End Get
            End Property
            Public ReadOnly Property Parm2() As String
                Get
                    Return cParm2
                End Get
            End Property
            Public ReadOnly Property Parm3() As String
                Get
                    Return cParm3
                End Get
            End Property
            Public ReadOnly Property Parm4() As String
                Get
                    Return cParm4
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
                MyBase.New(DG.ColumnType.Icon, ItemName, Nothing)
                cOnClickMethod = OnClickMethod
                cImagePath = ImagePath
                cImageAttributes = ImageAttributes
                cTitle = Title
                cRight = Right
                cPermissionFldName = PermissionFldName
                cParm2 = Parm2
                cParm3 = Parm3
                cParm4 = Parm4
                cUseDefaultImage = False
                cIsImageType = True
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
                MyBase.New(DG.ColumnType.Icon, ItemName, Nothing)
                cOnClickMethod = OnClickMethod
                cDefaultImage = DefaultImage
                cTitle = Title
                cRight = Right
                cPermissionFldName = PermissionFldName
                cParm2 = Parm2
                cParm3 = Parm3
                cParm4 = Parm4
                cUseDefaultImage = True
                cIsImageType = True
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal CellText As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String, ByVal IDoNothing1 As String, ByVal IDoNothing2 As String)
                MyBase.New(DG.ColumnType.Icon, ItemName, Nothing)
                cOnClickMethod = OnClickMethod
                cDefaultImage = DefaultImage
                cTitle = Title
                cRight = Right
                cPermissionFldName = PermissionFldName
                cParm2 = Parm2
                cParm3 = Parm3
                cParm4 = Parm4
                cUseDefaultImage = True
                cIsFreeForm = True
                cCellText = CellText
            End Sub
        End Class

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

    Public Class HiddenItems
        Inherits DataBoundColumn

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String)
            MyBase.New(ColumnType, ItemName, DataFldName)
        End Sub
    End Class

    Public Class CheckboxToggleColumnItems
        Inherits DataBoundColumn
        Private cTrueText As String
        Private cFalseText As String
        Private cTestFld As String
        Private cPermissionFldName As String

        Public ReadOnly Property TestFld() As String
            Get
                Return cTestFld
            End Get
        End Property

        Public ReadOnly Property TrueText() As String
            Get
                Return cTrueText
            End Get
        End Property
        Public ReadOnly Property FalseText() As String
            Get
                Return cFalseText
            End Get
        End Property
        Public ReadOnly Property PermissionFldName() As String
            Get
                Return cPermissionFldName
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal PermissionFldName As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal TestFld As String, ByVal TrueText As String, ByVal FalseText As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, Nothing, TitleFldName, Attributes)
            cTrueText = TrueText
            cFalseText = FalseText
            cTestFld = TestFld
            cPermissionFldName = PermissionFldName
        End Sub
    End Class

    Public Class CheckboxRowSelectColumnItems
        Inherits DataBoundColumn
        Private cTrueText As String
        Private cFalseText As String
        Private cTestFld As String
        Private cPermissionFldName As String

        Public ReadOnly Property TestFld() As String
            Get
                Return cTestFld
            End Get
        End Property

        Public ReadOnly Property TrueText() As String
            Get
                Return cTrueText
            End Get
        End Property
        Public ReadOnly Property FalseText() As String
            Get
                Return cFalseText
            End Get
        End Property
        Public ReadOnly Property PermissionFldName() As String
            Get
                Return cPermissionFldName
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal PermissionFldName As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal TestFld As String, ByVal TrueText As String, ByVal FalseText As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, Nothing, TitleFldName, Attributes)
            cTrueText = TrueText
            cFalseText = FalseText
            cTestFld = TestFld
            cPermissionFldName = PermissionFldName
        End Sub
    End Class

    Public Class LinkColumnItems
        Inherits DataBoundColumn
        Private cHRef As String
        Private cAddlParm As String
        Private cRight As String
        Private cPermissionFldName As String

        Public ReadOnly Property Right() As String
            Get
                Return cRight
            End Get
        End Property
        Public ReadOnly Property PermissionFldName() As String
            Get
                Return cPermissionFldName
            End Get
        End Property

        Public ReadOnly Property AddlParm() As String
            Get
                Return cAddlParm
            End Get
        End Property
        Public ReadOnly Property HRef() As String
            Get
                Return cHRef
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal HRef As String, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal AddlParm As String, ByVal Right As String, ByVal PermissionFldName As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes)
            cHRef = HRef
            cAddlParm = AddlParm
            cRight = Right
            cPermissionFldName = PermissionFldName
        End Sub
    End Class

    Public Class CompoundDataboundColumn
        Inherits DataBoundColumn
        Private cDataFldName2 As String
        Private cDataFldName3 As String
        Private cLineBreak1 As Boolean
        Private cLineBreak2 As Boolean

        Public ReadOnly Property DataFldName2() As String
            Get
                Return cDataFldName2
            End Get
        End Property
        Public ReadOnly Property DataFldName3() As String
            Get
                Return cDataFldName3
            End Get
        End Property
        Public ReadOnly Property LineBreak1() As Boolean
            Get
                Return cLineBreak1
            End Get
        End Property
        Public ReadOnly Property LineBreak2() As Boolean
            Get
                Return cLineBreak2
            End Get
        End Property

        Public Sub New(ByVal ColumnType As ColumnType, ByVal ItemName As String, ByVal DataFldName1 As String, ByVal LineBreak1 As Boolean, ByVal DataFldName2 As String, ByVal LineBreak2 As Boolean, ByVal DataFldName3 As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName1, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes)
            cDataFldName2 = DataFldName2
            cDataFldName3 = DataFldName3
            cLineBreak1 = LineBreak1
            cLineBreak2 = LineBreak2
        End Sub
    End Class

    Public Class DateColumn
        Inherits DataBoundColumn
        Private cTitle As String

        Public ReadOnly Property Title() As String
            Get
                Return cTitle
            End Get
        End Property

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal Title As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, Nothing, Attributes)
            cTitle = Title
        End Sub
    End Class

    Public Class BooleanColumn
        Inherits DataBoundColumn
        Private cTrueValue As String
        Private cTrueText As String
        Private cFalseText As String
        Private cTitle As String
        Public ReadOnly Property TrueText() As String
            Get
                Return cTrueText
            End Get
        End Property
        Public ReadOnly Property FalseText() As String
            Get
                Return cFalseText
            End Get
        End Property
        Public ReadOnly Property Truevalue() As String
            Get
                Return cTrueValue
            End Get
        End Property
        Public ReadOnly Property Title() As String
            Get
                Return cTitle
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal TrueValue As String, ByVal TrueText As String, ByVal FalseText As String, ByVal Title As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, Nothing, Nothing, Attributes)
            cTrueValue = TrueValue
            cTrueText = TrueText
            cFalseText = FalseText
            cTitle = Title
        End Sub
    End Class

    Public Class DataBoundColumnSpecial
        Inherits DataBoundColumn
        Private cDisplayFldName As String

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal DisplayFldName As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes)
            cDisplayFldName = DisplayFldName
        End Sub

        Public ReadOnly Property DisplayFldName() As String
            Get
                Return cDisplayFldName
            End Get
        End Property
    End Class

    Public Class DataBoundColumn
        Private cColumnType As DG.ColumnType
        Private cItemName As String
        Private cDataFldName As String
        Private cHeaderText As String
        Private cSortExpression As String
        Private cVisible As Boolean
        Private cDataFormatString As String
        Private cTitleFldName As String
        Private cAttributes As String

        Public ReadOnly Property ColumnType() As DG.ColumnType
            Get
                Return cColumnType
            End Get
        End Property
        Public ReadOnly Property TitleFldName() As String
            Get
                Return cTitleFldName
            End Get
        End Property
        Public ReadOnly Property ItemName() As String
            Get
                Return cItemName
            End Get
        End Property
        Public ReadOnly Property DataFldName() As String
            Get
                Return cDataFldName
            End Get
        End Property
        Public ReadOnly Property HeaderText() As String
            Get
                Return cHeaderText
            End Get
        End Property
        Public ReadOnly Property SortExpression() As String
            Get
                Return cSortExpression
            End Get
        End Property
        Public ReadOnly Property Visible() As Boolean
            Get
                Return cVisible
            End Get
        End Property
        Public ReadOnly Property DataFormatString() As String
            Get
                Return cDataFormatString
            End Get
        End Property
        Public ReadOnly Property Attributes() As String
            Get
                Return cAttributes
            End Get
        End Property

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String)
            cColumnType = ColumnType
            cItemName = ItemName
            cDataFldName = DataFldName
        End Sub

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
            cColumnType = ColumnType
            cItemName = ItemName
            cDataFldName = DataFldName
            cHeaderText = HeaderText
            cSortExpression = SortExpression
            cVisible = Visible
            cDataFormatString = DataFormatString
            cTitleFldName = TitleFldName
            If Attributes = Nothing Then
                cAttributes = String.Empty
            Else
                cAttributes = Attributes
            End If
        End Sub
    End Class
#End Region

#Region " SortItem "
    Public Class SortItem
        Private cSortExpression As String
        Private cPosition As String
        Private cSortDirection As String
        Private cLastFieldSorted As String
        Public ReadOnly Property SortExpression() As String
            Get
                Return cSortExpression
            End Get
        End Property
        Public ReadOnly Property Position() As String
            Get
                Return cPosition
            End Get
        End Property
        Public Property SortDirection() As String
            Get
                Return cSortDirection
            End Get
            Set(ByVal Value As String)
                cSortDirection = Value
            End Set
        End Property
        Public Property LastFieldSorted() As String
            Get
                Return cLastFieldSorted
            End Get
            Set(ByVal Value As String)
                cLastFieldSorted = Value
            End Set
        End Property
        Public Sub New(ByVal SortExpression As String, ByVal Position As String)
            cSortExpression = SortExpression
            cPosition = Position
            cLastFieldSorted = "F"
            cSortDirection = "N"
        End Sub
    End Class
#End Region

#Region " AnchorObject "
    Public Class AnchorObject
        Private cKey2 As String
        Private cKey3 As String
        Private cKey4 As String

        Public Sub New(ByVal Key2 As String, ByVal Key3 As String, ByVal Key4 As String)
            cKey2 = Key2
            cKey3 = Key3
            cKey4 = Key4
        End Sub

        Public ReadOnly Property Key2() As String
            Get
                Return cKey2
            End Get
        End Property
        Public ReadOnly Property Key3() As String
            Get
                Return cKey3
            End Get
        End Property
        Public ReadOnly Property Key4() As String
            Get
                Return cKey4
            End Get
        End Property
    End Class
#End Region

#Region " Total "
    Public Class Total
        Private cColl As New Collection
        Private cColumnColl As Collection

        Public Sub New(ByRef ColumnColl As Collection)
            cColumnColl = ColumnColl
        End Sub

        Public Sub Add(ByVal ItemName As String, ByVal FormatString As String)
            Dim i As Integer
            Dim ColNum As Integer

            For i = 1 To cColumnColl.Count
                If cColumnColl(i).Visible Then
                    ColNum += 1
                    If cColumnColl(i).ItemName = ItemName Then
                        cColl.Add(New TotalItem(ItemName, FormatString, ColNum), ItemName)
                        Exit For
                    End If
                End If
            Next
        End Sub

        Default Public ReadOnly Property Item(ByVal ItemName As String) As Filter.FilterItem
            Get
                Return cColl(ItemName)
            End Get
        End Property

        Public ReadOnly Property Coll() As Collection
            Get
                Return cColl
            End Get
        End Property

        Public Class TotalItem
            Private cItemName As String
            Private cValue As Object
            Private cFormatString As String
            Private cColNum As Integer

            Public Sub New(ByVal ItemName As String, ByVal FormatString As String, ByVal ColNum As Integer)
                cItemName = ItemName
                cFormatString = FormatString
                cColNum = ColNum
            End Sub

            Public ReadOnly Property ItemName() As String
                Get
                    Return cItemName
                End Get
            End Property

            Public ReadOnly Property FormatString() As String
                Get
                    Return cFormatString
                End Get
            End Property

            Public Property Value() As String
                Get
                    If IsNumeric(cValue) Then
                        Return CType(cValue, System.Decimal).ToString(cFormatString)
                    Else
                        Return String.Empty
                    End If
                End Get
                Set(ByVal Value As String)
                    cValue = Value
                End Set
            End Property

            Public ReadOnly Property ColNum() As Integer
                Get
                    Return cColNum
                End Get
            End Property
        End Class
    End Class
#End Region

#Region " Filter "
    Public Class Filter
        Private cColl As New Collection

        Default Public ReadOnly Property Item(ByVal ItemName As String) As Filter.FilterItem
            Get
                Return cColl(ItemName)
            End Get
        End Property

        Public ReadOnly Property Coll() As Collection
            Get
                Return cColl
            End Get
        End Property

        Public Sub AddTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer)
            cColl.Add(New TextboxCtl(ItemName, DG.FilterTypeEnum.Textbox, DataFldName, MaxLength, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal FilterField As String, ByVal DefaultValue As String)
            cColl.Add(New TextboxCtl(ItemName, DG.FilterTypeEnum.Textbox, DataFldName, MaxLength, FilterField, DefaultValue), ItemName)
        End Sub

        Public Sub AddNameTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal LastNameFld As String, ByVal FirstNameFld As String, ByVal MaxLength As Integer, ByVal DefaultValue As String)
            cColl.Add(New NameTextboxCtl(ItemName, DG.FilterTypeEnum.NameTextbox, DataFldName, LastNameFld, FirstNameFld, MaxLength, DefaultValue), ItemName)
        End Sub

        Public Sub AddDropdown(ByVal ItemName As String, ByVal DataFldName As String)
            cColl.Add(New DropdownCtl(ItemName, DG.FilterTypeEnum.Dropdown, True, False, DataFldName), ItemName)
        End Sub

        Public Sub AddExtendedDropdown(ByVal ItemName As String, ByVal DataFldName As String)
            cColl.Add(New DropdownCtl(ItemName, DG.FilterTypeEnum.Dropdown, False, True, DataFldName), ItemName)
        End Sub

        Public Sub AddDateCtlYM(ByVal ItemName As String, ByVal DataFldName As String, ByVal Caption As String, ByVal EventString As String)
            cColl.Add(New DateCtlYM(ItemName, DG.FilterTypeEnum.DateCtlYM, DataFldName, Caption, EventString), ItemName)
        End Sub

        Public Sub AddDateCtlYMD(ByVal ItemName As String, ByVal DataFldName As String, ByVal Caption As String, ByVal EventString As String, ByVal Attributes As String)
            cColl.Add(New DateCtlYMD(ItemName, DG.FilterTypeEnum.DateCtlYMD, DataFldName, Caption, EventString, Attributes), ItemName)
        End Sub

        Public Sub AddCheckboxRowSelect(ByVal ItemName As String)
            cColl.Add(New CheckboxRowSelect(ItemName, DG.FilterTypeEnum.CheckboxRowSelect, Nothing, Nothing, Nothing, Nothing), ItemName)
        End Sub

        Public Class CheckboxRowSelect
            Inherits FilterItem

            Private cClearAll As Boolean

            Public Sub New(ByVal ItemName As String, ByVal FilterTypeEnum As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal Caption As String, ByVal EventString As String, ByVal Attributes As String)
                MyBase.New(ItemName, FilterTypeEnum, DataFldName)
            End Sub

            Public Property ClearAll() As Boolean
                Get
                    Return cClearAll
                End Get
                Set(ByVal Value As Boolean)
                    cClearAll = Value
                End Set
            End Property
        End Class

        Public Class DateCtlYMD
            Inherits FilterItem

            Private cText As String
            Private cEventString As String
            Private cCaption As String
            Private cAttributes As String

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal Caption As String, ByVal EventString As String, ByVal Attributes As String)
                MyBase.New(ItemName, FilterType, DataFldName)
                cCaption = Caption
                cEventString = EventString
                cAttributes = Attributes
            End Sub

            Public ReadOnly Property Caption() As String
                Get
                    Return cCaption
                End Get
            End Property
            Public ReadOnly Property EventString() As String
                Get
                    Return cEventString
                End Get
            End Property

            Public Overrides Function GetValue() As String
                Return cText
            End Function
            Public Overrides Sub SetValue(ByVal Value As String)
                cText = Value
            End Sub
            Public ReadOnly Property FilterField() As String
                Get
                    Return Nothing
                End Get
            End Property
            Public ReadOnly Property Attributes() As String
                Get
                    Return cAttributes
                End Get
            End Property
        End Class

        Public Class DateCtlYM
            Inherits FilterItem

            Private cText As String
            Private cEventString As String
            Private cCaption As String

            Public ReadOnly Property Caption() As String
                Get
                    Return cCaption
                End Get
            End Property
            Public ReadOnly Property EventString() As String
                Get
                    Return cEventString
                End Get
            End Property
            Public Overrides Function GetValue() As String
                Return cText
            End Function

            Public Overrides Sub SetValue(ByVal Value As String)
                cText = Value
            End Sub

            Public ReadOnly Property FilterField() As String
                Get
                    Return Nothing
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal Caption As String, ByVal EventString As String)
                MyBase.New(ItemName, FilterType, DataFldName)
                cCaption = Caption
                cEventString = EventString
            End Sub
        End Class

        Public Class DropdownCtl
            Inherits FilterItem
            Private cColl As New Collection
            Private cSelectedValue As String = String.Empty
            Private cIsStandard As Boolean
            Private cIsExtended As Boolean

            Public ReadOnly Property IsStandard() As Boolean
                Get
                    Return cIsStandard
                End Get
            End Property

            Public ReadOnly Property IsExtended() As Boolean
                Get
                    Return cIsExtended
                End Get
            End Property

            Public ReadOnly Property Coll() As Collection
                Get
                    Return cColl
                End Get
            End Property

            Public Overrides Sub SetValue(ByVal Value As String)
                cSelectedValue = Value
            End Sub

            Public Overrides Function GetValue() As String
                Return cSelectedValue
            End Function

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal IsStandard As Boolean, ByVal IsExtended As Boolean, ByVal DataFldName As String)
                MyBase.new(ItemName, FilterType, DataFldName)
                cIsStandard = IsStandard
                cIsExtended = IsExtended
            End Sub

            Public Overrides Sub AddDropdownItem(ByVal Value As String, ByVal Text As String, Optional ByVal Selected As Boolean = False)
                If Selected Then
                    ' cSelectedValue = Value
                    SetFilterValue(Value)
                End If
                cColl.Add(New DropdownItem(Value, Text), Value)
            End Sub

            Public Overrides Sub AddExtendedDropdownItem(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
                If Selected Then
                    'cSelectedValue = Value
                    SetFilterValue(Value)
                End If
                cColl.Add(New DropdownItem(Value, Text, Sql), Value)
            End Sub

            Public Class DropdownItem
                Private cValue As String
                Private cText As String
                Private cSql As String

                Public ReadOnly Property Sql() As String
                    Get
                        Return cSql
                    End Get
                End Property
                Public ReadOnly Property Value() As String
                    Get
                        Return cValue
                    End Get
                End Property
                Public ReadOnly Property Text() As String
                    Get
                        Return cText
                    End Get
                End Property

                Public Sub New(ByVal Value As String, ByVal Text As String)
                    cValue = Value
                    cText = Text
                End Sub

                Public Sub New(ByVal Value As String, ByVal Text As String, ByVal Sql As String)
                    cValue = Value
                    cText = Text
                    cSql = Sql
                End Sub

            End Class

        End Class

        Public Class TextboxCtl
            Inherits FilterItem
            Private cMaxLength As Integer
            Private cFilterField As String
            Private cText As String

            Public Overrides Function GetValue() As String
                Return cText
            End Function

            Public Overrides Sub SetValue(ByVal Value As String)
                cText = Value
            End Sub

            Public ReadOnly Property MaxLength() As Integer
                Get
                    Return cMaxLength
                End Get
            End Property

            Public ReadOnly Property FilterField() As String
                Get
                    Return cFilterField
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal MaxLength As Integer)
                MyBase.New(ItemName, FilterType, DataFldName)
                cMaxLength = MaxLength
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal FilterField As String, ByVal Text As String)
                MyBase.New(ItemName, FilterType, DataFldName)
                cMaxLength = MaxLength
                cFilterField = FilterField
                cText = Text
            End Sub
        End Class

        Public Class NameTextboxCtl
            Inherits FilterItem

            Private cLastNameFld As String
            Private cFirstNameFld As String
            Private cMaxLength As Integer
            Private cFilterField As String
            Private cText As String

            Public Overrides Function GetValue() As String
                Return cText
            End Function

            Public Overrides Sub SetValue(ByVal Value As String)
                cText = Value
            End Sub

            Public ReadOnly Property LastNameFld() As String
                Get
                    Return cLastNameFld
                End Get
            End Property

            Public ReadOnly Property FirstNameFld() As String
                Get
                    Return cFirstNameFld
                End Get
            End Property

            Public ReadOnly Property MaxLength() As Integer
                Get
                    Return cMaxLength
                End Get
            End Property

            Public ReadOnly Property FilterField() As String
                Get
                    Return cFilterField
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal LastNameFld As String, ByVal FirstNameFld As String, ByVal MaxLength As Integer)
                MyBase.New(ItemName, FilterType, DataFldName)
                cLastNameFld = LastNameFld
                cFirstNameFld = FirstNameFld
                cMaxLength = MaxLength
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String, ByVal LastNameFld As String, ByVal FirstNameFld As String, ByVal MaxLength As Integer, ByVal Text As String)
                MyBase.New(ItemName, FilterType, DataFldName)
                cLastNameFld = LastNameFld
                cFirstNameFld = FirstNameFld
                cMaxLength = MaxLength
                cFilterField = FilterField
                cText = Text
            End Sub
        End Class

        Public Class FilterItem
            Private cItemName As String
            Private cDataFldName As String
            Private cFilterType As DG.FilterTypeEnum
            Private cFilterValue As String

            Public ReadOnly Property CtlName() As String
                Get
                    Select Case cFilterType
                        Case FilterTypeEnum.Textbox, FilterTypeEnum.NameTextbox
                            Return "txt" & cDataFldName
                        Case FilterTypeEnum.Dropdown
                            Return "dd" & cDataFldName
                        Case Else
                            Throw New Exception("Error #DG429: DG.FilterItem CtlName. Unable to perform match against " & cFilterType.ToString)
                    End Select
                End Get
            End Property
            Public ReadOnly Property ItemName() As String
                Get
                    Return cItemName
                End Get
            End Property
            Public ReadOnly Property DataFldName() As String
                Get
                    Return cDataFldName
                End Get
            End Property
            Public ReadOnly Property FilterType() As FilterTypeEnum
                Get
                    Return cFilterType
                End Get
            End Property

            Public Sub SetFilterValue(ByVal Value As String)
                cFilterValue = Value
            End Sub

            Public Function GetFilterValue() As String
                Return cFilterValue
            End Function

            Public Sub New(ByVal ItemName As String, ByVal FilterType As DG.FilterTypeEnum, ByVal DataFldName As String)
                cItemName = ItemName
                cDataFldName = DataFldName
                cFilterType = FilterType
            End Sub

            Public Overridable Sub AddDropdownItem(ByVal Value As String, ByVal Text As String, Optional ByVal Selected As Boolean = False)
            End Sub

            Public Overridable Sub AddExtendedDropdownItem(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
            End Sub

            Public Overridable Function GetValue() As String
                Return String.Empty
            End Function

            Public Overridable Sub SetValue(ByVal Value As String)
            End Sub

        End Class

    End Class
#End Region

#Region " ExternalFilter "
    Public Class ExternalFilter
        Private cColl As New Collection

        Default Public ReadOnly Property Item(ByVal ItemName As String) As ExternalFilter.FilterItem
            Get
                Return cColl(ItemName)
            End Get
        End Property

        Public ReadOnly Property Coll() As Collection
            Get
                Return cColl
            End Get
        End Property

        Public Sub AddTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal Position As Integer)
            cColl.Add(New TextboxCtl(ItemName, DataFldName, MaxLength, Nothing, Nothing, Position), ItemName)
        End Sub

        Public Sub AddTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal FilterField As String, ByVal DefaultValue As String, ByVal Position As Integer)
            cColl.Add(New TextboxCtl(ItemName, DataFldName, MaxLength, FilterField, DefaultValue, Position), ItemName)
        End Sub

        Public Sub AddDropdown(ByVal ItemName As String, ByVal DataFldName As String, ByVal Position As Integer, Optional ByVal EventString As String = Nothing)
            cColl.Add(New DropdownCtl(ItemName, True, False, DataFldName, Position, EventString), ItemName)
        End Sub

        Public Sub AddExtendedDropdown(ByVal ItemName As String, ByVal DataFldName As String, ByVal Position As Integer, Optional ByVal EventString As String = Nothing)
            cColl.Add(New DropdownCtl(ItemName, False, True, DataFldName, Position, EventString), ItemName)
        End Sub

        Public Sub AddLinkItem(ByVal ItemName As String, ByVal DataFldName As String, ByVal Position As Integer, ByVal EventString As String)
            cColl.Add(New LinkFilterCtl(ItemName, DataFldName, Position, EventString), ItemName)
        End Sub

        Public Class DropdownCtl
            Inherits FilterItem
            Private cColl As New Collection
            Private cSelectedValue As String = String.Empty
            Private cIsStandard As Boolean
            Private cIsExtended As Boolean
            Private cEventString As String

            Public ReadOnly Property EventString() As String
                Get
                    Return cEventString
                End Get
            End Property

            Public ReadOnly Property IsStandard() As Boolean
                Get
                    Return cIsStandard
                End Get
            End Property

            Public ReadOnly Property IsExtended() As Boolean
                Get
                    Return cIsExtended
                End Get
            End Property

            Public ReadOnly Property Coll() As Collection
                Get
                    Return cColl
                End Get
            End Property

            Public Overrides Sub SetValue(ByVal Value As String)
                cSelectedValue = Value
            End Sub

            Public Overrides Function GetValue() As String
                Return cSelectedValue
                'Dim i As Integer
                'For i = 1 To cColl.Count
                '    If cColl(i).Selected = True Then
                '        Return cColl(i).Value
                '    End If
                'Next
                'Return String.Empty
            End Function

            Public Sub New(ByVal ItemName As String, ByVal IsStandard As Boolean, ByVal IsExtended As Boolean, ByVal DataFldName As String, ByVal Position As Integer, ByVal EventString As String)
                MyBase.new(ItemName, DataFldName, False, True, False, Position)
                cIsStandard = IsStandard
                cIsExtended = IsExtended
                cEventString = EventString
            End Sub

            Public Overrides Sub AddDropdownItem(ByVal Value As String, ByVal Text As String, Optional ByVal Selected As Boolean = False)
                If Selected Then
                    cSelectedValue = Value
                End If
                cColl.Add(New DropdownItem(Value, Text), Value)
            End Sub

            Public Overrides Sub AddExtendedDropdownItem(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
                If Selected Then
                    cSelectedValue = Value
                End If
                cColl.Add(New DropdownItem(Value, Text, Sql), Value)
            End Sub

            Public Class DropdownItem
                Private cValue As String
                Private cText As String
                Private cSql As String

                Public ReadOnly Property Sql() As String
                    Get
                        Return cSql
                    End Get
                End Property
                Public ReadOnly Property Value() As String
                    Get
                        Return cValue
                    End Get
                End Property
                Public ReadOnly Property Text() As String
                    Get
                        Return cText
                    End Get
                End Property

                Public Sub New(ByVal Value As String, ByVal Text As String)
                    cValue = Value
                    cText = Text
                End Sub

                Public Sub New(ByVal Value As String, ByVal Text As String, ByVal Sql As String)
                    cValue = Value
                    cText = Text
                    cSql = Sql
                End Sub

            End Class

        End Class

        Public Class LinkFilterCtl
            Inherits FilterItem

            Private cText As String
            Private cFilterExpression As String
            Private cEventString As String


            Public ReadOnly Property EventString() As String
                Get
                    Return cEventString
                End Get
            End Property
            Public Property Text() As String
                Get
                    Return cText
                End Get
                Set(ByVal Value As String)
                    cText = Value
                End Set
            End Property
            Public Property FilterExpression() As String
                Get
                    Return cFilterExpression
                End Get
                Set(ByVal Value As String)
                    cFilterExpression = Value
                End Set
            End Property

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal Position As Integer, ByVal EventString As String)
                MyBase.New(ItemName, DataFldName, False, False, True, Position)

                cEventString = EventString
            End Sub

            Public Sub SetLinkValues(ByVal Text As String, ByVal FilterExpression As String)
                cText = Text
                cFilterExpression = FilterExpression
            End Sub

        End Class

        Public Class TextboxCtl
            Inherits FilterItem
            Private cMaxLength As Integer
            Private cInternalFilterField As String
            Private cText As String

            Public Overrides Function GetValue() As String
                Return cText
            End Function
            Public Overrides Sub SetValue(ByVal Value As String)
                cText = Value
            End Sub

            Public ReadOnly Property MaxLength() As Integer
                Get
                    Return cMaxLength
                End Get
            End Property

            Public ReadOnly Property FilterField() As String
                Get
                    Return cInternalFilterField
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal Position As Integer)
                MyBase.New(ItemName, DataFldName, True, False, False, Position)
                cMaxLength = MaxLength
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal FilterField As String, ByVal Text As String, ByVal Position As Integer)
                MyBase.New(ItemName, DataFldName, True, False, False, Position)
                cMaxLength = MaxLength
                cInternalFilterField = FilterField
                cText = Text
            End Sub
        End Class

        Public Class FilterItem
            Private cItemName As String
            Private cDataFldName As String
            Private cIsTextBox As Boolean
            Private cIsDropdown As Boolean
            Private cIsLink As Boolean
            Private cPosition As Integer
            Private cOverrideValue As String
            Private cControlType As ControlTypeEnum

            Public Enum ControlTypeEnum
                Textbox = 1
                StandardDropdown = 2
                ExtendedDropdown = 3
                Link = 4
            End Enum

            Public Sub SetOverrideValue(ByVal Value As String)
                cOverrideValue = Value
            End Sub

            Public Function GetOverrideValue() As String
                Return cOverrideValue
            End Function

            Public ReadOnly Property ControlType() As ControlTypeEnum
                Get
                    Return cControlType
                End Get
            End Property

            Public ReadOnly Property IsLink() As Boolean
                Get
                    Return cIsLink
                End Get
            End Property
            Public ReadOnly Property Position() As Integer
                Get
                    Return cPosition
                End Get
            End Property

            Public ReadOnly Property CtlName() As String
                Get
                    Dim Value As String = Nothing
                    If cIsTextBox Then
                        Value = "txt" & cDataFldName
                    ElseIf cIsDropdown Then
                        Value = "dd" & cDataFldName
                    ElseIf cIsLink Then
                        Value = "lnk" & cDataFldName
                    End If
                    Return Value
                End Get
            End Property
            Public ReadOnly Property ItemName() As String
                Get
                    Return cItemName
                End Get
            End Property
            Public ReadOnly Property DataFldName() As String
                Get
                    Return cDataFldName
                End Get
            End Property
            Public ReadOnly Property IsTextBox() As Boolean
                Get
                    Return cIsTextBox
                End Get
            End Property
            Public ReadOnly Property IsDropdown() As Boolean
                Get
                    Return cIsDropdown
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal IsTextBox As Boolean, ByVal IsDropdown As Boolean, ByVal IsLink As Boolean, ByVal Position As Integer)
                cItemName = ItemName
                cDataFldName = DataFldName
                cIsTextBox = IsTextBox
                cIsDropdown = IsDropdown
                cIsLink = IsLink
                cPosition = Position
                If IsTextBox Then
                    cControlType = ControlTypeEnum.Textbox
                ElseIf IsDropdown Then
                    cControlType = ControlTypeEnum.StandardDropdown
                ElseIf IsLink Then
                    cControlType = ControlTypeEnum.Link
                End If
            End Sub

            Public Overridable Sub AddDropdownItem(ByVal Value As String, ByVal Text As String, Optional ByVal Selected As Boolean = False)
                'cColl.Add(New DropdownItem(Value, Text, Selected), Value)
            End Sub

            Public Overridable Sub AddLink(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
            End Sub

            Public Overridable Sub AddExtendedDropdownItem(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
            End Sub

            Public Overridable Function GetValue() As String
                Return String.Empty
            End Function

            Public Overridable Sub SetValue(ByVal Value As String)
            End Sub

        End Class

    End Class
#End Region

#Region " ChildTable "
    Public Class ChildTableClass
        Private cItemName As String
        'Private cDataFldName As String
        Private cPermissionFldName As String
        Private cChildTableSelectColumn As ChildTableSelectColumn

        Public Property ChildTableSelectColumn() As DG.ChildTableSelectColumn
            Get
                Return cChildTableSelectColumn
            End Get
            Set(ByVal Value As DG.ChildTableSelectColumn)
                cChildTableSelectColumn = Value
            End Set
        End Property

        Public ReadOnly Property ItemName() As String
            Get
                Return cItemName
            End Get
        End Property
        Public ReadOnly Property PermissionFldName() As String
            Get
                Return cPermissionFldName
            End Get
        End Property

        Public Sub New(ByVal ItemName As String)
            cItemName = ItemName
            'cPermissionFldName = PermissionFldName
        End Sub
    End Class
#End Region

#Region " Menu "
    Public Class Menu
        Private cColl As Collection
        Private cCellWidthPercent As Integer

        Public Enum ObjectTypeEnum
            IsLink = 1
            IsButton = 2
        End Enum

        Public ReadOnly Property CellWidthPercent() As Integer
            Get
                Return cCellWidthPercent
            End Get
        End Property

        Public ReadOnly Property Coll() As Collection
            Get
                Return cColl
            End Get
        End Property

        Public Sub New(ByVal CellWidthPercent As Integer)
            cColl = New Collection
            cCellWidthPercent = CellWidthPercent
        End Sub

        Public Sub AddItem(ByVal ObjectType As ObjectTypeEnum, ByVal OnClickMethod As String, ByVal Text As String, ByVal Right As String)
            cColl.Add(New MenuItem(ObjectType, OnClickMethod, Text, Right))
        End Sub

        Public Class MenuItem
            Private cIsLink As Boolean
            Private cIsButton As Boolean
            Private cOnClickMethod As String
            Private cText As String
            Private cIsVisible As Boolean = True
            Private cRight As String

            Public ReadOnly Property Right() As String
                Get
                    Return cRight
                End Get
            End Property

            Public Property IsVisible() As Boolean
                Get
                    Return cIsVisible
                End Get
                Set(ByVal Value As Boolean)
                    cIsVisible = Value
                End Set
            End Property

            Public ReadOnly Property IsLink() As Boolean
                Get
                    Return cIsLink
                End Get
            End Property
            Public ReadOnly Property IsButton() As Boolean
                Get
                    Return cIsButton
                End Get
            End Property
            Public ReadOnly Property OnClickMethod() As String
                Get
                    Return cOnClickMethod
                End Get
            End Property

            Public Sub New(ByVal ObjectType As ObjectTypeEnum, ByVal OnClickMethod As String, ByVal Text As String, ByVal Right As String)
                If ObjectType = ObjectTypeEnum.IsButton Then
                    cIsButton = True
                Else
                    cIsLink = True
                End If
                cOnClickMethod = OnClickMethod
                cText = Text
                cRight = Right
            End Sub
        End Class
    End Class
#End Region

#Region " MenuLevel2 "
    Public Class MenuLevel2Class
        Private cItem As String
        Private cAttributes As String
        Private cIsVisible As Boolean = True

        Public Sub AddItem(ByVal Item As String, ByVal Attributes As String)
            cItem = Item
            cAttributes = Attributes
        End Sub
        Public ReadOnly Property Item() As String
            Get
                Return cItem
            End Get
        End Property
        Public ReadOnly Property Attributes() As String
            Get
                Return cAttributes
            End Get
        End Property
        Public Property IsVisible() As Boolean
            Get
                Return cIsVisible
            End Get
            Set(ByVal Value As Boolean)
                cIsVisible = Value
            End Set
        End Property
    End Class
#End Region
End Class
