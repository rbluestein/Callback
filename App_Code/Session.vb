'Public MustInherit Class DataSession
'End Class

Public Class PageSession
    Private cPageInitiallyLoaded As Boolean
    Private cPageReturnOnLoadMessasge As String


    Public Property PageReturnOnLoadMessage() As String
        Get
            Return cPageReturnOnLoadMessasge
        End Get
        Set(ByVal Value As String)
            cPageReturnOnLoadMessasge = Value
        End Set
    End Property
    Public Property PageInitiallyLoaded() As Boolean
        Get
            Return cPageInitiallyLoaded
        End Get
        Set(ByVal Value As Boolean)
            cPageInitiallyLoaded = Value
        End Set
    End Property
End Class

Public Class DataGridSession
    Private cSortReference As String = String.Empty
    Private cFilterOnOffState As String = String.Empty
    Private cSql As String
    Private cExcessiveRecordsWarningInEffect As Boolean
    Private cInitialReportDataSuppressInEffect As Boolean
    Private cJumpToRecID As String = String.Empty
    Private cActiveSubTableRecID As String = String.Empty
    Private cSubTableInd As String = String.Empty

    Public Property SortReference() As String
        Get
            Return cSortReference
        End Get
        Set(ByVal Value As String)
            cSortReference = Value
        End Set
    End Property

    Public Property FilterOnOffState() As String
        Get
            Return cFilterOnOffState
        End Get
        Set(ByVal Value As String)
            cFilterOnOffState = Value
        End Set
    End Property

    Public Property ExcessiveRecordsWarningInEffect() As Boolean
        Get
            Return cExcessiveRecordsWarningInEffect
        End Get
        Set(ByVal Value As Boolean)
            cExcessiveRecordsWarningInEffect = Value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Return cSql
        End Get
        Set(ByVal Value As String)
            cSql = Value
        End Set
    End Property

    Public Property InitialReportDataSuppressInEffect() As Boolean
        Get
            Return cInitialReportDataSuppressInEffect
        End Get
        Set(ByVal Value As Boolean)
            cInitialReportDataSuppressInEffect = Value
        End Set
    End Property

    Public Property JumpToRecID() As String
        Get
            Return cJumpToRecID
        End Get
        Set(ByVal Value As String)
            cJumpToRecID = Value
        End Set
    End Property
    Public Property ActiveSubTableRecID() As String
        Get
            Return cActiveSubTableRecID
        End Get
        Set(ByVal Value As String)
            cActiveSubTableRecID = Value
        End Set
    End Property
    Public Property SubTableInd() As String
        Get
            Return cSubTableInd
        End Get
        Set(ByVal value As String)
            cSubTableInd = value
        End Set
    End Property
End Class

Public Class IndexSession
    Inherits PageSession

    Private cCallbackID As String = String.Empty
    Private cCallbackAttemptID As String = String.Empty
    Private cDGSess As New DataGridSession
    Private cFilter As New IndexSession.FilterClass
    Private cUserConfirmationResolvedAsTrue As Boolean

    Public ReadOnly Property DGSess() As DataGridSession
        Get
            Return cDGSess
        End Get
    End Property

    Public ReadOnly Property Filter() As IndexSession.FilterClass
        Get
            Return cFilter
        End Get
    End Property

    Public Property CallbackID() As String
        Get
            Return cCallbackID
        End Get
        Set(ByVal Value As String)
            cCallbackID = Value
        End Set
    End Property
    Public Property CallbackAttemptID() As String
        Get
            Return cCallbackAttemptID
        End Get
        Set(ByVal value As String)
            cCallbackAttemptID = value
        End Set
    End Property

    Public Class FilterClass
        Private cDateRange As String = "|||||"
        Private cEmpName As String
        Private cClientID As String
        Private cState As String
        Private cCallPurposeCode As String
        Private cPriorityTagInd As String
        Private cPreferSpanishInd As String
        Private cNumAttempts As String
        Private cStatusCodeAdjDetail As String

        Public Property DateRange() As String
            Get
                Return cDateRange
            End Get
            Set(ByVal Value As String)
                cDateRange = Value
            End Set
        End Property
        Public Property EmpName() As String
            Get
                Return cEmpName
            End Get
            Set(ByVal value As String)
                cEmpName = value
            End Set
        End Property
        Public Property ClientID() As String
            Get
                Return cClientID
            End Get
            Set(ByVal Value As String)
                cClientID = Value
            End Set
        End Property
        Public Property State() As String
            Get
                Return cState
            End Get
            Set(ByVal value As String)
                cState = value
            End Set
        End Property
        Public Property CallPurposeCode() As String
            Get
                Return cCallPurposeCode
            End Get
            Set(ByVal value As String)
                cCallPurposeCode = value
            End Set
        End Property
        Public Property StatusCodeAdjDetail() As String
            Get
                Return cStatusCodeAdjDetail
            End Get
            Set(ByVal value As String)
                cStatusCodeAdjDetail = value
            End Set
        End Property
        Public Property PriorityTagInd() As String
            Get
                Return cPriorityTagInd
            End Get
            Set(ByVal value As String)
                cPriorityTagInd = value
            End Set
        End Property
        Public Property PreferSpanishInd() As String
            Get
                Return cPreferSpanishInd
            End Get
            Set(ByVal value As String)
                cPreferSpanishInd = value
            End Set
        End Property
        Public Property NumAttempts() As String
            Get
                Return cNumAttempts
            End Get
            Set(ByVal value As String)
                cNumAttempts = value
            End Set
        End Property
    End Class
End Class