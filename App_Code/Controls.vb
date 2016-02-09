Imports System.Data

Public Class CollFS
    Private cRows As New CollX

    Public ReadOnly Property Rows As CollX
        Get
            Return cRows
        End Get
    End Property

    Public ReadOnly Property GetValueX(ByVal RowID As String, ByVal FldName As String) As Object
        Get
            Dim i As Integer
            Dim RowIDX As Integer
            Dim Value As Object = Nothing

            If IsNumeric(RowID) Then
                RowIDX = CType(RowID, System.Int64)
                For i = 1 To cRows.Count
                    If i = RowIDX Then
                        Value = cRows(i)(FldName).Value
                        Exit For
                    End If
                Next
            Else
                Value = GetValue(RowID, FldName)
            End If

            Return Value
        End Get
    End Property
    Public ReadOnly Property GetValue(ByVal RowID As String, ByVal FldName As String) As Object
        Get
            Dim Value As Object = Nothing
            RowID = RowID.ToUpper
            For i = 1 To cRows.Count
                If cRows.Key(i).ToUpper = RowID Then
                    Value = cRows(i)(FldName).Value()
                End If
            Next

            Return Value
        End Get
    End Property
    'Default Public ReadOnly Property GetValue(ByVal Idx As Integer, ByVal FldName As String, Optional ByVal CallingMethod As String = Nothing) As Object
    '    Get
    '        If Idx <= cRows.Count - 1 Then
    '            Return cRows(Idx).Value
    '        Else
    '            Throw New CollXError("Error #3606: CollX.Coll item not found error. Idx: " & Idx.ToString & IIf(CallingMethod <> Nothing, ". Calling method: " & CallingMethod, ""))  'CollXError
    '        End If
    '    End Get
    'End Property
    Public ReadOnly Property Count As Integer
        Get
            Return cRows.Count
        End Get
    End Property

    'Public Sub AddRow(ByVal RowID As String)
    '    cRows.Assign(RowID, New CollX)
    'End Sub
    Public Sub AddRow(ByVal KeyFldName As String, ByVal Value As Object, ByVal DataType As DataTypeEnum)
        cRows.Assign(Value, New CollX)
        Select Case DataType
            Case DataTypeEnum.String
                cRows(Value).assign(KeyFldName, New StrItem(KeyFldName, Value))
            Case DataTypeEnum.Integer
                cRows(Value).assign(KeyFldName, New IntItem(KeyFldName, Value))
            Case DataTypeEnum.Decimal
                cRows(Value).assign(KeyFldName, New DecItem(KeyFldName, Value))
            Case DataTypeEnum.Boolean
                cRows(Value).assign(KeyFldName, New BoolItem(KeyFldName, Value))
            Case DataTypeEnum.Date
                cRows(Value).assign(KeyFldName, New DateItem(KeyFldName, Value))
        End Select
    End Sub
    Public Sub AssignStrItem(ByVal RowID As String, ByVal FldName As String, ByVal Value As String)
        cRows(RowID).assign(FldName, New StrItem(FldName, Value))
    End Sub
    Public Sub AssignIntItem(ByVal RowID As String, ByVal FldName As String, ByVal Value As Integer)
        cRows(RowID).Assign(FldName, New IntItem(FldName, Value))
    End Sub
    Public Sub AssignDecItem(ByVal RowID As String, ByVal FldName As String, ByVal Value As Decimal)
        cRows(RowID).Assign(FldName, New DecItem(FldName, Value))
    End Sub
    Public Sub AssignBoolItem(ByVal RowID As String, ByVal FldName As String, ByVal Value As Boolean)
        cRows(RowID).Assign(FldName, New BoolItem(FldName, Value))
    End Sub
    Public Sub AssignDateItem(ByVal RowID As String, ByVal FldName As String, ByVal Value As DateTime)
        cRows(RowID).Assign(FldName, New DateItem(FldName, Value))
    End Sub
    Public Sub RemoveRowItem(ByVal RowID As String)
        cRows.Remove(RowID)
    End Sub
    Public Sub RemoveColItem(ByVal RowID As String, ByVal FldName As String)
        cRows(RowID).Remove(FldName)
    End Sub

#Region " DataType items "
    Public Class BaseItem
        Private cFldName As String
        Private cFldNameNumeric As Boolean

        Public Sub New(ByVal FldName As String)
            cFldName = FldName
            If IsNumeric(FldName) Then
                cFldNameNumeric = True
            End If
        End Sub
        Public ReadOnly Property FldName As String
            Get
                Return cFldName
            End Get
        End Property
        Public ReadOnly Property FldNameNumeric As Boolean
            Get
                Return cFldNameNumeric
            End Get
        End Property
    End Class
    Public Class StrItem
        Inherits BaseItem
        Private cValue As String

        Public Sub New(ByVal FldName As String, ByVal Value As String)
            MyBase.New(FldName)
            cValue = Value
        End Sub
        Public Property Value As String
            Set(value As String)
                cValue = value
            End Set
            Get
                Return cValue
            End Get
        End Property
    End Class
    Public Class IntItem
        Inherits BaseItem
        Private cValue As Integer

        Public Sub New(ByVal FldName As String, ByVal Value As Integer)
            MyBase.new(FldName)
            cValue = Value
        End Sub
        Public Property Value As Integer
            Set(value As Integer)
                cValue = value
            End Set
            Get
                Return cValue
            End Get
        End Property
    End Class
    Public Class DecItem
        Inherits BaseItem
        Private cValue As Decimal

        Public Sub New(ByVal FldName As String, ByVal Value As Decimal)
            MyBase.New(FldName)
            cValue = Value
        End Sub
        Public Property Value As Decimal
            Set(value As Decimal)
                cValue = value
            End Set
            Get
                Return cValue
            End Get
        End Property
    End Class
    Public Class BoolItem
        Inherits BaseItem
        Private cValue As Boolean

        Public Sub New(ByVal FldName As String, ByVal Value As Boolean)
            MyBase.New(FldName)
            cValue = Value
        End Sub
        Public Property Value As Boolean
            Set(value As Boolean)
                cValue = value
            End Set
            Get
                Return cValue
            End Get
        End Property
    End Class
    Public Class DateItem
        Inherits BaseItem
        Private cValue As DateTime

        Public Sub New(ByVal FldName As String, ByVal Value As DateTime)
            MyBase.New(FldName)
            cValue = Value
        End Sub

        Public Property Value As DateTime
            Set(value As DateTime)
                cValue = value
            End Set
            Get
                Return cValue
            End Get
        End Property
    End Class
#End Region
End Class

Public Class CollX
    Inherits System.Collections.CollectionBase
    ' Private Bittem As ListItem

    Public Sub New()
        List.Add(DBNull.Value)
    End Sub

    Public Overloads ReadOnly Property Count() As Integer
        Get
            Return List.Count - 1
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Idx As Integer, Optional ByVal CallingMethod As String = Nothing) As Object
        'Get
        '    Return List(Idx).Value
        'End Get
        Get
            If Idx <= List.Count - 1 Then
                Return List(Idx).Value
            Else
                Throw New CollXError("Error #3606: CollX.Coll item not found error. Idx: " & Idx.ToString & IIf(CallingMethod <> Nothing, ". Calling method: " & CallingMethod, ""))  'CollXError
            End If
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Key As String, Optional ByVal CallingMethod As String = Nothing) As Object
        Get
            Dim i As Integer
            Dim KeyUpper As String
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3604: CollX.Coll item not found error. Key: " & Key & IIf(CallingMethod <> Nothing, ". Calling method: " & CallingMethod, ""))  'CollXError
        End Get
    End Property

    'Default Public ReadOnly Property Coll(ByVal Idx As Integer, CallingMethod As String) As Object
    '    'Get
    '    '    Return List(Idx).Value
    '    'End Get
    '    Get
    '        If Idx <= List.Count - 1 Then
    '            Return List(Idx).Value
    '        Else
    '            Throw New CollXError("Error #3606: CollX.Coll item not found error. Idx: " & Idx.ToString & IIf(CallingMethod <> Nothing, ". Calling method: " & CallingMethod, ""))  'CollXError
    '        End If
    '    End Get
    'End Property


    'Default Public ReadOnly Property Coll(ByVal Key As String, ByVal CallingMethod As String) As Object
    '    Get
    '        Dim i As Integer
    '        Dim KeyUpper As String
    '        KeyUpper = Key.ToUpper
    '        For i = 1 To List.Count - 1
    '            If List(i).Key.ToUpper = KeyUpper Then
    '                Return List(i).Value
    '            End If
    '        Next
    '        Throw New CollXError("Error #3604: CollX.Coll item not found error. Key: " & Key & IIf(CallingMethod <> Nothing, ". Calling method: " & CallingMethod, ""))  'CollXError
    '    End Get
    'End Property

    Public Function TreatKeyAsString(ByVal Key As String, Optional ByVal CallingMethod As String = Nothing) As Object
        Dim i As Integer
        Dim KeyUpper As String

        Try
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3611.1: CallX.TreatKeyAsString item not found error. Key: " & Key & IIf(CallingMethod <> Nothing, ". Calling method: " & CallingMethod, ""))  'CollXError

        Catch ex As Exception
            Throw New CollXError("Error #3611: CallX.TreatKeyAsString item not found error. Key: " & Key)  'CollXError
        End Try
    End Function

    Public ReadOnly Property KeyFromValue(ByVal Input As String) As String
        Get
            Dim i As Integer
            Input = Input.ToUpper

            For i = 1 To List.Count - 1
                If List(i).Value.ToUpper = Input Then
                    'Return List(i).Value
                    Return List(i).Key
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Key(ByVal Idx As Integer) As String
        Get
            Dim i As Integer
            For i = 1 To List.Count - 1
                If i = Idx Then
                    'Return List(i).Value
                    Return List(i).Key
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property DoesKeyExist(ByVal Key As String) As Boolean
        Get
            Dim i As Integer
            If IsNumeric(Key) Then
                For i = 1 To List.Count - 1
                    If List(i).Key = Key Then
                        Return True
                    End If
                Next
            Else
                Key = Key.ToUpper
                For i = 1 To List.Count - 1
                    If List(i).Key.ToUpper = Key Then
                        Return True
                    End If
                Next
            End If
            Return False
        End Get
    End Property

    Public Sub RenameKey(ByVal CurrentKey As String, ByVal NewKey As String)
        Dim i As Integer
        If IsNumeric(CurrentKey) Then
            For i = 1 To List.Count - 1
                If List(i).Key = CurrentKey Then
                    List(i).Key = NewKey
                    Return
                End If
            Next
        Else
            CurrentKey = CurrentKey.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = CurrentKey.ToUpper Then
                    List(i).Key = NewKey
                    Return
                End If
            Next
        End If
    End Sub

    Public Sub Assign(ByVal Key As String, ByVal Value As Object)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key Then
                    List(i).Value = Nothing
                    List(i).Value = Value
                    Found = True
                    If Found Then
                        Exit For
                    End If
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key, Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3614: CallX.Assign. Item not found error. Key: " & Key)
        End Try
    End Sub

    Public Sub Assign(ByVal Key_Value As String)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key_Value Then
                    List(i).Value = Nothing
                    List(i).Value = Key_Value
                    Found = True
                End If
            Next

            If Not Found Then
                List.Add(New KeyValuePair(Key_Value, Key_Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3605: CallX.Assign. " & ex.Message)
        End Try
    End Sub

    Public Sub ConvertRow(ByRef dr As DataRow)
        Try

            Dim i As Integer
            For i = 0 To dr.ItemArray.GetUpperBound(0)
                Assign(dr.Table.Columns(i).ColumnName, dr(i))
            Next

        Catch ex As Exception
            Throw New CollXError("Error #3610: CallX.ConvertRow. " & ex.Message)
        End Try
    End Sub

    Public Function ConvertToStr(ByVal Delimiter As String) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        Try

            For i = 1 To List.Count - 1
                If i < List.Count - 1 Then
                    sb.Append(List(i).Value & Delimiter)
                Else
                    sb.Append(List(i).Value)
                End If
            Next
            Return sb.ToString
        Catch ex As Exception
            Throw New CollXError("Error #3612: CallX.ConvertToStr. " & ex.Message)
        End Try
    End Function

    Public Function ConvertToStr(ByVal RowDelimiter As String, ByVal Key_ValueDelimiter As String) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        Try

            For i = 1 To List.Count - 1
                If i < List.Count - 1 Then
                    sb.Append(List(i).Key & Key_ValueDelimiter & List(i).Value & RowDelimiter)
                Else
                    sb.Append(List(i).Key & Key_ValueDelimiter & List(i).Value)
                End If
            Next
            Return sb.ToString
        Catch ex As Exception
            Throw New CollXError("Error #3615: CallX.ConvertToStr. " & ex.Message)
        End Try
    End Function

    Public Function CollxToSql() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i < List.Count - 1 Then
                sb.Append(List(i).Key & "=" & List(i).Value & ", ")
            Else
                sb.Append(List(i).Key & "=" & List(i).Value)
            End If
        Next
        Return sb.ToString
    End Function

    Public Function CollxToParameters() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i = 1 Then
                sb.Append("?" & List(i).Key & "=" & List(i).Value)
            Else
                sb.Append("&" & List(i).Key & "=" & List(i).Value)
            End If
        Next

        Return sb.ToString
    End Function

#Region " New from... "
    Public Shared Function NewFromList(ByVal Input As String, ByVal Delimiter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Coll As New CollX
        Box = Input.Split(Delimiter)
        If Box.GetUpperBound(0) > -1 Then
            For i = 0 To Box.GetUpperBound(0)
                Coll.Assign(Box(i))
            Next
        End If
        Return Coll
    End Function

    Public Shared Function NewFromDataRow(ByRef dr As DataRow) As CollX
        Dim i As Integer
        Dim dt As DataTable
        Dim Coll As New CollX
        dt = dr.Table
        For i = 0 To dt.Columns.Count - 1
            Coll.Assign(dt.Columns(i).ColumnName, dr(i))
        Next
        Return Coll
    End Function

    Public Shared Function NewFromTable(ByRef dt As DataTable) As CollX
        Dim i As Integer
        Dim Coll As New CollX

        If dt.Columns.Count = 1 Then
            For i = 0 To dt.Rows.Count - 1
                Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(0))
            Next
        Else
            For i = 0 To dt.Rows.Count - 1
                Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(1))
            Next
        End If

        Return Coll
    End Function

    Public Shared Function NewFromKeyValue(ByVal Input As String, ByVal RowDelimter As String, ByVal ColDelimter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Box2 As String()
        Dim Coll As New CollX
        Box = Input.Split(RowDelimter)
        For i = 0 To Box.GetUpperBound(0)
            Box2 = Box(i).Split(ColDelimter)
            Coll.Assign(Box2(0), Box2(1))
        Next
        Return Coll
    End Function
#End Region
    Public Overloads Sub Clear()
        Dim i As Integer
        For i = List.Count - 1 To 1 Step -1
            List.Remove(List(i))
        Next
    End Sub
    Public Overloads Sub RemoveAt(ByVal Index As Integer)
        List.RemoveAt(Index)
    End Sub

    Public Overloads Sub Remove(ByVal Key As String)
        Dim i As Integer
        For i = 1 To List.Count - 1
            If List(i).Key.ToUpper = Key.ToUpper Then
                List.Remove(List(i))
                Exit For
            End If
        Next
    End Sub

    Public Overloads Sub RemoveAndDestroy(ByVal Key As String)
        Dim i As Integer
        For i = 1 To List.Count - 1
            If List(i).Key.ToUpper = Key.ToUpper Then
                List(i).Value = Nothing
                List.Remove(List(i))
                Exit For
            End If
        Next
    End Sub

    Public Shared Function GetDistinct(ByRef dt As DataTable, ByVal FieldName As String) As CollX
        Dim i As Integer
        Dim Coll As New CollX
        Dim FirstValue As Object = DBNull.Value

        If dt.Rows.Count = 0 Then
            Return Nothing
        Else
            Coll.Assign(dt.Rows(0)(FieldName))
            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(0)(FieldName) <> Coll(Coll.Count) Then
                    Coll.Assign(dt.Rows(i)(FieldName))
                End If
            Next
        End If
        Return Coll
    End Function

    Public ReadOnly Property View() As String
        Get
            Dim sbOutput As New System.Text.StringBuilder
            Dim Val As String

            Try
                sbOutput.Append(Environment.NewLine)
                For i = 1 To List.Count - 1
                    Try
                        Val = List(i).Value
                    Catch ex As Exception
                        Val = "<object>"
                    End Try
                    sbOutput.Append("(" & i.ToString & ") " & List(i).Key & "|" & Val & Environment.NewLine)
                Next
                Return sbOutput.ToString

            Catch ex As Exception
                Throw New CollXError("Error #3608: CallX.View. " & ex.Message)
            End Try
        End Get
    End Property

    Public Shared Function Clone(ByVal InputColl As CollX) As CollX
        Dim i As Integer
        Dim OutputColl As New CollX

        Try

            For i = 1 To InputColl.Count
                OutputColl.Assign(InputColl.Key(i), InputColl(i))
            Next
            Return OutputColl

        Catch ex As Exception
            Throw New CollXError("Error #3609: CallX.Clone. " & ex.Message)
        End Try
    End Function

    Public Shared Function CollxToDT(ByRef Coll As CollX) As DataTable
        Dim i As Integer
        Dim dt As New DataTable
        Dim dr As DataRow

        dt.Columns.Add(New DataColumn("Key", GetType(System.String)))
        dt.Columns.Add(New DataColumn("Value", GetType(System.String)))

        For i = 1 To Coll.Count
            dr = dt.NewRow
            dr(0) = Coll.Key(i)
            dr(1) = Coll(i)
            dt.Rows.Add(dr)
        Next

        Return dt
    End Function

    Public Class KeyValuePair
        Private cKey As String
        Private cValue As Object

        Public Sub New(ByVal Key As String, ByVal Value As Object)
            cKey = Key
            cValue = Value
        End Sub
        Public Property Key() As String
            Get
                Return cKey
            End Get
            Set(ByVal value As String)
                cKey = value
            End Set
        End Property

        Public Property Value() As Object
            Get
                Return cValue
            End Get
            Set(ByVal value As Object)
                cValue = value
            End Set
        End Property
    End Class
End Class

Public Class CollXError
    Inherits Exception
    Private cMessage As String

    Public Sub New(ByVal Message As String)
        cMessage = Message
    End Sub

    Public Overrides ReadOnly Property Message() As String
        Get
            Return cMessage
        End Get
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

' Alternate approach
'Public Class xxxxCollXtreme
'    Dim CollX() As KeyValuePair

'    Public Sub Add(ByVal Key As String, ByVal Value As Object)
'        Dim Found As Boolean

'        If CollX IsNot Nothing Then
'            For Each Item In CollX
'                If Item.Key = Key Then
'                    Item.Value = Value
'                    Found = True
'                End If
'            Next
'        End If

'        If Not Found Then
'            ReDim Preserve CollX(0)
'            CollX(0) = New KeyValuePair(Key, Value)
'        End If
'    End Sub

'    Public Class KeyValuePair
'        Private cKey As String
'        Private cValue As Object

'        Public Property Key() As String
'            Get
'                Return cKey
'            End Get
'            Set(ByVal value As String)
'                cKey = value
'            End Set
'        End Property

'        Public Property Value() As Object
'            Get
'                Return cValue
'            End Get
'            Set(ByVal value As Object)
'                cValue = value
'            End Set
'        End Property

'        Public Sub New(ByVal Key As String, ByVal Value As Object)
'            cKey = Key
'            cValue = Value
'        End Sub
'    End Class
'End Class
'<input 
'name="ctl00$MainContent$txtAnnualContribution" 
'type="text" 
'value="520.00" 
'id="MainContent_txtAnnualContribution" 
'onkeypress="return NumericOnlyDec(this, 7, event)" onkeyup="CalcPerPay(this)"
' />

Public Class CalendarX
    Private cCalID As String
    Private cCaption As String
    Private cValue As String
    Private cIsVisible As Boolean = True
    Private cIsEditable As Boolean
    Private cTabIndex As Integer

    Public Sub New()
    End Sub

    Public Sub New(ByVal CalID As String)
        cCalID = CalID
    End Sub

    Public Sub New(ByVal CalID As String, ByVal Caption As String, ByVal Value As String, ByVal IsEditable As Boolean, ByVal TabIndex As Integer)
        cCalID = CalID
        cCaption = Caption
        cValue = Value
        cIsEditable = IsEditable
        cTabIndex = TabIndex
    End Sub

    Public Function Render() As System.Text.StringBuilder
        Return GetCalendarX(cCalID, cCaption, cValue, cIsVisible, cIsEditable, cTabIndex)
    End Function

    Private Function GetCalendarX(ByVal CalID As String, ByVal Caption As String, ByVal Value As String, ByVal IsVisible As Boolean, ByVal IsEditable As Boolean, ByVal TabIndex As Integer) As System.Text.StringBuilder
        Dim sb As New System.Text.StringBuilder
        Dim Editable As String = "0"
        Dim Visible As String = "1"

        Try
            If IsVisible Then
                Visible = "1"
            Else
                Visible = "0"
            End If
            If IsEditable Then
                Editable = "1"
            Else
                TabIndex = "-1"
            End If

            sb.Append("<script type=""text/javascript""> ")
            'sb.Append("new CalendarX(""" & CalID & """, """ & Caption & """, """ & Value & """, """ & Editable & """,  """ & TabIndex & """) ")
            sb.Append("new CalendarX(""" & CalID & """, """ & Caption & """, """ & Value & """, """ & Visible & """, """ & Editable & """,  """ & TabIndex.ToString & """) ")
            sb.Append("</script> ")

            Return sb

        Catch ex As Exception
            Throw New Exception("Error #2234: Common.GetCalendarX. " & ex.Message)
        End Try
    End Function

    Public WriteOnly Property CalID() As String
        Set(ByVal value As String)
            cCalID = value
        End Set
    End Property
    Public WriteOnly Property Caption() As String
        Set(ByVal value As String)
            cCaption = value
        End Set
    End Property
    Public Property Value() As String
        Get
            Return cValue
        End Get
        Set(ByVal value As String)
            cValue = value
        End Set
    End Property
    Public WriteOnly Property Visible() As Boolean
        Set(ByVal value As Boolean)
            cIsVisible = value
        End Set
    End Property
    Public WriteOnly Property IsEditable() As Boolean
        Set(ByVal value As Boolean)
            cIsEditable = value
        End Set
    End Property
    Public WriteOnly Property [ReadOnly]() As Boolean
        Set(ByVal value As Boolean)
            cIsEditable = Not value
        End Set
    End Property

    Public WriteOnly Property TabIndex() As Integer
        Set(ByVal value As Integer)
            cTabIndex = value
        End Set
    End Property
    Public ReadOnly Property ID() As String
        Get
            Return cCalID
        End Get
    End Property
End Class

Public Class LabelX
    Private cID As String
    Private cText As String

    Public Sub New(ByVal ID As String)
        cID = ID
    End Sub
    Public Sub New(ByVal ID As String, ByVal Text As String)
        cID = ID
        cText = Text
    End Sub

    Public Function Render() As System.Text.StringBuilder
        Return GetLabelX(cID, cText)
    End Function

    Private Function GetLabelX(ByVal ID As String, ByVal Text As String) As System.Text.StringBuilder
        Dim sb As New System.Text.StringBuilder

        Try
            sb.Append("<span id=""" & ID & """>" & Text & "</span>")
            Return sb

        Catch ex As Exception
            Throw New Exception("Error #4001: TextboxX.GetTextboxX. " & ex.Message)
        End Try
    End Function

    Public WriteOnly Property ID() As String
        Set(ByVal value As String)
            cID = value
        End Set
    End Property
    Public Property Text() As String
        Get
            Return cText
        End Get
        Set(ByVal value As String)
            cText = value
        End Set
    End Property
End Class

Public Class TextboxX
    Private cID As String
    Private cText As String = String.Empty
    Private cMaxLength As Integer = -1
    Private cIsReadOnly As Boolean
    Private cTabIndex As Integer
    Private cEvents As String
    Private cClassName As String
    Private cIsPassword As Boolean

    Public Sub New(ByVal ID As String)
        cID = ID
    End Sub

    Public Sub New(ByVal ID As String, ByVal Text As String, ByVal MaxLength As Integer, ByVal IsReadOnly As Boolean, ByVal Events As String, ByVal ClassName As String)
        cID = ID
        cText = Text
        cMaxLength = MaxLength
        cIsReadOnly = IsReadOnly
        cEvents = Events
        ClassName = ClassName
    End Sub

    Public Function Render() As System.Text.StringBuilder
        Return GetTextboxX(cID, cText, cMaxLength, cIsReadOnly, cTabIndex, cEvents, cClassName, cIsPassword)
    End Function

    Private Function GetTextboxX(ByVal ID As String, ByVal Text As String, ByVal MaxLength As Integer, ByVal IsReadOnly As Boolean, ByVal TabIndex As Integer, ByVal Events As String, ByVal ClassName As String, ByVal IsPassword As Boolean) As System.Text.StringBuilder
        Dim sb As New System.Text.StringBuilder
        Dim Type As String

        Try

            If IsPassword Then
                Type = "password"
            Else
                Type = "text"
            End If

            sb.Append("<input type=""" & Type & """ id=""" & ID & """ name=""" & ID & """  ")
            sb.Append("value=""" & Text & """ ")

            If MaxLength > -1 Then
                sb.Append("maxlength=""" & MaxLength & """ ")
            End If

            If IsReadOnly Then
                ClassName = "no_edit"
            End If
            If ClassName <> Nothing AndAlso ClassName.Length > 0 Then
                sb.Append("class=""" & ClassName & """ ")
            End If

            If IsReadOnly Then
                sb.Append("readonly=""readonly"" ")
                sb.Append("tabindex=""-1"" ")
            ElseIf TabIndex <> Nothing Then
                sb.Append("tabindex=""" & TabIndex & """ ")
            End If

            If Events <> Nothing Then
                sb.Append(Events & " ")
            End If

            'sb.Append("</input>")
            sb.Append(" />")

            Return sb

        Catch ex As Exception
            Throw New Exception("Error #4001: TextboxX.GetTextboxX. " & ex.Message)
        End Try
    End Function

    Public Property Text() As String
        Get
            Return cText
        End Get
        Set(ByVal value As String)
            cText = value
        End Set
    End Property
    Public WriteOnly Property MaxLength() As Integer
        Set(ByVal value As Integer)
            cMaxLength = value
        End Set
    End Property

    Public WriteOnly Property IsReadOnly() As Boolean
        Set(ByVal value As Boolean)
            cIsReadOnly = value
        End Set
    End Property
    Public WriteOnly Property TabIndex() As Integer
        Set(ByVal value As Integer)
            cTabIndex = value
        End Set
    End Property
    Public WriteOnly Property Events() As String
        Set(ByVal value As String)
            cEvents = value
        End Set
    End Property
    Public WriteOnly Property ClassName() As String
        Set(ByVal value As String)
            cClassName = value
        End Set
    End Property
    Public Property IsPassword() As Boolean
        Get
            Return cIsPassword
        End Get
        Set(ByVal value As Boolean)
            cIsPassword = value
        End Set
    End Property
End Class

Public Class CheckboxX
    Private cID As String
    Private cContainerObjName As String = Nothing
    Private cIsReadOnly As Boolean
    Private cIsChecked As Boolean
    Private cOnClick As String = Nothing
    Private cValue2 As String = Nothing

    Public Sub New(ByVal ID As String)
        cID = ID
    End Sub

    Public Sub New(ByVal ID As String, ByVal ContainerObjName As String, ByVal IsReadOnly As Boolean, ByVal IsChecked As Boolean, ByVal OnClick As String, ByVal Value2 As String)
        cID = ID
        cContainerObjName = ContainerObjName
        cIsReadOnly = IsReadOnly
        cIsChecked = IsChecked
        cOnClick = OnClick
        cValue2 = Value2
    End Sub

    Public Function Render() As System.Text.StringBuilder
        Return GetCheckboxX(cContainerObjName, cID, cIsReadOnly, cIsChecked, cOnClick, cValue2)
    End Function

    Private Function GetCheckboxX(ByVal ContainerObjName As String, ByVal ID As String, ByVal IsReadOnly As Boolean, ByVal IsChecked As Boolean, ByVal OnClick As String, ByVal Value2 As String) As System.Text.StringBuilder
        Dim sb As New System.Text.StringBuilder
        Dim ClassName As String
        Dim ClickAdj As String = Nothing
        Dim OnClickEvent As Boolean
        Dim chkXClick As String

        Try

            If IsReadOnly Then
                If IsChecked Then
                    ClassName = "checkbox_readonly_on"
                Else
                    ClassName = "checkbox_readonly_off"
                End If
            Else
                If IsChecked Then
                    ClassName = "checkbox_on"
                Else
                    ClassName = "checkbox_off"
                End If
            End If

            sb.Append("<a id=""" & ID & """ class=""" & ClassName & """ ")

            If OnClick <> Nothing Then
                OnClickEvent = True
            End If

            If ContainerObjName = Nothing Then
                chkXClick = "chkXClick('" & ID & "')"
            Else
                chkXClick = ContainerObjName & ".chkXClick('" & ID & "')"
            End If

            Select Case IsReadOnly
                Case True
                    Select Case OnClickEvent
                        Case True
                            ClickAdj = " onclick=""" & OnClick & """ "
                    End Select
                Case False
                    Select Case OnClickEvent
                        Case True
                            ClickAdj = " onclick=""" & chkXClick & "; " & OnClick & """ "
                        Case False
                            ClickAdj = " onclick=""" & chkXClick & """ "
                    End Select
            End Select

            If ClickAdj <> Nothing Then
                sb.Append(ClickAdj)
            End If

            sb.Append("></a>")

            If IsChecked Then
                sb.Append("<input type=""hidden"" id=""hd" & ID & """ name=""hd" & ID & """ value=""1"" /> ")
            Else
                sb.Append("<input type=""hidden"" id=""hd" & ID & """ name=""hd" & ID & """ value=""0"" /> ")
            End If

            If Value2 <> Nothing Then
                sb.Append("<input type=""hidden"" id=""hd" & ID & "2"" name=""hd" & ID & "2"" value=""" & Value2 & """ /> ")
            End If

            Return sb
        Catch ex As Exception
            Throw New Exception("Error #4403: CheckboxX.GetCheckboxX. " & ex.Message)
        End Try
    End Function

    Public Property IsReadOnly() As Boolean
        Get
            Return cIsReadOnly
        End Get
        Set(ByVal value As Boolean)
            cIsReadOnly = value
        End Set
    End Property
    Public Property IsChecked() As Boolean
        Get
            Return cIsChecked
        End Get
        Set(ByVal value As Boolean)
            cIsChecked = value
        End Set
    End Property
    Public Property OnClick() As String
        Get
            Return cOnClick
        End Get
        Set(ByVal value As String)
            cOnClick = value
        End Set
    End Property
End Class

Public Class DropdownX
    Private cID As String
    Private cSelectedValue As String
    Private cText As String
    Private cColl As New CollX
    Private cIsReadOnly As Boolean
    Private cIsVisible As Boolean = True
    Private cOnClick As String
    Private cTabIndex As Integer

    Public Function Render() As System.Text.StringBuilder
        Return GetDropdownX(cID, cSelectedValue, cColl, cIsReadOnly, cIsVisible, cTabIndex, cOnClick)
    End Function

    Private Function GetDropdownX(ByVal FldName As String, ByVal SelectedValue As String, ByVal Coll As CollX, ByVal IsReadonly As Boolean, ByVal IsVisible As Boolean, ByVal TabIndex As Integer, ByVal OnClick As String) As System.Text.StringBuilder
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        Dim Idx As Integer = 1
        Dim SelectedText As String = String.Empty
        Dim Text As String = Nothing
        Dim Value As String = Nothing
        Dim ClickAdj As String

        Try

            If Not IsVisible Then
                Return sb
            End If

            ' ___ Find selected
            If Common.IsNotBlank(SelectedValue) Then
                For i = 1 To Coll.Count
                    If Coll.Key(i) = SelectedValue Then
                        Idx = i
                        SelectedText = Coll(i)
                        Exit For
                    End If
                Next
            End If

            If IsReadonly Then
                TabIndex = -1
                sb.Append("<div class=""input_field""><input class=""no_edit"" type=""text"" readonly=""readonly"" tabindex=""-1"" value=""" & SelectedText & """ /></div>" & Environment.NewLine)
            Else

                sb.Append("<div class=""input_field"">")

                ' ___ Selected
                sb.Append("<div class=""dropdownX"">")
                sb.Append("<a class=""dropdownX-field"" ></a><input type=""text"" id=""" & FldName & "Select"" readonly=""readonly"" tabindex=""" & TabIndex.ToString & """ value=""" & SelectedText & """ />")

                ' ___ List
                sb.Append("<div class=""dropdownX-list"" id=""" & FldName & "List"" style=""display:none""><ul>")
                For i = 1 To Coll.Count
                    Value = Coll.Key(i)
                    Text = Coll(i)

                    If OnClick = Nothing Then
                        ClickAdj = " onclick=""ddXClick('" & FldName & "', '" & Value & "',  '" & Text & "')"" "
                    Else
                        ClickAdj = " onclick=""ddXClick('" & FldName & "', '" & Value & "',  '" & Text & "'); " & OnClick & """ "
                    End If

                    If Value = SelectedValue Then
                        sb.Append("<li id=""" & Value & """><a class=""dropdownX-selected""" & ClickAdj & ">" & Coll(i) & "</a></li>")
                    Else
                        sb.Append("<li id=""" & Value & """><a " & ClickAdj & ">" & Text & "</a></li>")
                    End If

                Next
                sb.Append("</ul></div></div>")
                sb.Append("</div>")
            End If

            ' ___ Hidden
            sb.Append("<input type=""hidden"" id=""hd" & FldName & """ name=""hd" & FldName & """ value=""" & SelectedValue & """  />")
            sb.Append("<input type=""hidden"" id=""hd" & FldName & "Text"" name=""hd" & FldName & "Text"" value=""" & SelectedText & """  />")

            Return sb

        Catch ex As Exception
            Throw New Exception("Error #3801: DropdownX.GetDropdownX. " & ex.Message)
        End Try
    End Function

    Public ReadOnly Property View() As String
        Get
            Return "ID: " & cID & "|SelectedValue: " & cSelectedValue & "|IsReadOnly " & CType(cIsReadOnly, System.String) & "|Coll.Count: " & cColl.Count
        End Get
    End Property

    Public Property ID() As String
        Get
            Return cID
        End Get
        Set(ByVal value As String)
            cID = value
        End Set
    End Property

    Public Property SelectedValue() As String
        Get
            Return cSelectedValue
        End Get
        Set(ByVal value As String)
            cSelectedValue = value
        End Set
    End Property
    Public WriteOnly Property IsReadOnly() As Boolean
        Set(ByVal value As Boolean)
            cIsReadOnly = value
        End Set
    End Property
    Public WriteOnly Property [ReadOnly]() As Boolean
        Set(ByVal value As Boolean)
            cIsReadOnly = value
        End Set
    End Property
    Public WriteOnly Property Visible() As Boolean
        Set(ByVal value As Boolean)
            cIsVisible = value
        End Set
    End Property
    Public WriteOnly Property TabIndex() As Integer
        Set(ByVal value As Integer)
            cTabIndex = value
        End Set
    End Property
    Public Property Coll() As CollX
        Get
            Return cColl
        End Get
        Set(ByVal value As CollX)
            cColl = value
        End Set
    End Property
    Public WriteOnly Property OnClick() As String
        Set(ByVal value As String)
            cOnClick = value
        End Set
    End Property
End Class

Public Class VarX
    Private cID As String
    Private cNameAttributeInd As Boolean
    Private cValue As String = String.Empty

    Public Sub New(ByVal ID As String, ByVal NameAttributeInd As Boolean)
        cID = ID
        cNameAttributeInd = NameAttributeInd
    End Sub

    Public Sub New(ByVal ID As String, ByVal NameAttributeInd As Boolean, ByVal Value As String)
        cID = ID
        cNameAttributeInd = NameAttributeInd
        cValue = Value
    End Sub
    Public Function Render() As System.Text.StringBuilder
        Return GetVarX(cID, cNameAttributeInd, cValue)
    End Function

    Private Function GetVarX(ByVal ID As String, ByVal NameAttributeInd As String, ByVal Value As String) As System.Text.StringBuilder
        Dim sb As New System.Text.StringBuilder

        Try

            If NameAttributeInd Then
                sb.Append("<input type=""hidden"" id=""hd" & ID & """ name=""hd" & ID & """ value=""" & Value & """ /> ")
            Else
                sb.Append("<input type=""hidden"" id=""hd" & ID & """ value=""" & Value & """ /> ")
            End If

            Return sb
        Catch ex As Exception
            Throw New Exception("Error #4404: VarX.GetVarX. " & ex.Message)
        End Try
    End Function

    Public Property Value() As String
        Get
            Return cValue
        End Get
        Set(ByVal value As String)
            cValue = value
        End Set
    End Property
End Class