' Validate number of active emps in callbackmaster. Multiple nonpurged emps cause usp_AcesCloseoutCallback to close out the wrong record.
'use Callback

'Declare @StartDate datetime
'SET @StartDate = '7/31/2012'

'select cm1.empid, count(cm2.empid)
'from callbackmaster cm1
'inner join callbackmaster cm2 on
'cm1.empid = cm2.empid and
'cm1.adddate > @StartDate and
'cm2.adddate > @StartDate and
'isnull(cm1.ispurged, 0) = 0 and
'isnull(cm2.ispurged, 0) = 0
'group by cm1.empid
'having COUNT(cm2.empid) >  1

Public Class Results
    Private cSuccess As Boolean
    Private cMessage As String
    Private cObtainConfirm As Boolean
    Private cValue As Object
    Private cValue2 As Object

    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property
    Public Property Message() As String
        Get
            Return cMessage
        End Get
        Set(ByVal Value As String)
            cMessage = Value
        End Set
    End Property
    Public Property ObtainConfirm() As Boolean
        Get
            Return cObtainConfirm
        End Get
        Set(ByVal Value As Boolean)
            cObtainConfirm = Value
        End Set
    End Property
    Public Property Value() As Object
        Get
            Return cValue
        End Get
        Set(ByVal Value As Object)
            cValue = Value
        End Set
    End Property
    Public Property Value2() As Object
        Get
            Return cValue2
        End Get
        Set(ByVal Value As Object)
            cValue2 = Value
        End Set
    End Property
End Class


Public Class RequestActionResults
    Private cResponseAction As ResponseActionEnum
    Private cMessage As String
    Private cObtainConfirm As Boolean

    Public Property ResponseAction() As ResponseActionEnum
        Get
            Return cResponseAction
        End Get
        Set(ByVal value As ResponseActionEnum)
            cResponseAction = value
        End Set
    End Property
    Public Property Message() As String
        Get
            Return cMessage
        End Get
        Set(ByVal Value As String)
            cMessage = Value
        End Set
    End Property
    Public Property ObtainConfirm() As Boolean
        Get
            Return cObtainConfirm
        End Get
        Set(ByVal value As Boolean)
            cObtainConfirm = value
        End Set
    End Property
End Class

Public Class Style
    Public Enum StyleType
        NormalEditable = 1
        NoneditableGrayed = 3
        NoneditableWhite = 2
        NotVisible = 4
    End Enum

    Public Shared Sub AddCustomStyle(ByVal tb As TextBox, ByRef Elements As Collection, ByVal Visible As Boolean, ByVal [ReadOnly] As Boolean)
        Dim StyleStr As String = Nothing
        Dim i As Integer
        For i = 1 To Elements.Count
            StyleStr &= Elements(i).Key & ":" & Elements(i).Value & ";"
        Next
        tb.Attributes.Add("style", StyleStr)
        tb.Visible = Visible
        tb.ReadOnly = [ReadOnly]
    End Sub

    Public Shared Sub AddStyle(ByVal tb As TextBox, ByVal StyleType As StyleType, ByVal Width As Integer, Optional ByVal IsTextArea As Boolean = False)
        If Not IsTextArea Then
            Select Case StyleType
                Case StyleType.NormalEditable
                    tb.Attributes.Add("style", "width:" & Width & "px;")
                    tb.Visible = True
                    tb.ReadOnly = False

                Case StyleType.NoneditableGrayed
                    tb.Attributes.Add("style", "width:" & Width & "px;border-width:1px;background: #eeeedd;readOnly: true;")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NoneditableWhite
                    tb.Attributes.Add("style", "width:" & Width & "px;border-width:1px;border-style:solid;border-color:cccccc;background: #ffffff")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NotVisible
                    tb.Visible = False
            End Select
        Else                                                                ' is a textarea
            Select Case StyleType
                Case StyleType.NormalEditable
                    tb.Attributes.Add("style", "width:" & Width & "px;FONT: 10pt Arial, Helvetica, sans-serif;")
                    tb.Visible = True
                    tb.ReadOnly = False

                Case StyleType.NoneditableGrayed
                    ' tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:hidden;readOnly: true")
                    tb.Attributes.Add("style", "width:" & Width & "px;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:auto;readOnly: true")
                    tb.Visible = True
                    tb.ReadOnly = True


                Case StyleType.NoneditableWhite
                    'tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;overflow:hidden;readOnly: true")
                    tb.Attributes.Add("style", "width:" & Width & "px;FONT: 10pt Arial, Helvetica, sans-serif;overflow:auto;readOnly: true")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NotVisible
                    tb.Visible = False
            End Select
        End If

    End Sub
    Public Shared Sub AddStyle(ByVal tb As TextBox, ByVal IsVisible As Boolean, ByVal IsReadOnly As Boolean, ByVal Width As Integer, Optional ByVal IsTextArea As Boolean = False)
        tb.Visible = IsVisible
        tb.ReadOnly = IsReadOnly
        If Not IsTextArea Then
            If IsVisible Then
                If IsReadOnly Then
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;background: #eeeedd;readOnly: true;")
                    tb.ReadOnly = True
                Else
                    'tb.Attributes.Add("style", "width:" & Width & ";background: #ffffff;")
                    tb.Attributes.Add("style", "width:" & Width & ";")
                End If
            Else
                ' tb.Attributes.Add("style", "VISIBILITY: hidden")
                tb.Attributes.Add("style", "display: none")
            End If
        Else
            If IsVisible Then
                If IsReadOnly Then
                    'tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:hidden;readOnly: true")
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:auto;readOnly: true")
                    tb.ReadOnly = True
                Else
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;")
                End If
            End If
        End If
    End Sub
End Class


Public Class ErrorArgs
    Private cHeaderMessage As String
    Private cErrorMessage As String
 
    Public Property HeaderMessage() As String
        Get
            Return cHeaderMessage
        End Get
        Set(ByVal Value As String)
            cHeaderMessage = Value
        End Set
    End Property
    Public Property ErrorMessage() As String
        Get
            Return cErrorMessage
        End Get
        Set(ByVal Value As String)
            cErrorMessage = Value
        End Set
    End Property
End Class

Public Class AcesInterfaceIssue
    Private cClientID As String
    Private cErrorMessage As String
    Private cErrorType As ErrorTypeEnum

    Public Enum ErrorTypeEnum
        EmpTransmittalDataChanged = 1
        EmpProductTransmittalDataChanged = 2
    End Enum

    Public Property ClientID() As String
        Get
            Return cClientID
        End Get
        Set(ByVal Value As String)
            cClientID = Value
        End Set
    End Property
    Public Property ErrorMessage() As String
        Get
            Return cErrorMessage
        End Get
        Set(ByVal Value As String)
            cErrorMessage = Value
        End Set
    End Property
    Public Property ErrorType() As ErrorTypeEnum
        Get
            Return cErrorType
        End Get
        Set(ByVal Value As ErrorTypeEnum)
            cErrorType = Value
        End Set
    End Property
End Class

Public Class DataChangedPack
    Private cFldName As String
    Private cTableData As String
    Private cPageData As String

    Public Sub New(ByVal FldName As String, ByVal TableData As String, ByVal PageData As String)
        cFldName = FldName
        cTableData = TableData
        cPageData = PageData
    End Sub
    Public ReadOnly Property FldName() As String
        Get
            Return cFldName
        End Get
    End Property
    Public ReadOnly Property TableData() As String
        Get
            Return cTableData
        End Get
    End Property
    Public ReadOnly Property PageData() As String
        Get
            Return cPageData
        End Get
    End Property
End Class

'Public Class CmdAsst
'    Private cSqlCmd As SqlClient.SqlCommand
'    Private cAppSession As AppSession
'    Private cDBName As String

'    Public Sub New(ByRef AppSession As AppSession, ByVal CmdType As CommandType, ByRef SPNameOrSql As String, Optional ByVal DBName As String = Nothing)
'        cAppSession = AppSession
'        cSqlCmd = New SqlClient.SqlCommand(SPNameOrSql)
'        cSqlCmd.CommandType = CmdType
'        cDBName = DBName
'    End Sub

'#Region " Add Parameters "
'    Public Sub AddInt(ByVal VarName As String, ByVal Value As Object)
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.ParameterDirection.Input, False, CType(10, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddBit(ByVal VarName As String, ByVal Value As Object)
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Bit, 1"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddDateTime(ByVal VarName As String, ByVal Value As Object) ' datetime
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.DateTime, 8"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddMoney(ByVal VarName As String, ByVal Value As Object) ' decimal
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Money, 4, System.Data.ParameterDirection.Input, False, CType(19, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddVarChar(ByVal VarName As String, ByVal Value As Object, ByVal Length As Integer)
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.VarChar, "))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'    Public Sub AddFloat(ByVal VarName As String, ByVal Value As Object) ' double
'        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Float, 8, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
'        cSqlCmd.Parameters("@" & VarName).Value = Value
'    End Sub
'#End Region

'    Public Function Execute() As DBase.QueryPack
'        Dim QueryPack As New DBase.QueryPack
'        Dim DataAdapter As SqlDataAdapter
'        Dim dt As New DataTable

'        If cDBName = Nothing Then
'            cSqlCmd.Connection = New SqlConnection(cAppSession.GetConnectionString)
'        Else
'            cSqlCmd.Connection = New SqlConnection(cAppSession.GetConnectionString(cAppSession.DBHost, cDBName))
'        End If
'        DataAdapter = New SqlDataAdapter(cSqlCmd)
'        Try
'            DataAdapter.Fill(dt)
'            QueryPack.Success = True
'            QueryPack.dt = dt
'        Catch ex As Exception
'            QueryPack.Success = False
'            QueryPack.TechErrMsg = ex.Message
'        End Try
'        DataAdapter.Dispose()
'        cSqlCmd.Dispose()
'        Return QueryPack
'    End Function
'End Class

Public Class KeyValueObj
    Private cKey As String
    Private cValue As String

    Public Property Key() As String
        Get
            Return cKey
        End Get
        Set(ByVal Value As String)
            cKey = Value
        End Set
    End Property
    Public Property Value() As String
        Get
            Return cValue
        End Get
        Set(ByVal Value As String)
            cValue = Value
        End Set
    End Property
    Public Sub New(ByVal Key As String, ByVal Value As String)
        cKey = Key
        cValue = Value
    End Sub
End Class

Public Class AESEncryption
    Public Shared Function AES_Encrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim Encrypted As String
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            'AES.Mode = Security.Cryptography.CipherMode.ECB
            AES.Mode = System.Security.Cryptography.CipherMode.ECB
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
            Encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return Encrypted
        Catch ex As Exception
            Throw New Exception("Error #8601: AES_Encryption.AES_Encrypt. " & ex.Message)
        End Try
    End Function

    Public Shared Function AES_Decrypt(ByVal input As String, ByVal pass As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim Decrypted As String
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = System.Security.Cryptography.CipherMode.ECB
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = Convert.FromBase64String(input)
            Decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return Decrypted
        Catch ex As Exception
            Throw New Exception("Error #8602: AES_Encryption.AES_Decrypt. " & ex.Message)
        End Try
    End Function
End Class
