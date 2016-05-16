'Author:		<Sanjeev Kochumon>
'Testing
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.Data
Imports System.Configuration

#Region "DB Utils"

Public Class clsConnection
    Public db_Connection As New SqlConnection
    Public gEBSCnn As New OracleClient.OracleConnection
    Public cmd As SqlCommand
    Private strQry As String
    Public dr As SqlDataReader
    Public adpt As SqlDataAdapter
    Private connStr As String
    Private ebsConnStr As String
    Public Function getEBSConnection() As Boolean
        ' DATA SOURCE=crptest.alabbargroup.com:1525/CRPDEV;PERSIST SECURITY INFO=True;USER ID=APPS
        If gEBSCnn Is Nothing Then gEBSCnn = New OracleClient.OracleConnection
        If gEBSCnn.State = ConnectionState.Broken Then gEBSCnn.Close()
        If gEBSCnn.State = ConnectionState.Closed Then
            gEBSCnn.ConnectionString = ebsConnStr
            Try
                gEBSCnn.Open()
                getEBSConnection = True
            Catch ex As SqlException
                getEBSConnection = False
            End Try
        Else
            getEBSConnection = True
        End If
    End Function
    Public Sub closeEBSConnection()
        If gEBSCnn Is Nothing Then Exit Sub
        If gEBSCnn.State = ConnectionState.Open Then
            gEBSCnn.Close()
            gEBSCnn = Nothing
        End If
    End Sub
    Public Function getDataTableOracle(ByVal strQry As String) As DataTable
        Try
            If getEBSConnection() Then
                Dim cmd As New OracleClient.OracleCommand(strQry, gEBSCnn)
                Dim dr As OracleClient.OracleDataReader = cmd.ExecuteReader
                Dim dt As New DataTable
                dt.Load(dr)
                closeEBSConnection()
                Return dt
            End If
        Catch ex As Exception
            Return Nothing
        Finally
            closeEBSConnection()
        End Try
    End Function
    Public Sub New()
        connStr = ConfigurationManager.ConnectionStrings("IntegraConnectionString").ConnectionString
        ebsConnStr = ConfigurationManager.ConnectionStrings("EBSConnectionString").ConnectionString
    End Sub
    Public WriteOnly Property Query As String
        Set(value As String)
            strQry = value
        End Set
    End Property
    Public Function getConnection() As Boolean
        If db_Connection.State <> ConnectionState.Open Then 'If connection closed
            db_Connection.ConnectionString = connStr
            Try
                db_Connection.Open()
                getConnection = True
            Catch ex As Exception
                getConnection = False
            End Try
        Else
            getConnection = True
        End If
    End Function

    Public Sub closeConnection()
        If db_Connection.State = ConnectionState.Open Then db_Connection.Close()
    End Sub
    Public Function getSingleValue() As Object
        Dim strVal As Object
        If getConnection() Then
            cmd = New SqlCommand(strQry, db_Connection)
            Try
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    dr.Read()
                    strVal = dr.Item(0)
                End If
                dr.Close()
                closeConnection()
            Catch ex As Exception

            End Try
        End If
        getSingleValue = strVal
    End Function
    Public Function getDataTable() As DataTable
        If getConnection() Then
            cmd = New SqlCommand(strQry, db_Connection)
            adpt = New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            adpt.Fill(dt)
            closeConnection()
            Return dt
        End If
    End Function
    Public Sub ExecuteQuery()
        If getConnection() Then
            cmd = New SqlCommand(strQry, db_Connection)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Sub
    Public Function GetLatestNo() As Integer
        GetLatestNo = 1
        If getConnection() Then
            Dim cmd As New SqlClient.SqlCommand(strQry, db_Connection)
            Dim rdSer As SqlClient.SqlDataReader
            rdSer = cmd.ExecuteReader
            With rdSer
                Try
                    If .Read() Then
                        GetLatestNo = .Item(0) + 1
                    Else
                        GetLatestNo = 1
                    End If
                Catch ex As Exception
                End Try
                .Close()
                cmd.Dispose()
            End With
            closeConnection()
        End If
    End Function
End Class
#End Region

#Region "Integra"
Public Class clsIntegra
    Inherits clsConnection
    Private strUname As String
    Private strPassword As String
    Private appusername As String
    Private intUserId As Integer
    Private bolApprove As Boolean = False
    Private bolMTOAccess As Boolean = False
    Private bolEBSCostAccess As Boolean = False
    Private bolIsAdmin As Boolean = False

    Public WriteOnly Property AuthenUsername As String
        Set(value As String)
            strUname = value
        End Set
    End Property
    Public WriteOnly Property AuthenPassword As String
        Set(value As String)
            strPassword = value
        End Set
    End Property
    Public WriteOnly Property Username As String
        Set(value As String)
            appusername = value
        End Set
    End Property
    Public WriteOnly Property UserID As Integer
        Set(value As Integer)
            intUserId = value
        End Set
    End Property
    Public Function GetFullName() As String
        Dim strQry = "select full_name from users where uid=" & intUserId
        Query = strQry
        Return getSingleValue()
    End Function
    Public ReadOnly Property Approver As Boolean
        Get
            Return bolApprove
        End Get
    End Property
    Public ReadOnly Property MTOAccess As Boolean
        Get
            Return bolMTOAccess
        End Get
    End Property
    Public ReadOnly Property EBSCostAccess As Boolean
        Get
            Return bolEBSCostAccess
        End Get
    End Property
    Public ReadOnly Property IsAdmin As Boolean
        Get
            Return bolIsAdmin
        End Get
    End Property
    Public Function GetUserRights() As DataTable
        Dim dt As New DataTable
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandType = Data.CommandType.StoredProcedure
            cmd.CommandText = "Web_Get_Users_Access"
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@userid"
            param1.DbType = Data.DbType.Int32
            param1.Value = intUserId
            cmd.Parameters.Add(param1)
            Try
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    dt.Load(dr)
                    dr.Close()
                    cmd.Dispose()
                End If
            Catch ex As Exception
            End Try
            Return dt
        End If
    End Function
    Public Sub GetUserAccess()
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandType = Data.CommandType.StoredProcedure
            cmd.CommandText = "Web_Get_User_Rights"
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@userid"
            param1.DbType = Data.DbType.Int32
            param1.Value = intUserId
            cmd.Parameters.Add(param1)
            Try
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    dr.Read()
                    If dr.Item("MTOEnter") = True Then
                        bolMTOAccess = True
                    Else
                        bolMTOAccess = False
                    End If
                    If dr.Item("MTOApprove") = True Then
                        bolApprove = True
                    Else
                        bolApprove = False
                    End If
                    If dr.Item("show_ebs_cost") = True Then
                        bolEBSCostAccess = True
                    Else
                        bolEBSCostAccess = False
                    End If
                    If dr.Item("is_admin") = True Then
                        bolIsAdmin = True
                    Else
                        bolIsAdmin = False
                    End If
                End If
                dr.Close()
                cmd.Dispose()
                closeConnection()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Function GetUserName() As String
        Dim strUsername As String = ""
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_User_Validate"
            cmd.Connection = db_Connection
            cmd.CommandType = Data.CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@id"
            param1.DbType = Data.DbType.String
            param1.Value = strUname
            Dim param2 As New SqlParameter
            param2.ParameterName = "@pass"
            param2.DbType = Data.DbType.String
            param2.Value = Generate_Hash(strPassword, 2)
            Dim param3 As New SqlParameter
            param3.ParameterName = "@username"
            param3.DbType = Data.DbType.Int32
            param3.Direction = Data.ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            Try
                cmd.ExecuteScalar()
                strUsername = param3.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        Return strUsername
    End Function
    Public Function GetUsers() As DataTable
        Dim strQry As String
        strQry = "select uid,full_name from users where is_locked=0"
        Query = strQry
        Return getDataTable()
    End Function
    Public Function ValidateUser() As String
        Dim strQry As String, strUsername As String = ""
        If getConnection() Then
            strQry = "select username from users where username='" & strUname & "' and u_password='" & Generate_Hash(strPassword, 2) & "' and is_locked=0"
            cmd = New SqlCommand(strQry, db_Connection)
            Try
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    dr.Read()
                    strUsername = dr.Item(0)
                Else
                    strUsername = ""
                End If
            Catch ex As Exception

            End Try
            closeConnection()
            dr.Close()
            cmd.Dispose()
        End If
        Return strUsername
    End Function
    Public Function getEBSItems() As DataTable
        Dim strQry As String
        strQry = "select ITEMNMBR as ITEM,ITEMDESC as DESCRIPTION,ORACLE_CAT as CATEGORY,LENGTH as UNIT_LENGTH,WIDTH as UNIT_WIDTH,WEIGHT as UNIT_WEIGHT,BASEUOFM as PRIMARY_UOM_CODE," _
            & "CURRCOST as ITEM_COST,STNDCOST as LIST_PRICE from Oracle_Items where ORACLE_CAT<>'ASSEMBLY' and ORACLE_CAT<>'GLASS-ASSEMBLY'"
        Query = strQry
        Return getDataTable()
    End Function
    Public Function Generate_Hash(ByVal strPassword As String, Optional ByVal intHashType As Integer = 0) As String
        ' Create an Encoding object so that you can use the convenient GetBytes 
        ' method to obtain byte arrays.
        Dim uEncode As New System.Text.UnicodeEncoding()
        ' Create a byte array from the source text passed as an argument.
        Dim bytPassword() As Byte = uEncode.GetBytes(strPassword)

        ' The code is almost identical for all three hash types.
        Dim hash() As Byte

        Select Case intHashType
            Case 1
                ' MD5 hash value.
                Dim md5 As New Security.Cryptography.MD5CryptoServiceProvider()
                hash = md5.ComputeHash(bytPassword)
            Case 2
                ' SHA1 hash value.
                Dim sha1 As New Security.Cryptography.SHA1CryptoServiceProvider()
                hash = sha1.ComputeHash(bytPassword)
            Case 3
                ' SHA384 hash value.
                Dim sha384 As New Security.Cryptography.SHA384Managed()
                hash = sha384.ComputeHash(bytPassword)
            Case Else
                ' Default MD5 hash value.
                Dim md5 As New Security.Cryptography.MD5CryptoServiceProvider()
                hash = md5.ComputeHash(bytPassword)
        End Select
        ' Base64 is a method of encoding binary data as ASCII text.
        Return Convert.ToBase64String(hash)
    End Function
    Public Function GetUserID() As Integer
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandType = Data.CommandType.StoredProcedure
            cmd.CommandText = "Web_Get_UserID"
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@username"
            param1.DbType = Data.DbType.String
            param1.Value = appusername
            Dim param2 As New SqlParameter
            param2.ParameterName = "@uid"
            param2.DbType = Data.DbType.Int32
            param2.Direction = Data.ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Try
                cmd.ExecuteScalar()
                Return param2.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Function

End Class
#End Region

#Region "Oracle"
Public Class clsOracle
    Inherits clsConnection
    Public Function getEBSInventory() As DataTable
        Dim strQry As String
        strQry = "select ITEM,DESCRIPTION,round(UNIT_WEIGHT,2) as UNIT_WEIGHT,ORGANIZATION_ID,CATEGORY,PRIMARY_UOM_CODE,round(UNIT_LENGTH,2) as UNIT_LENGTH,LIST_PRICE_PER_UNIT as LIST_PRICE, " _
                & "round(UNIT_WIDTH,2) as UNIT_WIDTH,round(ITEM_COST,2) as ITEM_COST,SUBINVENTORY_CODE,round(ONHAND,2) as ONHAND,LOCATOR,PROJECT_NUMBER,round(VALUE,4) as Value from APPS.XXAAB_ITEM_STOCK_VIEW " _
                & " where ITEM_TYPE='P'AND CATEGORY<>'GLASS-ASSEMBLY'"
        Return getDataTableOracle(strQry)
    End Function
End Class
#End Region

#Region "Material Take Off"
Public Class clsMTO
    Inherits clsConnection
    Friend refer_no As String
    Friend ext_ref As String
    Friend var_ref As String
    Friend loc As String
    Friend user_id As Integer
    Friend remarks As String
    Friend project_id As String
    Friend doc_ref As String
    Friend boq_id As Integer
    Friend area_code As Integer
    Friend zone_code As Integer
    Friend cut_length As Double
    Friend cut_width As Double
    Friend qty As Double
    Friend stock_code As String
    Friend boq_type As Integer
    Friend category As String
    Friend header_id As Integer
    Friend detail_id As Integer
    Friend revision_no As String
    Friend field_name As String
    Friend table_name As String
    Friend primary_field As String
    Friend revision_comment As String
    Friend item_code As String
    Friend cutsize_code As String
    Friend is_mor As Boolean
    Friend MORhdrid As Integer
    
    Public Sub New()
        MyBase.New()
    End Sub
   
    Public WriteOnly Property MTOReferenceNo As String
        Set(value As String)
            refer_no = value
        End Set
    End Property
    Public WriteOnly Property PrimaryField As String
        Set(value As String)
            primary_field = value
        End Set
    End Property
    Public WriteOnly Property ItemCode As String
        Set(value As String)
            item_code = value
        End Set
    End Property
    Public WriteOnly Property MORHeaderID As Integer
        Set(value As Integer)
            MORhdrid = value
        End Set
    End Property
    Public WriteOnly Property IsMOR As Boolean
        Set(value As Boolean)
            is_mor = value
        End Set
    End Property
    Public WriteOnly Property CutsizeCode As String
        Set(value As String)
            cutsize_code = value
        End Set
    End Property
    Public WriteOnly Property RevisionComment As String
        Set(value As String)
            revision_comment = value
        End Set
    End Property
    Public WriteOnly Property FieldName As String
        Set(value As String)
            field_name = value
        End Set
    End Property
    Public WriteOnly Property TableName As String
        Set(value As String)
            table_name = value
        End Set
    End Property
    Public WriteOnly Property RevisionNo As String
        Set(value As String)
            revision_no = value
        End Set
    End Property
    Public WriteOnly Property BoqType As String
        Set(value As String)
            boq_type = value
        End Set
    End Property
    Public WriteOnly Property MtoHeader As Integer
        Set(value As Integer)
            header_id = value
        End Set
    End Property
    Public WriteOnly Property MtoDetail As Integer
        Set(value As Integer)
            detail_id = value
        End Set
    End Property
    Public WriteOnly Property CutLength As String
        Set(value As String)
            cut_length = value
        End Set
    End Property
    Public WriteOnly Property CutWidth As String
        Set(value As String)
            cut_width = value
        End Set
    End Property
    Public WriteOnly Property StockCode As String
        Set(value As String)
            stock_code = value
        End Set
    End Property
    Public WriteOnly Property Quantity As String
        Set(value As String)
            qty = value
        End Set
    End Property
    Public WriteOnly Property ExternalRef As String
        Set(value As String)
            ext_ref = value
        End Set
    End Property
    Public WriteOnly Property VarianceRef As String
        Set(value As String)
            var_ref = value
        End Set
    End Property
    Public WriteOnly Property Location As String
        Set(value As String)
            loc = value
        End Set
    End Property
    Public WriteOnly Property UserID As Integer
        Set(value As Integer)
            user_id = value
        End Set
    End Property
    Public WriteOnly Property MTORemarks As String
        Set(value As String)
            remarks = value
        End Set
    End Property

    Public WriteOnly Property ProjectID As String
        Set(value As String)
            project_id = value
        End Set
    End Property
    Public WriteOnly Property DocumentRef As String
        Set(value As String)
            doc_ref = value
        End Set
    End Property
    Public WriteOnly Property BOQId As Integer
        Set(value As Integer)
            boq_id = value
        End Set
    End Property
    Public WriteOnly Property Area As Integer
        Set(value As Integer)
            area_code = value
        End Set
    End Property
    Public WriteOnly Property Zone As Integer
        Set(value As Integer)
            zone_code = value
        End Set
    End Property
    Public WriteOnly Property ItemCategory As String
        Set(value As String)
            category = value
        End Set
    End Property
    Public Function GetMORRefernce() As String
        Query = "select referno from mor_hdr where morhdr_id=" & MORhdrid & ""
        Return getSingleValue()
    End Function
    Public Function GetWeight(ByVal fltQty As Double, fltArea As Double) As Double
        Return ((Val(fltQty) * Val(fltArea)) / 1000) ' IN TON
    End Function
    Public Function GetSquareMeter(ByVal fltQty As Double, fltLength As Double, fltWidth As Double) As Double
        Return (Val(fltQty) * Val(fltLength) * Val(fltWidth))
    End Function
    Public Function GetCost(ByVal fltRate As Double, ByVal fltQty As Double) As Double
        Return (fltRate * fltQty)
    End Function
    Public Function GetFieldValue() As Object
        Query = "select " & field_name & " from " & table_name & " where " & primary_field & "=" & header_id
        Return getSingleValue()
    End Function

    Public Sub PerformCutsizeMaintainence(ByVal intCrud As Integer)
        '1. Insert
        '3. Update
        '4. Delete
        cmd = New SqlCommand
        cmd.Connection = db_Connection
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "Web_Create_CutSize"
        Dim param1 As New SqlParameter
        param1.ParameterName = "@itemcode"
        param1.DbType = DbType.String
        param1.Value = item_code
        Dim param2 As New SqlParameter
        param2.ParameterName = "@stockcode"
        param2.DbType = DbType.String
        param2.Value = stock_code
        Dim param3 As New SqlParameter
        param3.ParameterName = "@cutsizecode"
        param3.DbType = DbType.String
        param3.Value = cutsize_code
        Dim param4 As New SqlParameter
        param4.ParameterName = "@cutlength"
        param4.DbType = DbType.Double
        param4.Value = cut_length
        Dim param5 As New SqlParameter
        param5.ParameterName = "@cutwidth"
        param5.DbType = DbType.Double
        param5.Value = cut_width
        Dim param6 As New SqlParameter
        param6.ParameterName = "@uid"
        param6.DbType = DbType.Int32
        param6.Value = user_id
        Dim param7 As New SqlParameter
        param7.ParameterName = "@crud"
        param7.DbType = DbType.Int32
        param7.Value = intCrud
        cmd.Parameters.Add(param1)
        cmd.Parameters.Add(param2)
        cmd.Parameters.Add(param3)
        cmd.Parameters.Add(param4)
        cmd.Parameters.Add(param5)
        cmd.Parameters.Add(param6)
        cmd.Parameters.Add(param7)
        If getConnection() Then
            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Sub
    Public Function GetCutSizeCost() As Object
        If getConnection() Then
            cmd = New SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandText = "Get_Cut_Size_Cost"
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@cutlength"
            param1.DbType = DbType.Double
            param1.Value = cut_length
            Dim param2 As New SqlParameter
            param2.ParameterName = "@cutwidth"
            param2.DbType = DbType.Double
            param2.Value = cut_width
            Dim param3 As New SqlParameter
            param3.ParameterName = "@stockcode "
            param3.DbType = DbType.String
            param3.Value = stock_code
            Dim param4 As New SqlParameter
            param4.ParameterName = "@priceperpiece"
            param4.DbType = DbType.Double
            param4.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            Try
                cmd.ExecuteScalar()
                Return param4.Value
            Catch ex As Exception
            Finally
                cmd.Dispose()
                closeConnection()
            End Try
        End If
    End Function
    Public Function ItemExists() As Boolean
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Check_Item_Exists"
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@headerid"
            param1.DbType = DbType.Int32
            param1.Value = header_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@itemref"
            param2.DbType = DbType.String
            param2.Value = stock_code
            Dim param3 As New SqlParameter
            param3.ParameterName = "@mtodtid"
            param3.Direction = ParameterDirection.Output
            param3.DbType = DbType.Int32
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            Try
                cmd.ExecuteScalar()
                If param3.Value = 1 Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
            Finally
                cmd.Dispose()
                closeConnection()
            End Try
        End If
    End Function
    Public Function ReviseMto() As Boolean
        Dim intRet As Integer = 0
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Revise_Mto"
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@mtohdrid"
            param1.DbType = DbType.Int32
            param1.Value = header_id
            Dim param2 As New SqlParameter
            param2.DbType = DbType.String
            param2.ParameterName = "@revisno"
            param2.Value = revision_no
            Dim param3 As New SqlParameter
            param3.DbType = DbType.Int32
            param3.ParameterName = "@userid"
            param3.Value = user_id
            Dim param4 As New SqlParameter
            param4.DbType = DbType.Int32
            param4.ParameterName = "@sp"
            param4.Direction = ParameterDirection.Output
            Dim param5 As New SqlParameter
            param5.DbType = DbType.String
            param5.ParameterName = "@revis"
            param5.Value = revision_comment
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            Try
                cmd.ExecuteNonQuery()
                intRet = param4.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        If intRet = 1 Then Return True Else Return False
    End Function
    Public Function LoadItems() As DataTable
        Dim dt As New DataTable
        cmd = New SqlCommand
        cmd.CommandText = "Web_Get_Mto_Items"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = db_Connection
        Dim param1 As New SqlParameter
        param1.ParameterName = "@headerid"
        param1.DbType = DbType.Int32
        param1.Value = header_id
        cmd.Parameters.Add(param1)
        Try
            If getConnection() Then
                dr = cmd.ExecuteReader
                dt.Load(dr)
                dr.Close()
                cmd.Dispose()
            End If
        Catch ex As Exception
        Finally
            closeConnection()
        End Try
        Return dt
    End Function
    Public Function GetBudgetConsumption() As DataTable
        Dim dt As New DataTable
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Get_Budget_Consumption"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@jobno"
            param1.DbType = DbType.String
            param1.Value = project_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@boqid"
            param2.DbType = DbType.Int32
            param2.Value = boq_id
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Try
                dr = cmd.ExecuteReader
                dt.Load(dr)
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        Return dt
    End Function
    Public Function ProjectAccess() As Boolean
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Check_Project_Access"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@userid"
            param1.DbType = DbType.Int32
            param1.Value = user_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@projectid"
            param2.DbType = DbType.String
            param2.Value = project_id
            Dim param3 As New SqlParameter
            param3.ParameterName = "@exists"
            param3.DbType = DbType.Int32
            param3.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            Try
                cmd.ExecuteScalar()
                If param3.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Function
    Public Function GetBudgetConsumptionPercent(ByVal intUOM As Integer) As Double
        If getConnection() Then
            Dim cmd As New SqlClient.SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "Get_Category_Vs_DBudget_Consumption_Percent"
            Dim param1 As New SqlClient.SqlParameter
            Dim param2 As New SqlClient.SqlParameter
            Dim param3 As New SqlClient.SqlParameter
            Dim param4 As New SqlClient.SqlParameter
            Dim param5 As New SqlClient.SqlParameter
            Dim param6 As New SqlClient.SqlParameter
            Dim param7 As New SqlClient.SqlParameter
            Dim param8 As New SqlClient.SqlParameter

            param1.ParameterName = "@project_ref"
            param1.DbType = DbType.String
            param1.Value = project_id

            param2.ParameterName = "@area_id"
            param2.DbType = DbType.Int32
            param2.Value = area_code

            param3.ParameterName = "@zone_id"
            param3.DbType = DbType.Int32
            param3.Value = zone_code

            param4.ParameterName = "@category"
            param4.DbType = DbType.String
            param4.Value = category

            param5.ParameterName = "@boq_id"
            param5.DbType = DbType.Int32
            param5.Value = boq_id

            param6.ParameterName = "@var_ref"
            param6.DbType = DbType.String
            param6.Value = var_ref

            param7.ParameterName = "@uom"
            param7.DbType = DbType.Int32
            param7.Value = intUOM

            param8.ParameterName = "@percent"
            param8.Direction = ParameterDirection.Output
            param8.DbType = DbType.Double

            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            cmd.Parameters.Add(param8)

            Try
                cmd.ExecuteScalar()
                If Not IsDBNull(param8.Value) Then
                    If param8.Value > 100 Then param8.Value = 100
                    Return Math.Round(param8.Value, 2)
                Else
                    Return 0
                End If
            Catch ex As Exception
                'MsgBox(ex.Message)
            Finally
                cmd.Dispose()
                closeConnection()
            End Try
        End If
    End Function
    Public Function ImportItems(ByVal fields As List(Of Object)) As String
        Dim fieldValues() As Object, intI As Integer, intStatus As Integer, strResult As String = ""
        For intI = 0 To fields.Count - 1
            fieldValues = fields(intI)
            category = fieldValues(1)
            intStatus = GetBudgetStatus()
            If intStatus = 3 Or intStatus = 2 Then
                Query = "select item_no from mto_details where item_no='" & fieldValues(2) & "' and mtohdr_id=" & header_id & ""
                If String.IsNullOrEmpty(getSingleValue()) Then
                    ImportSingleItem(Val(fieldValues(0)))
                End If
            ElseIf intStatus = 1 Then
                strResult = strResult & fieldValues(2) & " (Category not yet budgeted)." & Environment.NewLine
            ElseIf intStatus = 4 Then
                strResult = strResult & fieldValues(2) & " (Design budget does not exists)." & Environment.NewLine
            End If
        Next
        Return strResult
    End Function
    Public Sub ImportSingleItem(ByVal intDtId As Integer)
        cmd = New SqlCommand()
        cmd.CommandText = "Web_Import_Mto_Item"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = db_Connection
        Dim param1 As New SqlParameter
        param1.DbType = DbType.Int32
        param1.ParameterName = "@headerid"
        param1.Value = header_id
        Dim param2 As New SqlParameter
        param2.ParameterName = "@dtid"
        param2.DbType = DbType.Int32
        Query = "select max(mtodt_id)from mto_details"
        param2.Value = GetLatestNo()
        Dim param3 As New SqlParameter
        param3.ParameterName = "@mtodtid"
        param3.DbType = DbType.Int32
        param3.Value = intDtId
        cmd.Parameters.Add(param1)
        cmd.Parameters.Add(param2)
        cmd.Parameters.Add(param3)
        If getConnection() Then
            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            Catch ex As Exception

            Finally
                closeConnection()
            End Try
        End If
    End Sub

    Public Function GetBudgetStatus() As Integer
        Dim intStatus As Integer = 0
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Get_Budget_Status"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@jobno"
            param1.DbType = DbType.String
            param1.Value = project_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@category"
            param2.DbType = DbType.String
            param2.Value = category
            Dim param3 As New SqlParameter
            param3.ParameterName = "@type"
            param3.DbType = DbType.Int32
            param3.Value = boq_type
            Dim param4 As New SqlParameter
            param4.ParameterName = "@varref"
            param4.DbType = DbType.String
            param4.Value = var_ref
            Dim param5 As New SqlParameter
            param5.ParameterName = "@areaid"
            param5.DbType = DbType.Int32
            param5.Value = area_code
            Dim param6 As New SqlParameter
            param6.ParameterName = "@zoneid"
            param6.DbType = DbType.Int32
            param6.Value = zone_code
            Dim param7 As New SqlParameter
            param7.ParameterName = "@status"
            param7.DbType = DbType.Int32
            param7.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            Try
                cmd.ExecuteScalar()
                intStatus = param7.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        Return intStatus
    End Function
    Public Function GetMtoConsumption(ByVal bolIsValue As Boolean, ByVal bolIsDesign As Boolean, ByVal intNeglectID As Integer) As Double
        Dim fltCons As Double = 0
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.Connection = db_Connection
            cmd.CommandText = "Web_Get_Mto_Consumption"
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@jobno"
            param1.DbType = DbType.String
            param1.Value = project_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@category"
            param2.DbType = DbType.String
            param2.Value = category
            Dim param3 As New SqlParameter
            param3.ParameterName = "@areaid"
            param3.DbType = DbType.Int32
            param3.Value = area_code
            Dim param4 As New SqlParameter
            param4.ParameterName = "@zoneid"
            param4.DbType = DbType.Int32
            param4.Value = zone_code
            Dim param5 As New SqlParameter
            param5.ParameterName = "@boqid"
            param5.DbType = DbType.Int32
            param5.Value = boq_id
            Dim param6 As New SqlParameter
            param6.ParameterName = "@design"
            param6.DbType = DbType.Int32
            param6.Value = bolIsDesign
            Dim param7 As New SqlParameter
            param7.ParameterName = "@value"
            param7.DbType = DbType.Int32
            param7.Value = bolIsValue
            Dim param8 As New SqlParameter
            param8.ParameterName = "@mtodtid"
            param8.DbType = DbType.Int32
            param8.Value = intNeglectID
            Dim param9 As New SqlParameter
            param9.ParameterName = "@cons"
            param9.DbType = DbType.Double
            param9.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            cmd.Parameters.Add(param8)
            cmd.Parameters.Add(param9)
            Try
                cmd.ExecuteScalar()
                fltCons = param9.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        Return fltCons
    End Function
    Public Function GetBudget(ByVal bolIsValue As Boolean) As Double
        Dim fltBudget As Double = 0
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Get_Budget"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@jobno"
            param1.DbType = DbType.String
            param1.Value = project_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@category"
            param2.DbType = DbType.String
            param2.Value = category
            Dim param3 As New SqlParameter
            param3.ParameterName = "@type"
            param3.DbType = DbType.Int32
            param3.Value = boq_type
            Dim param4 As New SqlParameter
            param4.ParameterName = "@varref"
            param4.DbType = DbType.String
            param4.Value = var_ref
            Dim param5 As New SqlParameter
            param5.ParameterName = "@areaid"
            param5.DbType = DbType.Int32
            param5.Value = area_code
            Dim param6 As New SqlParameter
            param6.ParameterName = "@zoneid"
            param6.DbType = DbType.Int32
            param6.Value = zone_code
            Dim param8 As New SqlParameter
            param8.ParameterName = "@value"
            param8.DbType = DbType.Int32
            param8.Value = bolIsValue
            Dim param7 As New SqlParameter
            param7.ParameterName = "@budget"
            param7.DbType = DbType.Double
            param7.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param8)
            cmd.Parameters.Add(param7)
            Try
                cmd.ExecuteScalar()
                fltBudget = param7.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        Return fltBudget
    End Function
    Public Function GetConsumptionBase() As Integer
        Query = "select cons_uom from cost_category_budget where jobno='" & project_id & "' and " _
                & "item_cls_id='" & category & "'and type=" & boq_type & " and var_ref='" & var_ref & "'"
        If getConnection() Then
            Return Val(getSingleValue())
        End If
    End Function
    Public Function GetStockQty() As Double
        Dim fltStkQty As Double = 0
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Get_Stock_Quantity"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@Cut_Length"
            param1.DbType = DbType.Double
            param1.Value = cut_length
            Dim param2 As New SqlParameter
            param2.ParameterName = "@Cut_Width"
            param2.DbType = DbType.Double
            param2.Value = cut_width
            Dim param3 As New SqlParameter
            param3.ParameterName = "@Quantity"
            param3.DbType = DbType.Int32
            param3.Value = qty
            Dim param4 As New SqlParameter
            param4.ParameterName = "@Cut_Size"
            param4.DbType = DbType.Int32
            param4.Value = 1
            Dim param5 As New SqlParameter
            param5.ParameterName = "@Stock_Code"
            param5.DbType = DbType.String
            param5.Value = stock_code
            Dim param6 As New SqlParameter
            param6.ParameterName = "@Stock_Qty"
            param6.Direction = ParameterDirection.Output
            param6.DbType = DbType.Double
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            Try
                cmd.ExecuteScalar()
                fltStkQty = param6.Value
            Catch ex As Exception

            End Try
            closeConnection()
        End If
        Return Math.Round(fltStkQty, 4)
    End Function
    Public Sub CreateMTO()
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Create_MTO"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@referno"
            param1.DbType = DbType.String
            param1.Value = refer_no
            Dim param2 As New SqlParameter
            param2.ParameterName = "@ext_ref"
            param2.DbType = DbType.String
            param2.Value = ext_ref
            Dim param3 As New SqlParameter
            param3.ParameterName = "@variance_ref"
            param3.DbType = DbType.String
            param3.Value = var_ref
            Dim param4 As New SqlParameter
            param4.ParameterName = "@location"
            param4.DbType = DbType.String
            param4.Value = loc
            Dim param5 As New SqlParameter
            param5.ParameterName = "@userid"
            param5.DbType = DbType.Int32
            param5.Value = user_id
            Dim param6 As New SqlParameter
            param6.ParameterName = "@remarks"
            param6.DbType = DbType.String
            param6.Value = remarks
            Dim param7 As New SqlParameter
            param7.ParameterName = "@project_id"
            param7.DbType = DbType.String
            param7.Value = project_id
            Dim param8 As New SqlParameter
            param8.ParameterName = "@doc_ref"
            param8.DbType = DbType.String
            param8.Value = doc_ref
            Dim param9 As New SqlParameter
            param9.ParameterName = "@boq_id"
            param9.DbType = DbType.Int32
            param9.Value = boq_id
            Dim param10 As New SqlParameter
            param10.ParameterName = "@area_code"
            param10.DbType = DbType.Int32
            param10.Value = area_code
            Dim param11 As New SqlParameter
            param11.ParameterName = "@zone_code"
            param11.DbType = DbType.Int32
            param11.Value = zone_code
            Dim param12 As New SqlParameter
            param12.ParameterName = "is_mor"
            param12.DbType = DbType.Boolean
            param12.Value = is_mor
            Dim param13 As New SqlParameter
            param13.ParameterName = "@morhdr_id"
            param13.DbType = DbType.Int32
            param13.Value = MORhdrid
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            cmd.Parameters.Add(param8)
            cmd.Parameters.Add(param9)
            cmd.Parameters.Add(param10)
            cmd.Parameters.Add(param11)
            cmd.Parameters.Add(param12)
            cmd.Parameters.Add(param13)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Sub
    Public Function LoadProject() As DataTable
        If getConnection() Then
            cmd = New SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandText = "Web_Get_Project_Info"
            cmd.CommandType = Data.CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@userid"
            param1.DbType = Data.DbType.Int32
            param1.Value = user_id
            cmd.Parameters.Add(param1)
            Dim dt As New DataTable
            Try
                dr = cmd.ExecuteReader()
                dt.Load(dr)
                closeConnection()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Return dt
        End If
    End Function
    Public Function LoadBoq() As DataTable
        Dim strQry As String
        strQry = "select row_id,case boq_type when 0 then 'Original' else 'Variation' end as type,boq_desc,variation_ref from boq_header " _
               & " where project_id='" & project_id & "' order by row_id "
        Query = strQry
        Return getDataTable()
    End Function
    Public Function LoadArea() As DataTable
        Dim strQry As String
        strQry = "select area_desc,area_code from al_job_area where jobno='" & project_id & "'"
        Query = strQry
        Return getDataTable()
    End Function
    Public Function LoadZone(ByVal intArea As Integer) As DataTable
        Dim strQry As String
        strQry = "select zone_desc,zone_code from al_job_zone where jobno='" & project_id & "' and " _
                    & "area_code=" & intArea & ""
        Query = strQry
        Return getDataTable()
    End Function
    Public Function LoadFinishes() As DataTable
        Dim strQry As String
        strQry = "select finishdescr,finishid from finish_hdr,project_finish where finishid=finish_id and project_ref='" & project_id & "'"
        Query = strQry
        Return getDataTable()
    End Function
    Public Function GetMtoRefNumber() As String
        If getConnection() Then
            Dim strQry As String
            strQry = "select referno from mto_hdr where mtohdr_id=" & header_id & ""
            Query = strQry
            Return getSingleValue()
        End If
    End Function
    Public Function GetMtos(ByVal intStatus As Integer) As DataTable
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Get_Mtos"
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@projectid"
            param1.DbType = DbType.String
            param1.Value = project_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@status "
            param2.DbType = DbType.Int32
            param2.Value = intStatus
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Dim dt As New DataTable
            Try
                dr = cmd.ExecuteReader
                dt.Load(dr)
            Catch ex As Exception

            End Try
            closeConnection()
            Return dt
        End If
    End Function
    Public Function GetCutSizeInfo() As SqlDataReader
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Get_CutSize"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@item_code"
            param1.DbType = DbType.String
            param1.Value = stock_code
            cmd.Parameters.Add(param1)
            Try
                dr = cmd.ExecuteReader()
            Catch ex As Exception

            End Try
            Return dr
        End If
    End Function
    Public Function CutSizeAlreadyUsed() As Integer
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Check_CutSize_Used"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@item_code"
            param1.DbType = DbType.String
            param1.Value = stock_code
            Dim param2 As New SqlParameter
            param2.ParameterName = "@return"
            param2.DbType = DbType.Int32
            param2.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Try
                cmd.ExecuteScalar()
                closeConnection()
                Return (param2.Value)
            Catch ex As Exception

            End Try
        End If
    End Function
    Public Function CutSizeAlreadyExists() As Integer
        cmd = New SqlCommand
        cmd.CommandText = "Web_Check_CutSize_Exists"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = db_Connection
        Dim param1 As New SqlParameter
        param1.ParameterName = "@item_code"
        param1.DbType = DbType.String
        param1.Value = stock_code
        Dim param2 As New SqlParameter
        param2.DbType = DbType.Int32
        param2.ParameterName = "@return"
        param2.Direction = ParameterDirection.Output
        cmd.Parameters.Add(param1)
        cmd.Parameters.Add(param2)
        If getConnection() Then
            Try
                cmd.ExecuteScalar()
                closeConnection()
                Return param2.Value
            Catch ex As Exception

            End Try
        End If
    End Function

    Public Function GetMTONumber() As String
        If getConnection() Then
            Dim strQry As String, strMTO As String = ""
            strQry = "select * from msettings where prjid='" & project_id & "'"
            cmd = New SqlCommand(strQry, db_Connection)
            Try
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    dr.Read()
                    strMTO = dr.Item("MTOINDX") & "-" & dr.Item("Prjid") & "-" & dr.Item("NextSlno")
                End If
                dr.Close()
                cmd.Dispose()
                closeConnection()
            Catch ex As Exception

            End Try
            Return strMTO
        End If
    End Function

    Public Function LoadCutSizes(ByRef grd As DevExpress.Web.ASPxGridView.ASPxGridView, ByVal strItem As String) As String
        If getConnection() Then
            Dim strQry As String
            strQry = "select item_code,stk_item_code,cut_size_code,clength,cwidth from mat_cut_sizes where stk_item_code='" & strItem & "' order by cut_size_code"
            cmd = New SqlCommand(strQry, db_Connection)
            adpt = New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            adpt.Fill(ds, "CutSize")
            grd.DataSource = ds.Tables("CutSize")
            grd.DataBind()
            If ds.Tables(0).Rows.Count = 0 Then
                Return "Item Cut Sizes not yet defined"
            Else
                Return ""
            End If
        End If
    End Function
    Public Function Get_Rate(ByVal GPRate As Double) As Double
        Dim SqlStr As String
        Dim dr As SqlClient.SqlDataReader
        SqlStr = "select " & GPRate & " * CurrencyRate as CurCost from Projects where Project_Id='" & project_id & "'"
        If getConnection() Then
            Dim cmd As New SqlClient.SqlCommand(SqlStr, db_Connection)
            Try
                dr = cmd.ExecuteReader
                If dr.Read() = True Then
                    Get_Rate = dr.Item("CurCost")
                End If
            Catch
                dr.Close()
            End Try
            closeConnection()
        End If
    End Function

End Class
#End Region

#Region "Take Off Preparation"

Public Class clsTakeOff
    Inherits clsMTO
    Friend filter As String
    Friend length As Double
    Friend width As Double
    Friend weight As Double
    Friend rate As Double
    Friend baseunit As String
    Friend cutsize As Boolean
    Friend budgetuom As Integer
    Friend description As String
    Friend finishdesc As String
    Friend finishvalue As Integer
    Friend itemarea As Double
    Friend stkqty As Double

    Public WriteOnly Property MorReferenceNo As String
        Set(value As String)
            refer_no = value
        End Set
    End Property
    Public WriteOnly Property MorDetailID As Integer
        Set(value As Integer)
            detail_id = value
        End Set
    End Property
    Public WriteOnly Property Finish As String
        Set(value As String)
            finishdesc = value
        End Set
    End Property
    Public WriteOnly Property StockDescription As String
        Set(value As String)
            description = value
        End Set
    End Property
    Public WriteOnly Property IsCutSize As Boolean
        Set(value As Boolean)
            cutsize = value
        End Set
    End Property
    Public WriteOnly Property StockBaseUnit As String
        Set(value As String)
            baseunit = value
        End Set
    End Property
    Public WriteOnly Property StockLength As Double
        Set(value As Double)
            length = value
        End Set
    End Property
    Public WriteOnly Property StockQuantity As Double
        Set(value As Double)
            stkqty = value
        End Set
    End Property
    Public WriteOnly Property CutRate As Double
        Set(value As Double)
            rate = value
        End Set
    End Property
    Public WriteOnly Property StockArea As Double
        Set(value As Double)
            itemarea = value
        End Set
    End Property
    Public WriteOnly Property StockWeight As Double
        Set(value As Double)
            weight = value
        End Set
    End Property
    Public WriteOnly Property StockWidth As Double
        Set(value As Double)
            width = value
        End Set
    End Property
    Public WriteOnly Property MorHeader As Integer
        Set(value As Integer)
            header_id = value
        End Set
    End Property
    Public WriteOnly Property FinishID As Integer
        Set(value As Integer)
            finishvalue = value
        End Set
    End Property
    Public WriteOnly Property BudgetedUOM As Integer
        Set(value As Integer)
            budgetuom = value
        End Set
    End Property
    Public WriteOnly Property MorFilter As String
        Set(value As String)
            filter = value
        End Set
    End Property
    Public Sub MoveMorItemToMto()
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "Web_Insert_Item_Into_MTO_From_MOR"
            Dim param1 As New SqlParameter
            param1.ParameterName = "@mtohdrid"
            param1.DbType = DbType.Int32
            param1.Value = MORhdrid
            Dim param2 As New SqlParameter
            param2.ParameterName = "@mordtid"
            param2.DbType = DbType.Int32
            param2.Value = detail_id
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
            Finally
                closeConnection()
            End Try
        End If
    End Sub
    Public Function GetStockArea() As Double
        Return IIf(Val(length) = 0, 1, Val(length)) * IIf(Val(width) = 0, 1, Val(width)) * IIf(Val(weight) = 0, 1, Val(weight))
    End Function
    Public Function GetInventoryItemDetails() As DataTable
        Query = "select itemdesc,length,width,stndcost,currcost,unitweight,baseuofm from vw_abal_item_mstr where itemnmbr='" & stock_code & "'"
        Return getDataTable()
    End Function
    Public Function GetMors() As DataTable
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Get_Mors"
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@projectid"
            param1.DbType = DbType.String
            param1.Value = project_id
            cmd.Parameters.Add(param1)
            Dim dt As New DataTable
            Try
                dr = cmd.ExecuteReader
                dt.Load(dr)
            Catch ex As Exception

            End Try
            closeConnection()
            Return dt
        End If
    End Function
    Public Function MORItemExists() As Boolean
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Check_MOR_Item_Exists"
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@headerid"
            param1.DbType = DbType.Int32
            param1.Value = header_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@itemref"
            param2.DbType = DbType.String
            param2.Value = cutsize_code
            Dim param3 As New SqlParameter
            param3.ParameterName = "@mordtid"
            param3.Direction = ParameterDirection.Output
            param3.DbType = DbType.Int32
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            Try
                cmd.ExecuteScalar()
                If param3.Value = 1 Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
            Finally
                cmd.Dispose()
                closeConnection()
            End Try
        End If
    End Function
    Public Function GetTakeOffCutSizeCode() As String
        Dim strCode As String
        strCode = Trim(stock_code)
        strCode = strCode + "#CS#"
        If Val(cut_width) <> 0 Then
            strCode = strCode + (Val(cut_width) * 1000).ToString + "X" + (Val(cut_length) * 1000).ToString
        Else
            strCode = strCode + (Val(cut_length) * 1000).ToString
        End If
        Return strCode
    End Function

    Public Function GetCutSizeArea() As Double
        Return IIf(Val(cut_length) = 0, 1, Val(cut_length)) * IIf(Val(cut_width) = 0, 1, _
                        Val(cut_width)) * IIf(Val(weight) = 0, 1, Val(weight))
    End Function
    Public Sub UpdateMOR()
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "Web_Update_MOR"
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@detailid"
            param1.DbType = DbType.Int32
            param1.Value = detail_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@stockqty"
            param2.DbType = DbType.Double
            param2.Value = stkqty
            Dim param3 As New SqlParameter
            param3.ParameterName = "@qty"
            param3.DbType = DbType.Double
            param3.Value = qty
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
            Finally
                closeConnection()
            End Try
        End If
    End Sub
    Public Sub UpdateMorFromMto()
        Query = "update mto_hdr set morhdr_id=" & MORhdrid & ",is_mor=1 where mtohdr_id=" & header_id & ""
        ExecuteQuery()
    End Sub

    Public Function CreateMtoFromMor() As Integer
        If getConnection() Then
            cmd = New SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "Web_Create_MTO_From_MOR"
            Dim param1 As New SqlParameter
            param1.ParameterName = "@morhdrid"
            param1.DbType = DbType.Int32
            param1.Value = MORhdrid
            Dim param2 As New SqlParameter
            param2.ParameterName = "@project_id"
            param2.DbType = DbType.String
            param2.Value = project_id
            Dim param3 As New SqlParameter
            param3.ParameterName = "@referno"
            param3.DbType = DbType.String
            param3.Value = refer_no
            Dim param4 As New SqlParameter
            param4.ParameterName = "@userid"
            param4.DbType = DbType.Int32
            param4.Value = user_id
            Dim param5 As New SqlParameter
            param5.DbType = DbType.Int32
            param5.ParameterName = "@returnid"
            param5.Direction = ParameterDirection.Output
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            Try
                cmd.ExecuteNonQuery()
                Return param5.Value
            Catch ex As Exception
            Finally
                closeConnection()
            End Try
        End If
    End Function
    Public Function GetOptimization() As DataTable
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Get_TakeOff_Optimization"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@Category"
            param1.DbType = DbType.String
            param1.Value = category
            Dim param2 As New SqlParameter
            param2.ParameterName = "@CutLength"
            param2.DbType = DbType.Double
            param2.Value = cut_length
            Dim param3 As New SqlParameter
            param3.ParameterName = "@CutWidth"
            param3.DbType = DbType.Double
            param3.Value = cut_width
            Dim param4 As New SqlParameter
            param4.ParameterName = "@Qty"
            param4.DbType = DbType.Int32
            param4.Value = qty
            Dim param5 As New SqlParameter
            param5.ParameterName = "@Inv_Item"
            param5.DbType = DbType.String
            param5.Value = item_code
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            Dim dt As New DataTable
            Try
                dr = cmd.ExecuteReader
                dt.Load(dr)
            Catch ex As Exception
            Finally
                dr.Close()
                closeConnection()
            End Try
            Return dt
        End If
    End Function
    Public Function GetMtoWithoutMor() As DataTable
        If getConnection() Then
            cmd = New SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandText = "Web_Get_Mtos_Without_Mors"
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.DbType = DbType.Int32
            param1.ParameterName = "@morhdrid"
            param1.Value = header_id
            cmd.Parameters.Add(param1)
            Dim dt As New DataTable
            Try
                dr = cmd.ExecuteReader
                dt.Load(dr)
            Catch ex As Exception
            Finally
                closeConnection()
            End Try
            Return dt
        End If
    End Function


    Public Function GetBudgetedCategory() As DataTable
        If getConnection() Then
            cmd = New SqlCommand
            cmd.CommandText = "Web_Get_Budgeted_Category"
            cmd.Connection = db_Connection
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@jobno"
            param1.DbType = DbType.String
            param1.Value = project_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@type"
            param2.DbType = DbType.Int32
            param2.Value = boq_type
            Dim param3 As New SqlParameter
            param3.ParameterName = "@varref"
            param3.DbType = DbType.String
            param3.Value = var_ref
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            Dim dt As New DataTable
            Try
                dr = cmd.ExecuteReader
                dt.Load(dr)
            Catch ex As Exception
            Finally
                closeConnection()
            End Try
            Return dt
        End If
    End Function
    Public Sub CreateCutSize()
        If getConnection() Then
            cmd = New SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandText = "Web_Create_Take_Off_CutSizes"
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@cutsizecode"
            param1.DbType = DbType.String
            param1.Value = cutsize_code
            Dim param2 As New SqlParameter
            param2.ParameterName = "@itemcode"
            param2.DbType = DbType.String
            param2.Value = item_code
            Dim param3 As New SqlParameter
            param3.ParameterName = "@stockcode"
            param3.DbType = DbType.String
            param3.Value = stock_code
            Dim param4 As New SqlParameter
            param4.ParameterName = "@cutlength"
            param4.DbType = DbType.Double
            param4.Value = cut_length
            Dim param5 As New SqlParameter
            param5.ParameterName = "@cutwidth"
            param5.DbType = DbType.Double
            param5.Value = cut_width
            Dim param6 As New SqlParameter
            param6.ParameterName = "@uid"
            param6.DbType = DbType.Int32
            param6.Value = user_id
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Sub
    Public Sub InsertMORDetails()
        If getConnection() Then
            cmd = New SqlCommand
            cmd.Connection = db_Connection
            cmd.CommandText = "Web_Insert_MOR_Items"
            cmd.CommandType = CommandType.StoredProcedure
            Dim param1 As New SqlParameter
            param1.ParameterName = "@morhdrid"
            param1.DbType = DbType.Int32
            param1.Value = header_id
            Dim param2 As New SqlParameter
            param2.ParameterName = "@category"
            param2.DbType = DbType.String
            param2.Value = category
            Dim param3 As New SqlParameter
            param3.ParameterName = "@cutsizecode"
            param3.DbType = DbType.String
            param3.Value = cutsize_code
            Dim param4 As New SqlParameter
            param4.ParameterName = "@cutlength"
            param4.DbType = DbType.Double
            param4.Value = cut_length
            Dim param5 As New SqlParameter
            param5.ParameterName = "@cutwidth"
            param5.DbType = DbType.Double
            param5.Value = cut_width
            Dim param6 As New SqlParameter
            param6.ParameterName = "@qty"
            param6.DbType = DbType.Int32
            param6.Value = qty
            Dim param7 As New SqlParameter
            param7.ParameterName = "@stockcode"
            param7.DbType = DbType.String
            param7.Value = stock_code
            Dim param8 As New SqlParameter
            param8.ParameterName = "@morfilter"
            param8.DbType = DbType.String
            param8.Value = filter
            Dim param9 As New SqlParameter
            param9.ParameterName = "@length"
            param9.DbType = DbType.Double
            param9.Value = length
            Dim param10 As New SqlParameter
            param10.ParameterName = "@width"
            param10.DbType = DbType.Double
            param10.Value = width
            Dim param11 As New SqlParameter
            param11.ParameterName = "@weight"
            param11.DbType = DbType.Double
            param11.Value = weight
            Dim param12 As New SqlParameter
            param12.ParameterName = "@cost"
            param12.DbType = DbType.Double
            param12.Value = rate
            Dim param13 As New SqlParameter
            param13.ParameterName = "@baseunit"
            param13.DbType = DbType.String
            param13.Value = baseunit
            Dim param14 As New SqlParameter
            param14.ParameterName = "@budgetuom"
            param14.DbType = DbType.Int32
            param14.Value = budgetuom
            Dim param15 As New SqlParameter
            param15.ParameterName = "@description"
            param15.DbType = DbType.String
            param15.Value = description
            Dim param16 As New SqlParameter
            param16.ParameterName = "@area"
            param16.DbType = DbType.Double
            param16.Value = itemarea
            Dim param17 As New SqlParameter
            param17.ParameterName = "@finishid"
            param17.DbType = DbType.Int32
            param17.Value = finishvalue
            Dim param18 As New SqlParameter
            param18.ParameterName = "@finish"
            param18.DbType = DbType.String
            param18.Value = finishdesc
            Dim param19 As New SqlParameter
            param19.ParameterName = "@stockqty"
            param19.DbType = DbType.Double
            param19.Value = stkqty
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            cmd.Parameters.Add(param8)
            cmd.Parameters.Add(param9)
            cmd.Parameters.Add(param10)
            cmd.Parameters.Add(param11)
            cmd.Parameters.Add(param12)
            cmd.Parameters.Add(param13)
            cmd.Parameters.Add(param14)
            cmd.Parameters.Add(param15)
            cmd.Parameters.Add(param16)
            cmd.Parameters.Add(param17)
            cmd.Parameters.Add(param18)
            cmd.Parameters.Add(param19)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
            Finally
                cmd.Dispose()
                closeConnection()
            End Try
        End If
    End Sub
    Public Sub CreateMOR()
        If getConnection() Then
            cmd = New SqlCommand()
            cmd.CommandText = "Web_Create_MOR"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = db_Connection
            Dim param1 As New SqlParameter
            param1.ParameterName = "@referno"
            param1.DbType = DbType.String
            param1.Value = refer_no
            Dim param2 As New SqlParameter
            param2.ParameterName = "@ext_ref"
            param2.DbType = DbType.String
            param2.Value = ext_ref
            Dim param3 As New SqlParameter
            param3.ParameterName = "@variance_ref"
            param3.DbType = DbType.String
            param3.Value = var_ref
            Dim param4 As New SqlParameter
            param4.ParameterName = "@location"
            param4.DbType = DbType.String
            param4.Value = loc
            Dim param5 As New SqlParameter
            param5.ParameterName = "@userid"
            param5.DbType = DbType.Int32
            param5.Value = user_id
            Dim param6 As New SqlParameter
            param6.ParameterName = "@remarks"
            param6.DbType = DbType.String
            param6.Value = remarks
            Dim param7 As New SqlParameter
            param7.ParameterName = "@project_id"
            param7.DbType = DbType.String
            param7.Value = project_id
            Dim param8 As New SqlParameter
            param8.ParameterName = "@doc_ref"
            param8.DbType = DbType.String
            param8.Value = doc_ref
            Dim param9 As New SqlParameter
            param9.ParameterName = "@boq_id"
            param9.DbType = DbType.Int32
            param9.Value = boq_id
            Dim param10 As New SqlParameter
            param10.ParameterName = "@area_code"
            param10.DbType = DbType.Int32
            param10.Value = area_code
            Dim param11 As New SqlParameter
            param11.ParameterName = "@zone_code"
            param11.DbType = DbType.Int32
            param11.Value = zone_code
            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)
            cmd.Parameters.Add(param3)
            cmd.Parameters.Add(param4)
            cmd.Parameters.Add(param5)
            cmd.Parameters.Add(param6)
            cmd.Parameters.Add(param7)
            cmd.Parameters.Add(param8)
            cmd.Parameters.Add(param9)
            cmd.Parameters.Add(param10)
            cmd.Parameters.Add(param11)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception

            End Try
            closeConnection()
        End If
    End Sub
    Public Function GetMORNumber() As String
        If getConnection() Then
            Dim strQry As String, strMTO As String = ""
            strQry = "select * from mor_settings where prjid='" & project_id & "'"
            cmd = New SqlCommand(strQry, db_Connection)
            Try
                dr = cmd.ExecuteReader
                If dr.HasRows Then
                    dr.Read()
                    strMTO = dr.Item("MORINDX") & "-" & dr.Item("Prjid") & "-" & dr.Item("NextSlno")
                End If
                dr.Close()
                cmd.Dispose()
                closeConnection()
            Catch ex As Exception

            End Try
            Return strMTO
        End If
    End Function
End Class

#End Region


