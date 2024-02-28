VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanFeeAllowableDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPlanFeeAllowableDB
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Plan Fee Allowable table.
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CPlanFeeAllowableDB"
Private Const TABLE_NAME As String = "tblPlanFeeAllowable"

Public Function Insert(ByVal lngPlanID As Long, _
                       ByVal strCPTCode As String, ByVal lngProviderCredential As Long, _
                       ByVal curFeeAllowable As Currency)
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblPlanCopay table utilizing
'              a stored procedure.
'Parameters: lngPlanID -   Plan ID number to insert into table
'            strCPTCode - The ten digit CPT Code that will be inserted into the table.
'            lngProviderCredential - Numerical representation of providers credentials.
'            curFeeAllowable - Amount paid to provider
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspInsPlanFeeAllowable"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPlanID)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@ProviderCredential", adInteger, adParamInput, , lngProviderCredential)
        .Parameters.Append .CreateParameter("@FeeAllowable", adCurrency, adParamInput, , curFeeAllowable)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    
    Insert = cmdSQL.Parameters(0).Value

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    Call RaiseError(Err, cmdSQL.Parameters("@SQLErrorNum").Value, CLASS_NAME, "")

End Function


Public Function Update(ByVal lngPlanFeeID As Long, ByVal curFeeAllowable As Currency)
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Updates a row into the tblPlanFeeAllowable table utilizing
'              a stored procedure.
'Parameters: lngPlanFeeID -   Plan Fee ID number to insert into table
'            curFeeAllowable - Amount paid to provider
'
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspUpdPlanFeeAllowable"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , lngPlanFeeID)
        .Parameters.Append .CreateParameter("@FeeAllowable", adCurrency, adParamInput, , curFeeAllowable)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    Call RaiseError(Err, cmdSQL.Parameters("@SQLErrorNum").Value, CLASS_NAME, "")

End Function


Public Function Exists(ByVal lngPlanID As Long, _
                       ByVal strCPTCode As String, ByVal lngProviderCredential As Long)
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Checks for exisiting records in tblPlanCopay
'Parameters: lngPlanFeeID -   Plan ID number to check for existing entries in table
'
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    'Prepare the SQL statement
    
    strSQL = "SELECT "
    strSQL = strSQL & "COUNT(*) AS TOTAL "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldPlanID = " & lngPlanID
    strSQL = strSQL & " AND "
    strSQL = strSQL & " fldCPTCode = '" & strCPTCode & "' "
    strSQL = strSQL & " AND "
    strSQL = strSQL & " fldProviderCredential = " & lngProviderCredential
    'Instantiate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    ' Populate the recordset
    Call rstSQL.Open(strSQL, cnnSQL, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch)
    
    'Disconnect the recordset, close the connection and return the recordset
    'to the calling environment.
    Set rstSQL.ActiveConnection = Nothing

    'Close the connection and free all resources
    cnnSQL.Close
    Set cnnSQL = Nothing

    'Get the value from the Recordset then destroy it.
    If rstSQL.Fields("TOTAL").Value > 0 Then
        Exists = True
    Else
        Exists = False
    End If
    Set rstSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function FetchByPlanCPTCredential(ByVal lngPlanID As Long, ByVal strCPTCode As String, _
                                 ByVal lngProviderCredential As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPlanFeeAllowable table.'
'Parameters: lngPlanID -   Plan ID number to fetch from table       '
'            strCPTCode - 10 digit CPT Code                         '
'            lngProviderCredential - numerical representation of Providers credentials                                                       '
'Returns: Recordset of requested providers                          '
'-------------------------------------------------------------------'

    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelFeeByPlanCPTCode"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngPlanID)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@ProviderCredential", adInteger, adParamInput, , lngProviderCredential)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByPlanCPTCredential = rst
    
    'Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function FetchByPlan(ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPlanFeeAllowable table.'
'Parameters: lngPlanID -   Plan ID number to fetch from table       '
'                                                                 '
'Returns: Recordset of requested providers                          '
'-------------------------------------------------------------------'

    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelFeeByPlan"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngPlanID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByPlan = rst
    
    'Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
