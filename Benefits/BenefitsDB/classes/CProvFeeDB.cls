VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProvFeeDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CProvFeeDB
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Provider Fee table.
'--------------------------------------------------------------------
'Revisions
'    R000 - Pena: 12/7/2000 - Modified insert and update method calls to include auditing
'    R001 - Richkun: 05/08/2001 - Added FetchByID() method
'    R002 - Richkun: 05/08/2001 - Altered Update() to include CPTCode, ProviderID parameter
'    R003 - Richkun: 05/08/2001 - Added Delete() method
'    R004 - Richkun: 05/08/2001 - Altered Exists() method to search for duplicates based on
'                   conditional parameter.  Business rules are slightly different when checking
'                   for duplicates during Insert than when checking for duplicates during Update
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CProvFeeDB"
Private Const TABLE_NAME As String = "tblProviderFee"

Public Function Insert(ByVal lngProviderID As Long, _
                       ByVal strCPTCode As String, ByVal strCPTModifier As String, _
                       ByVal strCPTDescription As String, ByVal curFee As Currency, _
                       ByVal curContractAmount As Currency, ByVal blnPatSched As Boolean, _
                       ByVal lngDuration As Long, ByVal lngPlanID As Long, _
                       ByVal strNotes As String, ByVal strAddedBy As String)
'--------------------------------------------------------------------
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderID - Provider ID number to insert into table
'            strCPTCode - The ten digit CPT Code that will be inserted into the table.
'            dblFee - The monetary amount that will be inserted into the table.
'
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspInsProvFee"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@CPTModifier", adVarChar, adParamInput, 2, strCPTModifier) 'R002
        .Parameters.Append .CreateParameter("@CPTDescription", adVarChar, adParamInput, 60, strCPTDescription) 'R002
        .Parameters.Append .CreateParameter("@Fee", adCurrency, adParamInput, , curFee)
        .Parameters.Append .CreateParameter("@ContractAmount", adCurrency, adParamInput, , IIf(IsNull(curContractAmount), 0, curContractAmount))
        .Parameters.Append .CreateParameter("@PatSched", adVarChar, adParamInput, 1, IIf(blnPatSched = True, "Y", "N"))
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , lngDuration)
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , IIf(IsNull(lngPlanID), 0, lngPlanID))
        .Parameters.Append .CreateParameter("@Notes", adVarChar, adParamInput, 255, strNotes)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strAddedBy)
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

Public Function Update(ByVal lngProviderFeeID As Long, ByVal lngProviderID As Long, _
                       ByVal strCPTCode As String, ByVal strCPTModifier As String, _
                       ByVal strCPTDescription As String, ByVal curFee As Currency, _
                       ByVal curContractAmount As Currency, ByVal blnPatSched As Boolean, _
                       ByVal lngDuration As Long, ByVal lngPlanID As Long, _
                       ByVal strNotes As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderFeeID - Provider Fee ID number to update table
'            lngProviderID - ID of Provider whose fee is being updated
'            strCPTCode - CPT Code
'            strCPTModifier - CPT Modifier
'            strCPTDescription - CPT Description
'            curFee - The monetary amount that will be updated
'            strUserName - User name identifying the user making the change
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspUpdProvFee"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@ProviderFeeID", adInteger, adParamInput, , lngProviderFeeID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID) 'R002
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode) 'R002
        .Parameters.Append .CreateParameter("@CPTModifier", adVarChar, adParamInput, 2, strCPTModifier) 'R002
        .Parameters.Append .CreateParameter("@CPTDescription", adVarChar, adParamInput, 60, strCPTDescription) 'R002
        .Parameters.Append .CreateParameter("@Fee", adCurrency, adParamInput, , curFee)
        .Parameters.Append .CreateParameter("@ContractAmount", adCurrency, adParamInput, , IIf(IsNull(curContractAmount), 0, curContractAmount))
        .Parameters.Append .CreateParameter("@PatSched", adVarChar, adParamInput, 1, IIf(blnPatSched = True, "Y", "N"))
        .Parameters.Append .CreateParameter("@Duration", adInteger, adParamInput, , IIf(IsNull(lngDuration), 0, lngDuration))
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , IIf(IsNull(lngPlanID), 0, lngPlanID))
        .Parameters.Append .CreateParameter("@Notes", adVarChar, adParamInput, 255, strNotes)
        .Parameters.Append .CreateParameter("@ModifiedBy", adVarChar, adParamInput, 30, strUserName)
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

Public Function Exists(ByVal lngProviderID As Long, ByVal strCPTCode As String, _
                       ByVal lngPlanID As Long, Optional ByVal lngProviderFeeID As Long)
'--------------------------------------------------------------------
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  Checks for duplicate entries in the tblProviderFee table
'Parameters: lngProviderID - ID of Provider
'            strCPTCode - CPT Code
'            strAction - Optional identifier to determine if checking for
'               existance during Insert or Update.  Update requires that
'               CPTCode/ProviderID combination utilizes the same Primary Key ID
'               in order to be successful.
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim rstSQL As ADODB.Recordset
    Dim strSQL As String
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    'Prepare the SQL statement
    If lngProviderFeeID = 0 Then 'R004
        strSQL = "SELECT "
        strSQL = strSQL & "COUNT(*) AS TOTAL "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & TABLE_NAME
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " fldProviderID = " & lngProviderID
        strSQL = strSQL & " AND"
        strSQL = strSQL & " fldPlanID = " & lngPlanID
        strSQL = strSQL & " AND"
        strSQL = strSQL & " fldCPTCode = '" & strCPTCode & "'"
    Else  'R004
        strSQL = "SELECT "
        strSQL = strSQL & "fldProviderFeeID "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & TABLE_NAME
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " fldProviderID = " & IIf(IsNull(lngProviderID), 0, lngProviderID)
        strSQL = strSQL & " AND"
        strSQL = strSQL & " fldPlanID = " & IIf(IsNull(lngPlanID), 0, lngPlanID)
        strSQL = strSQL & " AND"
        strSQL = strSQL & " fldCPTCode = '" & strCPTCode & "'"
    End If
        
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

    If lngProviderFeeID = 0 Then
        'This condition called when checking for duplicate during Insert 'R004
        If rstSQL.Fields("TOTAL").Value > 0 Then
            Exists = True
        Else
            Exists = False
        End If
    Else
        'This condition called when checking for duplicate during Update 'R004
        If Not (rstSQL Is Nothing) Then
            Do While Not (rstSQL.EOF)
                If rstSQL.Fields("fldProviderFeeID").Value <> lngProviderFeeID Then
                    Exists = True
                    Exit Do
                Else
                    Exists = False
                End If
                rstSQL.MoveNext
            Loop
        Else
            Exists = False
        End If
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

Public Function FetchByProvider(ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/25/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderFee table.     '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'
'Returns: Recordset of requested providers
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
        .CommandText = "uspSelFeeByProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByProvider = rst
    
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

Public Function FetchByProviderCPTCode(ByVal lngProviderID As Long, _
                                       ByVal strCPTCode As String, _
                                       ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/25/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderFee table.     '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'            strCPTCode - CPT Code to fetch from table
'Returns: RecordSet with requested fee
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
        .CommandText = "uspSelFeeByProviderCPTCode"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , lngPlanID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByProviderCPTCode = rst
    
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


Public Function FetchByID(ByVal lngProviderFeeID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/08/2001
'Author: Dave Richkun
'Description:  Retrieves a record from the tblProviderFee table.
'Parameters: lngProviderFeeID - Value of the Primary Key identifiying the
'               row that will be retrieved.
'Returns: ADO Recordset
'-------------------------------------------------------------------'
' R001 - Created
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
        .CommandText = "uspSelFeeByID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ProviderFeeID", adInteger, adParamInput, , lngProviderFeeID)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByID = rst
    
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

Public Function FetchByPlan(ByVal lngProviderID As Long, _
                            ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 03/01/2005                                                   '
'Author: Duane C Orth                                              '
'Description:  Retrieves records from the tblProviderFee table.'
'Parameters: lngPlanID -   Plan ID number to fetch from table       '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'                                                                   '
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
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , lngPlanID)
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


Public Function Delete(ByVal lngProviderFeeID As Long) 'R003
'--------------------------------------------------------------------
'Date: 05/08/2001
'Author: Dave Richkun
'Description:  Deletes a row from the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderFeeID - ID of row to delete
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspDelProvFee"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@ProviderFeeID", adInteger, adParamInput, , lngProviderFeeID)
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

