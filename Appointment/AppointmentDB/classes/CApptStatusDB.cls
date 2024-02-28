VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApptStatusDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatApptDB                                             '
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS object designed to host methods associated with  '
'              data affecting the tblApptStatus table.              '
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CApptStatusDB"
Private Const TABLE_NAME As String = "tblApptStatus"

'--------------------------------------------------------------------
' Public Methods    +++++++++++++++++++++++++++++++++++++++++++++++++
'--------------------------------------------------------------------

Public Function Insert(ByVal strDescription As String) As Long
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Inserts an appointment status record into tblApptStatus
'Parameters:  strDescription - Description of the new appointment   '
'               status to be inserted                               '
'Returns:  ApptStatus ID on success, -1 otherwise                   '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspInsApptStatus"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 20, strDescription)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        
        Insert = .Parameters("@ID").Value
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            GoTo ErrTrap
        End If
        
    End With
    
    ' Clean house
    Set cnn = Nothing
    Set cmd = Nothing
    
    ' Signal completion
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
    
End Function

Public Function Update(ByVal lngID As Long, ByVal strDescription As String)
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Updates an appointment status record in tblApptStatus'
'Parameters:  lngID - ID of appt status to update                   '
'             strDescription - Updated desscription of the appt status
'Returns:  Null                                                     '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
            
    'Create the parameter objects
    With cmd
        .CommandText = "uspUpdApptStatus"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@ApptStatusID", adInteger, adParamInput, , lngID)
        .Parameters.Append .CreateParameter("@Description", adVarChar, adParamInput, 10, strDescription)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Call cnn.Open(CONST_PSYQUEL_CNN)
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
        
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            GoTo ErrTrap
        End If
        
    End With
    
    ' Clean house
    Set cnn = Nothing
    Set cmd = Nothing
    
    ' Signal completion
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
    
End Function

Public Function Fetch(Optional ByVal blnIncludeDisabled As Boolean = False) _
    As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Retrieves records from the tblApptStatus table.      '
'Parameters: blnIncludeDisabled - Optional parameter that identifies'
'               if records flagged as 'Disabled' or 'De-activated'  '
'               are to be included in the record set. The default   '
'               value is False.                                     '
'Returns: Recordset of requested Appt statuses                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    Dim strShowDisabled As String
    Dim SQLErrorNum As Long
    
    On Error GoTo ErrTrap:

    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    If blnIncludeDisabled Then
        strShowDisabled = "Y"
    Else
        strShowDisabled = "N"
    End If
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelApptStatus"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ShowDisabled", adVarChar, adParamInput, 1, strShowDisabled)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set Fetch = rst
    
    'Disconnect the recordset and clean house
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


Public Function Deleted(ByVal blnDeleted As Boolean, ByVal lngID As Long)
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Flags a row in the tblApptStatus table marking the   '
'               row as deleted or undeleted.                        '
'Parameters:   blnDeleted - Boolean value identifying if the record '
'                   is to be deleted (True) or undeleted (False).   '
'            lngID - ID of the row in the table whose value will be '
'               updated.                                            '
'Returns:  Null                                                     '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String
    Dim blnFlag As String

    On Error GoTo ErrTrap:

    If blnDeleted = True Then
        blnFlag = "Y"
    Else
        blnFlag = "N"
    End If

    'Prepare the SQL statement
    strSQL = "UPDATE "
    strSQL = strSQL & TABLE_NAME
    strSQL = strSQL & " SET "
    strSQL = strSQL & " fldDisabledYN = '" & blnFlag & "' "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "fldApptStatusID = " & lngID

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = strSQL
    cmdSQL.CommandType = adCmdText
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute , , adExecuteNoRecords

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function Exists(ByVal strDescription As String) As Boolean
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Determines if an Appointment Status description already
'               exists in tblApptStatus                             '
'Parameters:  strDescription - Appointment Status description being '
'               sought in tblApptStatus                             '
'Returns:  TRUE if association is found, FALSE otherwise            '
'--------------------------------------------------------------------
'Revision History:                                                  '
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
    strSQL = strSQL & " fldDescription = '" & strDescription & "' "
    
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

'--------------------------------------------------------------------
' PrivateMethods    +++++++++++++++++++++++++++++++++++++++++++++++++
'--------------------------------------------------------------------
