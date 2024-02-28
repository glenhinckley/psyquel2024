VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CCertAuditDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CCertAuditDB                                           '
'Date: 05/03/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS object designed to host methods associated with  '
'              data affecting the tblCertAudit table.               '
'--------------------------------------------------------------------


Public Function Insert(ByVal lngCertID As Long, _
                       ByVal strUserName As String, _
                       ByVal strNote As String) As Long
'--------------------------------------------------------------------
'Date: 05/03/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Inserts a single record into the tblCertAudit table  '
'Parameters: lngCertID - ID of cert to link CPT Code with           '
'            strUserName - Name of user executing insert call       '
'Returns: ID of record just inserted(on success)                    '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim lngErr As Long
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
        
    Call cnn.Open(CONST_PSYQUEL_CNN)
        
    'Create the parameter objects
    With cmd
        .CommandText = "uspInsCertAudit"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@CertID", adInteger, adParamInput, , lngCertID)
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@Note", adVarChar, adParamInput, 1000, Left(strNote, 1000))
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
        
        'Assign the connection to the Command object and execute the stored procedure
        Set .ActiveConnection = cnn
        .Execute ', , adExecuteNoRecords
    
        Insert = .Parameters("@ID").Value
        
        'Check the ErrorNum parameter before deciding to commit the transaction
        If .Parameters("@ErrorNum").Value <> 0 Then
            lngErr = .Parameters("@ErrorNum").Value
            GoTo ErrTrap
        End If
        
    End With
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    Call RaiseError(Err, lngErr, CLASS_NAME)
End Function

Public Function FetchCertAuditByCert(ByVal lngCertID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/03/2001                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Retrieves recordset of audits for the given cert     '
'Parameters: lngCertID - ID of Cert to fetch certs for              '
'Returns: Recordset of audit trail                                  '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    
    Call cnn.Open(CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelCertAuditByCert"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@CertID", adInteger, adParamInput, , lngCertID)
    End With
    
    '  Instantiate recordset and populate
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly

    Set rst.ActiveConnection = Nothing
    cnn.Close
    Set cnn = Nothing
    Set cmd = Nothing
    
    Set FetchCertAuditByCert = rst
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cmd = Nothing
    Set cnn = Nothing
    Set rst = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

