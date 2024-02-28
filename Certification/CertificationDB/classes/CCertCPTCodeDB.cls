VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CCertCPTCodeDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CCertCPTCodeDB                                         '
'Date: 03/08/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS object designed to host methods associated with  '
'              data affecting the tblCertCPTCode table.             '
'--------------------------------------------------------------------

Option Explicit

Private Const CLASS_NAME As String = "CCertCPTCodeDB"
Private Const TABLE_NAME As String = "tblCertCPTCode"

Public Function Insert(ByVal lngCertID As Long, ByVal strCPTCode As String, _
                       ByVal lngCertUnitID As Long, ByVal strUserName As String) As Long
'--------------------------------------------------------------------
'Date: 03/08/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Inserts a single record into the tblCertCPTCode      '
'               table                                               '
'Parameters: lngCertID - ID of cert to link CPT Code with           '
'            strCPTCode - CPT Code to associate with cert           '
'            lngCertUnitID - Cert Unit to associate with pair       '
'            strUserName - Name of user executing insert call
'Returns: ID of record just inserted(on success)                    '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    Dim lngErrNum As Long
    Dim strErrMsg As String

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
        
    Call cnn.Open(CONST_PSYQUEL_CNN)
        
    'Create the parameter objects
    With cmd
        .CommandText = "uspInsCertCPTCode"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@CertID", adInteger, adParamInput, , lngCertID)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@CertUnitID", adInteger, adParamInput, , lngCertUnitID)
        .Parameters.Append .CreateParameter("@UserName", adVarChar, adParamInput, 30, strUserName)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
        
    'Assign the connection to the Command object and execute the stored procedure
    Set cmd.ActiveConnection = cnn
    cmd.Execute ', , adExecuteNoRecords

    'Check the ErrorNum parameter before deciding to commit the transaction
    If IsEmpty(cmd.Parameters("@ErrorNum").Value) Then
        lngErrNum = vbObjectError
        strErrMsg = "Procedure uspInsCertCPTCode has failed."
        GoTo ErrTrap:
    End If
    
    If cmd.Parameters("@ErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    
    Insert = cmd.Parameters("@ID").Value
    
    Set cnn = Nothing
    Set cmd = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cnn = Nothing
    Set cmd = Nothing
    If lngErrNum <> 0 Then
        Err.Raise lngErrNum, CLASS_NAME, strErrMsg
    Else
        Err.Raise cmd.Parameters("@ErrorNum").Value, CLASS_NAME, Err.Description
    End If
End Function


Public Sub DeleteByUnitCode(ByVal lngCertUnitID As Long, _
                            ByVal strCPTCode As String)
'--------------------------------------------------------------------
'Date: 10/30/2000                                                   '
'Author: Dave Richkun
'Description:  Deletes rows from tblCertCPTCode and tblCertUnit where
'              lngCertUnitID is referenced
'Parameters: lngCertUnitID - ID of CertUnitID to delete
'Returns: Null
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    
    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
        
    'Create the parameter objects
    With cmd
        .CommandText = "uspDelCertCPTByUnitCode"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@CertUnitID", adInteger, adParamInput, , lngCertUnitID)
        .Parameters.Append .CreateParameter("@CPTCode", adVarChar, adParamInput, 10, strCPTCode)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
        
    'Open the database connection.
    Call cnn.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmd.ActiveConnection = cnn
    cmd.Execute ', , adExecuteNoRecords

    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmd.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    
    Set cmd = Nothing
    Set cnn = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cmd = Nothing
    Set cnn = Nothing
    Call RaiseError(Err, cmd.Parameters("@SQLErrorNum").Value, CLASS_NAME)
End Sub

