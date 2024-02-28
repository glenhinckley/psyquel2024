VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "CCertProviderDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CCertProviderDB                                        '
'Date: 03/07/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  MTS object designed to host methods associated with  '
'              data affecting the tblCertProvider table.            '
'--------------------------------------------------------------------
'Revision History:
'   R001: 10/26/2000 Richkun - Added Update() method
'--------------------------------------------------------------------

Option Explicit

Private Const CLASS_NAME As String = "CCertProviderDB"
Private Const TABLE_NAME As String = "tblCertProvider"


Public Function Insert(ByVal lngCertID As Long, _
                       ByVal lngProviderID As Long) As Long
'--------------------------------------------------------------------
'Date: 03/07/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Inserts a single record into the tblCertProvider     '
'               table                                               '
'Parameters: lngCertID - ID of cert to link provider with           '
'            lngProviderID - ID of provider to link cert with       '
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
        .CommandText = "uspInsCertProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@CertID", adInteger, adParamInput, , lngCertID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
        
    'Assign the connection to the Command object and execute the stored procedure
    Set cmd.ActiveConnection = cnn
    cmd.Execute ', , adExecuteNoRecords

    'Check the ErrorNum parameter before deciding to commit the transaction
    If IsEmpty(cmd.Parameters("@ErrorNum").Value) Then
        lngErrNum = vbObjectError
        strErrMsg = "Procedure uspInsCertProvider has failed."
        GoTo ErrTrap:
    End If
    
    If cmd.Parameters("@ErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    
    Insert = cmd.Parameters("@ID").Value
    
    Set cmd = Nothing
    Set cnn = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cmd = Nothing
    Set cnn = Nothing
    If lngErrNum <> 0 Then
        Err.Raise lngErrNum, CLASS_NAME, strErrMsg
    Else
        Err.Raise cmd.Parameters("@ErrorNum").Value, CLASS_NAME, Err.Description
    End If
End Function


Public Sub Update(ByVal lngCertProviderID As Long, ByVal lngProviderID As Long)
'--------------------------------------------------------------------
'Date: 10/26/2000                                                   '
'Author: Dave Richkun
'Description:  Updates a single record into the tblCertProvider table                                               '
'Parameters: lngCertProviderID - ID of the row in the table to update
'            lngProviderID - ID of provider associated with cert.
'Returns: Null
'--------------------------------------------------------------------
'Revision History:                                                  '
'  R001 - Created
'--------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
        
    'Create the parameter objects
    With cmd
        .CommandText = "uspUpdCertProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@CertProviderID", adInteger, adParamInput, , lngCertProviderID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
        
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
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
End Sub
Public Sub Delete(ByVal lngCertID As Long, ByVal lngProviderID As Long)
'--------------------------------------------------------------------
'Date: 10/30/2001
'Author: Eric Pena
'Description:  Deletes a single record into the tblCertProvider table                                               '
'Parameters: lngCertProviderID - ID of the row in the table to update
'            lngProviderID - ID of provider associated with cert.
'Returns: Null
'--------------------------------------------------------------------
'Revision History:                                                  '
'  R001 - Created
'--------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection

    On Error GoTo ErrTrap
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command
        
    'Create the parameter objects
    With cmd
        .CommandText = "uspDelCertProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@CertID", adInteger, adParamInput, , lngCertID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
        
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
    Call RaiseError(Err, cmd.Parameters("@ErrorNum").Value, CLASS_NAME)
End Sub



