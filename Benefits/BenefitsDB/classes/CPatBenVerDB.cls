VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatBEnVerDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatBenVerDB
'Date: 10/26/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Patient Benefits Verification table.
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CPatBenVerDB"
Private Const TABLE_NAME As String = "tblPatientBenefitVerification"

Public Function Insert(ByVal lngPatRPPlanID As Long, _
                       ByVal strUser As String, ByVal dteDateTime As Date, _
                       ByVal strContact As String, ByVal strNotes As String) As Long
'--------------------------------------------------------------------
'Date: 10/26/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblPatientBenefitVerfication table utilizing
'              a stored procedure.
'Parameters: lngPatRPPlanID - Patient Resposible Party Plan ID number to insert into table
'            strUser - System User that will be inserted into the table.
'            dteDateTime - The date/time of the entry that will be inserted into the table.
'            strContact - The Name of the person contacted about the verification
'            strNotes - Notes about the verification
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspInsPatBenefitVerification"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@ID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@User", adVarChar, adParamInput, 30, strUser)
        .Parameters.Append .CreateParameter("@DateTime", adDBTimeStamp, adParamInput, , dteDateTime)
        .Parameters.Append .CreateParameter("@Contact", adVarChar, adParamInput, 30, strContact)
        .Parameters.Append .CreateParameter("@Notes", adVarChar, adParamInput, 256, strNotes)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    
    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords

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
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Function FetchByPatRPPlan(ByVal lngPatRPPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/26/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPatientBenefitVerfication table.     '
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
        .CommandText = "uspSelBenefitVerificationByPatRPPlanID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatRPPlan", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchByPatRPPlan = rst
    
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

Public Function FetchPendingVerif(ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/26/2005                                                   '
'Author: Duane C Orth                                              '
'Description:  Retrieves records from the tblPatientBenefitVerfication table.     '
'Parameters: lngUserID - Provider ID number to fetch from table '
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
        .CommandText = "uspSelPatientsToBeVerifiedByUserID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@UserID", adInteger, adParamInput, , lngUserID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
    
    Set FetchPendingVerif = rst
    
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
