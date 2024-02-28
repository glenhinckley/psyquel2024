VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEligibilityDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CEligibilityDB
'Date: 12/20/2023
'Author: DCO
'Description:  MTS object designed to host methods associated with data affecting Eligibility tables.
'--------------------------------------------------------------------
' Revision History:
'
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CEligibilityDB"
Private Const TABLE_NAME As String = "tblEligibility"

Public Function FetchByID(ByVal lngID As Long, Optional ByVal strDataBase As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Retrieves records from the tblEligibility table utilizing
'              a stored procedure.
'Parameters: None
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command

    On Error GoTo ErrTrap:

    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspSelEligibilityByID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@EligibilityID", adInteger, adParamInput, , lngID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    Select Case strDataBase
          Case "Test"
               cnn.Open (CONST_TEST_CNN)
          Case "PsyquelDirect"
               cnn.Open (CONST_DIRECT_CNN)
          Case Else
               cnn.Open (CONST_PSYQUEL_CNN)
    End Select
    Set cmd.ActiveConnection = cnn

    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly

    Set FetchByID = rst
    
    'Disconnect the recordset
    Set cmd = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set rst = Nothing
    Set cmd = Nothing
    Set cnn = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Insert(ByVal strCompanyID As String, ByVal strInsuredFirstName As String, ByVal strInsuredLastName As String, ByVal strInsuredDOB As Date, _
      ByVal lngCPCID As Long, ByVal strPayerCode As String, ByVal strPatientRelation As String, ByVal strProviderFirstName As String, _
      ByVal strProviderLastName As String, ByVal strProviderNPI As String, ByVal strTIN As String, ByVal strTINType As String, _
      ByVal strCardNum As String, ByVal strGroupNum As String, ByVal strPatientLastName As String, ByVal strPatientFirstName As String, _
      ByVal dtePatientDOB As Date, ByVal strPatientSex As String, ByVal strGroupName As String, ByVal strCPT As String, ByVal strSecCPT As String, _
      ByVal strQueuedYN As String, ByVal strUserName As String, Optional ByVal strDataBase As String) As Long
 
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Inserts a row into the tblEligibility table utilizing
'              a stored procedure.
'Parameters: All data values to be inserted into the table
'Returns: ID (Primary Key) of the inserted row.
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspInsEligibility"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
         .Parameters.Append .CreateParameter("@EligibilityID", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("@CompanyID", adVarChar, adParamInput, 30, strCompanyID)
         .Parameters.Append .CreateParameter("@InsuredFirstName", adVarChar, adParamInput, 60, strInsuredFirstName)
         .Parameters.Append .CreateParameter("@InsuredLastName", adVarChar, adParamInput, 60, strInsuredLastName)
         .Parameters.Append .CreateParameter("@InsuredDOB", adDBTimeStamp, adParamInput, , IIf(CLng(strInsuredDOB) = 0, Null, strInsuredDOB))
         .Parameters.Append .CreateParameter("@CPCID", adInteger, adParamInput, , lngCPCID)
         .Parameters.Append .CreateParameter("@PayerCode", adVarChar, adParamInput, 12, strPayerCode)
         .Parameters.Append .CreateParameter("@PatientRelation", adVarChar, adParamInput, 20, strPatientRelation)
         .Parameters.Append .CreateParameter("@ProviderFirstName", adVarChar, adParamInput, 60, strProviderFirstName)
         .Parameters.Append .CreateParameter("@ProviderLastName", adVarChar, adParamInput, 60, strProviderLastName)
         .Parameters.Append .CreateParameter("@ProviderNPI", adVarChar, adParamInput, 10, strProviderNPI)
         .Parameters.Append .CreateParameter("@TIN", adVarChar, adParamInput, 30, strTIN)
         .Parameters.Append .CreateParameter("@TINType", adVarChar, adParamInput, 1, strTINType)
         .Parameters.Append .CreateParameter("@CardNum", adVarChar, adParamInput, 30, strCardNum)
         .Parameters.Append .CreateParameter("@GroupNum", adVarChar, adParamInput, 30, strGroupNum)
         .Parameters.Append .CreateParameter("@PatientLastName", adVarChar, adParamInput, 60, strPatientLastName)
         .Parameters.Append .CreateParameter("@PatientFirstName", adVarChar, adParamInput, 60, strPatientFirstName)
         .Parameters.Append .CreateParameter("@PatientDOB", adDBTimeStamp, adParamInput, , IIf(CLng(dtePatientDOB) = 0, Null, dtePatientDOB))
         .Parameters.Append .CreateParameter("@PatientSex", adVarChar, adParamInput, 1, strPatientSex)
         .Parameters.Append .CreateParameter("@GroupName", adVarChar, adParamInput, 60, strGroupName)
         .Parameters.Append .CreateParameter("@CPT", adVarChar, adParamInput, 5, strCPT)
         .Parameters.Append .CreateParameter("@SecCPT", adVarChar, adParamInput, 5, strSecCPT)
         .Parameters.Append .CreateParameter("@QueuedYN", adVarChar, adParamInput, 1, strQueuedYN)
         .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 50, strUserName)
         .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Select Case strDataBase
          Case "Test"
               cnnSQL.Open (CONST_TEST_CNN)
          Case "PsyquelDirect"
               cnnSQL.Open (CONST_DIRECT_CNN)
          Case Else
               cnnSQL.Open (CONST_PSYQUEL_CNN)
    End Select
    
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
Public Sub Update(ByVal lngEligibilityID As Long, ByVal strCompanyID As String, ByVal strInsuredFirstName As String, ByVal strInsuredLastName As String, ByVal strInsuredDOB As Date, _
      ByVal lngCPCID As Long, ByVal strPayerCode As String, ByVal strPatientRelation As String, ByVal strProviderFirstName As String, _
      ByVal strProviderLastName As String, ByVal strProviderNPI As String, ByVal strTIN As String, ByVal strTINType As String, _
      ByVal strCardNum As String, ByVal strGroupNum As String, ByVal strPatientLastName As String, ByVal strPatientFirstName As String, _
      ByVal dtePatientDOB As Date, ByVal strPatientSex As String, ByVal strGroupName As String, ByVal strCPT As String, ByVal strSecCPT As String, _
      ByVal strQueuedYN As String, ByVal dteDateProcessed As Date, ByVal strUserName As String, Optional ByVal strDataBase As String)
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Updates a row into tblEligibility table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspUpdEligibility"
    cmdSQL.CommandType = adCmdStoredProc

    With cmdSQL
         .Parameters.Append .CreateParameter("@EligibilityID", adInteger, adParamInput, , lngEligibilityID)
         .Parameters.Append .CreateParameter("@CompanyID", adVarChar, adParamInput, 30, strCompanyID)
         .Parameters.Append .CreateParameter("@InsuredFirstName", adVarChar, adParamInput, 60, strInsuredFirstName)
         .Parameters.Append .CreateParameter("@InsuredLastName", adVarChar, adParamInput, 60, strInsuredLastName)
         .Parameters.Append .CreateParameter("@InsuredDOB", adDBTimeStamp, adParamInput, , IIf(CLng(strInsuredDOB) = 0, Null, strInsuredDOB))
         .Parameters.Append .CreateParameter("@CPCID", adInteger, adParamInput, , lngCPCID)
         .Parameters.Append .CreateParameter("@PayerCode", adVarChar, adParamInput, 12, strPayerCode)
         .Parameters.Append .CreateParameter("@PatientRelation", adVarChar, adParamInput, 20, strPatientRelation)
         .Parameters.Append .CreateParameter("@ProviderFirstName", adVarChar, adParamInput, 60, strProviderFirstName)
         .Parameters.Append .CreateParameter("@ProviderLastName", adVarChar, adParamInput, 60, strProviderLastName)
         .Parameters.Append .CreateParameter("@ProviderNPI", adVarChar, adParamInput, 10, strProviderNPI)
         .Parameters.Append .CreateParameter("@TIN", adVarChar, adParamInput, 30, strTIN)
         .Parameters.Append .CreateParameter("@TINType", adVarChar, adParamInput, 1, strTINType)
         .Parameters.Append .CreateParameter("@CardNum", adVarChar, adParamInput, 30, strCardNum)
         .Parameters.Append .CreateParameter("@GroupNum", adVarChar, adParamInput, 30, strGroupNum)
         .Parameters.Append .CreateParameter("@PatientLastName", adVarChar, adParamInput, 60, strPatientLastName)
         .Parameters.Append .CreateParameter("@PatientFirstName", adVarChar, adParamInput, 60, strPatientFirstName)
         .Parameters.Append .CreateParameter("@PatientDOB", adDBTimeStamp, adParamInput, , IIf(CLng(dtePatientDOB) = 0, Null, dtePatientDOB))
         .Parameters.Append .CreateParameter("@PatientSex", adVarChar, adParamInput, 1, strPatientSex)
         .Parameters.Append .CreateParameter("@GroupName", adVarChar, adParamInput, 60, strGroupName)
         .Parameters.Append .CreateParameter("@CPT", adVarChar, adParamInput, 5, strCPT)
         .Parameters.Append .CreateParameter("@SecCPT", adVarChar, adParamInput, 5, strSecCPT)
         .Parameters.Append .CreateParameter("@QueuedYN", adVarChar, adParamInput, 1, strQueuedYN)
         .Parameters.Append .CreateParameter("@DateProcessed", adDBTimeStamp, adParamInput, , IIf(CLng(dteDateProcessed) = 0, Null, dteDateProcessed))
         .Parameters.Append .CreateParameter("@ProcessedBy", adVarChar, adParamInput, 50, strUserName)
         .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With

    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Select Case strDataBase
          Case "Test"
               cnnSQL.Open (CONST_TEST_CNN)
          Case "PsyquelDirect"
               cnnSQL.Open (CONST_DIRECT_CNN)
          Case Else
               cnnSQL.Open (CONST_PSYQUEL_CNN)
    End Select

    'Assign the connection to the Command object and execute the stored procedure
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords

    If cmdSQL.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If

    'Close the connection and free all resources
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Sub

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Public Function Exists(ByVal lngID As Long, Optional ByVal strDataBase As String)

'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Checks for existing entries in tblEligibilityTx
'Parameters: lngID - fldEligibilityID number to check for existing entries in table
'
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
    strSQL = strSQL & " fldEligibilityID = " & lngID
        
    'Instantiate the Recordset.
    Set rstSQL = New ADODB.Recordset
    rstSQL.CursorLocation = adUseClient
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Select Case strDataBase
          Case "Test"
               cnnSQL.Open (CONST_TEST_CNN)
          Case "PsyquelDirect"
               cnnSQL.Open (CONST_DIRECT_CNN)
          Case Else
               cnnSQL.Open (CONST_PSYQUEL_CNN)
    End Select
    
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




