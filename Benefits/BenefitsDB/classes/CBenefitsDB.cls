VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBenefitsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CBenefitsDB
'Date: 09/11/2000
'Author: Eric Pena
'Description:  MTS object designed to host methods associated with data affecting Benefits tables.
'--------------------------------------------------------------------
' Revision History:
'   R001: 11/06/2001 Richkun: Added FetchCoPayByPatientProvider() method
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CBenefitsDB"
Private Const TABLE_NAME As String = ""

Public Function FetchBenefitsByPatient(ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 09/11/2000
'Author: Eric Pena
'Description:  Retrieves records from the tblInsurance table utilizing
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
        .CommandText = "uspSelBenefitsByPatient"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn

    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly

    Set FetchBenefitsByPatient = rst
    
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


Public Function FetchCoPayByPatientProvider(ByVal lngPatientID As Long, _
       ByVal lngProviderID As Long) As Double
'--------------------------------------------------------------------
'Date: 11/07/2001
'Author: Dave Richkun
'Description:  Returns the CoPay amount for a patient's Primary plan
'              based on the credentials of a Provider
'Parameters: lngPatientID - ID of the patient
'            lngProviderID - ID of the Provider
'Returns: Copay amount
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
        .CommandText = "uspSelCoPayByPatientProvider"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatientID)
        .Parameters.Append .CreateParameter("@ProviderID", adInteger, adParamInput, , lngProviderID)
    End With
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn

    'Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly

    FetchCoPayByPatientProvider = rst.Fields("fldCopay").Value
    
    'Disconnect the recordset
    Set rst = Nothing
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


