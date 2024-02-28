VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatRPPlanRuleDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatRPPlanRuleDB
'Date: 11/07/2000
'Author: Morgan Travis
'Description:  MTS object designed to allow insertions into the tblPatRPPlanRule table.
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
Option Explicit

Private Const TABLE_NAME As String = "tblPatRPPlanRule"
Private Const CLASS_NAME As String = "CPatRPPlanRuleDB"



Public Function Update(ByVal lngPatRPPlanID As Long, ByVal strRPRelat As String, ByVal strCardNum As String, ByVal strGroupNum As String, _
                                    ByVal strMedicareHICNum As String, ByVal curAmtOfGrpDeductPd As Currency, ByVal dteFiscalEndDate As Date, _
                                    ByVal curDeduct As Currency, ByVal strGroupDeduct As String, ByVal curDeductRem As Currency, _
                                    ByVal curDollarLimit As Currency, ByVal curDollarsApplied As Currency, ByVal intVisitLimit As Integer, _
                                    ByVal intVisitsApplied As Integer, ByVal strCatCap As String, ByVal curCatCapAmt As Currency, _
                                    ByVal strCertReqd As String, ByVal strNoHCFA As String, ByVal strNoInvoice As String, ByVal strOutpatientRpt As String, _
                                    ByVal strVerbalRept As String, ByVal intNumVisit As Integer, ByVal intDaysPrior As Integer, ByVal strElectClaim As String, ByVal strRecurrCopay As String)
'--------------------------------------------------------------------
'Date: 03/01/2000
'Author: Eric Pena
'Description:  Creates a billing association for the given patient-rp-plan combonation
'Parameters: lngPatRPPlanID - The ID of the key in tblPatRPPlanRule
'                  All other fields represent columns to be updated
'Returns: T if the association is enabled/exists, F otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String
    Dim intCTR As Integer
    Dim lngID As Long
    On Error GoTo ErrTrap:
    
    'Acquire the database connection.
    Set cnnSQL = New ADODB.Connection
    Call cnnSQL.Open(CONST_PSYQUEL_CNN)
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    With cmdSQL
        .CommandText = "uspUpdPatRPPlanRule"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@RPRelat", adVarChar, adParamInput, 30, strRPRelat)
        .Parameters.Append .CreateParameter("@CardNum", adVarChar, adParamInput, 30, strCardNum)
        .Parameters.Append .CreateParameter("@GroupNum", adVarChar, adParamInput, 30, strGroupNum)
        .Parameters.Append .CreateParameter("@MedicareHICNum", adVarChar, adParamInput, 30, strMedicareHICNum)
        .Parameters.Append .CreateParameter("@AmtOfGrpDeductPd", adCurrency, adParamInput, , curAmtOfGrpDeductPd)
        .Parameters.Append .CreateParameter("@FiscalEndDate", adDBTimeStamp, adParamInput, , dteFiscalEndDate)
        .Parameters.Append .CreateParameter("@Deduct", adCurrency, adParamInput, , curDeduct)
        .Parameters.Append .CreateParameter("@GroupDeduct", adVarChar, adParamInput, 1, strGroupDeduct)
        .Parameters.Append .CreateParameter("@DeductRem", adCurrency, adParamInput, , curDeductRem)
        .Parameters.Append .CreateParameter("@DollarLimit", adCurrency, adParamInput, , curDollarLimit)
        .Parameters.Append .CreateParameter("@DollarsApplied", adCurrency, adParamInput, , curDollarsApplied)
        .Parameters.Append .CreateParameter("@VisitLimit", adInteger, adParamInput, , intVisitLimit)
        .Parameters.Append .CreateParameter("@VisitsApplied", adInteger, adParamInput, , intVisitsApplied)
        .Parameters.Append .CreateParameter("@CatCap", adVarChar, adParamInput, 1, strCatCap)
        .Parameters.Append .CreateParameter("@CatCapAmt", adCurrency, adParamInput, , curCatCapAmt)
        .Parameters.Append .CreateParameter("@CertReqd", adVarChar, adParamInput, 1, strCertReqd)
        .Parameters.Append .CreateParameter("@NoHCFA", adVarChar, adParamInput, 1, strNoHCFA)
        .Parameters.Append .CreateParameter("@NoInvoice", adVarChar, adParamInput, 1, strNoInvoice)
        .Parameters.Append .CreateParameter("@OutpatientRpt", adVarChar, adParamInput, 1, strOutpatientRpt)
        .Parameters.Append .CreateParameter("@VerbalRept", adVarChar, adParamInput, 1, strVerbalRept)
        .Parameters.Append .CreateParameter("@NumVisit", adInteger, adParamInput, , intNumVisit)
        .Parameters.Append .CreateParameter("@DaysPrior", adInteger, adParamInput, , intDaysPrior)
        .Parameters.Append .CreateParameter("@ElectClaim", adVarChar, adParamInput, 1, strElectClaim)
        .Parameters.Append .CreateParameter("@RecurrCopay", adVarChar, adParamInput, 1, strRecurrCopay)
        .Parameters.Append .CreateParameter("@ErrorNum", adInteger, adParamOutput, , 0)
    End With
    Set cmdSQL.ActiveConnection = cnnSQL
    cmdSQL.Execute ', , adExecuteNoRecords
    'Check the ErrorNum parameter before deciding to commit the transaction
    If cmdSQL.Parameters("@ErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    cnnSQL.Close
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    'Signal successful completion
    GetObjectContext.SetComplete
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set cmdSQL.ActiveConnection = Nothing
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function FetchPatRPPlanRule(ByVal lngPatID As Long, _
                                ByVal lngRPID As Long, _
                                ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 04/19/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Returns a row from tblPatRPPlanRule with the         '
'              matching parameters                                  '
'Parameters: lngPatID - ID of Patient                               '
'            lngRPID - ID of RP                                     '
'            lngPlanID - ID of Plan                                 '
'Returns: ADODB Recordset if match found, otherwise Null            '
'--------------------------------------------------------------------
'Note:  Notice how much easier to read this format is?...Eric       '
'-------------------------------------------------------------------'
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    
    
    On Error GoTo ErrTrap:
    
    
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
     
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    ' Create the parameter objects
    With cmd
        .CommandText = "uspSelPatientRPPlanRule"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatientID", adInteger, adParamInput, , lngPatID)
        .Parameters.Append .CreateParameter("@RPID", adInteger, adParamInput, , lngRPID)
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , lngPlanID)
    End With
    
    ' Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
     
    Set FetchPatRPPlanRule = rst
     
    ' Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
     
    ' Signal successful completion
    GetObjectContext.SetComplete
     
    Exit Function
     
ErrTrap:
    ' Signal incompletion and raise the error to the calling environment.
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set rst = Nothing
    Set cnn = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function

Public Function FetchPatRPPlanRuleByPatRPPlanID(ByVal lngPatRPPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/10/2001                                                   '
'Author: Morgan Travis                                              '
'Description:  Returns a row from tblPatRPPlanRule with the         '
'              matching parameters                                  '
'Parameters: lngPatRPPlanID - ID of Patient                         '
'                                                                   '
'Returns: ADODB Recordset if match found, otherwise Null            '
'--------------------------------------------------------------------

    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim cnn As ADODB.Connection
    
    
    On Error GoTo ErrTrap:
    
    
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
     
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    ' Create the parameter objects
    With cmd
        .CommandText = "uspSelPatientRPPlanRuleByPatRPPlanID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
    End With
    
    ' Execute the stored procedure
    rst.CursorLocation = adUseClient
    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly
     
    Set FetchPatRPPlanRuleByPatRPPlanID = rst
     
    ' Disconnect the recordset
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set cnn = Nothing
     
    ' Signal successful completion
    GetObjectContext.SetComplete
     
    Exit Function
     
ErrTrap:
    ' Signal incompletion and raise the error to the calling environment.
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rst.ActiveConnection = Nothing
    Set rst = Nothing
    Set cnn = Nothing
    GetObjectContext.SetAbort
    Call RaiseError(Err, Err.Number, CLASS_NAME)
 End Function

Public Function UpdatePatientCounters(ByVal lngPatRPPlanID As Long, ByVal intVisitCounter As Integer, ByVal curAmtApplied As Currency, ByVal curDeductible As Currency)
'--------------------------------------------------------------------
'Date: 01/10/2001
'Author: Morgan Travis
'Description:  Updates Patient visit, amount applied, and deductible counters
'Parameters:
'Returns:
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command

    On Error GoTo ErrTrap:

    Set cmd = New ADODB.Command
    Set cnn = New ADODB.Connection
    
    'Create the parameter objects
    With cmd
        .CommandText = "uspUpdPatientCounters"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@VisitCounter", adInteger, adParamInput, , intVisitCounter)
        .Parameters.Append .CreateParameter("@AmtApplied", adInteger, adParamInput, , curAmtApplied)
        .Parameters.Append .CreateParameter("@Deductible", adInteger, adParamInput, , curDeductible)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
    End With
    
    'Acquire the database connection.
    cnn.Open (CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn

    'Execute the stored procedure
    cmd.Execute , , adExecuteNoRecords

    If cmd.Parameters("@SQLErrorNum").Value <> 0 Then
        GoTo ErrTrap
    End If
    
    'Disconnect the recordset
    Set cmd = Nothing
    Set cnn = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set cmd = Nothing
    Set cnn = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



