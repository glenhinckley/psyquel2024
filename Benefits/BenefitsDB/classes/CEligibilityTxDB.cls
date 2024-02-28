VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEligibilityTxDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CEligibilityTxDB
'Date: 12/20/2023
'Author: DCO
'Description:  MTS object designed to host methods associated with data affecting EligibilityTx tables.
'--------------------------------------------------------------------
' Revision History:
'
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CEligibilityTxDB"
Private Const TABLE_NAME As String = "tblEligibilityTx"

Public Function FetchByID(ByVal lngID As Long, Optional ByVal strDataBase As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Retrieves records from the tblEligibilityTx table utilizing
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
        .CommandText = "uspSelEligibilityTxByID"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@EligibilityTxID", adInteger, adParamInput, , lngID)
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
Public Function Insert(ByVal lngEligibilityID As Long, ByVal lngVisitLimit As Long, ByVal lngVisitsApplied As Long, ByVal strHasCoverageYN As String, ByVal strOutOfNetworkYN As String, _
         ByVal lngNumVisit As Long, ByVal strBenefitsText As String, ByVal strNotes As String, _
         ByVal strPlanName As String, ByVal strPlanAddress As String, ByVal strPlanCity As String, ByVal strPlanState As String, ByVal strPlanZip As String, _
         ByVal dteEffDate As Date, ByVal dteTermDate As Date, ByVal strMSPCode As String, ByVal dteVerifyDate As Date, ByVal strQueuedYN As String, _
         ByRef varINN As Variant, ByRef varOON As Variant, ByRef varCoPay As Variant, ByRef varCert As Variant, _
         ByVal dteDateProcessed As Date, ByVal strUserName As String, Optional ByVal strDataBase As String)

'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Inserts a row into the tblEligibilityTx table utilizing
'              a stored procedure.
'Parameters: All data values to be inserted into the table
'Returns: ID (Primary Key) of the inserted row.
'--------------------------------------------------------------------
   Dim curINIndDeduct, curINIndDeductRem, curINIndDeductPaid, curINIndOOP, curINIndOOPRem, curINIndOOPPaid As Currency
   Dim curINFamDeduct, curINFamDeductRem, curINFamDeductPaid, curINFamOOP, curINFamOOPRem, curINFamOOPPaid As Currency
   Dim curOONIndDeduct, curOONIndDeductRem, curOONIndDeductPaid, curOONIndOOP, curOONIndOOPRem, curOONIndOOPPaid As Currency
   Dim curOONFamDeduct, curOONFamDeductRem, curOONFamDeductPaid, curOONFamOOP, curOONFamOOPRem, curOONFamOOPPaid As Currency
   Dim lngINPriCopayTypeID As Long
   Dim curINPriCopayAmt As Currency
   Dim lngINSecCopayTypeID As Long
   Dim curINSecCopayAmt As Currency
   Dim lngINTerCopayTypeID As Long
   Dim curINTerCopayAmt As Currency
   Dim lngOONPriCopayTypeID As Long
   Dim curOONPriCopayAmt As Currency
   Dim lngOONSecCopayTypeID As Long
   Dim curOONSecCopayAmt As Currency
   Dim lngOONTerCopayTypeID As Long
   Dim curOONTerCopayAmt As Currency
   Dim strCertReqd, strCertNumber As String
   Dim dteCertStartDate, dteCertEndDate As Date
   Dim lngCertUnits As Long
   Dim strCertCPT, strTestCertReqd, strTestCertNumber As String
   Dim dteTestCertStartDate, dteTestCertEndDate As Date
   Dim lngTestCertUnits As Long
   Dim strTestCertCPT As String
   
   Dim intCtr As Integer
   Dim lngID As Long
   Dim arrINN() As Variant
   ReDim arrINN(12)
   Dim arrOON() As Variant
   ReDim arrOON(12)
   Dim arrCoPay() As Variant
   ReDim arrCoPay(12)
   Dim arrCert() As Variant
   ReDim arrCert(12)

   Dim cnnSQL As ADODB.Connection
   Dim cmdSQL As ADODB.Command

   On Error GoTo ErrTrap:
    
   If Not IsArray(varINN) Or IsEmpty(varINN) Then
      For intCtr = 0 To UBound(arrINN)
        arrINN(intCtr) = 0
      Next intCtr
   Else
      For intCtr = 0 To UBound(varINN)
        arrINN(intCtr) = IfNull(varINN(intCtr), 0)
      Next intCtr
   End If
    
   If Not IsArray(varOON) Or IsEmpty(varOON) Then
      For intCtr = 0 To UBound(arrOON)
        arrOON(intCtr) = 0
      Next intCtr
   Else
      For intCtr = 0 To UBound(varOON)
        arrOON(intCtr) = IfNull(varOON(intCtr), 0)
      Next intCtr
   End If
    
   If Not IsArray(varCoPay) Or IsEmpty(varCoPay) Then
      For intCtr = 0 To UBound(arrCoPay)
        arrCoPay(intCtr) = 0
      Next intCtr
   Else
      For intCtr = 0 To UBound(varCoPay)
        arrCoPay(intCtr) = IfNull(varCoPay(intCtr), 0)
      Next intCtr
   End If
    
   If Not IsArray(varCert) Or IsEmpty(varCert) Then
      arrCert(0) = "N"
      arrCert(1) = ""
      arrCert(2) = Null
      arrCert(3) = Null
      arrCert(4) = 0
      arrCert(5) = ""
      arrCert(6) = "N"
      arrCert(7) = ""
      arrCert(8) = Null
      arrCert(9) = Null
      arrCert(10) = 0
      arrCert(11) = ""
   Else
      arrCert(0) = IfNull(varCert(0), "N")
      arrCert(1) = IfNull(varCert(1), "")
      arrCert(2) = IIf(CLng(varCert(2)) = 0, Null, varCert(2))
      arrCert(3) = IIf(CLng(varCert(3)) = 0, Null, varCert(2))
      arrCert(4) = IfNull(varCert(4), 0)
      arrCert(5) = IfNull(varCert(5), "")
      arrCert(6) = IfNull(varCert(6), "")
      arrCert(7) = IfNull(varCert(7), "")
      arrCert(8) = IIf(CLng(varCert(8)) = 0, Null, varCert(8))
      arrCert(9) = IIf(CLng(varCert(9)) = 0, Null, varCert(9))
      arrCert(10) = IfNull(varCert(10), 0)
      arrCert(11) = IfNull(varCert(11), "")
   End If
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspInsEligibilityTx"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@EligibilityID", adInteger, adParamInput, , lngEligibilityID)
        .Parameters.Append .CreateParameter("@INIndDeduct", adCurrency, adParamInput, , arrINN(0))
        .Parameters.Append .CreateParameter("@INIndDeductRem", adCurrency, adParamInput, , arrINN(1))
        .Parameters.Append .CreateParameter("@INIndDeductPaid", adCurrency, adParamInput, , arrINN(2))
        .Parameters.Append .CreateParameter("@INIndOOP", adCurrency, adParamInput, , arrINN(3))
        .Parameters.Append .CreateParameter("@INIndOOPRem", adCurrency, adParamInput, , arrINN(4))
        .Parameters.Append .CreateParameter("@INIndOOPPaid", adCurrency, adParamInput, , arrINN(5))
        .Parameters.Append .CreateParameter("@INFamDeduct", adCurrency, adParamInput, , arrINN(6))
        .Parameters.Append .CreateParameter("@INFamDeductRem", adCurrency, adParamInput, , arrINN(7))
        .Parameters.Append .CreateParameter("@INFamDeductPaid", adCurrency, adParamInput, , arrINN(8))
        .Parameters.Append .CreateParameter("@INFamOOP", adCurrency, adParamInput, , arrINN(9))
        .Parameters.Append .CreateParameter("@INFamOOPRem", adCurrency, adParamInput, , arrINN(10))
        .Parameters.Append .CreateParameter("@INFamOOPPaid", adCurrency, adParamInput, , arrINN(11))
        .Parameters.Append .CreateParameter("@OONIndDeduct", adCurrency, adParamInput, , arrOON(0))
        .Parameters.Append .CreateParameter("@OONIndDeductRem", adCurrency, adParamInput, , arrOON(1))
        .Parameters.Append .CreateParameter("@OONIndDeductPaid", adCurrency, adParamInput, , arrOON(2))
        .Parameters.Append .CreateParameter("@OONIndOOP", adCurrency, adParamInput, , arrOON(3))
        .Parameters.Append .CreateParameter("@OONIndOOPRem", adCurrency, adParamInput, , arrOON(4))
        .Parameters.Append .CreateParameter("@OONIndOOPPaid", adCurrency, adParamInput, , arrOON(5))
        .Parameters.Append .CreateParameter("@OONFamDeduct", adCurrency, adParamInput, , arrOON(6))
        .Parameters.Append .CreateParameter("@OONFamDeductRem", adCurrency, adParamInput, , arrOON(7))
        .Parameters.Append .CreateParameter("@OONFamDeductPaid", adCurrency, adParamInput, , arrOON(8))
        .Parameters.Append .CreateParameter("@OONFamOOP", adCurrency, adParamInput, , arrOON(9))
        .Parameters.Append .CreateParameter("@OONFamOOPRem", adCurrency, adParamInput, , arrOON(10))
        .Parameters.Append .CreateParameter("@OONFamOOPPaid", adCurrency, adParamInput, , arrOON(11))
        .Parameters.Append .CreateParameter("@VisitLimit", adInteger, adParamInput, , lngVisitLimit)
        .Parameters.Append .CreateParameter("@VisitsApplied", adInteger, adParamInput, , lngVisitsApplied)
        .Parameters.Append .CreateParameter("@HasCoverageYN", adVarChar, adParamInput, 1, strHasCoverageYN)
        .Parameters.Append .CreateParameter("@OutOfNetworkYN", adVarChar, adParamInput, 1, strOutOfNetworkYN)
        .Parameters.Append .CreateParameter("@NumVisit", adInteger, adParamInput, , lngNumVisit)
        .Parameters.Append .CreateParameter("@BenefitsText", adLongVarChar, adParamInput, IIf(Len(strBenefitsText) = 0, 1, Len(strBenefitsText)))
        .Parameters("@BenefitsText").AppendChunk strBenefitsText
        .Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, IIf(Len(strNotes) = 0, 1, Len(strNotes)))
        .Parameters("@Notes").AppendChunk strNotes
        .Parameters.Append .CreateParameter("@INPriCopayTypeID", adInteger, adParamInput, , arrCoPay(0))
        .Parameters.Append .CreateParameter("@INPriCopayAmt", adCurrency, adParamInput, , arrCoPay(1))
        .Parameters.Append .CreateParameter("@INSecCopayTypeID", adInteger, adParamInput, , arrCoPay(2))
        .Parameters.Append .CreateParameter("@INSecCopayAmt", adCurrency, adParamInput, , arrCoPay(3))
        .Parameters.Append .CreateParameter("@INTerCopayTypeID", adInteger, adParamInput, , arrCoPay(4))
        .Parameters.Append .CreateParameter("@INTerCopayAmt", adCurrency, adParamInput, , arrCoPay(5))
        .Parameters.Append .CreateParameter("@OONPriCopayTypeID", adInteger, adParamInput, , arrCoPay(6))
        .Parameters.Append .CreateParameter("@OONPriCopayAmt", adCurrency, adParamInput, , arrCoPay(7))
        .Parameters.Append .CreateParameter("@OONSecCopayTypeID", adInteger, adParamInput, , arrCoPay(8))
        .Parameters.Append .CreateParameter("@OONSecCopayAmt", adCurrency, adParamInput, , arrCoPay(9))
        .Parameters.Append .CreateParameter("@OONTerCopayTypeID", adInteger, adParamInput, , arrCoPay(10))
        .Parameters.Append .CreateParameter("@OONTerCopayAmt", adCurrency, adParamInput, , arrCoPay(11))
        .Parameters.Append .CreateParameter("@PlanName", adVarChar, adParamInput, 30, strPlanName)
        .Parameters.Append .CreateParameter("@PlanAddress", adVarChar, adParamInput, 50, strPlanAddress)
        .Parameters.Append .CreateParameter("@PlanCity", adVarChar, adParamInput, 50, strPlanCity)
        .Parameters.Append .CreateParameter("@PlanState", adVarChar, adParamInput, 2, strPlanState)
        .Parameters.Append .CreateParameter("@PlanZip", adVarChar, adParamInput, 10, strPlanZip)
        .Parameters.Append .CreateParameter("@EffDate", adDBTimeStamp, adParamInput, , IIf(CLng(dteEffDate) = 0, Null, dteEffDate))
        .Parameters.Append .CreateParameter("@TermDate", adDBTimeStamp, adParamInput, , IIf(CLng(dteTermDate) = 0, Null, dteTermDate))
        .Parameters.Append .CreateParameter("@CertReqd", adVarChar, adParamInput, 1, arrCert(0))
        .Parameters.Append .CreateParameter("@CertNumber", adVarChar, adParamInput, 30, arrCert(1))
        .Parameters.Append .CreateParameter("@CertStartDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(2)) = 0, Null, arrCert(2)))
        .Parameters.Append .CreateParameter("@CertEndDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(3)) = 0, Null, arrCert(3)))
        .Parameters.Append .CreateParameter("@CertUnits", adInteger, adParamInput, , arrCert(4))
        .Parameters.Append .CreateParameter("@CertCPT", adVarChar, adParamInput, 5, arrCert(5))
        .Parameters.Append .CreateParameter("@TestCertReqd", adVarChar, adParamInput, 1, arrCert(6))
        .Parameters.Append .CreateParameter("@TestCertNumber", adVarChar, adParamInput, 10, arrCert(7))
        .Parameters.Append .CreateParameter("@TestCertStartDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(8)) = 0, Null, arrCert(8)))
        .Parameters.Append .CreateParameter("@TestCertEndDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(9)) = 0, Null, arrCert(9)))
        .Parameters.Append .CreateParameter("@TestCertUnits", adInteger, adParamInput, , arrCert(10))
        .Parameters.Append .CreateParameter("@TestCertCPT", adVarChar, adParamInput, 5, arrCert(11))
        .Parameters.Append .CreateParameter("@MSPcode", adVarChar, adParamInput, 2, strMSPCode)
        .Parameters.Append .CreateParameter("@VerifyDate", adDBTimeStamp, adParamInput, , IIf(CLng(dteVerifyDate) = 0, Null, dteVerifyDate))
        .Parameters.Append .CreateParameter("@QueuedYN", adVarChar, adParamInput, 1, strQueuedYN)
        .Parameters.Append .CreateParameter("@DateProcessed", adDBTimeStamp, adParamInput, , IIf(CLng(dteDateProcessed) = 0, Null, dteDateProcessed))
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
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set cnnSQL = Nothing
    Set cmdSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub Update(ByVal lngEligibilityID As Long, ByVal lngVisitLimit As Long, ByVal lngVisitsApplied As Long, ByVal strHasCoverageYN As String, ByVal strOutOfNetworkYN As String, _
         ByVal lngNumVisit As Long, ByVal strBenefitsText As String, ByVal strNotes As String, _
         ByVal strPlanName As String, ByVal strPlanAddress As String, ByVal strPlanCity As String, ByVal strPlanState As String, ByVal strPlanZip As String, _
         ByVal dteEffDate As Date, ByVal dteTermDate As Date, ByVal strMSPCode As String, ByVal dteVerifyDate As Date, ByVal strQueuedYN As String, _
         ByRef varINN As Variant, ByRef varOON As Variant, ByRef varCoPay As Variant, ByRef varCert As Variant, _
         ByVal dteDateProcessed As Date, ByVal strUserName As String, Optional ByVal strDataBase As String)
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Updates a row into tblEligibilityTx table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

   Dim intCtr As Integer
   Dim lngID As Long
   Dim arrINN() As Variant
   ReDim arrINN(12)
   Dim arrOON() As Variant
   ReDim arrOON(12)
   Dim arrCoPay() As Variant
   ReDim arrCoPay(12)
   Dim arrCert() As Variant
   ReDim arrCert(12)

   Dim cnnSQL As ADODB.Connection
   Dim cmdSQL As ADODB.Command

   On Error GoTo ErrTrap:
    
   If Not IsArray(varINN) Or IsEmpty(varINN) Then
      For intCtr = 0 To UBound(arrINN)
        arrINN(intCtr) = 0
      Next intCtr
   Else
      For intCtr = 0 To UBound(varINN)
        arrINN(intCtr) = IfNull(varINN(intCtr), 0)
      Next intCtr
   End If
    
   If Not IsArray(varOON) Or IsEmpty(varOON) Then
      For intCtr = 0 To UBound(arrOON)
        arrOON(intCtr) = 0
      Next intCtr
   Else
      For intCtr = 0 To UBound(varOON)
        arrOON(intCtr) = IfNull(varOON(intCtr), 0)
      Next intCtr
   End If
    
   If Not IsArray(varCoPay) Or IsEmpty(varCoPay) Then
      For intCtr = 0 To UBound(arrCoPay)
        arrCoPay(intCtr) = 0
      Next intCtr
   Else
      For intCtr = 0 To UBound(varCoPay)
        arrCoPay(intCtr) = IfNull(varCoPay(intCtr), 0)
      Next intCtr
   End If
    
   If Not IsArray(varCert) Or IsEmpty(varCert) Then
      arrCert(0) = "N"
      arrCert(1) = ""
      arrCert(2) = Null
      arrCert(3) = Null
      arrCert(4) = 0
      arrCert(5) = ""
      arrCert(6) = "N"
      arrCert(7) = ""
      arrCert(8) = Null
      arrCert(9) = Null
      arrCert(10) = 0
      arrCert(11) = ""
   Else
      arrCert(0) = IfNull(varCert(0), "N")
      arrCert(1) = IfNull(varCert(1), "")
      arrCert(2) = IIf(CLng(varCert(2)) = 0, Null, varCert(2))
      arrCert(3) = IIf(CLng(varCert(3)) = 0, Null, varCert(2))
      arrCert(4) = IfNull(varCert(4), 0)
      arrCert(5) = IfNull(varCert(5), "")
      arrCert(6) = IfNull(varCert(6), "")
      arrCert(7) = IfNull(varCert(7), "")
      arrCert(8) = IIf(CLng(varCert(8)) = 0, Null, varCert(8))
      arrCert(9) = IIf(CLng(varCert(9)) = 0, Null, varCert(9))
      arrCert(10) = IfNull(varCert(10), 0)
      arrCert(11) = IfNull(varCert(11), "")
   End If

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspUpdEligibilityTx"
    cmdSQL.CommandType = adCmdStoredProc

    With cmdSQL
        .Parameters.Append .CreateParameter("@EligibilityID", adInteger, adParamInput, , lngEligibilityID)
        .Parameters.Append .CreateParameter("@INIndDeduct", adCurrency, adParamInput, , arrINN(0))
        .Parameters.Append .CreateParameter("@INIndDeductRem", adCurrency, adParamInput, , arrINN(1))
        .Parameters.Append .CreateParameter("@INIndDeductPaid", adCurrency, adParamInput, , arrINN(2))
        .Parameters.Append .CreateParameter("@INIndOOP", adCurrency, adParamInput, , arrINN(3))
        .Parameters.Append .CreateParameter("@INIndOOPRem", adCurrency, adParamInput, , arrINN(4))
        .Parameters.Append .CreateParameter("@INIndOOPPaid", adCurrency, adParamInput, , arrINN(5))
        .Parameters.Append .CreateParameter("@INFamDeduct", adCurrency, adParamInput, , arrINN(6))
        .Parameters.Append .CreateParameter("@INFamDeductRem", adCurrency, adParamInput, , arrINN(7))
        .Parameters.Append .CreateParameter("@INFamDeductPaid", adCurrency, adParamInput, , arrINN(8))
        .Parameters.Append .CreateParameter("@INFamOOP", adCurrency, adParamInput, , arrINN(9))
        .Parameters.Append .CreateParameter("@INFamOOPRem", adCurrency, adParamInput, , arrINN(10))
        .Parameters.Append .CreateParameter("@INFamOOPPaid", adCurrency, adParamInput, , arrINN(11))
        .Parameters.Append .CreateParameter("@OONIndDeduct", adCurrency, adParamInput, , arrOON(0))
        .Parameters.Append .CreateParameter("@OONIndDeductRem", adCurrency, adParamInput, , arrOON(1))
        .Parameters.Append .CreateParameter("@OONIndDeductPaid", adCurrency, adParamInput, , arrOON(2))
        .Parameters.Append .CreateParameter("@OONIndOOP", adCurrency, adParamInput, , arrOON(3))
        .Parameters.Append .CreateParameter("@OONIndOOPRem", adCurrency, adParamInput, , arrOON(4))
        .Parameters.Append .CreateParameter("@OONIndOOPPaid", adCurrency, adParamInput, , arrOON(5))
        .Parameters.Append .CreateParameter("@OONFamDeduct", adCurrency, adParamInput, , arrOON(6))
        .Parameters.Append .CreateParameter("@OONFamDeductRem", adCurrency, adParamInput, , arrOON(7))
        .Parameters.Append .CreateParameter("@OONFamDeductPaid", adCurrency, adParamInput, , arrOON(8))
        .Parameters.Append .CreateParameter("@OONFamOOP", adCurrency, adParamInput, , arrOON(9))
        .Parameters.Append .CreateParameter("@OONFamOOPRem", adCurrency, adParamInput, , arrOON(10))
        .Parameters.Append .CreateParameter("@OONFamOOPPaid", adCurrency, adParamInput, , arrOON(11))
        .Parameters.Append .CreateParameter("@VisitLimit", adInteger, adParamInput, , lngVisitLimit)
        .Parameters.Append .CreateParameter("@VisitsApplied", adInteger, adParamInput, , lngVisitsApplied)
        .Parameters.Append .CreateParameter("@HasCoverageYN", adVarChar, adParamInput, 1, strHasCoverageYN)
        .Parameters.Append .CreateParameter("@OutOfNetworkYN", adVarChar, adParamInput, 1, strOutOfNetworkYN)
        .Parameters.Append .CreateParameter("@NumVisit", adInteger, adParamInput, , lngNumVisit)
        .Parameters.Append .CreateParameter("@BenefitsText", adLongVarChar, adParamInput, IIf(Len(strBenefitsText) = 0, 1, Len(strBenefitsText)))
        .Parameters("@BenefitsText").AppendChunk strBenefitsText
        .Parameters.Append .CreateParameter("@Notes", adLongVarChar, adParamInput, IIf(Len(strNotes) = 0, 1, Len(strNotes)))
        .Parameters("@Notes").AppendChunk strNotes
        .Parameters.Append .CreateParameter("@INPriCopayTypeID", adInteger, adParamInput, , arrCoPay(0))
        .Parameters.Append .CreateParameter("@INPriCopayAmt", adCurrency, adParamInput, , arrCoPay(1))
        .Parameters.Append .CreateParameter("@INSecCopayTypeID", adInteger, adParamInput, , arrCoPay(2))
        .Parameters.Append .CreateParameter("@INSecCopayAmt", adCurrency, adParamInput, , arrCoPay(3))
        .Parameters.Append .CreateParameter("@INTerCopayTypeID", adInteger, adParamInput, , arrCoPay(4))
        .Parameters.Append .CreateParameter("@INTerCopayAmt", adCurrency, adParamInput, , arrCoPay(5))
        .Parameters.Append .CreateParameter("@OONPriCopayTypeID", adInteger, adParamInput, , arrCoPay(6))
        .Parameters.Append .CreateParameter("@OONPriCopayAmt", adCurrency, adParamInput, , arrCoPay(7))
        .Parameters.Append .CreateParameter("@OONSecCopayTypeID", adInteger, adParamInput, , arrCoPay(8))
        .Parameters.Append .CreateParameter("@OONSecCopayAmt", adCurrency, adParamInput, , arrCoPay(9))
        .Parameters.Append .CreateParameter("@OONTerCopayTypeID", adInteger, adParamInput, , arrCoPay(10))
        .Parameters.Append .CreateParameter("@OONTerCopayAmt", adCurrency, adParamInput, , arrCoPay(11))
        .Parameters.Append .CreateParameter("@PlanName", adVarChar, adParamInput, 30, strPlanName)
        .Parameters.Append .CreateParameter("@PlanAddress", adVarChar, adParamInput, 50, strPlanAddress)
        .Parameters.Append .CreateParameter("@PlanCity", adVarChar, adParamInput, 50, strPlanCity)
        .Parameters.Append .CreateParameter("@PlanState", adVarChar, adParamInput, 2, strPlanState)
        .Parameters.Append .CreateParameter("@PlanZip", adVarChar, adParamInput, 10, strPlanZip)
        .Parameters.Append .CreateParameter("@EffDate", adDBTimeStamp, adParamInput, , IIf(CLng(dteEffDate) = 0, Null, dteEffDate))
        .Parameters.Append .CreateParameter("@TermDate", adDBTimeStamp, adParamInput, , IIf(CLng(dteTermDate) = 0, Null, dteTermDate))
        .Parameters.Append .CreateParameter("@CertReqd", adVarChar, adParamInput, 1, arrCert(0))
        .Parameters.Append .CreateParameter("@CertNumber", adVarChar, adParamInput, 30, arrCert(1))
        .Parameters.Append .CreateParameter("@CertStartDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(2)) = 0, Null, arrCert(2)))
        .Parameters.Append .CreateParameter("@CertEndDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(3)) = 0, Null, arrCert(3)))
        .Parameters.Append .CreateParameter("@CertUnits", adInteger, adParamInput, , arrCert(4))
        .Parameters.Append .CreateParameter("@CertCPT", adVarChar, adParamInput, 5, arrCert(5))
        .Parameters.Append .CreateParameter("@TestCertReqd", adVarChar, adParamInput, 1, arrCert(6))
        .Parameters.Append .CreateParameter("@TestCertNumber", adVarChar, adParamInput, 10, arrCert(7))
        .Parameters.Append .CreateParameter("@TestCertStartDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(8)) = 0, Null, arrCert(8)))
        .Parameters.Append .CreateParameter("@TestCertEndDate", adDBTimeStamp, adParamInput, , IIf(CLng(arrCert(9)) = 0, Null, arrCert(9)))
        .Parameters.Append .CreateParameter("@TestCertUnits", adInteger, adParamInput, , arrCert(10))
        .Parameters.Append .CreateParameter("@TestCertCPT", adVarChar, adParamInput, 5, arrCert(11))
        .Parameters.Append .CreateParameter("@MSPcode", adVarChar, adParamInput, 2, strMSPCode)
        .Parameters.Append .CreateParameter("@VerifyDate", adDBTimeStamp, adParamInput, , IIf(CLng(dteVerifyDate) = 0, Null, dteVerifyDate))
        .Parameters.Append .CreateParameter("@QueuedYN", adVarChar, adParamInput, 1, strQueuedYN)
        .Parameters.Append .CreateParameter("@DateProcessed", adDBTimeStamp, adParamInput, , IIf(CLng(dteDateProcessed) = 0, Null, dteDateProcessed))
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





