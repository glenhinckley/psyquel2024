VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatientBenefitDB"
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
' Revision History:
'  R001: 02/16/2001 Richkun - Altered Insert() method so that a limited parameter
'       set is used.  Parameter set is based on changes made to Benefits
'       entry the week of Feb 12th, 2001 where it was decided that only
'       Co-pays and general information would be captured for patient benefits.
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CPatientBenefitDB"
Private Const TABLE_NAME As String = "tblPatRPPlanRule"

Public Function Insert(ByVal lngPatRPPlanID As Long, ByVal strCardNumber As String, _
                       ByVal strGroupNumber As String, ByVal strCertRequired As String, _
                       ByVal lngPriCopayID As Long, ByVal lngSecCopayID As Long, ByVal lngTerCopayID As Long, _
                       ByVal dblPriCopayAmt As Double, ByVal dblSecCopayAmt As Double, _
                       ByVal dblTerCopayAmt As Double, _
                       ByVal strAdditionalInfo As String, ByVal strRPRelat As String, ByVal lngVerPlanID As Long, ByVal strPlanName As String, _
                       ByVal strPlanAddress As String, ByVal strPlanCity As String, ByVal strPlanState As String, _
                       ByVal strPlanZip As String, ByVal strInsPhone As String, ByVal strInsdName As String, _
                       ByVal strPayerCode As String, ByVal strVerifyElectronic As String, _
                       ByVal strOutOfNetwork As String, ByVal strContracted As String, ByVal dblDeduct As Double, ByVal dblDeductRem As Double, _
                       ByVal strMSPCode As String, ByVal strUserName As String) As Long
'--------------------------------------------------------------------
'Date: 02/16/2001
'Author: Dave Richkun
'Description:  Inserts a row into the tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters: All data values to be inserted into the table
'Returns: ID (Primary Key) of the inserted row.
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:
    
    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspInsPatRPPlanRule"
    cmdSQL.CommandType = adCmdStoredProc
    
    With cmdSQL
        .Parameters.Append .CreateParameter("@PatRPPlanRuleID", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@CardNumber", adVarChar, adParamInput, 30, strCardNumber)
        .Parameters.Append .CreateParameter("@GrpNumber", adVarChar, adParamInput, 30, strGroupNumber)
        .Parameters.Append .CreateParameter("@CertRequired", adVarChar, adParamInput, 1, strCertRequired)
        .Parameters.Append .CreateParameter("@PriCopayID", adInteger, adParamInput, , lngPriCopayID)
        .Parameters.Append .CreateParameter("@SecCopayID", adInteger, adParamInput, , lngSecCopayID)
        .Parameters.Append .CreateParameter("@TerCopayID", adInteger, adParamInput, , lngTerCopayID)
        .Parameters.Append .CreateParameter("@PriCopayAmt", adCurrency, adParamInput, , dblPriCopayAmt)
        .Parameters.Append .CreateParameter("@SecCopayAmt", adCurrency, adParamInput, , dblSecCopayAmt)
        .Parameters.Append .CreateParameter("@TerCopayAmt", adCurrency, adParamInput, , dblTerCopayAmt)
        .Parameters.Append .CreateParameter("@Deduct", adCurrency, adParamInput, , dblDeduct)
        .Parameters.Append .CreateParameter("@DeductRem", adCurrency, adParamInput, , dblDeductRem)
        .Parameters.Append .CreateParameter("@BenefitsNotes", adVarChar, adParamInput, 1024, strAdditionalInfo)
        .Parameters.Append .CreateParameter("@RPRelat", adVarChar, adParamInput, 20, strRPRelat)
        .Parameters.Append .CreateParameter("@VerPlanID", adInteger, adParamInput, , lngVerPlanID)
        .Parameters.Append .CreateParameter("@PlanName", adVarChar, adParamInput, 30, strPlanName)
        .Parameters.Append .CreateParameter("@PlanAddress", adVarChar, adParamInput, 50, strPlanAddress)
        .Parameters.Append .CreateParameter("@PlanCity", adVarChar, adParamInput, 50, strPlanCity)
        .Parameters.Append .CreateParameter("@PlanState", adVarChar, adParamInput, 2, strPlanState)
        .Parameters.Append .CreateParameter("@PlanZip", adVarChar, adParamInput, 10, strPlanZip)
        .Parameters.Append .CreateParameter("@InsPhone", adVarChar, adParamInput, 15, NumbersOnly(strInsPhone))
        .Parameters.Append .CreateParameter("@InsdName", adVarChar, adParamInput, 40, strInsdName)
        .Parameters.Append .CreateParameter("@PayerCode", adVarChar, adParamInput, 15, strPayerCode)
        .Parameters.Append .CreateParameter("@VerifyElectronic", adVarChar, adParamInput, 1, strVerifyElectronic)
        .Parameters.Append .CreateParameter("@OutOfNetwork", adVarChar, adParamInput, 1, strOutOfNetwork)
        .Parameters.Append .CreateParameter("@Contracted", adVarChar, adParamInput, 1, strContracted)
        .Parameters.Append .CreateParameter("@MSPCode", adVarChar, adParamInput, 2, strMSPCode)
        .Parameters.Append .CreateParameter("@AddedBy", adVarChar, adParamInput, 30, strUserName)
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


Public Sub Update(ByVal lngPatRPPlanID As Long, ByVal strRelation As String, ByVal strCardNumber As String, _
                  ByVal strGroupNumber As String, ByVal strCertRequired As String, ByVal strPayerCode As String, _
                  ByVal lngPriCopayID As Long, ByVal lngSecCopayID As Long, ByVal lngTerCopayID As Long, _
                  ByVal dblPriCopayAmt As Double, ByVal dblSecCopayAmt As Double, ByVal dblTerCopayAmt As Double, _
                  ByVal strVerifyElectronic As String, _
                  ByVal strReVerifyElectronic As String, ByVal strOutOfNetwork As String, ByVal strContracted As String, _
                  ByVal strBenefitsNotes As String, _
                  ByVal dblDeduct As Double, ByVal dblDeductRem As Double, _
                  ByVal strMSPCode As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 02/15/2001
'Author: Dave Richkun
'Description:  Updates a row into tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspUpdPatientBenefit"
    cmdSQL.CommandType = adCmdStoredProc

    With cmdSQL
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@Relation", adVarChar, adParamInput, 20, strRelation)
        .Parameters.Append .CreateParameter("@CardNumber", adVarChar, adParamInput, 30, strCardNumber)
        .Parameters.Append .CreateParameter("@GroupNumber", adVarChar, adParamInput, 30, strGroupNumber)
        .Parameters.Append .CreateParameter("@CertRequired", adVarChar, adParamInput, 1, strCertRequired)
        .Parameters.Append .CreateParameter("@PayerCode", adVarChar, adParamInput, 15, strPayerCode)
        .Parameters.Append .CreateParameter("@PriCopayID", adInteger, adParamInput, , lngPriCopayID)
        .Parameters.Append .CreateParameter("@SecCopayID", adInteger, adParamInput, , lngSecCopayID)
        .Parameters.Append .CreateParameter("@TerCopayID", adInteger, adParamInput, , lngTerCopayID)
        .Parameters.Append .CreateParameter("@PriCopayAmt", adCurrency, adParamInput, , dblPriCopayAmt)
        .Parameters.Append .CreateParameter("@SecCopayAmt", adCurrency, adParamInput, , dblSecCopayAmt)
        .Parameters.Append .CreateParameter("@TerCopayAmt", adCurrency, adParamInput, , dblTerCopayAmt)
        .Parameters.Append .CreateParameter("@Deduct", adCurrency, adParamInput, , dblDeduct)
        .Parameters.Append .CreateParameter("@DeductRem", adCurrency, adParamInput, , dblDeductRem)
        .Parameters.Append .CreateParameter("@VerifyElectronic", adVarChar, adParamInput, 1, strVerifyElectronic)
        .Parameters.Append .CreateParameter("@ReVerifyElectronic", adVarChar, adParamInput, 1, strReVerifyElectronic)
        .Parameters.Append .CreateParameter("@OutOfNetwork", adVarChar, adParamInput, 1, strOutOfNetwork)
        .Parameters.Append .CreateParameter("@Contracted", adVarChar, adParamInput, 1, strContracted)
        .Parameters.Append .CreateParameter("@MSPCode", adVarChar, adParamInput, 2, strMSPCode)
        .Parameters.Append .CreateParameter("@BenefitsNotes", adVarChar, adParamInput, 1024, strBenefitsNotes)
        .Parameters.Append .CreateParameter("@ModifiedBy", adVarChar, adParamInput, 30, strUserName)
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
Public Sub UpdateVerify(ByVal lngPatRPPlanID As Long, ByVal strRelation As String, ByVal strCardNumber As String, _
                  ByVal strGroupNumber As String, ByVal strCertRequired As String, ByVal strPayerCode As String, _
                  ByVal lngPriCopayID As Long, ByVal lngSecCopayID As Long, ByVal lngTerCopayID As Long, _
                  ByVal dblPriCopayAmt As Double, ByVal dblSecCopayAmt As Double, ByVal dblTerCopayAmt As Double, _
                  ByVal strVerifyText As String, _
                  ByVal strVerifyRejectYN As String, ByVal strOutOfNetwork As String, ByVal strContracted As String, _
                  ByVal dblDeduct As Double, ByVal dblDeductRem As Double, ByVal strMSPCode As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 02/14/2018
'Author: Duane C Orth
'Description:  Updates a row in tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command

    On Error GoTo ErrTrap:

    'Instantiate and prepare the Command object.
    Set cmdSQL = New ADODB.Command
    cmdSQL.CommandText = "uspUpdPatientBenefitVerify"
    cmdSQL.CommandType = adCmdStoredProc
 
    With cmdSQL
        .Parameters.Append .CreateParameter("@PatRPPlanID", adInteger, adParamInput, , lngPatRPPlanID)
        .Parameters.Append .CreateParameter("@Relation", adVarChar, adParamInput, 20, strRelation)
        .Parameters.Append .CreateParameter("@CardNumber", adVarChar, adParamInput, 30, strCardNumber)
        .Parameters.Append .CreateParameter("@GroupNumber", adVarChar, adParamInput, 30, strGroupNumber)
        .Parameters.Append .CreateParameter("@CertRequired", adVarChar, adParamInput, 1, strCertRequired)
        .Parameters.Append .CreateParameter("@PayerCode", adVarChar, adParamInput, 15, strPayerCode)
        .Parameters.Append .CreateParameter("@PriCopayID", adInteger, adParamInput, , lngPriCopayID)
        .Parameters.Append .CreateParameter("@SecCopayID", adInteger, adParamInput, , lngSecCopayID)
        .Parameters.Append .CreateParameter("@TerCopayID", adInteger, adParamInput, , lngTerCopayID)
        .Parameters.Append .CreateParameter("@PriCopayAmt", adCurrency, adParamInput, , dblPriCopayAmt)
        .Parameters.Append .CreateParameter("@SecCopayAmt", adCurrency, adParamInput, , dblSecCopayAmt)
        .Parameters.Append .CreateParameter("@TerCopayAmt", adCurrency, adParamInput, , dblTerCopayAmt)
        .Parameters.Append .CreateParameter("@Deduct", adCurrency, adParamInput, , dblDeduct)
        .Parameters.Append .CreateParameter("@DeductRem", adCurrency, adParamInput, , dblDeductRem)
        .Parameters.Append .CreateParameter("@VerifyText", adLongVarChar, adParamInput, IIf(Len(strVerifyText) = 0, 1, Len(strVerifyText)))
        .Parameters("@VerifyText").AppendChunk strVerifyText
        .Parameters.Append .CreateParameter("@VerifyRejectYN", adVarChar, adParamInput, 1, strVerifyRejectYN)
        .Parameters.Append .CreateParameter("@OutOfNetwork", adVarChar, adParamInput, 1, strOutOfNetwork)
        .Parameters.Append .CreateParameter("@Contracted", adVarChar, adParamInput, 1, strContracted)
        .Parameters.Append .CreateParameter("@MSPCode", adVarChar, adParamInput, 2, strMSPCode)
        .Parameters.Append .CreateParameter("@ModifiedBy", adVarChar, adParamInput, 30, strUserName)
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

Public Sub UpdateFileID(ByVal lngPatRPPlanID As Long, ByVal lngPlanID As Long, ByVal strFileID As String)
    
    Dim cnnSQL As ADODB.Connection
    Dim cmdSQL As ADODB.Command
    Dim strSQL As String

    On Error GoTo ErrTrap:

    'Prepare the SQL statement
    strSQL = "UPDATE tblPatRPPlanRule "
    strSQL = strSQL & " SET "
    strSQL = strSQL & " fldVerifyElectronicYN = 'Y', fldVerifyRejectYN = 'N', fldReVerifyElectronicYN = 'N', "
    strSQL = strSQL & " fldVerifyFileID = '" & strFileID & "' "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " fldPatRPPlanID = " & lngPatRPPlanID & " AND "
    strSQL = strSQL & " fldVerPlanID = " & lngPlanID

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
    Set cmdSQL.ActiveConnection = Nothing
    Set cmdSQL = Nothing
    cnnSQL.Close
    Set cnnSQL = Nothing

    'Signal successful completion
    'GetObjectContext.SetComplete

    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    'GetObjectContext.SetAbort
    Set cmdSQL = Nothing
    Set cnnSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


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
        .CommandText = "uspSelPatientBenefitByPatRPPlanID"
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

Public Function Exists(ByVal lngPatRPPlanID As Long)

'--------------------------------------------------------------------
'Date: 11/02/2000
'Author: Morgan Travis
'Description:  Checks for existing entries in tblPatientBenefit
'Parameters: lngPatRPPlanID - Patient RP Plan ID number to check for existing entries in table
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
    strSQL = strSQL & " fldPatRPPlanID = " & lngPatRPPlanID
        
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

Public Function VerifyPolicyNum(ByVal intPlanId As Integer, ByVal intInsID As Integer, ByVal strGroupNum As String, ByVal strPolicyNum As String) As Integer
'--------------------------------------------------------------------
'Date: 11/23/2021
'Author: Jacob Orth
'Description:  Verifies the policy number is a valid format.
'Parameters:
'Returns: The ID associated with the UserName.
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim SQLErrorNum As Long

    On Error GoTo ErrTrap:
    
    Set cnn = New ADODB.Connection
    Set cmd = New ADODB.Command

    Call cnn.Open(CONST_PSYQUEL_CNN)
    Set cmd.ActiveConnection = cnn
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "uspVerifyPolicyNum"
        .Parameters.Append .CreateParameter("@PlanID", adInteger, adParamInput, , intPlanId)
        .Parameters.Append .CreateParameter("@InsID", adInteger, adParamInput, , intInsID)
        .Parameters.Append .CreateParameter("@GroupNum", adVarChar, adParamInput, 30, strGroupNum)
        .Parameters.Append .CreateParameter("@PolicyNum", adVarChar, adParamInput, 30, strPolicyNum)
        .Parameters.Append .CreateParameter("@SQLErrorNum", adInteger, adParamOutput, , 0)
        
        Set .ActiveConnection = cnn
        .Execute
    End With

    VerifyPolicyNum = cmd.Parameters("@SQLErrorNum").Value

    'Close the connection and free all resources
    cnn.Close
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
    Call MsgBox("Error: " & Err.Description & vbLf & vbLf & "Error Number: " _
                & Err.Number, vbOKOnly + vbCritical, "Error")

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

