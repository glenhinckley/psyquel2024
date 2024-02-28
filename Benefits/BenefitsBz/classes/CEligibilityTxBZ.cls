VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEligibilityTxBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatientBenefitBZ
'Date: 12/20/2023
'Author: DCO
'Description:  MTS object designed to host methods associated with
'               data affecting Patient Benefits Verification table.
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CEligibilityTxBZ"

Public Function FetchByID(ByVal lngID As Long, Optional ByVal strDataBase As String) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  Retrieves records from the tblEligibilityTx table.   '
'Parameters: lngID - EligibilityID number to fetch from table       '
'
'Returns: Recordset of requested Eligibility
'-------------------------------------------------------------------'

    
    Dim obj As BenefitsDB.CEligibilityTxDB
    Dim strErrMsg As String
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set obj = CreateObject("BenefitsDB.CEligibilityDB")
    Set rstSQL = obj.FetchByID(lngID)
      
    Set FetchByID = rstSQL
    Set obj = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set obj = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function


Public Function Insert(ByVal lngID As Long, ByVal lngVisitLimit As Long, ByVal lngVisitsApplied As Long, ByVal strHasCoverageYN As String, ByVal strOutOfNetworkYN As String, _
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

    Dim obj As BenefitsDB.CEligibilityTxDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set obj = CreateObject("BenefitsDB.CEligibilityTxDB")
    Call obj.Insert(lngID, lngVisitLimit, lngVisitsApplied, strHasCoverageYN, strOutOfNetworkYN, _
         lngNumVisit, strBenefitsText, strNotes, strPlanName, strPlanAddress, strPlanCity, strPlanState, strPlanZip, _
         dteEffDate, dteTermDate, strMSPCode, dteVerifyDate, strQueuedYN, varINN, varOON, varCoPay, varCert, _
         dteDateProcessed, strUserName, strDataBase)
    Set obj = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set obj = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function

Public Function Update(ByVal lngID As Long, ByVal lngVisitLimit As Long, ByVal lngVisitsApplied As Long, ByVal strHasCoverageYN As String, ByVal strOutOfNetworkYN As String, _
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

    Dim obj As BenefitsDB.CEligibilityTxDB
    Dim strErrMsg As String

    On Error GoTo ErrTrap:

    Set obj = CreateObject("BenefitsDB.CEligibilityTxDB")
    Call obj.Update(lngID, lngVisitLimit, lngVisitsApplied, strHasCoverageYN, strOutOfNetworkYN, _
         lngNumVisit, strBenefitsText, strNotes, strPlanName, strPlanAddress, strPlanCity, strPlanState, strPlanZip, _
         dteEffDate, dteTermDate, strMSPCode, dteVerifyDate, strQueuedYN, varINN, varOON, varCoPay, varCert, _
         dteDateProcessed, strUserName, strDataBase)
    Set obj = Nothing

    GetObjectContext.SetComplete

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set obj = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function

Private Function VerifyData(ByVal lngID As Long, _
                       ByVal dteFiscalStart As Date, ByVal dteFiscalEnd As Date, _
                       ByVal curDeductible As Currency, ByVal strGroupDeductible As String, _
                       ByVal curDeductibleMet As Currency, ByVal curAnnualDollarLimit As Currency, _
                       ByVal curAmountApplied As Currency, ByVal lngAnnualVisitLimit As Long, _
                       ByVal lngVisits As Long, ByVal strCatastrophicCap As String, _
                       ByVal curCatastrophicCapAmt As Currency, ByVal strCertRequired As String, _
                       ByVal strNoHCFA As String, ByVal strNoInvoice As String, _
                       ByVal strOutPatientRept As String, ByVal strVerbRept As String, _
                       ByVal lngNumVisit As Long, ByVal lngDaysPrior As Long, _
                       ByVal strElectClaim As String, ByVal strRecurrCopay As String, ByVal strErrMsg As String, ByVal blnExists As Boolean) As Boolean
    
    VerifyData = False
    If lngID < 0 Then
        strErrMsg = "Patient RP Plan ID Not Valid"
        Exit Function
    End If
    
    If IsDate(dteFiscalStart) = False Or IsDate(dteFiscalEnd) = False Then
        strErrMsg = "Fiscal Start Date or Fiscal End Date Not Valid"
        Exit Function
    End If
    
    If dteFiscalStart > dteFiscalEnd Then
        strErrMsg = "Fiscal Start Date Greater Than Fiscal End Date"
        Exit Function
    End If
    
    If curDeductible < 0 Then
        strErrMsg = "Invalid Deductible"
        Exit Function
    End If
    
    If curDeductibleMet < 0 Then
        strErrMsg = "Invalid Deductible Met"
        Exit Function
    End If
    
    If Not ((strGroupDeductible <> "Y") Or (strGroupDeductible <> "y") Or _
       (strGroupDeductible <> "N") Or (strGroupDeductible <> "n")) Then
        strErrMsg = "Group Deductible Not Valid"
        Exit Function
    End If
    
    If curAnnualDollarLimit < 0 Then
        strErrMsg = "Invalid Annual Dollar Limit"
        Exit Function
    End If
    
    If curAmountApplied < 0 Then
        strErrMsg = "Invalid Amount Applied"
        Exit Function
    End If
    
    If curCatastrophicCapAmt < 0 Then
        strErrMsg = "Invalid Catastrophic Cap Amount"
        Exit Function
    End If
        
    If lngAnnualVisitLimit < 0 Then
        strErrMsg = "Annual Visit Limt Not Valid"
        Exit Function
    End If
    
    If lngVisits < 0 Then
        strErrMsg = "Visits Not Valid"
        Exit Function
    End If
    
    If Not ((strCatastrophicCap <> "Y") Or (strCatastrophicCap <> "y") Or _
       (strCatastrophicCap <> "N") Or (strCatastrophicCap <> "n")) Then
        strErrMsg = "Catastrophic Cap Not Valid"
        Exit Function
    End If
    
    If Not ((strCertRequired <> "Y") Or (strCertRequired <> "y") Or _
       (strCertRequired <> "N") Or (strCertRequired <> "n")) Then
        strErrMsg = "Cert Required Not Valid"
        Exit Function
    End If
    
    If Not ((strNoHCFA <> "Y") Or (strNoHCFA <> "y") Or _
       (strNoHCFA <> "N") Or (strNoHCFA <> "n")) Then
        strErrMsg = "No HCFA Not Valid"
        Exit Function
    End If
    
    If Not ((strNoInvoice <> "Y") Or (strNoInvoice <> "y") Or _
       (strNoInvoice <> "N") Or (strNoInvoice <> "n")) Then
        strErrMsg = "No Invoice Not Valid"
        Exit Function
    End If
    
    If Not ((strOutPatientRept <> "Y") Or (strOutPatientRept <> "y") Or _
       (strOutPatientRept <> "N") Or (strOutPatientRept <> "n")) Then
        strErrMsg = "Out Patient Report Not Valid"
        Exit Function
    End If
    
    If Not ((strVerbRept <> "Y") Or (strVerbRept <> "y") Or _
       (strVerbRept <> "N") Or (strVerbRept <> "n")) Then
        strErrMsg = "Verbal Report Not Valid"
        Exit Function
    End If
    
    If Not ((strElectClaim <> "Y") Or (strElectClaim <> "y") Or _
       (strElectClaim <> "N") Or (strElectClaim <> "n")) Then
        strErrMsg = "Elect Claim Not Valid"
        Exit Function
    End If
    
    If Not ((strRecurrCopay <> "Y") Or (strRecurrCopay <> "y") Or _
       (strRecurrCopay <> "N") Or (strRecurrCopay <> "n")) Then
        strErrMsg = "Recurring Copay Not Valid"
        Exit Function
    End If
    
    If lngNumVisit < 0 Then
        strErrMsg = "Number of Visits Not Valid"
        Exit Function
    End If
    
    If lngDaysPrior < 0 Then
        strErrMsg = "Days Prior Not Valid"
        Exit Function
    End If
    
    If Not blnExists Then
        If Exists(lngID) Then
            strErrMsg = "Patient Responsible Party Plan ID already exists"
            Exit Function
        End If
        
    End If
    
    VerifyData = True

End Function

Public Function Exists(ByVal lngID As Long, Optional ByVal strDataBase As String) As Boolean
'--------------------------------------------------------------------
'Date: 12/20/2023
'Author: DCO
'Description:  CHecks for existing entries in tblEligibilityTx.
'Parameters: lngID - EligibilityID number to check for existing entries in table
'
'
'--------------------------------------------------------------------
    Dim obj As BenefitsDB.CEligibilityTxDB
    Dim blnExists As Boolean
    
    On Error GoTo ErrTrap:
    
    Set obj = CreateObject("BenefitsDB.CEligibilityTxDB")
    blnExists = obj.Exists(lngID, strDataBase)
    Set obj = Nothing
    Exists = blnExists
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set obj = Nothing
        
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function
