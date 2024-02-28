VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatientBenefitBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatientBenefitBZ
'Date: 11/02/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Patient Benefits Verification table.
'--------------------------------------------------------------------
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CPatientBenefitBZ"


Public Function Insert(ByVal lngPatRPPlanID As Long, ByVal strCardNumber As String, _
                       ByVal strGroupNumber As String, ByVal strCertRequired As String, _
                       ByVal lngPriCopayID As Long, ByVal lngSecCopayID As Long, ByVal lngTerCopayID As Long, _
                       ByVal dblPriCopayAmt As Double, ByVal dblSecCopayAmt As Double, ByVal dblTerCopayAmt As Double, _
                       ByVal strAdditionalInfo As String, ByVal strRPRelat As String, ByVal lngVerPlanID As Long, ByVal strPlanName As String, _
                       ByVal strPlanAddress As String, ByVal strPlanCity As String, ByVal strPlanState As String, _
                       ByVal strPlanZip As String, ByVal strInsPhone As String, ByVal strInsdName As String, _
                       ByVal strPayerCode As String, ByVal strVerifyElectronic As String, _
                       ByVal strOutOfNetwork As String, ByVal strContracted As String, ByVal dblDeduct As Double, _
                       ByVal dblDeductRem As Double, ByVal strMSPCode As String, ByVal strUserName As String) As Long
'--------------------------------------------------------------------
'Date: 02/16/2001
'Author: Dave Richkun
'Description:  Inserts a row into the tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters: All data values to be inserted into the table
'Returns: ID (Primary Key) of the inserted row.
'--------------------------------------------------------------------

    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim strErrMsg As String
    Dim lngID As Long
    
    On Error GoTo ErrTrap:
    
    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    Insert = objPatBen.Insert(lngPatRPPlanID, strCardNumber, strGroupNumber, _
                strCertRequired, lngPriCopayID, lngSecCopayID, lngTerCopayID, _
                dblPriCopayAmt, dblSecCopayAmt, dblTerCopayAmt, strAdditionalInfo, strRPRelat, _
                lngVerPlanID, strPlanName, strPlanAddress, strPlanCity, strPlanState, strPlanZip, _
                strInsPhone, strInsdName, strPayerCode, strVerifyElectronic, strOutOfNetwork, strContracted, dblDeduct, dblDeductRem, strMSPCode, strUserName)
    Set objPatBen = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBen = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function

Public Function Update(ByVal lngPatRPPlanID As Long, ByVal strRPRelat As String, ByVal strCardNumber As String, _
                       ByVal strGroupNumber As String, ByVal strCertRequired As String, ByVal strPayerCode As String, _
                       ByVal lngPriCopayID As Long, ByVal lngSecCopayID As Long, ByVal lngTerCopayID As Long, _
                       ByVal dblPriCopayAmt As Double, ByVal dblSecCopayAmt As Double, ByVal dblTerCopayAmt As Double, _
                       ByVal strVerifyElectronic As String, _
                       ByVal strReVerifyElectronic As String, ByVal strOutOfNetwork As String, ByVal strContracted As String, _
                       ByVal strBenefitsNotes As String, _
                       ByVal dblDeduct As Double, ByVal dblDeductRem As Double, ByVal strMSPCode As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 02/15/2001
'Author: Dave Richkun
'Description:  Updates a row into tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim strErrMsg As String

    On Error GoTo ErrTrap:

    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    Call objPatBen.Update(lngPatRPPlanID, strRPRelat, strCardNumber, strGroupNumber, _
                      strCertRequired, strPayerCode, lngPriCopayID, lngSecCopayID, lngTerCopayID, _
                      dblPriCopayAmt, dblSecCopayAmt, dblTerCopayAmt, _
                      strVerifyElectronic, strReVerifyElectronic, strOutOfNetwork, strContracted, strBenefitsNotes, dblDeduct, dblDeductRem, strMSPCode, strUserName)
    Set objPatBen = Nothing

    GetObjectContext.SetComplete

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBen = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function
Public Function UpdateVerify(ByVal lngPatRPPlanID As Long, ByVal strRPRelat As String, ByVal strCardNumber As String, _
                       ByVal strGroupNumber As String, ByVal strCertRequired As String, ByVal strPayerCode As String, _
                       ByVal lngPriCopayID As Long, ByVal lngSecCopayID As Long, ByVal lngTerCopayID As Long, _
                       ByVal dblPriCopayAmt As Double, ByVal dblSecCopayAmt As Double, ByVal dblTerCopayAmt As Double, _
                       ByVal strVerifyText As String, _
                       ByVal strVerifyRejectYN As String, ByVal strOutOfNetwork As String, ByVal strContracted As String, _
                       ByVal dblDeduct As Double, ByVal dblDeductRem As Double, ByVal strMSPCode As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 02/15/2001
'Author: Dave Richkun
'Description:  Updates a row into tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim strErrMsg As String

    On Error GoTo ErrTrap:

    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    Call objPatBen.UpdateVerify(lngPatRPPlanID, strRPRelat, strCardNumber, strGroupNumber, _
                      strCertRequired, strPayerCode, lngPriCopayID, lngSecCopayID, lngTerCopayID, _
                      dblPriCopayAmt, dblSecCopayAmt, dblTerCopayAmt, strVerifyText, _
                      strVerifyRejectYN, strOutOfNetwork, strContracted, dblDeduct, dblDeductRem, strMSPCode, strUserName)
    Set objPatBen = Nothing

    GetObjectContext.SetComplete

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBen = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function
Public Function UpdateFileID(ByVal lngPatRPPlanID As Long, ByVal lngPlanID As Long, ByVal strFileID As String)
'--------------------------------------------------------------------
'Date: 02/15/2001
'Author: Dave Richkun
'Description:  Updates a row in the tblPatientRPPlanRule table utilizing
'              a stored procedure.
'Parameters:
'--------------------------------------------------------------------

    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim strErrMsg As String

    On Error GoTo ErrTrap:

    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    Call objPatBen.UpdateFileID(lngPatRPPlanID, lngPlanID, strFileID)
    Set objPatBen = Nothing

    GetObjectContext.SetComplete

    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBen = Nothing

    If Err.Number = 0 Then
        Err.Raise vbObjectError + 512 + 16, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If

End Function
Public Function FetchByPatRPPlan(ByVal lngPatRPPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/26/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPatientBenefitVerfication table.     '
'Parameters: lngPatRPPlanID - Patient RP Plan ID number to fetch from table '
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    
    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim strErrMsg As String
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    Set rstSQL = objPatBen.FetchByPatRPPlan(lngPatRPPlanID)
      
    Set FetchByPatRPPlan = rstSQL
    Set objPatBen = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBen = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function


Private Function VerifyData(ByVal lngPatRPPlanID As Long, _
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
    If lngPatRPPlanID < 0 Then
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
        If Exists(lngPatRPPlanID) Then
            strErrMsg = "Patient Responsible Party Plan ID already exists"
            Exit Function
        End If
        
    End If
    
    VerifyData = True

End Function

Public Function Exists(ByVal lngPatRPPlanID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 11/02/2000
'Author: Morgan Travis
'Description:  CHecks for existing entries in tblPatientBenefit.
'Parameters: lngPatRPPlanID - PatRPPlanID number to check for existing entries in table
'
'
'--------------------------------------------------------------------
    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim blnExists As Boolean
    
    On Error GoTo ErrTrap:
    
    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    blnExists = objPatBen.Exists(lngPatRPPlanID)
    Set objPatBen = Nothing
    Exists = blnExists
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPatBen = Nothing
        
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

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

    Dim objPatBen As BenefitsDB.CPatientBenefitDB
    Dim strMessage As String

    On Error GoTo ErrTrap:

    Set objPatBen = CreateObject("BenefitsDB.CPatientBenefitDB")
    
    VerifyPolicyNum = objPatBen.VerifyPolicyNum(intPlanId, intInsID, strGroupNum, strPolicyNum)
    
    'Signal successful completion
    GetObjectContext.SetComplete
        
    Set objPatBen = Nothing
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise Err.Number, CLASS_NAME, strMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

