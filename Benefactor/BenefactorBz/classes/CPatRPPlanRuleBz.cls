VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatRPPlanRuleBz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatRPPlanRuleBz
'Date: 11/07/2000
'Author: Morgan Travis
'Description:  MTS object designed to fetch from the tblPatRPPlanRule table.
'--------------------------------------------------------------------
'Revision History:                                                  '
'
'--------------------------------------------------------------------
Option Explicit

Private Const TABLE_NAME As String = "tblPatRPPlanRule"


Public Function Update(ByVal lngPatRPPlanID As Long, ByVal strRPRelat As String, ByVal strCardNum As String, ByVal strGroupNum As String, _
                                    ByVal strMedicareHICNum As String, ByVal curAmtOfGrpDeductPd As Currency, ByVal dteFiscalEndDate As Date, _
                                    ByVal curDeduct As Currency, ByVal strGroupDeduct As String, ByVal curDeductRem As Currency, _
                                    ByVal curDollarLimit As Currency, ByVal curDollarsApplied As Currency, ByVal intVisitLimit As Integer, _
                                    ByVal intVisitsApplied As Integer, ByVal strCatCap As String, ByVal curCatCapAmt As Currency, _
                                    ByVal strCertReqd As String, ByVal strNoHCFA As String, ByVal strNoInvoice As String, ByVal strOutpatientRpt As String, _
                                    ByVal strVerbalRept As String, ByVal intNumVisit As Integer, ByVal intDaysPrior As Integer, ByVal strElectClaim As String, ByVal strRecurrCopay As String)
'---------------------------------------------------------------------
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

    Dim objPatRPPlan As BenefactorDB.CPatRPPlanRuleDB
    

    Set objPatRPPlan = CreateObject("BenefactorDB.CPatRPPlanRuleDB")
    Call objPatRPPlan.Update(lngPatRPPlanID, strRPRelat, strCardNum, strGroupNum, strMedicareHICNum, curAmtOfGrpDeductPd, dteFiscalEndDate, _
                                     curDeduct, strGroupDeduct, curDeductRem, curDollarLimit, curDollarsApplied, intVisitLimit, intVisitsApplied, strCatCap, _
                                     curCatCapAmt, strCertReqd, strNoHCFA, strNoInvoice, strOutpatientRpt, strVerbalRept, intNumVisit, intDaysPrior, strElectClaim, strRecurrCopay)
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objPatRPPlan = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
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
    
    Dim objBenefactor As BenefactorDB.CPatRPPlanRuleDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanRuleDB")
    Set FetchPatRPPlanRule = objBenefactor.FetchPatRPPlanRule(lngPatID, lngRPID, lngPlanID)
    
    ' Signal successful completion
    GetObjectContext.SetComplete
    
    Set objBenefactor = Nothing
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FetchPatRPPlanRuleByPatRPPlanID(ByVal lngPatRPPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/10/2001                                                   '
'Author: Morgan Travis                                              '
'Description:  Returns a row from tblPatRPPlanRule with the         '
'              matching parameters                                  '
'Parameters: lngPatRPPlanID                                         '
'Returns: ADODB Recordset if match found, otherwise Null            '
'--------------------------------------------------------------------

    
    Dim objBenefactor As BenefactorDB.CPatRPPlanRuleDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanRuleDB")
    Set FetchPatRPPlanRuleByPatRPPlanID = objBenefactor.FetchPatRPPlanRuleByPatRPPlanID(lngPatRPPlanID)
    
    ' Signal successful completion
    GetObjectContext.SetComplete
    
    Set objBenefactor = Nothing
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function UpdatePatientCounters(ByVal lngPatRPPlanID As Long, ByVal curAmtApplied As Currency)
'--------------------------------------------------------------------
'Date: 01/10/2001
'Author: Morgan Travis
'Description:  Updates Patients visit,amount applied, and deductible counters
'Parameters:
'Returns:
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    Dim lngCounter As Long
    Dim objBene As BenefactorDB.CPatRPPlanRuleDB
    Dim objPatRPPlan As BenefactorBz.CPatRPPlanRuleBz
    Dim rst As ADODB.Recordset
    Dim VisitCounter As Integer
    Dim AmountApplied As Long
    Dim Deductible As Long
    
    On Error GoTo ErrTrap:
    
    
    Set objPatRPPlan = CreateObject("BenefactorBz.CPatRPPlanRuleBz")
    Set rst = objPatRPPlan.FetchPatRPPlanRuleByPatRPPlanID(lngPatRPPlanID)
    
    If curAmtApplied > 0 Then
    
        VisitCounter = rst.Fields("fldVisitsApplied").Value + 1
        AmountApplied = rst.Fields("fldDollarsApplied").Value + curAmtApplied
        Deductible = rst.Fields("fldDeductRem").Value - curAmtApplied
    
    End If
    
    If curAmtApplied < 0 Then
    
        VisitCounter = rst.Fields("fldVisitsApplied").Value - 1
        AmountApplied = rst.Fields("fldDollarsApplied").Value + curAmtApplied
        Deductible = rst.Fields("fldDeductRem").Value - curAmtApplied
    
    End If
    
    Set objBene = CreateObject("BenefactorDB.CPatRPPlanRuleDB")
    Call objBene.UpdatePatientCounters(lngPatRPPlanID, VisitCounter, AmountApplied, Deductible)
    
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBene = Nothing
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objBene = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function



