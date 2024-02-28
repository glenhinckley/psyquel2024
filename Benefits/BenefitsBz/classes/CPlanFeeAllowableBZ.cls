VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanFeeAllowableBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPlanFeeAllowableBZ
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Plan Fee Allowable table.
'--------------------------------------------------------------------
'   R000 - epena - 12/8/00 - fixed object so that is works with CPlanFeeAllowableDB
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CPlanFeeAllowableBZ"

Public Function Insert(ByVal lngPlanID As Long, _
                       ByVal strCPTCode As String, ByVal lngProviderCredential As Long, _
                       ByVal curFeeAllowable As Currency)
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblPlanFeeAllowable table utilizing
'              a stored procedure.
'Parameters: lngPlanID -  Plan ID number to insert into table
'            strCPTCode - The ten digit CPT Code that will be inserted into the table.
'            lngProviderCredential - Numerical representation of providers credentials.
'            curFeeAllowable - Amount paid to provider
'--------------------------------------------------------------------

    Dim objPlanFeeAllowable As BenefitsDB.CPlanFeeAllowableDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objPlanFeeAllowable = CreateObject("BenefitsDB.CPlanFeeAllowableDB")
    If VerifyData(lngPlanID, strCPTCode, lngProviderCredential, curFeeAllowable, strErrMsg, 0) Then
        Call objPlanFeeAllowable.Insert(lngPlanID, strCPTCode, lngProviderCredential, curFeeAllowable)
    Else
        GoTo ErrTrap
    End If
    Set objPlanFeeAllowable = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPlanFeeAllowable = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
End Function
Public Function Update(ByVal lngPlanFeeID As Long, ByVal curFeeAllowable As Currency)   'r000
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Updates a row into the tblPlanCopay table utilizing
'              a stored procedure.
'Parameters: lngPlanFeeID -  Plan Fee ID number to insert into table
'            curFeeAllowable - Amount paid to provider
'            lngPlanID - Plan ID
'            strCPTode - CPT Code
'--------------------------------------------------------------------
    
    Dim objPlanFeeAllowable As BenefitsDB.CPlanFeeAllowableDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    If curFeeAllowable < 0 Then
        strErrMsg = "Fee allowables must be positive amounts"
        GoTo ErrTrap
    End If
    
    Set objPlanFeeAllowable = CreateObject("BenefitsDB.CPlanFeeAllowableDB")
    Call objPlanFeeAllowable.Update(lngPlanFeeID, curFeeAllowable)
    
    Set objPlanFeeAllowable = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPlanFeeAllowable = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
    

End Function


Public Function Exists(ByVal lngPlanID As Long, _
                       ByVal strCPTCode As String, ByVal lngProviderCredential As Long) As Boolean
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Checks for existing entries into tblPlanCopay
'Parameters: lngPlanFeeID - Plan Fee ID number to check for existing entries in table
'
'
'--------------------------------------------------------------------
    Dim objPlanFeeAllowable As BenefitsDB.CPlanFeeAllowableDB
    Dim blnExists As Boolean
    
    On Error GoTo ErrTrap:
    
    Set objPlanFeeAllowable = CreateObject("BenefitsDB.CPlanFeeAllowableDB")
    blnExists = objPlanFeeAllowable.Exists(lngPlanID, strCPTCode, lngProviderCredential)
    Set objPlanFeeAllowable = Nothing
    Exists = blnExists
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPlanFeeAllowable = Nothing
        
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Public Function FetchByPlanCPTCredential(ByVal lngPlanID As Long, ByVal strCPTCode As String, _
                                 ByVal lngProviderCredential As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPlanFeeAllowable table.     '
'Parameters: lngPlanID -  Plan ID number
'            strCPTCode - The ten digit CPT Code.
'            lngProviderCredential - Numerical representation of providers credentials.
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objPlanFeeAllowable As BenefitsDB.CPlanFeeAllowableDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objPlanFeeAllowable = CreateObject("BenefitsDB.CPlanFeeAllowableDB")
    Set rstSQL = objPlanFeeAllowable.FetchByPlanCPTCredential(lngPlanID, strCPTCode, lngProviderCredential)
      
    Set FetchByPlanCPTCredential = rstSQL
    Set objPlanFeeAllowable = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPlanFeeAllowable = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Public Function FetchByPlan(ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblPlanFeeAllowable table.     '
'Parameters: lngPlanID -  Plan ID number
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objPlanFeeAllowable As BenefitsDB.CPlanFeeAllowableDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objPlanFeeAllowable = CreateObject("BenefitsDB.CPlanFeeAllowableDB")
    Set rstSQL = objPlanFeeAllowable.FetchByPlan(lngPlanID)
      
    Set FetchByPlan = rstSQL
    Set objPlanFeeAllowable = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objPlanFeeAllowable = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Private Function VerifyData(ByVal lngPlanID As Long, _
                            ByVal strCPTCode As String, ByVal lngProviderCredential As Long, _
                            ByVal curFeeAllowable As Currency, ByRef strErrMsg As String, ByVal lngPlanFeeID As Long) As Boolean

    Dim objCPT As ListBz.CCPTCodeBz
        
    VerifyData = False
    If lngPlanFeeID = 0 Then
    
        If lngPlanID < 1 Then
            strErrMsg = "Plan RP Plan ID Not Valid"
            Exit Function
        End If
    
        If lngProviderCredential < 1 Then
            strErrMsg = "Provider Credential ID Not Valid"
            Exit Function
        End If
        
        Set objCPT = CreateObject("ListBz.CCPTCodeBz")
        
        If Not objCPT.Exists(strCPTCode) Then
            strErrMsg = "Invalid CPT Code"
            Exit Function
        End If
    
        If curFeeAllowable < 0 Then
            strErrMsg = "Invalid amount"
            Exit Function
        End If
        
        If Exists(lngPlanID, strCPTCode, lngProviderCredential) Then
            strErrMsg = "Provider/CPT Code combination already exists"
            Exit Function
        End If
        VerifyData = True
    Else
        If curFeeAllowable < 0 Then
            strErrMsg = "Invalid amount"
            Exit Function
        End If
        VerifyData = True
    End If

    

End Function

