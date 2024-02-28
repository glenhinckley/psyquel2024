VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProviderPanelBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'--------------------------------------------------------------------
'Class Name: CPlanCopayBZ
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Plan Copay table.
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CProviderPanelBZ"

Public Function Insert(ByVal lngPlanID As Long, ByVal lngProviderID As Long)
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblProviderPlan table utilizing
'              a stored procedure.
'Parameters: lngPlanID -  Plan ID number to insert into table
'            lngProviderID - Provider ID number.
'
'--------------------------------------------------------------------

    Dim objProviderPlan As BenefitsDB.CProviderPanelDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objProviderPlan = CreateObject("BenefitsDB.CProviderPanelDB")
    If VerifyData(lngPlanID, lngProviderID, strErrMsg) Then
        Call objProviderPlan.Insert(lngPlanID, lngProviderID)
    Else
        GoTo ErrTrap
    End If
    Set objProviderPlan = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProviderPlan = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
End Function


Public Function Delete(ByVal lngProviderPanelID As Long, ByVal lngProviderID As Long, _
                       ByVal lngPlanID As Long)
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Delete a row into the tblProviderPlan table utilizing
'              a stored procedure.
'Parameters: lngProviderPanelID -  Provider Panel ID number to delete from table
'
'
'--------------------------------------------------------------------
    
    Dim objProviderPlan As BenefitsDB.CProviderPanelDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objProviderPlan = CreateObject("BenefitsDB.CProviderPanelDB")
    If Exists(lngProviderID, lngPlanID) Then
        Call objProviderPlan.Delete(lngProviderPanelID)
    Else
        GoTo ErrTrap
    End If
    Set objProviderPlan = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProviderPlan = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
    

End Function


Public Function Exists(ByVal lngProviderID As Long, ByVal lngPlanID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 10/31/2000
'Author: Morgan Travis
'Description:  Checks for existing entries into tblPlanCopay
'Parameters: lngProviderPanelID - Provider Panel ID Number.
'
'--------------------------------------------------------------------
    Dim objProviderPlan As BenefitsDB.CProviderPanelDB
    Dim blnExists As Boolean
    
    On Error GoTo ErrTrap:
    
    Set objProviderPlan = CreateObject("BenefitsDB.CProviderPanelDB")
    blnExists = objProviderPlan.Exists(lngProviderID, lngPlanID)
    Set objProviderPlan = Nothing
    Exists = blnExists
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProviderPlan = Nothing
        
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Public Function FetchByPlan(ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderPanel table.   '
'Parameters: lngPlanID -  Plan ID number to match
'
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objProviderPlan As BenefitsDB.CProviderPanelDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProviderPlan = CreateObject("BenefitsDB.CProviderPanelDB")
    Set rstSQL = objProviderPlan.FetchByPlan(lngPlanID)
      
    Set FetchByPlan = rstSQL
    Set objProviderPlan = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProviderPlan = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Public Function FetchByProvider(ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderPanel table.   '
'Parameters: lngProviderID -  Provider ID number to match
'
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objProviderPlan As BenefitsDB.CProviderPanelDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProviderPlan = CreateObject("BenefitsDB.CProviderPanelDB")
    Set rstSQL = objProviderPlan.FetchByProvider(lngProviderID)
      
    Set FetchByProvider = rstSQL
    Set objProviderPlan = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProviderPlan = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Public Function FetchByPlanProvider(ByVal lngPlanID As Long, ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/31/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderPanel table.   '
'Parameters: lngPlanID -  Plan ID number to match
'
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objProviderPlan As BenefitsDB.CProviderPanelDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProviderPlan = CreateObject("BenefitsDB.CProviderPanelDB")
    Set rstSQL = objProviderPlan.FetchByPlanProvider(lngPlanID, lngProviderID)
      
    Set FetchByPlanProvider = rstSQL
    Set objProviderPlan = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProviderPlan = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function


Private Function VerifyData(ByVal lngPlanID As Long, ByVal lngProviderID As Long, _
                            ByRef strErrMsg As String) As Boolean

    VerifyData = False
        
    If lngPlanID < 1 Then
         strErrMsg = "Plan ID Not Valid"
         Exit Function
    End If
    
    If lngProviderID < 1 Then
        strErrMsg = "Provider ID Not Valid"
        Exit Function
    End If
    
    If Exists(lngProviderID, lngPlanID) Then
       strErrMsg = "Provider/Plan combination already exists"
       Exit Function
    End If
    
    VerifyData = True

End Function



