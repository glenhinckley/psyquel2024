VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProvFeeBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CProvFeeBZ
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  MTS object designed to host methods associated with
'               data affecting Provider Fee table.
'--------------------------------------------------------------------
'Revision History:
'  R001: 05/08/2001 Richkun: Added FetchByID() method
'  R002: 05/08/2001 Richkun: Altered Update() method to include CPTCode, ProviderID parameter
'  R003: 05/08/2001 Richkun: Added Delete() method
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CProvFeeBZ"


Public Function Insert(ByVal lngProviderID As Long, _
                       ByVal strCPTCode As String, ByVal strCPTModifier As String, _
                       ByVal strCPTDescription As String, ByVal curFee As Currency, _
                       ByVal curContractAmount As Currency, ByVal blnPatSched As Boolean, _
                       ByVal lngDuration As Long, ByVal lngPlanID As Long, _
                       ByVal strNotes As String, ByVal strAddedBy As String)
'--------------------------------------------------------------------
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderID - Provider ID number to insert into table
'            strCPTCode - The ten digit CPT Code that will be inserted into the table.
'            dblFee - The monetary amount that will be inserted into the table.
'
'--------------------------------------------------------------------

    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    If VerifyData(lngProviderID, strCPTCode, curFee, curContractAmount, strErrMsg, lngPlanID, 0) Then
        Call objProvFee.Insert(lngProviderID, strCPTCode, strCPTModifier, _
                        strCPTDescription, curFee, curContractAmount, blnPatSched, lngDuration, _
                        lngPlanID, strNotes, strAddedBy)
    Else
        GoTo ErrTrap
    End If
    Set objProvFee = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
End Function

Public Function Update(ByVal lngProviderFeeID As Long, ByVal lngProviderID As Long, _
                       ByVal strCPTCode As String, ByVal strCPTModifier As String, _
                       ByVal strCPTDescription As String, ByVal curFee As Currency, _
                       ByVal curContractAmount As Currency, ByVal blnPatSched As Boolean, _
                       ByVal lngDuration As Long, ByVal lngPlanID As Long, _
                       ByVal strNotes As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  Updates a row in the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderFeeID - Provider Fee ID number to update table
'            lngProviderID - ID of the Provider whose fee is being updated
'            strCPTCode - CPT Code
'            curFee - The monetary amount that will be inserted into the table.
'            strUserName - User name identifying user making the change
'--------------------------------------------------------------------
    
    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    
    If curFee < 0 Then
        strErrMsg = "Fees must be positive amounts."
        GoTo ErrTrap
    End If
    
    If Exists(lngProviderID, strCPTCode, lngPlanID, lngProviderFeeID) Then
        strErrMsg = "Update would produce duplicate fee."
        GoTo ErrTrap
    End If
     
    Call objProvFee.Update(lngProviderFeeID, lngProviderID, strCPTCode, _
                strCPTModifier, strCPTDescription, curFee, curContractAmount, _
                blnPatSched, lngDuration, lngPlanID, strNotes, strUserName)
    
    Set objProvFee = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
    

End Function

Public Function Exists(ByVal lngProviderID As Long, ByVal strCPTCode As String, _
                       ByVal lngPlanID As Long, Optional ByVal lngProviderFeeID As Long)
'--------------------------------------------------------------------
'Date: 10/25/2000
'Author: Morgan Travis
'Description:  Inserts a row into the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderID - Provider ID number to check for existing entries in table
'            strCPTCode - CPT Code
'--------------------------------------------------------------------
    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim blnExists As Boolean
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    blnExists = objProvFee.Exists(lngProviderID, strCPTCode, lngPlanID, lngProviderFeeID)
    Set objProvFee = Nothing
    Exists = blnExists
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
        
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function

Public Function FetchByProvider(ByVal lngProviderID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/25/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderFee table.     '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    Set rstSQL = objProvFee.FetchByProvider(lngProviderID)
      
    Set FetchByProvider = rstSQL
    Set objProvFee = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
    

End Function
Public Function FetchByPlan(ByVal lngProviderID As Long, _
                            ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/25/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderFee table.     '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'
'Returns: Recordset of requested providers
'-------------------------------------------------------------------'

    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    Set rstSQL = objProvFee.FetchByPlan(lngProviderID, lngPlanID)
      
    Set FetchByPlan = rstSQL
    Set objProvFee = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description
End Function
Public Function FetchByProviderCPTCode(ByVal lngProviderID As Long, _
                                       ByVal strCPTCode As String, _
                                       ByVal lngPlanID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 10/25/2000                                                   '
'Author: Morgan Travis                                              '
'Description:  Retrieves records from the tblProviderFee table.     '
'Parameters: lngProviderID - Provider ID number to fetch from table '
'            strCPTCode - CPT Code to fetch from table
'Returns: RecordSet with requested fee
'-------------------------------------------------------------------'

    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    Set rstSQL = objProvFee.FetchByProviderCPTCode(lngProviderID, strCPTCode, lngPlanID)
      
    Set FetchByProviderCPTCode = rstSQL
    Set objProvFee = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function

Private Function VerifyData(ByVal lngProviderID As Long, ByVal strCPTCode As String, _
                            ByVal curFee As Currency, ByVal curContractAmount As Currency, _
                            ByRef strErrMsg As String, _
                            ByVal lngPlanID As Long, ByVal lngProviderFeeID As Long) As Boolean

    Dim objCPT As ListBz.CCPTCodeBz
    
    VerifyData = False
    If lngProviderFeeID = 0 Then
    
        If lngProviderID < 1 Then
            strErrMsg = "Provider ID Not Valid"
            Exit Function
        End If
    
        If curFee < 0 Then
            strErrMsg = "Fee Amount Not Valid"
            Exit Function
        End If
    
        If curContractAmount < 0 Then
            strErrMsg = "Contract Amount Not Valid"
            Exit Function
        End If
        
        Set objCPT = CreateObject("ListBz.CCPTCodeBz")
    
        If Not objCPT.Exists(strCPTCode) Then
            strErrMsg = "Invalid CPT Code"
            Exit Function
        End If
        
        If Exists(lngProviderID, strCPTCode, lngPlanID, 0) Then
                strErrMsg = "Provider/CPT Code combination already exists"
                Exit Function
        End If
        
        VerifyData = True
    Else
        If curFee < 0 Then
            strErrMsg = "Fee Amount Not Valid"
            Exit Function
        End If
        
        If curContractAmount < 0 Then
            strErrMsg = "Contract Amount Not Valid"
            Exit Function
        End If
        
        VerifyData = True
    End If
End Function

Public Function FetchByID(ByVal lngProviderFeeID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/08/2001
'Author: Dave Richkun
'Description:  Retrieves a record from the tblProviderFee table.
'Parameters: lngProviderFeeID - Value of the Primary Key identifiying the
'               row that will be retrieved.
'Returns: ADO Recordset
'-------------------------------------------------------------------'
' R001 - Created
'-------------------------------------------------------------------'

    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim rstSQL As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    Set rstSQL = objProvFee.FetchByID(lngProviderFeeID)
      
    Set FetchByID = rstSQL
    Set objProvFee = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, CLASS_NAME, Err.Description

End Function

Public Sub Delete(ByVal lngProviderFeeID As Long)
'--------------------------------------------------------------------
'Date: 05/08/2001
'Author: Dave Richkun
'Description:  Deletes a row from the tblProviderFee table utilizing
'              a stored procedure.
'Parameters: lngProviderFeeID - ID of row to delete
'--------------------------------------------------------------------
'  R003 - Created
'--------------------------------------------------------------------
    
    Dim objProvFee As BenefitsDB.CProvFeeDB
    Dim strErrMsg As String
    
    On Error GoTo ErrTrap:
    
    Set objProvFee = CreateObject("BenefitsDB.CProvFeeDB")
    Call objProvFee.Delete(lngProviderFeeID)
    Set objProvFee = Nothing
    
    GetObjectContext.SetComplete
    
    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    On Error GoTo 0
    GetObjectContext.SetAbort
    Set objProvFee = Nothing
    
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, CLASS_NAME, Err.Description
    End If
    
End Sub

