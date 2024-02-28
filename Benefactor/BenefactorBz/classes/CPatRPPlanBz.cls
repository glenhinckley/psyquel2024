VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatRPPlanBz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatRPPlanBz
'Date: 02/06/2000
'Author: Eric Pena
'Description:  MTS object designed to fetch from the tblPatRPPlan table.
'--------------------------------------------------------------------
'Revision History:                                                  '
' R001 - 11/07/2000 Travis:  Moved FetchPatRPPlanRule to CPatRPPlanRuleBZ               '
' R002 - 11/07/2000 Travis:  Moved original Update() to CPatRPPlanRuleBZ
' R003 - 06/15/2000 Richkun: Removed IsEnabled() check from SetOrder.
'           The Pat-RP-Plan record will we re-enabled if it is found
'           to be disabled.
' R004 - 07/24/2000 Richkun: Added Append() method.
' R005 - 03/22/2001 Nichol: Added ExistsPatRPPlan()
' R006 - 05/16/2001 Richkun: Added FetchDisabledRPPlansByPatient()
' R007 - 01/24/2002 Richkun: Added FetchPatientRPNames()
'--------------------------------------------------------------------
Option Explicit

Private Const TABLE_NAME As String = "tblPatRPPlan"
Private Const CLASS_NAME As String = "CPatRPPlanDB"

Public Function FetchPlans(ByVal lngPatientID As Long, ByVal lngRPID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Returns a variant 2d-array containing a list of all Plans and levels to which a Patient-RP combo is associated.
'Parameters: lngPatientID - The ID of the Patient
'                    lngRPID - - The ID of the RP
'Returns: An ADODB.Recordset containing a list of all plans and levels  to which a Patient-RP combo is associated.
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objPatRPPlan As BenefactorDB.CPatRPPlanDB
    
    Set objPatRPPlan = CreateObject("BenefactorDB.CPatRPPlanDB")
    Set FetchPlans = objPatRPPlan.FetchPlans(lngPatientID, lngRPID)
    Set objPatRPPlan = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objPatRPPlan = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IsEnabled(ByVal lngPatientID As Long, ByVal lngRPID As Long, ByVal lngPlanID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Lets the caller know if a given Patient-RP-Plan combo association is active
'Parameters: lngPatientID - The ID of the Patient
'                  lngRPID - The ID of the RP
'                  lngPlanID - The ID of the Plan
'Returns: T if the association is enabled/exists, F otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objPatRPPlan As BenefactorDB.CPatRPPlanDB
    
    Set objPatRPPlan = CreateObject("BenefactorDB.CPatRPPlanDB")
    If objPatRPPlan.IsEnabled(lngPatientID, lngRPID, lngPlanID) > 0 Then
        IsEnabled = True
    Else
        IsEnabled = False
    End If
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objPatRPPlan = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objPatRPPlan = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function Append(ByVal lngPatientID As Long, ByVal lngRPID As Long, ByVal lngPlanID As Long, _
                       ByVal strAddedBy As String) As Long
'--------------------------------------------------------------------
'Date: 07/24/2000
'Author: Dave Richkun
'Description:  Creates a new Patient-RP-Plan combination.  The Order value
'              is automatically assigned the next available number based on
'              existing plan assignments.
'Parameters: lngPatientID - The ID of the Patient
'            lngRPID - The ID of the RP
'            lngPlanID - The ID of the Plan
'Returns: ID of the new row inserted into tblPatRPPlan
'--------------------------------------------------------------------
'Revision History:
'  R004 - Created
'--------------------------------------------------------------------

    
    Dim objPatRPPlan As BenefactorDB.CPatRPPlanDB
    
    
    Set objPatRPPlan = CreateObject("BenefactorDB.CPatRPPlanDB")
    
    Append = objPatRPPlan.Append(lngPatientID, lngRPID, lngPlanID, strAddedBy)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objPatRPPlan = Nothing
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objPatRPPlan = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function Insert(ByVal lngPatientID As Long, ByVal lngRPID As Long, ByVal lngPlanID As Long, ByVal intOrder As Integer, ByVal strAddedBy As String) As Long
'--------------------------------------------------------------------
'Date: 03/01/2000
'Author: Eric Pena
'Description:  Creates a billing association for the given patient-rp-plan combonation
'Parameters: lngPatientID - The ID of the Patient
'                  lngRPID - The ID of the RP
'                  lngPlanID - The ID of the Plan
'Returns: T if the association is enabled/exists, F otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objPatRPPlan As BenefactorDB.CPatRPPlanDB
    
    Set objPatRPPlan = CreateObject("BenefactorDB.CPatRPPlanDB")
    Insert = objPatRPPlan.Insert(lngPatientID, lngRPID, lngPlanID, intOrder, strAddedBy)
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objPatRPPlan = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objPatRPPlan = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngPatientID As Long, ByVal lngRPID As Long, ByVal lngPlanID As Long, ByVal strDeletedBy As String)
'--------------------------------------------------------------------
'Date: 03/01/2000
'Author: Eric Pena
'Description:  Flags a row in the tblBenefactor table marking the row as
'              deleted or undeleted.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngID - ID of the row in the table whose value will be
'               updated.
'            strUserName - Login name of the user responsible for
'               marking the row as deleted.
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    
    'TODO: Check for active claims
    Call objBenefactor.Deleted(blnDeleted, lngPatientID, lngRPID, lngPlanID, strDeletedBy)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub DeleteRP(ByVal lngPatientID As Long, ByVal lngRPID As Long, _
                    ByVal strDeletedBy As String)
'--------------------------------------------------------------------
'Date: 07/31/2000
'Author: Dave Richkun
'Description:  Marks every Plan record associated with a Patient's Responsible
'              Party as disabled.
'Parameters:  lngPatientID - The ID of the Patient
'             lngRPID - The ID of the RP
'             strUserName - Login name of the user responsible for
'               marking the rows as deleted.
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    
    Dim objPRPP As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    
    Set objPRPP = CreateObject("BenefactorDB.CPatRPPlanDB")
    
    Call objPRPP.DeleteRP(lngPatientID, lngRPID, strDeletedBy)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objPRPP = Nothing
    
    Exit Sub
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objPRPP = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub


Public Function SetOrder(ByVal lngPatientID As Long, ByVal lngRPID As Long, _
                    ByVal lngPlanID As Long, ByVal intOrder As Integer) As Boolean
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Changes the order of the given Patient-RP-Plan combo association above
'Parameters: lngPatientID - The ID of the Patient
'                  lngRPID - The ID of the RP
'                  lngPlanID - The ID of the Plan
'                  intOrder - The order to assign to the association
'Returns: True if the record was succesfully updated, otherwise False.
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objPatRPPlan As BenefactorDB.CPatRPPlanDB
    
    
    Set objPatRPPlan = CreateObject("BenefactorDB.CPatRPPlanDB")
'    Removed R003
'    If Not IsEnabled(lngPatientID, lngRPID, lngPlanID) Then
'        getobjectcontext.SetAbort
'        GoTo ErrTrap
'    End If
    SetOrder = objPatRPPlan.SetOrder(lngPatientID, lngRPID, lngPlanID, intOrder)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objPatRPPlan = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objPatRPPlan = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function FetchRPsByPat(ByVal intPatient As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   intProvider - provider to match
'                   strLast - last name to match
'                   strFirst - first name to match
'Returns: ADORst
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    Set FetchRPsByPat = objBenefactor.FetchRPsByPat(intPatient, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function FetchPlansByRP(ByVal intRPID As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns rows from the tblBenefactor table utilizing a stored procedure
'Parameters: blnIncludeDisabled - determines if disabled records will be returned
'                   intProvider - provider to match
'                   strLast - last name to match
'                   strFirst - first name to match
'Returns: ADORst
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    Set FetchPlansByRP = objBenefactor.FetchPlansByRP(intRPID, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FetchRPPlansByPat(ByVal lngPatientID As Long, Optional blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 3/08/2000
'Author: Eric Pena
'Description:  Returns plan information for a specific patient
'Parameters: lngPatientID - ID of patient whose plans will be retrieved
'            blnIncludeDisabled - determines if disabled plans will be included in
'                   the resultset
'Returns: Recordset of patient plans
'--------------------------------------------------------------------
    
    Dim objBFact As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    Set objBFact = CreateObject("BenefactorDB.CPatRPPlanDB")
    Set FetchRPPlansByPat = objBFact.FetchRPPlansByPat(lngPatientID, blnIncludeDisabled)
    Set objBFact = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBFact = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function FetchDisabledRPPlansByPatient(ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/16/2001
'Author: Dave Richkun
'Description:  Returns disabled Patient-RP-Plan rows from the tblPatRPPlan table
'              utilizing a stored procedure
'Parameters: lngPatientID - ID of the patient whose disabled palns will be retrieved
'Returns: ADO Recordset
'--------------------------------------------------------------------
'  R006 - Created
'--------------------------------------------------------------------

    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    Set FetchDisabledRPPlansByPatient = objBenefactor.FetchDisabledRPPlansByPatient(lngPatientID)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



Public Function ExistsPatRPPlan(ByVal lngPatientID As Long, ByVal lngRPID As Long, ByVal lngPlanID As Long, _
                            Optional ByVal blnIncludeDisabled As Boolean = False) As Long
'--------------------------------------------------------------------
'Date: 3/21/2001
'Author: David Nichol
'Description:  If an patient/RP/plan combination already exists, returns its ID in tblPatientRPPlan.
'               Otherwise it returns 0.
'Parameters: lngPatientID - patient ID
'            lngRPID - responsible party ID
'            lngPlanID - plan ID
'            blnIncludeDisabled - Y/N as to whether to include disabled records, defaulting to false
'Returns: Long - ID of the existing record, or 0
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    ExistsPatRPPlan = objBenefactor.ExistsPatRPPlan(lngPatientID, lngRPID, lngPlanID, blnIncludeDisabled)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'++++++++++++                R001         +++++++++++++++++++++'

'---------------------- Begin R005 -----------------------------
Public Function ReorderPlan(ByVal lngPatientID As Long, ByVal intOldPriority As Integer, _
                            ByVal intNewPriority As Integer) As Boolean
'--------------------------------------------------------------------
'Date: 01/12/2001
'Author: Dave Nichol
'Description:  Calls DB method that changes the order of the a Patient-RP-Plan
'   combo association, then adjusts the order of the other pat-rpplan
'   records as needed.
'Parameters: lngPatientID - The ID of the Patient
'            lngOldPriority - The current order # of the plan
'            lngNewPriority - The new order # of the plan
'Returns: True if the record was updated, False otherwise
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------

    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    Dim strErrMessage As String
    
    On Error GoTo ErrTrap:
    
    'verify parameter
    If Not lngPatientID > 0 Then
        strErrMessage = "Patient ID is not valid."
        GoTo ErrTrap
    End If
    If Not intOldPriority > 0 Then
        strErrMessage = "Old priority is not valid."
        GoTo ErrTrap
    End If
    If Not intNewPriority > 0 Then
        strErrMessage = "New priority is not valid."
        GoTo ErrTrap
    End If

    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    ReorderPlan = objBenefactor.ReorderPlan(lngPatientID, intOldPriority, intNewPriority)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function

'---------------------- End R005 -----------------------------

Public Function FetchPatientRPNames(ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 01/24/2002
'Author: Dave Richkun
'Description:  Retrieves list of Responsible Party names associated with a patient
'Parameters: lngPatientID - ID of the Patient
'Returns: Recordset of RP names
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatRPPlanDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatRPPlanDB")
    Set FetchPatientRPNames = objBenefactor.FetchPatientRPNames(lngPatientID)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBenefactor = Nothing
    
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



