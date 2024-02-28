VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBenefitsBz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CInsuranceBz
'Date: 09/11/2000
'Author: Eric Pena
'Description:  MTS business object designed to call methods associated
'              with the CInsuranceDB class.
'--------------------------------------------------------------------
' Revision History:
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CBenefitsBz"
Private Const TABLE_NAME As String = ""

Public Function FetchBenefitsByPatient(ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 09/11/2000
'Author: Eric Pena
'Description:  Retrieves records from the tblInsCompany table utilizing a
'            stored procedure.
'Parameters: blnIncludeDisabled - Optional parameter that identifies if
'               records flagged as 'Disabled' or 'De-activated' are to be
'               included in the record set. The default value is False.
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    Dim rstSQL As ADODB.Recordset
    Dim objBene As BenefitsDB.CBenefitsDB
    
    On Error GoTo ErrTrap:

    Set objBene = CreateObject("BenefitsDB.CBenefitsDB")
    Set rstSQL = objBene.FetchBenefitsByPatient(lngPatientID)
    Set FetchBenefitsByPatient = rstSQL
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBene = Nothing
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objBene = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function FetchCoPayByPatientProvider(ByVal lngPatientID As Long, _
       ByVal lngProviderID As Long) As Double
'--------------------------------------------------------------------
'Date: 11/07/2001
'Author: Dave Richkun
'Description:  Returns the CoPay amount for a patient's Primary plan
'              based on the credentials of a Provider
'Parameters: lngPatientID - ID of the patient
'            lngProviderID - ID of the Provider
'Returns: Copay amount
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    Dim objBnft As BenefitsDB.CBenefitsDB
    
    On Error GoTo ErrTrap:

    Set objBnft = CreateObject("BenefitsDB.CBenefitsDB")
    FetchCoPayByPatientProvider = objBnft.FetchCoPayByPatientProvider(lngPatientID, lngProviderID)
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Set objBnft = Nothing
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objBnft = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

