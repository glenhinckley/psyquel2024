VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBenefactorClinicBz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CBenefactorClinicBz
'Date: 02/06/2000
'Author: Eric Pena
'Description:  MTS object designed to fetch from the tblBenefactorClinic table.
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------

Option Explicit
Private Const TABLE_NAME As String = "tblBenefactorClinic"
Private Const CLASS_NAME As String = "CBenefactorClinicDB"
Public Function FetchClinics(ByVal lngBenefactorID As Long) As Variant
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Returns a variant array containing a list of all Clinics to which a Benefactor is associated.
'Parameters: lngBenefactorID - The ID of the Benefactor
'Returns: A variant array containing a list of all clinics to which a Benefactor is associated.
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactorClinic As BenefactorDB.CBenefactorClinicDB
    
    
    Set objBenefactorClinic = CreateObject("BenefactorDB.CBenefactorClinicDB")
    FetchClinics = objBenefactorClinic.FetchClinics(lngBenefactorID)
    Set objBenefactorClinic = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactorClinic = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function IsEnabled(ByVal lngBenefactorID As Long, ByVal lngClinicID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 02/06/2000
'Author: Eric Pena
'Description:  Lets the caller know if a given Benefactor/clinic association is active
'Parameters: lngBenefactorID - The ID of the Benefactor
'                  lngClinicID - The ID of the Clinic
'Returns: T if the association is enabled/exists, F otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactorClinic As BenefactorDB.CBenefactorClinicDB
    
    
    Set objBenefactorClinic = CreateObject("BenefactorDB.CBenefactorClinicDB")
    IsEnabled = objBenefactorClinic.IsEnabled(lngBenefactorID, lngClinicID)
    Set objBenefactorClinic = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    Exit Function
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactorClinic = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


