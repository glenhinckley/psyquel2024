VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatientProviderBz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatientProviderBz
'Date: 05/31/2000
'Author: Eric Pena
'Description:  MTS business object designed to call methods associated
'              with the CPatientProviderDB class.
'--------------------------------------------------------------------
'Revisions:
'   R001 01/31/2001 D. Nichol: Added optional parameter blnIncludeDisabled to
'              FetchPatientProviders
'   R002 Valentines Day 2001!  Nichol: Added method InsertWeb() for insertion
'        of pat-provider association from the web page--that is, allows the
'        association to be seen on the web site by both providers.  Also added
'        methods DeletedByAssoc() and UpdatePatProviderWebAssoc()
'   R003 03/30/2001 D. Nichol: Added DeletedByProvider()
'   R004 04/10/2001 Richkun: Added strUserName parameter to Deleted() method
'   R005 02/08/2002 Richkun: Added method FetchPatientUserProviders()
'--------------------------------------------------------------------
Option Explicit

Private Const CLASS_NAME As String = "CBenefactorBz"
Private Const TABLE_NAME As String = "tblBenefactor"

Public Function FetchPatientProviders(ByVal lngPatientID As Long, Optional ByVal blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 05/25/2000
'Author: Eric Pena
'Description:  Retrieves all Providers associated with a specific patient.
'Parameters: lngPatientID - ID of the patient whose Provider records will be
'                 retrieved.
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB

    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Set FetchPatientProviders = objBenefactor.FetchPatientProviders(lngPatientID, blnIncludeDisabled)
    
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

Public Function FetchPatientUserProviders(ByVal lngPatientID As Long, _
                ByVal lngUserID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/08/2002
'Author: Dave Richkun
'Description:  Retrieves all Providers associated with a specific patient where
'               the provider is also serviced by the user (Provider/Office Manager).
'Parameters: lngPatientID - ID of the patient whose Provider records will be retrieved.
'            lngUserID - ID of the User (Provider/Office Manager)
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB

    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Set FetchPatientUserProviders = objBenefactor.FetchPatientUserProviders(lngPatientID, lngUserID)
    Set objBenefactor = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function FetchWebPatientProviders(ByVal lngPatientID As Long, Optional ByVal blnIncludeDisabled As Boolean = False) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/14/2001
'Author: Dave Nichol
'Description:  Retrieves all Providers associated with a specific patient,
'               _specifically who were shared by the web site, that is, have
'               a 'yes' in their tblPatientProvider.fldWebAssocYN association with
'               the patient.
'Parameters: lngPatientID - ID of the patient whose Provider records will be
'                 retrieved.
'Returns: ADODB.Recordset
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB

    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Set FetchWebPatientProviders = objBenefactor.FetchWebPatientProviders(lngPatientID, blnIncludeDisabled)
    
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

Public Function Insert(ByVal lngPatientID As Long, ByVal lngProviderID As Long, ByVal strAddedBy As String) As Long
'--------------------------------------------------------------------
'Date: 05/31/2000
'Author: Eric Pena
'Description: Creates a provider-patient association                '
'Parameters: lngPatientID - Patient ID                              '
'            lngProviderID - Provider ID                            '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    Dim strErrMessage As String
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    If Exists(lngPatientID, lngProviderID) Then
        Call DeletedByAssoc(False, lngPatientID, lngProviderID)
        Insert = 0
    Else
        Insert = objBenefactor.Insert(lngPatientID, lngProviderID, strAddedBy)
    End If
    Set objBenefactor = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
Public Function InsertWeb(ByVal lngPatientID As Long, ByVal lngProviderID As Long, _
                          ByVal lngUserID As Long, ByVal strAddedBy As String) As Long
'--------------------------------------------------------------------
'Date: 02/14/2001
'Author: Dave Nichol
'Description: Creates a provider-patient association                '
'Parameters: lngPatientID - Patient ID                              '
'            lngProviderID - Provider ID                            '
'            lngUserID - User ID of inserter                        '
'            strAddedBy - username of inserter                      '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    Dim strErrMessage As String
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    If Exists(lngPatientID, lngProviderID) Then
        Call DeletedByAssoc(False, lngPatientID, lngProviderID)
        Call UpdatePatProviderWebAssoc(True, lngPatientID, lngProviderID, lngUserID)
        'InsertWeb = 0
    Else
        InsertWeb = objBenefactor.InsertWeb(lngPatientID, lngProviderID, lngUserID, strAddedBy)
    End If
    Set objBenefactor = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function

ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
Public Sub Deleted(ByVal blnDeleted As Boolean, ByVal lngID As Long, _
                   ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 05/31/2000
'Author: Eric Pena
'Description:  Flags a row in the tbPatientProvider table marking the row as
'              deleted or undeleted.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngID - ID of the row in the table whose value will be
'               updated.
'            strUserName - User name of the user responsible for deleting the
'               Patient-Provider relationship R004
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------
    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Call objBenefactor.Deleted(blnDeleted, lngID, strUserName)
    
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

Public Sub DeletedByAssoc(ByVal blnDeleted As Boolean, ByVal lngPatientID As Long, ByVal lngProviderID As Long)
'--------------------------------------------------------------------
'Date: 02/14/2001
'Author: David Nichol
'Description:  Flags a row in the tbPatientProvider table marking the row as
'              deleted or undeleted.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngPatientID - ID of the patient
'            lngProviderID - ID of the provider
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Call objBenefactor.DeletedByAssoc(blnDeleted, lngPatientID, lngProviderID)
    
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
Public Sub DeletedByProvider(ByVal blnDeleted As Boolean, ByVal lngProviderID As Long, ByVal lngOwnerID)
'--------------------------------------------------------------------
'Date: 03/30/2001
'Author: Dave Nichol & Joshua Lockwood (and a little bit of Segura)
'Description:  Flags a row in the tbPatientProvider table marking the row as
'              deleted or undeleted, by providerID--deleting all the associations
'              where a given provider is the owner.
'Parameters: blnDeleted - Boolean value identifying if the record is to
'               be deleted (True) or undeleted (False).
'            lngProviderID - ID of the provider
'            lngOwnerID - ID of the owner of the patient.
'Returns: Null
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    
    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Call objBenefactor.DeletedByProvider(blnDeleted, lngProviderID, lngOwnerID)
    
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
Public Function Exists(ByVal lngPatientID As Long, ByVal lngProviderID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 05/31/2000
'Author: Eric Pena
'Description:  Determines if the association identical to the parameters already exists in the table.
'Parameters:  all are values to be checked
'Returns: True if the assoc exists, false otherwise
'--------------------------------------------------------------------
'Revision History:
'
'--------------------------------------------------------------------

    
    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    blnExists = objBenefactor.Exists(lngPatientID, lngProviderID)

    Exists = blnExists

    Set objBenefactor = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub UpdatePatProviderWebAssoc(ByVal blnWebAssoc As Boolean, ByVal lngPatientID As Long, _
                                     ByVal lngProviderID As Long, ByVal lngUserID As Long)
'--------------------------------------------------------------------
'Date: 02/14/2001
'Author: David Nichol
'Description:  Updates tblPatientProvider.fldWebAssocYN for a patient/provider
'               pair.
'Parameters: blnWebAssoc - True if the association is to be visible to web users.
'            lngPatientID - ID of the patient
'            lngProviderID - ID of the provider
'            lngUserID - ID of the user who is inserting the association
'Returns: Null
'--------------------------------------------------------------------

    Dim objBenefactor As BenefactorDB.CPatientProviderDB
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:
    
    Set objBenefactor = CreateObject("BenefactorDB.CPatientProviderDB")
    Call objBenefactor.UpdatePatProviderWebAssoc(blnWebAssoc, lngPatientID, lngProviderID, lngUserID)

    Set objBenefactor = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    Set objBenefactor = Nothing
    GetObjectContext.SetAbort
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

