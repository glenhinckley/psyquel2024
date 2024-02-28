VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatApptBZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Class Name: CPatApptBz                                             '
'Date: 08/29/2000                                                   '
'Author: Rick Segura                                    '
'Description:  MTS business object designed to call methods         '
'              associated with the CPatApptDB class.                '
'--------------------------------------------------------------------
' Revision History:
'   R001: 10/25/2001 Richkun: Created ChangeStatus2() method as starting
'       point to replace the mess that is ChangeStatus().  The ChangeStatus()
'       method will continue to be used at the Check-In page until that
'       page is also redone.
'   R002: 11/06/2001 Richkun: Created FetchCheckInDetail() method
'   R003: 12/19/2001 Richkun: Altered ChangeStatus2() to support CancelFee and
'       CancelExplain parameters.
'   R004: 01/09/2002 Pena: Altered ChangeStatus2() to support tblProvider.fldCtr_Unbilled
'   R004: 02/12/2002 Richkun: Altered ChangeStatus2() so No-Show:Bill Patient claims
'       add a charge to patient, and no longer submit 'dummy' encounter.
'--------------------------------------------------------------------

Option Explicit

Private Const CLASS_NAME As String = "CPatApptBz"

Private Const SCHEDULED_STATUS As Long = 1
Private Const CONFIRMED_STATUS As Long = 2
Private Const ATTENDED_STATUS As Long = 3
Private Const CANCELLED_STATUS As Long = 4
Private Const HOLD_STATUS As Long = 5
Private Const NO_SHOW_STATUS As Long = 6
Private Const PHONE_CONTACT_STATUS As Long = 7
Private Const DO_NOT_COUNT_STATUS As Long = 8
Private Const WALK_IN_STATUS As Long = 9
Private Const NOT_OK As Long = 0
Private Const ALL_OK As Long = 1
Private Const IGNORE As Long = 2
Private Const NO_SHOW_CPTCODE As String = "00000"
Private Const TX_ID_NO_SHOW As Integer = 210

Public Function FetchByAppt(ByVal lngApptID As Long, _
                            Optional ByVal blnIncludeCancelled As Boolean = False) _
                            As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Retrieves records from the tblPatAppt table.         '
'Parameters: lngApptID - Appt ID to seek associations for           '
'            blnIncludeCancelled - Optional parameter that identifies
'               if records flagged as 'Cancelled' are to be included'
'               in the recordset. The default value is False.       '
'Returns: Recordset of requested markets                            '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim objPatAppt As ApptDB.CPatApptDB
    Dim rstSQL As ADODB.Recordset

    On Error GoTo ErrTrap:

    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
    Set rstSQL = objPatAppt.FetchByAppt(lngApptID)

    Set FetchByAppt = rstSQL
    
    Set objPatAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objPatAppt = Nothing
    Set rstSQL = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function FetchByApptPatient(ByVal lngApptID As Long, ByVal lngPatientID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 02/13/2002
'Author: Dave Richkun
'Description:  Retrieves appointment information using ApptID, PatientID
'Parameters: lngApptID - ID of the appointment
'            lngPatientID - ID of the patient
'Returns: Recordset
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------

    Dim objPatAppt As ApptDB.CPatApptDB

    On Error GoTo ErrTrap:

    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
    Set FetchByApptPatient = objPatAppt.FetchByApptPatient(lngApptID, lngPatientID)
    Set objPatAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objPatAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function FetchByID(ByVal lngPatApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 08/31/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description: Retrieves a single record in tblPatAppt having the ID '
'               passed in as lngPatApptID                           '
'Parameters:  lngApptID - System ID of the record to return         '
'Returns:  Recordset of 1 row                                       '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim objPatAppt As ApptDB.CPatApptDB
    Dim rstSQL As ADODB.Recordset
    Dim strErrMessage As String
    
    On Error GoTo ErrTrap:

    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
    Set rstSQL = objPatAppt.FetchByID(lngPatApptID)

    Set FetchByID = rstSQL
    
    If rstSQL.RecordCount = 0 Then
        strErrMessage = "An error occured while verifying the status of an appointment to update."
        GoTo ErrTrap
    End If
    
    Set objPatAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objPatAppt = Nothing
    Set rstSQL = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function


Public Function Insert(ByVal lngApptID As Long, ByVal lngPatientID As Long) As Long
'--------------------------------------------------------------------
'Date: 07/13/2002
'Author: Dave Richkun
'Description:  Inserts a patient appointment record into tblPatAppt
'Parameters:  lngApptID - ID of appointment
'             lngPatientID - ID of the patient
'Returns:  ID of PatientAppt record on success, -1 on failure
'--------------------------------------------------------------------
'Revision History:
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CPatApptDB

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CPatApptDB")
    Insert = objAppt.Insert(lngApptID, lngPatientID)
    Set objAppt = Nothing

    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Sub Update(ByVal lngID As Long, ByVal lngApptID As Long, ByVal lngPatientID As Long)
'--------------------------------------------------------------------
'Date: 08/29/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Updates a row into the tblPatAppt table.             '
'Parameters:  lngID - ID of the row in the table whose value will be'
'               updated.                                            '
'             lngApptID - System ID of appointment to asociate with '
'               the patient.                                        '
'             lngPatientID - System ID of the patient to associate  '
'               with the appointment.                               '
'Returns: Null                                                      '
'--------------------------------------------------------------------
'Revision History:                                                  '
'                                                                   '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CPatApptDB
    Dim strErrMessage As String

    On Error GoTo ErrTrap:

    'Verify data before proceeding.
    If Not VerifyData(lngID, lngApptID, lngPatientID, strErrMessage) Then
        GoTo ErrTrap
    End If
    
    Set objAppt = CreateObject("ApptDB.CPatApptDB")

    Call objAppt.Update(lngID, lngApptID, lngPatientID)

    Set objAppt = Nothing
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMessage
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Sub


Public Sub Delete(ByVal lngApptID As Long, ByVal lngPatientID As Long)
'--------------------------------------------------------------------
'Date: 08/29/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Physically deletes a record in tblPatAppt            '
'Parameters:  lngApptID - System ID of appointment associated with  '
'               the patient.                                        '
'             lngPatientID - System ID of the patient associated    '
'               with the appointment.                               '
'Returns:  Null                                                     '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CPatApptDB

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CPatApptDB")

    Call objAppt.Delete(lngApptID, lngPatientID)
    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
Public Sub DeleteByID(ByVal lngPatApptID As Long)
'--------------------------------------------------------------------
'Date: 12/18/2008                                                   '
'Author: Duane C Orth                                               '
'Description:  Physically deletes a record in tblPatAppt            '
'Parameters:  lngPatApptID - System ID of appointment associated with  '
'               the patient.                                        '
'Returns:  Null                                                     '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CPatApptDB

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CPatApptDB")

    Call objAppt.DeleteByID(lngPatApptID)
    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function Exists(ByVal lngApptID As Long, ByVal lngPatientID As Long) As Boolean
'--------------------------------------------------------------------
'Date: 08/29/2000
'Author: Rick "Boom Boom" Segura
'Description: Determines if an Appointment/Patient association exists in tblPatAppointment                         '
'Parameters:  lngApptID - System ID of appointment associated with the patient.                                        '
'             lngPatientID - System ID of the patient associated with the appointment.                               '
'Returns:  TRUE if association is found, FALSE otherwise            '
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------
    Dim objAppt As ApptDB.CPatApptDB
    Dim blnExists As Boolean

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CPatApptDB")
    
    blnExists = objAppt.Exists(lngApptID, lngPatientID)
    Exists = blnExists

    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'''Public Sub ChangeStatus(ByVal lngPatApptID As Long, ByVal lngStatusID As Long, _
'''        Optional ByVal lngPatientID As Long, Optional ByVal lngProviderID As Long, _
'''        Optional ByVal lngClinicID As Long, Optional ByVal dteDOS As Date, _
'''        Optional ByVal strCPTCode As String, Optional ByVal strModifier As String, _
'''        Optional ByVal lngDSM_IV_ID As Long, Optional ByVal dblFee As Double, _
'''        Optional ByVal intUnits As Integer, Optional ByVal strUserName As String, _
'''        Optional ByVal strTaxID As String, Optional ByVal dblAmtExpected As Double, _
'''        Optional ByVal dblAmtCollected As Double, Optional ByVal strReferPhy As String, _
'''        Optional ByVal strReferPhyID As String, Optional ByVal lngEncLogID As Long, _
'''        Optional ByVal strPatPmtType As String, Optional ByVal lngTransTypeID As Long, _
'''        Optional ByVal lngPatPmtID As Long = 0, Optional ByVal strPmtNotes As String, _
'''        Optional ByVal strCheckNum As String, Optional ByVal dteCheckDate As Date, _
'''        Optional ByVal strCertNum As String, Optional ByVal strNoShowFlag As String, _
'''        Optional ByVal lngApptID As Long, Optional ByVal lngCategoryID As Long, _
'''        Optional ByVal lngApptType As Integer, Optional ByVal dteStartDateTime As Date, _
'''        Optional ByVal dteEndDateTime As Date, Optional ByVal lngDuration As Long, _
'''        Optional ByVal strApptNote As String, Optional ByVal strApptDescription As String, _
'''        Optional ByVal blnUpdatePatPmtInfo As Boolean = False, _
'''        Optional ByVal blnUpdateApptInfo As Boolean = False)
''''--------------------------------------------------------------------
''''Date: 08/28/2000                                                   '
''''Author: Rick "Boom Boom" Segura                                    '
''''Description:  Function has been modified to be a transactional     '
''''           for updating the status of a patient appt, updating     '
''''           the core appt info., posting patient payments and       '
''''           inserting entries into the Encounter Log.               '
''''Parameters:  The values to be checked.                             '
''''Returns: Nothing                                                   '
''''--------------------------------------------------------------------
'''    Dim objPatAppt As ApptDB.CPatApptDB
'''    Dim objAppt As ApptDB.CApptDB
'''    Dim objElBill As ELBillBz.CELBillBZ
'''    Dim objPatPmt As EncounterLogBz.CPatPmtLogBz
'''    'Dim objBill As BillingBz.CBillCreationBz
'''    Dim rstPatAppt As ADODB.Recordset
'''    Dim strErrMsg As String
'''    Dim blnApplyPatientPayment As Boolean
'''    Dim lngEncID As Long
'''    Dim lngBillID As Long
'''
'''    On Error GoTo Err_Handler
'''    strErrMsg = ""
'''
'''    ' Fetch current patient appt info to compare to new changes
'''    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
'''    Set rstPatAppt = objPatAppt.FetchByID(lngPatApptID)
'''    'Set objBill = CreateObject("BillingBz.CBillCreationBz")
'''
'''    If rstPatAppt.RecordCount = 0 Then
'''        strErrMsg = "An error occurred while trying to validate the patient appointment."
'''        GoTo Err_Handler
'''    End If
'''
'''    Set objElBill = CreateObject("ELBillBz.CELBillBZ")
'''    'Set objEncLog = CreateObject("EncounterLogBz.CEncounterLogBz")
'''    Set objPatPmt = CreateObject("EncounterLogBz.CPatPmtLogBz")
'''
'''    ' Status Change should always occur, validate change 1st
'''    With rstPatAppt
'''
'''        Select Case ValidateChange(.Fields("fldApptStatusID"), .Fields("fldNoShowFlag"), _
'''                                    lngStatusID, strNoShowFlag, strErrMsg)
'''
'''            Case NOT_OK
'''                ' An invalid request was made, raise error and exit
'''                strErrMsg = "An error occurred while trying to validate the patient appointment."
'''                GoTo Err_Handler
'''
'''            Case ALL_OK
'''                ' Valid request made, change status
'''                Call objPatAppt.ChangeStatus(lngPatApptID, lngStatusID, strNoShowFlag, strUserName)
'''
'''                ' Now determine what to do with patient payments, if any
'''                Select Case lngStatusID
'''                    Case ATTENDED_STATUS
'''                    ' Make an entry in the Encounter Log appropriatley
'''                        If lngPatPmtID = 0 Then
'''                            lngEncID = objElBill.Insert(lngPatientID, lngProviderID, lngClinicID, _
'''                                                  dteDOS, strCPTCode, strModifier, lngDSM_IV_ID, strCertNum, _
'''                                                  dblFee, intUnits, strUserName, strTaxID, _
'''                                                  dblAmtExpected, dblAmtCollected, strReferPhy, strReferPhyID, lngApptID)
'''
'''                        Else
'''                            lngEncID = objElBill.Insert(lngPatientID, lngProviderID, lngClinicID, _
'''                                                  dteDOS, strCPTCode, strModifier, lngDSM_IV_ID, strCertNum, _
'''                                                  dblFee, intUnits, strUserName, strTaxID, _
'''                                                  dblAmtExpected, 0, strReferPhy, strReferPhyID, lngApptID)
'''                        End If
'''                        If lngEncID <= 0 Then
'''                            strErrMsg = "An error occurred while making an Encounter Log entry."
'''                            GoTo Err_Handler
'''                        'Else
'''                        '    lngBillID = objBill.CreateBill(lngEncID, strUserName)
'''                        '    If lngBillID <= 0 Then
'''                        '        strErrMsg = "An error occurred while attempting to bill this D.O.S."
'''                        '        GoTo Err_Handler
'''                        '    End If
'''                        End If
'''
'''                    Case NO_SHOW_STATUS
'''                        If InStr("PI", strNoShowFlag) Then
'''                            ' Billable No-Show
'''                            lngEncID = objElBill.Insert(lngPatientID, lngProviderID, lngClinicID, _
'''                                                  dteDOS, strCPTCode, strModifier, lngDSM_IV_ID, strCertNum, _
'''                                                  dblFee, intUnits, strUserName, strTaxID, _
'''                                                  dblAmtExpected, dblAmtCollected, strReferPhy, strReferPhyID, lngApptID)
'''                            If lngEncID <= 0 Then
'''                                strErrMsg = "An error occurred while making an Encounter Log entry."
'''                                GoTo Err_Handler
'''                            'Else
'''                            '    lngBillID = objBill.CreateBill(lngEncID, strUserName)
'''                            '    If lngBillID <= 0 Then
'''                            '        strErrMsg = "An error occurred while attempting to bill this D.O.S."
'''                            '        GoTo Err_Handler
'''                            '    End If
'''                            End If
'''                        End If
'''
'''                    Case Else
'''                    ' Determine if a payment needs to be inserted or updated
'''                        If dblAmtCollected > 0 Then
'''                            If lngPatPmtID = 0 Then
'''                            ' New patient payment
'''                                Call objPatPmt.InsertPosting(lngPatientID, lngProviderID, lngEncLogID, _
'''                                                             dblAmtExpected, dblAmtCollected, strPatPmtType, _
'''                                                             lngTransTypeID, Date, strCheckNum, strPmtNotes, _
'''                                                             strUserName, dteCheckDate, lngApptID)
'''                            Else
'''                            ' Update patient payment
'''                                Call objPatPmt.UpdatePosting(lngPatPmtID, lngPatientID, lngProviderID, lngEncLogID, _
'''                                                             dblAmtExpected, dblAmtCollected, Date, strPatPmtType, _
'''                                                             lngTransTypeID, strCheckNum, strPmtNotes, dteCheckDate, _
'''                                                             lngApptID)
'''                            End If
'''                        End If
'''
'''                End Select
'''
'''            Case IGNORE
'''                ' New status was same as old status, do not update patient appt info
'''                ' Determine if a payment needs to be inserted or updated
'''                Select Case lngStatusID
'''                    Case ATTENDED_STATUS
'''                        blnApplyPatientPayment = False
'''                    Case NO_SHOW_STATUS
'''                        If InStr("PI", strNoShowFlag) > 0 Then
'''                            blnApplyPatientPayment = False
'''                        Else
'''                            blnApplyPatientPayment = True
'''                        End If
'''
'''                    Case SCHEDULED_STATUS, CONFIRMED_STATUS, CANCELLED_STATUS
'''                        ' Need to apply change
'''                        blnApplyPatientPayment = True
'''                End Select
'''
'''                If blnApplyPatientPayment And blnUpdatePatPmtInfo Then
'''                    If dblAmtCollected > 0 Then
'''                        If lngPatPmtID = 0 Then
'''                        ' New patient payment
'''                            Call objPatPmt.InsertPosting(lngPatientID, lngProviderID, lngEncLogID, _
'''                                                         dblAmtExpected, dblAmtCollected, strPatPmtType, _
'''                                                         lngTransTypeID, Date, strCheckNum, strPmtNotes, _
'''                                                         strUserName, dteCheckDate, lngApptID)
'''                        Else
'''                        ' Update patient payment
'''                            Call objPatPmt.UpdatePosting(lngPatPmtID, lngPatientID, lngProviderID, lngEncLogID, _
'''                                                         dblAmtExpected, dblAmtCollected, Date, strPatPmtType, _
'''                                                         lngTransTypeID, strCheckNum, strPmtNotes, dteCheckDate, _
'''                                                         lngApptID)
'''                        End If
'''                    End If
'''                End If
'''
'''        End Select
'''    End With
'''
'''    Set objPatPmt = Nothing
'''    Set objElBill = Nothing
'''    Set objPatAppt = Nothing
'''    Set rstPatAppt = Nothing
'''    'Set objBill = Nothing
'''
'''    ' Update Core Appt info, this also should always be done
'''    If blnUpdateApptInfo Then
'''        Set objAppt = CreateObject("ApptDB.CApptDB")
'''        Call objAppt.Update(lngApptID, lngProviderID, lngClinicID, lngCategoryID, dteStartDateTime, _
'''                            dteEndDateTime, lngDuration, strApptNote, strUserName, strApptDescription, _
'''                            strCPTCode)
'''    End If
'''    Set objAppt = Nothing
'''
'''    ' Signal completion
'''    GetObjectContext.SetComplete
'''    Exit Sub
'''
'''Err_Handler:
'''    ' Signal incompletion
'''    GetObjectContext.SetAbort
'''    ' Free up the resources we used
'''    Set objPatPmt = Nothing
'''    Set objElBill = Nothing
'''    Set objPatAppt = Nothing
'''    Set rstPatAppt = Nothing
'''    Set objAppt = Nothing
'''    'Set objBill = Nothing
'''
'''    ' Raise the error
'''    If Err.Number = 0 Then
'''        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
'''    Else
'''        Err.Raise Err.Number, Err.Source, Err.Description
'''    End If
'''
'''End Sub


Public Sub ChangeStatus2(ByVal lngPatApptID As Long, ByVal lngStatusID As Long, _
        Optional ByVal lngPatientID As Long, Optional ByVal lngProviderID As Long, _
        Optional ByVal lngClinicID As Long, Optional ByVal dteDOS As Date, _
        Optional ByVal strCPTCode As String, Optional ByVal strModifier As String, _
        Optional ByVal lngDSMIV As Long, Optional ByVal dblFee As Double, _
        Optional ByVal intUnits As Double, Optional ByVal strUserName As String, _
        Optional ByVal strTaxID As String, Optional ByVal dblAmtExpected As Double, _
        Optional ByVal dblAmtCollected As Double, Optional ByVal strReferPhy As String, _
        Optional ByVal strReferPhyID As String, Optional ByVal dtAdmitDate As Date, _
        Optional ByVal dtDischargeDate As Date, Optional ByVal lngEncLogID As Long, _
        Optional ByVal strCertNum As String, Optional ByVal strNoShowFlag As String, _
        Optional ByVal strPatPmtType As String, Optional ByVal lngTransTypeID As Long, _
        Optional ByVal lngPatPmtID As Long = 0, Optional ByVal strPmtNotes As String, _
        Optional ByVal strCheckNum As String, Optional ByVal dteCheckDate As Date, _
        Optional ByVal lngApptID As Long, Optional ByVal blnRecurInstance As Boolean = False, _
        Optional ByVal dblCancelFee As Double = 0, Optional ByVal strCancelExplain As String)
'--------------------------------------------------------------------
'Date: 10/26/2001
'Author: Dave Richkun
'Description:  This 2nd-generation function is designed to replace the
'           original ChangeStatus() function.  This function is responsible
'           for changing the status of patient appointment records, along
'           with managing any required billing and/or payment posting.
'Parameters:
'

'             blnRecurInstance - Boolean value identifying if the appointment we are updating
'                     is a single instance from a recurring patient appointment series.  Any changes made
'                     to an instance of a recurring patient appointment forces the appointment instance
'                     to become a new one-time patient, so that billing procedures can be properly applied
'                     to the instance.

'Returns: Null
'--------------------------------------------------------------------
' Revision History:
'  R003
'--------------------------------------------------------------------

    Dim objPatAppt As ApptDB.CPatApptDB
    Dim objAppt As ApptDB.CApptDB
    Dim objApptBz As ApptBZ.CApptBZ
    '''Dim objElBill As ELBillBz.CELBillBZ
    '''Dim objPatPmt As EncounterLogBz.CPatPmtLogBz
    '''Dim objEncLog As EncounterLogBz.CEncounterLogBz
    Dim rstPatAppt As ADODB.Recordset
    Dim rstPat As ADODB.Recordset
    Dim strErrMsg As String
    Dim blnApplyPatPmt As Boolean
    Dim lngEncID As Long
    Dim lngBillID As Long
    Dim blnRuleByPass As Boolean
    Dim lngNewPatApptID As Long
    Dim intCtr As Integer
    Dim objPatRPPlan As BenefactorBz.CPatRPPlanBz
    Dim rstPatRPPlan  As ADODB.Recordset
    
    On Error GoTo ErrTrap:
    strErrMsg = ""
    
    'If the appointment represents a single instance of a recurring appointment, create the
    'new appointment.  We will be working with the new ApptID and PatientApptID values.
    If blnRecurInstance = True Then
        'Create a new one-time patient appointment to replace this instance.
        Set objApptBz = CreateObject("ApptBZ.CApptBZ")
        lngApptID = objApptBz.CloneInstance(lngApptID, dteDOS, strUserName)
        Set objApptBz = Nothing
        'Obtain the ID of the new Patient-Appt record
        Set rstPatAppt = FetchByAppt(lngApptID)
        For intCtr = 1 To rstPatAppt.RecordCount
            If rstPatAppt.Fields("fldPatientID").Value = lngPatientID Then
                lngPatApptID = rstPatAppt.Fields("fldPatApptID").Value
                Exit For
            End If
            rstPatAppt.MoveNext
        Next intCtr
        Set rstPatAppt = Nothing
    End If
    
    'Fetch current patient appt info and compare to new changes
    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
    Set rstPatAppt = objPatAppt.FetchByID(lngPatApptID)
    
    If rstPatAppt.RecordCount = 0 Then
        strErrMsg = "An error occurred while trying to validate the patient appointment."
        GoTo ErrTrap:
    End If
    
'''    Set objPatPmt = CreateObject("EncounterLogBz.CPatPmtLogBz")
'''    Set objEncLog = CreateObject("EncounterLogBz.CEncounterLogBz")
    
    'get the providerid if not passed in (ie, cancellation)
    If lngProviderID < 1 Then
        Set objApptBz = CreateObject("ApptBZ.CApptBZ")
        Set rstPat = objApptBz.FetchPatientApptByID(rstPatAppt.Fields("fldApptID"))
        If Not rstPat.EOF Then lngProviderID = rstPat.Fields("fldProviderID").Value
        Set rstPat = Nothing
        Set objApptBz = Nothing
    End If
    
    'Validate change before comitting to it.
    Select Case ValidateChange(rstPatAppt.Fields("fldApptStatusID"), rstPatAppt.Fields("fldNoShowFlag"), _
                    lngStatusID, strNoShowFlag, strErrMsg)
        Case NOT_OK
            ' An invalid request was made, raise error and exit
            strErrMsg = "An error occurred while trying to validate the patient appointment."
            GoTo ErrTrap:
            
        Case ALL_OK
            'R004
            setUnbilledCounter lngProviderID, dteDOS, rstPatAppt.Fields("fldApptStatusID"), lngStatusID
            
            ' Valid request made, change status
            Call objPatAppt.ChangeStatus(lngPatApptID, lngStatusID, strNoShowFlag, _
                dblCancelFee, strCancelExplain, strUserName)
            
            Select Case lngStatusID
                Case ATTENDED_STATUS
                ' Make an entry in the Encounter Log appropriatley
'''                    lngEncID = objElBill.Insert(lngPatientID, lngProviderID, lngClinicID, _
'''                        dteDOS, strCPTCode, strModifier, lngDSMIV, strCertNum, _
'''                        dblFee, intUnits, strUserName, strTaxID, dblAmtExpected, _
'''                        dblAmtCollected, strReferPhy, strReferPhyID, dtAdmitDate, dtDischargeDate, _
'''                        lngApptID, False, lngTransTypeID, strCheckNum, Date, "", lngPatPmtID)
                    
                    If lngEncID <= 0 Then
                        strErrMsg = "An error occurred while making an Encounter Log entry."
                        GoTo ErrTrap:
                    End If
                    
                Case NO_SHOW_STATUS
                    blnRuleByPass = True
                    Select Case strNoShowFlag
                        Case "N"
                            'If changing to 'Do Not Bill' from 'Bill Patient' or 'Bill Insurance', we must
                            'delete old EL and PP records.
                            If lngPatPmtID > 0 Then
                                '''Call objPatPmt.Delete(lngPatPmtID)
                                blnRuleByPass = False
                            End If
                        
                            If lngEncLogID > 0 Then
                                '''Call objElBill.Delete(lngEncLogID, strUserName, True)
                            End If
                        Case "I"
                            If lngPatPmtID > 0 Then
                                '''Call objPatPmt.Delete(lngPatPmtID)
                            End If
                            blnRuleByPass = False
                        
                            'Insert or Update the bill to the Insurance Company
                            If lngEncLogID <= 0 Then
'''                                lngEncID = objElBill.Insert(lngPatientID, lngProviderID, lngClinicID, _
'''                                            dteDOS, strCPTCode, strModifier, lngDSMIV, strCertNum, _
'''                                            dblFee, intUnits, strUserName, strTaxID, _
'''                                            dblAmtExpected, dblAmtCollected, strReferPhy, strReferPhyID, _
'''                                            dtAdmitDate, dtDischargeDate, lngApptID, , lngTransTypeID)
                                If lngEncID <= 0 Then
                                    strErrMsg = "An error occurred while making an Encounter Log entry."
                                    GoTo ErrTrap:
                                End If
                            Else
'''                                Call objEncLog.Update(lngEncLogID, lngPatientID, lngProviderID, lngClinicID, dteDOS, _
'''                                    strCPTCode, strModifier, lngDSMIV, strCertNum, dblFee, intUnits, strTaxID, dblAmtExpected, _
'''                                    dblAmtCollected, strReferPhy, strReferPhyID, dtAdmitDate, dtDischargeDate, strUserName, _
'''                                    lngTransTypeID, strCheckNum, dteCheckDate, strPmtNotes, blnRuleByPass)
                            End If
                        Case "P"
                            If dblFee > 0 Then 'R004
                                'Insert charge for patient No-Show
                                Set objPatRPPlan = CreateObject("BenefactorBz.CPatRPPlanBz")
                                Set rstPatRPPlan = objPatRPPlan.FetchRPPlansByPat(lngPatientID)
                                Set objPatRPPlan = Nothing
                                If lngPatPmtID = 0 Then
'''                                    Call objPatPmt.InsertPosting(lngPatientID, rstPatRPPlan.Fields("fldRPID").Value, _
'''                                        lngProviderID, lngEncLogID, dblFee, 0, "C", TX_ID_NO_SHOW, Date, strCheckNum, _
'''                                        "", strUserName, dteCheckDate, lngApptID)
                                End If
                                Set rstPatRPPlan = Nothing
                            End If
                    End Select
                
                Case Else
                    'If appointment status was changed back to 'Scheduled' or 'Confirmed' or 'Cancelled' from a 'No-Show' status,
                    'old Encounter and payment records must be removed.
                    If rstPatAppt.Fields("fldApptStatusID").Value = 6 Then
                        'The old status was a No-Show
                        If lngStatusID = SCHEDULED_STATUS Or lngStatusID = CONFIRMED_STATUS Or lngStatusID = CANCELLED_STATUS Then
                            If lngPatPmtID > 0 Then
                                '''Call objPatPmt.Delete(lngPatPmtID)
                                lngPatPmtID = 0
                            End If
                        
                            If lngEncLogID > 0 Then
                                '''Call objElBill.Delete(lngEncLogID, strUserName, True)
                                lngEncLogID = 0
                            End If
                        End If
                    End If
                    
                    'If appointment is being Cancelled, delete any Encounter Log, Payment records that may have existed.
                    If lngStatusID = CANCELLED_STATUS Then
                        If lngPatPmtID > 0 Then
                            '''Call objPatPmt.Delete(lngPatPmtID)
                        End If
                    
                        If lngEncLogID > 0 Then
                            '''Call objElBill.Delete(lngEncLogID, strUserName, True)
                        End If
                    End If
                    
                    ' Determine if a payment needs to be inserted or updated
                    If dblAmtCollected > 0 Then
                        Set objPatRPPlan = CreateObject("BenefactorBz.CPatRPPlanBz")
                        Set rstPatRPPlan = objPatRPPlan.FetchRPPlansByPat(lngPatientID)
                        Set objPatRPPlan = Nothing
                        If lngPatPmtID = 0 Then
                        ' New patient payment
'''                            Call objPatPmt.InsertPosting(lngPatientID, rstPatRPPlan.Fields("fldRPID").Value, lngProviderID, lngEncLogID, _
'''                                    dblAmtExpected, dblAmtCollected, strPatPmtType, _
'''                                    lngTransTypeID, Date, strCheckNum, strPmtNotes, _
'''                                    strUserName, dteCheckDate, lngApptID)
                        Else
                        ' Update patient payment
'''                            Call objPatPmt.UpdatePosting(lngPatPmtID, lngPatientID, rstPatRPPlan.Fields("fldRPID").Value, lngProviderID, lngEncLogID, _
'''                                    dblAmtExpected, dblAmtCollected, Date, strPatPmtType, _
'''                                    lngTransTypeID, strCheckNum, strPmtNotes, dteCheckDate, _
'''                                    lngApptID)
                        End If
                        Set rstPatRPPlan = Nothing
                    End If
            End Select
    End Select
    
    '''Set objEncLog = Nothing
    '''Set objPatPmt = Nothing
    '''Set objElBill = Nothing
    Set objPatAppt = Nothing
    Set rstPatAppt = Nothing
    
    ' Signal completion
    GetObjectContext.SetComplete
    Exit Sub
    
ErrTrap:
    ' Signal incompletion
    GetObjectContext.SetAbort
    ' Free up the resources we used
    '''Set objPatPmt = Nothing
    '''Set objElBill = Nothing
    Set objPatAppt = Nothing
    Set rstPatAppt = Nothing
    Set objAppt = Nothing
    Set rstPat = Nothing
    Set rstPatRPPlan = Nothing
    Set objPatRPPlan = Nothing
    
    ' Raise the error
    If Err.Number = 0 Then
        Err.Raise vbObjectError, CLASS_NAME, strErrMsg
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Sub

Public Function FetchCheckInDetails(ByVal lngPatApptID As Long) As ADODB.Recordset
'--------------------------------------------------------------------
'Date: 11/06/2001
'Author: Dave Richkun
'Description:  Retrieves detailed Appointment information for a specific
'              Patient-Appointment ID
'Parameters: lngPatApptID - ID of the patient-appointment
'Returns: ADO Recordset
'--------------------------------------------------------------------
'Revision History:
'  R001: Created
'--------------------------------------------------------------------
    Dim objPatAppt As ApptDB.CPatApptDB
    
    On Error GoTo ErrTrap:
    
    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
    Set FetchCheckInDetails = objPatAppt.FetchCheckInDetails(lngPatApptID)
    
    Set objPatAppt = Nothing
    GetObjectContext.SetComplete
    
    Exit Function
    
ErrTrap:
    Set objPatAppt = Nothing
    GetObjectContext.SetAbort

End Function


'----------------------------------
' Private Methods
'----------------------------------

Private Function ValidateChange(ByVal lngCurrentStatusID As Long, _
                                ByVal varCurrentNoShowFlag As Variant, _
                                ByVal lngNewStatusID As Long, _
                                ByRef strNoShowFlag As String, _
                                ByRef strErrMsg As String) As Long
'--------------------------------------------------------------------
'Date: 09/05/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Validates a patient/appointment status change        '
'Parameters:  The current and new status IDs                        '
'Returns: TRUE if the change is allowed, FALSE otherwise            '
'-------------------------------------------------------------------'
    
    Select Case lngCurrentStatusID
    
        Case ATTENDED_STATUS
            ' Only allow ATTENDED (no change)
            If lngNewStatusID <> ATTENDED_STATUS Then
                ValidateChange = NOT_OK
                strErrMsg = "Cannot change the appointment status of an 'Attended' appointment."
                Exit Function
            Else
                ValidateChange = IGNORE
            End If
        Case SCHEDULED_STATUS, CONFIRMED_STATUS
            ' Allow CANCELLED, ATTENDED, NO SHOW
            If lngNewStatusID = HOLD_STATUS Then
                ValidateChange = NOT_OK
                strErrMsg = "Cannot hold this appointment."
                Exit Function
            End If

        Case HOLD_STATUS
            If Not (lngNewStatusID = SCHEDULED_STATUS) Then
                ValidateChange = NOT_OK
                strErrMsg = "Can only change a 'Held' appointment to a 'Scheduled' status."
                Exit Function
            End If
            
        Case NO_SHOW_STATUS
        Case CANCELLED_STATUS
            
    End Select
    
    If lngNewStatusID = NO_SHOW_STATUS Then
        If Not strNoShowFlag > "" Or InStr("IPN", strNoShowFlag) = 0 Then
            ValidateChange = NOT_OK
            strErrMsg = "An invalid NO_SHOW Flag was supplied."
            Exit Function
        End If
    Else
        strNoShowFlag = ""
    End If
    
    If ValidateChange = IGNORE Then Exit Function
    ValidateChange = ALL_OK
End Function

Private Function VerifyData(ByVal lngID As Long, ByVal lngApptID As Long, _
                            ByVal lngPatientID As Long, strErrMessage As String) As Boolean
'--------------------------------------------------------------------
'Date: 08/28/2000                                                   '
'Author: Rick "Boom Boom" Segura                                    '
'Description:  Verifies all required data has been provided by the user.
'Parameters:  The values to be checked.                             '
'Returns: Boolean value identifying if all criteria has been satisfied.
'--------------------------------------------------------------------
    
    'Validation rules here
    If lngApptID < 1 Then
        strErrMessage = "Invalid Appointment ID passed."
        VerifyData = False
        Exit Function
    End If
    
    If lngPatientID < 1 Then
        strErrMessage = "Invalid Patient ID passed."
        VerifyData = False
        Exit Function
    End If
    
    'If we get here, all is well...
    VerifyData = True

End Function

Private Sub setUnbilledCounter(ByVal lngProvID As Long, ByVal dteDOS As Date, ByVal lngOldStatus As Long, ByVal lngNewStatus As Long) 'R004
'--------------------------------------------------------------------
'Date: 01/09/2002
'Author: Eric Pena
'Description:  Inc(dec)rements tblProvider.fldCtr_Unbilled as necessary
'Parameters:  lngProvID - ID of provider whose appointment is being changed
'               dteDOS - DOS of appt
'               lngOldStatus - old appt status
'               lngNewStatus - new appt status
'Returns: null
'--------------------------------------------------------------------
    On Error GoTo ErrHand
    Dim objAppt As ApptDB.CPatApptDB
    
    'only affects appts older than 3 days
    If DateDiff("d", dteDOS, Date) <= 3 Then Exit Sub
    
    Set objAppt = CreateObject("ApptDB.CPatApptDB")
    
    If lngOldStatus < ATTENDED_STATUS And lngNewStatus > CONFIRMED_STATUS Then
        objAppt.UpdateUnbilledCounter lngProvID, -1
    End If
    
        
    Set objAppt = Nothing
    Exit Sub
    
ErrHand:
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ChangeStatus(ByVal lngPatApptID As Long, ByVal lngApptStatusID As Long, _
        ByVal strNoShowFlag As String, ByVal dblCancelFee As Double, _
        ByVal strCancelExplain As String, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 07/09/2002
'Author: Dave Richkun
'Description:  Updates the status of a single Patient-Appointment record
'Parameters:  lngPatApptID - ID of the row in tblPatientAppt whose status value will be updated
'             lngApptStatusID - The value representing the updated appointment status
'             strNoShowFlag - If appointment status is 'No-Show', represents the user's decision
'                   to manage the No-Show.  May be one of:
'                    N - Do not Bill for No-Show
'                    P - Bill patient for No-Show
'                    I - Bill Insurance Company for No-Show
'             dblCancelFee - If appointment status is 'Cancelled', represents the amount
'                   charged to patient as Cancellation Fee
'             strCancelExplain - If appointment is 'Cancelled', represents the user's explanation
'                   for cancellation (assumed to be given by patient)
'             strUserName - User performng action
'Returns:  Null
'--------------------------------------------------------------------

    Dim objPatAppt As ApptDB.CPatApptDB

    On Error GoTo ErrTrap:

    Set objPatAppt = CreateObject("ApptDB.CPatApptDB")
    Call objPatAppt.ChangeStatus(lngPatApptID, lngApptStatusID, strNoShowFlag, _
            dblCancelFee, strCancelExplain, strUserName)
    Set objPatAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete
    
    'Release resources
    Set objPatAppt = Nothing
    
    Exit Sub
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.
    GetObjectContext.SetAbort
    Set objPatAppt = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function Reset(ByVal lngPatApptID As Long, ByVal strUserName As String)
'--------------------------------------------------------------------
'Date: 02/15/2002
'Author: Dave Richkun
'Description:  Resets a patient appointment to a status of 'Attended' and
'              reverses any prior Cancellation and No-Show cancellations
'              applied against the patient.
'Parameters:  lngPatApptID - ID of the patient appointment
'             strUserName - User name of the user initiating the method
'Returns: Null
'--------------------------------------------------------------------
'Revision History:                                                  '
'--------------------------------------------------------------------

    Dim objAppt As ApptDB.CPatApptDB

    On Error GoTo ErrTrap:

    Set objAppt = CreateObject("ApptDB.CPatApptDB")
    Call objAppt.Reset(lngPatApptID, strUserName)
    Set objAppt = Nothing
    
    'Signal successful completion
    GetObjectContext.SetComplete

    Exit Function
    
ErrTrap:
    'Signal incompletion and raise the error to the calling environment.  The
    'condition handles custom business rule checks we may have established.
    GetObjectContext.SetAbort
    Set objAppt = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



